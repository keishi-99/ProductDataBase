# 成績書・リスト・チェックシート生成の非同期化 設計書

**作成日:** 2026-03-19
**対象ブランチ:** Claude

---

## 概要

成績書・リスト・チェックシートの Excel 生成処理を非同期化し、処理中に LoadingForm（GIF アニメーション）を表示することで UI フリーズを解消する。

---

## 背景・動機

現状、3つの Excel 生成処理はいずれも UI スレッドで同期実行されており、ClosedXML のファイル読み書きや COM Interop（Excel 起動）の間、アプリケーション全体がフリーズする。

現在の対応は `Cursor.Current = Cursors.WaitCursor` のみで、ユーザーが処理中か否かを判別しにくい。

基板情報ボタン（`SubstrateRegistrationWindow`）では `RunOnStaThreadAsync` + ボタン無効化のパターンで既に非同期化済み。印刷処理（`Print()`）では `Task.Run + LoadingForm.ShowDialog()` パターンで非同期化済み。今回はこれらのパターンを改良して 3 つの Excel 生成処理に適用する。

---

## 技術制約

- **COM Interop（Excel 起動）は STA スレッドが必要。** `Task.Run` のスレッドプールスレッドは MTA のため使用不可。専用 STA スレッド（`RunOnStaThreadAsync`）を使う。
- **`SaveFileDialog`・`OpenFileDialog` も STA スレッド要件あり。** ただし LoadingForm 表示中に別スレッドからダイアログを出すと UX が混乱するため、これらは事前に UI スレッドで実行する。
- **`MessageBox` はどのスレッドでも動作する。**
- **`LoadingForm.Invoke` はフォームのハンドルが作成される前に呼ぶことができない。** `ContinueWith` で `Invoke` を呼ぶパターン（既存 `Print()` のパターン）は、処理が `ShowDialog` 呼び出し前に完了した場合に競合する。今回は `LoadingForm.Load` イベントで STA スレッドを開始するパターンを採用し、この競合を解消する。

---

## 設計方針（アプローチ B: ダイアログを前処理として分離）

各処理を以下の 2 段階に分割する。

```
[UIスレッド] 前処理ダイアログ（あれば）→ キャンセル時はここで終了
    ↓
[UIスレッド] ボタン無効化 + LoadingForm.ShowDialog 開始
    ↓ (Load イベントで)
[STAスレッド] 重い処理（ClosedXML + COM Interop）
    ↓
[UIスレッド] LoadingForm 閉じる
    ↓
[UIスレッド] ボタン再有効化 + 完了通知 or エラーメッセージ
```

---

## 各処理の担当分け

| 処理 | UIスレッド（前処理） | STAスレッド（LoadingForm 中） |
|------|---------------------|-------------------------------|
| 成績書 | ConfigXlsm 読み込み + FindExcelFile + SaveFileDialog | テンプレート読み込み + 書き込み + SaveAs |
| リスト | なし | ClosedXML 全処理 + COM Interop（Excel 起動） |
| チェックシート | ConfigXlsm 読み込み + InputDialog（温度・湿度） | ClosedXML 全処理 + COM Interop（Excel 起動） |

---

## 変更詳細

### 1. `CommonUtils.cs` への `RunOnStaThreadAsync` 移動

`SubstrateRegistrationWindow` に存在する `private static Task RunOnStaThreadAsync(Action action)` を `CommonUtils` に `public static` として移動し、全フォームから参照できるようにする。

`SubstrateRegistrationWindow` は `CommonUtils.RunOnStaThreadAsync` を使う形に変更する。

```csharp
// CommonUtils に追加
public static Task RunOnStaThreadAsync(Action action)
{
    var tcs = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
    var thread = new Thread(() =>
    {
        try { action(); tcs.SetResult(); }
        catch (Exception ex) { tcs.SetException(ex); }
    });
    thread.SetApartmentState(ApartmentState.STA);
    thread.Start();
    return tcs.Task;
}
```

---

### 2. `ExcelServiceClosedXml.cs` の変更

#### 2-1. `ReportGeneratorClosedXml` の分割

**新メソッド `PrepareReport`（UIスレッドで呼ぶ）:**
- `LoadConfigWorkbook()` を呼ぶ
- `GetReportConfig(configWorkbook, productModel)` を呼ぶ（`FindExcelFile` を含む。ファイル複数一致時は `MessageBox` + `OpenFileDialog` が出る）
- `SaveReport` 内にあった `SaveFileDialog` を呼んで保存先パスを取得
- `OperationCanceledException`（OpenFileDialog または SaveFileDialog のキャンセル）は catch して `null` を返す
- 戻り値: `(string templateFilePath, string savePath)?`（キャンセル時は `null`）

**新メソッド `ExecuteReport`（STAスレッドで呼ぶ）:**
- テンプレートファイルを ClosedXML で開く
- `PopulateReportSheet()` を呼ぶ
- 必要に応じて `GetUsedSubstrateData()` + `UpdateSubstrateDetailsInExcel()` を呼ぶ
- `ForcePageBreakPreview()` を呼ぶ
- `workbook.SaveAs(savePath)` で保存
- MessageBox は含めない。例外はすべて呼び出し元に伝播する

**旧 `GenerateReport` の除去:** 呼び出し元が `PrepareReport` + `ExecuteReport` を使う形に変更されるため削除する。

---

#### 2-2. `CheckSheetGeneratorClosedXml` の分割

**新メソッド `PrepareCheckSheet`（UIスレッドで呼ぶ）:**
- ConfigCheckSheet.xlsm を ClosedXML で開く
- `LoadAndExtractConfig(workBook, productMaster)` を呼んで `CheckSheetConfigData` を取得
- `GetTemperatureAndHumidity(excelData)` を呼んで温度・湿度を取得（InputDialog1 が出る）
- `OperationCanceledException`（InputDialog1 のキャンセル）は catch して `null` を返す
- 戻り値: `(CheckSheetConfigData configData, string temperature, string humidity)?`（キャンセル時は `null`）

**新メソッド `ExecuteCheckSheet`（STAスレッドで呼ぶ）:**
- `configData`、`temperature`、`humidity` を引数として受け取る
- `FormatDate()`
- `LoadWorkbookOrDefault()` でターゲット Workbook を読み込む（`FileStream` は `using` で管理）
- `PopulateExcelSheets()`
- `HideSheets()`
- `SaveWorkbook()`
- `PrintExcelFile()`（COM Interop）
- 例外はすべて呼び出し元に伝播する

**旧 `GenerateCheckSheet` の除去。**

---

#### 2-3. `ListGeneratorClosedXml` の変更

`GenerateList` の内部にある `try-catch`（例外を握りつぶして MessageBox を出している）を除去し、例外を呼び出し元に伝播するよう変更する。

```csharp
// 変更前
public static void GenerateList(...) {
    try { ... }
    catch (Exception ex) {
        MessageBox.Show(ex.Message, ...); // 握りつぶし → 除去
    }
}

// 変更後
public static void GenerateList(...) {
    // try-catch なし。例外は呼び出し元に伝播
    ...
}
```

---

### 3. 呼び出し側の変更（`ProductRegistration2Window`, `HistoryWindow`, `SubstrateChange2`）

各フォームの `GenerateReport`・`GenerateList`・`GenerateCheckSheet` メソッドを以下のパターンで統一する。

**`LoadingForm.Load` イベントで STA スレッドを開始するパターン**を採用する。これにより、フォームのハンドル作成前に `Invoke` が呼ばれる競合を防ぐ。

**注意: `async (_, _) => { ... }` は `async void` のイベントハンドラ。`try` ブロックは必ず `await` を含む全体を包むこと。`await` より前に同期処理を追加してはならない（`await` 前の例外は `taskException` に保存されず `SynchronizationContext` に投げられる）。**

#### 成績書の例

```csharp
private void GenerateReport()
{
    // [UIスレッド] 前処理: Config読み込み + FindExcelFile + SaveFileDialog
    // PrepareReport がスローする例外（FileNotFoundException 等）もここで捕捉する
    (string templateFilePath, string savePath)? prepared;
    try {
        prepared = ExcelServiceClosedXml.ReportGeneratorClosedXml.PrepareReport(
            _productMaster.ProductModel);
    } catch (Exception ex) {
        MessageBox.Show(ex.Message,
            $"[{nameof(GenerateReport)}]エラー",
            MessageBoxButtons.OK, MessageBoxIcon.Error);
        return;
    }
    if (prepared is null) return; // キャンセル（OpenFileDialog or SaveFileDialog）

    // [UIスレッド] ボタン無効化 + LoadingForm 表示
    GenerateReportButton.Enabled = false;
    Exception? taskException = null;
    using var loadingForm = new LoadingForm();
    loadingForm.Load += async (_, _) => {
        // ※ try は await を含む全体を包む。await より前に同期処理を追加しないこと。
        try {
            await CommonUtils.RunOnStaThreadAsync(() =>
                ExcelServiceClosedXml.ReportGeneratorClosedXml.ExecuteReport(
                    _productMaster, _productRegisterWork,
                    prepared.Value.templateFilePath, prepared.Value.savePath));
        } catch (Exception ex) {
            taskException = ex;
        } finally {
            loadingForm.Close();
        }
    };
    loadingForm.ShowDialog(this); // Load イベント完了後にリターン

    // [UIスレッド] 完了・エラー処理
    GenerateReportButton.Enabled = true;
    if (taskException is not null) {
        MessageBox.Show(taskException.Message,
            $"[{nameof(GenerateReport)}]エラー",
            MessageBoxButtons.OK, MessageBoxIcon.Error);
    } else {
        MessageBox.Show("成績書が正常に生成されました。", "完了",
            MessageBoxButtons.OK, MessageBoxIcon.Information);
    }
}
```

#### リストの例

```csharp
private void GenerateList()
{
    // [UIスレッド] 前処理なし
    GenerateListButton.Enabled = false;
    Exception? taskException = null;
    using var loadingForm = new LoadingForm();
    loadingForm.Load += async (_, _) => {
        try {
            await CommonUtils.RunOnStaThreadAsync(() =>
                ExcelServiceClosedXml.ListGeneratorClosedXml.GenerateList(
                    _productMaster, _productRegisterWork));
        } catch (Exception ex) {
            taskException = ex;
        } finally {
            loadingForm.Close();
        }
    };
    loadingForm.ShowDialog(this);

    GenerateListButton.Enabled = true;
    if (taskException is not null) {
        MessageBox.Show(taskException.Message,
            $"[{nameof(GenerateList)}]エラー",
            MessageBoxButtons.OK, MessageBoxIcon.Error);
    }
}
```

#### チェックシートの例

```csharp
private void GenerateCheckSheet()
{
    // [UIスレッド] 前処理: Config読み込み + InputDialog（温度・湿度）
    // PrepareCheckSheet がスローする例外（FileNotFoundException 等）もここで捕捉する
    (CheckSheetConfigData configData, string temperature, string humidity)? prepared;
    try {
        prepared = ExcelServiceClosedXml.CheckSheetGeneratorClosedXml.PrepareCheckSheet(
            _productMaster);
    } catch (Exception ex) {
        MessageBox.Show(ex.Message,
            $"[{nameof(GenerateCheckSheet)}]エラー",
            MessageBoxButtons.OK, MessageBoxIcon.Error);
        return;
    }
    if (prepared is null) return; // キャンセル（InputDialog）

    CheckSheetPrintButton.Enabled = false;
    Exception? taskException = null;
    using var loadingForm = new LoadingForm();
    loadingForm.Load += async (_, _) => {
        try {
            await CommonUtils.RunOnStaThreadAsync(() =>
                ExcelServiceClosedXml.CheckSheetGeneratorClosedXml.ExecuteCheckSheet(
                    _productMaster, _productRegisterWork,
                    prepared.Value.configData,
                    prepared.Value.temperature, prepared.Value.humidity));
        } catch (Exception ex) {
            taskException = ex;
        } finally {
            loadingForm.Close();
        }
    };
    loadingForm.ShowDialog(this);

    CheckSheetPrintButton.Enabled = true;
    if (taskException is not null) {
        MessageBox.Show(taskException.Message,
            $"[{nameof(GenerateCheckSheet)}]エラー",
            MessageBoxButtons.OK, MessageBoxIcon.Error);
    }
}
```

**注意: `HistoryWindow` では `GenerateReportButton`・`GenerateListButton`・`GenerateCheckSheetButton` を、`SubstrateChange2` では `SubstrateListPrintButton` を同様にパターン適用する。**

---

## 変更ファイル一覧

| ファイル | 変更内容 |
|---------|---------|
| `Other/CommonUtils.cs` | `RunOnStaThreadAsync` を `public static` として追加 |
| `Substrate/SubstrateRegistrationWindow.cs` | `private static RunOnStaThreadAsync` を削除し `CommonUtils` を参照 |
| `ExcelService/ExcelServiceClosedXml.cs` | `GenerateReport` → `PrepareReport`+`ExecuteReport` に分割（try-catch 除去）; `GenerateCheckSheet` → `PrepareCheckSheet`+`ExecuteCheckSheet` に分割（try-catch 除去）; `GenerateList` の try-catch を除去 |
| `Product/ProductRegistration2Window.cs` | `GenerateReport`・`GenerateList`・`GenerateCheckSheet` を `LoadingForm.Load` パターンに変更 |
| `Other/HistoryWindow.cs` | 同上（3メソッド） |
| `Product/SubstrateChange2.cs` | `GenerateList` を `LoadingForm.Load` パターンに変更 |

---

## エラーハンドリング

- `ExecuteReport` / `ExecuteCheckSheet` / `GenerateList`（変更後）は例外をそのままスローし、呼び出し側の `loadingForm.Load` ハンドラ内で捕捉して `taskException` に保存する
- `PrepareReport` / `PrepareCheckSheet` 内での `OperationCanceledException`（ダイアログキャンセル）は catch して `null` を返す
- その他の例外は `null` を返さずそのままスローし、呼び出し側に伝播する

---

## 補足

### COM Interop と Excel プロセスのライフサイクル

`SaveAndPrintExcel`（リスト）および `PrintExcelFile`（チェックシート）は `xlApp.Visible = true` に設定し、COM オブジェクト解放後も Excel プロセスが生存し続ける仕様（意図的）。STAスレッドが終了した後も Excel.exe は起動したまま残る。これは既存の動作と変わらない。

### `PrepareCheckSheet` 内での XLWorkbook の Dispose

`PrepareCheckSheet` は ConfigCheckSheet.xlsm を ClosedXML で開き、`CheckSheetConfigData` を取り出した後に確実に `Dispose`（`using` を使用）しなければならない。`CheckSheetConfigData` は純粋なデータオブジェクト（`XLWorkbook` への参照を持たない）であり、`ExecuteCheckSheet` では ConfigCheckSheet.xlsm を再度開く必要がある。実装時にリソースリークに注意すること。

---

## 既存動作への影響確認

- `SubstrateRegistrationWindow.OpenSubstrateInformation` の `RunOnStaThreadAsync` パターンは `CommonUtils` を参照するよう変更されるが動作は同じ
- `Print()` メソッドの `Task.Run + LoadingForm` パターンは今回変更しない（今回の範囲外）
- `GenerateList`（`ExcelServiceClosedXml`）の内部ロジック（ClosedXML・COM Interop の処理内容）は変更なし。try-catch の除去のみ。

---

## テスト観点

- 成績書・リスト・チェックシート生成中に LoadingForm が表示されること
- 生成完了後に LoadingForm が閉じ、ボタンが再有効化されること
- SaveFileDialog / OpenFileDialog / InputDialog キャンセル時に LoadingForm が表示されず処理が中断されること
- 例外発生時に適切なエラーメッセージが表示されること（ボタンも再有効化されること）
- 生成処理中に他の UI 操作（ウィンドウ移動等）がブロックされないこと
- `SubstrateRegistrationWindow` の基板情報ボタンが引き続き正常動作すること
