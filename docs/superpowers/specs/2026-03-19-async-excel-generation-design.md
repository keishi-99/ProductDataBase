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

基板情報ボタン（`SubstrateRegistrationWindow`）では `RunOnStaThreadAsync` + ボタン無効化のパターンで既に非同期化済み。印刷処理（`Print()`）では `Task.Run + LoadingForm.ShowDialog()` パターンで非同期化済み。今回はこれらのパターンを 3 つの Excel 生成処理に適用する。

---

## 技術制約

- **COM Interop（Excel 起動）は STA スレッドが必要。** `Task.Run` のスレッドプールスレッドは MTA のため使用不可。専用 STA スレッド（`RunOnStaThreadAsync`）を使う。
- **`SaveFileDialog`・`OpenFileDialog` も STA スレッド要件あり。** ただし LoadingForm 表示中に別スレッドからダイアログを出すと UX が混乱するため、これらは事前に UI スレッドで実行する。
- **`MessageBox` はどのスレッドでも動作する。**

---

## 設計方針（アプローチ B: ダイアログを前処理として分離）

各処理を以下の 2 段階に分割する。

```
[UIスレッド] 前処理ダイアログ（あれば）
    ↓
[UIスレッド] ボタン無効化 + LoadingForm 表示開始
    ↓
[STAスレッド] 重い処理（ClosedXML + COM Interop）
    ↓
[UIスレッド] LoadingForm 閉じる + ボタン再有効化 + 完了通知
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
- `GetReportConfig(configWorkbook, productModel)` を呼ぶ（`FindExcelFile` 含む。ファイル複数一致時は `OpenFileDialog` が出る）
- `SaveReport` 内にあった `SaveFileDialog` を呼んで保存先パスを取得
- 戻り値: `(string templateFilePath, string savePath)?`（キャンセル時は `null`）

**新メソッド `ExecuteReport`（STAスレッドで呼ぶ）:**
- テンプレートファイルを ClosedXML で開く
- `PopulateReportSheet()` を呼ぶ
- 必要に応じて `GetUsedSubstrateData()` + `UpdateSubstrateDetailsInExcel()` を呼ぶ
- `ForcePageBreakPreview()` を呼ぶ
- `workbook.SaveAs(savePath)` で保存（MessageBox は含めない）

**旧 `GenerateReport` の除去:** 呼び出し元が `PrepareReport` + `ExecuteReport` を使う形に変更されるため、`GenerateReport` は削除する。

---

#### 2-2. `CheckSheetGeneratorClosedXml` の分割

**新メソッド `PrepareCheckSheet`（UIスレッドで呼ぶ）:**
- ConfigCheckSheet.xlsm を ClosedXML で開く
- `LoadAndExtractConfig(workBook, productMaster)` を呼んで `CheckSheetConfigData` を取得
- `GetTemperatureAndHumidity(excelData)` を呼んで温度・湿度を取得（InputDialog1 が出る）
- 戻り値: `(CheckSheetConfigData configData, string temperature, string humidity)?`（キャンセル時は `null`）

**新メソッド `ExecuteCheckSheet`（STAスレッドで呼ぶ）:**
- `configData`、`temperature`、`humidity` を引数として受け取る
- `FormatDate()`
- `LoadWorkbookOrDefault()` でターゲット Workbook を読み込む
- `PopulateExcelSheets()`
- `HideSheets()`
- `SaveWorkbook()`
- `PrintExcelFile()`（COM Interop）

**旧 `GenerateCheckSheet` の除去。**

---

#### 2-3. `ListGeneratorClosedXml` の変更なし

`GenerateList` はそのまま使い、呼び出し側から `RunOnStaThreadAsync` で包む。

---

### 3. 呼び出し側の変更（`ProductRegistration2Window`, `HistoryWindow`, `SubstrateChange2`）

各フォームの `GenerateReport`・`GenerateList`・`GenerateCheckSheet` メソッドを `async void` に変更し、以下のパターンで統一する。

#### 成績書の例

```csharp
private async void GenerateReport()
{
    // [UIスレッド] 前処理（Config 読み込み + ファイル選択 + 保存先選択）
    var prepared = ExcelServiceClosedXml.ReportGeneratorClosedXml.PrepareReport(
        _productMaster.ProductModel);
    if (prepared is null) return; // キャンセル

    // [UIスレッド] ボタン無効化 + LoadingForm 表示
    GenerateReportButton.Enabled = false;
    using var loadingForm = new LoadingForm();
    var task = CommonUtils.RunOnStaThreadAsync(() =>
        ExcelServiceClosedXml.ReportGeneratorClosedXml.ExecuteReport(
            _productMaster, _productRegisterWork,
            prepared.Value.templateFilePath, prepared.Value.savePath));
    _ = task.ContinueWith(_ => loadingForm.Invoke(loadingForm.Close));
    loadingForm.ShowDialog(this);

    // [UIスレッド] 完了・エラー処理
    try {
        await task;
        MessageBox.Show("成績書が正常に生成されました。", "完了",
            MessageBoxButtons.OK, MessageBoxIcon.Information);
    } catch (OperationCanceledException) {
        // キャンセルは無視
    } catch (Exception ex) {
        MessageBox.Show(ex.Message,
            $"[{nameof(GenerateReport)}]エラー",
            MessageBoxButtons.OK, MessageBoxIcon.Error);
    } finally {
        GenerateReportButton.Enabled = true;
    }
}
```

#### リストの例

```csharp
private async void GenerateList()
{
    GenerateListButton.Enabled = false;
    using var loadingForm = new LoadingForm();
    var task = CommonUtils.RunOnStaThreadAsync(() =>
        ExcelServiceClosedXml.ListGeneratorClosedXml.GenerateList(
            _productMaster, _productRegisterWork));
    _ = task.ContinueWith(_ => loadingForm.Invoke(loadingForm.Close));
    loadingForm.ShowDialog(this);

    try {
        await task;
    } catch (Exception ex) {
        MessageBox.Show(ex.Message,
            $"[{nameof(GenerateList)}]エラー",
            MessageBoxButtons.OK, MessageBoxIcon.Error);
    } finally {
        GenerateListButton.Enabled = true;
    }
}
```

#### チェックシートの例

```csharp
private async void GenerateCheckSheet()
{
    var prepared = ExcelServiceClosedXml.CheckSheetGeneratorClosedXml.PrepareCheckSheet(
        _productMaster);
    if (prepared is null) return;

    CheckSheetPrintButton.Enabled = false;
    using var loadingForm = new LoadingForm();
    var task = CommonUtils.RunOnStaThreadAsync(() =>
        ExcelServiceClosedXml.CheckSheetGeneratorClosedXml.ExecuteCheckSheet(
            _productMaster, _productRegisterWork,
            prepared.Value.configData,
            prepared.Value.temperature, prepared.Value.humidity));
    _ = task.ContinueWith(_ => loadingForm.Invoke(loadingForm.Close));
    loadingForm.ShowDialog(this);

    try {
        await task;
    } catch (Exception ex) {
        MessageBox.Show(ex.Message,
            $"[{nameof(GenerateCheckSheet)}]エラー",
            MessageBoxButtons.OK, MessageBoxIcon.Error);
    } finally {
        CheckSheetPrintButton.Enabled = true;
    }
}
```

---

## 変更ファイル一覧

| ファイル | 変更内容 |
|---------|---------|
| `Other/CommonUtils.cs` | `RunOnStaThreadAsync` を追加 |
| `Substrate/SubstrateRegistrationWindow.cs` | `RunOnStaThreadAsync` を削除し `CommonUtils` を参照 |
| `ExcelService/ExcelServiceClosedXml.cs` | `GenerateReport` → `PrepareReport`+`ExecuteReport` に分割; `GenerateCheckSheet` → `PrepareCheckSheet`+`ExecuteCheckSheet` に分割 |
| `Product/ProductRegistration2Window.cs` | `GenerateReport`・`GenerateList`・`GenerateCheckSheet` を async 化 |
| `Other/HistoryWindow.cs` | 同上（3メソッド） |
| `Product/SubstrateChange2.cs` | `GenerateList` を async 化 |

---

## エラーハンドリング

- `ExecuteReport` / `ExecuteCheckSheet` / `GenerateList` は例外をそのままスローし、呼び出し側の `catch` でハンドリングする（`ExcelServiceClosedXml` 内で MessageBox を使わない）
- `OperationCanceledException`（SaveFileDialog キャンセル等）は呼び出し側で握りつぶす
- `PrepareReport`・`PrepareCheckSheet` 内での例外はそのまま呼び出し側に伝播

---

## 既存動作への影響確認

- `SubstrateRegistrationWindow.OpenSubstrateInformation` のパターン（`RunOnStaThreadAsync`）は引き続き同じ動作
- `Print()` メソッドの `Task.Run + LoadingForm` パターンは変更なし
- `GenerateList`（`ExcelServiceClosedXml`）の内部ロジックは変更なし
- `SubstrateChange2.GenerateList` は `HistoryWindow` と同じパターンに統一

---

## テスト観点

- 成績書・リスト・チェックシート生成中に LoadingForm が表示されること
- 生成完了後に LoadingForm が閉じ、ボタンが再有効化されること
- SaveFileDialog/InputDialog キャンセル時に処理が中断されること
- 例外発生時に適切なエラーメッセージが表示されること
- 生成処理中に他の UI 操作（ウィンドウ移動等）がブロックされないこと
