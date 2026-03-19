# 成績書・リスト・チェックシート生成の非同期化 実装計画

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 成績書・リスト・チェックシートの Excel 生成処理を非同期化し、LoadingForm を表示することで UI フリーズを解消する

**Architecture:** UIダイアログ（SaveFileDialog・OpenFileDialog・InputDialog）を事前に UIスレッドで実行し、ClosedXML の重い処理と COM Interop（Excel 起動）だけを専用 STA スレッドで実行する。LoadingForm は `Load` イベントで STA スレッドを開始するパターンを採用し、ハンドル未作成競合を防ぐ。

**Tech Stack:** C# / .NET 8 / WinForms / ClosedXML / COM Interop (Excel.Interop)

---

## ファイルマップ

| ファイル | 変更種別 | 変更内容 |
|---------|---------|---------|
| `ProductDataBase/Other/CommonUtils.cs` | 変更 | `RunOnStaThreadAsync` を `public static` として追加 |
| `ProductDataBase/Substrate/SubstrateRegistrationWindow.cs` | 変更 | `private static RunOnStaThreadAsync` を削除 → `CommonUtils.RunOnStaThreadAsync` を呼ぶように変更 |
| `ProductDataBase/ExcelService/ExcelServiceClosedXml.cs` | 変更 | ① `GenerateReport` を `PrepareReport`+`ExecuteReport` に分割・旧メソッド削除 ② `GenerateCheckSheet` を `PrepareCheckSheet`+`ExecuteCheckSheet` に分割・旧メソッド削除 ③ `GenerateList` の try-catch 除去 |
| `ProductDataBase/Product/ProductRegistration2Window.cs` | 変更 | `GenerateReport`/`GenerateList`/`GenerateCheckSheet` を `LoadingForm.Load` パターンに変更 |
| `ProductDataBase/Other/HistoryWindow.cs` | 変更 | 同上（3メソッド） |
| `ProductDataBase/Product/SubstrateChange2.cs` | 変更 | `GenerateList` を `LoadingForm.Load` パターンに変更 |

---

## Task 1: CommonUtils に RunOnStaThreadAsync を追加

**Files:**
- Modify: `ProductDataBase/Other/CommonUtils.cs`

> `internal partial class CommonUtils` に `public static Task RunOnStaThreadAsync(Action action)` を追加する。
> COM Interop には STA スレッドが必要。`Task.Run` は MTA のため使えない。
> `TaskCreationOptions.RunContinuationsAsynchronously` により `await` 後の継続が STA スレッドで実行されるのを防ぐ。

- [ ] **Step 1: CommonUtils.cs に RunOnStaThreadAsync を追加する**

`ProductDataBase/Other/CommonUtils.cs` の `internal partial class CommonUtils` 内（`ToMonthCode` の前など）に以下を追加する:

```csharp
// STAスレッドで処理を実行し Task として返す（COM Interop・ダイアログ用）
// RunContinuationsAsynchronously: await後の継続処理がSTAスレッドで実行されるのを防ぐ
public static Task RunOnStaThreadAsync(Action action)
{
    var tcs = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
    var thread = new Thread(() =>
    {
        try
        {
            action();
            tcs.SetResult();
        }
        catch (Exception ex)
        {
            tcs.SetException(ex);
        }
    });
    thread.SetApartmentState(ApartmentState.STA);
    thread.Start();
    return tcs.Task;
}
```

- [ ] **Step 2: ビルドして追加したメソッドにコンパイルエラーがないことを確認する**

```
dotnet build ProductDataBase/ProductDataBase.csproj
```

期待結果: エラーなし

- [ ] **Step 3: コミットする**

```bash
git add ProductDataBase/Other/CommonUtils.cs
git commit -m "feat: CommonUtils に RunOnStaThreadAsync を追加（STA スレッド用ヘルパー）"
```

---

## Task 2: SubstrateRegistrationWindow の RunOnStaThreadAsync を CommonUtils に委譲

**Files:**
- Modify: `ProductDataBase/Substrate/SubstrateRegistrationWindow.cs:629-647`

> `SubstrateRegistrationWindow` にある `private static Task RunOnStaThreadAsync(Action action)` を削除し、`CommonUtils.RunOnStaThreadAsync` を呼ぶようにする。
> 呼び出し箇所: `SubstrateRegistrationWindow.cs` 728行目の `await RunOnStaThreadAsync(OpenSubstrateInformation);`

- [ ] **Step 1: SubstrateRegistrationWindow の private static RunOnStaThreadAsync メソッドを削除する**

`ProductDataBase/Substrate/SubstrateRegistrationWindow.cs` の以下のブロック（約629〜647行）を削除する:

```csharp
// STAスレッドで処理を実行し、Taskとして返す（COM Interop用）
// RunContinuationsAsynchronously: await後の継続処理がSTAスレッドで実行されるのを防ぐ
private static Task RunOnStaThreadAsync(Action action)
{
    var tcs = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
    var thread = new Thread(() =>
    {
        try
        {
            action();
            tcs.SetResult();
        }
        catch (Exception ex)
        {
            tcs.SetException(ex);
        }
    });
    thread.SetApartmentState(ApartmentState.STA);
    thread.Start();
    return tcs.Task;
}
```

- [ ] **Step 2: 呼び出し箇所を CommonUtils.RunOnStaThreadAsync に変更する**

`SubstrateRegistrationWindow.cs` 728行目付近:

```csharp
// 変更前
await RunOnStaThreadAsync(OpenSubstrateInformation);

// 変更後
await CommonUtils.RunOnStaThreadAsync(OpenSubstrateInformation);
```

- [ ] **Step 3: ビルドしてエラーがないことを確認する**

```
dotnet build ProductDataBase/ProductDataBase.csproj
```

期待結果: エラーなし

- [ ] **Step 4: コミットする**

```bash
git add ProductDataBase/Substrate/SubstrateRegistrationWindow.cs
git commit -m "refactor: SubstrateRegistrationWindow の RunOnStaThreadAsync を CommonUtils に委譲"
```

---

## Task 3: GenerateReport を PrepareReport + ExecuteReport に分割

**Files:**
- Modify: `ProductDataBase/ExcelService/ExcelServiceClosedXml.cs:37-72`（`ReportGeneratorClosedXml` クラス内）

> `GenerateReport` メソッドを `PrepareReport`（UIスレッド用）と `ExecuteReport`（STAスレッド用）に分割する。
>
> `PrepareReport` の責務:
> - `LoadConfigWorkbook()` で ConfigReport.xlsm を読む
> - `GetReportConfig()` を呼ぶ（内部で `FindExcelFile` が OpenFileDialog を出す可能性あり）
> - `SaveFileDialog` を表示して保存先パスを取得
> - `OperationCanceledException` (OpenFileDialog/SaveFileDialog のキャンセル) は catch して null を返す
>
> `ExecuteReport` の責務:
> - テンプレートファイルを ClosedXML で開く
> - `PopulateReportSheet()` でデータ書き込み
> - 必要なら `GetUsedSubstrateData()` + `UpdateSubstrateDetailsInExcel()`
> - `ForcePageBreakPreview()` 後に `workbook.SaveAs(savePath)`
> - MessageBox を含めない。例外はすべて呼び出し元に伝播する

- [ ] **Step 1: ExcelServiceClosedXml.cs の ReportGeneratorClosedXml クラスに PrepareReport を追加する**

既存の `GenerateReport` の前後（37行目付近）に以下を追加する:

```csharp
// UIスレッドで呼ぶ: Config読み込み + ファイル選択 + 保存先選択
// OperationCanceledException（ダイアログキャンセル）は null を返す
// その他の例外は呼び出し元に伝播する
public static (string templateFilePath, string savePath)? PrepareReport(string productModel)
{
    try {
        var configWorkbook = LoadConfigWorkbook();
        var reportConfig = GetReportConfig(configWorkbook, productModel);

        // SaveFileDialog を表示して保存先を取得
        var savePath = ShowSaveDialog(reportConfig);
        if (savePath is null) return null; // キャンセル

        return (reportConfig.FilePath, savePath);
    } catch (OperationCanceledException) {
        return null; // OpenFileDialog キャンセル
    }

    static string? ShowSaveDialog(ReportConfigClosedXml config) {
        using SaveFileDialog saveFileDialog = new() {
            Filter = $"Excel Files (*{config.FileExtension})|*{config.FileExtension}|All Files (*.*)|*.*",
            FileName = $"{config.FileName} のコピー{{0}}{config.FileExtension}",
            Title = "保存先を選択してください",
            InitialDirectory = string.IsNullOrWhiteSpace(config.SaveDirectory)
                ? Environment.CurrentDirectory
                : config.SaveDirectory
        };
        return saveFileDialog.ShowDialog() == DialogResult.OK
            ? saveFileDialog.FileName
            : null;
    }
}
```

**注意:** `FileName` の `{0}` プレースホルダーは `ProductNumber` を埋め込む箇所。後述の `ExecuteReport` には `productRegisterWork` が渡されるが、`PrepareReport` は `productModel` しか知らない。`SaveFileDialog` のデフォルト名は現状の実装に合わせて `ProductNumber` を含める必要があるため、`PrepareReport` のシグネチャを `(string productModel, string productNumber)` にするか、または仮のデフォルト名にしておいて後から変更可能にする。

**現状の `SaveReport` の実装（317行目付近）を参照して、`ShowSaveDialog` の `FileName` を正確に合わせること:**

```csharp
// 現状の SaveReport の SaveFileDialog 部分
FileName = $"{fileName} のコピー{productRegisterWork.ProductNumber}{fileExtension}",
```

→ `PrepareReport` に `productNumber` を引数として追加する形が正確。

**修正後のシグネチャ:**

```csharp
public static (string templateFilePath, string savePath)? PrepareReport(
    string productModel, string productNumber)
```

- [ ] **Step 2: ExecuteReport を追加する**

```csharp
// STAスレッドで呼ぶ: ClosedXML 読み書き + SaveAs のみ
// MessageBox を含めない。例外はすべて呼び出し元に伝播する
public static void ExecuteReport(
    ProductMaster productMaster,
    ProductRegisterWork productRegisterWork,
    string templateFilePath,
    string savePath)
{
    using var fs = new FileStream(templateFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
    var reportWorkbook = new XLWorkbook(fs);

    // レポートシートにデータを挿入
    PopulateReportSheet(reportWorkbook, productMaster, productRegisterWork,
        // GetReportConfig を再取得するのではなく、PrepareReport で取得済みの config を渡す設計が理想だが、
        // 現状の PopulateReportSheet は ReportConfigClosedXml を必要とする。
        // Task 3 Step 3 で詳細を検討する。
        );

    if (reportConfig.HasSubstrateInput) {
        var usedSubstrate = GetUsedSubstrateData(productRegisterWork);
        UpdateSubstrateDetailsInExcel(reportWorkbook, reportConfig, usedSubstrate);
    }

    ForcePageBreakPreview(reportWorkbook);
    reportWorkbook.SaveAs(savePath);
}
```

**実装上の重要な課題:** `PopulateReportSheet` と `UpdateSubstrateDetailsInExcel` は `ReportConfigClosedXml` を必要とする。`PrepareReport` でこの情報を取得して `ExecuteReport` に渡す必要がある。

**解決策:** `PrepareReport` の戻り値を以下に拡張する:

```csharp
// PrepareReport の戻り値
public static (string templateFilePath, string savePath, ReportConfigClosedXml config)?
    PrepareReport(string productModel, string productNumber)
```

`ReportConfigClosedXml` はすでに public class なので、呼び出し元がこれを保持して `ExecuteReport` に渡せる。

**修正後の ExecuteReport シグネチャ:**

```csharp
public static void ExecuteReport(
    ProductMaster productMaster,
    ProductRegisterWork productRegisterWork,
    string templateFilePath,
    string savePath,
    ReportConfigClosedXml config)
{
    using var fs = new FileStream(templateFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
    var reportWorkbook = new XLWorkbook(fs);

    PopulateReportSheet(reportWorkbook, productMaster, productRegisterWork, config);

    if (config.HasSubstrateInput) {
        var usedSubstrate = GetUsedSubstrateData(productRegisterWork);
        UpdateSubstrateDetailsInExcel(reportWorkbook, config, usedSubstrate);
    }

    ForcePageBreakPreview(reportWorkbook);
    reportWorkbook.SaveAs(savePath);
}
```

- [ ] **Step 3: 旧 GenerateReport を削除する**

`GenerateReport` メソッド（37〜72行）を丸ごと削除する。

また、`SaveReport` メソッド内の `SaveFileDialog` 部分は `PrepareReport` に移動したため、`SaveReport` は savePath を引数として受け取る形に変更するか、`ExecuteReport` の中で直接 `workbook.SaveAs(savePath)` を呼ぶ形に整理する（`SaveReport` メソッド自体も削除してよい）。

- [ ] **Step 4: ビルドしてエラーがないことを確認する**

```
dotnet build ProductDataBase/ProductDataBase.csproj
```

期待結果: `GenerateReport` を参照していた箇所（`ProductRegistration2Window.cs`, `HistoryWindow.cs`）でコンパイルエラーが発生する。これは正常（Task 6,7 で修正予定）。
`ExcelServiceClosedXml` 自体のエラーがないことを確認する。

- [ ] **Step 5: コミットする**

```bash
git add ProductDataBase/ExcelService/ExcelServiceClosedXml.cs
git commit -m "feat: ReportGeneratorClosedXml を PrepareReport+ExecuteReport に分割"
```

---

## Task 4: GenerateCheckSheet を PrepareCheckSheet + ExecuteCheckSheet に分割

**Files:**
- Modify: `ProductDataBase/ExcelService/ExcelServiceClosedXml.cs:609-673`（`CheckSheetGeneratorClosedXml` クラス内）

> `GenerateCheckSheet` を `PrepareCheckSheet`（UIスレッド用）と `ExecuteCheckSheet`（STAスレッド用）に分割する。
>
> `PrepareCheckSheet` の責務:
> - ConfigCheckSheet.xlsm を ClosedXML で開く（`using` で Dispose を保証）
> - `LoadAndExtractConfig()` で `CheckSheetConfigData` を取得
> - `GetTemperatureAndHumidity()` で InputDialog1 を表示
> - `OperationCanceledException`（InputDialog キャンセル）は catch して null を返す
>
> `ExecuteCheckSheet` の責務:
> - `configData`・`temperature`・`humidity` を引数として受け取る
> - ConfigCheckSheet.xlsm を再度開く（`using` で FileStream + XLWorkbook を管理）
> - `PopulateExcelSheets()`・`HideSheets()`・`SaveWorkbook()`
> - `PrintExcelFile()`（COM Interop）
> - 例外はすべて呼び出し元に伝播する

- [ ] **Step 1: PrepareCheckSheet を追加する**

```csharp
// UIスレッドで呼ぶ: Config読み込み + InputDialog（温度・湿度）
// OperationCanceledException（InputDialog キャンセル）は null を返す
// その他の例外は呼び出し元に伝播する
public static (CheckSheetConfigData configData, string temperature, string humidity)?
    PrepareCheckSheet(ProductMaster productMaster)
{
    var configPath = Path.Combine(Environment.CurrentDirectory,
        "config", "General", "Excel", "ConfigCheckSheet.xlsm");

    if (!File.Exists(configPath)) {
        throw new FileNotFoundException($"設定ファイルが見つかりません: {configPath}");
    }

    CheckSheetConfigData excelData;
    using (var fs = new FileStream(configPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
    using (var workBook = new XLWorkbook(fs)) {
        (_, excelData) = LoadAndExtractConfig(workBook, productMaster);
    }

    try {
        var (temperature, humidity) = GetTemperatureAndHumidity(excelData);
        return (excelData, temperature, humidity);
    } catch (OperationCanceledException) {
        return null; // InputDialog キャンセル
    }
}
```

- [ ] **Step 2: ExecuteCheckSheet を追加する**

```csharp
// STAスレッドで呼ぶ: ClosedXML 処理 + COM Interop（Excel 起動）
// 例外はすべて呼び出し元に伝播する
public static void ExecuteCheckSheet(
    ProductMaster productMaster,
    ProductRegisterWork productRegisterWork,
    CheckSheetConfigData excelData,
    string temperature,
    string humidity)
{
    var configPath = Path.Combine(Environment.CurrentDirectory,
        "config", "General", "Excel", "ConfigCheckSheet.xlsm");
    var temporarilyPath = Path.Combine(Environment.CurrentDirectory,
        "config", "General", "Excel", "temporarilyCheckSheet.xlsm");

    // 日付のフォーマット
    var formattedDate = FormatDate(productRegisterWork.RegDate, excelData.DateFormat);

    // ターゲット Workbook の読み込み
    using var configFs = new FileStream(configPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
    using var configBook = new XLWorkbook(configFs);

    XLWorkbook targetBook;
    FileStream? targetFs = null;

    if (string.IsNullOrWhiteSpace(excelData.BaseFilePath)) {
        targetBook = configBook;
    } else {
        targetFs = new FileStream(excelData.BaseFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        targetBook = new XLWorkbook(targetFs);
    }

    try {
        PopulateExcelSheets(targetBook, productMaster, productRegisterWork,
            temperature, humidity, formattedDate, excelData);
        HideSheets(targetBook, excelData.SheetNames);
        SaveWorkbook(targetBook, temporarilyPath);
    } finally {
        if (targetBook != configBook)
            targetBook.Dispose();
        targetFs?.Dispose();
    }

    PrintExcelFile(temporarilyPath);
}
```

- [ ] **Step 3: 旧 GenerateCheckSheet を削除する**

`GenerateCheckSheet` メソッド（609〜673行）を丸ごと削除する。

- [ ] **Step 4: ビルドしてエラーがないことを確認する**

```
dotnet build ProductDataBase/ProductDataBase.csproj
```

期待結果: `GenerateCheckSheet` を参照していた箇所でコンパイルエラーが発生する（Task 6,7 で修正予定）。`ExcelServiceClosedXml` 自体のエラーがないことを確認する。

- [ ] **Step 5: コミットする**

```bash
git add ProductDataBase/ExcelService/ExcelServiceClosedXml.cs
git commit -m "feat: CheckSheetGeneratorClosedXml を PrepareCheckSheet+ExecuteCheckSheet に分割"
```

---

## Task 5: GenerateList の try-catch を除去する

**Files:**
- Modify: `ProductDataBase/ExcelService/ExcelServiceClosedXml.cs:340-378`（`ListGeneratorClosedXml` クラス内）

> `GenerateList` 内の `try-catch`（例外を握りつぶして MessageBox を出していた）を除去し、例外を呼び出し元に伝播するよう変更する。
> 内部ロジック（ClosedXML・COM Interop の処理）は変更しない。

- [ ] **Step 1: GenerateList の try-catch ブロックを除去する**

```csharp
// 変更前
public static void GenerateList(ProductMaster productMaster, ProductRegisterWork productRegisterWork) {
    try {
        // ... 処理 ...
    } catch (Exception ex) {
        MessageBox.Show(ex.Message, $"[{...}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
    }
}

// 変更後（try-catch を除去。処理内容はそのまま）
public static void GenerateList(ProductMaster productMaster, ProductRegisterWork productRegisterWork) {
    // ... 処理（変更なし）...
}
```

- [ ] **Step 2: ビルドしてエラーがないことを確認する**

```
dotnet build ProductDataBase/ProductDataBase.csproj
```

期待結果: エラーなし（GenerateList 自体はシグネチャ変更なし）

- [ ] **Step 3: コミットする**

```bash
git add ProductDataBase/ExcelService/ExcelServiceClosedXml.cs
git commit -m "refactor: GenerateList の try-catch を除去し例外を呼び出し元へ伝播"
```

---

## Task 6: ProductRegistration2Window の 3メソッドを非同期化

**Files:**
- Modify: `ProductDataBase/Product/ProductRegistration2Window.cs:1025-1059`

> `GenerateReport`（1025行）・`GenerateList`（1037行）・`GenerateCheckSheet`（1049行）を `LoadingForm.Load` イベントパターンに変更する。
>
> **共通パターン:**
> 1. `PrepareXxx` を try-catch で囲んで呼ぶ（キャンセルなら return、エラーなら MessageBox→return）
> 2. ボタン無効化
> 3. `LoadingForm.Load` に `async void` ハンドラを登録（try-await-catch-finally(Close)）
> 4. `loadingForm.ShowDialog(this)` でブロック
> 5. ボタン再有効化、結果に応じてメッセージ表示

- [ ] **Step 1: GenerateReport を変更する**

```csharp
// 登録済み製品情報をもとにExcel成績書を生成する
private void GenerateReport()
{
    // [UIスレッド] 前処理: Config読み込み + ファイル選択 + 保存先選択
    (string templateFilePath, string savePath, ReportGeneratorClosedXml.ReportConfigClosedXml config)? prepared;
    try {
        prepared = ExcelServiceClosedXml.ReportGeneratorClosedXml.PrepareReport(
            _productMaster.ProductModel, _productRegisterWork.ProductNumber);
    } catch (Exception ex) {
        MessageBox.Show(ex.Message,
            $"[{nameof(GenerateReport)}]エラー",
            MessageBoxButtons.OK, MessageBoxIcon.Error);
        return;
    }
    if (prepared is null) return; // キャンセル

    GenerateReportButton.Enabled = false;
    Exception? taskException = null;
    using var loadingForm = new LoadingForm();
    loadingForm.Load += async (_, _) => {
        // ※ try は await を含む全体を包む。await より前に同期処理を追加しないこと。
        try {
            await CommonUtils.RunOnStaThreadAsync(() =>
                ExcelServiceClosedXml.ReportGeneratorClosedXml.ExecuteReport(
                    _productMaster, _productRegisterWork,
                    prepared.Value.templateFilePath,
                    prepared.Value.savePath,
                    prepared.Value.config));
        } catch (Exception ex) {
            taskException = ex;
        } finally {
            loadingForm.Close();
        }
    };
    loadingForm.ShowDialog(this); // Load イベント完了後にリターン

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

- [ ] **Step 2: GenerateList を変更する**

```csharp
// 登録済み製品情報をもとにExcel製品一覧を生成する
private void GenerateList()
{
    // 前処理なし（UIダイアログ不要）
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

**注意:** `ProductRegistration2Window` にはリストボタンが `SubstrateListPrintButton` という名前。`GenerateListButton` ではなく `SubstrateListPrintButton.Enabled` を使う（既存コードを確認して正しいコントロール名を使うこと）。

- [ ] **Step 3: GenerateCheckSheet を変更する**

```csharp
// 登録済み製品情報をもとにExcelチェックシートを生成する
private void GenerateCheckSheet()
{
    (CheckSheetGeneratorClosedXml.CheckSheetConfigData configData, string temperature, string humidity)? prepared;
    try {
        prepared = ExcelServiceClosedXml.CheckSheetGeneratorClosedXml.PrepareCheckSheet(
            _productMaster);
    } catch (Exception ex) {
        MessageBox.Show(ex.Message,
            $"[{nameof(GenerateCheckSheet)}]エラー",
            MessageBoxButtons.OK, MessageBoxIcon.Error);
        return;
    }
    if (prepared is null) return; // キャンセル

    CheckSheetPrintButton.Enabled = false;
    Exception? taskException = null;
    using var loadingForm = new LoadingForm();
    loadingForm.Load += async (_, _) => {
        try {
            await CommonUtils.RunOnStaThreadAsync(() =>
                ExcelServiceClosedXml.CheckSheetGeneratorClosedXml.ExecuteCheckSheet(
                    _productMaster, _productRegisterWork,
                    prepared.Value.configData,
                    prepared.Value.temperature,
                    prepared.Value.humidity));
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

- [ ] **Step 4: 旧 GenerateReport/GenerateList/GenerateCheckSheet の `finally { Cursor.Current = Cursors.WaitCursor; }` バグも一緒に除去されていることを確認する**

（既存の `GenerateCheckSheet` の finally は `Cursors.Default` ではなく `Cursors.WaitCursor` になっているバグあり。新しい実装では `Cursor` は使わない）

- [ ] **Step 5: ビルドしてエラーがないことを確認する**

```
dotnet build ProductDataBase/ProductDataBase.csproj
```

期待結果: エラーなし

- [ ] **Step 6: コミットする**

```bash
git add ProductDataBase/Product/ProductRegistration2Window.cs
git commit -m "feat: ProductRegistration2Window の成績書・リスト・チェックシート生成を非同期化"
```

---

## Task 7: HistoryWindow の 3メソッドを非同期化

**Files:**
- Modify: `ProductDataBase/Other/HistoryWindow.cs:1145-1217`

> `HistoryWindow` の `GenerateReport`・`GenerateList`・`GenerateCheckSheet` を同じパターンで変更する。
>
> **HistoryWindow 固有の注意点:**
> - 各メソッドの先頭で DataGridView の選択行からデータを `_productRegisterWork` と `_productMaster` にセットしている（この前処理は UIスレッドで行うため変更不要）
> - ボタン名: `GenerateReportButton`・`GenerateListButton`・`GenerateCheckSheetButton`
> - `SetBottomControlsEnabled(bool)` メソッドが存在するが、今回は個別のボタンのみを無効化する（他のボタンはそのままにする）

- [ ] **Step 1: GenerateReport を変更する**

既存の `GenerateReport`（1145〜1167行）を以下に置き換える:

```csharp
// 選択行の製品情報をセットしてExcel製造報告書を生成する
private void GenerateReport()
{
    if (DataBaseDataGridView.CurrentCell is null && DataBaseDataGridView.SelectedCells.Count <= 0) { return; }
    var selectRow = DataBaseDataGridView.SelectedCells[0].RowIndex;
    _productRegisterWork.RowID = Convert.ToInt32(DataBaseDataGridView.Rows[selectRow].Cells["ID"].Value);
    _productRegisterWork.OrderNumber = DataBaseDataGridView.Rows[selectRow].Cells["OrderNumber"].Value.ToString() ?? string.Empty;
    _productRegisterWork.ProductNumber = DataBaseDataGridView.Rows[selectRow].Cells["ProductNumber"].Value.ToString() ?? string.Empty;
    _productMaster.ProductModel = DataBaseDataGridView.Rows[selectRow].Cells["ProductModel"].Value.ToString() ?? string.Empty;
    _productRegisterWork.Quantity = int.TryParse(DataBaseDataGridView.Rows[selectRow].Cells["Quantity"].Value?.ToString(), out var quantity) ? quantity : 0;
    _productRegisterWork.RegDate = DataBaseDataGridView.Rows[selectRow].Cells["RegDate"].Value.ToString() ?? string.Empty;
    _productRegisterWork.SerialFirst = DataBaseDataGridView.Rows[selectRow].Cells["SerialFirst"].Value.ToString() ?? string.Empty;
    _productRegisterWork.SerialLast = DataBaseDataGridView.Rows[selectRow].Cells["SerialLast"].Value.ToString() ?? string.Empty;

    (string templateFilePath, string savePath, ExcelServiceClosedXml.ReportGeneratorClosedXml.ReportConfigClosedXml config)? prepared;
    try {
        prepared = ExcelServiceClosedXml.ReportGeneratorClosedXml.PrepareReport(
            _productMaster.ProductModel, _productRegisterWork.ProductNumber);
    } catch (Exception ex) {
        MessageBox.Show(ex.Message,
            $"[{nameof(GenerateReport)}]エラー",
            MessageBoxButtons.OK, MessageBoxIcon.Error);
        return;
    }
    if (prepared is null) return;

    GenerateReportButton.Enabled = false;
    Exception? taskException = null;
    using var loadingForm = new LoadingForm();
    loadingForm.Load += async (_, _) => {
        try {
            await CommonUtils.RunOnStaThreadAsync(() =>
                ExcelServiceClosedXml.ReportGeneratorClosedXml.ExecuteReport(
                    _productMaster, _productRegisterWork,
                    prepared.Value.templateFilePath,
                    prepared.Value.savePath,
                    prepared.Value.config));
        } catch (Exception ex) {
            taskException = ex;
        } finally {
            loadingForm.Close();
        }
    };
    loadingForm.ShowDialog(this);

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

- [ ] **Step 2: GenerateList を変更する**

```csharp
// 選択行の製品情報をセットしてExcel製品一覧を生成する
private void GenerateList()
{
    if (DataBaseDataGridView.CurrentCell is null && DataBaseDataGridView.SelectedCells.Count <= 0) { return; }
    var selectRow = DataBaseDataGridView.SelectedCells[0].RowIndex;
    _productRegisterWork.RowID = Convert.ToInt32(DataBaseDataGridView.Rows[selectRow].Cells["ID"].Value);
    _productRegisterWork.OrderNumber = DataBaseDataGridView.Rows[selectRow].Cells["OrderNumber"].Value.ToString() ?? string.Empty;
    _productRegisterWork.ProductNumber = DataBaseDataGridView.Rows[selectRow].Cells["ProductNumber"].Value.ToString() ?? string.Empty;
    _productMaster.ProductModel = DataBaseDataGridView.Rows[selectRow].Cells["ProductModel"].Value.ToString() ?? string.Empty;
    _productRegisterWork.Quantity = Convert.ToInt32(DataBaseDataGridView.Rows[selectRow].Cells["Quantity"].Value);
    _productRegisterWork.RegDate = DataBaseDataGridView.Rows[selectRow].Cells["RegDate"].Value.ToString() ?? string.Empty;
    _productRegisterWork.SerialFirst = DataBaseDataGridView.Rows[selectRow].Cells["SerialFirst"].Value.ToString() ?? string.Empty;
    _productRegisterWork.SerialLast = DataBaseDataGridView.Rows[selectRow].Cells["SerialLast"].Value.ToString() ?? string.Empty;
    _productRegisterWork.Comment = DataBaseDataGridView.Rows[selectRow].Cells["Comment"].Value.ToString() ?? string.Empty;

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

- [ ] **Step 3: GenerateCheckSheet を変更する**

```csharp
// 選択行の製品情報をセットしてExcelチェックシートを生成する
private void GenerateCheckSheet()
{
    if (DataBaseDataGridView.CurrentCell is null && DataBaseDataGridView.SelectedCells.Count <= 0) { return; }
    var selectRow = DataBaseDataGridView.SelectedCells[0].RowIndex;
    _productRegisterWork.ProductNumber = DataBaseDataGridView.Rows[selectRow].Cells["ProductNumber"].Value.ToString() ?? string.Empty;
    _productRegisterWork.OrderNumber = DataBaseDataGridView.Rows[selectRow].Cells["OrderNumber"].Value.ToString() ?? string.Empty;
    _productMaster.ProductModel = DataBaseDataGridView.Rows[selectRow].Cells["ProductModel"].Value.ToString() ?? string.Empty;
    _productRegisterWork.Quantity = int.TryParse(DataBaseDataGridView.Rows[selectRow].Cells["Quantity"].Value?.ToString(), out var quantity) ? quantity : 0;
    _productRegisterWork.RegDate = DataBaseDataGridView.Rows[selectRow].Cells["RegDate"].Value.ToString() ?? string.Empty;
    _productRegisterWork.SerialFirst = DataBaseDataGridView.Rows[selectRow].Cells["SerialFirst"].Value.ToString() ?? string.Empty;
    _productRegisterWork.SerialLast = DataBaseDataGridView.Rows[selectRow].Cells["SerialLast"].Value.ToString() ?? string.Empty;

    (ExcelServiceClosedXml.CheckSheetGeneratorClosedXml.CheckSheetConfigData configData, string temperature, string humidity)? prepared;
    try {
        prepared = ExcelServiceClosedXml.CheckSheetGeneratorClosedXml.PrepareCheckSheet(
            _productMaster);
    } catch (Exception ex) {
        MessageBox.Show(ex.Message,
            $"[{nameof(GenerateCheckSheet)}]エラー",
            MessageBoxButtons.OK, MessageBoxIcon.Error);
        return;
    }
    if (prepared is null) return;

    GenerateCheckSheetButton.Enabled = false;
    Exception? taskException = null;
    using var loadingForm = new LoadingForm();
    loadingForm.Load += async (_, _) => {
        try {
            await CommonUtils.RunOnStaThreadAsync(() =>
                ExcelServiceClosedXml.CheckSheetGeneratorClosedXml.ExecuteCheckSheet(
                    _productMaster, _productRegisterWork,
                    prepared.Value.configData,
                    prepared.Value.temperature,
                    prepared.Value.humidity));
        } catch (Exception ex) {
            taskException = ex;
        } finally {
            loadingForm.Close();
        }
    };
    loadingForm.ShowDialog(this);

    GenerateCheckSheetButton.Enabled = true;
    if (taskException is not null) {
        MessageBox.Show(taskException.Message,
            $"[{nameof(GenerateCheckSheet)}]エラー",
            MessageBoxButtons.OK, MessageBoxIcon.Error);
    }
}
```

- [ ] **Step 4: ビルドしてエラーがないことを確認する**

```
dotnet build ProductDataBase/ProductDataBase.csproj
```

期待結果: エラーなし

- [ ] **Step 5: コミットする**

```bash
git add ProductDataBase/Other/HistoryWindow.cs
git commit -m "feat: HistoryWindow の成績書・リスト・チェックシート生成を非同期化"
```

---

## Task 8: SubstrateChange2 の GenerateList を非同期化

**Files:**
- Modify: `ProductDataBase/Product/SubstrateChange2.cs:541-553`

> `SubstrateChange2.GenerateList`（541行）を `LoadingForm.Load` パターンに変更する。
> `SubstrateListPrintButton` が対応ボタン名（618行の `SubstrateListPrintButton_Click` を参照）。

- [ ] **Step 1: GenerateList を変更する**

```csharp
// 基板変更済み製品情報をもとにExcel製品一覧を生成する
private void GenerateList()
{
    SubstrateListPrintButton.Enabled = false;
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

    SubstrateListPrintButton.Enabled = true;
    if (taskException is not null) {
        MessageBox.Show(taskException.Message,
            $"[{nameof(GenerateList)}]エラー",
            MessageBoxButtons.OK, MessageBoxIcon.Error);
    }
}
```

- [ ] **Step 2: ビルドしてエラーがないことを確認する**

```
dotnet build ProductDataBase/ProductDataBase.csproj
```

期待結果: エラーなし

- [ ] **Step 3: コミットする**

```bash
git add ProductDataBase/Product/SubstrateChange2.cs
git commit -m "feat: SubstrateChange2 のリスト生成を非同期化"
```

---

## Task 9: 動作確認

> 自動テストがないため、手動スモークテストを実施する。

- [ ] **Step 1: アプリケーションを起動してビルドが通ることを確認する**

```
dotnet build ProductDataBase/ProductDataBase.csproj --configuration Release
```

期待結果: エラーなし、警告のみ

- [ ] **Step 2: 成績書生成の動作確認**

1. `ProductRegistration2Window` で製品登録後、「成績書作成」ボタンをクリック
2. LoadingForm（GIFアニメーション）が表示されることを確認
3. SaveFileDialog が **LoadingForm 表示前**（前処理として）表示されることを確認
4. 保存完了後に「成績書が正常に生成されました。」MessageBox が出ることを確認
5. ボタンが再有効化されていることを確認

- [ ] **Step 3: リスト生成の動作確認**

1. 「リスト作成」ボタンをクリック
2. LoadingForm が表示されることを確認
3. ClosedXML 処理後に Excel が起動することを確認
4. ボタンが再有効化されていることを確認

- [ ] **Step 4: チェックシート生成の動作確認**

1. 「チェックシート」ボタンをクリック
2. 温度・湿度 InputDialog が **LoadingForm 表示前**（前処理として）表示されることを確認
3. キャンセル時は LoadingForm が表示されず終了することを確認
4. 入力後に LoadingForm が表示され、Excel が起動することを確認

- [ ] **Step 5: HistoryWindow でも同様に動作することを確認する**

- [ ] **Step 6: SubstrateRegistrationWindow の基板情報ボタンが引き続き正常動作することを確認する**

(`RunOnStaThreadAsync` を CommonUtils に移動した影響がないことを確認)

- [ ] **Step 7: 最終コミット**

```bash
git add -u
git commit -m "docs: 実装完了 - 成績書・リスト・チェックシート生成の非同期化"
```

---

## 実装上の注意事項

### PrepareReport の戻り値に ReportConfigClosedXml を含める

`ExecuteReport` では `PopulateReportSheet` と `UpdateSubstrateDetailsInExcel` が `ReportConfigClosedXml` を必要とする。`PrepareReport` の戻り値をタプルで `config` も含める形にする（Task 3 参照）。

### GenerateList のボタン名

- `ProductRegistration2Window`: `SubstrateListPrintButton`（`GenerateListButton` ではない）
- `HistoryWindow`: `GenerateListButton`
- `SubstrateChange2`: `SubstrateListPrintButton`

### CheckSheetConfigData のアクセス修飾子

現状 `CheckSheetGeneratorClosedXml` 内の `CheckSheetConfigData` クラスは `public` であることを確認する（呼び出し元がこの型を参照するため）。

### async void Load ハンドラの注意

`try` ブロックは必ず `await` の前から始めること。`await` より前に同期処理を追加すると、その例外は `taskException` に保存されず `SynchronizationContext` に投げられる。
