using ClosedXML.Excel;
using Dapper;
using Microsoft.Data.Sqlite;
using ProductDatabase.Models;
using ProductDatabase.Other;
using ZXing;
using ZXing.QrCode;
using ZXing.QrCode.Internal;
using ZXing.Windows.Compatibility;
using static ProductDatabase.Data.ProductRepository;

namespace ProductDatabase.ExcelService {
    internal class ExcelServiceClosedXml {
        // 成績書
        public static class ReportGeneratorClosedXml {
            // 成績書生成データを保持するクラス
            public class ReportConfigClosedXml {
                public required IXLWorksheet Sheet { get; set; }
                public required int ResultRow { get; set; }
                public required string DirectoryPath { get; set; }
                public required string FilePath { get; set; }
                public required string FileName { get; set; }
                public required string FileExtension { get; set; }
                public required string SearchFileName { get; set; }
                public required string SheetName { get; set; }
                public string? ProductNumberRange { get; set; }
                public string? OrderNumberRange { get; set; }
                public string? QuantityRange { get; set; }
                public string? SerialFirstRange { get; set; }
                public string? SerialLastRange { get; set; }
                public string? ProductModelRange { get; set; }
                public string? SaveDirectory { get; set; }
                public bool HasSubstrateInput { get; set; }
            }

            // UIスレッドで呼ぶ: Config読み込み + ファイル選択 + 保存先選択
            // OperationCanceledException（ダイアログキャンセル）は null を返す
            // その他の例外は呼び出し元に伝播する
            public static (string templateFilePath, string savePath, ReportConfigClosedXml config)?
                PrepareReport(string productModel, string productNumber) {
                try {
                    var configWorkbook = LoadConfigWorkbook();
                    var reportConfig = GetReportConfig(configWorkbook, productModel);

                    var savePath = ShowSaveDialog(reportConfig, productNumber);
                    if (savePath is null) return null; // SaveFileDialog キャンセル

                    return (reportConfig.FilePath, savePath, reportConfig);
                } catch (OperationCanceledException) {
                    return null; // OpenFileDialog キャンセル
                }

                static string? ShowSaveDialog(ReportConfigClosedXml config, string productNumber) {
                    if (!Directory.Exists(config.SaveDirectory)) {
                        MessageBox.Show($"設定されている保存先が見つかりませんでした。\r\n{config.SaveDirectory}", "ファイル選択", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }

                    using SaveFileDialog saveFileDialog = new() {
                        Filter = $"Excel Files (*{config.FileExtension})|*{config.FileExtension}|All Files (*.*)|*.*",
                        FileName = $"{config.FileName} のコピー{productNumber}{config.FileExtension}",
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

            // STAスレッドで呼ぶ: ClosedXML 読み書き + SaveAs のみ
            // MessageBox を含めない。例外はすべて呼び出し元に伝播する
            public static void ExecuteReport(
                ProductMaster productMaster,
                ProductRegisterWork productRegisterWork,
                string templateFilePath,
                string savePath,
                ReportConfigClosedXml config) {
                if (!File.Exists(templateFilePath)) {
                    throw new FileNotFoundException($"設定ファイルが見つかりません: {templateFilePath}");
                }

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

            // 設定Excelワークブックを読み込む
            private static XLWorkbook LoadConfigWorkbook() {
                try {
                    var configPath = Path.Combine(Environment.CurrentDirectory, "config", "General", "Excel", "ConfigReport.xlsm");

                    if (!File.Exists(configPath)) {
                        throw new FileNotFoundException($"設定ファイルが見つかりません: {configPath}");
                    }

                    // Excelが開かれていても読み取れるようにFileShare.ReadWrite指定
                    using var fs = new FileStream(configPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

                    // FileStreamからXLWorkbookを読み込む
                    return new XLWorkbook(fs);
                } catch (Exception ex) {
                    throw new Exception($"設定ファイルの読み込み中にエラーが発生しました: {ex.Message}", ex);
                }
            }

            // 設定ワークブックから製品に対応する設定を抽出する
            private static ReportConfigClosedXml GetReportConfig(XLWorkbook configWorkbook, string productModel) {
                var configSheet = configWorkbook.Worksheet("Sheet1")
                    ?? throw new Exception("設定ファイルに 'Sheet1' が見つかりません。");

                // --- 品目番号の検索 ---
                var targetCell = configSheet
                    .Column(1) // 1列目を取得
                    .CellsUsed() // 値が入っているセルのみ
                    .FirstOrDefault(c => c.GetValue<string>() == productModel)
                    ?? throw new Exception($"Configに品目番号:[{productModel}]が見つかりません。");

                var resultRow = targetCell.Address.RowNumber;

                // --- シート名 ---
                var sheetName = configSheet.Cell(resultRow, 5).GetString();
                if (string.IsNullOrEmpty(sheetName)) {
                    throw new Exception($"設定ファイルのシート名が空です。");
                }

                // --- パスとファイル名の取得 ---
                var directoryPath = configSheet.Cell(resultRow, 3).GetString().Trim('"');
                if (string.IsNullOrEmpty(directoryPath)) {
                    directoryPath = configSheet.Cell(2, 3).GetString().Trim('"');
                }

                if (!Directory.Exists(directoryPath)) {
                    throw new DirectoryNotFoundException($"指定されたフォルダが存在しません: {directoryPath}");
                }

                var saveDirectory = configSheet.Cell(resultRow, 12).GetString().Trim('"');
                if (string.IsNullOrEmpty(saveDirectory)) {
                    saveDirectory = configSheet.Cell(2, 12).GetString().Trim('"');
                }

                var searchName = configSheet.Cell(resultRow, 4).GetString().Trim('"');
                if (string.IsNullOrEmpty(searchName)) {
                    throw new DirectoryNotFoundException($"検索ファイル名が空です。");
                }

                // --- ファイル検索 ---
                var filePath = FindExcelFile(directoryPath, searchName);
                var fileName = Path.GetFileNameWithoutExtension(filePath);
                var fileExtension = Path.GetExtension(filePath).ToLowerInvariant();

                // ---  基板入力があるか確認 ---
                int startColumn = 12;
                int endColumn = 27;

                var targetRowCells = configSheet.Row(resultRow).Cells(startColumn, endColumn);
                bool hasValue = targetRowCells.Any(c => !string.IsNullOrWhiteSpace(c.GetValue<string>()));

                // --- ReportConfig を構築 ---
                return new ReportConfigClosedXml {
                    Sheet = configSheet,
                    ResultRow = resultRow,
                    DirectoryPath = directoryPath,
                    FilePath = filePath,
                    FileName = fileName,
                    FileExtension = fileExtension,
                    SearchFileName = searchName,
                    SheetName = sheetName,
                    ProductNumberRange = configSheet.Cell(resultRow, 6).GetString(),
                    OrderNumberRange = configSheet.Cell(resultRow, 7).GetString(),
                    QuantityRange = configSheet.Cell(resultRow, 8).GetString(),
                    SerialFirstRange = configSheet.Cell(resultRow, 9).GetString(),
                    SerialLastRange = configSheet.Cell(resultRow, 10).GetString(),
                    ProductModelRange = configSheet.Cell(resultRow, 11).GetString(),
                    SaveDirectory = saveDirectory,
                    HasSubstrateInput = hasValue,
                };
            }
            // 検索名でフォルダを検索し一致するExcelファイルのパスを返す（複数一致時はダイアログで選択）
            private static string FindExcelFile(string directoryPath, string searchName) {
                var filePaths = string.IsNullOrEmpty(searchName)
                    ? []
                    : Directory.GetFiles(directoryPath, $"{searchName}*", SearchOption.AllDirectories);

                if (filePaths.Length == 1) {
                    return Path.GetFullPath(filePaths[0]);
                }

                if (filePaths.Length > 1) {
                    MessageBox.Show(
                        "複数のファイルが見つかりました。1つ選択してください。",
                        "ファイル選択",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                }

                var filePath = ShowExcelFileDialog(directoryPath)
                    ?? throw new OperationCanceledException("ファイル選択がキャンセルされました。");

                return filePath;

                static string? ShowExcelFileDialog(string initialDirectory) {
                    using var dialog = new OpenFileDialog {
                        InitialDirectory = initialDirectory,
                        Filter = "Excel ファイル|*.xlsx;*.xlsm|すべてのファイル (*.*)|*.*",
                        Multiselect = false
                    };

                    return dialog.ShowDialog() == DialogResult.OK
                        ? dialog.FileName
                        : null;
                }
            }

            // レポートシートに製品情報を挿入する
            private static void PopulateReportSheet(XLWorkbook reportWorkbook, ProductMaster productMaster, ProductRegisterWork productRegisterWork, ReportConfigClosedXml config) {
                var workSheetTemp = reportWorkbook.Worksheet(config.SheetName)
                    ?? throw new Exception($"シート '{config.SheetName}' が見つかりません。");

                var productNumber = productRegisterWork.ProductNumber.Split('-')[0] ?? string.Empty;
                SetValue(config.ProductNumberRange, productNumber);
                SetValue(config.OrderNumberRange, productRegisterWork.OrderNumber);
                SetValue(config.QuantityRange, productRegisterWork.Quantity.ToString());
                SetValue(config.SerialFirstRange, productRegisterWork.SerialFirst);
                SetValue(config.SerialLastRange, productRegisterWork.SerialLast);
                SetValue(config.ProductModelRange, productMaster.ProductModel);

                void SetValue(string? range, string? value) {
                    if (string.IsNullOrEmpty(range) || string.IsNullOrEmpty(value)) { return; }
                    workSheetTemp.Cell(range).Value = value;
                }
            }

            // データベースから使用済み基板情報を取得するメソッド
            private static List<(string SubstrateModel, List<string> SubstrateNumbers)> GetUsedSubstrateData(ProductRegisterWork productRegisterWork) {
                using var con = new SqliteConnection(GetConnectionRegistration());

                var sql =
                    $"""
                    SELECT
                        SubstrateModel,
                        SubstrateNumber,
                        Decrease
                    FROM
                        {Constants.VSubstrateTableName}
                    WHERE
                        UseID = @ID
                    ORDER BY
                        SubstrateModel ASC
                    """;

                var results = con.Query<(string SubstrateModel, string SubstrateNumber, int Decrease)>(
                    sql,
                    new { ID = productRegisterWork.RowID }
                );

                var usedSubstrate = new List<(string, List<string>)>();

                foreach (var (substrateModel, substrateNumber, decrease) in results) {
                    var existingSubstrateIndex = usedSubstrate.FindIndex(x => x.Item1 == substrateModel);

                    if (existingSubstrateIndex != -1) {
                        usedSubstrate[existingSubstrateIndex].Item2.Add(substrateNumber);
                    }
                    else {
                        usedSubstrate.Add((substrateModel, new List<string> { substrateNumber }));
                    }
                }

                return usedSubstrate;
            }

            // 基板情報をExcelシートに書き込むメソッド
            private static void UpdateSubstrateDetailsInExcel(XLWorkbook reportWorkbook, ReportConfigClosedXml config, List<(string SubstrateModel, List<string> SubstrateNumbers)> usedSubstrate) {

                var targetRow = config.ResultRow;
                var configSheet = config.Sheet;
                var startColumn = 12;
                var endColumn = 27;
                var targetRowCells = configSheet.Row(targetRow).Cells(startColumn, endColumn);

                foreach (var (substrateModel, substrateNumbers) in usedSubstrate) {
                    var searchValue = substrateModel;

                    var searchAddressResult = targetRowCells
                        .FirstOrDefault(c => c.GetValue<string>() == searchValue);

                    if (searchAddressResult == null) { continue; }

                    var foundColumn = searchAddressResult.Address.ColumnNumber;

                    var configCellValue = configSheet.Cell(targetRow, foundColumn + 1).GetString()
                        ?? throw new Exception($"Configシートの行 {targetRow}, 列 {foundColumn + 1} にセル範囲が設定されていません。");

                    var targetCellValue = string.Join(",", substrateNumbers.Select(substrateNumber => substrateNumber.Split('-')[0]));

                    var targetSheet = reportWorkbook.Worksheet(config.SheetName)
                        ?? throw new Exception($"シート '{config.SheetName}' が見つかりません。");
                    targetSheet.Cell(configCellValue).Value = targetCellValue;
                }
            }

        }

        // リスト
        public static class ListGeneratorClosedXml {
            // ProductCellRanges クラスを定義
            private class ProductCellRanges {
                public string? ProductNameRange { get; set; }
                public string? ProductNumberRange { get; set; }
                public string? OrderNumberRange { get; set; }
                public string? RegDateRange { get; set; }
                public string? ProductModelConfig { get; set; }
                public string? ProductModelRange { get; set; }
                public string? QuantityRange { get; set; }
                public string? SerialFirstRange { get; set; }
                public string? SerialLastRange { get; set; }
                public string? CommentRange { get; set; }
                public string? QrCodeRange { get; set; }
                public bool HasSubstrateInput { get; set; }
            }
            // リストを生成するメインメソッド
            public static void GenerateList(ProductMaster productMaster, ProductRegisterWork productRegisterWork) {
                try {
                    // 1. Excel設定の読み込みとワークブックの準備
                    var configPath = Path.Combine(Environment.CurrentDirectory, "config", "General", "Excel", "ConfigList.xlsm");
                    if (!File.Exists(configPath)) {
                        throw new FileNotFoundException($"設定ファイルが見つかりません: {configPath}");
                    }

                    // Excelが開かれていても読み取れるようにFileShare.ReadWrite指定
                    using var fs = new FileStream(configPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

                    // FileStreamからXLWorkbookを読み込む
                    using var workBook = new XLWorkbook(fs);

                    var (targetSheetName, productName, resultRow, configSheet) = LoadExcelConfiguration(workBook, productMaster.ProductModel);

                    // 2. 製品情報の設定とExcelへの書き込み
                    var productCellRanges = GetProductCellRanges(configSheet, resultRow);
                    var targetSheet = workBook.Worksheet(targetSheetName) ?? throw new Exception($"テンプレートシート:[{targetSheetName}]が見つかりません。");
                    PopulateProductDetails(targetSheet, productMaster, productRegisterWork, productCellRanges, productName);

                    // 3. データベースから使用済み基板情報を取得
                    var usedSubstrate = GetUsedSubstrateData(productRegisterWork);

                    // 4. 基板情報のExcelへの書き込み
                    UpdateSubstrateDetailsInExcel(configSheet, targetSheet, resultRow, usedSubstrate);

                    // 5. QRコードの生成と埋め込み
                    if (!string.IsNullOrEmpty(productCellRanges.QrCodeRange)) {
                        GenerateAndEmbedQrCode(targetSheet, productMaster, productRegisterWork, productCellRanges.QrCodeRange);
                    }

                    // 6. Excelファイルの保存と印刷
                    var temporarilyPath = Path.Combine(Environment.CurrentDirectory, "config", "General", "Excel", "temporarilyList.xlsm");
                    SaveAndPrintExcel(workBook, temporarilyPath, targetSheetName);
                } catch (Exception ex) {
                    MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            // Excel設定を読み込むメソッド
            private static (string sheetName, string productName, int resultRow, IXLWorksheet configSheet) LoadExcelConfiguration(XLWorkbook workBook, string productModel) {
                var configSheetName = "Sheet1";
                var configSheet = workBook.Worksheet(configSheetName)
                    ?? throw new Exception($"設定シート:[{configSheetName}]が見つかりません。");

                // --- 品目番号の検索 ---
                var targetCell = configSheet
                    .Column(1) // 1列目を取得
                    .CellsUsed() // 値が入っているセルのみ
                    .FirstOrDefault(c => c.GetValue<string>() == productModel)
                    ?? throw new Exception($"Configに品目番号:[{productModel}]が見つかりません。");

                var resultRow = targetCell.Address.RowNumber;

                // --- シート名 ---
                var sheetName = configSheet.Cell(resultRow, 2).GetString();
                if (string.IsNullOrEmpty(sheetName)) {
                    throw new Exception($"設定シートの行 {resultRow}, 列 2 にシート名が設定されていません。");
                }

                var productName = configSheet.Cell(resultRow, 3).GetString();

                return (sheetName, productName, resultRow, configSheet);
            }

            // 製品情報に関連するExcelのセル範囲を取得するヘルパーメソッド
            private static ProductCellRanges GetProductCellRanges(IXLWorksheet configSheet, int resultRow) {
                return new ProductCellRanges {
                    ProductNameRange = configSheet.Cell(resultRow, 4).GetString(),
                    ProductNumberRange = configSheet.Cell(resultRow, 5).GetString(),
                    OrderNumberRange = configSheet.Cell(resultRow, 6).GetString(),
                    RegDateRange = configSheet.Cell(resultRow, 7).GetString(),
                    ProductModelConfig = configSheet.Cell(resultRow, 8).GetString(),
                    ProductModelRange = configSheet.Cell(resultRow, 9).GetString(),
                    QuantityRange = configSheet.Cell(resultRow, 10).GetString(),
                    SerialFirstRange = configSheet.Cell(resultRow, 11).GetString(),
                    SerialLastRange = configSheet.Cell(resultRow, 12).GetString(),
                    CommentRange = configSheet.Cell(resultRow, 13).GetString(),
                    QrCodeRange = configSheet.Cell(resultRow, 14).GetString(),
                };
            }

            // 製品情報をExcelシートに書き込むメソッド
            private static void PopulateProductDetails(IXLWorksheet targetSheet, ProductMaster productMaster, ProductRegisterWork productRegisterWork, ProductCellRanges ranges, string productName) {

                SetValue(ranges.ProductNameRange, productName);
                SetValue(ranges.ProductNumberRange, productRegisterWork.ProductNumber);
                SetValue(ranges.OrderNumberRange, productRegisterWork.OrderNumber);
                SetValue(ranges.RegDateRange, productRegisterWork.RegDate);
                SetValue(ranges.ProductModelRange, productMaster.ProductModel);
                SetValue(ranges.QuantityRange, productRegisterWork.Quantity.ToString());
                SetValue(ranges.SerialFirstRange, productRegisterWork.SerialFirst);
                SetValue(ranges.SerialLastRange, productRegisterWork.SerialLast);
                SetValue(ranges.CommentRange, productRegisterWork.Comment);

                void SetValue(string? range, string? value) {
                    if (string.IsNullOrEmpty(range) || string.IsNullOrEmpty(value)) { return; }
                    targetSheet.Cell(range).Value = value;
                }
            }

            // データベースから使用済み基板情報を取得するメソッド
            private static List<(string SubstrateModel, List<string> SubstrateNumbers, List<int> Decreases)> GetUsedSubstrateData(ProductRegisterWork productRegisterWork) {
                using var con = new SqliteConnection(GetConnectionRegistration());

                var sql =
                    $"""
                    SELECT
                        SubstrateModel,
                        SubstrateNumber,
                        Decrease
                    FROM
                        {Constants.VSubstrateTableName}
                    WHERE
                        UseID = @ID
                    ORDER BY
                        SubstrateModel ASC
                    """;

                var results = con.Query<(string SubstrateModel, string SubstrateNumber, int Decrease)>(
                    sql,
                    new { ID = productRegisterWork.RowID }
                );

                var usedSubstrate = new List<(string, List<string>, List<int>)>();

                foreach (var (substrateModel, substrateNumber, decrease) in results) {
                    var existingSubstrateIndex = usedSubstrate.FindIndex(x => x.Item1 == substrateModel);

                    if (existingSubstrateIndex != -1) {
                        usedSubstrate[existingSubstrateIndex].Item2.Add(substrateNumber);
                        usedSubstrate[existingSubstrateIndex].Item3.Add(-decrease);
                    }
                    else {
                        usedSubstrate.Add((substrateModel, new List<string> { substrateNumber }, new List<int> { -decrease }));
                    }
                }

                return usedSubstrate;
            }

            // 基板情報をExcelシートに書き込むメソッド
            private static void UpdateSubstrateDetailsInExcel(IXLWorksheet configSheet, IXLWorksheet targetSheet, int resultRow, List<(string SubstrateModel, List<string> SubstrateNumbers, List<int> Decreases)> usedSubstrate) {

                var targetRow = resultRow;

                foreach (var (substrateModel, substrateNumbers, decreases) in usedSubstrate) {
                    var searchValue = substrateModel;

                    var searchAddressResult = configSheet
                        .Row(targetRow)
                        .CellsUsed()
                        .FirstOrDefault(c => c.GetValue<string>() == searchValue);

                    if (searchAddressResult == null) { continue; }

                    var foundColumn = searchAddressResult.Address.ColumnNumber;

                    var configCellValue = configSheet.Cell(targetRow, foundColumn + 1).GetString()
                        ?? throw new Exception($"Configシートの行 {targetRow}, 列 {foundColumn + 1} にセル範囲が設定されていません。");

                    var targetCellValue = string.Join("   ", substrateNumbers.Select((substrateNumber, k) => $"{substrateNumber}({decreases[k]})"));

                    targetSheet.Cell(configCellValue).Value = targetCellValue;
                }
            }

            // QRコードを生成し、Excelシートに埋め込むメソッド
            private static void GenerateAndEmbedQrCode(IXLWorksheet targetSheet, ProductMaster productMaster, ProductRegisterWork productRegisterWork, string qrCodeRange) {
                BarcodeWriter<Bitmap> qr = new() {
                    Format = BarcodeFormat.QR_CODE,
                    Options = new QrCodeEncodingOptions {
                        ErrorCorrection = ErrorCorrectionLevel.L,
                        CharacterSet = "Shift_JIS",
                        Width = 90,
                        Height = 90,
                    },
                    Renderer = new BitmapRenderer {
                        Foreground = Color.Gray,
                        Background = Color.White,
                    }
                };

                // QRコードのデータ文字列を構築
                var qrData = $"{productRegisterWork.OrderNumber};{productRegisterWork.ProductNumber};{productMaster.ProductModel};{productRegisterWork.Quantity};{productRegisterWork.SerialFirst};{productRegisterWork.SerialLast}";
                var qrBitmap = qr.Write(qrData);

                // Excelに画像を埋め込む
                //var picture = workSheetTemp.Drawings.AddPicture("QR", qrBitmap);
                using var ms = new MemoryStream();
                qrBitmap.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                ms.Position = 0; // ストリームの先頭に戻す

                var picture = targetSheet.AddPicture(ms).MoveTo(targetSheet.Cell(qrCodeRange));
            }

            // Excelファイルを保存し、印刷処理を行うメソッド
            private static void SaveAndPrintExcel(XLWorkbook workBook, string temporarilyPath, string sheetName) {
                // 保存前に改ページプレビューを強制設定
                ForcePageBreakPreview(workBook);
                workBook.SaveAs(temporarilyPath);

                // 印刷 (COM Interop)
                Microsoft.Office.Interop.Excel.Application? xlApp = null;
                Microsoft.Office.Interop.Excel.Workbooks? xlBooks = null;
                Microsoft.Office.Interop.Excel.Workbook? xlBook = null;
                Microsoft.Office.Interop.Excel.Sheets? xlSheets = null;
                Microsoft.Office.Interop.Excel.Worksheet? xlSheet = null;

                try {
                    xlApp = new Microsoft.Office.Interop.Excel.Application {
                        Visible = true // Excelウィンドウを表示します。
                    };

                    // ワークブック開く
                    xlBooks = xlApp.Workbooks;
                    xlBook = xlBooks.Open(temporarilyPath, ReadOnly: true);

                    // ワークシート選択
                    xlSheets = xlBook.Sheets;
                    xlSheet = (Microsoft.Office.Interop.Excel.Worksheet?)xlSheets[sheetName];

                    // ワークシート表示
                    xlSheet?.Activate();

                    // 印刷処理が必要な場合は、ここに xlSheet.PrintOut() などを追加します。
                    // 例: xlSheet.PrintOut();

                    // ワークブックを閉じてExcelを終了 (必要に応じてコメント解除)
                    // xlBook.Close(false);
                    // xlApp.Quit();
                } finally {
                    // COMオブジェクトの解放
                    ReleaseComObject(xlSheet);
                    ReleaseComObject(xlSheets);
                    ReleaseComObject(xlBook);
                    ReleaseComObject(xlBooks);
                    ReleaseComObject(xlApp);

                    static void ReleaseComObject(object? comObj) {
                        if (comObj is not null) {
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(comObj);
                        }
                    }
                }
            }
        }

        // チェックシート
        public static class CheckSheetGeneratorClosedXml {

            // チェックシート生成データを保持するクラス
            public class CheckSheetConfigData {
                public string? ProductModelRange { get; set; }
                public string? ProductNumberRange { get; set; }
                public string? OrderNumberRange { get; set; }
                public string? QuantityRange { get; set; }
                public string? SerialFirstRange { get; set; }
                public string? SerialLastRange { get; set; }
                public string? RegDateRange { get; set; }
                public string? DateFormat { get; set; }
                public string? RegTemperatureRange { get; set; }
                public string? RegHumidityRange { get; set; }
                public string? BaseFilePath { get; set; }
                public List<string> SheetNames { get; set; } = [];
            }

            // Excelチェックシートを生成し、データを書き込み、印刷します。
            public static void GenerateCheckSheet(ProductMaster productMaster, ProductRegisterWork productRegisterWork) {
                // 設定ファイルのパスを構築
                var configPath = Path.Combine(Environment.CurrentDirectory, "config", "General", "Excel", "ConfigCheckSheet.xlsm");
                var temporarilyPath = Path.Combine(Environment.CurrentDirectory, "config", "General", "Excel", "temporarilyCheckSheet.xlsm");

                if (!File.Exists(configPath)) {
                    throw new FileNotFoundException($"設定ファイルが見つかりません: {configPath}");
                }

                // Excelが開かれていても読み取れるようにFileShare.ReadWrite指定
                var fs = new FileStream(configPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                var workBook = new XLWorkbook(fs);

                try {

                    // 1. 設定ファイルの読み込みとメインシートの取得
                    var (configBook, excelData) = LoadAndExtractConfig(workBook, productMaster);

                    // 2. 温度・湿度入力ダイアログの表示と値の取得
                    (string temperature, string humidity) = GetTemperatureAndHumidity(excelData);

                    // 3. 日付のフォーマット
                    var formattedDate = FormatDate(productRegisterWork.RegDate, excelData.DateFormat);

                    // 4. EPPlusでExcelファイルを編集

                    // 対象のブックを読み込み
                    XLWorkbook targetBook;
                    FileStream? targetFs = null; // FileStream を保持（Dispose 用）

                    if (string.IsNullOrWhiteSpace(excelData.BaseFilePath)) {
                        targetBook = workBook; // 既存の configBook を使う
                    }
                    else {
                        targetFs = new FileStream(excelData.BaseFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                        targetBook = new XLWorkbook(targetFs);
                    }

                    try {
                        // 各シートに対して値を書き込む
                        PopulateExcelSheets(targetBook, productMaster, productRegisterWork, temperature, humidity, formattedDate, excelData);

                        // 不要なシートを非表示にする
                        HideSheets(targetBook, excelData.SheetNames);

                        // ブックを保存
                        SaveWorkbook(targetBook, temporarilyPath);
                    } finally {
                        // 新規に作った場合だけ Dispose
                        if (targetBook != workBook)
                            targetBook.Dispose();

                        // FileStream がある場合は Dispose
                        targetFs?.Dispose();
                    }

                    // 5. Excel Interopを使用して印刷
                    PrintExcelFile(temporarilyPath);
                } catch (Exception ex) {
                    // エラーメッセージを表示
                    MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                } finally {
                    workBook.Dispose();
                    fs.Dispose();
                }
            }

            // Excel設定ファイルを読み込み、メインシートから設定データを抽出します。
            private static (XLWorkbook workBook, CheckSheetConfigData excelData) LoadAndExtractConfig(XLWorkbook workBook, ProductMaster productMaster) {

                var productModel = productMaster.ProductModel;

                // 既存ワークシートを取得
                var configSheetName = "Sheet1";
                var configSheet = workBook.Worksheet(configSheetName);

                // --- 品目番号の検索 ---
                var targetCell = configSheet
                    .Column(1) // 1列目を取得
                    .CellsUsed() // 値が入っているセルのみ
                    .FirstOrDefault(c => c.GetValue<string>() == productModel)
                    ?? throw new Exception($"Configに品目番号:[{productModel}]が見つかりません。");

                var resultRow = targetCell.Address.RowNumber;

                // ワークシートのセルから値を取得し、ExcelConfigDataオブジェクトに格納
                var excelData = new CheckSheetConfigData {
                    ProductModelRange = configSheet.Cell(resultRow, 3).GetString(),
                    ProductNumberRange = configSheet.Cell(resultRow, 4).GetString(),
                    OrderNumberRange = configSheet.Cell(resultRow, 5).GetString(),
                    QuantityRange = configSheet.Cell(resultRow, 6).GetString(),
                    SerialFirstRange = configSheet.Cell(resultRow, 7).GetString(),
                    SerialLastRange = configSheet.Cell(resultRow, 8).GetString(),
                    RegDateRange = configSheet.Cell(resultRow, 9).GetString(),
                    DateFormat = configSheet.Cell(resultRow, 10).GetString(),
                    RegTemperatureRange = configSheet.Cell(resultRow, 11).GetString(),
                    RegHumidityRange = configSheet.Cell(resultRow, 12).GetString(),
                    BaseFilePath = configSheet.Cell(resultRow, 13).GetString(),
                    SheetNames = [.. Enumerable.Range(14, 21)
                        .Select(column => configSheet.Cell(resultRow, column).GetString())
                        .TakeWhile(sheetName => !string.IsNullOrEmpty(sheetName))]
                };

                return (workBook, excelData);
            }

            // BaseFilePath が有効ならそこから Workbook を読み込み、なければデフォルトの configBook を返す。
            private static XLWorkbook LoadWorkbookOrDefault(XLWorkbook defaultBook, string? baseFilePath = null) {
                if (string.IsNullOrWhiteSpace(baseFilePath)) {
                    return defaultBook;
                }

                // ファイルを読み込んで ClosedXML の Workbook を作成する
                using var fs = new FileStream(baseFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                return new XLWorkbook(fs);
            }

            // 温度と湿度の入力ダイアログを表示し、ユーザーからの値を取得します。
            private static (string, string) GetTemperatureAndHumidity(CheckSheetConfigData excelData) {
                // 温度セルまたは湿度セルが設定されていない場合は空を返す
                if (string.IsNullOrEmpty(excelData.RegTemperatureRange) && string.IsNullOrEmpty(excelData.RegHumidityRange)) {
                    return (string.Empty, string.Empty);
                }

                using var dialog = new InputDialog1();
                if (dialog.ShowDialog() != DialogResult.OK) {
                    throw new OperationCanceledException("ユーザーによって温度・湿度入力がキャンセルされました。");
                }

                return (dialog.Temperature, dialog.Humidity);
            }

            // 日付を指定されたフォーマットで文字列に変換します。
            private static string FormatDate(string dateString, string? dateFormat) {
                var formattedDate = string.Empty;
                if (DateTime.TryParse(dateString, out var date)) {
                    formattedDate = dateFormat switch {
                        "1" => date.ToString("yyyy年MM月dd日"),
                        "2" => date.ToString("yyyy-MM-dd"),
                        _ => formattedDate
                    };
                }
                return formattedDate;
            }

            // 指定されたワークブックの各シートに製品情報を書き込みます。
            private static void PopulateExcelSheets(XLWorkbook targetBook, ProductMaster productMaster, ProductRegisterWork productRegisterWork, string temperature, string humidity, string formattedDate, CheckSheetConfigData excelData) {
                var sheetNames = excelData.SheetNames;
                foreach (var sheetName in sheetNames) {
                    var targetSheet = targetBook.Worksheet(sheetName) ?? throw new Exception($"シート[{sheetName}]が見つかりません。");

                    // 各セルに値を書き込む
                    WriteCellValue(targetSheet, excelData.ProductModelRange, productMaster.ProductModel);
                    WriteCellValue(targetSheet, excelData.ProductNumberRange, productRegisterWork.ProductNumber);
                    WriteCellValue(targetSheet, excelData.OrderNumberRange, productRegisterWork.OrderNumber);
                    WriteCellValue(targetSheet, excelData.QuantityRange, productRegisterWork.Quantity.ToString());
                    WriteCellValue(targetSheet, excelData.SerialFirstRange, productRegisterWork.SerialFirst);
                    WriteCellValue(targetSheet, excelData.SerialLastRange, productRegisterWork.SerialLast);
                    WriteCellValue(targetSheet, excelData.RegDateRange, formattedDate);
                    WriteCellValue(targetSheet, excelData.RegTemperatureRange, temperature);
                    WriteCellValue(targetSheet, excelData.RegHumidityRange, humidity);
                }
            }

            // 指定されたシートのセルに値を書き込みます。
            private static void WriteCellValue(IXLWorksheet targetSheet, string? range, string value) {
                if (string.IsNullOrEmpty(range)) {
                    return;
                }

                targetSheet.Cell(range).Value = value;
            }

            // 指定されたワークブックの不要なシートを非表示にします。
            private static void HideSheets(XLWorkbook targetBook, List<string> sheetsToKeep) {

                // すべてのシート名を取得
                var allSheetNames = targetBook.Worksheets.Select(ws => ws.Name).ToList();

                // 非表示にする対象シートを抽出
                var sheetsToHide = allSheetNames
                    .Where(name => !sheetsToKeep.Contains(name))
                    .ToList();

                // "Sheet1" は常に非表示にする
                if (!sheetsToHide.Contains("Sheet1") && targetBook.Worksheets.Any(ws => ws.Name == "Sheet1")) {
                    sheetsToHide.Add("Sheet1");
                }

                // 非表示に設定
                foreach (var sheetName in sheetsToHide) {
                    var ws = targetBook.Worksheets.Worksheet(sheetName);
                    ws.Visibility = XLWorksheetVisibility.VeryHidden;
                }
            }

            // ワークブックをファイルに保存します。
            private static void SaveWorkbook(XLWorkbook targetBook, string outputPath) {
                try {
                    // 保存前に改ページプレビューを強制設定
                    ForcePageBreakPreview(targetBook);
                    targetBook.SaveAs(outputPath);
                } catch (IOException ex) {
                    throw new IOException($"Excelファイルの保存に失敗しました: {outputPath}. 詳細: {ex.Message}", ex);
                }
            }

            // Excel Interopを使用して指定されたExcelファイルを開きます。
            private static void PrintExcelFile(string filePath) {
                Microsoft.Office.Interop.Excel.Application? xlApp = null;
                Microsoft.Office.Interop.Excel.Workbooks? xlBooks = null;
                Microsoft.Office.Interop.Excel.Workbook? xlBook = null;

                try {
                    xlApp = new Microsoft.Office.Interop.Excel.Application {
                        Visible = true // Excelウィンドウを表示します。
                    };

                    xlBooks = xlApp.Workbooks;
                    xlBook = xlBooks.Open(filePath, ReadOnly: true);

                    //// ワークブックを閉じる (falseは変更を保存しない)
                    //xlBook.Close(false);
                    //// Excelアプリケーションを終了
                    //xlApp.Quit();
                } catch (Exception ex) {
                    throw new Exception($"エラーが発生しました。詳細: {ex.Message}", ex);
                } finally {
                    // COMオブジェクトの解放
                    ReleaseComObject(xlBook);
                    ReleaseComObject(xlBooks);
                    ReleaseComObject(xlApp);

                    static void ReleaseComObject(object? comObj) {
                        if (comObj is not null) {
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(comObj);
                        }
                    }
                }
            }
        }

        // 保存前に全シートのビューを改ページプレビューに強制設定する
        // ClosedXML が view 属性を正しく読み取れないバグへの暫定対応
        // バグ修正後は CaptureSheetViews / RestoreSheetViews に切り替える
        private static void ForcePageBreakPreview(XLWorkbook workbook) {
            foreach (var ws in workbook.Worksheets) {
                ws.SheetView.View = XLSheetViewOptions.PageBreakPreview;
            }
        }

        // ClosedXML バグ修正後に使用予定
        // private static Dictionary<string, XLSheetViewOptions> CaptureSheetViews(XLWorkbook workbook) {
        //     return workbook.Worksheets.ToDictionary(ws => ws.Name, ws => ws.SheetView.View);
        // }
        //
        // private static void RestoreSheetViews(XLWorkbook workbook, Dictionary<string, XLSheetViewOptions> views) {
        //     foreach (var ws in workbook.Worksheets) {
        //         if (views.TryGetValue(ws.Name, out var view)) {
        //             ws.SheetView.View = view;
        //         }
        //     }
        // }

        // 基板設定を開く
        public static class SubstrateInformationClosedXml {
            // 基板設定ファイルを開くメソッド（例外は呼び出し元でハンドルする）
            public static void OpenSubstrateInformationClosedXml(SubstrateMaster substrateMaster) {
                // 1. 設定ファイルを読み込み、レポート設定を取得
                var configBook = LoadConfigWorkbook();
                var (filePath, fileName) = GetSubstrateConfig(configBook, substrateMaster.SubstrateModel);

                OpenExcel(filePath, fileName);
            }

            // 設定Excelワークブックを読み込む
            private static XLWorkbook LoadConfigWorkbook() {
                try {
                    var configPath = Path.Combine(Environment.CurrentDirectory, "config", "General", "Excel", "ConfigSubstrateInformation.xlsm");

                    if (!File.Exists(configPath)) {
                        throw new FileNotFoundException($"設定ファイルが見つかりません: {configPath}");
                    }

                    // Excelが開かれていても読み取れるようにFileShare.ReadWrite指定
                    using var fs = new FileStream(configPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

                    // FileStreamからXLWorkbookを読み込む
                    return new XLWorkbook(fs);
                } catch (Exception ex) {
                    throw new Exception($"設定ファイルの読み込み中にエラーが発生しました: {ex.Message}", ex);
                }
            }

            // 設定ワークブックから製品に対応する設定を抽出する
            private static (string filePath, string sheetName) GetSubstrateConfig(XLWorkbook configWorkbook, string substrateModel) {
                var configSheet = configWorkbook.Worksheet("Sheet1") ?? throw new Exception("設定ファイルのシートに Sheet1 が見つかりません。");

                // --- 品目番号の検索 ---
                var targetCell = configSheet
                    .Column(1) // 1列目を取得
                    .CellsUsed() // 値が入っているセルのみ
                    .FirstOrDefault(c => c.GetValue<string>() == substrateModel)
                    ?? throw new Exception($"Configに品目番号:[{substrateModel}]が見つかりません。");

                var resultRow = targetCell.Address.RowNumber;

                // ディレクトリパスとファイル名を取得(ディレクトリパスが空の場合は、Configの2行目の値を使用する)
                var directoryPath = configSheet.Cell(resultRow, 4).GetString().Trim('"');
                if (string.IsNullOrEmpty(directoryPath)) {
                    directoryPath = configSheet.Cell(2, 4).GetString().Trim('"');
                }

                var fileName = configSheet.Cell(resultRow, 5).GetString().Trim('"');
                if (string.IsNullOrEmpty(fileName)) {
                    throw new Exception($"設定ファイルの型式 {substrateModel} の シート名 が空です。");
                }

                var filePath = Path.Combine(directoryPath, fileName);
                if (!File.Exists(filePath)) {
                    throw new FileNotFoundException($"指定されたファイルが存在しません: {filePath}");
                }

                var sheetName = configSheet.Cell(resultRow, 6).GetString().Trim('"');

                return (filePath, sheetName);
            }

            // Excelファイルを開くメソッド
            private static void OpenExcel(string filePath, string sheetName) {

                // COM Interopを使用してExcelを開く
                Microsoft.Office.Interop.Excel.Application? xlApp = null;
                Microsoft.Office.Interop.Excel.Workbooks? xlBooks = null;
                Microsoft.Office.Interop.Excel.Workbook? xlBook = null;
                Microsoft.Office.Interop.Excel.Sheets? xlSheets = null;
                Microsoft.Office.Interop.Excel.Worksheet? xlSheet = null;

                try {
                    xlApp = new Microsoft.Office.Interop.Excel.Application {
                        Visible = true // Excelウィンドウを表示します。
                    };

                    // ワークブック開く
                    xlBooks = xlApp.Workbooks;
                    xlBook = xlBooks.Open(filePath, ReadOnly: true);

                    // ワークシート選択
                    xlSheets = xlBook.Sheets;
                    xlSheet = (Microsoft.Office.Interop.Excel.Worksheet?)(string.IsNullOrEmpty(sheetName) ? xlSheets[1] : xlSheets[sheetName]);

                    // ワークシート表示
                    xlSheet?.Activate();

                } finally {
                    // COMオブジェクトの解放
                    ReleaseComObject(xlSheet);
                    ReleaseComObject(xlSheets);
                    ReleaseComObject(xlBook);
                    ReleaseComObject(xlBooks);
                    ReleaseComObject(xlApp);

                    static void ReleaseComObject(object? comObj) {
                        if (comObj is not null) {
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(comObj);
                        }
                    }
                }
            }

        }
    }
}

