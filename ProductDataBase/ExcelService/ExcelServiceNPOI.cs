using Microsoft.Data.Sqlite;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using ProductDatabase.Other;
using System.Drawing.Imaging;
using ZXing;
using ZXing.QrCode;
using ZXing.QrCode.Internal;
using ZXing.Windows.Compatibility;
using static ProductDatabase.MainWindow;

namespace ProductDatabase.ExcelService {
    internal class ExcelServiceNPOI {
        // 成績書
        public static class ReportGeneratorNPOI {
            // 成績書生成データを保持するクラス
            public class ReportConfigNPOI {
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
            }
            //// レポートを生成するメインメソッド
            public static void GenerateReport(ProductInformation productInfo) {
                try {
                    // 1. 設定ファイルを読み込み、レポート設定を取得
                    var configWorkbook = LoadConfigWorkbook();
                    var reportConfig = GetReportConfig(configWorkbook, productInfo.ProductModel);

                    // 2. レポートテンプレートを読み込み
                    var reportWorkbook = LoadReportTemplate(reportConfig.DirectoryPath, reportConfig.SearchFileName);

                    // 3. レポートシートにデータを挿入
                    PopulateReportSheet(reportWorkbook, productInfo, reportConfig);

                    // 4. レポートを保存
                    SaveReport(reportWorkbook, productInfo, reportConfig);

                    MessageBox.Show("成績書が正常に生成されました。", "完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
                } catch (Exception ex) {
                    // エラーメッセージをユーザーに表示
                    MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            // 設定Excelワークブックを読み込む
            private static XSSFWorkbook LoadConfigWorkbook() {
                var configPath = Path.Combine(Environment.CurrentDirectory, "config", "General", "Excel", "ConfigReport.xlsm");
                if (!File.Exists(configPath)) {
                    throw new FileNotFoundException($"設定ファイルが見つかりません: {configPath}");
                }

                try {
                    FileStream fileStreamConfig = new(configPath, FileMode.Open, FileAccess.Read, FileShare.Read);
                    return new XSSFWorkbook(fileStreamConfig);
                } catch (Exception ex) {
                    throw new Exception($"設定ファイルの読み込み中にエラーが発生しました: {ex.Message}", ex);
                }
            }

            // 設定ワークブックから製品に対応する設定を抽出する
            private static ReportConfigNPOI GetReportConfig(XSSFWorkbook configWorkbook, string productModel) {
                var workSheetMain = configWorkbook.GetSheet("Sheet1") ?? throw new Exception("設定ファイルに 'Sheet1' が見つかりません。");

                var targetColumnIndex = 0;
                var searchText = productModel;
                var resultRowIndex = -1;

                // 対象列だけ走査
                for (var rowIndex = workSheetMain.FirstRowNum; rowIndex <= workSheetMain.LastRowNum; rowIndex++) {
                    var row = workSheetMain.GetRow(rowIndex);
                    if (row == null) continue;

                    var cell = row.GetCell(targetColumnIndex);
                    if (cell == null) continue;

                    var cellValue = GetCellValue(cell) ?? string.Empty;

                    if (cellValue.Equals(searchText, StringComparison.OrdinalIgnoreCase)) {
                        resultRowIndex = rowIndex;
                    }
                }

                if (resultRowIndex == -1) { throw new Exception($"Configに品目番号:[{searchText}]が見つかりません。"); }

                // ワークシートのセルから値を取得し、ReportConfigオブジェクトに格納
                var resultRow = workSheetMain.GetRow(resultRowIndex);
                var directoryPath = GetCellValue(resultRow.GetCell(2))?.Trim('"') ?? string.Empty;
                if (string.IsNullOrEmpty(directoryPath)) { throw new Exception("Configのファイルパスが無効です。"); }
                if (!Directory.Exists(directoryPath)) { throw new FileNotFoundException($"指定されたフォルダが存在しません: {directoryPath}"); }

                var searchName = GetCellValue(resultRow.GetCell(3))?.Trim('"') ?? string.Empty;
                if (string.IsNullOrEmpty(searchName)) { throw new Exception("検索ファイル名が空です。"); }

                var filePaths = Directory.GetFiles(directoryPath, $"*{searchName}*", SearchOption.AllDirectories);
                var filePath = filePaths[0];
                var fileName = Path.GetFileNameWithoutExtension(filePath);
                var fileExtension = Path.GetExtension(filePath).ToLower();

                var sheetName = GetCellValue(resultRow.GetCell(4)) ?? string.Empty;
                return string.IsNullOrEmpty(sheetName)
                    ? throw new Exception("シート名がありません。")
                    : new ReportConfigNPOI {
                        DirectoryPath = directoryPath,
                        FilePath = filePath,
                        FileName = fileName,
                        FileExtension = fileExtension,
                        SearchFileName = searchName,
                        SheetName = sheetName,
                        ProductNumberRange = GetCellValue(resultRow.GetCell(5)) ?? string.Empty,
                        OrderNumberRange = GetCellValue(resultRow.GetCell(6)) ?? string.Empty,
                        QuantityRange = GetCellValue(resultRow.GetCell(7)) ?? string.Empty,
                        SerialFirstRange = GetCellValue(resultRow.GetCell(8)) ?? string.Empty,
                        SerialLastRange = GetCellValue(resultRow.GetCell(9)) ?? string.Empty,
                        ProductModelRange = GetCellValue(resultRow.GetCell(10)) ?? string.Empty,
                        SaveDirectory = GetCellValue(resultRow.GetCell(11)) ?? string.Empty
                    };
            }

            // レポートテンプレートExcelワークブックを読み込む
            private static XSSFWorkbook LoadReportTemplate(string directoryPath, string searchFileName) {
                var filePaths = !string.IsNullOrEmpty(searchFileName)
                    ? Directory.GetFiles(directoryPath, $"{searchFileName}*", SearchOption.AllDirectories)
                    : [];

                string filePath;

                if (filePaths.Length == 1) {
                    filePath = Path.GetFullPath(filePaths[0]);
                }
                else {
                    if (filePaths.Length > 1) {
                        MessageBox.Show("複数のファイルが見つかりました。1つ選択してください。", "ファイル選択", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }

                    filePath = ShowExcelFileDialog(directoryPath)
                        ?? throw new OperationCanceledException("ファイル選択がキャンセルされました。");
                }

                // --- FileStreamはusingで自動解放 ---
                using var fileStreamReport = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                return new XSSFWorkbook(fileStreamReport);

                static string? ShowExcelFileDialog(string initialDirectory) {
                    using var dialog = new OpenFileDialog {
                        InitialDirectory = initialDirectory,
                        Filter = "Excel ファイル|*.xlsx;*.xlsm|すべてのファイル (*.*)|*.*",
                        Multiselect = false
                    };

                    return dialog.ShowDialog() == DialogResult.OK ? dialog.FileName : null;
                }
            }

            // レポートシートに製品情報を挿入する
            private static void PopulateReportSheet(XSSFWorkbook reportWorkbook, ProductInformation productInfo, ReportConfigNPOI config) {
                var workSheetTemp = reportWorkbook.GetSheet(config.SheetName) ?? throw new Exception($"シート '{config.SheetName}' が見つかりません。");

                var productNumber = productInfo.ProductNumber.Split('-')[0] ?? string.Empty;
                SetValue(workSheetTemp, config.ProductNumberRange, productNumber);
                SetValue(workSheetTemp, config.OrderNumberRange, productInfo.OrderNumber);
                SetValue(workSheetTemp, config.QuantityRange, productInfo.Quantity.ToString());
                SetValue(workSheetTemp, config.SerialFirstRange, productInfo.SerialFirst);
                SetValue(workSheetTemp, config.SerialLastRange, productInfo.SerialLast);
                SetValue(workSheetTemp, config.ProductModelRange, productInfo.ProductModel);
            }

            // 変更されたレポートをファイルに保存する
            private static void SaveReport(XSSFWorkbook reportWorkbook, ProductInformation productInfo, ReportConfigNPOI config) {
                var fileName = config.FileName;
                var fileExtension = config.FileExtension;
                var initialDirectory = config.SaveDirectory;

                using SaveFileDialog saveFileDialog = new() {
                    Filter = $"*{fileExtension}|*{fileExtension}|All Files (*.*)|*.*",
                    FileName = $"{fileName} のコピー{productInfo.ProductNumber}{fileExtension}",
                    Title = "保存先を選択してください",
                    InitialDirectory = initialDirectory ?? Environment.CurrentDirectory // Nullの場合はデフォルトディレクトリを使用
                };

                if (saveFileDialog.ShowDialog() == DialogResult.OK) {
                    var outputPath = saveFileDialog.FileName;

                    using var fs3 = new FileStream(outputPath, FileMode.Create, FileAccess.Write);
                    reportWorkbook.SetForceFormulaRecalculation(true);
                    reportWorkbook.Write(fs3);
                }
                else {
                    // ユーザーが保存をキャンセルした場合の処理
                    throw new OperationCanceledException("キャンセルされました。");
                }
            }
        }

        // リスト
        public static class ListGeneratorNPOI {
            // リストを生成するメインメソッド
            public static void GenerateList(ProductInformation productInfo) {
                try {
                    // 1. Excel設定の読み込みとワークブックの準備
                    var workbook = LoadConfigWorkbook();

                    var (targetSheetName, productName, resultRow, resultRowIndex, workSheetMain) = LoadExcelConfiguration(workbook, productInfo.ProductModel);

                    // 2. 製品情報の設定とExcelへの書き込み
                    var productCellRanges = GetProductCellRanges(resultRow);
                    var workSheetTemp = workbook.GetSheet(targetSheetName) ?? throw new Exception($"テンプレートシート:[{targetSheetName}]が見つかりません。");
                    PopulateProductDetails(workSheetTemp, productInfo, productCellRanges, productName);

                    // 3. データベースから使用済み基板情報を取得
                    var usedSubstrate = GetUsedSubstrateData(productInfo);

                    // 4. 基板情報のExcelへの書き込み
                    UpdateSubstrateDetailsInExcel(workSheetMain, workSheetTemp, resultRowIndex, usedSubstrate);

                    // 5. QRコードの生成と埋め込み
                    if (!string.IsNullOrEmpty(productCellRanges.QrCodeRange)) {
                        GenerateAndEmbedQrCode(workSheetTemp, productInfo, productCellRanges.QrCodeRange);
                    }

                    // 6. Excelファイルの保存と印刷
                    var temporarilyPath = Path.Combine(Environment.CurrentDirectory, "config", "General", "Excel", "temporarilyList.xlsm");
                    SaveAndPrintExcel(workbook, temporarilyPath, targetSheetName);
                } catch (Exception ex) {
                    MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            // 設定Excelワークブックを読み込む
            private static XSSFWorkbook LoadConfigWorkbook() {
                var configPath = Path.Combine(Environment.CurrentDirectory, "config", "General", "Excel", "ConfigList.xlsm");
                if (!File.Exists(configPath)) {
                    throw new FileNotFoundException($"設定ファイルが見つかりません: {configPath}");
                }

                try {
                    FileStream fileStreamConfig = new(configPath, FileMode.Open, FileAccess.Read, FileShare.Read);
                    return new XSSFWorkbook(fileStreamConfig);
                } catch (Exception ex) {
                    throw new Exception($"設定ファイルの読み込み中にエラーが発生しました: {ex.Message}", ex);
                }
            }

            // Excel設定を読み込むメソッド
            private static (string sheetName, string productName, IRow resultRow, int resultRowIndex, ISheet workSheetMain) LoadExcelConfiguration(XSSFWorkbook workBook, string productModel) {
                var workSheetMain = workBook.GetSheet("Sheet1") ?? throw new Exception("設定ファイルに 'Sheet1' が見つかりません。");

                var targetColumnIndex = 0;
                var searchText = productModel;
                var resultRowIndex = -1;

                // 対象列だけ走査
                for (var rowIndex = workSheetMain.FirstRowNum; rowIndex <= workSheetMain.LastRowNum; rowIndex++) {
                    var row = workSheetMain.GetRow(rowIndex);
                    if (row == null) continue;

                    var cell = row.GetCell(targetColumnIndex);
                    if (cell == null) continue;

                    var cellValue = GetCellValue(cell) ?? string.Empty;

                    if (cellValue.Equals(searchText, StringComparison.OrdinalIgnoreCase)) {
                        resultRowIndex = rowIndex;
                    }
                }

                if (resultRowIndex == -1) { throw new Exception($"Configに品目番号:[{searchText}]が見つかりません。"); }

                var resultRow = workSheetMain.GetRow(resultRowIndex);

                var sheetName = GetCellValue(resultRow.GetCell(1)) ??
                    throw new Exception($"設定シートの行 {resultRow}, 列 2 にシート名が設定されていません。");
                var productName = GetCellValue(resultRow.GetCell(2)) ?? string.Empty;

                return (sheetName, productName, resultRow, resultRowIndex, workSheetMain);
            }

            // 製品情報に関連するExcelのセル範囲を取得するヘルパーメソッド
            private static ProductCellRanges GetProductCellRanges(IRow resultRow) {
                return new ProductCellRanges {
                    ProductNameRange = GetCellValue(resultRow.GetCell(3)) ?? string.Empty,
                    ProductNumberRange = GetCellValue(resultRow.GetCell(4)) ?? string.Empty,
                    OrderNumberRange = GetCellValue(resultRow.GetCell(5)) ?? string.Empty,
                    RegDateRange = GetCellValue(resultRow.GetCell(6)) ?? string.Empty,
                    ProductModelConfig = GetCellValue(resultRow.GetCell(7)) ?? string.Empty,
                    ProductModelRange = GetCellValue(resultRow.GetCell(8)) ?? string.Empty,
                    QuantityRange = GetCellValue(resultRow.GetCell(9)) ?? string.Empty,
                    SerialFirstRange = GetCellValue(resultRow.GetCell(10)) ?? string.Empty,
                    SerialLastRange = GetCellValue(resultRow.GetCell(11)) ?? string.Empty,
                    CommentRange = GetCellValue(resultRow.GetCell(12)) ?? string.Empty,
                    QrCodeRange = GetCellValue(resultRow.GetCell(13)) ?? string.Empty
                };
            }

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
            }

            // 製品情報をExcelシートに書き込むメソッド
            private static void PopulateProductDetails(ISheet workSheetTemp, ProductInformation productInfo, ProductCellRanges ranges, string productName) {

                SetValue(workSheetTemp, ranges.ProductNameRange, productName);
                SetValue(workSheetTemp, ranges.ProductNumberRange, productInfo.ProductNumber);
                SetValue(workSheetTemp, ranges.OrderNumberRange, productInfo.OrderNumber);
                SetValue(workSheetTemp, ranges.RegDateRange, productInfo.RegDate);
                SetValue(workSheetTemp, ranges.ProductModelRange, productInfo.ProductModel);
                SetValue(workSheetTemp, ranges.QuantityRange, productInfo.Quantity.ToString());
                SetValue(workSheetTemp, ranges.SerialFirstRange, productInfo.SerialFirst);
                SetValue(workSheetTemp, ranges.SerialLastRange, productInfo.SerialLast);
                SetValue(workSheetTemp, ranges.CommentRange, productInfo.Comment);
            }

            // データベースから使用済み基板情報を取得するメソッド
            private static List<(string SubstrateModel, List<string> SubstrateNumbers, List<int> Decreases)> GetUsedSubstrateData(ProductInformation productInfo) {
                List<(string, List<string>, List<int>)> usedSubstrate = [];

                using SqliteConnection con = new(GetConnectionRegistration());
                con.Open();
                using var cmd = con.CreateCommand();

                var tableName = $"[{productInfo.CategoryName}_Substrate]";
                cmd.CommandText =
                    $"""
                    SELECT
                        SubstrateModel,
                        SubstrateNumber,
                        Decrease
                    FROM
                        {tableName}
                    WHERE
                        UseID = @ID
                    ORDER BY
                        SubstrateModel ASC
                    ;
                    """;
                cmd.Parameters.Add("@ID", SqliteType.Text).Value = productInfo.ProductID;
                using var dr = cmd.ExecuteReader();

                while (dr.Read()) {
                    var substrateModel = dr.GetString(0);
                    var substrateNumber = dr.GetString(1);
                    var decrease = -1 * dr.GetInt32(2);

                    // 既存の substrateModel を検索し、見つかればリストに追加、なければ新しいエントリを作成
                    var existingSubstrateIndex = usedSubstrate.FindIndex(x => x.Item1 == substrateModel);

                    if (existingSubstrateIndex != -1) {
                        usedSubstrate[existingSubstrateIndex].Item2.Add(substrateNumber);
                        usedSubstrate[existingSubstrateIndex].Item3.Add(decrease);
                    }
                    else {
                        List<string> substrateNumbers = [substrateNumber];
                        List<int> decreases = [decrease];
                        usedSubstrate.Add((substrateModel, substrateNumbers, decreases));
                    }
                }
                return usedSubstrate;
            }

            // 基板情報をExcelシートに書き込むメソッド
            private static void UpdateSubstrateDetailsInExcel(ISheet workSheetMain, ISheet workSheetTemp, int resultRowIndex, List<(string SubstrateModel, List<string> SubstrateNumbers, List<int> Decreases)> usedSubstrate) {
                foreach (var (substrateModel, substrateNumbers, decreases) in usedSubstrate) {
                    //var targetRow = resultRow;
                    var searchText = substrateModel;
                    var resultColumnIndex = -1;

                    // 対象行だけ走査
                    var row = workSheetMain.GetRow(resultRowIndex);
                    if (row == null) continue;


                    // 各セルを走査
                    for (int colIndex = row.FirstCellNum; colIndex < row.LastCellNum; colIndex++) {
                        var cell = row.GetCell(colIndex);
                        if (cell == null) continue;

                        var cellValue = GetCellValue(cell) ?? string.Empty;

                        if (cellValue.Equals(searchText, StringComparison.OrdinalIgnoreCase)) {
                            resultColumnIndex = colIndex;
                        }
                    }

                    if (resultColumnIndex == -1) { throw new Exception($"Configに基板モデル:[{searchText}]が見つかりません。"); }

                    var cell2 = row.GetCell(resultColumnIndex + 1);

                    var targetCellRange = GetCellValue(cell2)
                        ?? throw new Exception($"Configシートの行 {resultRowIndex}, 列 {resultColumnIndex + 1} にセル範囲が設定されていません。");

                    var tempCellValue = string.Join("   ", substrateNumbers.Select((substrateNumber, k) => $"{substrateNumber}({decreases[k]})"));

                    SetValue(workSheetTemp, targetCellRange, tempCellValue);
                }
            }

            // QRコードを生成し、Excelシートに埋め込むメソッド
            private static void GenerateAndEmbedQrCode(ISheet workSheetTemp, ProductInformation productInfo, string qrCodeRange) {
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
                var qrData = $"{productInfo.OrderNumber};{productInfo.ProductNumber};{productInfo.ProductModel};{productInfo.Quantity};{productInfo.SerialFirst};{productInfo.SerialLast}";
                var qrBitmap = qr.Write(qrData);

                // Excelに画像を埋め込む
                GetRowColFromAddress(qrCodeRange, out var rowIndex, out var colIndex);
                InsertImageOriginalSize(workSheetTemp, qrBitmap, rowIndex, colIndex);
            }

            // Excelファイルを保存し、印刷処理を行うメソッド
            private static void SaveAndPrintExcel(XSSFWorkbook workBook, string outputPath, string sheetName) {

                using (var fs3 = new FileStream(outputPath, FileMode.Create, FileAccess.Write)) {
                    workBook.SetForceFormulaRecalculation(true);
                    workBook.Write(fs3);
                }

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
                    xlBook = xlBooks.Open(outputPath, ReadOnly: true);

                    // ワークシート選択
                    xlSheets = xlBook.Sheets;
                    xlSheet = xlSheets[sheetName];

                    // ワークシート表示
                    xlSheet.Activate();

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
        public static class CheckSheetGeneratorNPOI {
            // チェックシート生成データを保持するクラス
            public class CheckSheetConfigDataNPOI {
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
                public List<string> SheetNames { get; set; } = [];
            }
            // チェックシートを生成するメインメソッド
            public static void GenerateCheckSheet(ProductInformation productInfo) {
                try {
                    // 1. Excel設定の読み込みとワークブックの準備
                    var workbook = LoadConfigWorkbook();
                    var resultRow = LoadExcelConfiguration(workbook, productInfo.ProductModel);

                    // 2. 製品情報の設定とExcelへの書き込み
                    var productCellRanges = GetProductCellRanges(resultRow);

                    // 3. 温度・湿度入力ダイアログの表示と値の取得
                    (var temperature, var humidity) = GetTemperatureAndHumidity(productCellRanges);

                    // 4. 日付のフォーマット
                    var formattedDate = FormatDate(productInfo.RegDate, productCellRanges.DateFormat);

                    // 5. データの書き込み
                    foreach (var sheetName in productCellRanges.SheetNames) {
                        if (string.IsNullOrEmpty(sheetName)) { continue; }
                        var workSheetTemp = workbook.GetSheet(sheetName) ?? throw new Exception($"テンプレートシート:[{sheetName}]が見つかりません。");
                        PopulateProductDetails(workSheetTemp, productInfo, productCellRanges, temperature, humidity, formattedDate);
                    }

                    // 6. 不要なシートを非表示にする
                    HideSheets(workbook, productCellRanges.SheetNames);

                    // 7. Excelファイルの保存と印刷
                    var temporarilyPath = Path.Combine(Environment.CurrentDirectory, "config", "General", "Excel", "temporarilyCheckSheet.xlsm");
                    SaveAndPrintExcel(workbook, temporarilyPath);
                } catch (Exception ex) {
                    MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            // 設定Excelワークブックを読み込む
            private static XSSFWorkbook LoadConfigWorkbook() {
                var configPath = Path.Combine(Environment.CurrentDirectory, "config", "General", "Excel", "ConfigCheckSheet.xlsm");
                if (!File.Exists(configPath)) {
                    throw new FileNotFoundException($"設定ファイルが見つかりません: {configPath}");
                }

                try {
                    FileStream fileStreamConfig = new(configPath, FileMode.Open, FileAccess.Read, FileShare.Read);
                    return new XSSFWorkbook(fileStreamConfig);
                } catch (Exception ex) {
                    throw new Exception($"設定ファイルの読み込み中にエラーが発生しました: {ex.Message}", ex);
                }
            }

            // Excel設定を読み込むメソッド
            private static IRow LoadExcelConfiguration(XSSFWorkbook workBook, string productModel) {
                var workSheetMain = workBook.GetSheet("Sheet1") ?? throw new Exception("設定ファイルに 'Sheet1' が見つかりません。");

                var targetColumnIndex = 0;
                var searchText = productModel;
                var resultRowIndex = -1;

                // 対象列だけ走査
                for (var rowIndex = workSheetMain.FirstRowNum; rowIndex <= workSheetMain.LastRowNum; rowIndex++) {
                    var row = workSheetMain.GetRow(rowIndex);
                    if (row == null) continue;

                    var cell = row.GetCell(targetColumnIndex);
                    if (cell == null) continue;

                    var cellValue = GetCellValue(cell) ?? string.Empty;

                    if (cellValue.Equals(searchText, StringComparison.OrdinalIgnoreCase)) {
                        resultRowIndex = rowIndex;
                    }
                }

                if (resultRowIndex == -1) { throw new Exception($"Configに品目番号:[{searchText}]が見つかりません。"); }

                var resultRow = workSheetMain.GetRow(resultRowIndex);

                return resultRow;
            }

            // 温度と湿度の入力ダイアログを表示し、ユーザーからの値を取得します。
            private static (string temperature, string humidity) GetTemperatureAndHumidity(CheckSheetConfigDataNPOI excelData) {
                var temperature = string.Empty;
                var humidity = string.Empty;

                // 温度セルか湿度セルが設定されている場合、ダイアログを表示
                if (!string.IsNullOrEmpty(excelData.RegTemperatureRange) || !string.IsNullOrEmpty(excelData.RegHumidityRange)) {
                    var dialog = new InputDialog1();
                    var result = dialog.ShowDialog();
                    if (result != DialogResult.OK) {
                        // キャンセルされた場合は処理を中断
                        throw new OperationCanceledException("ユーザーによって温度・湿度入力がキャンセルされました。");
                    }
                    temperature = dialog.Temperature;
                    humidity = dialog.Humidity;
                }
                return (temperature, humidity);
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

            // 製品情報に関連するExcelのセル範囲を取得するヘルパーメソッド
            private static CheckSheetConfigDataNPOI GetProductCellRanges(IRow resultRow) {
                return new CheckSheetConfigDataNPOI {
                    ProductModelRange = GetCellValue(resultRow.GetCell(2)) ?? string.Empty,
                    ProductNumberRange = GetCellValue(resultRow.GetCell(3)) ?? string.Empty,
                    OrderNumberRange = GetCellValue(resultRow.GetCell(4)) ?? string.Empty,
                    QuantityRange = GetCellValue(resultRow.GetCell(5)) ?? string.Empty,
                    SerialFirstRange = GetCellValue(resultRow.GetCell(6)) ?? string.Empty,
                    SerialLastRange = GetCellValue(resultRow.GetCell(7)) ?? string.Empty,
                    RegDateRange = GetCellValue(resultRow.GetCell(8)) ?? string.Empty,
                    DateFormat = GetCellValue(resultRow.GetCell(9)) ?? string.Empty,
                    RegTemperatureRange = GetCellValue(resultRow.GetCell(10)) ?? string.Empty,
                    RegHumidityRange = GetCellValue(resultRow.GetCell(11)) ?? string.Empty,
                    SheetNames = [.. Enumerable.Range(12, 20)
                        .Select(column => GetCellValue(resultRow.GetCell(column)) ?? string.Empty)
                        .TakeWhile(sheetName => !string.IsNullOrEmpty(sheetName))],
                };
            }

            // 製品情報をExcelシートに書き込むメソッド
            private static void PopulateProductDetails(ISheet workSheetTemp, ProductInformation productInfo, CheckSheetConfigDataNPOI ranges, string temperature, string humidity, string formattedDate) {
                SetValue(workSheetTemp, ranges.ProductModelRange, productInfo.ProductModel);
                SetValue(workSheetTemp, ranges.ProductNumberRange, productInfo.ProductNumber);
                SetValue(workSheetTemp, ranges.OrderNumberRange, productInfo.OrderNumber);
                SetValue(workSheetTemp, ranges.QuantityRange, productInfo.Quantity.ToString());
                SetValue(workSheetTemp, ranges.SerialFirstRange, productInfo.SerialFirst);
                SetValue(workSheetTemp, ranges.SerialLastRange, productInfo.SerialLast);
                SetValue(workSheetTemp, ranges.RegDateRange, formattedDate);
                SetValue(workSheetTemp, ranges.RegTemperatureRange, temperature);
                SetValue(workSheetTemp, ranges.RegHumidityRange, humidity);
            }

            // 指定されたワークブックの不要なシートを非表示にします。
            private static void HideSheets(XSSFWorkbook workBook, List<string> sheetsToKeep) {
                // 非表示にするシートのインデックスを特定
                var allSheetNames = new List<string>();
                for (var i = 0; i < workBook.NumberOfSheets; i++) {
                    allSheetNames.Add(workBook.GetSheetName(i));
                }

                var sheetIndicesToHide = allSheetNames
                    .Select((name, index) => new { Name = name, Index = index })
                    .Where(sheet => !sheetsToKeep.Contains(sheet.Name))
                    .Select(sheet => sheet.Index)
                    .ToList();

                // "Sheet1" は常に非表示にする
                var sheet1Index = workBook.GetSheetIndex("Sheet1");
                if (sheet1Index != -1 && !sheetIndicesToHide.Contains(sheet1Index)) {
                    sheetIndicesToHide.Add(sheet1Index);
                }


                // シートを非表示に設定
                foreach (var sheetIndex in sheetIndicesToHide) {
                    workBook.SetSheetVisibility(sheetIndex, SheetVisibility.VeryHidden);
                }
            }

            // Excelファイルを保存し、印刷処理を行うメソッド
            private static void SaveAndPrintExcel(XSSFWorkbook workBook, string outputPath) {

                using (var fs3 = new FileStream(outputPath, FileMode.Create, FileAccess.Write)) {
                    workBook.SetForceFormulaRecalculation(true);
                    workBook.Write(fs3);
                }

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
                    xlBook = xlBooks.Open(outputPath, ReadOnly: true);

                    //// ワークシート選択
                    //xlSheets = xlBook.Sheets;
                    //xlSheet = xlSheets[sheetNames];

                    //// ワークシート表示
                    //xlSheet.Activate();

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

        // 基板設定を開く
        public static class SubstrateInformationNPOI {
            // 基板設定ファイルを開くメソッド
            public static void OpenSubstrateInformationNPOI(ProductInformation productInfo) {
                try {
                    // 1. 設定ファイルを読み込み、レポート設定を取得
                    var configWorkbook = LoadConfigWorkbook();
                    var (filePaths, fileName) = GetSubstrateConfig(configWorkbook, productInfo.SubstrateModel);

                    OpenExcel(filePaths, fileName);

                } catch (Exception ex) {
                    // エラーメッセージを表示
                    MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            // 設定Excelワークブックを読み込む
            private static XSSFWorkbook LoadConfigWorkbook() {
                var configPath = Path.Combine(Environment.CurrentDirectory, "config", "General", "Excel", "ConfigSubstrateInformation.xlsm");
                if (!File.Exists(configPath)) {
                    throw new FileNotFoundException($"設定ファイルが見つかりません: {configPath}");
                }

                try {
                    FileStream fileStreamConfig = new(configPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    return new XSSFWorkbook(fileStreamConfig);
                } catch (Exception ex) {
                    throw new Exception($"設定ファイルの読み込み中にエラーが発生しました: {ex.Message}", ex);
                }
            }

            // 設定ワークブックから製品に対応する設定を抽出する
            private static (string filePaths, string sheetName) GetSubstrateConfig(XSSFWorkbook configWorkbook, string substrateModel) {
                var workSheetMain = configWorkbook.GetSheet("Sheet1") ?? throw new Exception("設定ファイルのシートに Sheet1 が見つかりません。");

                var targetColumnIndex = 0;
                var searchText = substrateModel;
                var resultRowIndex = -1;

                // 対象列だけ走査
                for (var rowIndex = workSheetMain.FirstRowNum; rowIndex <= workSheetMain.LastRowNum; rowIndex++) {
                    var row = workSheetMain.GetRow(rowIndex);
                    if (row == null) continue;

                    var cell = row.GetCell(targetColumnIndex);
                    if (cell == null) continue;

                    var cellValue = GetCellValue(cell) ?? string.Empty;

                    if (cellValue.Equals(searchText, StringComparison.OrdinalIgnoreCase)) {
                        resultRowIndex = rowIndex;
                    }
                }

                if (resultRowIndex == -1) { throw new Exception($"Configに品目番号:[{searchText}]が見つかりません。"); }

                var resultRow = workSheetMain.GetRow(resultRowIndex);
                var defaultRow = workSheetMain.GetRow(1);


                // ディレクトリパスとファイル名を取得(ディレクトリパスが空の場合は、Configの2行目の値を使用する)
                var directoryPath = GetCellValue(resultRow.GetCell(3))?.Trim('"');
                if (string.IsNullOrEmpty(directoryPath)) { directoryPath = GetCellValue(defaultRow.GetCell(3))?.Trim('"'); }
                if (!Directory.Exists(directoryPath)) { throw new FileNotFoundException($"指定されたフォルダが存在しません: {directoryPath}"); }

                var fileName = GetCellValue(resultRow.GetCell(4))?.Trim('"');
                if (string.IsNullOrEmpty(fileName)) { throw new Exception($"設定ファイルの型式 {substrateModel} の ファイル名 が空です。"); }

                var filePaths = Path.Combine(directoryPath, fileName);
                if (!File.Exists(filePaths)) { throw new FileNotFoundException($"指定されたファイルが存在しません: {filePaths}"); }

                var sheetName = GetCellValue(resultRow.GetCell(5)) ?? string.Empty;

                return (filePaths, sheetName);
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
                    xlSheet = string.IsNullOrEmpty(sheetName) ? xlSheets[1] : xlSheets[sheetName];

                    // ワークシート表示
                    xlSheet.Activate();

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


        private static string? GetCellValue(ICell cell) {
            if (cell == null) return string.Empty;
            return cell.CellType switch {
                CellType.String => cell.StringCellValue,
                CellType.Numeric => cell.NumericCellValue.ToString(),
                CellType.Boolean => cell.BooleanCellValue.ToString(),
                CellType.Formula => cell.ToString(),
                _ => cell.ToString()
            };
        }
        private static void SetValue(ISheet sheet, string? address, string? value) {
            if (string.IsNullOrEmpty(address) || string.IsNullOrEmpty(value)) { return; }
            GetRowColFromAddress(address, out var rowIndex, out var colIndex);
            var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            var cell = row.GetCell(colIndex) ?? row.CreateCell(colIndex);
            cell.SetCellValue(value);
        }
        private static int ColumnNameToIndex(string columnName) {
            var index = 0;
            foreach (var c in columnName.ToUpper()) {
                if (c < 'A' || c > 'Z') throw new ArgumentException("Invalid column name");
                index = index * 26 + c - 'A' + 1;
            }
            return index - 1; // 0始まり
        }
        private static void GetRowColFromAddress(string address, out int rowIndex, out int colIndex) {
            // 数字の位置を検索
            var i = 0;
            while (i < address.Length && char.IsLetter(address[i])) i++;

            var colPart = address[..i];   // "A"
            var rowPart = address[i..];      // "2"

            colIndex = ColumnNameToIndex(colPart);
            rowIndex = int.Parse(rowPart) - 1; // 0始まり
        }
        private static void InsertImageOriginalSize(ISheet sheet, Bitmap bitmap, int rowIndex, int colIndex) {
            var workbook = sheet.Workbook;

            // Bitmap をメモリストリームに保存 → バイト配列化
            byte[] bytes;
            using (var ms = new MemoryStream()) {
                bitmap.Save(ms, ImageFormat.Png); // PNG形式で保存
                bytes = ms.ToArray();
            }

            // ワークブックに画像を追加
            var pictureIdx = workbook.AddPicture(bytes, PictureType.PNG);

            // 描画オブジェクトを取得
            var drawing = sheet.CreateDrawingPatriarch();
            var helper = workbook.GetCreationHelper();
            var anchor = helper.CreateClientAnchor();
            anchor.AnchorType = AnchorType.DontMoveAndResize;

            // 画像を作成
            drawing.CreatePicture(anchor, pictureIdx);

            anchor.Col1 = colIndex;   // 左上列
            anchor.Row1 = rowIndex;   // 左上行
            anchor.Col2 = anchor.Col1;
            anchor.Row2 = anchor.Row1;

        }
    }
}
