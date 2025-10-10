using Microsoft.Data.Sqlite;
using OfficeOpenXml;
using ZXing;
using ZXing.QrCode;
using ZXing.QrCode.Internal;
using ZXing.Windows.Compatibility;
using static ProductDatabase.MainWindow;

namespace ProductDatabase.ExcelService {
    internal class ExcelServiceEPPlus {
        // 成績書
        public static class ReportGeneratorEPPlus {
            // 成績書生成データを保持するクラス
            public class ReportConfigEPPlus {
                public required ExcelWorksheet Sheet { get; set; }
                public required int Row { get; set; }
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
            // レポートを生成するメインメソッド
            public static void GenerateReport(ProductInformation productInfo) {
                try {
                    // 1. 設定ファイルを読み込み、レポート設定を取得
                    var configWorkbook = LoadConfigWorkbook();
                    var reportConfig = GetReportConfig(configWorkbook, productInfo.ProductModel);

                    // 2. レポートテンプレートを読み込み
                    var reportWorkbook = LoadReportTemplate(reportConfig);

                    // 3. レポートシートにデータを挿入
                    PopulateReportSheet(reportWorkbook, productInfo, reportConfig);

                    // 4. データベースから使用済み基板情報を取得
                    var usedSubstrate = GetUsedSubstrateData(productInfo);

                    // 5. 基板情報のExcelへの書き込み
                    UpdateSubstrateDetailsInExcel(reportWorkbook, reportConfig, usedSubstrate);

                    // 6. レポートを保存
                    SaveReport(reportWorkbook, productInfo, reportConfig);

                    MessageBox.Show("成績書が正常に生成されました。", "完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
                } catch (Exception ex) {
                    // エラーメッセージをユーザーに表示
                    MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            // 設定Excelワークブックを読み込む
            private static ExcelPackage LoadConfigWorkbook() {
                var configPath = Path.Combine(Environment.CurrentDirectory, "config", "General", "Excel", "ConfigReport.xlsm");
                if (!File.Exists(configPath)) {
                    throw new FileNotFoundException($"設定ファイルが見つかりません: {configPath}");
                }

                try {
                    FileStream fileStreamConfig = new(configPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    return new ExcelPackage(fileStreamConfig);
                } catch (Exception ex) {
                    throw new Exception($"設定ファイルの読み込み中にエラーが発生しました: {ex.Message}", ex);
                }
            }

            // 設定ワークブックから製品に対応する設定を抽出する
            private static ReportConfigEPPlus GetReportConfig(ExcelPackage configWorkbook, string productModel) {
                var sheet = configWorkbook.Workbook.Worksheets["Sheet1"]
                    ?? throw new Exception("設定ファイルに 'Sheet1' が見つかりません。");

                // --- 品目番号の検索 ---
                var targetCell = sheet.Cells.FirstOrDefault(c => c.Start.Column == 1 && c.Value?.ToString() == productModel)
                    ?? throw new Exception($"Configに品目番号:[{productModel}]が見つかりません。");
                var row = targetCell.Start.Row;

                // --- シート名 ---
                var sheetName = sheet.Cells[row, 5].Text;
                if (string.IsNullOrEmpty(sheetName)) {
                    throw new Exception($"設定ファイルのシート名が空です。");
                }

                // --- パスとファイル名の取得 ---
                var directoryPath = sheet.Cells[row, 3].Value?.ToString()?.Trim('"')
                    ?? sheet.Cells[2, 3].Value?.ToString()?.Trim('"');

                if (string.IsNullOrEmpty(directoryPath)) { throw new Exception("Configのファイルパスが無効です。"); }

                if (!Directory.Exists(directoryPath)) {
                    throw new DirectoryNotFoundException($"指定されたフォルダが存在しません: {directoryPath}");
                }

                var saveDirectory = sheet.Cells[row, 12].Value?.ToString()?.Trim('"')
                    ?? sheet.Cells[2, 12].Value?.ToString()?.Trim('"');

                var searchName = sheet.Cells[row, 4].Text.Trim('"');
                if (string.IsNullOrEmpty(searchName)) { throw new DirectoryNotFoundException($"検索ファイル名が空です。"); }

                // --- ファイル検索 ---
                var filePath = FindExcelFile(directoryPath, searchName);
                var fileName = Path.GetFileNameWithoutExtension(filePath);
                var fileExtension = Path.GetExtension(filePath).ToLowerInvariant();

                // --- ReportConfig を構築 ---
                return new ReportConfigEPPlus {
                    Sheet = sheet,
                    Row = row,
                    DirectoryPath = directoryPath,
                    FilePath = filePath,
                    FileName = fileName,
                    FileExtension = fileExtension,
                    SearchFileName = searchName,
                    SheetName = sheetName,
                    ProductNumberRange = sheet.Cells[row, 6].Text ?? string.Empty,
                    OrderNumberRange = sheet.Cells[row, 7].Text ?? string.Empty,
                    QuantityRange = sheet.Cells[row, 8].Text ?? string.Empty,
                    SerialFirstRange = sheet.Cells[row, 9].Text ?? string.Empty,
                    SerialLastRange = sheet.Cells[row, 10].Text ?? string.Empty,
                    ProductModelRange = sheet.Cells[row, 11].Text ?? string.Empty,
                    SaveDirectory = saveDirectory
                };
            }
            private static string FindExcelFile(string directoryPath, string searchName) {
                var filePaths = !string.IsNullOrEmpty(searchName)
                    ? Directory.GetFiles(directoryPath, $"{searchName}*", SearchOption.AllDirectories)
                    : [];

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

                    return dialog.ShowDialog() == DialogResult.OK ? dialog.FileName : null;
                }
            }

            // レポートテンプレートExcelワークブックを読み込む
            private static ExcelPackage LoadReportTemplate(ReportConfigEPPlus config) {
                var filePath = config.FilePath;

                try {
                    FileStream fileStreamReport = new(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    return new ExcelPackage(fileStreamReport);
                } catch (Exception ex) {
                    throw new Exception($"レポートテンプレートの読み込み中にエラーが発生しました: {ex.Message}", ex);
                }
            }

            // レポートシートに製品情報を挿入する
            private static void PopulateReportSheet(ExcelPackage reportWorkbook, ProductInformation productInfo, ReportConfigEPPlus config) {
                var workSheetTemp = reportWorkbook.Workbook.Worksheets[config.SheetName] ?? throw new Exception($"シート '{config.SheetName}' が見つかりません。");

                var productNumber = productInfo.ProductNumber.Split('-')[0] ?? string.Empty;
                SetValue(config.ProductNumberRange, productNumber);
                SetValue(config.OrderNumberRange, productInfo.OrderNumber);
                SetValue(config.QuantityRange, productInfo.Quantity.ToString());
                SetValue(config.SerialFirstRange, productInfo.SerialFirst);
                SetValue(config.SerialLastRange, productInfo.SerialLast);
                SetValue(config.ProductModelRange, productInfo.ProductModel);

                void SetValue(string? range, string? value) {
                    if (!string.IsNullOrEmpty(range) && !string.IsNullOrEmpty(value)) {
                        workSheetTemp.Cells[range].Value = value;
                    }
                }
            }

            // データベースから使用済み基板情報を取得するメソッド
            private static List<(string SubstrateModel, List<string> SubstrateNumbers)> GetUsedSubstrateData(ProductInformation productInfo) {
                List<(string, List<string>)> usedSubstrate = [];

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

                    // 既存の substrateModel を検索し、見つかればリストに追加、なければ新しいエントリを作成
                    var existingSubstrateIndex = usedSubstrate.FindIndex(x => x.Item1 == substrateModel);

                    if (existingSubstrateIndex != -1) {
                        usedSubstrate[existingSubstrateIndex].Item2.Add(substrateNumber);
                    }
                    else {
                        List<string> substrateNumbers = [substrateNumber];
                        usedSubstrate.Add((substrateModel, substrateNumbers));
                    }
                }
                return usedSubstrate;
            }

            // 基板情報をExcelシートに書き込むメソッド
            private static void UpdateSubstrateDetailsInExcel(ExcelPackage reportWorkbook, ReportConfigEPPlus config, List<(string SubstrateModel, List<string> SubstrateNumbers)> usedSubstrate) {
                foreach (var (substrateModel, substrateNumbers) in usedSubstrate) {
                    var targetRow = config.Row;
                    var searchValue = substrateModel;
                    var foundColumn = 0;

                    var workSheetMain = config.Sheet;
                    var searchAddressResult2 = workSheetMain.Cells
                        .Where(x => x.Start.Row == targetRow && x.Value?.ToString() == searchValue)
                        .FirstOrDefault();

                    if (searchAddressResult2 == null) { continue; }

                    foundColumn = searchAddressResult2.Start.Column;

                    var mainCellValue = workSheetMain.Cells[targetRow, foundColumn + 1].Value?.ToString()
                        ?? throw new Exception($"Configシートの行 {targetRow}, 列 {foundColumn + 1} にセル範囲が設定されていません。");

                    var tempCellValue = string.Join(",", substrateNumbers.Select(substrateNumber => substrateNumber.Split('-')[0]));

                    var workSheetTemp = reportWorkbook.Workbook.Worksheets[config.SheetName]
                        ?? throw new Exception($"シート '{config.SheetName}' が見つかりません。");
                    workSheetTemp.Cells[mainCellValue].Value = tempCellValue;
                }
            }

            // 変更されたレポートをファイルに保存する
            private static void SaveReport(ExcelPackage reportWorkbook, ProductInformation productInfo, ReportConfigEPPlus config) {
                var fileName = config.FileName;
                var fileExtension = config.FileExtension;
                var initialDirectory = config.SaveDirectory;

                if (!Directory.Exists(initialDirectory)) {
                    MessageBox.Show($"設定されている保存先が見つかりませんでした。\r\n{initialDirectory}", "ファイル選択", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                using SaveFileDialog saveFileDialog = new() {
                    Filter = $"Excel Files (*{fileExtension})|*{fileExtension}|All Files (*.*)|*.*",
                    FileName = $"{fileName} のコピー{productInfo.ProductNumber}{fileExtension}",
                    Title = "保存先を選択してください",
                    InitialDirectory = string.IsNullOrWhiteSpace(initialDirectory)
                        ? Environment.CurrentDirectory  // Nullの場合はデフォルトディレクトリを使用
                        : initialDirectory
                };

                if (saveFileDialog.ShowDialog() == DialogResult.OK) {
                    var outputPath = saveFileDialog.FileName;
                    var fileInfo = new FileInfo(outputPath);
                    reportWorkbook.SaveAs(fileInfo);
                }
                else {
                    // ユーザーが保存をキャンセルした場合の処理
                    throw new OperationCanceledException("キャンセルされました。");
                }
            }
        }

        // リスト
        public static class ListGeneratorEPPlus {
            // リストを生成するメインメソッド
            public static void GenerateList(ProductInformation productInfo) {
                try {
                    // 1. Excel設定の読み込みとワークブックの準備
                    var configPath = Path.Combine(Environment.CurrentDirectory, "config", "General", "Excel", "ConfigList.xlsm");
                    using var workBook = new ExcelPackage(new FileInfo(configPath));
                    var (targetSheetName, productName, resultRow, workSheetMain) = LoadExcelConfiguration(workBook, productInfo.ProductModel);

                    // 2. 製品情報の設定とExcelへの書き込み
                    var productCellRanges = GetProductCellRanges(workSheetMain, resultRow);
                    var workSheetTemp = workBook.Workbook.Worksheets[targetSheetName] ?? throw new Exception($"テンプレートシート:[{targetSheetName}]が見つかりません。");
                    PopulateProductDetails(workSheetTemp, productInfo, productCellRanges, productName);

                    // 3. データベースから使用済み基板情報を取得
                    var usedSubstrate = GetUsedSubstrateData(productInfo);

                    // 4. 基板情報のExcelへの書き込み
                    UpdateSubstrateDetailsInExcel(workSheetMain, workSheetTemp, resultRow, usedSubstrate);

                    // 5. QRコードの生成と埋め込み
                    if (!string.IsNullOrEmpty(productCellRanges.QrCodeRange)) {
                        GenerateAndEmbedQrCode(workSheetTemp, productInfo, productCellRanges.QrCodeRange);
                    }

                    // 6. Excelファイルの保存と印刷
                    var temporarilyPath = Path.Combine(Environment.CurrentDirectory, "config", "General", "Excel", "temporarilyList.xlsm");
                    SaveAndPrintExcel(workBook, temporarilyPath, targetSheetName);
                } catch (Exception ex) {
                    MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            // Excel設定を読み込むメソッド
            private static (string sheetName, string productName, int resultRow, ExcelWorksheet workSheetMain) LoadExcelConfiguration(ExcelPackage workBook, string productModel) {
                var sheet = workBook.Workbook.Worksheets;
                var targetSheetName = "Sheet1";
                var workSheetMain = sheet[targetSheetName] ?? throw new Exception($"設定シート:[{targetSheetName}]が見つかりません。");

                // セル検索
                var searchAddressResult = workSheetMain.Cells.FirstOrDefault(x => x.Start.Column == 1 && x.Value?.ToString() == productModel)
                    ?? throw new Exception($"Configに品目番号:[{productModel}]が見つかりません。");
                var resultRow = searchAddressResult.Start.Row;

                var sheetName = workSheetMain.Cells[resultRow, 2].Value?.ToString()
                    ?? throw new Exception($"設定シートの行 {resultRow}, 列 2 にシート名が設定されていません。");
                var productName = workSheetMain.Cells[resultRow, 3].Value?.ToString() ?? string.Empty;

                return (sheetName, productName, resultRow, workSheetMain);
            }

            // 製品情報に関連するExcelのセル範囲を取得するヘルパーメソッド
            private static ProductCellRanges GetProductCellRanges(ExcelWorksheet workSheetMain, int resultRow) {
                return new ProductCellRanges {
                    ProductNameRange = workSheetMain.Cells[resultRow, 4].Value?.ToString() ?? string.Empty,
                    ProductNumberRange = workSheetMain.Cells[resultRow, 5].Value?.ToString() ?? string.Empty,
                    OrderNumberRange = workSheetMain.Cells[resultRow, 6].Value?.ToString() ?? string.Empty,
                    RegDateRange = workSheetMain.Cells[resultRow, 7].Value?.ToString() ?? string.Empty,
                    ProductModelConfig = workSheetMain.Cells[resultRow, 8].Value?.ToString() ?? string.Empty,
                    ProductModelRange = workSheetMain.Cells[resultRow, 9].Value?.ToString() ?? string.Empty,
                    QuantityRange = workSheetMain.Cells[resultRow, 10].Value?.ToString() ?? string.Empty,
                    SerialFirstRange = workSheetMain.Cells[resultRow, 11].Value?.ToString() ?? string.Empty,
                    SerialLastRange = workSheetMain.Cells[resultRow, 12].Value?.ToString() ?? string.Empty,
                    CommentRange = workSheetMain.Cells[resultRow, 13].Value?.ToString() ?? string.Empty,
                    QrCodeRange = workSheetMain.Cells[resultRow, 14].Value?.ToString() ?? string.Empty
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
            private static void PopulateProductDetails(ExcelWorksheet workSheetTemp, ProductInformation productInfo, ProductCellRanges ranges, string productName) {

                SetValue(ranges.ProductNameRange, productName);
                SetValue(ranges.ProductNumberRange, productInfo.ProductNumber);
                SetValue(ranges.OrderNumberRange, productInfo.OrderNumber);
                SetValue(ranges.RegDateRange, productInfo.RegDate);
                SetValue(ranges.ProductModelRange, productInfo.ProductModel);
                SetValue(ranges.QuantityRange, productInfo.Quantity.ToString());
                SetValue(ranges.SerialFirstRange, productInfo.SerialFirst);
                SetValue(ranges.SerialLastRange, productInfo.SerialLast);
                SetValue(ranges.CommentRange, productInfo.Comment);

                void SetValue(string? range, string? value) {
                    if (!string.IsNullOrEmpty(range) && !string.IsNullOrEmpty(value)) {
                        workSheetTemp.Cells[range].Value = value;
                    }
                }
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
            private static void UpdateSubstrateDetailsInExcel(ExcelWorksheet workSheetMain, ExcelWorksheet workSheetTemp, int resultRow, List<(string SubstrateModel, List<string> SubstrateNumbers, List<int> Decreases)> usedSubstrate) {
                foreach (var (substrateModel, substrateNumbers, decreases) in usedSubstrate) {
                    var targetRow = resultRow;
                    var searchValue = substrateModel;
                    var foundColumn = 0;

                    var searchAddressResult2 = workSheetMain.Cells
                        .Where(x => x.Start.Row == targetRow && x.Value?.ToString() == searchValue)
                        .FirstOrDefault();

                    foundColumn = searchAddressResult2 is not null
                        ? searchAddressResult2.Start.Column
                        : throw new Exception($"Configに基板モデル:[{searchValue}]が見つかりません。");

                    var mainCellValue = workSheetMain.Cells[resultRow, foundColumn + 1].Value?.ToString()
                        ?? throw new Exception($"Configシートの行 {resultRow}, 列 {foundColumn + 1} にセル範囲が設定されていません。");

                    var tempCellValue = string.Join("   ", substrateNumbers.Select((substrateNumber, k) => $"{substrateNumber}({decreases[k]})"));

                    workSheetTemp.Cells[mainCellValue].Value = tempCellValue;
                }
            }

            // QRコードを生成し、Excelシートに埋め込むメソッド
            private static void GenerateAndEmbedQrCode(ExcelWorksheet workSheetTemp, ProductInformation productInfo, string qrCodeRange) {
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
                var picture = workSheetTemp.Drawings.AddPicture("QR", qrBitmap);
                // QRコードのセル範囲の左上隅に画像を配置
                picture.SetPosition(workSheetTemp.Cells[qrCodeRange].Start.Row - 1, 0, workSheetTemp.Cells[qrCodeRange].Start.Column - 1, 0);
            }

            // Excelファイルを保存し、印刷処理を行うメソッド
            private static void SaveAndPrintExcel(ExcelPackage workBook, string temporarilyPath, string sheetName) {
                var fileInfo = new FileInfo(temporarilyPath);
                workBook.SaveAs(fileInfo);

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

        // チェックシート (EPPlusでは画像のセル配置が難しいため、コメントアウト ブック作成にNPOIを使用)
        public static class CheckSheetGeneratorEPPlus {
            //// チェックシート生成データを保持するクラス
            //public class CheckSheetConfigData {
            //    public string? ProductModelRange { get; set; }
            //    public string? ProductNumberRange { get; set; }
            //    public string? OrderNumberRange { get; set; }
            //    public string? QuantityRange { get; set; }
            //    public string? SerialFirstRange { get; set; }
            //    public string? SerialLastRange { get; set; }
            //    public string? RegDateRange { get; set; }
            //    public string? DateFormat { get; set; }
            //    public string? RegTemperatureRange { get; set; }
            //    public string? RegHumidityRange { get; set; }
            //    public List<string> SheetNames { get; set; } = [];
            //}
            //// Excelチェックシートを生成し、データを書き込み、印刷します。
            //public static void GenerateCheckSheet(ProductInformation productInfo) {
            //    try {
            //        // 設定ファイルのパスを構築
            //        var configPath = Path.Combine(Environment.CurrentDirectory, "config", "General", "Excel", "ConfigCheckSheet.xlsm");
            //        var temporarilyPath = Path.Combine(Environment.CurrentDirectory, "config", "General", "Excel", "temporarilyCheckSheet.xlsm");

            //        // 1. 設定ファイルの読み込みとメインシートの取得
            //        // EPPlusを使用して設定ファイルを読み込み、必要な情報を抽出
            //        var excelData = LoadAndExtractConfig(configPath, productInfo);

            //        // 2. 温度・湿度入力ダイアログの表示と値の取得
            //        (var temperature, var humidity) = GetTemperatureAndHumidity(excelData);

            //        // 3. 日付のフォーマット
            //        var formattedDate = FormatDate(productInfo.RegDate, excelData.DateFormat);

            //        // 4. EPPlusでExcelファイルを編集
            //        // 設定ファイルからEPPlusワークブックを作成
            //        using (var workBookEPPlus = WorkbookFactory.Create(configPath)) {
            //            // セル参照をCellReferenceオブジェクトとして作成
            //            var cellReferences = CreateCellReferences(excelData);

            //            // 各シートに対して値を書き込む
            //            PopulateExcelSheets(workBookEPPlus, productInfo, cellReferences, temperature, humidity, formattedDate, excelData.SheetNames);

            //            // 不要なシートを非表示にする
            //            HideSheets(workBookEPPlus, excelData.SheetNames);

            //            // ブックを保存
            //            SaveWorkbook(workBookEPPlus, temporarilyPath);
            //        }

            //        // 5. Excel Interopを使用して印刷
            //        PrintExcelFile(temporarilyPath);
            //    } catch (Exception ex) {
            //        // エラーメッセージを表示
            //        MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //    }
            //}

            //// Excel設定ファイルを読み込み、メインシートから設定データを抽出します。
            //private static CheckSheetConfigData LoadAndExtractConfig(string configPath, ProductInformation productInfo) {
            //    if (!File.Exists(configPath)) {
            //        throw new FileNotFoundException($"設定ファイルが見つかりません: {configPath}");
            //    }

            //    using FileStream fileStream = new(configPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            //    using var workBook = new ExcelPackage(fileStream);

            //    // 既存ワークシートを取得
            //    var sheet = workBook.Workbook.Worksheets;
            //    var targetSheetName = "Sheet1";
            //    var workSheetMain = sheet[targetSheetName];

            //    // セル検索: ProductModelに基づいて行を特定
            //    var searchAddressResult = workSheetMain.Cells.FirstOrDefault(x => x.Start.Column == 1 && x.Value?.ToString() == productInfo.ProductModel)
            //        ?? throw new Exception($"Configに品目番号:[{productInfo.ProductModel}]が見つかりません。");
            //    var resultRow = searchAddressResult.Start.Row;

            //    // ワークシートのセルから値を取得し、ExcelConfigDataオブジェクトに格納
            //    var excelData = new CheckSheetConfigData {
            //        ProductModelRange = workSheetMain.Cells[resultRow, 3].Value?.ToString() ?? string.Empty,
            //        ProductNumberRange = workSheetMain.Cells[resultRow, 4].Value?.ToString() ?? string.Empty,
            //        OrderNumberRange = workSheetMain.Cells[resultRow, 5].Value?.ToString() ?? string.Empty,
            //        QuantityRange = workSheetMain.Cells[resultRow, 6].Value?.ToString() ?? string.Empty,
            //        SerialFirstRange = workSheetMain.Cells[resultRow, 7].Value?.ToString() ?? string.Empty,
            //        SerialLastRange = workSheetMain.Cells[resultRow, 8].Value?.ToString() ?? string.Empty,
            //        RegDateRange = workSheetMain.Cells[resultRow, 9].Value?.ToString() ?? string.Empty,
            //        DateFormat = workSheetMain.Cells[resultRow, 10].Value?.ToString() ?? string.Empty,
            //        RegTemperatureRange = workSheetMain.Cells[resultRow, 11].Value?.ToString() ?? string.Empty,
            //        RegHumidityRange = workSheetMain.Cells[resultRow, 12].Value?.ToString() ?? string.Empty,
            //        SheetNames = [.. Enumerable.Range(13, 20)
            //            .Select(column => workSheetMain.Cells[resultRow, column].Value?.ToString() ?? string.Empty)
            //            .TakeWhile(sheetName => !string.IsNullOrEmpty(sheetName))]
            //    };

            //    return excelData.SheetNames.Count == 0 ? throw new Exception("対象シートがConfigファイルに設定されていません。") : excelData;
            //}

            //// 温度と湿度の入力ダイアログを表示し、ユーザーからの値を取得します。
            //private static (string temperature, string humidity) GetTemperatureAndHumidity(CheckSheetConfigData excelData) {
            //    var temperature = string.Empty;
            //    var humidity = string.Empty;

            //    // 温度セルか湿度セルが設定されている場合、ダイアログを表示
            //    if (!string.IsNullOrEmpty(excelData.RegTemperatureRange) || !string.IsNullOrEmpty(excelData.RegHumidityRange)) {
            //        var dialog = new InputDialog1();
            //        var result = dialog.ShowDialog();
            //        if (result != DialogResult.OK) {
            //            // キャンセルされた場合は処理を中断
            //            throw new OperationCanceledException("ユーザーによって温度・湿度入力がキャンセルされました。");
            //        }
            //        temperature = dialog.Temperature;
            //        humidity = dialog.Humidity;
            //    }
            //    return (temperature, humidity);
            //}

            //// 日付を指定されたフォーマットで文字列に変換します。
            //private static string FormatDate(string dateString, string? dateFormat) {
            //    var formattedDate = string.Empty;
            //    if (DateTime.TryParse(dateString, out var date)) {
            //        formattedDate = dateFormat switch {
            //            "1" => date.ToString("yyyy年MM月dd日"),
            //            "2" => date.ToString("yyyy-MM-dd"),
            //            _ => formattedDate
            //        };
            //    }
            //    return formattedDate;
            //}

            //// ExcelConfigDataからEPPlusのCellReferenceオブジェクトを作成します。
            //private static dynamic CreateCellReferences(CheckSheetConfigData excelData) {
            //    // セル参照をCellReferenceオブジェクトとして作成
            //    return new {
            //        ProductModel = string.IsNullOrEmpty(excelData.ProductModelRange) ? null : new CellReference(excelData.ProductModelRange),
            //        ProductNumber = string.IsNullOrEmpty(excelData.ProductNumberRange) ? null : new CellReference(excelData.ProductNumberRange),
            //        OrderNumber = string.IsNullOrEmpty(excelData.OrderNumberRange) ? null : new CellReference(excelData.OrderNumberRange),
            //        Quantity = string.IsNullOrEmpty(excelData.QuantityRange) ? null : new CellReference(excelData.QuantityRange),
            //        SerialFirst = string.IsNullOrEmpty(excelData.SerialFirstRange) ? null : new CellReference(excelData.SerialFirstRange),
            //        SerialLast = string.IsNullOrEmpty(excelData.SerialLastRange) ? null : new CellReference(excelData.SerialLastRange),
            //        RegDate = string.IsNullOrEmpty(excelData.RegDateRange) ? null : new CellReference(excelData.RegDateRange),
            //        RegTemperature = string.IsNullOrEmpty(excelData.RegTemperatureRange) ? null : new CellReference(excelData.RegTemperatureRange),
            //        RegHumidity = string.IsNullOrEmpty(excelData.RegHumidityRange) ? null : new CellReference(excelData.RegHumidityRange),
            //    };
            //}

            //// 指定されたEPPlusワークブックの各シートに製品情報を書き込みます。
            //private static void PopulateExcelSheets(IWorkbook workBookEPPlus, ProductInformation productInfo, dynamic cellReferences, string temperature, string humidity, string formattedDate, List<string> sheetNames) {
            //    foreach (var sheetName in sheetNames) {
            //        var sheetEPPlus = workBookEPPlus.GetSheet(sheetName) ?? throw new Exception($"シート[{sheetName}]が見つかりません。");
            //        sheetEPPlus.ForceFormulaRecalculation = true; // 数式の再計算を強制

            //        // 各セルに値を書き込む
            //        WriteCellValue(sheetEPPlus, cellReferences.ProductModel, productInfo.ProductModel);
            //        WriteCellValue(sheetEPPlus, cellReferences.ProductNumber, productInfo.ProductNumber);
            //        WriteCellValue(sheetEPPlus, cellReferences.OrderNumber, productInfo.OrderNumber);
            //        WriteCellValue(sheetEPPlus, cellReferences.Quantity, productInfo.Quantity.ToString());
            //        WriteCellValue(sheetEPPlus, cellReferences.SerialFirst, productInfo.SerialFirst);
            //        WriteCellValue(sheetEPPlus, cellReferences.SerialLast, productInfo.SerialLast);
            //        WriteCellValue(sheetEPPlus, cellReferences.RegDate, formattedDate);
            //        WriteCellValue(sheetEPPlus, cellReferences.RegTemperature, temperature);
            //        WriteCellValue(sheetEPPlus, cellReferences.RegHumidity, humidity);
            //    }
            //}

            //// 指定されたシートのセルに値を書き込みます。
            //private static void WriteCellValue(ISheet sheet, CellReference? reference, string value) {
            //    if (reference is null) {
            //        return; // 参照が nullの場合は何もしない
            //    }

            //    // 指定した行を取得できない時はエラーとならないよう新規作成している
            //    var row = sheet.GetRow(reference.Row) ?? sheet.CreateRow(reference.Row);
            //    // 一行上の処理の列版
            //    var cell = row.GetCell(reference.Col) ?? row.CreateCell(reference.Col);
            //    cell.SetCellValue(value);
            //}

            //// 指定されたワークブックの不要なシートを非表示にします。
            //private static void HideSheets(IWorkbook workBookEPPlus, List<string> sheetsToKeep) {
            //    // 非表示にするシートのインデックスを特定
            //    var allSheetNames = new List<string>();
            //    for (var i = 0; i < workBookEPPlus.NumberOfSheets; i++) {
            //        allSheetNames.Add(workBookEPPlus.GetSheetName(i));
            //    }

            //    var sheetIndicesToHide = allSheetNames
            //        .Select((name, index) => new { Name = name, Index = index })
            //        .Where(sheet => !sheetsToKeep.Contains(sheet.Name))
            //        .Select(sheet => sheet.Index)
            //        .ToList();

            //    // "Sheet1" は常に非表示にする
            //    var sheet1Index = workBookEPPlus.GetSheetIndex("Sheet1");
            //    if (sheet1Index != -1 && !sheetIndicesToHide.Contains(sheet1Index)) {
            //        sheetIndicesToHide.Add(sheet1Index);
            //    }

            //    // シートを非表示に設定
            //    foreach (var sheetIndex in sheetIndicesToHide) {
            //        workBookEPPlus.SetSheetVisibility(sheetIndex, SheetVisibility.VeryHidden);
            //    }
            //}

            //// EPPlusワークブックをファイルに保存します。
            //private static void SaveWorkbook(IWorkbook workBookEPPlus, string outputPath) {
            //    try {
            //        using var fs = new FileStream(outputPath, FileMode.Create);
            //        workBookEPPlus.Write(fs);
            //    } catch (IOException ex) {
            //        throw new IOException($"Excelファイルの保存に失敗しました: {outputPath}. 詳細: {ex.Message}", ex);
            //    }
            //}

            //// Excel Interopを使用して指定されたExcelファイルを開きます。
            //private static void PrintExcelFile(string filePath) {
            //    Microsoft.Office.Interop.Excel.Application? xlApp = null;
            //    Microsoft.Office.Interop.Excel.Workbooks? xlBooks = null;
            //    Microsoft.Office.Interop.Excel.Workbook? xlBook = null;

            //    try {
            //        xlApp = new Microsoft.Office.Interop.Excel.Application {
            //            Visible = true // Excelウィンドウを表示します。
            //        };

            //        xlBooks = xlApp.Workbooks;
            //        xlBook = xlBooks.Open(filePath, ReadOnly: true);

            //        //// ワークブックを閉じる (falseは変更を保存しない)
            //        //xlBook.Close(false);
            //        //// Excelアプリケーションを終了
            //        //xlApp.Quit();
            //    } catch (Exception ex) {
            //        throw new Exception($"エラーが発生しました。詳細: {ex.Message}", ex);
            //    } finally {
            //        // COMオブジェクトの解放
            //        ReleaseComObject(xlBook);
            //        ReleaseComObject(xlBooks);
            //        ReleaseComObject(xlApp);

            //        static void ReleaseComObject(object? comObj) {
            //            if (comObj is not null) {
            //                System.Runtime.InteropServices.Marshal.ReleaseComObject(comObj);
            //            }
            //        }
            //    }
            //}
        }

        // 基板設定を開く
        public static class SubstrateInformationEPPlus {
            // 基板設定ファイルを開くメソッド
            public static void OpenSubstrateInformationEPPlus(ProductInformation productInfo) {
                try {
                    // 1. 設定ファイルを読み込み、レポート設定を取得
                    var configWorkbook = LoadConfigWorkbook();
                    var (filePath, fileName) = GetSubstrateConfig(configWorkbook, productInfo.SubstrateModel);

                    OpenExcel(filePath, fileName);

                } catch (Exception ex) {
                    // エラーメッセージを表示
                    MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            // 設定Excelワークブックを読み込む
            private static ExcelPackage LoadConfigWorkbook() {
                var configPath = Path.Combine(Environment.CurrentDirectory, "config", "General", "Excel", "ConfigSubstrateInformation.xlsm");
                if (!File.Exists(configPath)) {
                    throw new FileNotFoundException($"設定ファイルが見つかりません: {configPath}");
                }

                try {
                    FileStream fileStreamConfig = new(configPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    return new ExcelPackage(fileStreamConfig);
                } catch (Exception ex) {
                    throw new Exception($"設定ファイルの読み込み中にエラーが発生しました: {ex.Message}", ex);
                }
            }

            // 設定ワークブックから製品に対応する設定を抽出する
            private static (string filePath, string sheetName) GetSubstrateConfig(ExcelPackage configWorkbook, string substrateModel) {
                var workSheetMain = configWorkbook.Workbook.Worksheets["Sheet1"] ?? throw new Exception("設定ファイルのシートに Sheet1 が見つかりません。");

                // セル検索
                var searchAddressResult = workSheetMain.Cells.FirstOrDefault(x => x.Start.Column == 1 && x.Value?.ToString() == substrateModel)
                    ?? throw new Exception($"Configに型式:[{substrateModel}]が見つかりません。");
                var searchAddressResultRow = searchAddressResult.Start.Row;

                // ディレクトリパスとファイル名を取得(ディレクトリパスが空の場合は、Configの2行目の値を使用する)
                var directoryPath = workSheetMain.Cells[searchAddressResultRow, 4].Value?.ToString()?.Trim('"') ?? workSheetMain.Cells[2, 4].Value?.ToString()?.Trim('"') ?? string.Empty;
                var fileName = workSheetMain.Cells[searchAddressResultRow, 5].Value?.ToString()?.Trim('"') ?? string.Empty;
                if (string.IsNullOrEmpty(fileName)) { throw new Exception($"設定ファイルの型式 {substrateModel} の ファイル名 が空です。"); }
                var filePath = Path.Combine(directoryPath, fileName);
                if (!File.Exists(filePath)) { throw new FileNotFoundException($"指定されたファイルが存在しません: {filePath}"); }

                var sheetName = workSheetMain.Cells[searchAddressResultRow, 6].Value?.ToString()?.Trim('"') ?? string.Empty;

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
    }
}
