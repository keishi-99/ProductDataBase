using ClosedXML.Excel;
using Microsoft.Data.Sqlite;
using ProductDatabase.Other;
using ZXing;
using ZXing.QrCode;
using ZXing.QrCode.Internal;
using ZXing.Windows.Compatibility;
using static ProductDatabase.MainWindow;

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

                    // 基板情報の入力がある場合のみ処理を実行
                    if (reportConfig.HasSubstrateInput) {
                        // 4. データベースから使用済み基板情報を取得
                        var usedSubstrate = GetUsedSubstrateData(productInfo);

                        // 5. 基板情報のExcelへの書き込み
                        UpdateSubstrateDetailsInExcel(reportWorkbook, reportConfig, usedSubstrate);
                    }

                    // 6. レポートを保存
                    SaveReport(reportWorkbook, productInfo, reportConfig);

                    MessageBox.Show("成績書が正常に生成されました。", "完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
                } catch (Exception ex) {
                    // エラーメッセージをユーザーに表示
                    MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            // 設定Excelワークブックを読み込む
            private static XLWorkbook LoadConfigWorkbook() {
                try {
                    var configPath = Path.Combine(Environment.CurrentDirectory, "config", "General", "Excel", "ConfigReport.xlsm");
                    if (!File.Exists(configPath)) {
                        throw new FileNotFoundException($"設定ファイルが見つかりません: {configPath}");
                    }
                    return new XLWorkbook(configPath);
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

                var reslutRow = targetCell.Address.RowNumber;

                // --- シート名 ---
                var sheetName = configSheet.Cell(reslutRow, 5).GetString();
                if (string.IsNullOrEmpty(sheetName)) {
                    throw new Exception($"設定ファイルのシート名が空です。");
                }

                // --- パスとファイル名の取得 ---
                var directoryPath = configSheet.Cell(reslutRow, 3).GetString().Trim('"');
                if (string.IsNullOrEmpty(directoryPath)) {
                    directoryPath = configSheet.Cell(2, 3).GetString().Trim('"');
                }

                if (!Directory.Exists(directoryPath)) {
                    throw new DirectoryNotFoundException($"指定されたフォルダが存在しません: {directoryPath}");
                }

                var saveDirectory = configSheet.Cell(reslutRow, 12).GetString().Trim('"');
                if (string.IsNullOrEmpty(saveDirectory)) {
                    saveDirectory = configSheet.Cell(2, 12).GetString().Trim('"');
                }

                var searchName = configSheet.Cell(reslutRow, 4).GetString().Trim('"');
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

                var targetRowCells = configSheet.Row(reslutRow).Cells(startColumn, endColumn);
                bool hasValue = targetRowCells.Any(c => !string.IsNullOrWhiteSpace(c.GetValue<string>()));

                // --- ReportConfig を構築 ---
                return new ReportConfigClosedXml {
                    Sheet = configSheet,
                    ResultRow = reslutRow,
                    DirectoryPath = directoryPath,
                    FilePath = filePath,
                    FileName = fileName,
                    FileExtension = fileExtension,
                    SearchFileName = searchName,
                    SheetName = sheetName,
                    ProductNumberRange = configSheet.Cell(reslutRow, 6).GetString(),
                    OrderNumberRange = configSheet.Cell(reslutRow, 7).GetString(),
                    QuantityRange = configSheet.Cell(reslutRow, 8).GetString(),
                    SerialFirstRange = configSheet.Cell(reslutRow, 9).GetString(),
                    SerialLastRange = configSheet.Cell(reslutRow, 10).GetString(),
                    ProductModelRange = configSheet.Cell(reslutRow, 11).GetString(),
                    SaveDirectory = saveDirectory,
                    HasSubstrateInput = hasValue,
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
            private static XLWorkbook LoadReportTemplate(ReportConfigClosedXml config) {
                var filePath = config.FilePath;

                try {
                    return new XLWorkbook(filePath);
                } catch (Exception ex) {
                    throw new Exception($"レポートテンプレートの読み込み中にエラーが発生しました: {ex.Message}", ex);
                }
            }

            // レポートシートに製品情報を挿入する
            private static void PopulateReportSheet(XLWorkbook reportWorkbook, ProductInformation productInfo, ReportConfigClosedXml config) {
                var workSheetTemp = reportWorkbook.Worksheet(config.SheetName)
                    ?? throw new Exception($"シート '{config.SheetName}' が見つかりません。");

                var productNumber = productInfo.ProductNumber.Split('-')[0] ?? string.Empty;
                SetValue(config.ProductNumberRange, productNumber);
                SetValue(config.OrderNumberRange, productInfo.OrderNumber);
                SetValue(config.QuantityRange, productInfo.Quantity.ToString());
                SetValue(config.SerialFirstRange, productInfo.SerialFirst);
                SetValue(config.SerialLastRange, productInfo.SerialLast);
                SetValue(config.ProductModelRange, productInfo.ProductModel);

                void SetValue(string? range, string? value) {
                    if (!string.IsNullOrEmpty(range) && !string.IsNullOrEmpty(value)) {
                        workSheetTemp.Cell(range).Value = value;
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

            // 変更されたレポートをファイルに保存する
            private static void SaveReport(XLWorkbook reportWorkbook, ProductInformation productInfo, ReportConfigClosedXml config) {
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
                    reportWorkbook.SaveAs(outputPath);
                }
                else {
                    // ユーザーが保存をキャンセルした場合の処理
                    throw new OperationCanceledException("キャンセルされました。");
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
            public static void GenerateList(ProductInformation productInfo) {
                try {
                    // 1. Excel設定の読み込みとワークブックの準備
                    var configPath = Path.Combine(Environment.CurrentDirectory, "config", "General", "Excel", "ConfigList.xlsm");
                    if (!File.Exists(configPath)) {
                        throw new FileNotFoundException($"設定ファイルが見つかりません: {configPath}");
                    }
                    using var workBook = new XLWorkbook(configPath);
                    var (targetSheetName, productName, resultRow, configSheet) = LoadExcelConfiguration(workBook, productInfo.ProductModel);

                    // 2. 製品情報の設定とExcelへの書き込み
                    var productCellRanges = GetProductCellRanges(configSheet, resultRow);
                    var targetSheet = workBook.Worksheet(targetSheetName) ?? throw new Exception($"テンプレートシート:[{targetSheetName}]が見つかりません。");
                    PopulateProductDetails(targetSheet, productInfo, productCellRanges, productName);

                    // 3. データベースから使用済み基板情報を取得
                    var usedSubstrate = GetUsedSubstrateData(productInfo);

                    // 4. 基板情報のExcelへの書き込み
                    UpdateSubstrateDetailsInExcel(configSheet, targetSheet, resultRow, usedSubstrate);

                    // 5. QRコードの生成と埋め込み
                    if (!string.IsNullOrEmpty(productCellRanges.QrCodeRange)) {
                        GenerateAndEmbedQrCode(targetSheet, productInfo, productCellRanges.QrCodeRange);
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
            private static void PopulateProductDetails(IXLWorksheet targetSheet, ProductInformation productInfo, ProductCellRanges ranges, string productName) {

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
                        targetSheet.Cell(range).Value = value;
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
            private static void GenerateAndEmbedQrCode(IXLWorksheet targetSheet, ProductInformation productInfo, string qrCodeRange) {
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
                //var picture = workSheetTemp.Drawings.AddPicture("QR", qrBitmap);
                using var ms = new MemoryStream();
                qrBitmap.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                ms.Position = 0; // ストリームの先頭に戻す

                var picture = targetSheet.AddPicture(ms).MoveTo(targetSheet.Cell(qrCodeRange));
            }

            // Excelファイルを保存し、印刷処理を行うメソッド
            private static void SaveAndPrintExcel(XLWorkbook workBook, string temporarilyPath, string sheetName) {
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
                public List<string> SheetNames { get; set; } = [];
            }

            // Excelチェックシートを生成し、データを書き込み、印刷します。
            public static void GenerateCheckSheet(ProductInformation productInfo) {
                try {
                    // 設定ファイルのパスを構築
                    var configPath = Path.Combine(Environment.CurrentDirectory, "config", "General", "Excel", "ConfigCheckSheet.xlsm");
                    var temporarilyPath = Path.Combine(Environment.CurrentDirectory, "config", "General", "Excel", "temporarilyCheckSheet.xlsm");

                    // 1. 設定ファイルの読み込みとメインシートの取得
                    var (configBook, excelData) = LoadAndExtractConfig(configPath, productInfo);

                    // 2. 温度・湿度入力ダイアログの表示と値の取得
                    (var temperature, var humidity) = GetTemperatureAndHumidity(excelData);

                    // 3. 日付のフォーマット
                    var formattedDate = FormatDate(productInfo.RegDate, excelData.DateFormat);

                    // 4. EPPlusでExcelファイルを編集

                    // 各シートに対して値を書き込む
                    PopulateExcelSheets(configBook, productInfo, temperature, humidity, formattedDate, excelData);

                    // 不要なシートを非表示にする
                    HideSheets(configBook, excelData.SheetNames);

                    // ブックを保存
                    SaveWorkbook(configBook, temporarilyPath);

                    // 5. Excel Interopを使用して印刷
                    PrintExcelFile(temporarilyPath);
                } catch (Exception ex) {
                    // エラーメッセージを表示
                    MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            // Excel設定ファイルを読み込み、メインシートから設定データを抽出します。
            private static (XLWorkbook workBook, CheckSheetConfigData excelData) LoadAndExtractConfig(string configPath, ProductInformation productInfo) {
                if (!File.Exists(configPath)) {
                    throw new FileNotFoundException($"設定ファイルが見つかりません: {configPath}");
                }

                var workBook = new XLWorkbook(configPath);
                var productModel = productInfo.ProductModel;

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
                    SheetNames = [.. Enumerable.Range(13, 20)
                        .Select(column => configSheet.Cell(resultRow, column).GetString())
                        .TakeWhile(sheetName => !string.IsNullOrEmpty(sheetName))]
                };

                return (workBook, excelData);
            }

            // 温度と湿度の入力ダイアログを表示し、ユーザーからの値を取得します。
            private static (string temperature, string humidity) GetTemperatureAndHumidity(CheckSheetConfigData excelData) {
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

            // 指定されたEPPlusワークブックの各シートに製品情報を書き込みます。
            private static void PopulateExcelSheets(XLWorkbook configBook, ProductInformation productInfo, string temperature, string humidity, string formattedDate, CheckSheetConfigData excelData) {
                var sheetNames = excelData.SheetNames;
                foreach (var sheetName in sheetNames) {
                    var targetSheet = configBook.Worksheet(sheetName) ?? throw new Exception($"シート[{sheetName}]が見つかりません。");

                    // 各セルに値を書き込む
                    WriteCellValue(targetSheet, excelData.ProductModelRange, productInfo.ProductModel);
                    WriteCellValue(targetSheet, excelData.ProductNumberRange, productInfo.ProductNumber);
                    WriteCellValue(targetSheet, excelData.OrderNumberRange, productInfo.OrderNumber);
                    WriteCellValue(targetSheet, excelData.QuantityRange, productInfo.Quantity.ToString());
                    WriteCellValue(targetSheet, excelData.SerialFirstRange, productInfo.SerialFirst);
                    WriteCellValue(targetSheet, excelData.SerialLastRange, productInfo.SerialLast);
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
            private static void HideSheets(XLWorkbook configBook, List<string> sheetsToKeep) {

                // すべてのシート名を取得
                var allSheetNames = configBook.Worksheets.Select(ws => ws.Name).ToList();

                // 非表示にする対象シートを抽出
                var sheetsToHide = allSheetNames
                    .Where(name => !sheetsToKeep.Contains(name))
                    .ToList();

                // "Sheet1" は常に非表示にする
                if (!sheetsToHide.Contains("Sheet1") && configBook.Worksheets.Any(ws => ws.Name == "Sheet1")) {
                    sheetsToHide.Add("Sheet1");
                }

                // 非表示に設定
                foreach (var sheetName in sheetsToHide) {
                    var ws = configBook.Worksheets.Worksheet(sheetName);
                    ws.Visibility = XLWorksheetVisibility.VeryHidden;
                }
            }

            // EPPlusワークブックをファイルに保存します。
            private static void SaveWorkbook(XLWorkbook configBook, string outputPath) {
                try {
                    configBook.SaveAs(outputPath);
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

        // 基板設定を開く
        public static class SubstrateInformationClosedXml {
            // 基板設定ファイルを開くメソッド
            public static void OpenSubstrateInformationClosedXml(ProductInformation productInfo) {
                try {
                    // 1. 設定ファイルを読み込み、レポート設定を取得
                    var configBook = LoadConfigWorkbook();
                    var (filePath, fileName) = GetSubstrateConfig(configBook, productInfo.SubstrateModel);

                    OpenExcel(filePath, fileName);

                } catch (Exception ex) {
                    // エラーメッセージを表示
                    MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            // 設定Excelワークブックを読み込む
            private static XLWorkbook LoadConfigWorkbook() {
                try {
                    var configPath = Path.Combine(Environment.CurrentDirectory, "config", "General", "Excel", "ConfigSubstrateInformation.xlsm");
                    if (!File.Exists(configPath)) {
                        throw new FileNotFoundException($"設定ファイルが見つかりません: {configPath}");
                    }

                    return new XLWorkbook(configPath);
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

                var reslutRow = targetCell.Address.RowNumber;

                // ディレクトリパスとファイル名を取得(ディレクトリパスが空の場合は、Configの2行目の値を使用する)
                var directoryPath = configSheet.Cell(reslutRow, 4).GetString().Trim('"');
                if (string.IsNullOrEmpty(directoryPath)) {
                    directoryPath = configSheet.Cell(2, 4).GetString().Trim('"');
                }

                var fileName = configSheet.Cell(reslutRow, 5).GetString().Trim('"');
                if (string.IsNullOrEmpty(fileName)) {
                    throw new Exception($"設定ファイルの型式 {substrateModel} の シート名 が空です。");
                }

                var filePath = Path.Combine(directoryPath, fileName);
                if (!File.Exists(filePath)) {
                    throw new FileNotFoundException($"指定されたファイルが存在しません: {filePath}");
                }

                var sheetName = configSheet.Cell(reslutRow, 6).GetString().Trim('"');

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
