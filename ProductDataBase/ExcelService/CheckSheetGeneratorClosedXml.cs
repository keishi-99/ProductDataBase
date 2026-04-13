using ClosedXML.Excel;
using ProductDatabase.Models;
using ProductDatabase.Other;

namespace ProductDatabase.ExcelService {
    // チェックシートのExcel生成を担当するクラス
    internal static class CheckSheetGeneratorClosedXml {

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

        // UIスレッドで呼ぶ: Config読み込み + InputDialog（温度・湿度）
        // OperationCanceledException（InputDialog キャンセル）は null を返す
        // その他の例外は呼び出し元に伝播する
        public static (CheckSheetConfigData configData, string temperature, string humidity)?
            PrepareCheckSheet(ProductMaster productMaster) {
            var configPath = Path.Combine(Environment.CurrentDirectory, "config", "General", "Excel", "ConfigCheckSheet.xlsm");

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

        // STAスレッドで呼ぶ: ClosedXML 処理 + COM Interop（Excel 起動）
        // 例外はすべて呼び出し元に伝播する
        public static void ExecuteCheckSheet(
            ProductMaster productMaster,
            ProductRegisterWork productRegisterWork,
            CheckSheetConfigData excelData,
            string temperature,
            string humidity) {
            var configPath      = Path.Combine(Environment.CurrentDirectory, "config", "General", "Excel", "ConfigCheckSheet.xlsm");
            var temporarilyPath = Path.Combine(Environment.CurrentDirectory, "config", "General", "Excel", "temporarilyCheckSheet.xlsm");

            var formattedDate = FormatDate(productRegisterWork.RegDate, excelData.DateFormat);

            using var configFs   = new FileStream(configPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var configBook = new XLWorkbook(configFs);

            XLWorkbook targetBook;
            FileStream? targetFs = null;

            if (string.IsNullOrWhiteSpace(excelData.BaseFilePath)) {
                targetBook = configBook;
            } else {
                targetFs   = new FileStream(excelData.BaseFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                targetBook = new XLWorkbook(targetFs);
            }

            try {
                PopulateExcelSheets(targetBook, productMaster, productRegisterWork, temperature, humidity, formattedDate, excelData);
                HideSheets(targetBook, excelData.SheetNames);
                SaveWorkbook(targetBook, temporarilyPath);
            } finally {
                if (targetBook != configBook)
                    targetBook.Dispose();
                targetFs?.Dispose();
            }

            PrintExcelFile(temporarilyPath);
        }

        // Excel設定ファイルを読み込み設定データを抽出する
        private static (XLWorkbook workBook, CheckSheetConfigData excelData) LoadAndExtractConfig(XLWorkbook workBook, ProductMaster productMaster) {
            var configSheet = workBook.Worksheet("Sheet1");

            var targetCell = configSheet
                .Column(1)
                .CellsUsed()
                .FirstOrDefault(c => c.GetValue<string>() == productMaster.ProductModel)
                ?? throw new Exception($"Configに品目番号:[{productMaster.ProductModel}]が見つかりません。");

            var resultRow = targetCell.Address.RowNumber;

            var excelData = new CheckSheetConfigData {
                ProductModelRange    = configSheet.Cell(resultRow, 3).GetString(),
                ProductNumberRange   = configSheet.Cell(resultRow, 4).GetString(),
                OrderNumberRange     = configSheet.Cell(resultRow, 5).GetString(),
                QuantityRange        = configSheet.Cell(resultRow, 6).GetString(),
                SerialFirstRange     = configSheet.Cell(resultRow, 7).GetString(),
                SerialLastRange      = configSheet.Cell(resultRow, 8).GetString(),
                RegDateRange         = configSheet.Cell(resultRow, 9).GetString(),
                DateFormat           = configSheet.Cell(resultRow, 10).GetString(),
                RegTemperatureRange  = configSheet.Cell(resultRow, 11).GetString(),
                RegHumidityRange     = configSheet.Cell(resultRow, 12).GetString(),
                BaseFilePath         = configSheet.Cell(resultRow, 13).GetString(),
                SheetNames = [.. Enumerable.Range(14, 21)
                    .Select(column => configSheet.Cell(resultRow, column).GetString())
                    .TakeWhile(sheetName => !string.IsNullOrEmpty(sheetName))]
            };

            return (workBook, excelData);
        }

        // 温度と湿度の入力ダイアログを表示しユーザーからの値を取得する
        private static (string, string) GetTemperatureAndHumidity(CheckSheetConfigData excelData) {
            if (string.IsNullOrEmpty(excelData.RegTemperatureRange) && string.IsNullOrEmpty(excelData.RegHumidityRange)) {
                return (string.Empty, string.Empty);
            }

            using var dialog = new InputDialog1();
            if (dialog.ShowDialog() != DialogResult.OK) {
                throw new OperationCanceledException("ユーザーによって温度・湿度入力がキャンセルされました。");
            }

            return (dialog.Temperature, dialog.Humidity);
        }

        // 日付を指定フォーマットで文字列に変換する
        private static string FormatDate(string dateString, string? dateFormat) {
            if (!DateTime.TryParse(dateString, out var date)) {
                return string.Empty;
            }

            return dateFormat switch {
                "1" => date.ToString("yyyy年MM月dd日"),
                "2" => date.ToString("yyyy-MM-dd"),
                _   => string.Empty
            };
        }

        // 指定ワークブックの各シートに製品情報を書き込む
        private static void PopulateExcelSheets(XLWorkbook targetBook, ProductMaster productMaster, ProductRegisterWork productRegisterWork, string temperature, string humidity, string formattedDate, CheckSheetConfigData excelData) {
            foreach (var sheetName in excelData.SheetNames) {
                var targetSheet = targetBook.Worksheet(sheetName) ?? throw new Exception($"シート[{sheetName}]が見つかりません。");

                WriteCellValue(targetSheet, excelData.ProductModelRange,   productMaster.ProductModel);
                WriteCellValue(targetSheet, excelData.ProductNumberRange,  productRegisterWork.ProductNumber);
                WriteCellValue(targetSheet, excelData.OrderNumberRange,    productRegisterWork.OrderNumber);
                WriteCellValue(targetSheet, excelData.QuantityRange,       productRegisterWork.Quantity.ToString());
                WriteCellValue(targetSheet, excelData.SerialFirstRange,    productRegisterWork.SerialFirst);
                WriteCellValue(targetSheet, excelData.SerialLastRange,     productRegisterWork.SerialLast);
                WriteCellValue(targetSheet, excelData.RegDateRange,        formattedDate);
                WriteCellValue(targetSheet, excelData.RegTemperatureRange, temperature);
                WriteCellValue(targetSheet, excelData.RegHumidityRange,    humidity);
            }
        }

        // 指定シートのセルに値を書き込む
        private static void WriteCellValue(IXLWorksheet targetSheet, string? range, string value) {
            if (string.IsNullOrEmpty(range)) { return; }
            targetSheet.Cell(range).Value = value;
        }

        // 指定シート以外を非表示にする（Sheet1 は常に非表示）
        private static void HideSheets(XLWorkbook targetBook, List<string> sheetsToKeep) {
            var sheetsToHide = targetBook.Worksheets
                .Select(ws => ws.Name)
                .Where(name => !sheetsToKeep.Contains(name))
                .ToList();

            if (!sheetsToHide.Contains("Sheet1") && targetBook.Worksheets.Any(ws => ws.Name == "Sheet1")) {
                sheetsToHide.Add("Sheet1");
            }

            foreach (var sheetName in sheetsToHide) {
                targetBook.Worksheets.Worksheet(sheetName).Visibility = XLWorksheetVisibility.VeryHidden;
            }
        }

        // ワークブックをファイルに保存する
        private static void SaveWorkbook(XLWorkbook targetBook, string outputPath) {
            try {
                ExcelHelper.ForcePageBreakPreview(targetBook);
                targetBook.SaveAs(outputPath);
            } catch (IOException ex) {
                throw new IOException($"Excelファイルの保存に失敗しました: {outputPath}. 詳細: {ex.Message}", ex);
            }
        }

        // COM Interop を使用してExcelファイルを開く
        private static void PrintExcelFile(string filePath) {
            Microsoft.Office.Interop.Excel.Application? xlApp   = null;
            Microsoft.Office.Interop.Excel.Workbooks?   xlBooks = null;
            Microsoft.Office.Interop.Excel.Workbook?    xlBook  = null;

            try {
                xlApp   = new Microsoft.Office.Interop.Excel.Application { Visible = true };
                xlBooks = xlApp.Workbooks;
                xlBook  = xlBooks.Open(filePath, ReadOnly: true);
            } catch (Exception ex) {
                throw new Exception($"エラーが発生しました。詳細: {ex.Message}", ex);
            } finally {
                ExcelHelper.ReleaseComObject(xlBook);
                ExcelHelper.ReleaseComObject(xlBooks);
                ExcelHelper.ReleaseComObject(xlApp);
            }
        }
    }
}
