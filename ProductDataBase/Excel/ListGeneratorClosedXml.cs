using ClosedXML.Excel;
using ProductDatabase.Data;
using ProductDatabase.Models;
using System.Data;
using ZXing;
using ZXing.QrCode;
using ZXing.QrCode.Internal;
using ZXing.Windows.Compatibility;

namespace ProductDatabase.Excel {
    // 製品登録リストのExcel生成を担当するクラス
    internal static class ListGeneratorClosedXml {

        // 製品セル範囲の設定データ
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

        // リストを生成するメインメソッド（例外は呼び出し元に伝播する）
        public static void GenerateList(ProductMaster productMaster, ProductRegisterWork productRegisterWork) {
            var configPath = Path.Combine(Environment.CurrentDirectory, "config", "General", "Excel", "ConfigList.xlsm");
            if (!File.Exists(configPath)) {
                throw new FileNotFoundException($"設定ファイルが見つかりません: {configPath}");
            }

            // Excelが開かれていても読み取れるようにFileShare.ReadWrite指定
            using var fs = new FileStream(configPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var workBook = new XLWorkbook(fs);

            var (targetSheetName, productName, resultRow, configSheet) = LoadExcelConfiguration(workBook, productMaster.ProductModel);

            var productCellRanges = GetProductCellRanges(configSheet, resultRow);
            var targetSheet = workBook.Worksheet(targetSheetName) ?? throw new Exception($"テンプレートシート:[{targetSheetName}]が見つかりません。");
            PopulateProductDetails(targetSheet, productMaster, productRegisterWork, productCellRanges, productName);

            var usedSubstrate = GetUsedSubstrateData(productRegisterWork);
            UpdateSubstrateDetailsInExcel(configSheet, targetSheet, resultRow, usedSubstrate);

            if (!string.IsNullOrEmpty(productCellRanges.QrCodeRange)) {
                GenerateAndEmbedQrCode(targetSheet, productMaster, productRegisterWork, productCellRanges.QrCodeRange);
            }

            var temporarilyPath = Path.Combine(Environment.CurrentDirectory, "config", "General", "Excel", "temporarilyList.xlsm");
            SaveAndPrintExcel(workBook, temporarilyPath, targetSheetName);
        }

        // Excel設定を読み込む
        private static (string sheetName, string productName, int resultRow, IXLWorksheet configSheet) LoadExcelConfiguration(XLWorkbook workBook, string productModel) {
            var configSheetName = "Sheet1";
            var configSheet = workBook.Worksheet(configSheetName)
                ?? throw new Exception($"設定シート:[{configSheetName}]が見つかりません。");

            var targetCell = configSheet
                .Column(1)
                .CellsUsed()
                .FirstOrDefault(c => c.GetValue<string>() == productModel)
                ?? throw new Exception($"Configに品目番号:[{productModel}]が見つかりません。");

            var resultRow = targetCell.Address.RowNumber;

            var sheetName = configSheet.Cell(resultRow, 2).GetString();
            if (string.IsNullOrEmpty(sheetName)) {
                throw new Exception($"設定シートの行 {resultRow}, 列 2 にシート名が設定されていません。");
            }

            var productName = configSheet.Cell(resultRow, 3).GetString();

            return (sheetName, productName, resultRow, configSheet);
        }

        // 製品情報に関連するExcelのセル範囲を取得する
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

        // 製品情報をExcelシートに書き込む
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

        // 使用済み基板情報をDBから取得しSubstrateModel単位にグループ化する
        private static List<(string SubstrateModel, List<string> SubstrateNumbers, List<int> Decreases)> GetUsedSubstrateData(ProductRegisterWork productRegisterWork) {
            var table = HistoryRepository.QueryUsedSubstrates(productRegisterWork.RowID);
            var usedSubstrate = new List<(string, List<string>, List<int>)>();

            foreach (DataRow row in table.Rows) {
                var substrateModel = row["SubstrateModel"]?.ToString() ?? string.Empty;
                var substrateNumber = row["SubstrateNumber"]?.ToString() ?? string.Empty;
                var decrease = row["Decrease"] is DBNull ? 0 : Convert.ToInt32(row["Decrease"]);

                var existingIndex = usedSubstrate.FindIndex(x => x.Item1 == substrateModel);
                if (existingIndex != -1) {
                    usedSubstrate[existingIndex].Item2.Add(substrateNumber);
                    usedSubstrate[existingIndex].Item3.Add(-decrease);
                }
                else {
                    usedSubstrate.Add((substrateModel, new List<string> { substrateNumber }, new List<int> { -decrease }));
                }
            }

            return usedSubstrate;
        }

        // 基板情報をExcelシートに書き込む
        private static void UpdateSubstrateDetailsInExcel(IXLWorksheet configSheet, IXLWorksheet targetSheet, int resultRow, List<(string SubstrateModel, List<string> SubstrateNumbers, List<int> Decreases)> usedSubstrate) {
            foreach (var (substrateModel, substrateNumbers, decreases) in usedSubstrate) {
                var searchAddressResult = configSheet
                    .Row(resultRow)
                    .CellsUsed()
                    .FirstOrDefault(c => c.GetValue<string>() == substrateModel);

                if (searchAddressResult == null) { continue; }

                var foundColumn = searchAddressResult.Address.ColumnNumber;

                var configCellValue = configSheet.Cell(resultRow, foundColumn + 1).GetString()
                    ?? throw new Exception($"Configシートの行 {resultRow}, 列 {foundColumn + 1} にセル範囲が設定されていません。");

                var targetCellValue = string.Join("   ", substrateNumbers.Select((substrateNumber, k) => $"{substrateNumber}({decreases[k]})"));
                targetSheet.Cell(configCellValue).Value = targetCellValue;
            }
        }

        // QRコードを生成しExcelシートに埋め込む
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

            var qrData = $"{productRegisterWork.OrderNumber};{productRegisterWork.ProductNumber};{productMaster.ProductModel};{productRegisterWork.Quantity};{productRegisterWork.SerialFirst};{productRegisterWork.SerialLast}";
            var qrBitmap = qr.Write(qrData);

            using var ms = new MemoryStream();
            qrBitmap.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
            ms.Position = 0;

            targetSheet.AddPicture(ms).MoveTo(targetSheet.Cell(qrCodeRange));
        }

        // Excelファイルを保存し、COM Interop でExcelを開く
        private static void SaveAndPrintExcel(XLWorkbook workBook, string temporarilyPath, string sheetName) {
            ExcelHelper.ForcePageBreakPreview(workBook);
            workBook.SaveAs(temporarilyPath);

            Microsoft.Office.Interop.Excel.Application? xlApp = null;
            Microsoft.Office.Interop.Excel.Workbooks? xlBooks = null;
            Microsoft.Office.Interop.Excel.Workbook? xlBook = null;
            Microsoft.Office.Interop.Excel.Sheets? xlSheets = null;
            Microsoft.Office.Interop.Excel.Worksheet? xlSheet = null;

            try {
                xlApp = new Microsoft.Office.Interop.Excel.Application { Visible = true };
                xlBooks = xlApp.Workbooks;
                xlBook = xlBooks.Open(temporarilyPath, ReadOnly: true);
                xlSheets = xlBook.Sheets;
                xlSheet = (Microsoft.Office.Interop.Excel.Worksheet?)xlSheets[sheetName];
                xlSheet?.Activate();
            } finally {
                ExcelHelper.ReleaseComObject(xlSheet);
                ExcelHelper.ReleaseComObject(xlSheets);
                ExcelHelper.ReleaseComObject(xlBook);
                ExcelHelper.ReleaseComObject(xlBooks);
                ExcelHelper.ReleaseComObject(xlApp);
            }
        }
    }
}
