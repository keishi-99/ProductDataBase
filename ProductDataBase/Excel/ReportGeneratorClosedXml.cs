using ClosedXML.Excel;
using ProductDatabase.Data;
using ProductDatabase.Models;
using System.Data;

namespace ProductDatabase.Excel {
    // 成績書のExcel生成を担当するクラス
    internal static class ReportGeneratorClosedXml {

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
                        ? AppDomain.CurrentDomain.BaseDirectory
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

            ExcelHelper.ForcePageBreakPreview(reportWorkbook);
            reportWorkbook.SaveAs(savePath);
        }

        // 設定Excelワークブックを読み込む
        private static XLWorkbook LoadConfigWorkbook() {
            try {
                var configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config", "General", "Excel", "ConfigReport.xlsm");

                if (!File.Exists(configPath)) {
                    throw new FileNotFoundException($"設定ファイルが見つかりません: {configPath}");
                }

                // Excelが開かれていても読み取れるようにFileShare.ReadWrite指定
                using var fs = new FileStream(configPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                return new XLWorkbook(fs);
            } catch (Exception ex) {
                throw new Exception($"設定ファイルの読み込み中にエラーが発生しました: {ex.Message}", ex);
            }
        }

        // 設定ワークブックから製品に対応する設定を抽出する
        private static ReportConfigClosedXml GetReportConfig(XLWorkbook configWorkbook, string productModel) {
            var configSheet = configWorkbook.Worksheet("Sheet1")
                ?? throw new Exception("設定ファイルに 'Sheet1' が見つかりません。");

            var targetCell = configSheet
                .Column(1)
                .CellsUsed()
                .FirstOrDefault(c => c.GetValue<string>() == productModel)
                ?? throw new Exception($"Configに品目番号:[{productModel}]が見つかりません。");

            var resultRow = targetCell.Address.RowNumber;

            var sheetName = configSheet.Cell(resultRow, 5).GetString();
            if (string.IsNullOrEmpty(sheetName)) {
                throw new Exception($"設定ファイルのシート名が空です。");
            }

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

            var filePath = FindExcelFile(directoryPath, searchName);
            var fileName = Path.GetFileNameWithoutExtension(filePath);
            var fileExtension = Path.GetExtension(filePath).ToLowerInvariant();

            int startColumn = 12;
            int endColumn = 27;
            var targetRowCells = configSheet.Row(resultRow).Cells(startColumn, endColumn);
            bool hasValue = targetRowCells.Any(c => !string.IsNullOrWhiteSpace(c.GetValue<string>()));

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

        // 使用済み基板情報をDBから取得しSubstrateModel単位にグループ化する
        private static List<(string SubstrateModel, List<string> SubstrateNumbers)> GetUsedSubstrateData(ProductRegisterWork productRegisterWork) {
            var table = HistoryRepository.QueryUsedSubstrates(productRegisterWork.RowID);
            var usedSubstrate = new List<(string, List<string>)>();

            foreach (DataRow row in table.Rows) {
                var substrateModel = row["SubstrateModel"]?.ToString() ?? string.Empty;
                var substrateNumber = row["SubstrateNumber"]?.ToString() ?? string.Empty;

                var existingIndex = usedSubstrate.FindIndex(x => x.Item1 == substrateModel);
                if (existingIndex != -1) {
                    usedSubstrate[existingIndex].Item2.Add(substrateNumber);
                }
                else {
                    usedSubstrate.Add((substrateModel, new List<string> { substrateNumber }));
                }
            }

            return usedSubstrate;
        }

        // 基板情報をExcelシートに書き込む
        private static void UpdateSubstrateDetailsInExcel(XLWorkbook reportWorkbook, ReportConfigClosedXml config, List<(string SubstrateModel, List<string> SubstrateNumbers)> usedSubstrate) {
            var targetRow = config.ResultRow;
            var configSheet = config.Sheet;
            var targetRowCells = configSheet.Row(targetRow).Cells(12, 27);

            foreach (var (substrateModel, substrateNumbers) in usedSubstrate) {
                var searchAddressResult = targetRowCells
                    .FirstOrDefault(c => c.GetValue<string>() == substrateModel);

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
}
