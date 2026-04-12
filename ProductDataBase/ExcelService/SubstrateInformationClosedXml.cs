using ClosedXML.Excel;
using ProductDatabase.Models;

namespace ProductDatabase.ExcelService {
    // 基板情報ExcelファイルをCOM Interopで開くクラス
    internal static class SubstrateInformationClosedXml {

        // 基板設定ファイルを開く（例外は呼び出し元でハンドルする）
        public static void OpenSubstrateInformationClosedXml(SubstrateMaster substrateMaster) {
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
                return new XLWorkbook(fs);
            } catch (Exception ex) {
                throw new Exception($"設定ファイルの読み込み中にエラーが発生しました: {ex.Message}", ex);
            }
        }

        // 設定ワークブックから基板モデルに対応する設定を抽出する
        private static (string filePath, string sheetName) GetSubstrateConfig(XLWorkbook configWorkbook, string substrateModel) {
            var configSheet = configWorkbook.Worksheet("Sheet1") ?? throw new Exception("設定ファイルのシートに Sheet1 が見つかりません。");

            var targetCell = configSheet
                .Column(1)
                .CellsUsed()
                .FirstOrDefault(c => c.GetValue<string>() == substrateModel)
                ?? throw new Exception($"Configに品目番号:[{substrateModel}]が見つかりません。");

            var resultRow = targetCell.Address.RowNumber;

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

        // COM Interop を使用してExcelファイルを開く
        private static void OpenExcel(string filePath, string sheetName) {
            Microsoft.Office.Interop.Excel.Application? xlApp    = null;
            Microsoft.Office.Interop.Excel.Workbooks?   xlBooks  = null;
            Microsoft.Office.Interop.Excel.Workbook?    xlBook   = null;
            Microsoft.Office.Interop.Excel.Sheets?      xlSheets = null;
            Microsoft.Office.Interop.Excel.Worksheet?   xlSheet  = null;

            try {
                xlApp   = new Microsoft.Office.Interop.Excel.Application { Visible = true };
                xlBooks = xlApp.Workbooks;
                xlBook  = xlBooks.Open(filePath, ReadOnly: true);
                xlSheets = xlBook.Sheets;
                xlSheet  = (Microsoft.Office.Interop.Excel.Worksheet?)(string.IsNullOrEmpty(sheetName) ? xlSheets[1] : xlSheets[sheetName]);
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
