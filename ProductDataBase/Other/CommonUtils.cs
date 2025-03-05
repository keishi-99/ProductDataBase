using OfficeOpenXml;
using System.Data;
using System.Data.SQLite;
using System.Drawing.Imaging;
using System.Text;
using ZXing;
using ZXing.QrCode;
using ZXing.QrCode.Internal;
using ZXing.Rendering;
using static ProductDatabase.MainWindow;

namespace ProductDatabase.Other {
    internal partial class CommonUtils {
        public static string s_networkPath = string.Empty; // ClonePathを保持する静的変数

        // ログ作成
        public static class Logger {
            private static readonly string s_logDirectory = Path.Combine(Environment.CurrentDirectory, "db", "logs"); // ログを保存するディレクトリ
            private static readonly object s_lockObject = new();

            /// <summary>
            /// 作業ログを追記します。
            /// </summary>
            /// <param name="message">記録する作業内容</param>
            public static void AppendLog(string message) {
                try {
                    lock (s_lockObject) {
                        // ディレクトリが存在しない場合は作成
                        if (!Directory.Exists(s_logDirectory)) {
                            Directory.CreateDirectory(s_logDirectory);
                        }

                        //// 年と月を含むログファイル名を生成
                        var logFileName = $"log_{DateTime.Now:yyyyMM}.csv"; // CSVファイルとして保存
                        var logFilePath = Path.Combine(s_logDirectory, logFileName);

                        // CSV形式でエスケープ処理
                        var escapedMessage = EscapeCsvField(message);
                        var logEntry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss},{escapedMessage}";
                        // ログ内容をファイルの末尾に追記
                        File.AppendAllText(logFilePath, logEntry + Environment.NewLine, Encoding.UTF8); // UTF-8で保存

                        if (!string.IsNullOrEmpty(s_networkPath)) {
                            var cloneFilePath = Path.Combine(s_networkPath, "db", "logs", logFileName);
                            if (cloneFilePath != logFilePath) {
                                File.Copy(logFilePath, cloneFilePath, true);
                            }
                        }
                    }
                } catch (Exception ex) {
                    MessageBox.Show($"ログの書き込み中にエラーが発生しました: {ex.Message}", $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        private static readonly char[] SPECIAL_CHARACTERS = { ',', '"', '\r', '\n' };
        private static string EscapeCsvField(string field) {
            if (string.IsNullOrEmpty(field)) {
                return "";
            }

            if (field.IndexOfAny(SPECIAL_CHARACTERS) == -1) {
                return field;
            }

            var sb = new StringBuilder();
            foreach (var c in field) {
                if (c == '"') {
                    sb.Append("\"\""); // ダブルクォーテーションをエスケープ
                }
                else {
                    sb.Append(c);
                }
            }

            return $"\"{sb.ToString()}\""; // ダブルクォーテーションで囲む
        }
        // バックアップ作成
        public static class BackupManager {
            private static readonly string s_backupDirectory = Path.Combine(Environment.CurrentDirectory, "db", "backup"); // バックアップを保存するディレクトリ
            private static readonly string s_originalFilePath = Path.Combine(Environment.CurrentDirectory, "db", "registration.db"); // 元ファイルパス
            //private static readonly string s_originalFilePath = Path.Combine(s_networkPath, "db", "registration.db"); // 元ファイルパス
            private static readonly int s_maxBackupFiles = 10; // 最大バックアップファイル数
            private static readonly object s_lockObject = new();

            /// <summary>
            /// バックアップを作成します。
            /// </summary>
            public static void CreateBackup() {
                try {
                    lock (s_lockObject) {
                        // バックアップ用ディレクトリが存在しない場合は作成
                        if (!Directory.Exists(s_backupDirectory)) {
                            Directory.CreateDirectory(s_backupDirectory);
                        }

                        // 日付と時間をファイル名に付加
                        var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                        var backupFileName = $"registration_{timestamp}.db";
                        var backupFilePath = Path.Combine(s_backupDirectory, backupFileName);

                        // 元ファイルをバックアップにコピー
                        File.Copy(s_originalFilePath, backupFilePath, true);
                        var networkFilePath = Path.Combine(s_networkPath, "db", "registration.db");
                        if (Environment.CurrentDirectory != s_networkPath) {
                            File.Copy(s_originalFilePath, networkFilePath, true);
                        }

                        // バックアップファイルを管理
                        ManageBackupFiles();
                    }
                } catch (Exception ex) {
                    MessageBox.Show($"バックアップの作成中にエラーが発生しました: {ex.Message}", $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            /// <summary>
            /// 古いバックアップファイルを削除します。
            /// </summary>
            private static void ManageBackupFiles() {
                try {
                    var backupFiles = Directory.GetFiles(s_backupDirectory, "registration_*.db")
                                               .OrderBy(File.GetCreationTime) // 作成日時順に並べる
                                               .ToList();

                    while (backupFiles.Count > s_maxBackupFiles) {
                        var oldestFile = backupFiles.First();
                        File.Delete(oldestFile);
                        backupFiles.RemoveAt(0);
                    }
                } catch (Exception ex) {
                    MessageBox.Show($"バックアップの作成中にエラーが発生しました: {ex.Message}", $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        // 成績書作成
        public static void GenerateReport(ProductInformation productInfo) {
            try {
                var configPath = Path.Combine(Environment.CurrentDirectory, "config", "General", "Excel", "ConfigReport.xlsx");
                using FileStream fileStreamConfig = new(configPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using var workBookConfig = new ExcelPackage(fileStreamConfig);
                //既存ワークシートを取得（workBookはExcelWorkbookクラスオブジェト）
                var sheet = workBookConfig.Workbook.Worksheets;
                var targetSheetName = "Sheet1";
                var workSheetMain = sheet[targetSheetName];

                // セル検索
                var searchAddressResult = workSheetMain.Cells.FirstOrDefault(x => x.Start.Column == 1 && x.Value?.ToString() == productInfo.ProductModel) ?? throw new Exception($"Configに品目番号:[{productInfo.ProductModel}]が見つかりません。");
                var searchAddressResultRow = searchAddressResult.Start.Row;

                // ワークシートのセルから値を取得
                var directoryPath = workSheetMain.Cells[searchAddressResultRow, 3].Value?.ToString()?.Trim('"');
                if (string.IsNullOrWhiteSpace(directoryPath)) { throw new Exception("Configのファイルパスが無効です。"); }
                if (!Directory.Exists(directoryPath)) { throw new FileNotFoundException($"指定されたフォルダが存在しません: {directoryPath}"); }

                var searchName = workSheetMain.Cells[searchAddressResultRow, 4].Value.ToString()?.Trim('"');
                if (string.IsNullOrWhiteSpace(searchName)) { throw new Exception("Configのファイル名が無効です。"); }

                var filePaths = Directory.GetFiles(directoryPath, $"*{searchName}*", SearchOption.TopDirectoryOnly);
                var filePath = filePaths[0];
                var fileName = Path.GetFileNameWithoutExtension(filePath);
                var fileExtension = Path.GetExtension(filePath).ToLower(); // 開いたファイルの拡張子取得

                var sheetName = !string.IsNullOrEmpty(workSheetMain.Cells[searchAddressResultRow, 5].Value.ToString())
                    ? workSheetMain.Cells[searchAddressResultRow, 5].Value.ToString()
                    : throw new Exception("シート名がありません。");
                var productNumberRange = workSheetMain.Cells[searchAddressResultRow, 6].Value?.ToString() ?? workSheetMain.Cells[2, 6].Value.ToString();
                var orderNumberRange = workSheetMain.Cells[searchAddressResultRow, 7].Value?.ToString() ?? workSheetMain.Cells[2, 7].Value.ToString();
                var quantityRange = workSheetMain.Cells[searchAddressResultRow, 8].Value?.ToString() ?? workSheetMain.Cells[2, 8].Value.ToString();
                var serialFirstRange = workSheetMain.Cells[searchAddressResultRow, 9].Value?.ToString() ?? workSheetMain.Cells[2, 9].Value.ToString();
                var serialLastRange = workSheetMain.Cells[searchAddressResultRow, 10].Value?.ToString() ?? workSheetMain.Cells[2, 10].Value.ToString();
                var productModelRange = workSheetMain.Cells[searchAddressResultRow, 11].Value?.ToString();
                var saveDirectory = workSheetMain.Cells[searchAddressResultRow, 12].Value?.ToString();
                using FileStream fileStreamReport = new(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using var workBookReport = new ExcelPackage(fileStreamReport);
                var reportSheet = workBookReport.Workbook.Worksheets;

                // セルに値を挿入
                var workSheetTemp = reportSheet[sheetName];
                if (!string.IsNullOrEmpty(productInfo.ProductNumber)) {
                    var productNumbers = productInfo.ProductNumber.Split("-");
                    workSheetTemp.Cells[productNumberRange].Value = productNumbers[0];
                }
                if (!string.IsNullOrEmpty(orderNumberRange) && !string.IsNullOrEmpty(productInfo.OrderNumber)) {
                    workSheetTemp.Cells[orderNumberRange].Value = productInfo.OrderNumber;
                }
                if (!string.IsNullOrEmpty(quantityRange) && productInfo.Quantity != 0) {
                    workSheetTemp.Cells[quantityRange].Value = productInfo.Quantity;
                }
                if (!string.IsNullOrEmpty(serialFirstRange) && !string.IsNullOrEmpty(productInfo.SerialFirst)) {
                    workSheetTemp.Cells[serialFirstRange].Value = productInfo.SerialFirst;
                }
                if (!string.IsNullOrEmpty(serialLastRange) && !string.IsNullOrEmpty(productInfo.SerialLast)) {
                    workSheetTemp.Cells[serialLastRange].Value = productInfo.SerialLast;
                }
                if (!string.IsNullOrEmpty(productModelRange) && !string.IsNullOrEmpty(productInfo.ProductModel)) {
                    workSheetTemp.Cells[productModelRange].Value = productInfo.ProductModel;
                }

                // ダイアログで保存先を選択
                using SaveFileDialog saveFileDialog = new() {
                    Filter = $"Excel Files (*{fileExtension})|*{fileExtension}|All Files (*.*)|*.*",
                    FileName = $"{fileName} のコピー{productInfo.ProductNumber}{fileExtension}",
                    Title = "保存先を選択してください",
                    InitialDirectory = saveDirectory
                };
                if (saveFileDialog.ShowDialog() == DialogResult.OK) {
                    var outputPath = saveFileDialog.FileName;
                    var fileInfo = new FileInfo(outputPath);
                    workBookReport.SaveAs(fileInfo);
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // リスト生成
        public static void GenerateList(ProductInformation productInfo) {
            try {
                var configPath = Path.Combine(Environment.CurrentDirectory, "config", "General", "Excel", "ConfigList.xlsx");
                using FileStream fileStream = new(configPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using var workBook = new ExcelPackage(fileStream);
                //既存ワークシートを取得（workBookはExcelWorkbookクラスオブジェクト）
                var sheet = workBook.Workbook.Worksheets;
                var targetSheetName = "Sheet1";
                var workSheetMain = sheet[targetSheetName];

                // セル検索
                var searchAddressResult = workSheetMain.Cells.FirstOrDefault(x => x.Start.Column == 1 && x.Value?.ToString() == productInfo.ProductModel) ?? throw new Exception($"Configに品目番号:[{productInfo.ProductModel}]が見つかりません。");
                var resultRow = searchAddressResult.Start.Row;

                // ワークシートのセルから値を取得
                var sheetName = workSheetMain.Cells[resultRow, 2].Value?.ToString();
                var productName = workSheetMain.Cells[resultRow, 3].Value?.ToString();
                var productNameRange = workSheetMain.Cells[resultRow, 4].Value?.ToString();
                var productNumberRange = workSheetMain.Cells[resultRow, 5].Value?.ToString();
                var orderNumberRange = workSheetMain.Cells[resultRow, 6].Value?.ToString();
                var regDateRange = workSheetMain.Cells[resultRow, 7].Value?.ToString();
                var productModel = workSheetMain.Cells[resultRow, 8].Value?.ToString();
                var productModelRange = workSheetMain.Cells[resultRow, 9].Value?.ToString();
                var quantityRange = workSheetMain.Cells[resultRow, 10].Value?.ToString();
                var serialFirstRange = workSheetMain.Cells[resultRow, 11].Value?.ToString();
                var serialLastRange = workSheetMain.Cells[resultRow, 12].Value?.ToString();
                var commentRange = workSheetMain.Cells[resultRow, 13].Value?.ToString();
                var qrCodeRange = workSheetMain.Cells[resultRow, 14].Value?.ToString();

                var workSheetTemp = sheet[sheetName];
                if (!string.IsNullOrEmpty(productNameRange)) { workSheetTemp.Cells[productNameRange].Value = productName; }
                if (!string.IsNullOrEmpty(productNumberRange)) { workSheetTemp.Cells[productNumberRange].Value = productInfo.ProductNumber; }
                if (!string.IsNullOrEmpty(orderNumberRange)) { workSheetTemp.Cells[orderNumberRange].Value = productInfo.OrderNumber; }
                if (!string.IsNullOrEmpty(regDateRange)) { workSheetTemp.Cells[regDateRange].Value = productInfo.RegDate; }
                if (!string.IsNullOrEmpty(productModelRange)) { workSheetTemp.Cells[productModelRange].Value = productInfo.ProductModel; }
                if (!string.IsNullOrEmpty(quantityRange)) { workSheetTemp.Cells[quantityRange].Value = productInfo.Quantity; }
                if (!string.IsNullOrEmpty(serialFirstRange)) { workSheetTemp.Cells[serialFirstRange].Value = productInfo.SerialFirst; }
                if (!string.IsNullOrEmpty(serialLastRange)) { workSheetTemp.Cells[serialLastRange].Value = productInfo.SerialLast; }
                if (!string.IsNullOrEmpty(commentRange)) { workSheetTemp.Cells[commentRange].Value = productInfo.Comment; }

                List<(string, List<string>, List<int>)> usedSubstrate = [];

                using SQLiteConnection con = new(GetConnectionRegistration());
                con.Open();
                using var cmd = con.CreateCommand();

                cmd.CommandText = $"""
                                SELECT
                                	SubstrateModel,
                                	SubstrateNumber,
                                	Decrease
                                FROM
                                	PA_Substrate
                                WHERE
                                	UseID = @ID
                                ORDER BY
                                	SubstrateModel ASC
                                """;
                cmd.Parameters.Add("@ID", DbType.String).Value = productInfo.ProductID;
                using var dr = cmd.ExecuteReader();

                while (dr.Read()) {
                    var substrateModel = dr.GetString(0);
                    var substrateNumber = dr.GetString(1);
                    var decrease = -1 * dr.GetInt32(2);

                    // 既存の substrateModel を検索
                    var existingSubstrate = usedSubstrate.FirstOrDefault(x => x.Item1 == substrateModel);

                    if (existingSubstrate != default) {
                        // 既存の substrateModel が見つかった場合、リストに追加
                        existingSubstrate.Item2.Add(substrateNumber);
                        existingSubstrate.Item3.Add(decrease);
                    }
                    else {
                        // 既存の substrateModel が見つからなかった場合、新しいエントリを追加
                        List<string> substrateNumbers = [substrateNumber];
                        List<int> decreases = [decrease];
                        (string, List<string>, List<int>) substrateData = (substrateModel, substrateNumbers, decreases);
                        usedSubstrate.Add(substrateData);
                    }
                }
                for (var i = 0; i <= usedSubstrate.Count - 1; i++) {
                    var targetRow = resultRow; // 検索対象の行番号
                    var searchValue = $"{usedSubstrate[i].Item1}";
                    var foundColumn = 0;

                    var searchAddressResult2 = workSheetMain.Cells
                        .Where(x => x.Start.Row == targetRow && x.Value?.ToString() == searchValue)
                        .FirstOrDefault();

                    if (searchAddressResult2 != null) {
                        // セルが見つかった場合の処理
                        foundColumn = searchAddressResult2.Start.Column;
                    }
                    if (foundColumn == 0) {
                        throw new Exception($"{usedSubstrate[i]}が見つかりません。");
                    }
                    var mainCellValue = workSheetMain.Cells[resultRow, foundColumn + 1].Value.ToString();
                    var tempCellValue = string.Join("    ", usedSubstrate[i].Item2.Select((subStrateNumber, k) => $"{subStrateNumber}({usedSubstrate[i].Item3[k]})"));

                    workSheetTemp.Cells[mainCellValue].Value = tempCellValue;
                }

                // QRコード
                if (!string.IsNullOrEmpty(qrCodeRange)) {
                    BarcodeWriter<PixelData> qr = new() {
                        Format = BarcodeFormat.QR_CODE,
                        Options = new QrCodeEncodingOptions {
                            ErrorCorrection = ErrorCorrectionLevel.L,
                            CharacterSet = "Shift_JIS",
                            Width = 100,
                            Height = 100,
                        },
                        Renderer = new PixelDataRenderer {
                            Foreground = new(Color.Gray.ToArgb()),
                            Background = new(Color.White.ToArgb()),
                        },
                    };

                    var pixelData = qr.Write($"{productInfo.OrderNumber};{productInfo.ProductNumber};{productModel};{productInfo.Quantity};{productInfo.SerialFirst};{productInfo.SerialLast}");

                    // PixelData を Bitmap に変換
                    using var bitmap = new Bitmap(pixelData.Width, pixelData.Height, PixelFormat.Format32bppArgb);
                    var bmpData = bitmap.LockBits(new System.Drawing.Rectangle(0, 0, bitmap.Width, bitmap.Height), ImageLockMode.WriteOnly, PixelFormat.Format32bppArgb);
                    System.Runtime.InteropServices.Marshal.Copy(pixelData.Pixels, 0, bmpData.Scan0, pixelData.Pixels.Length);
                    bitmap.UnlockBits(bmpData);
                    using MemoryStream stream = new();
                    bitmap.Save(stream, ImageFormat.Bmp);

                    var image = workSheetTemp.Drawings.AddPicture("QR", bitmap);
                    image.SetPosition(workSheetTemp.Cells[qrCodeRange].Start.Row - 1, 0, workSheetTemp.Cells[qrCodeRange].Start.Column - 1, 0);
                }

                //引数に保存先パスを指定
                var temporarilyPath = Path.Combine(Environment.CurrentDirectory, "config", "General", "Excel", "temporarilyList.xlsx");
                var fileInfo = new FileInfo(temporarilyPath);
                workBook.SaveAs(fileInfo);

                // 印刷
                Microsoft.Office.Interop.Excel.Application xlApp = new() {
                    Visible = true // Excelウィンドウを表示します。
                };

                // ワークブック開く
                var xlBooks = xlApp.Workbooks;
                var xlBook = xlBooks.Open(temporarilyPath, ReadOnly: true);

                // ワークシート選択
                var xlSheets = xlBook.Sheets;
                Microsoft.Office.Interop.Excel.Worksheet xlSheet = xlSheets[sheetName];

                // ワークシート表示
                xlSheet.Activate();

                // ワークブックを閉じてExcelを終了
                //xlBook.Close(false);
                //xlApp.Quit();

                _ = System.Runtime.InteropServices.Marshal.ReleaseComObject(xlSheet);
                _ = System.Runtime.InteropServices.Marshal.ReleaseComObject(xlSheets);
                _ = System.Runtime.InteropServices.Marshal.ReleaseComObject(xlBook);
                _ = System.Runtime.InteropServices.Marshal.ReleaseComObject(xlBooks);
                _ = System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);

            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // チェックシート生成
        public static void GenerateCheckSheet(ProductInformation productInfo) {
            try {
                var dialog = new InputDialog1();
                var result = dialog.ShowDialog();

                var temperature = string.Empty;
                var humidity = string.Empty;

                if (result == DialogResult.OK) {
                    temperature = dialog.Temperature;
                    humidity = dialog.Humidity;
                }
                else {
                    return;
                }

                var configPath = Path.Combine(Environment.CurrentDirectory, "config", "General", "Excel", "ConfigCheckSheet.xlsx");
                using FileStream fileStream = new(configPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using var workBook = new ExcelPackage(fileStream);
                //既存ワークシートを取得（workBookはExcelWorkbookクラスオブジェクト）
                var sheet = workBook.Workbook.Worksheets;
                var targetSheetName = "Sheet1";
                var workSheetMain = sheet[targetSheetName];

                // セル検索
                var searchAddressResult = workSheetMain.Cells.FirstOrDefault(x => x.Start.Column == 1 && x.Value?.ToString() == productInfo.ProductModel) ?? throw new Exception($"Configに品目番号:[{productInfo.ProductModel}]が見つかりません。");
                var resultRow = searchAddressResult.Start.Row;

                // ワークシートのセルから値を取得
                var orderNumberRange = workSheetMain.Cells[resultRow, 7].Value?.ToString();
                var quantityRange = workSheetMain.Cells[resultRow, 8].Value?.ToString();
                var serialFirstRange = workSheetMain.Cells[resultRow, 9].Value?.ToString();
                var serialLastRange = workSheetMain.Cells[resultRow, 10].Value?.ToString();
                var regDateRange = workSheetMain.Cells[resultRow, 11].Value?.ToString();
                var dateFormat = workSheetMain.Cells[resultRow, 12].Value?.ToString();
                //var regDateMonthRange = workSheetMain.Cells[resultRow, ].Value?.ToString();
                //var regDateDayRange = workSheetMain.Cells[resultRow, ].Value?.ToString();
                var regTemperatureRange = workSheetMain.Cells[resultRow, 13].Value?.ToString();
                var regHumidityRange = workSheetMain.Cells[resultRow, 14].Value?.ToString();

                const int StartColumn = 15;
                var sheetNames = Enumerable.Range(StartColumn, 20) // 無限の範囲
                    .Select(column => workSheetMain.Cells[resultRow, column].Value?.ToString())
                    .TakeWhile(sheetName => !string.IsNullOrWhiteSpace(sheetName)) // 空白でない間
                    .ToList();

                if (sheetNames.Count == 0) { throw new Exception("対象シートがありません。"); }

                var formattedDate = string.Empty;
                if (DateTime.TryParse(productInfo.RegDate, out var date)) {
                    switch (dateFormat) {
                        case "1":
                            formattedDate = date.ToString("yyyy年MM月dd日");
                            break;
                        case "2":
                            formattedDate = date.ToString("yyyy-MM-dd");
                            break;
                    }
                }

                foreach (var sheetName in sheetNames) {
                    var workSheetTemp = sheet[sheetName] ?? throw new Exception($"シート[{sheetName}]が見つかりません。");
                    if (!string.IsNullOrEmpty(orderNumberRange)) { workSheetTemp.Cells[orderNumberRange].Value = productInfo.OrderNumber; }
                    if (!string.IsNullOrEmpty(quantityRange)) { workSheetTemp.Cells[quantityRange].Value = productInfo.Quantity; }
                    if (!string.IsNullOrEmpty(serialFirstRange)) { workSheetTemp.Cells[serialFirstRange].Value = productInfo.SerialFirst; }
                    if (!string.IsNullOrEmpty(serialLastRange)) { workSheetTemp.Cells[serialLastRange].Value = productInfo.SerialLast; }
                    if (!string.IsNullOrEmpty(regDateRange)) { workSheetTemp.Cells[regDateRange].Value = formattedDate; }
                    //if (!string.IsNullOrEmpty(regDateMonthRange)) { workSheetTemp.Cells[regDateMonthRange].Value = date.Month; }
                    //if (!string.IsNullOrEmpty(regDateDayRange)) { workSheetTemp.Cells[regDateDayRange].Value = date.Day; }
                    if (!string.IsNullOrEmpty(regTemperatureRange)) { workSheetTemp.Cells[regTemperatureRange].Value = temperature; }
                    if (!string.IsNullOrEmpty(regHumidityRange)) { workSheetTemp.Cells[regHumidityRange].Value = humidity; }
                }

                // 不要なシートを非表示にする
                var allSheetName = sheet
                    .Select(sh => sh.Name.ToString())
                    .ToList();

                var hiddenSheetNames = allSheetName.Except(sheetNames.Where(name => name != null).Cast<string>()).ToList();
                foreach (var sheetName in hiddenSheetNames) {
                    sheet[sheetName].Hidden = eWorkSheetHidden.VeryHidden;
                }
                sheet["Sheet1"].Hidden = eWorkSheetHidden.VeryHidden;

                //引数に保存先パスを指定
                var temporarilyPath = Path.Combine(Environment.CurrentDirectory, "config", "General", "Excel", "temporarilyCheckSheet.xlsx");
                var fileInfo = new FileInfo(temporarilyPath);
                workBook.SaveAs(fileInfo);

                // 印刷
                Microsoft.Office.Interop.Excel.Application xlApp = new() {
                    Visible = true // Excelウィンドウを表示します。
                };

                // ワークブック開く
                var xlBooks = xlApp.Workbooks;
                var xlBook = xlBooks.Open(temporarilyPath, ReadOnly: true);

                //// ワークシート選択
                //var xlSheets = xlBook.Sheets;
                //Microsoft.Office.Interop.Excel.Worksheet xlSheet = xlSheets[0];

                //// ワークシート表示
                //xlSheet.Activate();

                //// ワークブックを閉じてExcelを終了
                //xlBook.Close(false);
                //xlApp.Quit();

                //_ = System.Runtime.InteropServices.Marshal.ReleaseComObject(xlSheet);
                //_ = System.Runtime.InteropServices.Marshal.ReleaseComObject(xlSheets);
                _ = System.Runtime.InteropServices.Marshal.ReleaseComObject(xlBook);
                _ = System.Runtime.InteropServices.Marshal.ReleaseComObject(xlBooks);
                _ = System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);

            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
