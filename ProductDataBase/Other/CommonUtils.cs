using Microsoft.Data.Sqlite;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using OfficeOpenXml;
using System.Data;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using ZXing;
using ZXing.QrCode;
using ZXing.QrCode.Internal;
using ZXing.Windows.Compatibility;
using static ProductDatabase.MainWindow;

namespace ProductDatabase.Other {

    internal static partial class NativeMethods {
        [LibraryImport("user32.dll", SetLastError = true)]
        public static partial void keybd_event(byte bVk, byte bScan, int dwFlags, int dwExtraInfo);
    }

    internal partial class CommonUtils {
        public static string s_networkPath = string.Empty; // ClonePathを保持する静的変数
        /// <summary>
        /// ファイルをコピーします。
        /// </summary>
        private static void CopyWithRetry(string sourceFileName, string destFileName, bool overwrite, int retryCount = 5, int delayMilliseconds = 2000) {
            for (var attempt = 1; attempt <= retryCount; attempt++) {
                try {
                    if (File.Exists(sourceFileName)) {
                        File.Copy(sourceFileName, destFileName, overwrite);
                    }
                    return;
                } catch (IOException) {
                    if (attempt == retryCount) {
                        throw new Exception("バックアップファイルがコピーできません。"); // 最後の試行で失敗したら例外を投げる
                    }
                    Thread.Sleep(delayMilliseconds); // 一定時間待機してリトライ
                }
            }
        }
        /// <summary>
        /// ファイルを削除します。
        /// </summary>
        private static void DeleteWithRetry(string filePath, int retryCount = 5, int delayMilliseconds = 2000) {
            for (var attempt = 1; attempt <= retryCount; attempt++) {
                try {
                    if (File.Exists(filePath)) {
                        File.Delete(filePath);
                    }
                    return;
                } catch (IOException) {
                    if (attempt == retryCount) {
                        throw new Exception("バックアップファイルが削除できません。"); // 最後の試行で失敗したら例外を投げる
                    }
                    Thread.Sleep(delayMilliseconds); // 一定時間待機してリトライ
                }
            }
        }
        // ログ作成
        public static class Logger {
            private static readonly string s_logDirectory = Path.Combine(Environment.CurrentDirectory, "db", "logs"); // ログを保存するディレクトリ
            private static readonly object s_lockObject = new();

            /// <summary>
            /// 作業ログを追記します。
            /// </summary>
            /// <param name="message">記録する作業内容</param>
            public static void AppendLog(string[] message) {
                try {
                    lock (s_lockObject) {
                        // ディレクトリが存在しない場合は作成
                        if (!Directory.Exists(s_logDirectory)) {
                            Directory.CreateDirectory(s_logDirectory);
                        }

                        //// 年と月を含むログファイル名を生成
                        var logFileName = $"log_{DateTime.Now:yyyyMM}.csv";
                        var logFilePath = Path.Combine(s_logDirectory, logFileName);

                        // CSV形式でログ内容をファイルの末尾に追記
                        var logEntry = $"\"{DateTime.Now:yyyy-MM-dd HH:mm:ss}\",{string.Join(",", message.Select(m => $"\"{m.Replace("\"", "\"\"").Replace("\r\n", " ").Replace("\r", " ").Replace("\n", " ")}\""))}";
                        // ログ内容をファイルの末尾に追記
                        File.AppendAllText(logFilePath, logEntry + Environment.NewLine);

                        if (!string.IsNullOrEmpty(s_networkPath)) {
                            var cloneFilePath = Path.Combine(s_networkPath, "db", "logs", logFileName);
                            if (cloneFilePath != logFilePath) {
                                CopyWithRetry(logFilePath, cloneFilePath, true);
                            }
                        }
                    }
                } catch (Exception ex) {
                    MessageBox.Show($"ログの書き込み中にエラーが発生しました: {ex.Message}", $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        // バックアップ作成
        public static class BackupManager {
            private static readonly string s_backupDirectory = Path.Combine(Environment.CurrentDirectory, "db", "backup"); // バックアップを保存するディレクトリ
            private static readonly string s_originalFilePath = Path.Combine(Environment.CurrentDirectory, "db", "registration.db"); // 元ファイルパス
            private static readonly int s_maxBackupFiles = 20; // 最大バックアップファイル数
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
                        CopyWithRetry(s_originalFilePath, backupFilePath, true);
                        // バックアップファイルを管理
                        ManageBackupFiles();

                        // ネットワークにバックアップ
                        var networkFilePath = Path.Combine(s_networkPath, "db", "registration.db");
                        if (Environment.CurrentDirectory != s_networkPath) {
                            CopyWithRetry(s_originalFilePath, networkFilePath, true);
                        }

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
                        DeleteWithRetry(oldestFile);
                        backupFiles.RemoveAt(0);
                    }
                } catch (Exception ex) {
                    MessageBox.Show($"バックアップの作成中にエラーが発生しました: {ex.Message}", $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        // CapsLockがオンになっていたらCapsLockを解除する
        public static partial class Keyboard {
            private const byte VK_CAPITAL = 0x14; // CapsLock の仮想キーコード
            private const int KEYEVENTF_EXTENDEDKEY = 0x1;
            private const int KEYEVENTF_KEYUP = 0x2;

            // CapsLock の状態を切り替える
            public static void CapsDisable() {
                if (Control.IsKeyLocked(Keys.CapsLock)) {
                    NativeMethods.keybd_event(VK_CAPITAL, 0, KEYEVENTF_EXTENDEDKEY, 0);
                    NativeMethods.keybd_event(VK_CAPITAL, 0, KEYEVENTF_KEYUP, 0);
                }
            }
        }

        // 成績書生成データを保持するクラス
        public class ReportConfig {
            public required string DirectoryPath { get; set; }
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
        // 成績書生成
        public static class ReportGeneratorNPOI {
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
                var configPath = Path.Combine(Environment.CurrentDirectory, "config", "General", "Excel", "ConfigReport.xlsx");
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
            private static ReportConfig GetReportConfig(XSSFWorkbook configWorkbook, string productModel) {
                var workSheetMain = configWorkbook.GetSheet("Sheet1") ?? throw new Exception("設定ファイルに 'Sheet1' が見つかりません。");

                var targetColumnIndex = 0;
                var searchText = productModel;
                var resultRowIndex = -1;

                // 対象列だけ走査
                for (int rowIndex = workSheetMain.FirstRowNum; rowIndex <= workSheetMain.LastRowNum; rowIndex++) {
                    IRow row = workSheetMain.GetRow(rowIndex);
                    if (row == null) continue;

                    ICell cell = row.GetCell(targetColumnIndex);
                    if (cell == null) continue;

                    string cellValue = GetCellValue(cell) ?? string.Empty;

                    if (cellValue.Equals(searchText, StringComparison.OrdinalIgnoreCase)) {
                        resultRowIndex = rowIndex;
                    }
                }

                if (resultRowIndex == -1) { throw new Exception($"Configに品目番号:[{searchText}]が見つかりません。"); }

                // ワークシートのセルから値を取得し、ReportConfigオブジェクトに格納
                IRow resultRow = workSheetMain.GetRow(resultRowIndex);
                var directoryPath = GetCellValue(resultRow.GetCell(2))?.Trim('"') ?? string.Empty;
                if (string.IsNullOrWhiteSpace(directoryPath)) { throw new Exception("Configのファイルパスが無効です。"); }
                if (!Directory.Exists(directoryPath)) { throw new FileNotFoundException($"指定されたフォルダが存在しません: {directoryPath}"); }

                var searchName = GetCellValue(resultRow.GetCell(3))?.Trim('"') ?? string.Empty;
                if (string.IsNullOrWhiteSpace(searchName)) { throw new Exception("Configのファイル名が無効です。"); }

                var filePaths = Directory.GetFiles(directoryPath, $"*{searchName}*", SearchOption.AllDirectories);
                var filePath = filePaths[0];
                var fileName = Path.GetFileNameWithoutExtension(filePath);
                var fileExtension = Path.GetExtension(filePath).ToLower();

                var sheetName = GetCellValue(resultRow.GetCell(4)) ?? string.Empty;
                return string.IsNullOrWhiteSpace(sheetName)
                    ? throw new Exception("シート名がありません。")
                    : new ReportConfig {
                        DirectoryPath = directoryPath,
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

                static string? GetCellValue(ICell cell) {
                    if (cell == null) return string.Empty;
                    return cell.CellType switch {
                        CellType.String => cell.StringCellValue,
                        CellType.Numeric => cell.NumericCellValue.ToString(),
                        CellType.Boolean => cell.BooleanCellValue.ToString(),
                        CellType.Formula => cell.ToString(),
                        _ => cell.ToString()
                    };
                }
            }

            // レポートテンプレートExcelワークブックを読み込む
            private static XSSFWorkbook LoadReportTemplate(string directoryPath, string searchFileName) {
                var filePaths = Directory.GetFiles(directoryPath, $"{searchFileName}*", SearchOption.AllDirectories);
                if (filePaths.Length == 0) {
                    throw new FileNotFoundException($"指定されたファイル名 '{searchFileName}' のファイルが '{directoryPath}' に見つかりません。");
                }

                string filePath;

                if (filePaths.Length == 1) {
                    // ファイルが1つだけ見つかった場合は、それを自動的に選択
                    filePath = filePaths[0];
                }
                else {
                    MessageBox.Show("複数のファイルが見つかりました。1つ選択してください。");
                    // 複数のファイルが見つかった場合は、OpenFileDialog を使用してユーザーに選択させる
                    using var openFileDialog = new OpenFileDialog();
                    openFileDialog.InitialDirectory = directoryPath;
                    openFileDialog.Filter = "Excel ファイル|*.xlsx;*.xlsm|すべてのファイル (*.*)|*.*";

                    // 複数ファイル選択を無効にする
                    openFileDialog.Multiselect = false;

                    // ファイル選択ダイアログを表示
                    if (openFileDialog.ShowDialog() == DialogResult.OK) {
                        filePath = openFileDialog.FileName;
                    }
                    else {
                        // キャンセルされた場合は、処理を中止
                        throw new OperationCanceledException("ファイル選択がキャンセルされました。");
                    }
                }

                try {
                    FileStream fileStreamReport = new(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    return new XSSFWorkbook(fileStreamReport);
                } catch (Exception ex) {
                    throw new Exception($"レポートテンプレートの読み込み中にエラーが発生しました: {ex.Message}", ex);
                }
            }

            // レポートシートに製品情報を挿入する
            private static void PopulateReportSheet(XSSFWorkbook reportWorkbook, ProductInformation productInfo, ReportConfig config) {
                var workSheetTemp = reportWorkbook.GetSheet(config.SheetName) ?? throw new Exception($"シート '{config.SheetName}' が見つかりません。");

                var productNumber = productInfo.ProductNumber.Split('-')[0] ?? string.Empty;
                SetValue(config.ProductNumberRange, productNumber);
                SetValue(config.OrderNumberRange, productInfo.OrderNumber);
                SetValue(config.QuantityRange, productInfo.Quantity.ToString());
                SetValue(config.SerialFirstRange, productInfo.SerialFirst);
                SetValue(config.SerialLastRange, productInfo.SerialLast);
                SetValue(config.ProductModelRange, productInfo.ProductModel);

                void SetValue(string? address, string? value) {
                    if (string.IsNullOrEmpty(address) || string.IsNullOrEmpty(value)) { return; }
                    GetRowColFromAddress(address, out int rowIndex, out int colIndex);
                    IRow row = workSheetTemp.GetRow(rowIndex) ?? workSheetTemp.CreateRow(rowIndex);
                    ICell cell = row.GetCell(colIndex) ?? row.CreateCell(colIndex);
                    cell.SetCellValue(value);
                }
                static int ColumnNameToIndex(string columnName) {
                    int index = 0;
                    foreach (char c in columnName.ToUpper()) {
                        if (c < 'A' || c > 'Z') throw new ArgumentException("Invalid column name");
                        index = index * 26 + (c - 'A' + 1);
                    }
                    return index - 1; // 0始まり
                }
                static void GetRowColFromAddress(string address, out int rowIndex, out int colIndex) {
                    // 数字の位置を検索
                    int i = 0;
                    while (i < address.Length && char.IsLetter(address[i])) i++;

                    string colPart = address.Substring(0, i);   // "A"
                    string rowPart = address.Substring(i);      // "2"

                    colIndex = ColumnNameToIndex(colPart);
                    rowIndex = int.Parse(rowPart) - 1; // 0始まり
                }
            }

            // 変更されたレポートをファイルに保存する
            private static void SaveReport(XSSFWorkbook reportWorkbook, ProductInformation productInfo, ReportConfig config) {
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
        public static class ReportGeneratorEPPlus {
            // レポートを生成するメインメソッド
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
            private static ExcelPackage LoadConfigWorkbook() {
                var configPath = Path.Combine(Environment.CurrentDirectory, "config", "General", "Excel", "ConfigReport.xlsx");
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
            private static ReportConfig GetReportConfig(ExcelPackage configWorkbook, string productModel) {
                var workSheetMain = configWorkbook.Workbook.Worksheets["Sheet1"] ?? throw new Exception("設定ファイルに 'Sheet1' が見つかりません。");

                // セル検索
                var searchAddressResult = workSheetMain.Cells.FirstOrDefault(x => x.Start.Column == 1 && x.Value?.ToString() == productModel)
                    ?? throw new Exception($"Configに品目番号:[{productModel}]が見つかりません。");
                var searchAddressResultRow = searchAddressResult.Start.Row;

                // ワークシートのセルから値を取得し、ReportConfigオブジェクトに格納
                var directoryPath = workSheetMain.Cells[searchAddressResultRow, 3].Value?.ToString()?.Trim('"');
                if (string.IsNullOrWhiteSpace(directoryPath)) { throw new Exception("Configのファイルパスが無効です。"); }
                if (!Directory.Exists(directoryPath)) { throw new FileNotFoundException($"指定されたフォルダが存在しません: {directoryPath}"); }

                var searchName = workSheetMain.Cells[searchAddressResultRow, 4].Value?.ToString()?.Trim('"');
                if (string.IsNullOrWhiteSpace(searchName)) { throw new Exception("Configのファイル名が無効です。"); }

                var filePaths = Directory.GetFiles(directoryPath, $"*{searchName}*", SearchOption.AllDirectories);
                var filePath = filePaths[0];
                var fileName = Path.GetFileNameWithoutExtension(filePath);
                var fileExtension = Path.GetExtension(filePath).ToLower();

                var sheetName = workSheetMain.Cells[searchAddressResultRow, 5].Value?.ToString();
                return string.IsNullOrWhiteSpace(sheetName)
                    ? throw new Exception("シート名がありません。")
                    : new ReportConfig {
                        DirectoryPath = directoryPath,
                        FileName = fileName,
                        FileExtension = fileExtension,
                        SearchFileName = searchName,
                        SheetName = sheetName,
                        ProductNumberRange = workSheetMain.Cells[searchAddressResultRow, 6].Value?.ToString() ?? string.Empty,
                        OrderNumberRange = workSheetMain.Cells[searchAddressResultRow, 7].Value?.ToString() ?? string.Empty,
                        QuantityRange = workSheetMain.Cells[searchAddressResultRow, 8].Value?.ToString() ?? string.Empty,
                        SerialFirstRange = workSheetMain.Cells[searchAddressResultRow, 9].Value?.ToString() ?? string.Empty,
                        SerialLastRange = workSheetMain.Cells[searchAddressResultRow, 10].Value?.ToString() ?? string.Empty,
                        ProductModelRange = workSheetMain.Cells[searchAddressResultRow, 11].Value?.ToString() ?? string.Empty,
                        SaveDirectory = workSheetMain.Cells[searchAddressResultRow, 12].Value?.ToString() ?? string.Empty
                    };
            }

            // レポートテンプレートExcelワークブックを読み込む
            private static ExcelPackage LoadReportTemplate(string directoryPath, string searchFileName) {
                var filePaths = Directory.GetFiles(directoryPath, $"{searchFileName}*", SearchOption.AllDirectories);
                if (filePaths.Length == 0) {
                    throw new FileNotFoundException($"指定されたファイル名 '{searchFileName}' のファイルが '{directoryPath}' に見つかりません。");
                }

                string filePath;

                if (filePaths.Length == 1) {
                    // ファイルが1つだけ見つかった場合は、それを自動的に選択
                    filePath = filePaths[0];
                }
                else {
                    MessageBox.Show("複数のファイルが見つかりました。1つ選択してください。");
                    // 複数のファイルが見つかった場合は、OpenFileDialog を使用してユーザーに選択させる
                    using var openFileDialog = new OpenFileDialog();
                    openFileDialog.InitialDirectory = directoryPath;
                    openFileDialog.Filter = "Excel ファイル|*.xlsx;*.xlsm|すべてのファイル (*.*)|*.*";

                    // 複数ファイル選択を無効にする
                    openFileDialog.Multiselect = false;

                    // ファイル選択ダイアログを表示
                    if (openFileDialog.ShowDialog() == DialogResult.OK) {
                        filePath = openFileDialog.FileName;
                    }
                    else {
                        // キャンセルされた場合は、処理を中止
                        throw new OperationCanceledException("ファイル選択がキャンセルされました。");
                    }
                }

                try {
                    FileStream fileStreamReport = new(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    return new ExcelPackage(fileStreamReport);
                } catch (Exception ex) {
                    throw new Exception($"レポートテンプレートの読み込み中にエラーが発生しました: {ex.Message}", ex);
                }
            }

            // レポートシートに製品情報を挿入する
            private static void PopulateReportSheet(ExcelPackage reportWorkbook, ProductInformation productInfo, ReportConfig config) {
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

            // 変更されたレポートをファイルに保存する
            private static void SaveReport(ExcelPackage reportWorkbook, ProductInformation productInfo, ReportConfig config) {
                var fileName = config.FileName;
                var fileExtension = config.FileExtension;
                var initialDirectory = config.SaveDirectory;

                using SaveFileDialog saveFileDialog = new() {
                    Filter = $"Excel Files (*{fileExtension})|*{fileExtension}|All Files (*.*)|*.*",
                    FileName = $"{fileName} のコピー{productInfo.ProductNumber}{fileExtension}",
                    Title = "保存先を選択してください",
                    InitialDirectory = initialDirectory ?? Environment.CurrentDirectory // Nullの場合はデフォルトディレクトリを使用
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

        // リスト生成
        public static class ListGeneratorNPOI {
            // リストを生成するメインメソッド
            public static void GenerateList(ProductInformation productInfo) {
                try {
                    // 1. Excel設定の読み込みとワークブックの準備
                    var workbook = LoadConfigWorkbook();

                    var (targetSheetName, productName, resultRow, resultRowIndex, workSheetMain) = LoadExcelConfiguration(workbook, productInfo.ProductModel);

                    // 2. 製品情報の設定とExcelへの書き込み
                    var productCellRanges = GetProductCellRanges(workSheetMain, resultRow);
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
                    var temporarilyPath = Path.Combine(Environment.CurrentDirectory, "config", "General", "Excel", "temporarilyList.xlsx");
                    SaveAndPrintExcel(workbook, temporarilyPath, targetSheetName);
                } catch (Exception ex) {
                    MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            // 設定Excelワークブックを読み込む
            private static XSSFWorkbook LoadConfigWorkbook() {
                var configPath = Path.Combine(Environment.CurrentDirectory, "config", "General", "Excel", "ConfigList.xlsx");
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
                for (int rowIndex = workSheetMain.FirstRowNum; rowIndex <= workSheetMain.LastRowNum; rowIndex++) {
                    IRow row = workSheetMain.GetRow(rowIndex);
                    if (row == null) continue;

                    ICell cell = row.GetCell(targetColumnIndex);
                    if (cell == null) continue;

                    string cellValue = GetCellValue(cell) ?? string.Empty;

                    if (cellValue.Equals(searchText, StringComparison.OrdinalIgnoreCase)) {
                        resultRowIndex = rowIndex;
                    }
                }

                if (resultRowIndex == -1) { throw new Exception($"Configに品目番号:[{searchText}]が見つかりません。"); }

                IRow resultRow = workSheetMain.GetRow(resultRowIndex);

                var sheetName = GetCellValue(resultRow.GetCell(1)) ??
                    throw new Exception($"設定シートの行 {resultRow}, 列 2 にシート名が設定されていません。");
                var productName = GetCellValue(resultRow.GetCell(2)) ?? string.Empty;

                return (sheetName, productName, resultRow, resultRowIndex, workSheetMain);
            }

            // 製品情報に関連するExcelのセル範囲を取得するヘルパーメソッド
            private static ProductCellRanges GetProductCellRanges(ISheet workSheetMain, IRow resultRow) {
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
                    IRow row = workSheetMain.GetRow(resultRowIndex);
                    if (row == null) continue;


                    // 各セルを走査
                    for (int colIndex = row.FirstCellNum; colIndex < row.LastCellNum; colIndex++) {
                        ICell cell = row.GetCell(colIndex);
                        if (cell == null) continue;

                        string cellValue = GetCellValue(cell) ?? string.Empty;

                        if (cellValue.Equals(searchText, StringComparison.OrdinalIgnoreCase)) {
                            resultColumnIndex = colIndex;
                        }
                    }

                    if (resultColumnIndex == -1) { throw new Exception($"Configに基板モデル:[{searchText}]が見つかりません。"); }

                    ICell cell2 = row.GetCell(resultColumnIndex + 1);

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
                GetRowColFromAddress(qrCodeRange, out int rowIndex, out int colIndex);
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
                GetRowColFromAddress(address, out int rowIndex, out int colIndex);
                IRow row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
                ICell cell = row.GetCell(colIndex) ?? row.CreateCell(colIndex);
                cell.SetCellValue(value);
            }
            private static void InsertImageOriginalSize(ISheet sheet, Bitmap bitmap, int rowIndex, int colIndex) {
                IWorkbook workbook = sheet.Workbook;

                // Bitmap をメモリストリームに保存 → バイト配列化
                byte[] bytes;
                using (var ms = new MemoryStream()) {
                    bitmap.Save(ms, ImageFormat.Png); // PNG形式で保存
                    bytes = ms.ToArray();
                }

                // ワークブックに画像を追加
                int pictureIdx = workbook.AddPicture(bytes, PictureType.PNG);

                // 描画オブジェクトを取得
                var drawing = sheet.CreateDrawingPatriarch();
                ICreationHelper helper = workbook.GetCreationHelper();
                var anchor = helper.CreateClientAnchor();
                anchor.AnchorType = AnchorType.DontMoveAndResize;

                // 画像を作成
                var picture = drawing.CreatePicture(anchor, pictureIdx);

                anchor.Col1 = colIndex;   // 左上列
                anchor.Row1 = rowIndex;   // 左上行
                anchor.Col2 = anchor.Col1;
                anchor.Row2 = anchor.Row1;

            }
            private static int ColumnNameToIndex(string columnName) {
                int index = 0;
                foreach (char c in columnName.ToUpper()) {
                    if (c < 'A' || c > 'Z') throw new ArgumentException("Invalid column name");
                    index = index * 26 + (c - 'A' + 1);
                }
                return index - 1; // 0始まり
            }
            private static void GetRowColFromAddress(string address, out int rowIndex, out int colIndex) {
                // 数字の位置を検索
                int i = 0;
                while (i < address.Length && char.IsLetter(address[i])) i++;

                string colPart = address.Substring(0, i);   // "A"
                string rowPart = address.Substring(i);      // "2"

                colIndex = ColumnNameToIndex(colPart);
                rowIndex = int.Parse(rowPart) - 1; // 0始まり
            }
        }
        public static class ListGeneratorEPPlus {
            // リストを生成するメインメソッド
            public static void GenerateList(ProductInformation productInfo) {
                try {
                    // 1. Excel設定の読み込みとワークブックの準備
                    var configPath = Path.Combine(Environment.CurrentDirectory, "config", "General", "Excel", "ConfigList.xlsx");
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
                    var temporarilyPath = Path.Combine(Environment.CurrentDirectory, "config", "General", "Excel", "temporarilyList.xlsx");
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
        // チェックシート生成
        public static class CheckSheetGeneratorNPOI {
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
                    var temporarilyPath = Path.Combine(Environment.CurrentDirectory, "config", "General", "Excel", "temporarilyCheckSheet.xlsx");
                    SaveAndPrintExcel(workbook, temporarilyPath);
                } catch (Exception ex) {
                    MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            // 設定Excelワークブックを読み込む
            private static XSSFWorkbook LoadConfigWorkbook() {
                var configPath = Path.Combine(Environment.CurrentDirectory, "config", "General", "Excel", "ConfigCheckSheet.xlsx");
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
                for (int rowIndex = workSheetMain.FirstRowNum; rowIndex <= workSheetMain.LastRowNum; rowIndex++) {
                    IRow row = workSheetMain.GetRow(rowIndex);
                    if (row == null) continue;

                    ICell cell = row.GetCell(targetColumnIndex);
                    if (cell == null) continue;

                    string cellValue = GetCellValue(cell) ?? string.Empty;

                    if (cellValue.Equals(searchText, StringComparison.OrdinalIgnoreCase)) {
                        resultRowIndex = rowIndex;
                    }
                }

                if (resultRowIndex == -1) { throw new Exception($"Configに品目番号:[{searchText}]が見つかりません。"); }

                IRow resultRow = workSheetMain.GetRow(resultRowIndex);

                return resultRow;
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

            // 製品情報に関連するExcelのセル範囲を取得するヘルパーメソッド
            private static CheckSheetConfigData GetProductCellRanges(IRow resultRow) {
                return new CheckSheetConfigData {
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
                        .TakeWhile(sheetName => !string.IsNullOrWhiteSpace(sheetName))],
                };
            }

            // 製品情報をExcelシートに書き込むメソッド
            private static void PopulateProductDetails(ISheet workSheetTemp, ProductInformation productInfo, CheckSheetConfigData ranges, string temperature, string humidity, string formattedDate) {
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

            // 指定されたEPPlusワークブックの不要なシートを非表示にします。
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

                // シート（"Sheet1"）を非表示リストに追加
                if (workBook.GetSheetIndex("Sheet1") != -1 && !sheetIndicesToHide.Contains(workBook.GetSheetIndex("Sheet1"))) {
                    sheetIndicesToHide.Add(workBook.GetSheetIndex("Sheet1"));
                }

                // シートを非表示に設定
                foreach (var sheetIndex in sheetIndicesToHide) {
                    workBook.SetSheetHidden(sheetIndex, true);
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
                GetRowColFromAddress(address, out int rowIndex, out int colIndex);
                IRow row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
                ICell cell = row.GetCell(colIndex) ?? row.CreateCell(colIndex);
                cell.SetCellValue(value);
            }
            private static int ColumnNameToIndex(string columnName) {
                int index = 0;
                foreach (char c in columnName.ToUpper()) {
                    if (c < 'A' || c > 'Z') throw new ArgumentException("Invalid column name");
                    index = index * 26 + (c - 'A' + 1);
                }
                return index - 1; // 0始まり
            }
            private static void GetRowColFromAddress(string address, out int rowIndex, out int colIndex) {
                // 数字の位置を検索
                int i = 0;
                while (i < address.Length && char.IsLetter(address[i])) i++;

                string colPart = address.Substring(0, i);   // "A"
                string rowPart = address.Substring(i);      // "2"

                colIndex = ColumnNameToIndex(colPart);
                rowIndex = int.Parse(rowPart) - 1; // 0始まり
            }
        }
        public static class CheckSheetGeneratorEPPlus {
            // Excelチェックシートを生成し、データを書き込み、印刷します。
            public static void GenerateCheckSheet(ProductInformation productInfo) {
                try {
                    // 設定ファイルのパスを構築
                    var configPath = Path.Combine(Environment.CurrentDirectory, "config", "General", "Excel", "ConfigCheckSheet.xlsx");
                    var temporarilyPath = Path.Combine(Environment.CurrentDirectory, "config", "General", "Excel", "temporarilyCheckSheet.xlsx");

                    // 1. 設定ファイルの読み込みとメインシートの取得
                    // EPPlusを使用して設定ファイルを読み込み、必要な情報を抽出
                    var excelData = LoadAndExtractConfig(configPath, productInfo);

                    // 2. 温度・湿度入力ダイアログの表示と値の取得
                    (var temperature, var humidity) = GetTemperatureAndHumidity(excelData);

                    // 3. 日付のフォーマット
                    var formattedDate = FormatDate(productInfo.RegDate, excelData.DateFormat);

                    // 4. EPPlusでExcelファイルを編集
                    // 設定ファイルからEPPlusワークブックを作成
                    using (var workBookEPPlus = WorkbookFactory.Create(configPath)) {
                        // セル参照をCellReferenceオブジェクトとして作成
                        var cellReferences = CreateCellReferences(excelData);

                        // 各シートに対して値を書き込む
                        PopulateExcelSheets(workBookEPPlus, productInfo, cellReferences, temperature, humidity, formattedDate, excelData.SheetNames);

                        // 不要なシートを非表示にする
                        HideSheets(workBookEPPlus, excelData.SheetNames);

                        // ブックを保存
                        SaveWorkbook(workBookEPPlus, temporarilyPath);
                    }

                    // 5. Excel Interopを使用して印刷
                    PrintExcelFile(temporarilyPath);
                } catch (Exception ex) {
                    // エラーメッセージを表示
                    MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            // Excel設定ファイルを読み込み、メインシートから設定データを抽出します。
            private static CheckSheetConfigData LoadAndExtractConfig(string configPath, ProductInformation productInfo) {
                if (!File.Exists(configPath)) {
                    throw new FileNotFoundException($"設定ファイルが見つかりません: {configPath}");
                }

                using FileStream fileStream = new(configPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using var workBook = new ExcelPackage(fileStream);

                // 既存ワークシートを取得
                var sheet = workBook.Workbook.Worksheets;
                var targetSheetName = "Sheet1";
                var workSheetMain = sheet[targetSheetName];

                // セル検索: ProductModelに基づいて行を特定
                var searchAddressResult = workSheetMain.Cells.FirstOrDefault(x => x.Start.Column == 1 && x.Value?.ToString() == productInfo.ProductModel)
                     ?? throw new Exception($"Configに品目番号:[{productInfo.ProductModel}]が見つかりません。");
                var resultRow = searchAddressResult.Start.Row;

                // ワークシートのセルから値を取得し、ExcelConfigDataオブジェクトに格納
                var excelData = new CheckSheetConfigData {
                    ProductModelRange = workSheetMain.Cells[resultRow, 3].Value?.ToString() ?? string.Empty,
                    ProductNumberRange = workSheetMain.Cells[resultRow, 4].Value?.ToString() ?? string.Empty,
                    OrderNumberRange = workSheetMain.Cells[resultRow, 5].Value?.ToString() ?? string.Empty,
                    QuantityRange = workSheetMain.Cells[resultRow, 6].Value?.ToString() ?? string.Empty,
                    SerialFirstRange = workSheetMain.Cells[resultRow, 7].Value?.ToString() ?? string.Empty,
                    SerialLastRange = workSheetMain.Cells[resultRow, 8].Value?.ToString() ?? string.Empty,
                    RegDateRange = workSheetMain.Cells[resultRow, 9].Value?.ToString() ?? string.Empty,
                    DateFormat = workSheetMain.Cells[resultRow, 10].Value?.ToString() ?? string.Empty,
                    RegTemperatureRange = workSheetMain.Cells[resultRow, 11].Value?.ToString() ?? string.Empty,
                    RegHumidityRange = workSheetMain.Cells[resultRow, 12].Value?.ToString() ?? string.Empty,
                    SheetNames = [.. Enumerable.Range(13, 20)
                        .Select(column => workSheetMain.Cells[resultRow, column].Value?.ToString() ?? string.Empty)
                        .TakeWhile(sheetName => !string.IsNullOrWhiteSpace(sheetName))]
                };

                return excelData.SheetNames.Count == 0 ? throw new Exception("対象シートがConfigファイルに設定されていません。") : excelData;
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

            // ExcelConfigDataからEPPlusのCellReferenceオブジェクトを作成します。
            private static dynamic CreateCellReferences(CheckSheetConfigData excelData) {
                // セル参照をCellReferenceオブジェクトとして作成
                return new {
                    ProductModel = string.IsNullOrEmpty(excelData.ProductModelRange) ? null : new CellReference(excelData.ProductModelRange),
                    ProductNumber = string.IsNullOrEmpty(excelData.ProductNumberRange) ? null : new CellReference(excelData.ProductNumberRange),
                    OrderNumber = string.IsNullOrEmpty(excelData.OrderNumberRange) ? null : new CellReference(excelData.OrderNumberRange),
                    Quantity = string.IsNullOrEmpty(excelData.QuantityRange) ? null : new CellReference(excelData.QuantityRange),
                    SerialFirst = string.IsNullOrEmpty(excelData.SerialFirstRange) ? null : new CellReference(excelData.SerialFirstRange),
                    SerialLast = string.IsNullOrEmpty(excelData.SerialLastRange) ? null : new CellReference(excelData.SerialLastRange),
                    RegDate = string.IsNullOrEmpty(excelData.RegDateRange) ? null : new CellReference(excelData.RegDateRange),
                    RegTemperature = string.IsNullOrEmpty(excelData.RegTemperatureRange) ? null : new CellReference(excelData.RegTemperatureRange),
                    RegHumidity = string.IsNullOrEmpty(excelData.RegHumidityRange) ? null : new CellReference(excelData.RegHumidityRange),
                };
            }

            // 指定されたEPPlusワークブックの各シートに製品情報を書き込みます。
            private static void PopulateExcelSheets(
                IWorkbook workBookEPPlus,
                ProductInformation productInfo,
                dynamic cellReferences,
                string temperature,
                string humidity,
                string formattedDate,
                List<string> sheetNames) {
                foreach (var sheetName in sheetNames) {
                    var sheetEPPlus = workBookEPPlus.GetSheet(sheetName) ?? throw new Exception($"シート[{sheetName}]が見つかりません。");
                    sheetEPPlus.ForceFormulaRecalculation = true; // 数式の再計算を強制

                    // 各セルに値を書き込む
                    WriteCellValue(sheetEPPlus, cellReferences.ProductModel, productInfo.ProductModel);
                    WriteCellValue(sheetEPPlus, cellReferences.ProductNumber, productInfo.ProductNumber);
                    WriteCellValue(sheetEPPlus, cellReferences.OrderNumber, productInfo.OrderNumber);
                    WriteCellValue(sheetEPPlus, cellReferences.Quantity, productInfo.Quantity.ToString());
                    WriteCellValue(sheetEPPlus, cellReferences.SerialFirst, productInfo.SerialFirst);
                    WriteCellValue(sheetEPPlus, cellReferences.SerialLast, productInfo.SerialLast);
                    WriteCellValue(sheetEPPlus, cellReferences.RegDate, formattedDate);
                    WriteCellValue(sheetEPPlus, cellReferences.RegTemperature, temperature);
                    WriteCellValue(sheetEPPlus, cellReferences.RegHumidity, humidity);
                }
            }

            // 指定されたシートのセルに値を書き込みます。
            private static void WriteCellValue(ISheet sheet, CellReference? reference, string value) {
                if (reference is null) {
                    return; // 参照が nullの場合は何もしない
                }

                // 指定した行を取得できない時はエラーとならないよう新規作成している
                var row = sheet.GetRow(reference.Row) ?? sheet.CreateRow(reference.Row);
                // 一行上の処理の列版
                var cell = row.GetCell(reference.Col) ?? row.CreateCell(reference.Col);
                cell.SetCellValue(value);
            }

            // 指定されたEPPlusワークブックの不要なシートを非表示にします。
            private static void HideSheets(IWorkbook workBookEPPlus, List<string> sheetsToKeep) {
                // 非表示にするシートのインデックスを特定
                var allSheetNames = new List<string>();
                for (var i = 0; i < workBookEPPlus.NumberOfSheets; i++) {
                    allSheetNames.Add(workBookEPPlus.GetSheetName(i));
                }

                var sheetIndicesToHide = allSheetNames
                    .Select((name, index) => new { Name = name, Index = index })
                    .Where(sheet => !sheetsToKeep.Contains(sheet.Name))
                    .Select(sheet => sheet.Index)
                    .ToList();

                // シート（"Sheet1"）を非表示にする
                if (workBookEPPlus.GetSheetIndex("Sheet1") != -1 && !sheetIndicesToHide.Contains(workBookEPPlus.GetSheetIndex("Sheet1"))) {
                    sheetIndicesToHide.Add(workBookEPPlus.GetSheetIndex("Sheet1"));
                }

                // シートを非表示に設定
                foreach (var sheetIndex in sheetIndicesToHide) {
                    workBookEPPlus.SetSheetHidden(sheetIndex, SheetVisibility.VeryHidden);
                }
            }

            // EPPlusワークブックをファイルに保存します。
            private static void SaveWorkbook(IWorkbook workBookEPPlus, string outputPath) {
                try {
                    using var fs = new FileStream(outputPath, FileMode.Create);
                    workBookEPPlus.Write(fs);
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
        public static class SubstrateInformation {
            // 基板設定ファイルを開くメソッド
            public static void OpenSubstrateInformation(ProductInformation productInfo) {
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
            private static ExcelPackage LoadConfigWorkbook() {
                var configPath = Path.Combine(Environment.CurrentDirectory, "config", "General", "Excel", "ConfigSubstrateInformation.xlsx");
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
            private static (string filePaths, string sheetName) GetSubstrateConfig(ExcelPackage configWorkbook, string substrateModel) {
                var workSheetMain = configWorkbook.Workbook.Worksheets["Sheet1"] ?? throw new Exception("設定ファイルのシートに Sheet1 が見つかりません。");

                // セル検索
                var searchAddressResult = workSheetMain.Cells.FirstOrDefault(x => x.Start.Column == 1 && x.Value?.ToString() == substrateModel)
                    ?? throw new Exception($"Configに型式:[{substrateModel}]が見つかりません。");
                var searchAddressResultRow = searchAddressResult.Start.Row;

                // ディレクトリパスとファイル名を取得(ディレクトリパスが空の場合は、Configの2行目の値を使用する)
                var directoryPath = workSheetMain.Cells[searchAddressResultRow, 4].Value?.ToString()?.Trim('"') ?? workSheetMain.Cells[2, 4].Value?.ToString()?.Trim('"') ?? string.Empty;
                var fileName = workSheetMain.Cells[searchAddressResultRow, 5].Value?.ToString()?.Trim('"') ?? string.Empty;
                if (string.IsNullOrEmpty(fileName)) { throw new Exception($"設定ファイルの型式 {substrateModel} の シート名 が空です。"); }
                var filePaths = Path.Combine(directoryPath, fileName);
                if (!File.Exists(filePaths)) { throw new FileNotFoundException($"指定されたファイルが存在しません: {filePaths}"); }

                var sheetName = workSheetMain.Cells[searchAddressResultRow, 6].Value?.ToString()?.Trim('"') ?? string.Empty;

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
    }
}
