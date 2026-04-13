namespace ProductDatabase.Other {
    // 月次CSVログファイルへの追記と共有フォルダへのコピーを管理するクラス
    internal static class Logger {
        private static readonly string _logDirectory = Path.Combine(Environment.CurrentDirectory, "db", "logs");
        private static readonly object _lockObject = new();

        /// <summary>
        /// 作業ログを追記します。
        /// </summary>
        /// <param name="message">記録する作業内容</param>
        public static void AppendLog(string[] message) {
            try {
                lock (_lockObject) {
                    if (!Directory.Exists(_logDirectory)) {
                        Directory.CreateDirectory(_logDirectory);
                    }

                    var logFileName = $"log_{DateTime.Now:yyyyMM}.csv";
                    var logFilePath = Path.Combine(_logDirectory, logFileName);

                    var logEntry = $"\"{DateTime.Now:yyyy-MM-dd HH:mm:ss}\",{string.Join(",", message.Select(m => $"\"{m.Replace("\"", "\"\"").Replace("\r\n", " ").Replace("\r", " ").Replace("\n", " ")}\""))}";

                    File.AppendAllText(logFilePath, logEntry + Environment.NewLine);

                    if (!string.IsNullOrEmpty(FileUtils.BackupPath)) {
                        var cloneFilePath = Path.Combine(FileUtils.BackupPath, "db", "logs", logFileName);
                        if (cloneFilePath != logFilePath) {
                            FileUtils.CopyWithRetry(logFilePath, cloneFilePath, true);
                        }
                    }
                }
            } catch (Exception ex) {
                MessageBox.Show($"ログの書き込み中にエラーが発生しました: {ex.Message}", $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
