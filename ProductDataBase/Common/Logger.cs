namespace ProductDatabase.Common {
    // 月次CSVログファイルへの追記と共有フォルダへのコピーを管理するクラス
    internal static class Logger {
        // AppDomain.CurrentDomain.BaseDirectory を使用してファイルダイアログ等による CurrentDirectory の変化を回避する
        internal static readonly string _logDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "db", "logs");
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
                System.Diagnostics.Debug.WriteLine($"ログの書き込み中にエラーが発生しました: {ex.Message}");
            }
        }

        /// <summary>
        /// エラーログを記録します。
        /// </summary>
        /// <param name="methodName">メソッド名</param>
        /// <param name="exception">例外</param>
        /// <param name="additionalInfo">追加情報（任意）</param>
        public static void AppendErrorLog(string methodName, Exception exception, string? additionalInfo = null) {
            try {
                lock (_lockObject) {
                    if (!Directory.Exists(_logDirectory)) {
                        Directory.CreateDirectory(_logDirectory);
                    }

                    var errorFileName = $"error_{DateTime.Now:yyyyMM}.csv";
                    var errorFilePath = Path.Combine(_logDirectory, errorFileName);

                    var errorMessage = exception.InnerException?.Message ?? exception.Message;
                    var stackTrace = exception.StackTrace?.Replace("\r\n", " ").Replace("\r", " ").Replace("\n", " ") ?? "";
                    var info = additionalInfo?.Replace("\"", "\"\"").Replace("\r\n", " ").Replace("\r", " ").Replace("\n", " ") ?? "";

                    var logEntry = $"\"{DateTime.Now:yyyy-MM-dd HH:mm:ss}\",\"{methodName}\",\"{exception.GetType().Name}\",\"{errorMessage.Replace("\"", "\"\"")}\",\"{stackTrace}\",\"{info}\"";

                    File.AppendAllText(errorFilePath, logEntry + Environment.NewLine);

                    if (!string.IsNullOrEmpty(FileUtils.BackupPath)) {
                        var cloneFilePath = Path.Combine(FileUtils.BackupPath, "db", "logs", errorFileName);
                        if (cloneFilePath != errorFilePath) {
                            FileUtils.CopyWithRetry(errorFilePath, cloneFilePath, true);
                        }
                    }
                }
            } catch {
                System.Diagnostics.Debug.WriteLine("エラーログの書き込み中に予期しないエラーが発生しました。");
            }
        }
    }
}
