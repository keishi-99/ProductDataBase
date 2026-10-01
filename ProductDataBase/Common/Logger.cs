namespace ProductDatabase.Common {
    // 月次CSVログファイルへの追記と共有フォルダへのコピーを管理するクラス
    // 【重要】WebViewer側(ProductWebViewer/Data/AuditLogger.cs)が同じ列構成・エスケープ処理で
    // 別ファイル(log_web_yyyyMM.csv)に書き込み、閲覧側で両ファイルをマージしている。
    // ここの列構成・エスケープ処理を変更する場合は、AuditLogger.cs・LogViewerWindow.cs・
    // ProductWebViewer/Data/LogRecordRepository.cs も必ず追従させること
    internal static class Logger {
        // AppDomain.CurrentDomain.BaseDirectory を使用してファイルダイアログ等による CurrentDirectory の変化を回避する
        internal static readonly string _logDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "db", "logs");
        private static readonly Lock _lockObject = new();

        /// <summary>
        /// 作業ログを追記します。
        /// </summary>
        /// <param name="message">記録する作業内容</param>
        public static void AppendLog(string[] message) {
            // コピー失敗時のメッセージ表示はロック解放後に行う（lock中にダイアログを出すと、応答待ちの間 他スレッドのログ書き込みも止まってしまうため）
            string? copyFailureMessage = null;
            try {
                lock (_lockObject) {
                    if (!Directory.Exists(_logDirectory)) {
                        Directory.CreateDirectory(_logDirectory);
                    }

                    var logFileName = $"log_{DateTime.Now:yyyyMM}.csv";
                    var logFilePath = Path.Combine(_logDirectory, logFileName);

                    var logEntry = $"\"{DateTime.Now:yyyy-MM-dd HH:mm:ss}\",{string.Join(",", message.Select(m => $"\"{m.Replace("\"", "\"\"").Replace("\r\n", " ").Replace("\r", " ").Replace("\n", " ")}\""))}";

                    File.AppendAllText(logFilePath, logEntry + Environment.NewLine);

                    if (!string.IsNullOrEmpty(FileUtils.BackupPath) && FileUtils.LogBackupEnabled) {
                        var cloneFilePath = Path.Combine(FileUtils.BackupPath, "db", "logs", logFileName);
                        if (cloneFilePath != logFilePath) {
                            // ログ本体の追記は成功しているので、コピー失敗時はメッセージだけ記録し処理を継続する
                            try {
                                FileUtils.CopyWithRetry(logFilePath, cloneFilePath, true);
                            } catch (Exception ex) {
                                copyFailureMessage = $"ログのバックアップコピーに失敗しました:\n{ex.Message}";
                            }
                        }
                    }
                }
            } catch (Exception ex) {
                System.Diagnostics.Debug.WriteLine($"ログの書き込み中にエラーが発生しました: {ex.Message}");
            }

            if (copyFailureMessage is not null) {
                MessageBox.Show(copyFailureMessage, "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        /// <summary>
        /// エラーログを記録します。
        /// </summary>
        /// <param name="methodName">メソッド名</param>
        /// <param name="exception">例外</param>
        /// <param name="additionalInfo">追加情報（任意）</param>
        public static void AppendErrorLog(string methodName, Exception exception, string? additionalInfo = null) {
            // コピー失敗時のメッセージ表示はロック解放後に行う（lock中にダイアログを出すと、応答待ちの間 他スレッドのログ書き込みも止まってしまうため）
            string? copyFailureMessage = null;
            try {
                lock (_lockObject) {
                    if (!Directory.Exists(_logDirectory)) {
                        Directory.CreateDirectory(_logDirectory);
                    }

                    var now = DateTime.Now;
                    var errorFileName = $"error_{now:yyyyMM}.csv";
                    var errorFilePath = Path.Combine(_logDirectory, errorFileName);

                    var errorMessage = exception.InnerException?.Message ?? exception.Message;
                    var stackTrace = exception.StackTrace?.Replace("\r\n", " ").Replace("\r", " ").Replace("\n", " ") ?? "";
                    var info = additionalInfo?.Replace("\"", "\"\"").Replace("\r\n", " ").Replace("\r", " ").Replace("\n", " ") ?? "";

                    var logEntry = $"\"{now:yyyy-MM-dd HH:mm:ss}\",\"{methodName}\",\"{exception.GetType().Name}\",\"{errorMessage.Replace("\"", "\"\"")}\",\"{stackTrace}\",\"{info}\"";

                    File.AppendAllText(errorFilePath, logEntry + Environment.NewLine);

                    if (!string.IsNullOrEmpty(FileUtils.BackupPath) && FileUtils.LogBackupEnabled) {
                        var cloneFilePath = Path.Combine(FileUtils.BackupPath, "db", "logs", errorFileName);
                        if (cloneFilePath != errorFilePath) {
                            // ログ本体の追記は成功しているので、コピー失敗時はメッセージだけ記録し処理を継続する
                            try {
                                FileUtils.CopyWithRetry(errorFilePath, cloneFilePath, true);
                            } catch (Exception ex) {
                                copyFailureMessage = $"エラーログのバックアップコピーに失敗しました:\n{ex.Message}";
                            }
                        }
                    }
                }
            } catch {
                System.Diagnostics.Debug.WriteLine("エラーログの書き込み中に予期しないエラーが発生しました。");
            }

            if (copyFailureMessage is not null) {
                MessageBox.Show(copyFailureMessage, "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
    }
}
