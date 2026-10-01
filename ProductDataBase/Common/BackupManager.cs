namespace ProductDatabase.Common {
    // DBファイルのタイムスタンプ付きバックアップ作成と古いバックアップの自動削除を管理するクラス
    internal static class BackupManager {
        private static readonly string _backupDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "db", "backup");
        private static readonly string _originalFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "db", "ProductRegistry.db");
        private static readonly int _maxBackupFiles = 20;
        private static readonly Lock _lockObject = new();

        /// <summary>
        /// バックアップを作成します。
        /// </summary>
        public static void CreateBackup() {
            try {
                lock (_lockObject) {
                    if (!Directory.Exists(_backupDirectory)) {
                        Directory.CreateDirectory(_backupDirectory);
                    }

                    var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                    var backupFileName = $"ProductRegistry_{timestamp}.db";
                    var backupFilePath = Path.Combine(_backupDirectory, backupFileName);

                    FileUtils.CopyWithRetry(_originalFilePath, backupFilePath, true);
                    ManageBackupFiles();
                }
            } catch (Exception ex) {
                Logger.AppendErrorLog(nameof(CreateBackup), ex, "バックアップ作成時にエラーが発生しました");
                // lockの外でダイアログを表示する（lock中に表示すると、応答待ちの間 他スレッドのバックアップ処理も止まってしまうため）
                MessageBox.Show($"バックアップの作成中にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 当日分のバックアップが未作成の場合のみDBをバックアップフォルダにコピーする
        public static void CreateDailyBackup() {
            // lockの外でダイアログを表示する（lock中に表示すると、応答待ちの間 他スレッドのバックアップ処理も止まってしまうため）
            string? infoMessage = null;
            try {
                lock (_lockObject) {
                    // フォルダ未設定
                    if (string.IsNullOrWhiteSpace(FileUtils.BackupPath)) {
                        Logger.AppendErrorLog(nameof(CreateDailyBackup), new InvalidOperationException("バックアップフォルダが設定されていません"), "設定確認が必要");
                        infoMessage = "フォルダが設定されていません。バックアップは保存されません。";
                        return;
                    }

                    // ネットワークフォルダが見つからない
                    if (!Directory.Exists(FileUtils.BackupPath)) {
                        Logger.AppendErrorLog(nameof(CreateDailyBackup), new DirectoryNotFoundException($"バックアップフォルダが見つかりません: {FileUtils.BackupPath}"), null);
                        infoMessage = $"'{FileUtils.BackupPath}'\nが見つかりません。バックアップは保存されません。";
                        return;
                    }

                    var today = DateTime.Today;
                    var year = today.Year;
                    var month = today.Month;
                    var day = today.Day;

                    var backupFolder = Path.Combine(FileUtils.BackupPath, "db", "backup", $"{year}", $"{month:00}");
                    var backupFile = Path.Combine(backupFolder, $"_bak_{year}-{month:00}-{day:00}.db");
                    var productRegistryFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "db", "ProductRegistry.db");

                    if (!File.Exists(backupFile)) {
                        // 年月ごとのサブフォルダは、親フォルダ(db/backup)の存在が EnsureDbBackupFolder で確認済みの前提のため、ここでは確認なしで自動作成する
                        Directory.CreateDirectory(backupFolder);
                        File.Copy(productRegistryFile, backupFile, overwrite: false);
                    }
                }
            } catch (Exception ex) {
                Logger.AppendErrorLog(nameof(CreateDailyBackup), ex, "日次バックアップの作成に失敗しました");
            } finally {
                if (infoMessage is not null) {
                    MessageBox.Show(infoMessage, string.Empty, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        /// <summary>
        /// 古いバックアップファイルを削除します。
        /// </summary>
        private static void ManageBackupFiles() {
            try {
                var backupFiles = Directory.GetFiles(_backupDirectory, "ProductRegistry_*.db")
                    .OrderBy(File.GetCreationTime)
                    .ToList();

                while (backupFiles.Count > _maxBackupFiles) {
                    var oldestFile = backupFiles.First();
                    FileUtils.DeleteWithRetry(oldestFile);
                    backupFiles.RemoveAt(0);
                }
            } catch (Exception ex) {
                Logger.AppendErrorLog(nameof(ManageBackupFiles), ex, $"古いバックアップファイルの削除に失敗しました");
            }
        }
    }
}
