namespace ProductDatabase.Other {
    // DBファイルのタイムスタンプ付きバックアップ作成と古いバックアップの自動削除を管理するクラス
    internal static class BackupManager {
        private static readonly string _backupDirectory = Path.Combine(Environment.CurrentDirectory, "db", "backup");
        private static readonly string _originalFilePath = Path.Combine(Environment.CurrentDirectory, "db", "ProductRegistry.db");
        private static readonly int _maxBackupFiles = 20;
        private static readonly object _lockObject = new();

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

                    if (!string.IsNullOrEmpty(FileUtils.BackupPath)) {
                        var backupPath = Path.Combine(FileUtils.BackupPath, "db", "ProductRegistry.db");
                        if (Environment.CurrentDirectory != FileUtils.BackupPath) {
                            FileUtils.CopyWithRetry(_originalFilePath, backupPath, true);
                        }
                    }
                }
            } catch (Exception ex) {
                MessageBox.Show($"バックアップの作成中にエラーが発生しました: {ex.Message}", $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 当日分のバックアップが未作成の場合のみDBをバックアップフォルダにコピーする
        public static void CreateDailyBackup() {
            // フォルダ未設定
            if (string.IsNullOrWhiteSpace(FileUtils.BackupPath)) {
                MessageBox.Show("フォルダが設定されていません。バックアップは保存されません。",
                    string.Empty, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // ネットワークフォルダが見つからない
            if (!Directory.Exists(FileUtils.BackupPath)) {
                MessageBox.Show($"'{FileUtils.BackupPath}'\nが見つかりません。バックアップは保存されません。",
                    string.Empty, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var today = DateTime.Today;
            var year = today.Year;
            var month = today.Month;
            var day = today.Day;

            var backupFolder = Path.Combine(FileUtils.BackupPath, "db", "backup", $"{year}", $"{month:00}");
            var backupFile = Path.Combine(backupFolder, $"_bak_{year}-{month:00}-{day:00}.db");
            var productRegistryFile = Path.Combine(Environment.CurrentDirectory, "db", "ProductRegistry.db");

            if (!File.Exists(backupFile)) {
                Directory.CreateDirectory(backupFolder);
                File.Copy(productRegistryFile, backupFile, overwrite: false);
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
                MessageBox.Show($"バックアップの作成中にエラーが発生しました: {ex.Message}", $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
