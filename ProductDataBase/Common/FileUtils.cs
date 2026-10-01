namespace ProductDatabase.Common {
    // ファイルコピー・削除のリトライ処理と共有バックアップパス管理を提供するユーティリティクラス
    internal static class FileUtils {

        // 共有バックアップパス（BackupManager と Logger が共同利用する）
        public static string BackupPath { get; set; } = string.Empty;

        // ログのバックアップ先フォルダ(db/logs)へのコピーを行うかどうか。
        // 起動時にフォルダ不在を確認しユーザーが作成を拒否した場合、毎回のコピー失敗によるリトライ待ちを避けるため false にする
        public static bool LogBackupEnabled { get; set; } = true;

        // ファイルをコピーする（失敗時はリトライ）
        internal static void CopyWithRetry(string sourceFileName, string destFileName, bool overwrite, int retryCount = 5, int delayMilliseconds = 2000) {
            for (var attempt = 1; attempt <= retryCount; attempt++) {
                try {
                    if (File.Exists(sourceFileName)) {
                        File.Copy(sourceFileName, destFileName, overwrite);
                    }
                    return;
                } catch (IOException) {
                    if (attempt == retryCount) {
                        throw new Exception("バックアップファイルがコピーできません。");
                    }
                    Thread.Sleep(delayMilliseconds);
                }
            }
        }

        // ファイルを削除する（失敗時はリトライ）
        internal static void DeleteWithRetry(string filePath, int retryCount = 5, int delayMilliseconds = 2000) {
            for (var attempt = 1; attempt <= retryCount; attempt++) {
                try {
                    if (File.Exists(filePath)) {
                        File.Delete(filePath);
                    }
                    return;
                } catch (IOException) {
                    if (attempt == retryCount) {
                        throw new Exception("バックアップファイルが削除できません。");
                    }
                    Thread.Sleep(delayMilliseconds);
                }
            }
        }
    }
}
