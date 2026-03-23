using System.Data;
using System.Runtime.InteropServices;

namespace ProductDatabase.Other {

    internal static partial class NativeMethods {
        [LibraryImport("user32.dll", SetLastError = true)]
        public static partial void keybd_event(byte bVk, byte bScan, int dwFlags, int dwExtraInfo);
    }

    internal partial class CommonUtils {
        public static string BackupPath { get; set; } = string.Empty;
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
                        throw new Exception("バックアップファイルがコピーできません。");
                    }
                    Thread.Sleep(delayMilliseconds);
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
                        throw new Exception("バックアップファイルが削除できません。");
                    }
                    Thread.Sleep(delayMilliseconds);
                }
            }
        }
        // 月次CSVログファイルへの追記と共有フォルダへのコピーを管理するクラス
        public static class Logger {
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

                        if (!string.IsNullOrEmpty(BackupPath)) {
                            var cloneFilePath = Path.Combine(BackupPath, "db", "logs", logFileName);
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
        // DBファイルのタイムスタンプ付きバックアップ作成と古いバックアップの自動削除を管理するクラス
        public static class BackupManager {
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

                        CopyWithRetry(_originalFilePath, backupFilePath, true);
                        ManageBackupFiles();

                        if (!string.IsNullOrEmpty(BackupPath)) {
                            var backupPath = Path.Combine(BackupPath, "db", "ProductRegistry.db");
                            if (Environment.CurrentDirectory != BackupPath) {
                                CopyWithRetry(_originalFilePath, backupPath, true);
                            }
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
                    var backupFiles = Directory.GetFiles(_backupDirectory, "ProductRegistry_*.db")
                        .OrderBy(File.GetCreationTime)
                        .ToList();

                    while (backupFiles.Count > _maxBackupFiles) {
                        var oldestFile = backupFiles.First();
                        DeleteWithRetry(oldestFile);
                        backupFiles.RemoveAt(0);
                    }
                } catch (Exception ex) {
                    MessageBox.Show($"バックアップの作成中にエラーが発生しました: {ex.Message}", $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        // STAスレッドで処理を実行し Task として返す（COM Interop・ダイアログ用）
        // RunContinuationsAsynchronously: await後の継続処理がSTAスレッドで実行されるのを防ぐ
        public static Task RunOnStaThreadAsync(Action action) {
            var tcs = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            var thread = new Thread(() => {
                try {
                    action();
                    tcs.SetResult();
                } catch (Exception ex) {
                    tcs.SetException(ex);
                }
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            return tcs.Task;
        }

        /// <summary>
        /// 月コードを取得します（10月→X, 11月→Y, 12月→Z, それ以外→MM形式）。
        /// </summary>
        public static string ToMonthCode(DateTime date) =>
            date.Month switch {
                10 => "X",
                11 => "Y",
                12 => "Z",
                var m => m.ToString("00")
            };

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
    }
}
