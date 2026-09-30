using ProductDatabase.Common;
using ProductDatabase.Data;
using ProductDatabase.Models;
using ProductDatabase.Services;

namespace ProductDatabase {
    // アプリケーション起動時の初期化を担当するクラス
    internal class ApplicationInitializer(string configPath) {

        private readonly string _jsonFilePath = configPath;

        // 設定ファイル読み込み
        public GeneralSettings LoadSettings() {
            try {
                return SettingsLoader.Load(_jsonFilePath);
            } catch (Exception ex) {
                Logger.AppendErrorLog(nameof(LoadSettings), ex, "設定ファイル読み込み失敗");
                throw;
            }
        }

        // バックアップ先にDBバックアップ用の親フォルダ(db/backup)がなければ確認の上で作成する。
        // 年月ごとのサブフォルダ(db/backup/年/月)は月が変わるたびに増えるため対象にせず、今までどおり自動作成する。
        // 作成しない場合は false を返し、呼び出し側で当日の日次バックアップをスキップする
        public bool EnsureDbBackupFolder(string backupPath) {
            try {
                if (string.IsNullOrWhiteSpace(backupPath) || !Directory.Exists(backupPath)) {
                    // バックアップ先自体が未設定・未検出の場合は CreateDailyBackup 側で警告済みのためここでは何もしない
                    return true;
                }

                var dbBackupFolder = Path.Combine(backupPath, "db", "backup");
                if (Directory.Exists(dbBackupFolder)) return true;

                var result = MessageBox.Show(
                    $"バックアップ先\n'{dbBackupFolder}'\nにDBバックアップ用フォルダがありません。作成しますか？\n作成しない場合、本日の日次バックアップは保存されません。",
                    "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (result != DialogResult.Yes) return false;

                Directory.CreateDirectory(dbBackupFolder);
                return true;
            } catch (Exception ex) {
                Logger.AppendErrorLog(nameof(EnsureDbBackupFolder), ex, "DBバックアップ用フォルダの確認に失敗しました");
                return false;
            }
        }

        // バックアップ作成（エラーは非致命的）
        public void CreateDailyBackup(string backupPath) {
            try {
                FileUtils.BackupPath = backupPath;
                BackupManager.CreateDailyBackup();
            } catch (Exception ex) {
                Logger.AppendErrorLog(nameof(CreateDailyBackup), ex, "日次バックアップ作成失敗");
                // 非致命的エラーなので継続
            }
        }

        // バックアップ先にログ用フォルダ(db/logs)がなければ確認の上で作成する。
        // 未作成のまま登録を行うと、ログ追記のたびにコピー失敗→リトライ待ち（最大10秒）が発生するため、起動時に一度だけ確認する
        public void EnsureLogBackupFolder(string backupPath) {
            try {
                if (string.IsNullOrWhiteSpace(backupPath) || !Directory.Exists(backupPath)) {
                    // バックアップ先自体が未設定・未検出の場合は CreateDailyBackup 側で警告済みのためここでは何もしない
                    return;
                }

                var logFolder = Path.Combine(backupPath, "db", "logs");
                if (Directory.Exists(logFolder)) return;

                var result = MessageBox.Show(
                    $"バックアップ先\n'{logFolder}'\nにログ用フォルダがありません。作成しますか？\n作成しない場合、ログのバックアップコピーは行われません。",
                    "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (result == DialogResult.Yes) {
                    Directory.CreateDirectory(logFolder);
                } else {
                    FileUtils.LogBackupEnabled = false;
                }
            } catch (Exception ex) {
                Logger.AppendErrorLog(nameof(EnsureLogBackupFolder), ex, "バックアップ用ログフォルダの確認に失敗しました");
                FileUtils.LogBackupEnabled = false;
            }
        }

        // DB読み込み
        public void LoadDatabase(ProductRepository productRepository) {
            try {
                ProductRepository.VerifyViewDefinitions();
                productRepository.LoadAll();
            } catch (Exception ex) {
                Logger.AppendErrorLog(nameof(LoadDatabase), ex, "DB読み込み失敗");
                throw;
            }
        }

        // 認証・権限設定
        public AppSettings ConfigureAppSettings(GeneralSettings generalSettings) {
            try {
                ArgumentNullException.ThrowIfNull(generalSettings);
                var appSettings = new AppSettings {
                    PersonList = generalSettings.Persons != null ? [.. generalSettings.Persons] : [],
                    IsAuthorizedUser = generalSettings.AuthorizedUsers?.Contains(Environment.UserName, StringComparer.OrdinalIgnoreCase) ?? false,
                    IsAdministrator = generalSettings.Administrators?.Contains(Environment.UserName, StringComparer.OrdinalIgnoreCase) ?? false
                };
                return appSettings;
            } catch (Exception ex) {
                Logger.AppendErrorLog(nameof(ConfigureAppSettings), ex, "アプリ設定に失敗");
                throw;
            }
        }

        // バーコード・QRサービス初期化
        public BarcodeService CreateBarcodeService(GeneralSettings generalSettings) {
            try {
                return new BarcodeService(generalSettings.DSN, generalSettings.UID, generalSettings.PWD);
            } catch (Exception ex) {
                Logger.AppendErrorLog(nameof(CreateBarcodeService), ex, "バーコードサービス初期化失敗");
                throw;
            }
        }
    }
}
