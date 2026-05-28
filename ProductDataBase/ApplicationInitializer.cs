using ProductDatabase.Common;
using ProductDatabase.Data;
using ProductDatabase.Models;
using ProductDatabase.Services;

namespace ProductDatabase {
    // アプリケーション起動時の初期化を担当するクラス
    internal class ApplicationInitializer {

        private readonly string _jsonFilePath;

        public ApplicationInitializer(string configPath) {
            _jsonFilePath = configPath;
        }

        // 設定ファイル読み込み
        public GeneralSettings LoadSettings() {
            try {
                return SettingsLoader.Load(_jsonFilePath);
            } catch (Exception ex) {
                Logger.AppendErrorLog(nameof(LoadSettings), ex, "設定ファイル読み込み失敗");
                throw;
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

        // DB読み込み
        public void LoadDatabase(ProductRepository productRepository) {
            try {
                ProductRepository.MigrateSerialProductId();
                ProductRepository.MigrateExclusiveGroup();
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
