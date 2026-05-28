using Microsoft.Data.Sqlite;

namespace ProductDatabase.Data {
    // データベース接続作成を統一管理するヘルパークラス
    internal static class DbConnectionHelper {

        // 接続文字列を取得する（プーリング無効・バックアップの安全性を優先）
        public static string GetConnectionString() {
            var productRegistryPath = Path.Combine(AppContext.BaseDirectory, "db", "ProductRegistry.db");
            if (!File.Exists(productRegistryPath)) {
                throw new FileNotFoundException("ファイルが見つかりません。", productRegistryPath);
            }
            return new SqliteConnectionStringBuilder() { DataSource = productRegistryPath, Pooling = false }.ToString();
        }

        // 接続を開いて返す（using ブロック用）
        public static SqliteConnection CreateAndOpenConnection() {
            var connection = new SqliteConnection(GetConnectionString());
            try {
                connection.Open();
                return connection;
            } catch {
                connection.Dispose();
                throw;
            }
        }
    }
}
