using Microsoft.Data.Sqlite;

namespace ProductDatabase.Other {
    public static class DatabaseHelper {
        public static string GetConnectionRegistration() {
            var productRegistryPath = Path.Combine(Environment.CurrentDirectory, "db", "ProductRegistry.db");
            return !File.Exists(productRegistryPath)
                ? throw new FileNotFoundException("ファイルが見つかりません。", productRegistryPath)
                : new SqliteConnectionStringBuilder() { DataSource = productRegistryPath, Pooling = false }.ToString();
        }
    }
}
