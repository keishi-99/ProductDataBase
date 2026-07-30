using Microsoft.Data.Sqlite;

namespace ProductWebViewer.Data {
    // 書き込み用リポジトリ(ProductWriteRepository/SubstrateWriteRepository)が共通で使う接続文字列の組み立てを担う
    internal static class WritableConnectionStringFactory {
        public static string Create(IConfiguration configuration) {
            var dbPath = configuration["DatabasePath"]
                ?? throw new InvalidOperationException("DatabasePath が appsettings.json に設定されていません。");
            var fullPath = Path.IsPathRooted(dbPath)
                ? dbPath
                : Path.Combine(AppContext.BaseDirectory, dbPath);

            // UNCパス(\\server\share\... または //server/share/...)経由の書き込みはSQLiteのネットワークファイルロックが不完全なため許可しない
            if (fullPath.StartsWith(@"\\", StringComparison.Ordinal) || fullPath.StartsWith("//", StringComparison.Ordinal)) {
                throw new InvalidOperationException(
                    $"DatabasePath がネットワーク共有パスを指しています: {fullPath}\n" +
                    "WebViewerからの書き込みは、DBファイルが存在するPC上でローカルディスクに対して行ってください。");
            }

            return new SqliteConnectionStringBuilder {
                DataSource = fullPath,
                Mode = SqliteOpenMode.ReadWrite,
                Pooling = false
            }.ToString();
        }
    }
}
