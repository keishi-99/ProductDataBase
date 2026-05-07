using Microsoft.Data.Sqlite;

namespace ProductWebViewer.Data {
    public abstract class RepositoryBase {
        protected readonly string _connectionString;

        protected RepositoryBase(IConfiguration configuration) {
            var dbPath = configuration["DatabasePath"]
                ?? throw new InvalidOperationException("DatabasePath が appsettings.json に設定されていません。");

            var fullPath = Path.IsPathRooted(dbPath)
                ? dbPath
                : Path.Combine(AppContext.BaseDirectory, dbPath);

            // ReadOnly: WebViewer からの誤った書き込みを防止する
            // Pooling=false: プールで保持された接続が WAL ロックを占有し、デスクトップ側の書き込みをブロックするのを避けるため
            _connectionString = new SqliteConnectionStringBuilder {
                DataSource = fullPath,
                Mode = SqliteOpenMode.ReadOnly,
                Pooling = false
            }.ToString();

            // 起動時に接続を試みることで、最初のリクエスト時ではなく起動直後にDB不在を検出する
            try {
                using var con = new SqliteConnection(_connectionString);
                con.Open();
            } catch (SqliteException ex) {
                throw new FileNotFoundException($"DB ファイルが見つかりません: {fullPath}", ex);
            }
        }

        // ORDER BY 句はパラメータ化できないため、cols ホワイトリストで SQLインジェクションを防ぐ
        protected static string BuildOrderBy(Dictionary<string, string> cols, string sortCol, string sortDir, string defaultOrder) =>
            cols.TryGetValue(sortCol ?? "", out var col)
                ? $"{col} {(sortDir == "asc" ? "ASC" : "DESC")}"
                : defaultOrder;

        // pageSize=0 は「全件取得」を意味するため空文字を返す
        protected static string BuildLimitOffset(int pageSize, int page) =>
            pageSize > 0 ? $"LIMIT {pageSize} OFFSET {(page - 1) * pageSize}" : "";
    }
}
