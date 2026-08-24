using Microsoft.Data.Sqlite;

namespace ProductWebViewer.Blazor.Data {
    // 本家 ProductWebViewer は「デスクトップアプリと同じDBファイルを共有する」ため
    // 読み取り専用接続と書き込み専用接続を分離しているが、
    // この練習用プロジェクトは他プロセスと共有しない専用DBのため、その分離は不要と判断しシンプルな1本の接続文字列にしている。
    public abstract class RepositoryBase {
        protected readonly string _connectionString;

        internal string ConnectionString => _connectionString;

        protected RepositoryBase(IConfiguration configuration) {
            var dbPath = configuration["DatabasePath"]
                ?? throw new InvalidOperationException("DatabasePath が appsettings.json に設定されていません。");

            var fullPath = Path.IsPathRooted(dbPath)
                ? dbPath
                : Path.Combine(AppContext.BaseDirectory, dbPath);

            _connectionString = new SqliteConnectionStringBuilder {
                DataSource = fullPath,
                Mode = SqliteOpenMode.ReadWriteCreate,
            }.ToString();
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
