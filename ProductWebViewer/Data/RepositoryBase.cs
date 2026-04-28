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

            _connectionString = new SqliteConnectionStringBuilder {
                DataSource = fullPath,
                Mode = SqliteOpenMode.ReadOnly,
                Pooling = false
            }.ToString();

            try {
                using var con = new SqliteConnection(_connectionString);
                con.Open();
            } catch (SqliteException ex) {
                throw new FileNotFoundException($"DB ファイルが見つかりません: {fullPath}", ex);
            }
        }

        protected static string BuildOrderBy(Dictionary<string, string> cols, string sortCol, string sortDir, string defaultOrder) =>
            cols.TryGetValue(sortCol ?? "", out var col)
                ? $"{col} {(sortDir == "asc" ? "ASC" : "DESC")}"
                : defaultOrder;

        protected static string BuildLimitOffset(int pageSize, int page) =>
            pageSize > 0 ? $"LIMIT {pageSize} OFFSET {(page - 1) * pageSize}" : "";
    }
}
