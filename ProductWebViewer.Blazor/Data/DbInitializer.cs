using Dapper;
using Microsoft.Data.Sqlite;

namespace ProductWebViewer.Blazor.Data {
    // 起動時にDBファイルが存在しなければ Schema.sql からスキーマを作成し、練習用のダミーデータを投入する。
    // 本番の ProductRegistry.db には一切接続しない（このアプリ専用の別DB）。
    public class DbInitializer(IConfiguration configuration, IWebHostEnvironment environment, ILogger<DbInitializer> logger) : IHostedService {
        public Task StartAsync(CancellationToken cancellationToken) {
            var dbPath = configuration["DatabasePath"]
                ?? throw new InvalidOperationException("DatabasePath が appsettings.json に設定されていません。");
            var fullPath = Path.IsPathRooted(dbPath) ? dbPath : Path.Combine(AppContext.BaseDirectory, dbPath);

            if (File.Exists(fullPath)) {
                logger.LogInformation("DBファイルは既に存在するため初期化をスキップします: {Path}", fullPath);
                return Task.CompletedTask;
            }

            Directory.CreateDirectory(Path.GetDirectoryName(fullPath)!);
            logger.LogInformation("DBファイルが存在しないため新規作成します: {Path}", fullPath);

            var schemaSql = File.ReadAllText(Path.Combine(environment.ContentRootPath, "Data", "Schema.sql"));

            using var con = new SqliteConnection(new SqliteConnectionStringBuilder {
                DataSource = fullPath,
                Mode = SqliteOpenMode.ReadWriteCreate,
            }.ToString());
            con.Open();

            try {
                con.Execute("PRAGMA journal_mode=WAL;");
                con.Execute(schemaSql);
                SeedDummyData.Run(con);
            } catch {
                // 途中で失敗した場合、不完全なDBファイルを残すと次回起動時に「既に存在する」と判定されて
                // 初期化がスキップされてしまうため、削除して次回やり直せるようにする
                con.Close();
                File.Delete(fullPath);
                throw;
            }

            logger.LogInformation("ダミーデータの投入が完了しました。");
            return Task.CompletedTask;
        }

        public Task StopAsync(CancellationToken cancellationToken) => Task.CompletedTask;
    }
}
