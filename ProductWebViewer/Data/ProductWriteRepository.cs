using Dapper;
using Microsoft.Data.Sqlite;
using ProductWebViewer.Models;

namespace ProductWebViewer.Data {
    // 製品登録内容の編集・削除（書き込み）を担当するリポジトリ
    // ReadOnlyのProductRecordRepositoryとは別に、書き込み可能な接続を独自に保持する
    public class ProductWriteRepository {
        private readonly string _connectionString;

        public ProductWriteRepository(IConfiguration configuration) {
            var dbPath = configuration["DatabasePath"]
                ?? throw new InvalidOperationException("DatabasePath が appsettings.json に設定されていません。");
            var fullPath = Path.IsPathRooted(dbPath)
                ? dbPath
                : Path.Combine(AppContext.BaseDirectory, dbPath);

            // UNCパス(\\server\share\...)経由の書き込みはSQLiteのネットワークファイルロックが不完全なため許可しない
            if (fullPath.StartsWith(@"\\", StringComparison.Ordinal)) {
                throw new InvalidOperationException(
                    $"DatabasePath がネットワーク共有パスを指しています: {fullPath}\n" +
                    "WebViewerからの書き込みは、DBファイルが存在するPC上でローカルディスクに対して行ってください。");
            }

            _connectionString = new SqliteConnectionStringBuilder {
                DataSource = fullPath,
                Mode = SqliteOpenMode.ReadWrite,
                Pooling = false
            }.ToString();
        }

        public ProductRecord? GetById(long id) {
            using var con = new SqliteConnection(_connectionString);
            return con.QueryFirstOrDefault<ProductRecord>("""
                SELECT
                    v.ID,
                    v.ProductID,
                    p.CategoryName,
                    v.ProductName,
                    v.ProductModel,
                    v.ProductType,
                    v.OrderNumber,
                    v.ProductNumber,
                    v.OLesNumber,
                    v.Quantity,
                    v.PersonInfo,
                    v.RegDate,
                    v.Revision,
                    v.SerialFirst,
                    v.SerialLast,
                    v.Comment,
                    v.CreatedAt
                FROM V_Product AS v
                LEFT JOIN M_ProductDef AS p ON v.ProductID = p.ProductID
                WHERE v.ID = @Id AND v.IsDeleted = 0
                """, new { Id = id });
        }

        // 担当者(PersonID)は編集対象に含めない（変更不要のため現状維持）
        // 対象行が別操作で既に削除されている場合は false を返す（呼び出し側は競合として扱う）
        public bool UpdateProduct(long id, string? orderNumber, string? productNumber, string? oLesNumber, string? regDate, string? revision, string? comment) {
            using var con = new SqliteConnection(_connectionString);
            con.Open();
            var affected = con.Execute("""
                UPDATE T_Product
                SET
                    OrderNumber   = @OrderNumber,
                    ProductNumber = @ProductNumber,
                    OLesNumber    = @OLesNumber,
                    RegDate       = @RegDate,
                    Revision      = @Revision,
                    Comment       = @Comment
                WHERE ID = @Id AND IsDeleted = 0
                """, new { Id = id, OrderNumber = orderNumber, ProductNumber = productNumber, OLesNumber = oLesNumber, RegDate = regDate, Revision = revision, Comment = comment });
            return affected > 0;
        }

        // 製品登録を論理削除し、連動する基板使用履歴を論理削除・シリアルを物理削除する
        // 対象行が既に削除されている場合は false を返し、関連データには一切触れない
        public bool DeleteProduct(long id) {
            using var con = new SqliteConnection(_connectionString);
            con.Open();
            using var tx = con.BeginTransaction();

            var affected = con.Execute("UPDATE T_Product SET IsDeleted = 1, DeletedAt = datetime('now', 'localtime') WHERE ID = @Id AND IsDeleted = 0", new { Id = id }, tx);
            if (affected == 0) {
                tx.Rollback();
                return false;
            }

            con.Execute("UPDATE T_Substrate SET IsDeleted = 1, DeletedAt = datetime('now', 'localtime') WHERE UseID = @Id AND IsDeleted = 0", new { Id = id }, tx);
            con.Execute("DELETE FROM T_Serial WHERE UsedID = @Id", new { Id = id }, tx);

            tx.Commit();
            return true;
        }
    }
}
