using Dapper;
using Microsoft.Data.Sqlite;
using ProductWebViewer.Blazor.Models;

namespace ProductWebViewer.Blazor.Data {
    // 製品登録内容の編集・削除（書き込み）を担当するリポジトリ
    public class ProductWriteRepository(IConfiguration configuration) : RepositoryBase(configuration) {

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

        // 担当者(PersonID)・登録日(RegDate)・Revisionは編集対象に含めない（本家 ProductWebViewer の仕様に合わせている）
        // 対象行が別操作で既に削除されている場合は false を返す（呼び出し側は競合として扱う）
        public bool UpdateProduct(long id, string? orderNumber, string? productNumber, string? oLesNumber, string? comment) {
            try {
                using var con = new SqliteConnection(_connectionString);
                con.Open();
                var affected = con.Execute("""
                    UPDATE T_Product
                    SET
                        OrderNumber   = @OrderNumber,
                        ProductNumber = @ProductNumber,
                        OLesNumber    = @OLesNumber,
                        Comment       = @Comment
                    WHERE ID = @Id AND IsDeleted = 0
                    """, new { Id = id, OrderNumber = orderNumber, ProductNumber = productNumber, OLesNumber = oLesNumber, Comment = comment });
                return affected > 0;
            } catch (Exception ex) {
                throw new Exception(SqliteBusyErrorHelper.GetUserMessage(ex), ex);
            }
        }

        // 製品登録を論理削除し、連動する基板使用履歴を論理削除・シリアルを物理削除する
        // 対象行が既に削除されている場合は Success=false を返し、関連データには一切触れない
        public ProductDeleteResult DeleteProduct(long id) {
            try {
                using var con = new SqliteConnection(_connectionString);
                con.Open();
                using var tx = con.BeginTransaction();

                var affected = con.Execute("UPDATE T_Product SET IsDeleted = 1, DeletedAt = datetime('now', 'localtime') WHERE ID = @Id AND IsDeleted = 0", new { Id = id }, tx);
                if (affected == 0) {
                    tx.Rollback();
                    return new ProductDeleteResult(false, 0, 0);
                }

                var deletedSubstrates = con.Execute("UPDATE T_Substrate SET IsDeleted = 1, DeletedAt = datetime('now', 'localtime') WHERE UseID = @Id AND IsDeleted = 0", new { Id = id }, tx);
                var deletedSerials = con.Execute("DELETE FROM T_Serial WHERE UsedID = @Id", new { Id = id }, tx);

                tx.Commit();
                return new ProductDeleteResult(true, deletedSubstrates, deletedSerials);
            } catch (Exception ex) {
                throw new Exception(SqliteBusyErrorHelper.GetUserMessage(ex), ex);
            }
        }
    }

    public record ProductDeleteResult(bool Success, int DeletedSubstrateCount, int DeletedSerialCount);
}
