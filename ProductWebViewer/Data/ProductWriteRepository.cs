using Dapper;
using Microsoft.Data.Sqlite;
using ProductWebViewer.Models;

namespace ProductWebViewer.Data {
    // 製品登録内容の編集・削除（書き込み）を担当するリポジトリ
    // ReadOnlyのProductRecordRepositoryとは別に、書き込み可能な接続を独自に保持する
    public class ProductWriteRepository(IConfiguration configuration) {
        private readonly string _connectionString = WritableConnectionStringFactory.Create(configuration);

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

        // 担当者(PersonID)・登録日(RegDate)・Revisionは編集対象に含めない
        // （メインアプリのHistoryEditDialogでも同項目はラベル表示のみで編集不可のため、それに合わせている）
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
        // DeletedSubstrates/DeletedSerials は監査ログ記録用に、削除直前の状態を返す
        public ProductDeleteResult DeleteProduct(long id) {
            try {
                using var con = new SqliteConnection(_connectionString);
                con.Open();
                using var tx = con.BeginTransaction();

                var affected = con.Execute("UPDATE T_Product SET IsDeleted = 1, DeletedAt = datetime('now', 'localtime') WHERE ID = @Id AND IsDeleted = 0", new { Id = id }, tx);
                if (affected == 0) {
                    tx.Rollback();
                    return new ProductDeleteResult(false, [], []);
                }

                // 監査ログ用に、連動削除される直前の基板使用履歴・シリアルを取得しておく
                var substrates = con.Query<CascadeSubstrateRow>("""
                    SELECT ID, OrderNumber, SubstrateNumber, ProductName, SubstrateName, SubstrateModel,
                           Increase, Decrease, Defect, RegDate, PersonInfo, Comment, UseID
                    FROM V_Substrate
                    WHERE UseID = @Id AND IsDeleted = 0
                    """, new { Id = id }, tx).ToList();
                var serials = con.Query<CascadeSerialRow>("""
                    SELECT rowid AS RowId, ProductName, Serial, UsedID
                    FROM V_Serial
                    WHERE UsedID = @Id
                    """, new { Id = id }, tx).ToList();

                con.Execute("UPDATE T_Substrate SET IsDeleted = 1, DeletedAt = datetime('now', 'localtime') WHERE UseID = @Id AND IsDeleted = 0", new { Id = id }, tx);
                con.Execute("DELETE FROM T_Serial WHERE UsedID = @Id", new { Id = id }, tx);

                tx.Commit();
                return new ProductDeleteResult(true, substrates, serials);
            } catch (Exception ex) {
                throw new Exception(SqliteBusyErrorHelper.GetUserMessage(ex), ex);
            }
        }
    }

    public record ProductDeleteResult(bool Success, IReadOnlyList<CascadeSubstrateRow> DeletedSubstrates, IReadOnlyList<CascadeSerialRow> DeletedSerials);

    // 製品削除に連動して削除される基板使用履歴・シリアルのスナップショット（監査ログ記録用）
    // PersonInfoはV_SubstrateビューでCOALESCE式から生成される列で宣言型を持たないため、
    // 該当データが0件の場合にMicrosoft.Data.Sqliteがbyte[]型と誤判定することがある。
    // positional recordだとDapperがコンストラクタの型完全一致を要求してしまうため、
    // プロパティ方式（列ごとに代入）に変更してこの型不一致を回避している。
    public record CascadeSubstrateRow {
        public long ID { get; init; }
        public string? OrderNumber { get; init; }
        public string? SubstrateNumber { get; init; }
        public string? ProductName { get; init; }
        public string? SubstrateName { get; init; }
        public string? SubstrateModel { get; init; }
        public long? Increase { get; init; }
        public long? Decrease { get; init; }
        public long? Defect { get; init; }
        public string? RegDate { get; init; }
        public string? PersonInfo { get; init; }
        public string? Comment { get; init; }
        public long? UseID { get; init; }
    }

    public record CascadeSerialRow(long RowId, string? ProductName, string? Serial, long? UsedID);
}
