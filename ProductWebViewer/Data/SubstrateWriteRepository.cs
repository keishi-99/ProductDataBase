using Dapper;
using Microsoft.Data.Sqlite;
using ProductWebViewer.Models;

namespace ProductWebViewer.Data {
    // 基板登録内容の削除（書き込み）を担当するリポジトリ
    // ReadOnlyのSubstrateRecordRepositoryとは別に、書き込み可能な接続を独自に保持する
    public class SubstrateWriteRepository {
        private readonly string _connectionString;

        public SubstrateWriteRepository(IConfiguration configuration) {
            _connectionString = WritableConnectionStringFactory.Create(configuration);
        }

        public SubstrateRecord? GetById(long id) {
            using var con = new SqliteConnection(_connectionString);
            return con.QueryFirstOrDefault<SubstrateRecord>("""
                SELECT
                    s.ID,
                    s.SubstrateID,
                    m.CategoryName,
                    m.ProductName,
                    s.SubstrateName,
                    s.SubstrateModel,
                    s.OrderNumber,
                    s.SubstrateNumber,
                    s.Increase,
                    s.Decrease,
                    s.Defect,
                    s.PersonInfo,
                    s.RegDate,
                    s.Comment,
                    s.CreatedAt,
                    s.UseID,
                    p.ProductName AS UseProductName,
                    p.OrderNumber AS UseOrderNumber,
                    p.ProductNumber AS UseProductNumber
                FROM V_Substrate AS s
                LEFT JOIN M_SubstrateDef AS m ON s.SubstrateID = m.SubstrateID
                LEFT JOIN V_Product AS p ON s.UseID = p.ID
                WHERE s.ID = @Id AND s.IsDeleted = 0
                """, new { Id = id });
        }

        // 基板登録履歴を論理削除する（在庫集計は IsDeleted=0 の行のみ合算するため、これだけで在庫数が戻る）
        // 対象行が既に削除されている場合は false を返す（呼び出し側は競合として扱う）
        public bool DeleteSubstrate(long id) {
            using var con = new SqliteConnection(_connectionString);
            con.Open();
            var affected = con.Execute("UPDATE T_Substrate SET IsDeleted = 1, DeletedAt = datetime('now', 'localtime') WHERE ID = @Id AND IsDeleted = 0", new { Id = id });
            return affected > 0;
        }
    }
}
