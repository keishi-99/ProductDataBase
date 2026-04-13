using Dapper;
using Microsoft.Data.Sqlite;
using ProductDatabase.Models;
using ProductDatabase.Other;
using System.Data;

namespace ProductDatabase.Data {
    // 基板登録処理（入荷/不良 INSERT、重複チェック、在庫確認）のDB操作を担当するリポジトリクラス
    internal static class SubstrateRegistrationRepository {

        // 基板IDと製番に対応する既存登録を取得する（入荷重複確認用）
        public static async Task<SubstrateStockInfo?> FindPreviousRegistrationAsync(long substrateId, string substrateNumber) {
            using var con = new SqliteConnection(ProductRepository.GetConnectionRegistration());
            await con.OpenAsync();

            var sql =
                $"""
                SELECT
                    SubstrateID,
                    SubstrateName,
                    SubstrateModel,
                    SubstrateNumber,
                    OrderNumber,
                    SUM(COALESCE(Increase, 0) + COALESCE(Decrease, 0) + COALESCE(Defect, 0)) AS Stock
                FROM
                    {Constants.VSubstrateTableName}
                WHERE
                    SubstrateID = @SubstrateID
                    AND SubstrateNumber = @SubstrateNumber
                GROUP BY
                    SubstrateID,
                    SubstrateName,
                    SubstrateModel,
                    SubstrateNumber,
                    OrderNumber
                ORDER BY
                    MIN(ID)
                LIMIT 1;
                """;

            return await con.QueryFirstOrDefaultAsync<SubstrateStockInfo>(
                sql,
                new { SubstrateID = substrateId, SubstrateNumber = substrateNumber });
        }

        // 基板ID・モデル・製番に対応する在庫情報を取得する（不良処理の在庫チェック用）
        public static SubstrateStockInfo? GetDefectStock(
            IDbConnection connection, IDbTransaction transaction,
            long substrateId, string substrateModel, string substrateNumber) {

            var sql =
                $"""
                SELECT
                    SubstrateID,
                    SubstrateName,
                    SubstrateModel,
                    SubstrateNumber,
                    OrderNumber,
                    SUM(COALESCE(Increase, 0) + COALESCE(Decrease, 0) + COALESCE(Defect, 0)) AS Stock
                FROM
                    {Constants.VSubstrateTableName}
                WHERE
                    SubstrateID = @SubstrateID
                    AND SubstrateModel = @SubstrateModel
                    AND SubstrateNumber = @SubstrateNumber
                GROUP BY
                    SubstrateID,
                    SubstrateName,
                    SubstrateModel,
                    SubstrateNumber,
                    OrderNumber
                ORDER BY
                    MIN(ID)
                LIMIT 1;
                """;

            return connection.QueryFirstOrDefault<SubstrateStockInfo>(
                sql,
                new { SubstrateID = substrateId, SubstrateModel = substrateModel, SubstrateNumber = substrateNumber },
                transaction: transaction);
        }

        // 基板登録テーブルに入荷/不良レコードをINSERTして生成されたROWIDを返す
        public static string InsertSubstrateRecord(
            IDbConnection connection, IDbTransaction transaction,
            long substrateId,
            string substrateNumber,
            string orderNumber,
            int? increase,
            int? defect,
            string person,
            string regDate,
            string comment) {

            var sql =
                $"""
                INSERT INTO {Constants.TSubstrateTableName}
                (
                    SubstrateID,
                    SubstrateNumber,
                    OrderNumber,
                    Increase,
                    Defect,
                    Person,
                    RegDate,
                    Comment
                )
                VALUES
                (
                    @SubstrateID,
                    @SubstrateNumber,
                    @OrderNumber,
                    @Increase,
                    @Defect,
                    @Person,
                    @RegDate,
                    @Comment
                );
                """;

            connection.Execute(sql, new {
                SubstrateID     = substrateId,
                SubstrateNumber = substrateNumber.NullIfWhiteSpace(),
                OrderNumber     = orderNumber.NullIfWhiteSpace(),
                Increase        = (object?)increase ?? DBNull.Value,
                Defect          = defect.HasValue ? (object)$"-{defect.Value}" : DBNull.Value,
                Person          = person.NullIfWhiteSpace(),
                RegDate         = regDate.NullIfWhiteSpace(),
                Comment         = comment.NullIfWhiteSpace()
            }, transaction: transaction);

            return connection.ExecuteScalar<string>("SELECT last_insert_rowid();", transaction: transaction) ?? string.Empty;
        }

        // 基板IDの合計在庫数を返す（在庫表示用）
        public static int? GetTotalStock(long substrateId) {
            using var con = new SqliteConnection(ProductRepository.GetConnectionRegistration());

            var sql =
                $"""
                SELECT
                    SUM(COALESCE(Increase, 0) + COALESCE(Decrease, 0) + COALESCE(Defect, 0)) AS Stock
                FROM {Constants.VSubstrateTableName}
                WHERE SubstrateID = @SubstrateID
                GROUP BY SubstrateID, SubstrateName, SubstrateModel
                """;

            return con.QueryFirstOrDefault<int?>(sql, new { SubstrateID = substrateId });
        }

        // 基板在庫情報DTO
        public sealed class SubstrateStockInfo {
            public long   SubstrateID     { get; set; }
            public string SubstrateName   { get; set; } = string.Empty;
            public string SubstrateModel  { get; set; } = string.Empty;
            public string SubstrateNumber { get; set; } = string.Empty;
            public string OrderNumber     { get; set; } = string.Empty;
            public int    Stock           { get; set; }
        }
    }
}
