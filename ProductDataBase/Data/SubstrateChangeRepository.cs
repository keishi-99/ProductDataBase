using Dapper;
using Microsoft.Data.Sqlite;
using ProductDatabase.Common;
using ProductDatabase.Models;
using System.Data;

namespace ProductDatabase.Data {
    // 基板変更（SubstrateChange1/2）のDB操作を担当するリポジトリクラス
    internal static class SubstrateChangeRepository {

        // 基板在庫と使用済み数量を結合して返す（SubstrateChange2 LoadEvents用）
        public static IEnumerable<SubstrateStockRow> GetSubstrateStock(long rowId, string substrateModel) {
            using var con = new SqliteConnection(ProductRepository.GetConnectionRegistration());

            var sql =
                $"""
                WITH Used AS
                (
                    SELECT
                        SubstrateNumber,
                        COALESCE(SUM(Decrease), 0) AS UsedDecrease
                    FROM
                        {Constants.VSubstrateTableName}
                    WHERE
                        SubstrateModel = @SubstrateModel
                        AND UseID = @ID
                        AND SubstrateNumber IS NOT NULL
                        AND IsDeleted = 0
                    GROUP BY
                        SubstrateNumber
                ),

                Stocked AS
                (
                    SELECT
                        SubstrateName,
                        SubstrateModel,
                        SubstrateNumber,
                        COALESCE(SUM(COALESCE(Increase, 0) + COALESCE(Decrease, 0) + COALESCE(Defect, 0)), 0) AS Stock
                    FROM
                        {Constants.VSubstrateTableName}
                    WHERE
                        SubstrateModel = @SubstrateModel
                        AND SubstrateNumber IS NOT NULL
                        AND IsDeleted = 0
                    GROUP BY
                        SubstrateName,
                        SubstrateModel,
                        SubstrateNumber
                )

                SELECT
                    s.SubstrateName,
                    s.SubstrateModel,
                    s.SubstrateNumber,
                    s.Stock,
                    COALESCE(u.UsedDecrease * -1, 0) AS UsedDecrease
                FROM
                    Stocked s
                    LEFT JOIN Used u ON s.SubstrateNumber = u.SubstrateNumber
                WHERE
                    s.Stock > 0
                    OR u.UsedDecrease IS NOT NULL
                ORDER BY
                    CASE WHEN COALESCE(u.UsedDecrease, 0) = 0 THEN 1 ELSE 0 END,
                    s.SubstrateNumber;
                """;

            return con.Query<SubstrateStockRow>(sql, new { ID = rowId, SubstrateModel = substrateModel });
        }

        // 基板番号に紐づく基板IDと注文番号を取得する（Registration UPDATE/INSERT 前の情報取得用）
        public static SubstrateLocationInfo? GetSubstrateInfo(
            IDbConnection connection, IDbTransaction transaction,
            long substrateId, string substrateNumber) {

            var sql =
                $"""
                SELECT
                    SubstrateID,
                    OrderNumber
                FROM
                    {Constants.VSubstrateTableName}
                WHERE
                    SubstrateID = @SubstrateID
                    AND SubstrateNumber = @SubstrateNumber
                    AND IsDeleted = 0
                GROUP BY
                    SubstrateID,
                    SubstrateNumber,
                    OrderNumber
                ORDER BY
                    MIN(ID)
                LIMIT 1
                """;

            return connection.QueryFirstOrDefault<SubstrateLocationInfo>(
                sql,
                new { SubstrateID = substrateId, SubstrateNumber = substrateNumber },
                transaction: transaction);
        }

        // 基板テーブルの使用数を更新して更新件数を返す（0件なら INSERT が必要）
        public static int UpdateSubstrateDecrease(
            IDbConnection connection, IDbTransaction transaction,
            int decrease, string person, string regDate, string comment,
            string substrateNumber, long useId) {

            var sql =
                $"""
                UPDATE
                    {Constants.TSubstrateTableName}
                SET
                    Decrease = @Decrease,
                    Person = @Person,
                    RegDate = @RegDate,
                    Comment = @Comment
                WHERE
                    SubstrateNumber = @SubstrateNumber
                    AND IsDeleted = 0
                    AND UseID = @UseID
                """;

            return connection.Execute(sql, new {
                Decrease = 0 - decrease,
                Person = person.NullIfWhiteSpace(),
                RegDate = regDate.NullIfWhiteSpace(),
                Comment = comment.NullIfWhiteSpace(),
                SubstrateNumber = substrateNumber,
                UseID = useId
            }, transaction: transaction);
        }

        // 基板テーブルに使用数レコードを新規挿入する
        public static void InsertSubstrateDecrease(
            IDbConnection connection, IDbTransaction transaction,
            string substrateId, string substrateNumber, string orderNumber,
            int decrease, string person, string regDate, string comment, long useId) {

            var sql =
                $"""
                INSERT INTO {Constants.TSubstrateTableName}
                    (
                        SubstrateID,
                        SubstrateNumber,
                        OrderNumber,
                        Decrease,
                        Person,
                        RegDate,
                        Comment,
                        UseID
                    )
                VALUES
                    (
                        @SubstrateID,
                        @SubstrateNumber,
                        @OrderNumber,
                        @Decrease,
                        @Person,
                        @RegDate,
                        @Comment,
                        @UseID
                    )
                """;

            connection.Execute(sql, new {
                SubstrateID = substrateId.NullIfWhiteSpace(),
                SubstrateNumber = substrateNumber.NullIfWhiteSpace(),
                OrderNumber = orderNumber.NullIfWhiteSpace(),
                Decrease = 0 - decrease,
                Person = person.NullIfWhiteSpace(),
                RegDate = regDate.NullIfWhiteSpace(),
                Comment = comment.NullIfWhiteSpace(),
                UseID = useId
            }, transaction: transaction);
        }

        // 基板テーブルの使用数を0にクリアする（チェックボックスOFF時）
        public static void ClearSubstrateDecrease(
            IDbConnection connection, IDbTransaction transaction,
            long substrateId, string substrateNumber, long useId) {

            var sql =
                $"""
                UPDATE
                    {Constants.TSubstrateTableName}
                SET
                    Decrease = @Decrease
                WHERE
                    SubstrateID = @SubstrateID
                    AND SubstrateNumber = @SubstrateNumber
                    AND IsDeleted = 0
                    AND UseID = @UseID
                """;

            connection.Execute(sql, new {
                Decrease = 0,
                SubstrateID = substrateId,
                SubstrateNumber = substrateNumber,
                UseID = useId
            }, transaction: transaction);
        }

        // 製品テーブルの登録情報を更新する
        public static void UpdateProduct(
            IDbConnection connection, IDbTransaction transaction,
            ProductRegisterWork work, int revisionGroup) {

            var sql =
                $"""
                UPDATE
                    {Constants.TProductTableName}
                SET
                    Quantity = @Quantity,
                    Person = @Person,
                    RegDate = @RegDate,
                    Revision = @Revision,
                    RevisionGroup = @RevisionGroup,
                    SerialLast = @SerialLast,
                    SerialLastNumber = @SerialLastNumber,
                    Comment = @Comment
                WHERE
                    ProductNumber = @ProductNumber
                    AND SerialFirst = @SerialFirst
                    AND IsDeleted = 0
                """;

            connection.Execute(sql, new {
                work.Quantity,
                Person = work.Person.NullIfWhiteSpace(),
                RegDate = work.RegDate.NullIfWhiteSpace(),
                Revision = work.Revision.NullIfWhiteSpace(),
                RevisionGroup = revisionGroup,
                SerialFirst = work.SerialFirst.NullIfWhiteSpace(),
                SerialLast = work.SerialLast.NullIfWhiteSpace(),
                work.SerialLastNumber,
                Comment = work.Comment.NullIfWhiteSpace(),
                ProductNumber = work.ProductNumber.NullIfWhiteSpace()
            }, transaction: transaction);
        }

        // 複数台登録対象の製品履歴をDBから取得して返す（SubstrateChange1 LoadEvents用）
        public static DataTable GetProductHistory(long productId) {
            using var con = new SqliteConnection(ProductRepository.GetConnectionRegistration());

            var sql =
                $"""
                SELECT
                    ID,
                    ProductName,
                    ProductType,
                    ProductModel,
                    OrderNumber,
                    ProductNumber,
                    Quantity,
                    SerialFirst,
                    SerialLast,
                    SerialLastNumber,
                    Revision,
                    RevisionGroup,
                    Person,
                    RegDate,
                    Comment,
                    CreatedAt
                FROM
                    {Constants.VProductTableName}
                WHERE
                    ProductID = @ProductID
                    AND Quantity > 1
                ORDER BY
                    ID DESC
                ;
                """;

            var result = con.ExecuteReader(sql, new { ProductID = productId });
            var table = new DataTable();
            table.Load(result);
            return table;
        }

        // 基板在庫行DTO
        public sealed class SubstrateStockRow {
            public string SubstrateName { get; init; } = string.Empty;
            public string SubstrateModel { get; init; } = string.Empty;
            public string SubstrateNumber { get; init; } = string.Empty;
            public int Stock { get; init; }
            public int UsedDecrease { get; init; }
        }

        // 基板ID・注文番号の位置情報DTO
        public sealed class SubstrateLocationInfo {
            public string SubstrateID { get; init; } = string.Empty;
            public string OrderNumber { get; init; } = string.Empty;
        }
    }
}
