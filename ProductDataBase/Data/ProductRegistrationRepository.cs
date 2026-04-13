using Dapper;
using ProductDatabase.Models;
using ProductDatabase.Other;
using System.Data;

namespace ProductDatabase.Data {
    // 製品登録処理（製品履歴・シリアル・基板使用INSERT、重複チェック）のDB操作を担当するリポジトリクラス
    internal static class ProductRegistrationRepository {

        // 製品履歴テーブルにレコードをINSERTして生成されたROWIDを返す
        public static long InsertProductRecord(
            IDbConnection connection, IDbTransaction transaction,
            ProductMaster productMaster, ProductRegisterWork productRegisterWork,
            string comment, int serialLastNumber) {

            var commandText =
                $"""
                INSERT INTO {Constants.TProductTableName}
                (
                    ProductID,
                    OrderNumber,
                    ProductNumber,
                    OLesNumber,
                    Quantity,
                    Person,
                    RegDate,
                    Revision,
                    RevisionGroup,
                    SerialFirst,
                    SerialLast,
                    SerialLastNumber,
                    Comment
                )
                VALUES
                (
                    @ProductID,
                    @OrderNumber,
                    @ProductNumber,
                    @OLesNumber,
                    @Quantity,
                    @Person,
                    @RegDate,
                    @Revision,
                    @RevisionGroup,
                    @SerialFirst,
                    @SerialLast,
                    @SerialLastNumber,
                    @Comment
                );
                """;

            connection.Execute(commandText, new {
                productMaster.ProductID,
                OrderNumber      = productRegisterWork.OrderNumber.NullIfWhiteSpace(),
                ProductNumber    = productRegisterWork.ProductNumber.NullIfWhiteSpace(),
                OLesNumber       = productRegisterWork.OLesNumber.NullIfWhiteSpace(),
                productRegisterWork.Quantity,
                Person           = productRegisterWork.Person.NullIfWhiteSpace(),
                RegDate          = productRegisterWork.RegDate.NullIfWhiteSpace(),
                Revision         = productRegisterWork.Revision.NullIfWhiteSpace(),
                productMaster.RevisionGroup,
                SerialFirst      = productRegisterWork.SerialFirst.NullIfWhiteSpace(),
                SerialLast       = productRegisterWork.SerialLast.NullIfWhiteSpace(),
                SerialLastNumber = productMaster.IsSerialGeneration ? (int?)serialLastNumber : null,
                Comment          = comment.NullIfWhiteSpace()
            }, transaction: transaction);

            return connection.ExecuteScalar<long>("SELECT last_insert_rowid();", transaction: transaction);
        }

        // シリアルリストを一括INSERTする
        public static void InsertSerials(IDbConnection connection, IDbTransaction transaction, IEnumerable<SerialInsertData> serialData) {
            var commandText =
                $"""
                INSERT INTO {Constants.TSerialTableName}
                (
                    Serial,
                    OLesSerial,
                    UsedID,
                    ProductName
                )
                VALUES
                (
                    @Serial,
                    @OLesSerial,
                    @ProductRowId,
                    @ProductName
                );
                """;
            connection.Execute(commandText, serialData, transaction: transaction);
        }

        // 基板IDと製造番号から注文番号を取得する
        public static string GetSubstrateOrderNumber(
            IDbConnection connection, IDbTransaction transaction,
            long substrateID, string substrateNumber) {

            var commandText =
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
                    AND IsDeleted = 0
                GROUP BY
                    OrderNumber
                ORDER BY ID ASC LIMIT 1;
                """;

            var result = connection.QueryFirstOrDefault<SubstrateStockInfo>(
                commandText,
                new { SubstrateID = substrateID, SubstrateNumber = substrateNumber },
                transaction: transaction);

            return result?.OrderNumber ?? string.Empty;
        }

        // 基板使用履歴テーブルに使用数（Decrease）をINSERTする
        public static void InsertSubstrateUsage(
            IDbConnection connection, IDbTransaction transaction,
            long substrateID, string substrateNumber, string orderNumber,
            long useValue, long? useID, string person, string regDate, string comment) {

            var commandText =
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
                );
                """;

            connection.Execute(commandText, new {
                SubstrateID     = substrateID,
                SubstrateNumber = substrateNumber.NullIfWhiteSpace(),
                OrderNumber     = orderNumber.NullIfWhiteSpace(),
                Decrease        = 0 - useValue,
                Person          = person.NullIfWhiteSpace(),
                RegDate         = regDate.NullIfWhiteSpace(),
                Comment         = comment.NullIfWhiteSpace(),
                UseID           = useID
            }, transaction: transaction);
        }

        // 登録したIDがDB上に存在するか確認し見つからない場合は例外をスローする
        public static void CheckRegistrationExists(IDbConnection connection, long rowId) {
            var exists = connection.ExecuteScalar<bool>(
                $"SELECT EXISTS(SELECT 1 FROM {Constants.VProductTableName} WHERE Id = @Id);",
                new { Id = rowId });

            if (!exists) {
                throw new Exception("登録に失敗しました。IDが見つかりません。");
            }
        }

        // 指定シリアルリストのうちDB既存シリアルを返す
        public static List<string> CheckSerialDuplication(IDbConnection connection, string productName, IList<string> serialList) {
            var sql =
                $"""
                SELECT
                    s.Serial
                FROM
                    {Constants.TSerialTableName} AS s
                LEFT JOIN
                    {Constants.VProductTableName} AS p
                ON
                    s.UsedID = p.ID
                WHERE
                    s.ProductName = @ProductName
                AND
                    s.Serial IN @SerialList;
                """;

            return connection.Query<string>(sql, new {
                ProductName = productName,
                SerialList  = serialList.Select(x => x.Trim()).ToList()
            }).ToList();
        }

        // 指定O-LesシリアルリストのうちDB既存シリアルを返す
        public static List<string> CheckOlesSerialDuplication(IDbConnection connection, string productName, IList<string> olesList) {
            var sql =
                $"""
                SELECT
                    s.OLesSerial
                FROM
                    {Constants.TSerialTableName} AS s
                WHERE
                    s.ProductName = @ProductName
                AND
                    s.OLesSerial IN @OlesList;
                """;

            return connection.Query<string>(sql, new {
                ProductName = productName,
                OlesList    = olesList
            }).ToList();
        }

        // OLesシリアルサフィックスを更新する
        public static void UpdateOLesSerialSuffix(IDbConnection connection, IDbTransaction transaction, long productId, string suffix) {
            connection.Execute(
                $"UPDATE {Constants.ProductTableName} SET OLesSerialSuffix = @Suffix WHERE ProductID = @ProductID;",
                new { Suffix = suffix, ProductID = productId },
                transaction: transaction);
        }

        // 基板IDの在庫データをDBから取得する（在庫0超の明細のみ）
        public static IEnumerable<SubstrateStockInfo> FetchSubstrateStock(IDbConnection connection, IDbTransaction transaction, long substrateID) {
            var commandText =
                $"""
                SELECT
                    SubstrateName,
                    SubstrateNumber,
                    SUM(COALESCE(Increase, 0) + COALESCE(Decrease, 0) + COALESCE(Defect, 0)) AS Stock
                FROM
                    {Constants.VSubstrateTableName}
                WHERE
                    SubstrateID = @SubstrateID
                    AND SubstrateNumber IS NOT NULL
                    AND IsDeleted = 0
                GROUP BY
                    SubstrateID,
                    SubstrateNumber
                HAVING
                    Stock > 0
                ORDER BY
                    MIN(ID);
                """;

            return connection.Query<SubstrateStockInfo>(commandText, new { SubstrateID = substrateID }, transaction: transaction);
        }

        // 製番重複チェック: 既存の製品モデルを返す（重複なければnull）
        public static string? FindDuplicateProductNumber(IDbConnection connection, IDbTransaction transaction, string productName, string productNumber) {
            var query = $"SELECT ProductModel FROM {Constants.VProductTableName} WHERE ProductName = @ProductName AND ProductNumber = @ProductNumber AND IsDeleted = 0 LIMIT 1";
            return connection.QueryFirstOrDefault<string>(query, new { ProductName = productName, ProductNumber = productNumber }, transaction: transaction);
        }

        // 注番重複チェック: 既存の製品モデルを返す（重複なければnull）
        public static string? FindDuplicateOrderNumber(IDbConnection connection, IDbTransaction transaction, string productName, string orderNumber) {
            var query = $"SELECT ProductModel FROM {Constants.VProductTableName} WHERE ProductName = @ProductName AND OrderNumber = @OrderNumber AND IsDeleted = 0 LIMIT 1";
            return connection.QueryFirstOrDefault<string>(query, new { ProductName = productName, OrderNumber = orderNumber }, transaction: transaction);
        }

        // 製品名・RevisionGroupに対応する最新Revisionを返す（なければnull）
        public static string? GetLatestRevision(IDbConnection connection, string productName, string revisionGroup) {
            var sql =
                $"""
                SELECT Revision
                FROM {Constants.VProductTableName}
                WHERE ProductName = @ProductName
                  AND RevisionGroup = @RevisionGroup
                  AND IsDeleted = 0
                ORDER BY ID DESC
                LIMIT 1
                """;
            return connection.QueryFirstOrDefault<string>(sql, new { ProductName = productName, RevisionGroup = revisionGroup });
        }

        // 製品名の最新SerialLastNumberを返す（なければnull）
        public static int? GetLatestSerialLastNumber(IDbConnection connection, string productName) {
            var sql =
                $"""
                SELECT SerialLastNumber
                FROM {Constants.VProductTableName}
                WHERE ProductName = @ProductName
                  AND IsDeleted = 0
                  AND SerialLastNumber IS NOT NULL
                ORDER BY ID DESC
                LIMIT 1
                """;
            return connection.QueryFirstOrDefault<int?>(sql, new { ProductName = productName });
        }

        // Revision変更レコードをINSERTして生成されたROWIDを返す
        public static long InsertRevisionChangeRecord(
            IDbConnection connection, IDbTransaction transaction,
            ProductMaster productMaster, ProductRegisterWork productRegisterWork, int serialLastNumber) {

            var sql =
                $"""
                INSERT INTO {Constants.TProductTableName}
                (
                    ProductID,
                    RegDate,
                    Revision,
                    RevisionGroup,
                    SerialLastNumber,
                    Comment
                )
                VALUES
                (
                    @ProductID,
                    @RegDate,
                    @Revision,
                    @RevisionGroup,
                    @SerialLastNumber,
                    @Comment
                );
                """;

            connection.Execute(sql, new {
                productMaster.ProductID,
                productRegisterWork.Revision,
                productRegisterWork.RegDate,
                productMaster.RevisionGroup,
                SerialLastNumber = serialLastNumber,
                Comment = productRegisterWork.Comment.NullIfWhiteSpace()
            }, transaction: transaction);

            return connection.ExecuteScalar<long>("SELECT last_insert_rowid();", transaction: transaction);
        }

        // 基板在庫情報（FetchSubstrateStock / GetSubstrateOrderNumber 共通）
        public sealed class SubstrateStockInfo {
            public string SubstrateName   { get; set; } = string.Empty;
            public string SubstrateNumber { get; set; } = string.Empty;
            public string OrderNumber     { get; set; } = string.Empty;
            public int    Stock           { get; set; }
        }
    }

    // シリアルINSERT用データレコード
    public sealed record SerialInsertData(string Serial, string? OLesSerial, int ProductRowId, string ProductName);
}
