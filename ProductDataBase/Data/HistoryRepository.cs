using Dapper;
using Microsoft.Data.Sqlite;
using ProductDatabase.Models;
using ProductDatabase.Other;
using System.Data;

namespace ProductDatabase.Data {
    // 履歴テーブルのクエリおよびCRUD操作を担当するリポジトリクラス
    internal static class HistoryRepository {

        // 基板登録履歴を取得する
        public static DataTable QuerySubstrateHistory(SubstrateMaster substrateMaster, bool allSubstrates) {
            var substrateCategoryFilter = !string.IsNullOrEmpty(substrateMaster.CategoryName) ? " AND s.CategoryName = @CategoryName" : string.Empty;
            var productNameFilter       = !string.IsNullOrEmpty(substrateMaster.ProductName)  ? " AND s.ProductName = @ProductName"  : string.Empty;
            var substrateIdFilter       = !allSubstrates ? " AND s.SubstrateID = @SubstrateID"                                        : string.Empty;

            var query = $"""
                SELECT
                    s.ID,
                    s.SubstrateID,
                    s.ProductName,
                    s.SubstrateName,
                    s.SubstrateModel,
                    s.OrderNumber,
                    s.SubstrateNumber,
                    s.Increase,
                    s.Decrease,
                    s.Defect,
                    p.ProductType,
                    p.ProductNumber,
                    p.OrderNumber,
                    s.Person,
                    s.RegDate,
                    s.Comment,
                    s.UseID,
                    s.CreatedAt
                FROM
                    {Constants.VSubstrateTableName} AS s
                LEFT JOIN
                    {Constants.VProductTableName} AS p
                ON
                    s.UseID = p.ID
                WHERE
                    s.IsDeleted = 0
                    {substrateCategoryFilter}
                    {productNameFilter}
                    {substrateIdFilter}
                ORDER BY
                    s.ID DESC;
                """;

            var p = new DynamicParameters();
            p.Add("@CategoryName", substrateMaster.CategoryName);
            p.Add("@ProductName",  substrateMaster.ProductName);
            p.Add("@SubstrateID",  substrateMaster.SubstrateID);
            return ExecuteQuery(query, p);
        }

        // 基板在庫サマリーを取得する
        public static DataTable QuerySubstrateStock(SubstrateMaster substrateMaster, bool allSubstrates, bool inStock, bool groupByModel) {
            var substrateCategoryFilter = !string.IsNullOrEmpty(substrateMaster.CategoryName) ? " AND CategoryName = @CategoryName" : string.Empty;
            var productNameFilter       = !string.IsNullOrEmpty(substrateMaster.ProductName)  ? " AND ProductName = @ProductName"  : string.Empty;
            var substrateId             = !allSubstrates ? substrateMaster.SubstrateID : 0;
            var substrateIdFilter       = (substrateId != 0) ? " AND SubstrateID = @SubstrateID" : string.Empty;
            var inStockFilter           = inStock ? " AND Stock > 0" : string.Empty;

            var selectClause = groupByModel
                ? "SubstrateID, ProductName, SubstrateName, SubstrateModel, SUM(COALESCE(Increase, 0) + COALESCE(Decrease, 0) + COALESCE(Defect, 0)) AS Stock"
                : "SubstrateID, ProductName, SubstrateName, SubstrateModel, SubstrateNumber, OrderNumber, SUM(COALESCE(Increase, 0) + COALESCE(Decrease, 0) + COALESCE(Defect, 0)) AS Stock";

            var groupByClause = groupByModel ? "SubstrateName, SubstrateID"                      : "SubstrateName, SubstrateID, SubstrateNumber, OrderNumber";
            var orderByClause = groupByModel ? "SubstrateModel"                                   : "MIN(ID) DESC";

            var query = $"""
                SELECT
                    {selectClause}
                FROM
                    {Constants.VSubstrateTableName}
                WHERE
                    IsDeleted = 0
                    {substrateCategoryFilter}
                    {productNameFilter}
                    {substrateIdFilter}
                GROUP BY
                    {groupByClause}
                HAVING
                    1=1 {inStockFilter}
                ORDER BY
                    {orderByClause};
                """;

            var p = new DynamicParameters();
            p.Add("@CategoryName", substrateMaster.CategoryName);
            p.Add("@ProductName",  substrateMaster.ProductName);
            p.Add("@SubstrateID",  substrateMaster.SubstrateID);
            return ExecuteQuery(query, p);
        }

        // 製品登録履歴を取得する
        public static DataTable QueryProductHistory(ProductMaster productMaster, bool allProducts) {
            var productCategoryFilter = !string.IsNullOrEmpty(productMaster.CategoryName) ? " AND CategoryName = @CategoryName" : string.Empty;
            var productNameFilter     = !string.IsNullOrEmpty(productMaster.ProductName)  ? " AND ProductName = @ProductName"  : string.Empty;
            var productId             = !allProducts ? productMaster.ProductID : 0;
            var productIdFilter       = (productId != 0) ? " AND ProductID = @ProductID" : string.Empty;

            var query = $"""
                SELECT
                    ID,
                    ProductID,
                    CategoryName,
                    ProductName,
                    ProductType,
                    ProductModel,
                    OrderNumber,
                    ProductNumber,
                    OLesNumber,
                    Quantity,
                    SerialFirst,
                    SerialLast,
                    Revision,
                    RevisionGroup,
                    SerialLastNumber,
                    Person,
                    RegDate,
                    Comment,
                    CreatedAt
                FROM
                    {Constants.VProductTableName}
                WHERE
                    IsDeleted = 0
                    {productCategoryFilter}
                    {productNameFilter}
                    {productIdFilter}
                ORDER BY
                    ID DESC;
                """;

            var p = new DynamicParameters();
            p.Add("@CategoryName", productMaster.CategoryName);
            p.Add("@ProductName",  productMaster.ProductName);
            p.Add("@ProductID",    productMaster.ProductID);
            return ExecuteQuery(query, p);
        }

        // シリアル番号履歴を取得する
        public static DataTable QuerySerialHistory(ProductMaster productMaster, bool allProducts) {
            var categoryFilter  = !string.IsNullOrEmpty(productMaster.CategoryName) ? " AND (p.CategoryName = @CategoryName OR p.CategoryName IS NULL)" : string.Empty;
            var productNameFilter = !string.IsNullOrEmpty(productMaster.ProductName) ? " AND s.ProductName = @ProductName" : string.Empty;
            var productId         = !allProducts ? productMaster.ProductID : 0;
            var productIdFilter   = (productId != 0) ? " AND ProductID = @ProductID" : string.Empty;

            var query = $"""
                SELECT
                    s.rowid,
                    s.Serial,
                    s.OLesSerial,
                    p.OrderNumber,
                    p.ProductNumber,
                    s.ProductName,
                    p.ProductType,
                    p.ProductModel,
                    p.RegDate,
                    s.usedID
                FROM
                    {Constants.TSerialTableName} AS s
                LEFT JOIN
                    {Constants.VProductTableName} AS p
                ON
                    s.UsedID = p.ID
                WHERE
                    1=1
                    {categoryFilter}
                    {productNameFilter}
                    {productIdFilter}
                ORDER BY
                    s.rowid DESC;
                """;

            var p = new DynamicParameters();
            p.Add("@CategoryName", productMaster.CategoryName);
            p.Add("@ProductName",  productMaster.ProductName);
            p.Add("@ProductID",    productMaster.ProductID);
            return ExecuteQuery(query, p);
        }

        // 再印刷履歴を取得する
        public static DataTable QueryReprintHistory(string productModel) {
            var productModelFilter = !string.IsNullOrEmpty(productModel) ? " AND ProductModel = @ProductModel" : string.Empty;

            var query = $"""
                SELECT
                    ID,
                    SerialPrintType,
                    ProductName,
                    OrderNumber,
                    ProductNumber,
                    ProductType,
                    ProductModel,
                    Quantity,
                    Person,
                    RegDate,
                    Revision,
                    SerialFirst,
                    SerialLast,
                    Comment,
                    CreatedAt
                FROM
                    T_Reprint
                WHERE
                    1=1
                    {productModelFilter}
                ORDER BY
                    ID DESC;
                """;

            var p = new DynamicParameters();
            p.Add("@ProductModel", productModel);
            return ExecuteQuery(query, p);
        }

        // 指定IDに紐づく使用基板一覧を取得する（使用基板サブウィンドウ用）
        public static DataTable QueryUsedSubstrates(long id) {
            var query = $"""
                SELECT
                    ID,
                    SubstrateName,
                    SubstrateModel,
                    SubstrateNumber,
                    Decrease
                FROM {Constants.VSubstrateTableName}
                WHERE UseID = @ID
                ORDER BY SubstrateModel ASC;
                """;

            return ExecuteQuery(query, new DynamicParameters(new { ID = id }));
        }

        // 基板IDで基板マスター情報を取得する（在庫調整画面用）
        public static dynamic? QuerySubstrateMasterById(long substrateId) {
            using var con = new SqliteConnection(ProductRepository.GetConnectionRegistration());
            var query = $"""
                SELECT
                    SubstrateID,
                    SubstrateName,
                    SubstrateModel,
                    ProductName,
                    RegType,
                    Checkbox,
                    SerialPrintType
                FROM {Constants.SubstrateTableName}
                WHERE SubstrateID = @SubstrateID;
                """;
            return con.QueryFirstOrDefault(query, new { SubstrateID = substrateId });
        }

        // 製品削除時に連動削除する基板履歴一覧を取得する
        public static IEnumerable<dynamic> GetSubstratesByUseId(IDbConnection connection, long id, IDbTransaction transaction) {
            var sql = $"""
                SELECT
                    ID,
                    OrderNumber,
                    SubstrateNumber,
                    ProductName,
                    SubstrateName,
                    SubstrateModel,
                    Increase,
                    Decrease,
                    Defect,
                    RegDate,
                    Person,
                    Comment,
                    UseID
                FROM {Constants.VSubstrateTableName}
                WHERE UseID = @ID;
                """;
            return connection.Query(sql, new { ID = id }, transaction);
        }

        // 製品削除時に連動削除するシリアル一覧を取得する
        public static IEnumerable<dynamic> GetSerialsByUsedId(IDbConnection connection, long id, IDbTransaction transaction) {
            var sql = $"""
                SELECT
                    rowid,
                    ProductName,
                    Serial,
                    UsedID
                FROM {Constants.TSerialTableName}
                WHERE UsedID = @ID;
                """;
            return connection.Query(sql, new { ID = id }, transaction);
        }

        // 基板履歴を更新する
        public static void UpdateSubstrateRow(IDbConnection connection, DataRow row, IDbTransaction transaction) {
            var sql = $"""
                UPDATE {Constants.TSubstrateTableName}
                SET
                    SubstrateNumber = @SubstrateNumber,
                    OrderNumber     = @OrderNumber,
                    Increase        = @Increase,
                    Decrease        = @Decrease,
                    Defect          = @Defect,
                    Person          = @Person,
                    RegDate         = @RegDate,
                    Comment         = @Comment,
                    UseId           = @UseId
                WHERE ID = @ID;
                """;
            connection.Execute(sql, new {
                SubstrateNumber = row["SubstrateNumber"],
                OrderNumber     = row["OrderNumber"],
                Increase        = row["Increase"],
                Decrease        = row["Decrease"],
                Defect          = row["Defect"],
                RegDate         = row["RegDate"],
                Person          = row["Person"],
                Comment         = row["Comment"],
                UseId           = row["UseId"],
                ID              = row["ID"]
            }, transaction);
        }

        // 基板履歴を論理削除する
        public static void DeleteSubstrateRow(IDbConnection connection, DataRow row, IDbTransaction transaction) {
            var sql = $"""
                UPDATE {Constants.TSubstrateTableName}
                SET
                    IsDeleted = 1,
                    DeletedAt = datetime('now', 'localtime')
                WHERE ID = @ID;
                """;
            connection.Execute(sql, new { ID = row["ID"] }, transaction);
        }

        // 製品履歴を更新する
        public static void UpdateProductRow(IDbConnection connection, DataRow row, IDbTransaction transaction) {
            var sql = $"""
                UPDATE {Constants.TProductTableName}
                SET
                    OrderNumber   = @OrderNumber,
                    ProductNumber = @ProductNumber,
                    OLesNumber    = @OLesNumber,
                    Person        = @Person,
                    RegDate       = @RegDate,
                    Revision      = @Revision,
                    RevisionGroup = @RevisionGroup,
                    Comment       = @Comment
                WHERE ID = @ID;
                """;
            connection.Execute(sql, new {
                ID            = row["ID"],
                OrderNumber   = row["OrderNumber"],
                ProductNumber = row["ProductNumber"],
                OLesNumber    = row["OLesNumber"],
                Person        = row["Person"],
                RegDate       = row["RegDate"],
                Revision      = row["Revision"],
                RevisionGroup = row["RevisionGroup"],
                Comment       = row["Comment"]
            }, transaction);
        }

        // 製品履歴を論理削除する
        public static void DeleteProductRow(IDbConnection connection, DataRow row, IDbTransaction transaction) {
            var sql = $"""
                UPDATE {Constants.TProductTableName}
                SET
                    IsDeleted = 1,
                    DeletedAt = datetime('now', 'localtime')
                WHERE ID = @ID;
                """;
            connection.Execute(sql, new { ID = row["ID"] }, transaction);
        }

        // 製品削除に連動して基板履歴を論理削除する
        public static void DeleteProductSubstrateRow(IDbConnection connection, DataRow row, IDbTransaction transaction) {
            var sql = $"""
                UPDATE {Constants.TSubstrateTableName}
                SET
                    IsDeleted = 1,
                    DeletedAt = datetime('now', 'localtime')
                WHERE UseID = @ID;
                """;
            connection.Execute(sql, new { ID = row["ID"] }, transaction);
        }

        // 製品削除に連動してシリアルを物理削除する
        public static void DeleteProductSerialRow(IDbConnection connection, DataRow row, IDbTransaction transaction) {
            var sql = $"""
                DELETE FROM {Constants.TSerialTableName}
                WHERE UsedID = @ID;
                """;
            connection.Execute(sql, new { ID = row["ID"] }, transaction);
        }

        // シリアルを物理削除する
        public static void DeleteSerialRow(IDbConnection connection, DataRow row, IDbTransaction transaction) {
            var sql = $"""
                DELETE FROM {Constants.TSerialTableName}
                WHERE rowid = @rowid;
                """;
            connection.Execute(sql, new { rowid = row["rowid"] }, transaction);
        }

        // クエリを実行してDataTableを返す共通ヘルパー
        private static DataTable ExecuteQuery(string query, DynamicParameters parameters) {
            using var con = new SqliteConnection(ProductRepository.GetConnectionRegistration());
            using var reader = con.ExecuteReader(query, parameters);
            var table = new DataTable();
            table.Load(reader);
            return table;
        }
    }
}
