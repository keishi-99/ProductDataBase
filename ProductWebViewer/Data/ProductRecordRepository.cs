using Dapper;
using Microsoft.Data.Sqlite;
using ProductWebViewer.Models;

namespace ProductWebViewer.Data {
    public class ProductRecordRepository : RepositoryBase {
        // クエリパラメータ名 → 実テーブルの列参照へのホワイトリスト（ORDER BY インジェクション対策）
        private static readonly Dictionary<string, string> _productSortCols = new(StringComparer.OrdinalIgnoreCase) {
            ["ID"] = "v.ID",
            ["CategoryName"] = "p.CategoryName",
            ["ProductName"] = "v.ProductName",
            ["ProductModel"] = "v.ProductModel",
            ["ProductType"] = "v.ProductType",
            ["OrderNumber"] = "v.OrderNumber",
            ["ProductNumber"] = "v.ProductNumber",
            ["OLesNumber"] = "v.OLesNumber",
            ["Quantity"] = "v.Quantity",
            ["Person"] = "v.Person",
            ["RegDate"] = "v.RegDate",
            ["Revision"] = "v.Revision",
            ["SerialFirst"] = "v.SerialFirst",
            ["SerialLast"] = "v.SerialLast",
            ["Comment"] = "v.Comment",
            ["CreatedAt"] = "v.CreatedAt",
        };

        // シリアル一覧用ホワイトリスト（同上）
        private static readonly Dictionary<string, string> _serialSortCols = new(StringComparer.OrdinalIgnoreCase) {
            ["RowId"] = "s.rowid",
            ["Serial"] = "s.Serial",
            ["OLesSerial"] = "s.OLesSerial",
            ["OrderNumber"] = "p.OrderNumber",
            ["ProductNumber"] = "p.ProductNumber",
            ["ProductName"] = "s.ProductName",
            ["ProductType"] = "p.ProductType",
            ["ProductModel"] = "p.ProductModel",
            ["RegDate"] = "p.RegDate",
            ["CreatedAt"] = "p.CreatedAt",
        };

        public ProductRecordRepository(IConfiguration configuration) : base(configuration) { }

        public IReadOnlyList<string> GetCategoryList() {
            using var con = new SqliteConnection(_connectionString);
            return con.Query<string>(
                "SELECT DISTINCT CategoryName FROM M_ProductDef WHERE Visible=1 ORDER BY CategoryName"
            ).AsList();
        }

        public IReadOnlyList<string> GetProductNameList(string? category) {
            using var con = new SqliteConnection(_connectionString);
            var sql = string.IsNullOrWhiteSpace(category)
                ? "SELECT DISTINCT ProductName FROM M_ProductDef WHERE Visible=1 ORDER BY ProductName"
                : "SELECT DISTINCT ProductName FROM M_ProductDef WHERE Visible=1 AND CategoryName=@Category ORDER BY ProductName";
            return con.Query<string>(sql, new { Category = category }).AsList();
        }

        public IReadOnlyList<string> GetProductTypeList(string? category, string? productName) {
            using var con = new SqliteConnection(_connectionString);
            var conditions = new List<string> { "Visible=1", "ProductType IS NOT NULL", "ProductType != ''" };
            if (!string.IsNullOrWhiteSpace(category))
                conditions.Add("CategoryName=@Category");
            if (!string.IsNullOrWhiteSpace(productName))
                conditions.Add("ProductName=@ProductName");
            var sql = $"SELECT DISTINCT ProductType FROM M_ProductDef WHERE {string.Join(" AND ", conditions)} ORDER BY ProductType";
            return con.Query<string>(sql, new { Category = category, ProductName = productName }).AsList();
        }

        public int GetCount(
            string? listCategory = null,
            string? listProductName = null,
            string? listProductType = null,
            string? productName = null,
            string? orderNumber = null,
            string? productNumber = null,
            string? regDateFrom = null,
            string? regDateTo = null) {

            using var con = new SqliteConnection(_connectionString);
            var (where, param) = BuildProductWhere(listCategory, listProductName, listProductType, productName, orderNumber, productNumber, regDateFrom, regDateTo);
            return con.ExecuteScalar<int>($"""
                SELECT COUNT(*)
                FROM V_Product AS v
                LEFT JOIN M_ProductDef AS p ON v.ProductID = p.ProductID
                WHERE {where}
                """, param);
        }

        public IReadOnlyList<ProductRecord> GetAll(
            string? listCategory = null,
            string? listProductName = null,
            string? listProductType = null,
            string? productName = null,
            string? orderNumber = null,
            string? productNumber = null,
            string? regDateFrom = null,
            string? regDateTo = null,
            string sortCol = "",
            string sortDir = "desc",
            int page = 1,
            int pageSize = 100) {

            using var con = new SqliteConnection(_connectionString);
            var (where, param) = BuildProductWhere(listCategory, listProductName, listProductType, productName, orderNumber, productNumber, regDateFrom, regDateTo);
            var orderBy = BuildOrderBy(_productSortCols, sortCol, sortDir, "v.ID DESC");
            var limitOffset = BuildLimitOffset(pageSize, page);

            return con.Query<ProductRecord>($"""
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
                    v.Person,
                    v.RegDate,
                    v.Revision,
                    v.SerialFirst,
                    v.SerialLast,
                    v.Comment,
                    v.CreatedAt
                FROM V_Product AS v
                LEFT JOIN M_ProductDef AS p ON v.ProductID = p.ProductID
                WHERE {where}
                ORDER BY {orderBy}
                {limitOffset}
                """, param).AsList();
        }

        public int GetSerialCount(
            string? listCategory = null,
            string? listProductName = null,
            string? listProductType = null,
            string? productName = null,
            string? orderNumber = null,
            string? productNumber = null,
            string? regDateFrom = null,
            string? regDateTo = null,
            string? serial = null
            ) {

            using var con = new SqliteConnection(_connectionString);
            var (where, param) = BuildSerialWhere(listCategory, listProductName, listProductType, productName, orderNumber, productNumber, regDateFrom, regDateTo, serial);
            return con.ExecuteScalar<int>($"""
                SELECT COUNT(*)
                FROM T_Serial AS s
                LEFT JOIN V_Product AS p ON s.UsedID = p.ID
                WHERE {where}
                """, param);
        }

        public IReadOnlyList<SerialRecord> GetSerialHistory(
            string? listCategory = null,
            string? listProductName = null,
            string? listProductType = null,
            string? productName = null,
            string? orderNumber = null,
            string? productNumber = null,
            string? regDateFrom = null,
            string? regDateTo = null,
            string? serial = null,
            string sortCol = "",
            string sortDir = "desc",
            int page = 1,
            int pageSize = 100) {

            using var con = new SqliteConnection(_connectionString);
            var (where, param) = BuildSerialWhere(listCategory, listProductName, listProductType, productName, orderNumber, productNumber, regDateFrom, regDateTo, serial);
            var orderBy = BuildOrderBy(_serialSortCols, sortCol, sortDir, "s.rowid DESC");
            var limitOffset = BuildLimitOffset(pageSize, page);

            return con.Query<SerialRecord>($"""
                SELECT
                    s.rowid AS RowId,
                    s.Serial,
                    s.OLesSerial,
                    p.OrderNumber,
                    p.ProductNumber,
                    s.ProductName,
                    p.ProductType,
                    p.ProductModel,
                    p.RegDate,
                    p.CreatedAt
                FROM T_Serial AS s
                LEFT JOIN V_Product AS p ON s.UsedID = p.ID
                WHERE {where}
                ORDER BY {orderBy}
                {limitOffset}
                """, param).AsList();
        }

        // UseID は V_Product.ID への外部キー（製品登録1件に紐づく使用基板を返す）
        public IReadOnlyList<UsedSubstrateRecord> GetUsedSubstrates(long id) {
            using var con = new SqliteConnection(_connectionString);
            return con.Query<UsedSubstrateRecord>("""
                SELECT
                    ID,
                    SubstrateName,
                    SubstrateModel,
                    SubstrateNumber,
                    Decrease
                FROM V_Substrate
                WHERE UseID = @ID
                ORDER BY SubstrateModel ASC
                """, new { ID = id }).AsList();
        }

        private static (string where, object param) BuildProductWhere(
            string? listCategory, string? listProductName, string? listProductType,
            string? productName, string? orderNumber, string? productNumber, string? regDateFrom, string? regDateTo) {

            var conditions = new List<string> { "v.IsDeleted = 0" };
            if (!string.IsNullOrWhiteSpace(listCategory))
                conditions.Add("p.CategoryName = @ListCategory");
            if (!string.IsNullOrWhiteSpace(listProductName))
                conditions.Add("v.ProductName = @ListProductName");
            if (!string.IsNullOrWhiteSpace(listProductType))
                conditions.Add("v.ProductType = @ListProductType");
            if (!string.IsNullOrWhiteSpace(productName))
                conditions.Add("v.ProductName LIKE '%' || @ProductName || '%'");
            if (!string.IsNullOrWhiteSpace(orderNumber))
                conditions.Add("v.OrderNumber LIKE '%' || @OrderNumber || '%'");
            if (!string.IsNullOrWhiteSpace(productNumber))
                conditions.Add("v.ProductNumber LIKE '%' || @ProductNumber || '%'");
            if (!string.IsNullOrWhiteSpace(regDateFrom))
                conditions.Add("v.RegDate >= @RegDateFrom");
            if (!string.IsNullOrWhiteSpace(regDateTo))
                conditions.Add("v.RegDate <= @RegDateTo");

            return (string.Join(" AND ", conditions), new {
                ListCategory = listCategory,
                ListProductName = listProductName,
                ListProductType = listProductType,
                ProductName = productName,
                OrderNumber = orderNumber,
                ProductNumber = productNumber,
                // SQLite はテキスト型で日付を "yyyy/MM/dd" 形式で保存しているため、
                // HTML date input の "yyyy-MM-dd" 形式から変換する
                RegDateFrom = regDateFrom?.Replace('-', '/'),
                RegDateTo = regDateTo?.Replace('-', '/')
            });
        }

        private static (string where, object param) BuildSerialWhere(
            string? listCategory, string? listProductName, string? listProductType, string? productName, string? orderNumber, string? productNumber, string? regDateFrom, string? regDateTo, string? serial) {

            var conditions = new List<string> { "1=1" };
            if (!string.IsNullOrWhiteSpace(listCategory))
                conditions.Add("p.CategoryName = @ListCategory");
            if (!string.IsNullOrWhiteSpace(listProductName))
                conditions.Add("p.ProductName = @ListProductName");
            if (!string.IsNullOrWhiteSpace(listProductType))
                conditions.Add("p.ProductType = @ListProductType");
            if (!string.IsNullOrWhiteSpace(productName))
                conditions.Add("s.ProductName LIKE '%' || @ProductName || '%'");
            if (!string.IsNullOrWhiteSpace(orderNumber))
                conditions.Add("p.OrderNumber LIKE '%' || @OrderNumber || '%'");
            if (!string.IsNullOrWhiteSpace(productNumber))
                conditions.Add("p.ProductNumber LIKE '%' || @ProductNumber || '%'");
            if (!string.IsNullOrWhiteSpace(regDateFrom))
                conditions.Add("p.RegDate >= @RegDateFrom");
            if (!string.IsNullOrWhiteSpace(regDateTo))
                conditions.Add("p.RegDate <= @RegDateTo");
            if (!string.IsNullOrWhiteSpace(serial))
                conditions.Add("s.Serial LIKE '%' || @Serial || '%'");

            return (string.Join(" AND ", conditions), new {
                ListCategory = listCategory,
                ListProductName = listProductName,
                ListProductType = listProductType,
                ProductName = productName,
                OrderNumber = orderNumber,
                ProductNumber = productNumber,
                RegDateFrom = regDateFrom?.Replace('-', '/'),
                RegDateTo = regDateTo?.Replace('-', '/'),
                Serial = serial
            });
        }

    }
}
