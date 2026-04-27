using Dapper;
using Microsoft.Data.Sqlite;
using ProductWebViewer.Models;

namespace ProductWebViewer.Data {
    public class ProductRecordRepository {
        private readonly string _connectionString;

        public ProductRecordRepository(IConfiguration configuration) {
            var dbPath = configuration["DatabasePath"]
                ?? throw new InvalidOperationException("DatabasePath が appsettings.json に設定されていません。");

            var fullPath = Path.IsPathRooted(dbPath)
                ? dbPath
                : Path.Combine(AppContext.BaseDirectory, dbPath);

            if (!File.Exists(fullPath)) {
                throw new FileNotFoundException($"DB ファイルが見つかりません: {fullPath}");
            }

            _connectionString = new SqliteConnectionStringBuilder {
                DataSource = fullPath,
                Mode = SqliteOpenMode.ReadOnly,
                Pooling = false
            }.ToString();
        }

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

        public IReadOnlyList<SerialRecord> GetSerialHistory(
            string? listProductName = null,
            string? listProductType = null,
            string? serial = null) {

            using var con = new SqliteConnection(_connectionString);

            var conditions = new List<string> { "1=1" };
            if (!string.IsNullOrWhiteSpace(listProductName))
                conditions.Add("s.ProductName = @ListProductName");
            if (!string.IsNullOrWhiteSpace(listProductType))
                conditions.Add("(p.ProductType = @ListProductType OR p.ProductType IS NULL)");
            if (!string.IsNullOrWhiteSpace(serial))
                conditions.Add("s.Serial LIKE '%' || @Serial || '%'");

            var sql = $"""
                SELECT
                    s.rowid AS RowId,
                    s.Serial,
                    s.OLesSerial,
                    p.OrderNumber,
                    p.ProductNumber,
                    s.ProductName,
                    p.ProductType,
                    p.ProductModel,
                    p.RegDate
                FROM T_Serial AS s
                LEFT JOIN V_Product AS p ON s.UsedID = p.ID
                WHERE {string.Join(" AND ", conditions)}
                ORDER BY s.rowid DESC
                LIMIT 500
                """;

            return con.Query<SerialRecord>(sql, new {
                ListProductName = listProductName,
                ListProductType = listProductType,
                Serial = serial
            }).AsList();
        }

        public IReadOnlyList<ProductRecord> GetAll(
            string? listCategory = null,
            string? listProductName = null,
            string? listProductType = null,
            string? productName = null,
            string? orderNumber = null,
            string? regDateFrom = null,
            string? regDateTo = null) {

            using var con = new SqliteConnection(_connectionString);

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
            if (!string.IsNullOrWhiteSpace(regDateFrom))
                conditions.Add("v.RegDate >= @RegDateFrom");
            if (!string.IsNullOrWhiteSpace(regDateTo))
                conditions.Add("v.RegDate <= @RegDateTo");

            var sql = $"""
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
                    v.Comment
                FROM V_Product AS v
                LEFT JOIN M_ProductDef AS p ON v.ProductID = p.ProductID
                WHERE {string.Join(" AND ", conditions)}
                ORDER BY v.ID DESC
                LIMIT 500
                """;

            return con.Query<ProductRecord>(sql, new {
                ListCategory = listCategory,
                ListProductName = listProductName,
                ListProductType = listProductType,
                ProductName = productName,
                OrderNumber = orderNumber,
                RegDateFrom = regDateFrom,
                RegDateTo = regDateTo
            }).AsList();
        }
    }
}
