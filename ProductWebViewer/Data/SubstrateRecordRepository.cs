using Dapper;
using Microsoft.Data.Sqlite;
using ProductWebViewer.Models;

namespace ProductWebViewer.Data {
    public class SubstrateRecordRepository {
        private readonly string _connectionString;

        public SubstrateRecordRepository(IConfiguration configuration) {
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
                "SELECT DISTINCT CategoryName FROM M_SubstrateDef WHERE Visible=1 ORDER BY CategoryName"
            ).AsList();
        }

        public IReadOnlyList<string> GetProductNameList(string? category) {
            using var con = new SqliteConnection(_connectionString);
            var sql = string.IsNullOrWhiteSpace(category)
                ? "SELECT DISTINCT ProductName FROM M_SubstrateDef WHERE Visible=1 ORDER BY ProductName"
                : "SELECT DISTINCT ProductName FROM M_SubstrateDef WHERE Visible=1 AND CategoryName=@Category ORDER BY ProductName";
            return con.Query<string>(sql, new { Category = category }).AsList();
        }

        public IReadOnlyList<string> GetSubstrateNameList(string? category, string? productName) {
            using var con = new SqliteConnection(_connectionString);
            var conditions = new List<string> { "Visible=1" };
            if (!string.IsNullOrWhiteSpace(category))
                conditions.Add("CategoryName=@Category");
            if (!string.IsNullOrWhiteSpace(productName))
                conditions.Add("ProductName=@ProductName");
            var sql = $"SELECT DISTINCT SubstrateName FROM M_SubstrateDef WHERE {string.Join(" AND ", conditions)} ORDER BY SubstrateName";
            return con.Query<string>(sql, new { Category = category, ProductName = productName }).AsList();
        }

        public IReadOnlyList<StockRecord> GetStock(
            string? listCategory = null,
            string? listProductName = null,
            string? listSubstrateName = null,
            bool groupByModel = true) {

            using var con = new SqliteConnection(_connectionString);

            var conditions = new List<string> { "s.IsDeleted = 0" };
            if (!string.IsNullOrWhiteSpace(listCategory))
                conditions.Add("m.CategoryName = @ListCategory");
            if (!string.IsNullOrWhiteSpace(listProductName))
                conditions.Add("m.ProductName = @ListProductName");
            if (!string.IsNullOrWhiteSpace(listSubstrateName))
                conditions.Add("s.SubstrateName = @ListSubstrateName");

            var where = string.Join(" AND ", conditions);

            var sql = groupByModel
                ? $"""
                    SELECT
                        s.SubstrateID,
                        m.CategoryName,
                        m.ProductName,
                        s.SubstrateName,
                        s.SubstrateModel,
                        SUM(COALESCE(s.Increase,0) + COALESCE(s.Decrease,0) + COALESCE(s.Defect,0)) AS Stock
                    FROM V_Substrate AS s
                    LEFT JOIN M_SubstrateDef AS m ON s.SubstrateID = m.SubstrateID
                    WHERE {where}
                    GROUP BY s.SubstrateID, s.SubstrateName, s.SubstrateModel
                    ORDER BY s.SubstrateModel
                    """
                : $"""
                    SELECT
                        s.SubstrateID,
                        m.CategoryName,
                        m.ProductName,
                        s.SubstrateName,
                        s.SubstrateModel,
                        s.SubstrateNumber,
                        s.OrderNumber,
                        SUM(COALESCE(s.Increase,0) + COALESCE(s.Decrease,0) + COALESCE(s.Defect,0)) AS Stock
                    FROM V_Substrate AS s
                    LEFT JOIN M_SubstrateDef AS m ON s.SubstrateID = m.SubstrateID
                    WHERE {where}
                    GROUP BY s.SubstrateID, s.SubstrateName, s.SubstrateModel, s.SubstrateNumber, s.OrderNumber
                    ORDER BY MIN(s.ID) DESC
                    LIMIT 500
                    """;

            return con.Query<StockRecord>(sql, new {
                ListCategory = listCategory,
                ListProductName = listProductName,
                ListSubstrateName = listSubstrateName
            }).AsList();
        }

        public IReadOnlyList<SubstrateRecord> GetAll(
            string? listCategory = null,
            string? listProductName = null,
            string? listSubstrateName = null,
            string? substrateName = null,
            string? orderNumber = null,
            string? regDateFrom = null,
            string? regDateTo = null) {

            using var con = new SqliteConnection(_connectionString);

            var conditions = new List<string> { "s.IsDeleted = 0" };
            if (!string.IsNullOrWhiteSpace(listCategory))
                conditions.Add("m.CategoryName = @ListCategory");
            if (!string.IsNullOrWhiteSpace(listProductName))
                conditions.Add("m.ProductName = @ListProductName");
            if (!string.IsNullOrWhiteSpace(listSubstrateName))
                conditions.Add("s.SubstrateName = @ListSubstrateName");
            if (!string.IsNullOrWhiteSpace(substrateName))
                conditions.Add("s.SubstrateName LIKE '%' || @SubstrateName || '%'");
            if (!string.IsNullOrWhiteSpace(orderNumber))
                conditions.Add("s.OrderNumber LIKE '%' || @OrderNumber || '%'");
            if (!string.IsNullOrWhiteSpace(regDateFrom))
                conditions.Add("s.RegDate >= @RegDateFrom");
            if (!string.IsNullOrWhiteSpace(regDateTo))
                conditions.Add("s.RegDate <= @RegDateTo");

            var sql = $"""
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
                    s.Person,
                    s.RegDate,
                    s.Comment
                FROM V_Substrate AS s
                LEFT JOIN M_SubstrateDef AS m ON s.SubstrateID = m.SubstrateID
                WHERE {string.Join(" AND ", conditions)}
                ORDER BY s.ID DESC
                LIMIT 500
                """;

            return con.Query<SubstrateRecord>(sql, new {
                ListCategory = listCategory,
                ListProductName = listProductName,
                ListSubstrateName = listSubstrateName,
                SubstrateName = substrateName,
                OrderNumber = orderNumber,
                RegDateFrom = regDateFrom,
                RegDateTo = regDateTo
            }).AsList();
        }
    }
}
