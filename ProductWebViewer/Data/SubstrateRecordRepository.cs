using Dapper;
using Microsoft.Data.Sqlite;
using ProductWebViewer.Models;

namespace ProductWebViewer.Data {
    public class SubstrateRecordRepository {
        private readonly string _connectionString;

        private static readonly Dictionary<string, string> _substrateSortCols = new(StringComparer.OrdinalIgnoreCase) {
            ["ID"]             = "s.ID",
            ["CategoryName"]   = "m.CategoryName",
            ["ProductName"]    = "m.ProductName",
            ["SubstrateName"]  = "s.SubstrateName",
            ["SubstrateModel"] = "s.SubstrateModel",
            ["OrderNumber"]    = "s.OrderNumber",
            ["SubstrateNumber"]= "s.SubstrateNumber",
            ["Increase"]       = "s.Increase",
            ["Decrease"]       = "s.Decrease",
            ["Defect"]         = "s.Defect",
            ["Person"]         = "s.Person",
            ["RegDate"]        = "s.RegDate",
            ["Comment"]        = "s.Comment",
        };

        private static readonly Dictionary<string, string> _stockSortCols = new(StringComparer.OrdinalIgnoreCase) {
            ["CategoryName"]   = "m.CategoryName",
            ["ProductName"]    = "m.ProductName",
            ["SubstrateName"]  = "s.SubstrateName",
            ["SubstrateModel"] = "s.SubstrateModel",
            ["SubstrateNumber"]= "s.SubstrateNumber",
            ["OrderNumber"]    = "s.OrderNumber",
            ["Stock"]          = "Stock",
        };

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

        public int GetCount(
            string? listCategory      = null,
            string? listProductName   = null,
            string? listSubstrateName = null,
            string? substrateName     = null,
            string? orderNumber       = null,
            string? regDateFrom       = null,
            string? regDateTo         = null) {

            using var con = new SqliteConnection(_connectionString);
            var (where, param) = BuildSubstrateWhere(listCategory, listProductName, listSubstrateName, substrateName, orderNumber, regDateFrom, regDateTo);
            return con.ExecuteScalar<int>($"""
                SELECT COUNT(*)
                FROM V_Substrate AS s
                LEFT JOIN M_SubstrateDef AS m ON s.SubstrateID = m.SubstrateID
                WHERE {where}
                """, param);
        }

        public IReadOnlyList<SubstrateRecord> GetAll(
            string? listCategory      = null,
            string? listProductName   = null,
            string? listSubstrateName = null,
            string? substrateName     = null,
            string? orderNumber       = null,
            string? regDateFrom       = null,
            string? regDateTo         = null,
            string  sortCol  = "",
            string  sortDir  = "desc",
            int     page     = 1,
            int     pageSize = 100) {

            using var con = new SqliteConnection(_connectionString);
            var (where, param) = BuildSubstrateWhere(listCategory, listProductName, listSubstrateName, substrateName, orderNumber, regDateFrom, regDateTo);
            var orderBy     = BuildOrderBy(_substrateSortCols, sortCol, sortDir, "s.ID DESC");
            var limitOffset = BuildLimitOffset(pageSize, page);

            return con.Query<SubstrateRecord>($"""
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
                WHERE {where}
                ORDER BY {orderBy}
                {limitOffset}
                """, param).AsList();
        }

        public int GetStockCount(
            string? listCategory      = null,
            string? listProductName   = null,
            string? listSubstrateName = null,
            bool    groupByModel      = true) {

            using var con = new SqliteConnection(_connectionString);
            var (where, param) = BuildStockWhere(listCategory, listProductName, listSubstrateName);
            var groupBy = groupByModel
                ? "s.SubstrateID, s.SubstrateName, s.SubstrateModel"
                : "s.SubstrateID, s.SubstrateName, s.SubstrateModel, s.SubstrateNumber, s.OrderNumber";

            return con.ExecuteScalar<int>($"""
                SELECT COUNT(*) FROM (
                    SELECT 1
                    FROM V_Substrate AS s
                    LEFT JOIN M_SubstrateDef AS m ON s.SubstrateID = m.SubstrateID
                    WHERE {where}
                    GROUP BY {groupBy}
                ) AS t
                """, param);
        }

        public IReadOnlyList<StockRecord> GetStock(
            string? listCategory      = null,
            string? listProductName   = null,
            string? listSubstrateName = null,
            bool    groupByModel      = true,
            string  sortCol  = "",
            string  sortDir  = "asc",
            int     page     = 1,
            int     pageSize = 100) {

            using var con = new SqliteConnection(_connectionString);
            var (where, param) = BuildStockWhere(listCategory, listProductName, listSubstrateName);
            var limitOffset = BuildLimitOffset(pageSize, page);

            if (groupByModel) {
                var orderBy = BuildOrderBy(_stockSortCols, sortCol, sortDir, "s.SubstrateModel ASC");
                return con.Query<StockRecord>($"""
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
                    ORDER BY {orderBy}
                    {limitOffset}
                    """, param).AsList();
            } else {
                var orderBy = BuildOrderBy(_stockSortCols, sortCol, sortDir, "MIN(s.ID) DESC");
                return con.Query<StockRecord>($"""
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
                    ORDER BY {orderBy}
                    {limitOffset}
                    """, param).AsList();
            }
        }

        private static (string where, object param) BuildSubstrateWhere(
            string? listCategory, string? listProductName, string? listSubstrateName,
            string? substrateName, string? orderNumber, string? regDateFrom, string? regDateTo) {

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
                conditions.Add("REPLACE(s.RegDate, '/', '-') >= @RegDateFrom");
            if (!string.IsNullOrWhiteSpace(regDateTo))
                conditions.Add("REPLACE(s.RegDate, '/', '-') <= @RegDateTo");

            return (string.Join(" AND ", conditions), new {
                ListCategory      = listCategory,
                ListProductName   = listProductName,
                ListSubstrateName = listSubstrateName,
                SubstrateName     = substrateName,
                OrderNumber       = orderNumber,
                RegDateFrom       = regDateFrom,
                RegDateTo         = regDateTo
            });
        }

        private static (string where, object param) BuildStockWhere(
            string? listCategory, string? listProductName, string? listSubstrateName) {

            var conditions = new List<string> { "s.IsDeleted = 0" };
            if (!string.IsNullOrWhiteSpace(listCategory))
                conditions.Add("m.CategoryName = @ListCategory");
            if (!string.IsNullOrWhiteSpace(listProductName))
                conditions.Add("m.ProductName = @ListProductName");
            if (!string.IsNullOrWhiteSpace(listSubstrateName))
                conditions.Add("s.SubstrateName = @ListSubstrateName");

            return (string.Join(" AND ", conditions), new {
                ListCategory      = listCategory,
                ListProductName   = listProductName,
                ListSubstrateName = listSubstrateName
            });
        }

        private static string BuildOrderBy(Dictionary<string, string> cols, string sortCol, string sortDir, string defaultOrder) =>
            cols.TryGetValue(sortCol ?? "", out var col)
                ? $"{col} {(sortDir == "asc" ? "ASC" : "DESC")}"
                : defaultOrder;

        private static string BuildLimitOffset(int pageSize, int page) =>
            pageSize > 0 ? $"LIMIT {pageSize} OFFSET {(page - 1) * pageSize}" : "";
    }
}
