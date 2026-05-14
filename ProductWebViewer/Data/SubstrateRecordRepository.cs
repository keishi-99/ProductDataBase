using Dapper;
using Microsoft.Data.Sqlite;
using ProductWebViewer.Models;

namespace ProductWebViewer.Data {
    public class SubstrateRecordRepository : RepositoryBase {
        // クエリパラメータ名 → 実テーブルの列参照へのホワイトリスト（ORDER BY インジェクション対策）
        private static readonly Dictionary<string, string> _substrateSortCols = new(StringComparer.OrdinalIgnoreCase) {
            ["ID"] = "s.ID",
            ["CategoryName"] = "m.CategoryName",
            ["ProductName"] = "m.ProductName",
            ["SubstrateName"] = "s.SubstrateName",
            ["SubstrateModel"] = "s.SubstrateModel",
            ["OrderNumber"] = "s.OrderNumber",
            ["SubstrateNumber"] = "s.SubstrateNumber",
            ["Increase"] = "s.Increase",
            ["Decrease"] = "s.Decrease",
            ["Defect"] = "s.Defect",
            ["Person"] = "s.Person",
            ["RegDate"] = "s.RegDate",
            ["Comment"] = "s.Comment",
            ["CreatedAt"] = "s.CreatedAt",
            ["UseProductName"]   = "p.ProductName",
            ["UseOrderNumber"]   = "p.OrderNumber",
            ["UseProductNumber"] = "p.ProductNumber",
        };

        // 在庫一覧用ホワイトリスト（同上）
        private static readonly Dictionary<string, string> _stockSortCols = new(StringComparer.OrdinalIgnoreCase) {
            ["CategoryName"] = "m.CategoryName",
            ["ProductName"] = "m.ProductName",
            ["SubstrateName"] = "s.SubstrateName",
            ["SubstrateModel"] = "s.SubstrateModel",
            ["SubstrateNumber"] = "s.SubstrateNumber",
            ["OrderNumber"] = "s.OrderNumber",
            ["Stock"] = "Stock",
        };

        public SubstrateRecordRepository(IConfiguration configuration) : base(configuration) { }

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
            string? listCategory = null,
            string? listProductName = null,
            string? listSubstrateName = null,
            string? substrateName = null,
            string? orderNumber = null,
            string? substrateNumber = null,
            string? dateType = null,
            string? dateFrom = null,
            string? dateTo = null) {

            using var con = new SqliteConnection(_connectionString);
            var (where, param) = BuildSubstrateWhere(listCategory, listProductName, listSubstrateName, substrateName, orderNumber, substrateNumber, dateType, dateFrom, dateTo);
            return con.ExecuteScalar<int>($"""
                SELECT COUNT(*)
                FROM V_Substrate AS s
                LEFT JOIN M_SubstrateDef AS m ON s.SubstrateID = m.SubstrateID
                WHERE {where}
                """, param);
        }

        public IReadOnlyList<SubstrateRecord> GetAll(
            string? listCategory = null,
            string? listProductName = null,
            string? listSubstrateName = null,
            string? substrateName = null,
            string? orderNumber = null,
            string? substrateNumber = null,
            string? dateType = null,
            string? dateFrom = null,
            string? dateTo = null,
            string sortCol = "",
            string sortDir = "desc",
            int page = 1,
            int pageSize = 100) {

            using var con = new SqliteConnection(_connectionString);
            var (where, param) = BuildSubstrateWhere(listCategory, listProductName, listSubstrateName, substrateName, orderNumber, substrateNumber, dateType, dateFrom, dateTo);
            var orderBy = BuildOrderBy(_substrateSortCols, sortCol, sortDir, "s.ID DESC");
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
                    s.Comment,
                    s.CreatedAt,
                    s.UseID,
                    p.ProductName AS UseProductName,
                    p.OrderNumber AS UseOrderNumber,
                    p.ProductNumber AS UseProductNumber
                FROM V_Substrate AS s
                LEFT JOIN M_SubstrateDef AS m ON s.SubstrateID = m.SubstrateID
                LEFT JOIN V_Product AS p ON s.UseID = p.ID
                WHERE {where}
                ORDER BY {orderBy}
                {limitOffset}
                """, param).AsList();
        }

        // groupByModel=true: 型式単位で集計（ロット違いを合算）、false: ロット・注文番号単位で集計
        public int GetStockCount(
            string? listCategory = null,
            string? listProductName = null,
            string? listSubstrateName = null,
            bool groupByModel = true,
            bool excludeZeroStock = false,
            string? substrateName = null,
            string? orderNumber = null,
            string? substrateNumber = null) {

            using var con = new SqliteConnection(_connectionString);
            var (where, param) = BuildStockWhere(listCategory, listProductName, listSubstrateName, substrateName, orderNumber, substrateNumber);
            var groupBy = groupByModel
                ? "s.SubstrateID, s.SubstrateName, s.SubstrateModel"
                : "s.SubstrateID, s.SubstrateName, s.SubstrateModel, s.SubstrateNumber, s.OrderNumber";
            var having = excludeZeroStock
                ? "HAVING SUM(COALESCE(s.Increase,0) + COALESCE(s.Decrease,0) + COALESCE(s.Defect,0)) > 0"
                : "";

            return con.ExecuteScalar<int>($"""
                SELECT COUNT(*) FROM (
                    SELECT 1
                    FROM V_Substrate AS s
                    LEFT JOIN M_SubstrateDef AS m ON s.SubstrateID = m.SubstrateID
                    WHERE {where}
                    GROUP BY {groupBy}
                    {having}
                ) AS t
                """, param);
        }

        public IReadOnlyList<StockRecord> GetStock(
            string? listCategory = null,
            string? listProductName = null,
            string? listSubstrateName = null,
            bool groupByModel = true,
            bool excludeZeroStock = false,
            string? substrateName = null,
            string? orderNumber = null,
            string? substrateNumber = null,
            string sortCol = "",
            string sortDir = "asc",
            int page = 1,
            int pageSize = 100) {

            using var con = new SqliteConnection(_connectionString);
            var (where, param) = BuildStockWhere(listCategory, listProductName, listSubstrateName, substrateName, orderNumber, substrateNumber);
            var limitOffset = BuildLimitOffset(pageSize, page);
            var having = excludeZeroStock
                ? "HAVING SUM(COALESCE(s.Increase,0) + COALESCE(s.Decrease,0) + COALESCE(s.Defect,0)) > 0"
                : "";

            // Decrease・Defect は DB に負数で格納されているため、そのまま合算すると在庫数になる
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
                    {having}
                    ORDER BY {orderBy}
                    {limitOffset}
                    """, param).AsList();
            }
            else {
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
                    {having}
                    ORDER BY {orderBy}
                    {limitOffset}
                    """, param).AsList();
            }
        }

        private static (string where, object param) BuildSubstrateWhere(
            string? listCategory, string? listProductName, string? listSubstrateName,
            string? substrateName, string? orderNumber, string? substrateNumber,
            string? dateType, string? dateFrom, string? dateTo) {

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
            if (!string.IsNullOrWhiteSpace(substrateNumber))
                conditions.Add("s.SubstrateNumber LIKE '%' || @SubstrateNumber || '%'");

            // CreatedAt は "yyyy-MM-dd HH:mm:ss" 形式のためDATE()で日付部分を抽出して比較する
            // RegDate は "yyyy/MM/dd" 形式のため "-" を "/" に変換して比較する
            var useCreatedAt = dateType == "createdAt";
            if (!string.IsNullOrWhiteSpace(dateFrom))
                conditions.Add(useCreatedAt ? "DATE(s.CreatedAt) >= @DateFrom" : "s.RegDate >= @DateFrom");
            if (!string.IsNullOrWhiteSpace(dateTo))
                conditions.Add(useCreatedAt ? "DATE(s.CreatedAt) <= @DateTo" : "s.RegDate <= @DateTo");

            return (string.Join(" AND ", conditions), new {
                ListCategory = listCategory,
                ListProductName = listProductName,
                ListSubstrateName = listSubstrateName,
                SubstrateName = substrateName,
                OrderNumber = orderNumber,
                SubstrateNumber = substrateNumber,
                DateFrom = useCreatedAt ? dateFrom : dateFrom?.Replace('-', '/'),
                DateTo   = useCreatedAt ? dateTo   : dateTo?.Replace('-', '/')
            });
        }

        // 在庫は全期間の累積値のためレジストリ日付フィルターは持たない
        private static (string where, object param) BuildStockWhere(
            string? listCategory, string? listProductName, string? listSubstrateName,
            string? substrateName = null, string? orderNumber = null, string? substrateNumber = null) {

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
            if (!string.IsNullOrWhiteSpace(substrateNumber))
                conditions.Add("s.SubstrateNumber LIKE '%' || @SubstrateNumber || '%'");

            return (string.Join(" AND ", conditions), new {
                ListCategory = listCategory,
                ListProductName = listProductName,
                ListSubstrateName = listSubstrateName,
                SubstrateName = substrateName,
                OrderNumber = orderNumber,
                SubstrateNumber = substrateNumber,
            });
        }

    }
}
