using System.Security.Cryptography;
using System.Text;
using Dapper;
using Microsoft.Data.Sqlite;

namespace ProductWebViewer.Data {
    // このアプリが依存しているビューが、DB上で想定通りの定義になっているかを検証する。
    // ProductDatabase側は以前、起動時にDROP+CREATEでビューを自動同期していたが、
    // このWebViewerを含む複数プロセスが同じDBに同時アクセスした際の書き込みロック競合を避けるため廃止し、
    // ビューの作成・変更はDB側で直接手動管理する運用にした（テーブルのDDLと同様の方式）。
    // そのため、このWebViewer側もDBの定義に一方的に依存するだけの存在になっており、
    // DB変更に対してこちらのコードが追従し忘れていないかを起動時に検証する。
    internal static class ViewDefinitionVerifier {
        // ProductDatabase/Data/ProductRepository.cs の定義と一致させること
        private static readonly Dictionary<string, string> _expectedViewDefinitions = new() {
            ["V_Product"] = """
                CREATE VIEW V_Product AS
                SELECT
                    t.ID,
                    t.ProductID,
                    m.CategoryName,
                    m.ProductName,
                    m.ProductType,
                    m.ProductModel,
                    t.OrderNumber,
                    t.ProductNumber,
                    t.OLesNumber,
                    t.Quantity,
                    t.SerialFirst,
                    t.SerialLast,
                    t.Revision,
                    t.RevisionGroup,
                    t.SerialLastNumber,
                    COALESCE(t.PersonID || '.' || p.PersonName, '') AS PersonInfo,
                    t.RegDate,
                    t.Comment,
                    t.CreatedAt,
                    t.IsDeleted,
                    t.DeletedAt
                FROM T_Product AS t
                LEFT JOIN M_ProductDef AS m ON t.ProductID = m.ProductID
                LEFT JOIN M_Person AS p ON t.PersonID = p.PersonID
                WHERE t.IsDeleted = 0
                """,
            ["V_Substrate"] = """
                CREATE VIEW V_Substrate AS
                SELECT
                    t.ID,
                    t.SubstrateID,
                    m.CategoryName,
                    m.ProductName,
                    m.SubstrateName,
                    m.SubstrateModel,
                    t.OrderNumber,
                    t.SubstrateNumber,
                    t.Increase,
                    t.Decrease,
                    t.Defect,
                    COALESCE(t.PersonID || '.' || p.PersonName, '') AS PersonInfo,
                    t.RegDate,
                    t.Comment,
                    t.UseID,
                    t.CreatedAt,
                    t.IsDeleted,
                    t.DeletedAt
                FROM T_Substrate AS t
                LEFT JOIN M_SubstrateDef AS m ON t.SubstrateID = m.SubstrateID
                LEFT JOIN M_Person AS p ON t.PersonID = p.PersonID
                WHERE t.IsDeleted = 0
                """,
            ["V_Serial"] = """
                CREATE VIEW V_Serial AS
                SELECT
                    s.rowid,
                    s.Serial,
                    s.OLesSerial,
                    s.UsedID,
                    s.ProductID,
                    m.ProductName,
                    m.CategoryName
                FROM T_Serial AS s
                LEFT JOIN M_ProductDef AS m ON s.ProductID = m.ProductID
                """,
        };

        public static void Verify(string connectionString) {
            using var con = new SqliteConnection(connectionString);
            con.Open();

            var mismatched = new List<string>();
            foreach (var (viewName, expectedSql) in _expectedViewDefinitions) {
                var actualSql = con.QuerySingleOrDefault<string>(
                    "SELECT sql FROM sqlite_master WHERE type = 'view' AND name = @ViewName",
                    new { ViewName = viewName });

                if (actualSql is null || ComputeHash(actualSql) != ComputeHash(expectedSql)) {
                    mismatched.Add(viewName);
                }
            }

            if (mismatched.Count > 0) {
                throw new InvalidOperationException(
                    $"DB上のビュー定義がこのアプリの想定と異なります: {string.Join(", ", mismatched)}\n" +
                    "DBの変更内容を確認し、アプリのコードを対応させてください。");
            }
        }

        private static string ComputeHash(string sql) {
            if (string.IsNullOrEmpty(sql)) { return string.Empty; }

            var normalized = string.Join(' ', sql.ToUpperInvariant().Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
            var bytes = SHA256.HashData(Encoding.UTF8.GetBytes(normalized));
            return Convert.ToHexString(bytes);
        }
    }
}
