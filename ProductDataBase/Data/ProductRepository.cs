using Dapper;
using Microsoft.Data.Sqlite;
using ProductDatabase.Common;
using ProductDatabase.Models;
using System.Data;

namespace ProductDatabase.Data {
    public class ProductRepository {

        public DataTable ProductDataTable { get; } = new();
        public DataTable SubstrateDataTable { get; } = new();
        public DataTable ProductUseSubstrate { get; } = new();

        internal static readonly CacheManager<(DataTable, DataTable, DataTable)> _cacheManager
            = new(TimeSpan.FromMinutes(5));
        private static readonly object _loadLock = new();

        // DBファイルのパスを検証しSQLite接続文字列を返す（DbConnectionHelper に委譲・プーリング無効）
        public static string GetConnectionRegistration() {
            return DbConnectionHelper.GetConnectionString();
        }

        // 製品・基板・使用基板の全マスターデータをDBから読み込む（TTL キャッシング機構付き）
        public void LoadAll() {
            lock (_loadLock) {
                // キャッシュが有効な場合はキャッシュからデータを復元して終了
                if (_cacheManager.TryGetCachedData(out var cached)) {
                    RestoreFromCache(cached);
                    return;
                }

                // キャッシュが無効な場合は DB から読み込む
                ProductDataTable.Clear();
                SubstrateDataTable.Clear();
                ProductUseSubstrate.Clear();

                using var con = new SqliteConnection(GetConnectionRegistration());
                con.Open();

                using (var reader = con.ExecuteReader($"SELECT * FROM {Constants.ProductTableName};")) {
                    ProductDataTable.Load(reader);
                }

                using (var reader = con.ExecuteReader($"SELECT * FROM {Constants.SubstrateTableName};")) {
                    SubstrateDataTable.Load(reader);
                }

                using (var reader = con.ExecuteReader($"SELECT * FROM {Constants.VProductUseSubstrate};")) {
                    ProductUseSubstrate.Load(reader);
                }

                // 読み込んだデータをキャッシュに保存
                _cacheManager.SetCache((ProductDataTable.Copy(), SubstrateDataTable.Copy(), ProductUseSubstrate.Copy()));
            }
        }

        // キャッシュからデータを現在の DataTable に復元する（UI 制約無効化でパフォーマンス向上）
        private void RestoreFromCache((DataTable, DataTable, DataTable) cached) {
            try {
                ProductDataTable.BeginLoadData();
                SubstrateDataTable.BeginLoadData();
                ProductUseSubstrate.BeginLoadData();

                ProductDataTable.Clear();
                SubstrateDataTable.Clear();
                ProductUseSubstrate.Clear();

                ProductDataTable.Merge(cached.Item1);
                SubstrateDataTable.Merge(cached.Item2);
                ProductUseSubstrate.Merge(cached.Item3);
            } finally {
                ProductDataTable.EndLoadData();
                SubstrateDataTable.EndLoadData();
                ProductUseSubstrate.EndLoadData();
            }
        }

        // 基板IDに一致する行をSubstrateDataTableから取得する
        public DataRow GetSubstrateById(long id) {
            return SubstrateDataTable
                .AsEnumerable()
                .SingleOrDefault(r => r.Field<long>("SubstrateID") == id)
                ?? throw new InvalidOperationException($"基板ID [{id}] が見つかりません。");
        }

        // 製品IDに一致する行をProductDataTableから取得する
        public DataRow GetProductById(long id) {
            return ProductDataTable
                .AsEnumerable()
                .SingleOrDefault(r => r.Field<long>("ProductID") == id)
                ?? throw new InvalidOperationException($"製品ID [{id}] が見つかりません。");
        }

        // 製品IDから使用基板リストを取得（M_SubstrateDef の ExclusiveGroupID を含む）
        public static List<SubstrateInfo> GetUseSubstrates(long productKey) {
            using var con = new SqliteConnection(GetConnectionRegistration());

            var sql =
                $"""
                SELECT
                    pus.SubstrateID,
                    s.SubstrateName,
                    s.SubstrateModel,
                    s.ExclusiveGroupID
                FROM {Constants.ProductUseSubstrateTableName} pus
                JOIN {Constants.SubstrateTableName} s ON pus.SubstrateID = s.SubstrateID
                WHERE pus.ProductID = @ProductKey
                """;

            return [.. con.Query<SubstrateInfo>(sql, new { ProductKey = productKey })];
        }

        // キャッシュ済みの全DataTableをクリアする
        public void Clear() {
            ProductDataTable.Clear();
            SubstrateDataTable.Clear();
            ProductUseSubstrate.Clear();
            _cacheManager.ClearCache();
        }

        // 製品マスターを新規登録し採番されたProductIDを返す
        public static long InsertProduct(ProductMaster product) {
            using var con = new SqliteConnection(GetConnectionRegistration());

            var sql =
                $"""
                INSERT INTO {Constants.ProductTableName}
                    (CategoryName, ProductName, ProductType, ProductModel, Initial,
                     OLesInitial, OLesSerialSuffix, RevisionGroup, RegType, SerialType, Checkbox,
                     SerialPrintType, SheetPrintType, Visible)
                VALUES
                    (@CategoryName, @ProductName, @ProductType, @ProductModel, @Initial,
                     @OLesInitial, @OLesSerialSuffix, @RevisionGroup, @RegType, @SerialType, @Checkbox,
                     @SerialPrintType, @SheetPrintType, @Visible);
                SELECT last_insert_rowid();
                """;

            var checkBin = Convert.ToString(product.CheckBin, 2).PadLeft(11, '0');

            var result = con.ExecuteScalar<long>(sql, new {
                product.CategoryName,
                product.ProductName,
                product.ProductType,
                product.ProductModel,
                product.Initial,
                product.OLesInitial,
                product.OLesSerialSuffix,
                product.RevisionGroup,
                product.RegType,
                SerialType = product.SerialDigitType,
                Checkbox = checkBin,
                product.SerialPrintType,
                product.SheetPrintType,
                Visible = product.Visible ? 1 : 0
            });

            // キャッシュをクリア（マスターデータが変更されたため）
            _cacheManager.ClearCache();

            return result;
        }

        // 製品マスターを更新する
        public static void UpdateProduct(ProductMaster product) {
            using var con = new SqliteConnection(GetConnectionRegistration());

            var sql =
                $"""
                UPDATE {Constants.ProductTableName} SET
                    CategoryName    = @CategoryName,
                    ProductName     = @ProductName,
                    ProductType     = @ProductType,
                    ProductModel    = @ProductModel,
                    Initial         = @Initial,
                    OLesInitial         = @OLesInitial,
                    OLesSerialSuffix    = @OLesSerialSuffix,
                    RevisionGroup       = @RevisionGroup,
                    RegType         = @RegType,
                    SerialType      = @SerialType,
                    Checkbox        = @Checkbox,
                    SerialPrintType = @SerialPrintType,
                    SheetPrintType  = @SheetPrintType,
                    Visible         = @Visible
                WHERE ProductID = @ProductID
                """;

            var checkBin = Convert.ToString(product.CheckBin, 2).PadLeft(11, '0');

            con.Execute(sql, new {
                product.CategoryName,
                product.ProductName,
                product.ProductType,
                product.ProductModel,
                product.Initial,
                product.OLesInitial,
                product.OLesSerialSuffix,
                product.RevisionGroup,
                product.RegType,
                SerialType = product.SerialDigitType,
                Checkbox = checkBin,
                product.SerialPrintType,
                product.SheetPrintType,
                Visible = product.Visible ? 1 : 0,
                product.ProductID
            });

            // キャッシュをクリア（マスターデータが変更されたため）
            _cacheManager.ClearCache();
        }

        // このコードが前提とするビュー定義。
        // 以前は起動時にDROP+CREATEで自動同期していたが、ProductWebViewer等の別プロセスが
        // 同じDBに同時アクセスした際に書き込みロックで競合しうるため廃止した。
        // ビューの作成・変更は（テーブルのDDLと同様）DB側で直接手動で行う運用とし、
        // ここではDB上の実際の定義とのズレをVerifyViewDefinitions()で検知するに留める。
        private static readonly Dictionary<string, string> _expectedViewDefinitions = new() {
        private static readonly Dictionary<string, string> _expectedViewDefinitions = new() {
            [Constants.VSerialTableName] = $"""
                CREATE VIEW {Constants.VSerialTableName} AS
                SELECT
                    s.rowid,
                    s.Serial,
                    s.OLesSerial,
                    s.UsedID,
                    s.ProductID,
                    m.ProductName,
                    m.CategoryName
                FROM {Constants.TSerialTableName} AS s
                LEFT JOIN {Constants.ProductTableName} AS m ON s.ProductID = m.ProductID
                """,
            [Constants.VProductTableName] = $"""
                CREATE VIEW {Constants.VProductTableName} AS
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
                FROM {Constants.TProductTableName} AS t
                LEFT JOIN {Constants.ProductTableName} AS m ON t.ProductID = m.ProductID
                LEFT JOIN {Constants.PersonTableName} AS p ON t.PersonID = p.PersonID
                WHERE t.IsDeleted = 0
                """,
            [Constants.VSubstrateTableName] = $"""
                CREATE VIEW {Constants.VSubstrateTableName} AS
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
                FROM {Constants.TSubstrateTableName} AS t
                LEFT JOIN {Constants.SubstrateTableName} AS m ON t.SubstrateID = m.SubstrateID
                LEFT JOIN {Constants.PersonTableName} AS p ON t.PersonID = p.PersonID
                WHERE t.IsDeleted = 0
                """,
            [Constants.VProductUseSubstrate] = $"""
                CREATE VIEW {Constants.VProductUseSubstrate} AS
                SELECT
                    p.ProductID AS P_ProductID,
                    s.SubstrateID AS S_SubstrateID,
                    p.ProductName,
                    p.ProductType,
                    s.SubstrateName,
                    s.SubstrateModel
                FROM {Constants.ProductUseSubstrateTableName} AS ps
                JOIN {Constants.ProductTableName} AS p ON p.ProductID = ps.ProductID
                JOIN {Constants.SubstrateTableName} AS s ON s.SubstrateID = ps.SubstrateID
                """,
            [Constants.VRePrintTableName] = $"""
                CREATE VIEW {Constants.VRePrintTableName} AS
                SELECT
                    t.ID,
                    t.SerialPrintType,
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
                    COALESCE(t.PersonID || '.' || p.PersonName, '') AS PersonInfo,
                    t.RegDate,
                    t.Comment,
                    t.CreatedAt
                FROM {Constants.TRePrintTableName} AS t
                LEFT JOIN {Constants.ProductTableName} AS m ON t.ProductID = m.ProductID
                LEFT JOIN {Constants.PersonTableName} AS p ON t.PersonID = p.PersonID
                """,
        };

        // DB上の実際のビュー定義が、コードが前提とする定義と一致するか検証する（不一致・不在なら例外を投げて起動を止める）
        public static void VerifyViewDefinitions() {
            using var con = new SqliteConnection(GetConnectionRegistration());
            con.Open();

            var mismatched = new List<string>();
            foreach (var (viewName, expectedSql) in _expectedViewDefinitions) {
                var actualSql = con.QuerySingleOrDefault<string>(
                    "SELECT sql FROM sqlite_master WHERE type = 'view' AND name = @ViewName",
                    new { ViewName = viewName });

                if (actualSql is null || ViewDefinitionHash.Compute(actualSql) != ViewDefinitionHash.Compute(expectedSql)) {
                    mismatched.Add(viewName);
                }
            }

            if (mismatched.Count > 0) {
                throw new InvalidOperationException(
                    $"DB上のビュー定義がアプリの想定と異なります: {string.Join(", ", mismatched)}\n" +
                    "DBの変更内容を確認し、アプリのコードを対応させてください。");
            }
        }

        // 製品マスターを物理削除する（実績存在チェック・関連紐づけ削除を含む）
        public static void DeleteProduct(long productId) {
            using var con = new SqliteConnection(GetConnectionRegistration());
            con.Open();
            using var tx = con.BeginTransaction();

            var count = con.ExecuteScalar<int>(
                $"SELECT COUNT(*) FROM {Constants.TProductTableName} WHERE ProductID = @ProductId",
                new { ProductId = productId }, tx);

            if (count > 0) {
                throw new InvalidOperationException("この製品には製品登録実績があるため削除できません。");
            }

            con.Execute(
                $"DELETE FROM {Constants.ProductUseSubstrateTableName} WHERE ProductID = @ProductId",
                new { ProductId = productId }, tx);

            con.Execute(
                $"DELETE FROM {Constants.ProductTableName} WHERE ProductID = @ProductId",
                new { ProductId = productId }, tx);

            tx.Commit();

            // キャッシュをクリア（マスターデータが変更されたため）
            _cacheManager.ClearCache();
        }

        // 指定ProductIDの製品実績が存在するか確認する
        public static bool ExistsProductResult(long productId) {
            using var con = new SqliteConnection(GetConnectionRegistration());
            return con.ExecuteScalar<int>(
                $"SELECT COUNT(*) FROM {Constants.TProductTableName} WHERE ProductID = @ProductId",
                new { ProductId = productId }) > 0;
        }

        // 品目番号でSQLiteを検索し一致する基板・製品情報のリストを返す
        public static List<ProductSearchItem> SearchByModel(string productModel) {
            using var con = new SqliteConnection(GetConnectionRegistration());
            var sql =
                $"""
                SELECT
                    s.SubItemNumber,
                    s.SubstrateName,
                    s.ProductName AS sName,
                    p.ProductName AS pName,
                    p.ProductType,
                    p.ProItemNumber
                FROM
                    {Constants.ProductTableName} AS p
                FULL JOIN
                    {Constants.SubstrateTableName} AS s
                ON
                    s.SubItemNumber = p.ProItemNumber
                WHERE
                    s.SubItemNumber LIKE '%' || @ProductModel || '%'
                OR
                    p.ProItemNumber LIKE '%' || @ProductModel || '%'
                """;

            var rows = con.Query(sql, new { ProductModel = productModel }).ToList();
            if (rows.Count == 0) {
                throw new Exception($"品目番号が見つかりません。\n品目番号:[{productModel}]");
            }

            var items = new List<ProductSearchItem>();
            foreach (var row in rows) {
                var colSub = row.SubItemNumber?.ToString() ?? string.Empty;
                var colPro = row.ProItemNumber?.ToString() ?? string.Empty;

                if (!string.IsNullOrWhiteSpace(colSub)) {
                    items.Add(new ProductSearchItem(
                        colSub,
                        row.sName?.ToString() ?? string.Empty,
                        string.Empty,
                        row.SubstrateName?.ToString() ?? string.Empty,
                        "1"));
                }
                if (!string.IsNullOrWhiteSpace(colPro)) {
                    items.Add(new ProductSearchItem(
                        colPro,
                        row.pName?.ToString() ?? string.Empty,
                        row.ProductType?.ToString() ?? string.Empty,
                        string.Empty,
                        "2"));
                }
            }
            return items;
        }

        // 製品-基板紐づけを全削除してから指定リストで再登録する
        public static void UpdateProductUseSubstrates(long productId, IEnumerable<long> substrateIds) {
            using var con = new SqliteConnection(GetConnectionRegistration());
            con.Open();
            using var tx = con.BeginTransaction();

            con.Execute(
                $"DELETE FROM {Constants.ProductUseSubstrateTableName} WHERE ProductID = @ProductId",
                new { ProductId = productId }, tx);

            var insertSql = $"INSERT INTO {Constants.ProductUseSubstrateTableName} (ProductID, SubstrateID) VALUES (@ProductId, @SubstrateId)";
            var insertData = substrateIds.Select(id => new { ProductId = productId, SubstrateId = id });
            con.Execute(insertSql, insertData, tx);

            tx.Commit();

            // キャッシュをクリア（マスターデータの関連が変更されたため）
            _cacheManager.ClearCache();
        }

        // キャッシュの状態を取得する（デバッグ・テスト用）
        public (bool IsValid, DateTime? LastLoadTime, TimeSpan Ttl) GetCacheStatus() {
            return _cacheManager.GetCacheStatus();
        }
    }

    // SearchByModel の検索結果1件を表すレコード
    public sealed record ProductSearchItem(
        string ItemNumber,
        string ProductName,
        string ProductType,
        string SubstrateName,
        string Type);
}
