using Dapper;
using Microsoft.Data.Sqlite;
using ProductDatabase.Models;
using ProductDatabase.Other;
using System.Data;

namespace ProductDatabase.Data {
    public class ProductRepository {

        public DataTable ProductDataTable { get; } = new();
        public DataTable SubstrateDataTable { get; } = new();
        public DataTable ProductUseSubstrate { get; } = new();

        // DBファイルのパスを検証しSQLite接続文字列を返す
        public static string GetConnectionRegistration() {
            var productRegistryPath = Path.Combine(AppContext.BaseDirectory, "db", "ProductRegistry.db");
            return !File.Exists(productRegistryPath)
                ? throw new FileNotFoundException("ファイルが見つかりません。", productRegistryPath)
                : new SqliteConnectionStringBuilder() { DataSource = productRegistryPath, Pooling = false }.ToString();
        }

        // 製品・基板・使用基板の全マスターデータをDBから読み込む
        public void LoadAll() {

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

        // 製品IDから使用基板リストを取得
        public static List<SubstrateInfo> GetUseSubstrates(long productKey) {
            using var con = new SqliteConnection(GetConnectionRegistration());

            var sql =
                $"""
                SELECT
                    S_SubstrateID as SubstrateID,
                    SubstrateName,
                    SubstrateModel
                FROM {Constants.VProductUseSubstrate}
                WHERE P_ProductID = @ProductKey
                """
                ;

            return [.. con.Query<SubstrateInfo>(sql, new { ProductKey = productKey })];
        }

        // キャッシュ済みの全DataTableをクリアする
        public void Clear() {
            ProductDataTable.Clear();
            SubstrateDataTable.Clear();
            ProductUseSubstrate.Clear();
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

            return con.ExecuteScalar<long>(sql, new {
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
                "DELETE FROM M_ProductUseSubstrate WHERE ProductID = @ProductId",
                new { ProductId = productId }, tx);

            con.Execute(
                $"DELETE FROM {Constants.ProductTableName} WHERE ProductID = @ProductId",
                new { ProductId = productId }, tx);

            tx.Commit();
        }

        // 指定ProductIDの製品実績が存在するか確認する
        public static bool ExistsProductResult(long productId) {
            using var con = new SqliteConnection(GetConnectionRegistration());
            return con.ExecuteScalar<int>(
                $"SELECT COUNT(*) FROM {Constants.TProductTableName} WHERE ProductID = @ProductId",
                new { ProductId = productId }) > 0;
        }

        // 基板マスターを新規登録し採番されたSubstrateIDを返す
        public static long InsertSubstrate(SubstrateMaster substrate) {
            using var con = new SqliteConnection(GetConnectionRegistration());

            var sql =
                $"""
                INSERT INTO {Constants.SubstrateTableName}
                    (CategoryName, ProductName, SubstrateName, SubstrateModel,
                     RegType, Checkbox, SerialPrintType, Visible)
                VALUES
                    (@CategoryName, @ProductName, @SubstrateName, @SubstrateModel,
                     @RegType, @Checkbox, @SerialPrintType, @Visible);
                SELECT last_insert_rowid();
                """;

            var checkBin = Convert.ToString(substrate.CheckBin, 2).PadLeft(11, '0');

            return con.ExecuteScalar<long>(sql, new {
                substrate.CategoryName,
                substrate.ProductName,
                substrate.SubstrateName,
                substrate.SubstrateModel,
                substrate.RegType,
                Checkbox = checkBin,
                substrate.SerialPrintType,
                Visible = substrate.Visible ? 1 : 0
            });
        }

        // 基板マスターを更新する
        public static void UpdateSubstrate(SubstrateMaster substrate) {
            using var con = new SqliteConnection(GetConnectionRegistration());

            var sql =
                $"""
                UPDATE {Constants.SubstrateTableName} SET
                    CategoryName    = @CategoryName,
                    ProductName     = @ProductName,
                    SubstrateName   = @SubstrateName,
                    SubstrateModel  = @SubstrateModel,
                    RegType         = @RegType,
                    Checkbox        = @Checkbox,
                    SerialPrintType = @SerialPrintType,
                    Visible         = @Visible
                WHERE SubstrateID = @SubstrateID
                """;

            var checkBin = Convert.ToString(substrate.CheckBin, 2).PadLeft(11, '0');

            con.Execute(sql, new {
                substrate.CategoryName,
                substrate.ProductName,
                substrate.SubstrateName,
                substrate.SubstrateModel,
                substrate.RegType,
                Checkbox = checkBin,
                substrate.SerialPrintType,
                Visible = substrate.Visible ? 1 : 0,
                substrate.SubstrateID
            });
        }

        // 基板マスターを物理削除する（実績存在チェック・関連紐づけ削除を含む）
        public static void DeleteSubstrate(long substrateId) {
            using var con = new SqliteConnection(GetConnectionRegistration());
            con.Open();
            using var tx = con.BeginTransaction();

            var count = con.ExecuteScalar<int>(
                $"SELECT COUNT(*) FROM {Constants.TSubstrateTableName} WHERE SubstrateID = @SubstrateId",
                new { SubstrateId = substrateId }, tx);

            if (count > 0) {
                throw new InvalidOperationException("この基板には基板登録実績があるため削除できません。");
            }

            con.Execute(
                "DELETE FROM M_ProductUseSubstrate WHERE SubstrateID = @SubstrateId",
                new { SubstrateId = substrateId }, tx);

            con.Execute(
                $"DELETE FROM {Constants.SubstrateTableName} WHERE SubstrateID = @SubstrateId",
                new { SubstrateId = substrateId }, tx);

            tx.Commit();
        }

        // 指定SubstrateIDの基板実績が存在するか確認する
        public static bool ExistsSubstrateResult(long substrateId) {
            using var con = new SqliteConnection(GetConnectionRegistration());
            return con.ExecuteScalar<int>(
                $"SELECT COUNT(*) FROM {Constants.TSubstrateTableName} WHERE SubstrateID = @SubstrateId",
                new { SubstrateId = substrateId }) > 0;
        }

        // SubstrateModelの重複を確認する（excludeIdは編集時に自身を除外するために使用）
        public static bool ExistsSubstrateModel(string substrateModel, long excludeId = 0) {
            using var con = new SqliteConnection(GetConnectionRegistration());
            return con.ExecuteScalar<int>(
                $"SELECT COUNT(*) FROM {Constants.SubstrateTableName} WHERE SubstrateModel = @SubstrateModel AND SubstrateID != @ExcludeId",
                new { SubstrateModel = substrateModel, ExcludeId = excludeId }) > 0;
        }

        // 製品-基板紐づけを全削除してから指定リストで再登録する
        public static void UpdateProductUseSubstrates(long productId, IEnumerable<long> substrateIds) {
            using var con = new SqliteConnection(GetConnectionRegistration());
            con.Open();
            using var tx = con.BeginTransaction();

            con.Execute(
                "DELETE FROM M_ProductUseSubstrate WHERE ProductID = @ProductId",
                new { ProductId = productId }, tx);

            var insertSql = "INSERT INTO M_ProductUseSubstrate (ProductID, SubstrateID) VALUES (@ProductId, @SubstrateId)";
            var insertData = substrateIds.Select(id => new { ProductId = productId, SubstrateId = id });
            con.Execute(insertSql, insertData, tx);

            tx.Commit();
        }
    }
}
