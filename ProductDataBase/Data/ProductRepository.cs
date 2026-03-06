using Dapper;
using Microsoft.Data.Sqlite;
using ProductDatabase.Other;
using System.Data;

namespace ProductDatabase {
    public class ProductRepository {

        public DataTable ProductDataTable { get; } = new();
        public DataTable SubstrateDataTable { get; } = new();
        public DataTable ProductUseSubstrate { get; } = new();

        public static string GetConnectionRegistration() {
            var productRegistryPath = Path.Combine(AppContext.BaseDirectory, "db", "ProductRegistry.db");
            return !File.Exists(productRegistryPath)
                ? throw new FileNotFoundException("ファイルが見つかりません。", productRegistryPath)
                : new SqliteConnectionStringBuilder() { DataSource = productRegistryPath, Pooling = false }.ToString();
        }

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

        public DataRow GetSubstrateById(long id) {
            return SubstrateDataTable
                .AsEnumerable()
                .SingleOrDefault(r => r.Field<long>("SubstrateID") == id)
                ?? throw new InvalidOperationException($"基板ID [{id}] が見つかりません。");
        }

        public DataRow GetProductById(long id) {
            return ProductDataTable
                .AsEnumerable()
                .SingleOrDefault(r => r.Field<long>("ProductID") == id)
                ?? throw new InvalidOperationException($"製品ID [{id}] が見つかりません。");
        }

        // 製品IDから使用基板リストを取得
        public static List<SubstrateInfo> GetUseSubstrates(int productKey) {
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

        public void Clear() {
            ProductDataTable.Clear();
            SubstrateDataTable.Clear();
            ProductUseSubstrate.Clear();
        }
    }
}
