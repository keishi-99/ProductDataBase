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

        // キャッシュ済みの全DataTableをクリアする
        public void Clear() {
            ProductDataTable.Clear();
            SubstrateDataTable.Clear();
            ProductUseSubstrate.Clear();
        }
    }
}
