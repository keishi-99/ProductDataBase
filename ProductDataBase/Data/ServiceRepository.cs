using Dapper;
using Microsoft.Data.Sqlite;
using ProductDatabase.Common;
using ProductDatabase.Models;
using System.Data;

namespace ProductDatabase.Data {
    // ServiceForm のDB操作を担当するリポジトリクラス
    internal static class ServiceRepository {

        // M_ProductDefから表示対象製品を取得して既存DataTableに読み込む
        public static void LoadVisibleProducts(DataTable table) {
            using var con = new SqliteConnection(ProductRepository.GetConnectionRegistration());
            con.Open();

            using var reader = con.ExecuteReader("SELECT * FROM M_ProductDef WHERE Visible = 1;");
            table.Load(reader);
        }

        // 製品IDに紐づく使用基板情報をV_ProductUseSubstrateから取得してリストで返す
        public static List<SubstrateInfo> GetUseSubstrates(long productId) {
            using var con = new SqliteConnection(ProductRepository.GetConnectionRegistration());
            con.Open();

            var query = $"SELECT * FROM {Constants.VProductUseSubstrate} WHERE P_ProductID = @ProductId AND S_SubstrateID IS NOT NULL;";
            var table = new DataTable();
            using (var reader = con.ExecuteReader(query, new { ProductId = productId })) {
                table.Load(reader);
            }

            return [.. table.AsEnumerable()
                .Select(r => new SubstrateInfo {
                    SubstrateID   = r.Field<long>("S_SubstrateID"),
                    SubstrateName = r["SubstrateName"]?.ToString() ?? "",
                    SubstrateModel = r["SubstrateModel"]?.ToString() ?? ""
                })];
        }
    }
}
