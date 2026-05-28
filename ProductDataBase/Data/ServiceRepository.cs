using Dapper;
using ProductDatabase.Models;
using System.Data;

namespace ProductDatabase.Data {
    // ServiceForm のDB操作を担当するリポジトリクラス
    internal static class ServiceRepository {

        // M_ProductDefから表示対象製品を取得して既存DataTableに読み込む
        public static void LoadVisibleProducts(DataTable table) {
            using var con = DbConnectionHelper.CreateAndOpenConnection();

            using var reader = con.ExecuteReader("SELECT * FROM M_ProductDef WHERE Visible = 1;");
            table.Load(reader);
        }

        // 製品IDに紐づく使用基板情報を取得してリストで返す（ExclusiveGroupID を含む）
        public static List<SubstrateInfo> GetUseSubstrates(long productId) {
            return ProductRepository.GetUseSubstrates(productId);
        }
    }
}
