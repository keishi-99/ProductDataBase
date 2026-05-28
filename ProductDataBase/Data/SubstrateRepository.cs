using Dapper;
using Microsoft.Data.Sqlite;
using ProductDatabase.Common;
using ProductDatabase.Models;

namespace ProductDatabase.Data {
    // 基板マスターの CRUD 操作を担当するリポジトリクラス
    public static class SubstrateRepository {

        // 基板マスターを新規登録し採番されたSubstrateIDを返す
        public static long InsertSubstrate(SubstrateMaster substrate) {
            using var con = DbConnectionHelper.CreateAndOpenConnection();

            var sql =
                $"""
                INSERT INTO {Constants.SubstrateTableName}
                    (CategoryName, ProductName, SubstrateName, SubstrateModel,
                     RegType, Checkbox, SerialPrintType, Visible, ExclusiveGroupID)
                VALUES
                    (@CategoryName, @ProductName, @SubstrateName, @SubstrateModel,
                     @RegType, @Checkbox, @SerialPrintType, @Visible, @ExclusiveGroupID);
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
                Visible = substrate.Visible ? 1 : 0,
                substrate.ExclusiveGroupID
            });
        }

        // 基板マスターを更新する
        public static void UpdateSubstrate(SubstrateMaster substrate) {
            using var con = DbConnectionHelper.CreateAndOpenConnection();

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
                    Visible         = @Visible,
                    ExclusiveGroupID = @ExclusiveGroupID
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
                substrate.ExclusiveGroupID,
                substrate.SubstrateID
            });
        }

        // 基板マスターを物理削除する（実績存在チェック・関連紐づけ削除を含む）
        public static void DeleteSubstrate(long substrateId) {
            using var con = DbConnectionHelper.CreateAndOpenConnection();
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
            using var con = DbConnectionHelper.CreateAndOpenConnection();
            return con.ExecuteScalar<int>(
                $"SELECT COUNT(*) FROM {Constants.TSubstrateTableName} WHERE SubstrateID = @SubstrateId",
                new { SubstrateId = substrateId }) > 0;
        }

        // 全基板（非表示含む）の ExclusiveGroupID 最大値を取得する（新規グループID採番用）
        public static int GetMaxExclusiveGroupID() {
            using var con = DbConnectionHelper.CreateAndOpenConnection();
            return con.ExecuteScalar<int>("SELECT COALESCE(MAX(ExclusiveGroupID), 0) FROM M_SubstrateDef");
        }

        // 既存の排他グループ一覧を取得する（GroupID → 基板名リスト、表示中の基板のみ）
        public static Dictionary<int, List<string>> GetExclusiveGroups() {
            using var con = DbConnectionHelper.CreateAndOpenConnection();
            var rows = con.Query(
                "SELECT ExclusiveGroupID, SubstrateName FROM M_SubstrateDef WHERE ExclusiveGroupID IS NOT NULL AND Visible = 1 ORDER BY ExclusiveGroupID, SubstrateName");
            var result = new Dictionary<int, List<string>>();
            foreach (var row in rows) {
                int groupId = (int)(long)row.ExclusiveGroupID;
                string name = row.SubstrateName?.ToString() ?? string.Empty;
                if (!result.TryGetValue(groupId, out var list)) {
                    list = [];
                    result[groupId] = list;
                }
                list.Add(name);
            }
            return result;
        }

        // SubstrateModelの重複を確認する（excludeIdは編集時に自身を除外するために使用）
        public static bool ExistsSubstrateModel(string substrateModel, long excludeId = 0) {
            using var con = DbConnectionHelper.CreateAndOpenConnection();
            return con.ExecuteScalar<int>(
                $"SELECT COUNT(*) FROM {Constants.SubstrateTableName} WHERE SubstrateModel = @SubstrateModel AND SubstrateID != @ExcludeId",
                new { SubstrateModel = substrateModel, ExcludeId = excludeId }) > 0;
        }
    }
}
