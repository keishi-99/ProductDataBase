using ProductDatabase.Common;
using ProductDatabase.Models;
using System.Data;

namespace ProductDatabase.Services {
    // 履歴編集・削除操作の監査ログエントリを整形するクラス
    internal static class HistoryAuditLogger {

        // Revision変更操作の監査ログをファイルに記録する
        public static void LogRevisionChange(ProductMaster productMaster, ProductRegisterWork productRegisterWork, long id) {
            Logger.AppendLog([
                "[Rev変更]",
                $"[{productMaster.CategoryName}]",
                $"[ID{id}]",
                $"[]",
                $"[]",
                $"[]",
                $"製品名[{productMaster.ProductName}]",
                $"タイプ[{productMaster.ProductType}]",
                $"型式[{productMaster.ProductModel}]",
                $"[]",
                $"[]",
                $"[]",
                $"Revision[{productRegisterWork.Revision}]",
                $"登録日[{productRegisterWork.RegDate}]",
                $"[]",
                $"コメント[{productRegisterWork.Comment}]"
            ]);
        }

        // 再印刷操作の監査ログをファイルに記録する
        public static void LogRePrint(ProductMaster productMaster, ProductRegisterWork productRegisterWork) {
            Logger.AppendLog([
                "[再印刷]",
                $"[{productMaster.CategoryName}]",
                $"[]",
                $"注文番号[{productRegisterWork.OrderNumber}]",
                $"製造番号[{productRegisterWork.ProductNumber}]",
                $"[]",
                $"製品名[{productMaster.ProductName}]",
                $"タイプ[{productMaster.ProductType}]",
                $"型式[{productMaster.ProductModel}]",
                $"数量[{productRegisterWork.Quantity}]",
                $"シリアル先頭[{productRegisterWork.SerialFirst}]",
                $"シリアル末尾[{productRegisterWork.SerialLast}]",
                $"Revision[{productRegisterWork.Revision}]",
                $"登録日[{productRegisterWork.RegDate}]",
                $"担当者[{productRegisterWork.Person}]",
                $"コメント[{productRegisterWork.Comment}]"
            ]);
        }

        // 基板登録操作の監査ログをファイルに記録する
        public static void LogSubstrateRegistration(
            SubstrateMaster substrateMaster, SubstrateRegisterWork substrateRegisterWork,
            string rowId, string logQuantity, string logDefectQuantity) {
            Logger.AppendLog([
                "[基板登録]",
                $"[{substrateMaster.CategoryName}]",
                $"ID[{rowId}]",
                $"注文番号[{substrateRegisterWork.OrderNumber}]",
                $"製造番号[{substrateRegisterWork.ProductNumber}]",
                $"[]",
                $"製品名[{substrateMaster.ProductName}]",
                $"基板名[{substrateMaster.SubstrateName}]",
                $"型式[{substrateMaster.SubstrateModel}]",
                $"追加数[{logQuantity}]",
                $"使用数[]",
                $"減少数[{logDefectQuantity}]",
                $"[]",
                $"登録日[{substrateRegisterWork.RegDate}]",
                $"担当者[{substrateRegisterWork.Person}]",
                $"コメント[{substrateRegisterWork.Comment}]"
            ]);
        }

        // 基板変更操作の監査ログをファイルに記録する
        public static void LogSubstrateChange(ProductMaster productMaster, ProductRegisterWork productRegisterWork) {
            Logger.AppendLog([
                "[基板変更]",
                $"[{productMaster.CategoryName}]",
                $"ID[{productRegisterWork.RowID}]",
                $"注文番号[{productRegisterWork.OrderNumber}]",
                $"製造番号[{productRegisterWork.ProductNumber}]",
                $"[]",
                $"製品名[{productMaster.ProductName}]",
                $"タイプ[{productMaster.ProductType}]",
                $"型式[{productMaster.ProductModel}]",
                $"数量[{productRegisterWork.Quantity}]",
                $"シリアル先頭[{productRegisterWork.SerialFirst}]",
                $"シリアル末尾[{productRegisterWork.SerialLast}]",
                $"Revision[{productRegisterWork.Revision}]",
                $"登録日[{productRegisterWork.RegDate}]",
                $"担当者[{productRegisterWork.Person}]",
                $"コメント[{productRegisterWork.Comment}]"
            ]);
        }

        // 製品登録操作の監査ログをファイルに記録する
        public static void LogProductRegistration(ProductMaster productMaster, ProductRegisterWork productRegisterWork) {
            Logger.AppendLog([
                "[製品登録]",
                $"[{productMaster.CategoryName}]",
                $"ID[{productRegisterWork.RowID}]",
                $"注文番号[{productRegisterWork.OrderNumber}]",
                $"製造番号[{productRegisterWork.ProductNumber}]",
                $"OLes番号[{productRegisterWork.OLesNumber}]",
                $"製品名[{productMaster.ProductName}]",
                $"タイプ[{productMaster.ProductType}]",
                $"型式[{productMaster.ProductModel}]",
                $"数量[{productRegisterWork.Quantity}]",
                $"シリアル先頭[{productRegisterWork.SerialFirst}]",
                $"シリアル末尾[{productRegisterWork.SerialLast}]",
                $"Revision[{productRegisterWork.Revision}]",
                $"登録日[{productRegisterWork.RegDate}]",
                $"担当者[{productRegisterWork.Person}]",
                $"コメント[{productRegisterWork.Comment}]"
            ]);
        }

        // 基板履歴編集の編集前後ログエントリを追加する
        public static void LogSubstrateEdit(DataRow row, List<string[]> pendingLogs, string categoryName) {
            pendingLogs.Add([
                "[基板履歴編集:前]",
                $"[{categoryName}]",
                $"ID[{GetValue(row, "ID", DataRowVersion.Original)}]",
                $"注文番号[{GetValue(row, "OrderNumber", DataRowVersion.Original)}]",
                $"製造番号[{GetValue(row, "SubstrateNumber", DataRowVersion.Original)}]",
                $"[]",
                $"製品名[{GetValue(row, "ProductName", DataRowVersion.Original)}]",
                $"基板名[{GetValue(row, "SubstrateName", DataRowVersion.Original)}]",
                $"型式[{GetValue(row, "SubstrateModel", DataRowVersion.Original)}]",
                $"追加数[{GetValue(row, "Increase", DataRowVersion.Original)}]",
                $"使用数[{GetValue(row, "Decrease", DataRowVersion.Original)}]",
                $"減少数[{GetValue(row, "Defect", DataRowVersion.Original)}]",
                $"[]",
                $"登録日[{GetValue(row, "RegDate", DataRowVersion.Original)}]",
                $"担当者[{GetValue(row, "Person", DataRowVersion.Original)}]",
                $"コメント[{GetValue(row, "Comment", DataRowVersion.Original)}]"
            ]);

            pendingLogs.Add([
                "[基板履歴編集:後]",
                $"[{categoryName}]",
                $"ID[{GetValue(row, "ID")}]",
                $"注文番号[{GetValue(row, "OrderNumber")}]",
                $"製造番号[{GetValue(row, "SubstrateNumber")}]",
                $"[]",
                $"製品名[{GetValue(row, "ProductName")}]",
                $"基板名[{GetValue(row, "SubstrateName")}]",
                $"型式[{GetValue(row, "SubstrateModel")}]",
                $"追加数[{GetValue(row, "Increase")}]",
                $"使用数[{GetValue(row, "Decrease")}]",
                $"減少数[{GetValue(row, "Defect")}]",
                $"[]",
                $"登録日[{GetValue(row, "RegDate")}]",
                $"担当者[{GetValue(row, "Person")}]",
                $"コメント[{GetValue(row, "Comment")}]"
            ]);
        }

        // 基板履歴削除のログエントリを追加する
        public static void LogSubstrateDelete(DataRow row, List<string[]> pendingLogs, string categoryName) {
            pendingLogs.Add([
                "[基板履歴削除]",
                $"[{categoryName}]",
                $"ID[{GetValue(row, "ID")}]",
                $"注文番号[{GetValue(row, "OrderNumber")}]",
                $"製造番号[{GetValue(row, "SubstrateNumber")}]",
                $"[]",
                $"製品名[{GetValue(row, "ProductName")}]",
                $"基板名[{GetValue(row, "SubstrateName")}]",
                $"型式[{GetValue(row, "SubstrateModel")}]",
                $"追加数[{GetValue(row, "Increase")}]",
                $"使用数[{GetValue(row, "Decrease")}]",
                $"減少数[{GetValue(row, "Defect")}]",
                $"[]",
                $"登録日[{GetValue(row, "RegDate")}]",
                $"担当者[{GetValue(row, "Person")}]",
                $"コメント[{GetValue(row, "Comment")}]"
            ]);
        }

        // 製品履歴編集の編集前後ログエントリを追加する
        public static void LogProductEdit(DataRow row, List<string[]> pendingLogs, string categoryName) {
            pendingLogs.Add([
                "[製品履歴編集:前]",
                $"[{categoryName}]",
                $"ID[{row["ID", DataRowVersion.Original]}]",
                $"注文番号[{row["OrderNumber", DataRowVersion.Original]}]",
                $"製造番号[{row["ProductNumber", DataRowVersion.Original]}]",
                $"OLes番号[{row["OLesNumber", DataRowVersion.Original]}]",
                $"製品名[{row["ProductName", DataRowVersion.Original]}]",
                $"タイプ[{row["ProductType", DataRowVersion.Original]}]",
                $"型式[{row["ProductModel", DataRowVersion.Original]}]",
                $"数量[{row["Quantity", DataRowVersion.Original]}]",
                $"シリアル先頭[{row["SerialFirst", DataRowVersion.Original]}]",
                $"シリアル末尾[{row["SerialLast", DataRowVersion.Original]}]",
                $"Revision[{row["Revision", DataRowVersion.Original]}]",
                $"登録日[{row["RegDate", DataRowVersion.Original]}]",
                $"担当者[{row["Person", DataRowVersion.Original]}]",
                $"コメント[{row["Comment", DataRowVersion.Original]}]"
            ]);

            pendingLogs.Add([
                "[製品履歴編集:後]",
                $"[{categoryName}]",
                $"ID[{GetValue(row, "ID")}]",
                $"注文番号[{GetValue(row, "OrderNumber")}]",
                $"製造番号[{GetValue(row, "ProductNumber")}]",
                $"OLes番号[{GetValue(row, "OLesNumber")}]",
                $"製品名[{GetValue(row, "ProductName")}]",
                $"タイプ[{GetValue(row, "ProductType")}]",
                $"型式[{GetValue(row, "ProductModel")}]",
                $"数量[{GetValue(row, "Quantity")}]",
                $"シリアル先頭[{GetValue(row, "SerialFirst")}]",
                $"シリアル末尾[{GetValue(row, "SerialLast")}]",
                $"Revision[{GetValue(row, "Revision")}]",
                $"登録日[{GetValue(row, "RegDate")}]",
                $"担当者[{GetValue(row, "Person")}]",
                $"コメント[{GetValue(row, "Comment")}]"
            ]);
        }

        // 製品履歴削除のログエントリを追加する
        public static void LogProductDelete(DataRow row, List<string[]> pendingLogs, string categoryName) {
            pendingLogs.Add([
                "[製品履歴削除]",
                $"[{categoryName}]",
                $"ID[{GetValue(row, "ID")}]",
                $"注文番号[{GetValue(row, "OrderNumber")}]",
                $"製造番号[{GetValue(row, "ProductNumber")}]",
                $"OLes番号[{GetValue(row, "OLesNumber")}]",
                $"製品名[{GetValue(row, "ProductName")}]",
                $"タイプ[{GetValue(row, "ProductType")}]",
                $"型式[{GetValue(row, "ProductModel")}]",
                $"数量[{GetValue(row, "Quantity")}]",
                $"シリアル先頭[{GetValue(row, "SerialFirst")}]",
                $"シリアル末尾[{GetValue(row, "SerialLast")}]",
                $"Revision[{GetValue(row, "Revision")}]",
                $"登録日[{GetValue(row, "RegDate")}]",
                $"担当者[{GetValue(row, "Person")}]",
                $"コメント[{GetValue(row, "Comment")}]"
            ]);
        }

        // 製品削除に伴う基板履歴削除のログエントリを追加する
        public static void LogProductSubstrateDelete(IEnumerable<dynamic> substrates, List<string[]> pendingLogs, string categoryName) {
            foreach (var item in substrates) {
                pendingLogs.Add([
                    "[製品削除に伴う基板削除]",
                    $"[{categoryName}]",
                    $"ID[{item.ID}]",
                    $"注文番号[{item.OrderNumber}]",
                    $"製造番号[{item.SubstrateNumber}]",
                    $"[]",
                    $"製品名[{item.ProductName}]",
                    $"基板名[{item.SubstrateName}]",
                    $"型式[{item.SubstrateModel}]",
                    $"追加数[{item.Increase}]",
                    $"使用数[{item.Decrease}]",
                    $"減少数[{item.Defect}]",
                    $"登録日[{item.RegDate}]",
                    $"担当者[{item.Person}]",
                    $"コメント[{item.Comment}]",
                    $"UseID[{item.UseID}]",
                ]);
            }
        }

        // 製品削除に伴うシリアル削除のログエントリを追加する
        public static void LogProductSerialDelete(IEnumerable<dynamic> serials, List<string[]> pendingLogs, string categoryName) {
            foreach (var item in serials) {
                pendingLogs.Add([
                    "[製品削除に伴うシリアル削除]",
                    $"[{categoryName}]",
                    $"ID[{item.rowid}]",
                    $"製品名[{item.ProductName}]",
                    $"Serial[{item.Serial}]",
                    $"UsedID[{item.UsedID}]",
                    $"[]", $"[]", $"[]", $"[]",
                    $"[]", $"[]", $"[]", $"[]",
                    $"[]", $"[]",
                ]);
            }
        }

        // シリアル履歴削除のログエントリを追加する
        public static void LogSerialDelete(DataRow row, List<string[]> pendingLogs, string categoryName) {
            pendingLogs.Add([
                "[シリアル履歴削除]",
                $"[{categoryName}]",
                $"ID[{GetValue(row, "rowid")}]",
                $"製品名[{GetValue(row, "ProductName")}]",
                $"Serial[{GetValue(row, "Serial")}]",
                $"UsedID[{GetValue(row, "UsedID")}]",
                $"[]", $"[]", $"[]", $"[]",
                $"[]", $"[]", $"[]", $"[]",
                $"[]", $"[]",
            ]);
        }

        // DataRowから指定カラムの値を文字列で取得しDBNullの場合は空文字を返す
        private static string GetValue(DataRow row, string columnName, DataRowVersion version = DataRowVersion.Current) {
            var value = row[columnName, version];
            return value == DBNull.Value ? "" : value.ToString() ?? "";
        }
    }
}
