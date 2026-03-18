using System.Data;

namespace ProductDatabase.Models {
    public class SubstrateMaster : PrintMasterBase {
        public int SubstrateID { get; set; }
        public string CategoryName { get; set; } = string.Empty;
        public string ProductName { get; set; } = string.Empty;
        public string SubstrateName { get; set; } = string.Empty;
        public string SubstrateModel { get; set; } = string.Empty;
        public int CheckBin { get; set; }
        public bool Visible { get; set; } = true;

        // DataRowから基板マスターの各フィールドを読み込む
        public void LoadFrom(DataRow row) {
            SubstrateID = Convert.ToInt32(row["SubstrateID"]);
            CategoryName = row.Field<string>("CategoryName") ?? string.Empty;
            ProductName = row.Field<string>("ProductName") ?? string.Empty;
            SubstrateName = row.Field<string>("SubstrateName") ?? string.Empty;
            SubstrateModel = row.Field<string>("SubstrateModel") ?? string.Empty;
            RegType = Convert.ToInt32(row["RegType"]);
            CheckBin = Convert.ToInt32(row["Checkbox"].ToString(), 2);
            SerialPrintType = Convert.ToInt32(row["SerialPrintType"]);
            Visible = row["Visible"] != DBNull.Value && Convert.ToInt32(row["Visible"]) == 1;
        }

        // 基板マスターデータを初期値にリセットする
        public void Reset() {
            SubstrateID = 0;
            CategoryName = string.Empty;
            ProductName = string.Empty;
            SubstrateName = string.Empty;
            SubstrateModel = string.Empty;
            RegType = 0;
            CheckBin = 0;
            SerialPrintType = 0;
            Visible = true;
        }
    }
}
