using System.Data;

namespace ProductDatabase.Models {
    public class SubstrateMaster : PrintMasterBase {
        public long SubstrateID { get; set; }
        public string CategoryName { get; set; } = string.Empty;
        public string ProductName { get; set; } = string.Empty;
        public string SubstrateName { get; set; } = string.Empty;
        public string SubstrateModel { get; set; } = string.Empty;
        public int CheckBin { get; set; }
        public bool Visible { get; set; } = true;

        // DataRowから基板マスターの各フィールドを読み込む
        public void LoadFrom(DataRow row) {
            SubstrateID = row.Field<long>("SubstrateID");
            CategoryName = row.Field<string>("CategoryName") ?? string.Empty;
            ProductName = row.Field<string>("ProductName") ?? string.Empty;
            SubstrateName = row.Field<string>("SubstrateName") ?? string.Empty;
            SubstrateModel = row.Field<string>("SubstrateModel") ?? string.Empty;
            RegType = (int)row.Field<long>("RegType");
            CheckBin = Convert.ToInt32(row["Checkbox"].ToString(), 2); // バイナリ文字列変換のため維持
            SerialPrintType = (int)row.Field<long>("SerialPrintType");
            Visible = row.Field<long?>("Visible") == 1;
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
