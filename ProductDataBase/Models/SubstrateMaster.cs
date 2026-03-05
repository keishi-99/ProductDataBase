using System.Data;

namespace ProductDatabase {
    public class SubstrateMaster : PrintMasterBase {
        public int SubstrateID { get; set; }
        public string CategoryName { get; set; } = string.Empty;
        public string ProductName { get; set; } = string.Empty;
        public string SubstrateName { get; set; } = string.Empty;
        public string SubstrateModel { get; set; } = string.Empty;
        public int CheckBin { get; set; }

        public void LoadFrom(DataRow row) {
            SubstrateID = Convert.ToInt32(row["SubstrateID"]);
            CategoryName = row.Field<string>("CategoryName") ?? string.Empty;
            ProductName = row.Field<string>("ProductName") ?? string.Empty;
            SubstrateName = row.Field<string>("SubstrateName") ?? string.Empty;
            SubstrateModel = row.Field<string>("SubstrateModel") ?? string.Empty;
            RegType = Convert.ToInt32(row["RegType"]);
            CheckBin = Convert.ToInt32(row["Checkbox"].ToString(), 2);
            SerialPrintType = Convert.ToInt32(row["SerialPrintType"]);
        }

        public void Reset() {
            SubstrateID = 0;
            CategoryName = string.Empty;
            ProductName = string.Empty;
            SubstrateName = string.Empty;
            SubstrateModel = string.Empty;
            RegType = 0;
            CheckBin = 0;
            SerialPrintType = 0;
        }
    }
}
