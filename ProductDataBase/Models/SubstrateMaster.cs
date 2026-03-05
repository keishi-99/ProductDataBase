using System.Data;

namespace ProductDatabase {
    public class SubstrateMaster {
        public int SubstrateID { get; set; }
        public string CategoryName { get; set; } = string.Empty;
        public string ProductName { get; set; } = string.Empty;
        public string SubstrateName { get; set; } = string.Empty;
        public string SubstrateModel { get; set; } = string.Empty;
        public int CheckBin { get; set; }

        private int _regType;
        public int RegType {
            get => _regType;
            set {
                _regType = value;
                UpdatePrintFlags();
            }
        }

        private int _serialPrintType;
        public int SerialPrintType {
            get => _serialPrintType;
            set {
                _serialPrintType = value;
                UpdatePrintFlags();
            }
        }

        // ===== 結果フラグ =====

        public bool IsLabelPrint { get; private set; }

        public bool IsSerialGeneration { get; private set; }

        // ===== 内部更新処理 =====

        private void UpdatePrintFlags() {
            UpdateRegTypeFlags();
            UpdateSerialPrintTypeFlags();
        }

        private void UpdateRegTypeFlags() {
            IsSerialGeneration = RegType is 1 or 2 or 3 or 9;
        }

        private void UpdateSerialPrintTypeFlags() {
            var flags = (SerialPrintTypeFlags)_serialPrintType;

            IsLabelPrint = flags.HasFlag(SerialPrintTypeFlags.Label);
        }

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
