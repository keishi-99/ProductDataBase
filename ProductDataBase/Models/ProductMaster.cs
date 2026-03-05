using System.Data;

namespace ProductDatabase {
    public class ProductMaster {
        public int ProductID { get; set; }
        public string CategoryName { get; set; } = string.Empty;
        public string ProductName { get; set; } = string.Empty;
        public string ProductModel { get; set; } = string.Empty;
        public string ProductType { get; set; } = string.Empty;
        public string Initial { get; set; } = string.Empty;
        public int SerialDigitType { get; set; }
        public int SerialDigit => SerialDigitType switch {
            3 or 101 or 102 => 3,
            4 => 4,
            _ => 0
        };
        public int RevisionGroup { get; set; }
        public int CheckBin { get; set; }
        public List<SubstrateInfo> UseSubstrates { get; set; } = [];

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

        private int _sheetPrintType;
        public int SheetPrintType {
            get => _sheetPrintType;
            set {
                _sheetPrintType = value;
                UpdatePrintFlags();
            }
        }

        // ===== 結果フラグ =====

        public bool IsLabelPrint { get; private set; }
        public bool IsBarcodePrint { get; private set; }
        public bool IsNameplatePrint { get; private set; }
        public bool IsLast4Digits { get; private set; }
        public bool IsUnderlinePrint { get; private set; }

        public bool IsCheckSheetPrint { get; private set; }
        public bool IsListPrint { get; private set; }

        public bool IsSerialGeneration { get; private set; }
        public bool IsRegType9 { get; private set; }

        // ===== 内部更新処理 =====

        private void UpdatePrintFlags() {
            UpdateRegTypeFlags();
            UpdateSerialPrintTypeFlags();
            UpdateSheetPrintTypeFlags();
        }

        private void UpdateRegTypeFlags() {
            IsRegType9 = RegType == 9;
            IsSerialGeneration = RegType is 1 or 2 or 3 or 9;
        }

        private void UpdateSerialPrintTypeFlags() {
            var flags = (SerialPrintTypeFlags)_serialPrintType;

            IsLabelPrint = flags.HasFlag(SerialPrintTypeFlags.Label);
            IsBarcodePrint = flags.HasFlag(SerialPrintTypeFlags.Barcode);
            IsNameplatePrint = flags.HasFlag(SerialPrintTypeFlags.Nameplate);
            IsUnderlinePrint = flags.HasFlag(SerialPrintTypeFlags.Underline);
            IsLast4Digits = flags.HasFlag(SerialPrintTypeFlags.Last4Digits);
        }

        private void UpdateSheetPrintTypeFlags() {
            var flags = (SheetPrintTypeFlags)_sheetPrintType;

            IsCheckSheetPrint = flags.HasFlag(SheetPrintTypeFlags.CheckSheet);
            IsListPrint = flags.HasFlag(SheetPrintTypeFlags.List);
        }

        /// <summary>
        /// SerialDigitType に応じたシリアル番号の範囲を返します。
        /// </summary>
        public (int minNumber, int maxNumber, int digit) GetSerialRange() => SerialDigitType switch {
            3 => (1, 999, 3),
            4 => (1, 9999, 4),
            101 => (1, 899, 3),
            102 => (901, 999, 3),
            _ => throw new InvalidOperationException("不明なシリアル桁数です。")
        };

        public void LoadFrom(DataRow row) {
            ProductID = Convert.ToInt32(row["ProductID"]);
            CategoryName = row.Field<string>("CategoryName") ?? string.Empty;
            ProductName = row.Field<string>("ProductName") ?? string.Empty;
            ProductType = row.Field<string>("ProductType") ?? string.Empty;
            ProductModel = row.Field<string>("ProductModel") ?? string.Empty;
            Initial = row.Field<string>("Initial") ?? string.Empty;
            RevisionGroup = Convert.ToInt32(row["RevisionGroup"]);
            RegType = Convert.ToInt32(row["RegType"]);
            CheckBin = Convert.ToInt32(row["Checkbox"].ToString(), 2);
            SerialDigitType = Convert.ToInt32(row["SerialType"]);
            SerialPrintType = Convert.ToInt32(row["SerialPrintType"]);
            SheetPrintType = Convert.ToInt32(row["SheetPrintType"]);
        }

        public void Reset() {
            ProductID = 0;
            CategoryName = string.Empty;
            ProductName = string.Empty;
            ProductModel = string.Empty;
            ProductType = string.Empty;
            Initial = string.Empty;
            SerialDigitType = 0;
            RevisionGroup = 0;
            CheckBin = 0;
            UseSubstrates = [];
            RegType = 0;
            SerialPrintType = 0;
            SheetPrintType = 0;
        }
    }
}
