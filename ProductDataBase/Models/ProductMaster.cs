using System.Data;

namespace ProductDatabase.Models {
    public class ProductMaster : PrintMasterBase {
        public long ProductID { get; set; }
        public string CategoryName { get; set; } = string.Empty;
        public string ProductName { get; set; } = string.Empty;
        public string ProductModel { get; set; } = string.Empty;
        public string ProductType { get; set; } = string.Empty;
        public string Initial { get; set; } = string.Empty;
        public string OLesInitial { get; set; } = string.Empty;
        public int SerialDigitType { get; set; }
        public int SerialDigit => SerialDigitType switch {
            3 or 101 or 102 => 3,
            4 => 4,
            _ => 0
        };
        public int RevisionGroup { get; set; }
        public int CheckBin { get; set; }
        public List<SubstrateInfo> UseSubstrates { get; set; } = [];

        public bool Visible { get; set; } = true;

        private int _sheetPrintType;
        public int SheetPrintType {
            get => _sheetPrintType;
            set {
                _sheetPrintType = value;
                UpdatePrintFlags();
            }
        }

        // ===== 製品固有フラグ =====

        public bool IsBarcodePrint { get; private set; }
        public bool IsNameplatePrint { get; private set; }
        public bool IsOLesLabelPrint { get; private set; }
        public bool IsLast4Digits { get; private set; }
        public bool IsUnderlinePrint { get; private set; }

        public bool IsCheckSheetPrint { get; private set; }
        public bool IsListPrint { get; private set; }

        public bool IsRegType9 { get; private set; }

        // ===== 内部更新処理（基底クラスを拡張）=====

        // 基底クラスのフラグ更新に加えSheetPrintTypeフラグも更新する
        protected override void UpdatePrintFlags() {
            base.UpdatePrintFlags();
            UpdateSheetPrintTypeFlags();
        }

        // 基底クラスの更新に加えIsRegType9フラグを設定する
        protected override void UpdateRegTypeFlags() {
            base.UpdateRegTypeFlags();
            IsRegType9 = RegType == 9;
        }

        // 基底クラスの更新に加え製品固有の印刷フラグを展開する
        protected override void UpdateSerialPrintTypeFlags() {
            base.UpdateSerialPrintTypeFlags();
            var flags = (SerialPrintTypeFlags)_serialPrintType;
            IsBarcodePrint = flags.HasFlag(SerialPrintTypeFlags.Barcode);
            IsNameplatePrint = flags.HasFlag(SerialPrintTypeFlags.Nameplate);
            IsOLesLabelPrint = flags.HasFlag(SerialPrintTypeFlags.OLesSerial);
            IsUnderlinePrint = flags.HasFlag(SerialPrintTypeFlags.Underline);
            IsLast4Digits = flags.HasFlag(SerialPrintTypeFlags.Last4Digits);
        }

        // SheetPrintTypeのビットフラグをboolプロパティに展開する
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

        // DataRowから製品マスターの各フィールドを読み込む
        public void LoadFrom(DataRow row) {
            ProductID = row.Field<long>("ProductID");
            CategoryName = row.Field<string>("CategoryName") ?? string.Empty;
            ProductName = row.Field<string>("ProductName") ?? string.Empty;
            ProductType = row.Field<string>("ProductType") ?? string.Empty;
            ProductModel = row.Field<string>("ProductModel") ?? string.Empty;
            Initial = row.Field<string>("Initial") ?? string.Empty;
            OLesInitial = row.Field<string>("OLesInitial") ?? string.Empty;
            RevisionGroup = (int)row.Field<long>("RevisionGroup");
            RegType = (int)row.Field<long>("RegType");
            CheckBin = Convert.ToInt32(row["Checkbox"].ToString(), 2); // バイナリ文字列変換のため維持
            SerialDigitType = (int)row.Field<long>("SerialType");
            SerialPrintType = (int)row.Field<long>("SerialPrintType");
            SheetPrintType = (int)row.Field<long>("SheetPrintType");
            Visible = row.Field<long?>("Visible") == 1;
        }

        // 製品マスターデータを初期値にリセットする
        public void Reset() {
            ProductID = 0;
            CategoryName = string.Empty;
            ProductName = string.Empty;
            ProductModel = string.Empty;
            ProductType = string.Empty;
            Initial = string.Empty;
            OLesInitial = string.Empty;
            SerialDigitType = 0;
            RevisionGroup = 0;
            CheckBin = 0;
            UseSubstrates = [];
            RegType = 0;
            SerialPrintType = 0;
            SheetPrintType = 0;
            Visible = true;
        }
    }
}
