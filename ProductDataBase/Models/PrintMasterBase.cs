namespace ProductDatabase {
    /// <summary>
    /// ProductMaster / SubstrateMaster の共通印刷フラグ管理ロジックを持つ基底クラス。
    /// </summary>
    public abstract class PrintMasterBase {

        protected int _regType;
        public int RegType {
            get => _regType;
            set {
                _regType = value;
                UpdatePrintFlags();
            }
        }

        protected int _serialPrintType;
        public int SerialPrintType {
            get => _serialPrintType;
            set {
                _serialPrintType = value;
                UpdatePrintFlags();
            }
        }

        // ===== 共通フラグ =====

        public bool IsLabelPrint { get; protected set; }
        public bool IsSerialGeneration { get; protected set; }

        // ===== 内部更新処理 =====

        // RegTypeとSerialPrintTypeの両フラグを更新する
        protected virtual void UpdatePrintFlags() {
            UpdateRegTypeFlags();
            UpdateSerialPrintTypeFlags();
        }

        // RegTypeに基づきシリアル生成フラグを更新する
        protected virtual void UpdateRegTypeFlags() {
            IsSerialGeneration = RegType is 1 or 2 or 3 or 9;
        }

        // SerialPrintTypeのビットフラグをboolプロパティに展開する
        protected virtual void UpdateSerialPrintTypeFlags() {
            var flags = (SerialPrintTypeFlags)_serialPrintType;
            IsLabelPrint = flags.HasFlag(SerialPrintTypeFlags.Label);
        }
    }
}
