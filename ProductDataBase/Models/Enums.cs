namespace ProductDatabase.Models {

    [Flags]
    public enum SerialPrintTypeFlags {
        None = 0,
        Label = 1 << 0,         // 00001
        Barcode = 1 << 1,       // 00010
        Nameplate = 1 << 2,     // 00100
        Underline = 1 << 3,     // 01000
        Last4Digits = 1 << 4,   // 10000
        OLesSerial = 1 << 5     // 100000
    }

    [Flags]
    public enum SheetPrintTypeFlags {
        None = 0,
        CheckSheet = 1 << 0,    // 00001
        List = 1 << 1,          // 00010
    }

    public enum ProductOperationMode {
        Register,
        RePrint,
        SubstrateChange
    }

    // MainWindowのCategoryRadioButton1〜4で選択するモード
    public enum RadioButtonMode {
        Substrate = 1,       // 基板登録
        ProductRegister = 2, // 製品登録
        RePrint = 3,          // 再印刷
        SubstrateChange = 4   // 基板変更
    }
}
