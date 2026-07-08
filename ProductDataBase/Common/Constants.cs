namespace ProductDatabase.Common {
    public static class Constants {
        public const string ProductTableName = "M_ProductDef";
        public const string SubstrateTableName = "M_SubstrateDef";
        public const string TProductTableName = "T_Product";
        public const string TSubstrateTableName = "T_Substrate";
        public const string TSerialTableName = "T_Serial";
        public const string VSerialTableName = "V_Serial";
        public const string TRePrintTableName = "T_RePrint";
        public const string VRePrintTableName = "V_RePrint";
        public const string VProductTableName = "V_Product";
        public const string VSubstrateTableName = "V_Substrate";
        public const string VProductUseSubstrate = "V_ProductUseSubstrate";

        // 基板在庫数を求める集計式（Decrease・Defect はDBに負数で格納されているため単純合算で在庫数になる）
        public const string SubstrateStockSumExpression = "SUM(COALESCE(Increase, 0) + COALESCE(Decrease, 0) + COALESCE(Defect, 0)) AS Stock";
    }
}
