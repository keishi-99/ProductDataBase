namespace ProductDatabase {
    public class ProductRegisterWork {
        public int RowID { get; set; }
        public int ProductID { get; set; }
        public string OrderNumber { get; set; } = string.Empty;
        public string ProductNumber { get; set; } = string.Empty;
        public string OLesNumber { get; set; } = string.Empty;
        public string SerialFirst { get; set; } = string.Empty;
        public string SerialLast { get; set; } = string.Empty;
        public int Quantity { get; set; }
        public string RegDate { get; set; } = string.Empty;
        public string Person { get; set; } = string.Empty;
        public string Revision { get; set; } = string.Empty;
        public string Comment { get; set; } = string.Empty;

        public int SerialFirstNumber { get; set; }
        public int SerialLastNumber { get; set; }

        // 製品登録作業データを初期値にリセットする
        public void Reset() {
            RowID = 0;
            ProductID = 0;
            ProductNumber = string.Empty;
            OrderNumber = string.Empty;
            SerialFirst = string.Empty;
            SerialLast = string.Empty;
            Quantity = 0;
            RegDate = string.Empty;
            Person = string.Empty;
            Revision = string.Empty;
            Comment = string.Empty;
            SerialFirstNumber = 0;
            SerialLastNumber = 0;
        }
    }
}
