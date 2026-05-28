namespace ProductDatabase.Models {
    public class SubstrateRegisterWork {
        public long SubstrateID { get; set; }

        public string ProductNumber { get; set; } = string.Empty;
        public string OrderNumber { get; set; } = string.Empty;
        public int AddQuantity { get; set; }
        public int DefectQuantity { get; set; }
        public int UseQuantity { get; set; }

        public long? PersonID { get; set; }
        public string PersonName { get; set; } = string.Empty;
        public string RegDate { get; set; } = string.Empty;
        public string Comment { get; set; } = string.Empty;

        // 基板登録作業データを初期値にリセットする
        public void Reset() {
            SubstrateID = 0;
            ProductNumber = string.Empty;
            OrderNumber = string.Empty;
            AddQuantity = 0;
            DefectQuantity = 0;
            UseQuantity = 0;
            PersonID = null;
            PersonName = string.Empty;
            RegDate = string.Empty;
            Comment = string.Empty;
        }
    }
}
