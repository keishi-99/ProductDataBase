namespace ProductWebViewer.Models {
    public class SubstrateRecord {
        public long Id { get; set; }
        public long SubstrateID { get; set; }
        public string? CategoryName { get; set; }
        public string? ProductName { get; set; }
        public string? SubstrateName { get; set; }
        public string? SubstrateModel { get; set; }
        public string? OrderNumber { get; set; }
        public string? SubstrateNumber { get; set; }
        public long? Increase { get; set; }
        // Decrease・Defect は DB に負数で格納されている（在庫計算は Increase + Decrease + Defect の合算）
        public long? Decrease { get; set; }
        public long? Defect { get; set; }
        public string? PersonInfo { get; set; }
        public string? RegDate { get; set; }
        public string? Comment { get; set; }
        public string? CreatedAt { get; set; }
        public long? UseID { get; set; }
        public string? UseProductName { get; set; }
        public string? UseOrderNumber { get; set; }
        public string? UseProductNumber { get; set; }
    }
}
