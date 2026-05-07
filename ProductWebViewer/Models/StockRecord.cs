namespace ProductWebViewer.Models {
    public class StockRecord {
        public long SubstrateID { get; set; }
        public string? CategoryName { get; set; }
        public string? ProductName { get; set; }
        public string? SubstrateName { get; set; }
        public string? SubstrateModel { get; set; }
        // groupByModel=false の場合のみ値が入る
        public string? SubstrateNumber { get; set; }
        public string? OrderNumber { get; set; }
        // Increase + Decrease(負数) + Defect(負数) の累積合算値
        public long Stock { get; set; }
    }
}
