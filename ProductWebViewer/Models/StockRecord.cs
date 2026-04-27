namespace ProductWebViewer.Models {
    public class StockRecord {
        public long SubstrateID { get; set; }
        public string? CategoryName { get; set; }
        public string? ProductName { get; set; }
        public string? SubstrateName { get; set; }
        public string? SubstrateModel { get; set; }
        public string? SubstrateNumber { get; set; }
        public string? OrderNumber { get; set; }
        public long Stock { get; set; }
    }
}
