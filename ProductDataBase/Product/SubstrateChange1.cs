using ProductDatabase.Data;
using ProductDatabase.Models;
using System.ComponentModel;
using System.Data;

namespace ProductDatabase {
    public partial class SubstrateChange1 : Form {

        private readonly ProductMaster _productMaster;
        private readonly ProductRegisterWork _productRegisterWork;
        private readonly AppSettings _appSettings;

        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public DataTable HistoryTable { get; set; } = new();

        private readonly List<string> _colFilter = [];

        public SubstrateChange1(ProductMaster productMaster, ProductRegisterWork productRegisterWork, AppSettings appSettings) {
            InitializeComponent();

            _productMaster = productMaster;
            _productRegisterWork = productRegisterWork;
            _appSettings = appSettings;

            SubstrateChangeDataGridView.AlternatingRowsDefaultCellStyle.BackColor = System.Drawing.Color.Lavender;
            SubstrateChangeDataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            SubstrateChangeDataGridView.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font(SubstrateChangeDataGridView.Font, FontStyle.Bold);
            SubstrateChangeDataGridView.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            SubstrateChangeDataGridView.ReadOnly = true;
            SubstrateChangeDataGridView.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            SubstrateChangeDataGridView.RowTemplate.DefaultCellStyle.Padding = new Padding(5);
            SubstrateChangeDataGridView.RowTemplate.Height += 10;

        }

        // フォームロード時にDBから対象製品の複数台登録履歴を取得しDataGridViewに表示する
        private void LoadEvents() {
            Font = new System.Drawing.Font(_appSettings.FontName, _appSettings.FontSize);

            HistoryTable = SubstrateChangeRepository.GetProductHistory(_productMaster.ProductID);

            SubstrateChangeDataGridView.DataSource = HistoryTable;

            _colFilter.Add("");
            for (var i = 0; i < SubstrateChangeDataGridView.ColumnCount; i++) {
                var headerValue = SubstrateChangeDataGridView.Columns[i].HeaderCell.Value?.ToString() ?? string.Empty;
                if (headerValue is not null) { _colFilter.Add(headerValue); }
            }

            if (Screen.PrimaryScreen is not null) {
                var h = Screen.PrimaryScreen.Bounds.Height;
                var w = Screen.PrimaryScreen.Bounds.Width;
                SubstrateChangeDataGridView.MaximumSize = new Size(w, h);
            }

            (string Name, string Header, int? Width)[] columnSettings = [
                ("ID", "ID", 40),
                ("ProductName", "製品名", null),
                ("OrderNumber", "注文番号", null),
                ("ProductNumber", "製造番号", 130),
                ("ProductType", "製品名", null),
                ("ProductModel", "製品型式", null),
                ("Quantity", "数量", 40),
                ("PersonInfo", "担当者", 70),
                ("RegDate", "登録日", 80),
                ("Revision", "Rev", 40),
                ("RevisionGroup", "RevGroup", null),
                ("SerialFirst", "シリアル先頭", null),
                ("SerialLast", "シリアル末尾", null),
                ("SerialLastNumber", "シリアル末番", 40),
                ("Comment", "コメント", null),
            ];
            foreach (var (name, header, width) in columnSettings) {
                var column = SubstrateChangeDataGridView.Columns[name]!;
                column.HeaderCell.Value = header;
                if (width is not null) { column.Width = width.Value; }
            }

        }

        // DataGridViewで選択した行の製品情報をWorkに格納しSubstrateChange2フォームをダイアログで開く
        private void OpenSubstrateChangeWindow() {

            var i = SubstrateChangeDataGridView.SelectedCells[0].RowIndex;

            _productRegisterWork.RowID = int.TryParse(SubstrateChangeDataGridView.Rows[i].Cells["ID"].Value?.ToString(), out var rowId) ? rowId : 0;
            _productRegisterWork.OrderNumber = SubstrateChangeDataGridView.Rows[i].Cells["OrderNumber"].Value?.ToString() ?? string.Empty;
            _productRegisterWork.ProductNumber = SubstrateChangeDataGridView.Rows[i].Cells["ProductNumber"].Value?.ToString() ?? string.Empty;
            _productRegisterWork.Quantity = int.TryParse(SubstrateChangeDataGridView.Rows[i].Cells["Quantity"].Value?.ToString(), out var qty) ? qty : 0;
            _productRegisterWork.Revision = SubstrateChangeDataGridView.Rows[i].Cells["Revision"].Value?.ToString() ?? string.Empty;
            _productRegisterWork.SerialFirst = SubstrateChangeDataGridView.Rows[i].Cells["SerialFirst"].Value?.ToString() ?? string.Empty;
            _productRegisterWork.SerialLast = SubstrateChangeDataGridView.Rows[i].Cells["SerialLast"].Value?.ToString() ?? string.Empty;
            _productRegisterWork.SerialLastNumber = int.TryParse(SubstrateChangeDataGridView.Rows[i].Cells["SerialLastNumber"].Value?.ToString(), out var sln) ? sln : 0;
            _productRegisterWork.Comment = SubstrateChangeDataGridView.Rows[i].Cells["Comment"].Value?.ToString() ?? string.Empty;
            using SubstrateChange2 window = new(_productMaster, _productRegisterWork, _appSettings);
            window.FormClosed += (s, e) => this.Close();
            window.ShowDialog(this);
        }

        private void SubstrateChange1_Load(object sender, EventArgs e) { LoadEvents(); }
        private void OKButton_Click(object sender, EventArgs e) { OpenSubstrateChangeWindow(); }
    }
}
