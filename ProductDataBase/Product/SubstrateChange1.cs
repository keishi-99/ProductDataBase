using Microsoft.Data.Sqlite;
using ProductDatabase.Other;
using System.Data;
using static ProductDatabase.MainWindow;

namespace ProductDatabase {
    public partial class SubstrateChange1 : Form {

        private readonly ProductMaster _productMaster;
        private readonly ProductRegisterWork _productRegisterWork;
        private readonly AppSettings _appSettings;

        public DataTable HistoryTable { get; set; } = new();

        private readonly List<string> _colFilter = [];

        public SubstrateChange1(ProductMaster productMaster, ProductRegisterWork productRegisterWork, AppSettings appSettings) {
            InitializeComponent();

            _productMaster = productMaster;
            _productRegisterWork = productRegisterWork;
            _appSettings = appSettings;

            SubstrateChangeDataGridView.RowTemplate.DefaultCellStyle.Padding = new Padding(5);
        }

        // ロードイベント
        private void LoadEvents() {
            Font = new Font(_appSettings.FontName, _appSettings.FontSize);

            using SqliteConnection con = new(GetConnectionRegistration());
            con.Open();
            HistoryTable.Clear();

            using var cmd = con.CreateCommand();
            cmd.CommandText =
                $"""
                    SELECT
                        ID,
                        ProductName,
                        ProductType,
                        ProductModel,
                        OrderNumber,
                        ProductNumber,
                        Quantity,
                        SerialFirst,
                        SerialLast,
                        SerialLastNumber,
                        Revision,
                        RevisionGroup,
                        Person,
                        RegDate,    
                        Comment,
                        CreatedAt
                    FROM
                        {Constants.VProductTableName}
                    WHERE
                        ProductID = @ProductID 
                        AND Quantity > 1
                    ORDER BY
                        ID DESC
                    ;
                    """;

            cmd.Parameters.Add("@ProductID", SqliteType.Text).Value = _productMaster.ProductID;

            using (var reader = cmd.ExecuteReader()) {
                HistoryTable.Load(reader);
            }

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
            SubstrateChangeDataGridView.ColumnHeadersDefaultCellStyle.Font = new Font(SubstrateChangeDataGridView.Font, FontStyle.Bold);
            SubstrateChangeDataGridView.RowHeadersVisible = true;
            SubstrateChangeDataGridView.AlternatingRowsDefaultCellStyle.BackColor = Color.Gainsboro;

            SubstrateChangeDataGridView.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            SubstrateChangeDataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;

            SubstrateChangeDataGridView.Columns["ID"].HeaderCell.Value = "ID";
            SubstrateChangeDataGridView.Columns["ID"].Width = 40;
            SubstrateChangeDataGridView.Columns["ProductName"].HeaderCell.Value = "製品名";
            SubstrateChangeDataGridView.Columns["OrderNumber"].HeaderCell.Value = "注文番号";
            SubstrateChangeDataGridView.Columns["ProductNumber"].HeaderCell.Value = "製造番号";
            SubstrateChangeDataGridView.Columns["ProductNumber"].Width = 130;
            SubstrateChangeDataGridView.Columns["ProductType"].HeaderCell.Value = "製品名";
            SubstrateChangeDataGridView.Columns["ProductModel"].HeaderCell.Value = "製品型式";
            SubstrateChangeDataGridView.Columns["Quantity"].HeaderCell.Value = "数量";
            SubstrateChangeDataGridView.Columns["Quantity"].Width = 40;
            SubstrateChangeDataGridView.Columns["Person"].HeaderCell.Value = "担当者";
            SubstrateChangeDataGridView.Columns["Person"].Width = 70;
            SubstrateChangeDataGridView.Columns["RegDate"].HeaderCell.Value = "登録日";
            SubstrateChangeDataGridView.Columns["RegDate"].Width = 80;
            SubstrateChangeDataGridView.Columns["Revision"].HeaderCell.Value = "Rev";
            SubstrateChangeDataGridView.Columns["Revision"].Width = 40;
            SubstrateChangeDataGridView.Columns["RevisionGroup"].HeaderCell.Value = "RevGroup";
            SubstrateChangeDataGridView.Columns["SerialFirst"].HeaderCell.Value = "シリアル先頭";
            SubstrateChangeDataGridView.Columns["SerialLast"].HeaderCell.Value = "シリアル末尾";
            SubstrateChangeDataGridView.Columns["SerialLastNumber"].HeaderCell.Value = "シリアル末番";
            SubstrateChangeDataGridView.Columns["SerialLastNumber"].Width = 40;
            SubstrateChangeDataGridView.Columns["Comment"].HeaderCell.Value = "コメント";

        }

        // 基板変更フォームを開く
        private void OpenSubstrateChangeWindow() {

            var i = SubstrateChangeDataGridView.SelectedCells[0].RowIndex;

            _productRegisterWork.RowID = Convert.ToInt32(SubstrateChangeDataGridView.Rows[i].Cells["ID"].Value);
            _productRegisterWork.OrderNumber = SubstrateChangeDataGridView.Rows[i].Cells["OrderNumber"].Value.ToString() ?? string.Empty;
            _productRegisterWork.ProductNumber = SubstrateChangeDataGridView.Rows[i].Cells["ProductNumber"].Value.ToString() ?? string.Empty;
            _productRegisterWork.Quantity = Convert.ToInt32(SubstrateChangeDataGridView.Rows[i].Cells["Quantity"].Value);
            _productRegisterWork.Revision = SubstrateChangeDataGridView.Rows[i].Cells["Revision"].Value.ToString() ?? string.Empty;
            _productRegisterWork.SerialFirst = SubstrateChangeDataGridView.Rows[i].Cells["SerialFirst"].Value.ToString() ?? string.Empty;
            _productRegisterWork.SerialLast = SubstrateChangeDataGridView.Rows[i].Cells["SerialLast"].Value.ToString() ?? string.Empty;
            _productRegisterWork.SerialLastNumber = Convert.ToInt32(SubstrateChangeDataGridView.Rows[i].Cells["SerialLastNumber"].Value);
            _productRegisterWork.Comment = SubstrateChangeDataGridView.Rows[i].Cells["Comment"].Value.ToString() ?? string.Empty;
            using SubstrateChange2 window = new(_productMaster, _productRegisterWork, _appSettings);
            window.Closed += (s, e) => this.Close();
            window.ShowDialog(this);
        }

        private void SubstrateChange1_Load(object sender, EventArgs e) { LoadEvents(); }
        private void OKButton_Click(object sender, EventArgs e) { OpenSubstrateChangeWindow(); }
    }
}
