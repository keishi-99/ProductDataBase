using System.Data;
using System.Data.SQLite;
using static ProductDatabase.MainWindow;

namespace ProductDatabase {
    public partial class SubstrateChange1 : Form {

        public ProductInformation ProductInfo { get; }

        public DataTable HistoryTable { get; set; } = new();

        private readonly List<string> _colFilter = [];

        public SubstrateChange1(ProductInformation productInfo) {
            InitializeComponent();
            ProductInfo = productInfo;
        }

        // ロードイベント
        private void LoadEvents() {
            try {
                Font = new Font(ProductInfo.FontName, ProductInfo.FontSize);

                using SQLiteConnection con = new(GetConnectionRegistration());
                HistoryTable.Clear();
                using SQLiteDataAdapter adapter = new($"""SELECT * FROM {ProductInfo.CategoryName}_Product WHERE ProductModel = '{ProductInfo.ProductModel}' ORDER BY "ID" DESC""", con);
                adapter.Fill(HistoryTable);

                SubstrateChangeDataGridView.DataSource = HistoryTable;

                _colFilter.Add("");
                for (var i = 0; i < SubstrateChangeDataGridView.ColumnCount; i++) {
                    var headerValue = SubstrateChangeDataGridView.Columns[i].HeaderCell.Value?.ToString() ?? string.Empty;
                    if (headerValue != null) { _colFilter.Add(headerValue); }
                }

                // 最大サイズをディスプレイサイズに合わせる
                if (Screen.PrimaryScreen != null) {
                    var h = Screen.PrimaryScreen.Bounds.Height;
                    var w = Screen.PrimaryScreen.Bounds.Width;
                    SubstrateChangeDataGridView.MaximumSize = new Size(w, h);
                }
                SubstrateChangeDataGridView.ColumnHeadersDefaultCellStyle.Font = new Font(SubstrateChangeDataGridView.Font, FontStyle.Bold);
                SubstrateChangeDataGridView.RowHeadersVisible = true;
                SubstrateChangeDataGridView.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                SubstrateChangeDataGridView.AlternatingRowsDefaultCellStyle.BackColor = Color.Gainsboro;
                //ヘッダーとすべてのセルの内容に合わせて、行の高さを自動調整する
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
                SubstrateChangeDataGridView.Columns["SerialLastNumber"].HeaderCell.Value = "末番";
                SubstrateChangeDataGridView.Columns["SerialLastNumber"].Width = 40;
                SubstrateChangeDataGridView.Columns["Comment"].HeaderCell.Value = "コメント";

            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } finally {
            }
        }

        // 基板変更フォームを開く
        private void OpenSubstrateChangeWindow() {

            var i = SubstrateChangeDataGridView.SelectedCells[0].RowIndex;

            ProductInfo.ProductID = Convert.ToInt32(SubstrateChangeDataGridView.Rows[i].Cells["ID"].Value);
            ProductInfo.OrderNumber = SubstrateChangeDataGridView.Rows[i].Cells["OrderNumber"].Value.ToString() ?? string.Empty;
            ProductInfo.ProductNumber = SubstrateChangeDataGridView.Rows[i].Cells["ProductNumber"].Value.ToString() ?? string.Empty;
            ProductInfo.Quantity = Convert.ToInt32(SubstrateChangeDataGridView.Rows[i].Cells["Quantity"].Value);
            ProductInfo.Revision = SubstrateChangeDataGridView.Rows[i].Cells["Revision"].Value.ToString() ?? string.Empty;
            ProductInfo.SerialFirst = SubstrateChangeDataGridView.Rows[i].Cells["SerialFirst"].Value.ToString() ?? string.Empty;
            ProductInfo.SerialLast = SubstrateChangeDataGridView.Rows[i].Cells["SerialLast"].Value.ToString() ?? string.Empty;
            ProductInfo.SerialLastNumber = Convert.ToInt32(SubstrateChangeDataGridView.Rows[i].Cells["SerialLastNumber"].Value);
            ProductInfo.Comment = SubstrateChangeDataGridView.Rows[i].Cells["Comment"].Value.ToString() ?? string.Empty;
            using SubstrateChange2 window = new();
            window.ProductInfo = ProductInfo;
            window.Closed += (s, e) => this.Close();
            window.ShowDialog(this);
        }

        private void SubstrateChange1_Load(object sender, EventArgs e) { LoadEvents(); }
        private void OKButton_Click(object sender, EventArgs e) { OpenSubstrateChangeWindow(); }
    }
}
