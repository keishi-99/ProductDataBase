using System.Data;
using System.Data.SQLite;
using static ProductDatabase.MainWindow;

namespace ProductDatabase {
    public partial class SubstrateChange1 : Form {

        public ProductInfomation ProductInfo { get; set; } = new ProductInfomation();

        public DataTable HistoryTable { get; set; } = new();

        private readonly List<string> _colFilter = [];

        public SubstrateChange1() {
            InitializeComponent();
        }

        // ロードイベント
        private void LoadEvents() {
            try {
                Font = new Font(ProductInfo.FontName, ProductInfo.FontSize);

                using SQLiteConnection con = new(GetConnectionString2());
                HistoryTable.Clear();
                using SQLiteDataAdapter adapter = new($"SELECT _rowid_, * FROM Product_Reg_{ProductInfo.ProductName} WHERE col_Product_Model = '{ProductInfo.ProductModel}' ORDER BY _rowid_ DESC", con);
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

                SubstrateChangeDataGridView.Columns[0].HeaderCell.Value = "ID";
                SubstrateChangeDataGridView.Columns[0].Width = 40;
                SubstrateChangeDataGridView.Columns[1].HeaderCell.Value = "注文番号";
                SubstrateChangeDataGridView.Columns[2].HeaderCell.Value = "製造番号";
                SubstrateChangeDataGridView.Columns[2].Width = 130;
                SubstrateChangeDataGridView.Columns[3].HeaderCell.Value = "製品名";
                SubstrateChangeDataGridView.Columns[4].HeaderCell.Value = "製品型式";
                SubstrateChangeDataGridView.Columns[5].HeaderCell.Value = "数量";
                SubstrateChangeDataGridView.Columns[5].Width = 40;
                SubstrateChangeDataGridView.Columns[6].HeaderCell.Value = "担当者";
                SubstrateChangeDataGridView.Columns[6].Width = 70;
                SubstrateChangeDataGridView.Columns[7].HeaderCell.Value = "登録日";
                SubstrateChangeDataGridView.Columns[7].Width = 80;
                SubstrateChangeDataGridView.Columns[8].HeaderCell.Value = "Rev";
                SubstrateChangeDataGridView.Columns[8].Width = 40;
                SubstrateChangeDataGridView.Columns[9].HeaderCell.Value = "シリアル先頭";
                SubstrateChangeDataGridView.Columns[10].HeaderCell.Value = "シリアル末尾";
                SubstrateChangeDataGridView.Columns[11].HeaderCell.Value = "末番";
                SubstrateChangeDataGridView.Columns[11].Width = 40;
                SubstrateChangeDataGridView.Columns[12].HeaderCell.Value = "コメント";
                SubstrateChangeDataGridView.Columns[13].HeaderCell.Value = "使用基板";
                SubstrateChangeDataGridView.Columns[13].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                //SubstrateChangeDataGridView.Columns[13].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;

            } catch (Exception ex) {
                MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } finally {
            }
        }

        // 基板変更フォームを開く
        private void OpenSubstrateChangeWindow() {
            using SubstrateChange2 window = new();

            var i = SubstrateChangeDataGridView.SelectedCells[0].RowIndex;

            ProductInfo.OrderNumber = SubstrateChangeDataGridView.Rows[i].Cells[1].Value.ToString() ?? string.Empty;
            ProductInfo.ProductNumber = SubstrateChangeDataGridView.Rows[i].Cells[2].Value.ToString() ?? string.Empty;
            ProductInfo.ProductType = SubstrateChangeDataGridView.Rows[i].Cells[3].Value.ToString() ?? string.Empty;
            ProductInfo.ProductModel = SubstrateChangeDataGridView.Rows[i].Cells[4].Value.ToString() ?? string.Empty;
            ProductInfo.Revision = SubstrateChangeDataGridView.Rows[i].Cells[8].Value.ToString() ?? string.Empty;
            ProductInfo.SerialFirst = SubstrateChangeDataGridView.Rows[i].Cells[9].Value.ToString() ?? string.Empty;
            ProductInfo.SerialLast = SubstrateChangeDataGridView.Rows[i].Cells[10].Value.ToString() ?? string.Empty;
            ProductInfo.Comment = SubstrateChangeDataGridView.Rows[i].Cells[12].Value.ToString() ?? string.Empty;
            ProductInfo.UsedSubstrate = SubstrateChangeDataGridView.Rows[i].Cells[13].Value.ToString() ?? string.Empty;
            ProductInfo.Quantity = Convert.ToInt32(SubstrateChangeDataGridView.Rows[i].Cells[5].Value);
            ProductInfo.SerialLastNumber = Convert.ToInt32(SubstrateChangeDataGridView.Rows[i].Cells[11].Value);
            window.ProductInfo = ProductInfo;
            window.ShowDialog(this);
            Close();
        }

        private void SubstrateChange1_Load(object sender, EventArgs e) { LoadEvents(); }
        private void OKButton_Click(object sender, EventArgs e) { OpenSubstrateChangeWindow(); }
    }
}
