using System.Data;
using System.Data.SQLite;

namespace ProductDatabase {
    public partial class SubstrateChange1 : Form {

        public string StrFontName { get; set; } = "Meiryo UI";
        public int IntFontSize { get; set; } = 9;

        public DataTable DtHistoryTable { get; set; } = new();

        private readonly List<string> _listColFilter = [];

        public int IntRadioBtnFlg { get; set; }

        public string StrProductName { get; set; } = string.Empty;
        public string StrStockName { get; set; } = string.Empty;
        public string StrProductModel { get; set; } = string.Empty;
        public string StrProductType { get; set; } = string.Empty;
        public string StrSubstrateName { get; set; } = string.Empty;
        public string StrSubstrateModel { get; set; } = string.Empty;
        public string StrUseSubstrate { get; set; } = string.Empty;

        public int IntPrintType { get; set; }
        public int IntRegType { get; set; }

        public SubstrateChange1() {
            InitializeComponent();
        }

        // ロードイベント
        private void LoadEvents() {
            try {
                Font = new Font(StrFontName, IntFontSize);

                switch (IntRadioBtnFlg) {
                    case 4:
                        using (SQLiteConnection con = new(MainWindow.GetConnectionString2())) {
                            DtHistoryTable.Clear();
                            using SQLiteDataAdapter adapter = new($"SELECT _rowid_, * FROM Product_Reg_{StrProductName} WHERE col_Product_Model = '{StrProductModel}' ORDER BY _rowid_ DESC", con);
                            adapter.Fill(DtHistoryTable);

                            SubstrateChangeDataGridView.DataSource = DtHistoryTable;

                            _listColFilter.Add("");
                            for (var i = 0; i < SubstrateChangeDataGridView.ColumnCount; i++) {
                                var headerValue = SubstrateChangeDataGridView.Columns[i].HeaderCell.Value?.ToString() ?? string.Empty;
                                if (headerValue != null) { _listColFilter.Add(headerValue); }
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

                        }
                        break;
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } finally {
            }
        }

        // 基板変更フォームを開く
        private void OpenSubstrateChangeWindow() {

            var i = SubstrateChangeDataGridView.SelectedCells[0].RowIndex;

            var orderNumber = SubstrateChangeDataGridView.Rows[i].Cells[1].Value.ToString() ?? string.Empty;
            var productNumber = SubstrateChangeDataGridView.Rows[i].Cells[2].Value.ToString() ?? string.Empty;
            var productType = SubstrateChangeDataGridView.Rows[i].Cells[3].Value.ToString() ?? string.Empty;
            var productModel = SubstrateChangeDataGridView.Rows[i].Cells[4].Value.ToString() ?? string.Empty;
            var revision = SubstrateChangeDataGridView.Rows[i].Cells[8].Value.ToString() ?? string.Empty;
            var strSerialFirstNumber = SubstrateChangeDataGridView.Rows[i].Cells[9].Value.ToString() ?? string.Empty;
            var strSerialLastNumber = SubstrateChangeDataGridView.Rows[i].Cells[10].Value.ToString() ?? string.Empty;
            var comment = SubstrateChangeDataGridView.Rows[i].Cells[12].Value.ToString() ?? string.Empty;
            var usedSubstrate = SubstrateChangeDataGridView.Rows[i].Cells[13].Value.ToString() ?? string.Empty;

            var intQuantity = Convert.ToInt32(SubstrateChangeDataGridView.Rows[i].Cells[5].Value);
            var intSerialLastNumber = Convert.ToInt32(SubstrateChangeDataGridView.Rows[i].Cells[11].Value);

            using SubstrateChange2 window = new();
            window.StrFontName = StrFontName;
            window.IntFontSize = IntFontSize;
            window.IntPrintType = IntPrintType;
            window.IntRegType = IntRegType;
            window.StrProductName = StrProductName;
            window.StrStockName = StrStockName;
            window.StrOrderNumber = orderNumber;
            window.StrProductNumber = productNumber;
            window.StrProductType = productType;
            window.StrProductModel = productModel;
            window.IntQuantity = intQuantity;
            window.StrRevision = revision;
            window.StrComment = comment;
            window.StrUseSubstrate = StrUseSubstrate;
            window.StrUsedSubstrate = usedSubstrate;
            window.StrSerialFirstNumber = strSerialFirstNumber;
            window.StrSerialLastNumber = strSerialLastNumber;
            window.IntSerialLastNumber = intSerialLastNumber;
            window.ShowDialog(this);
            Close();
        }

        private void SubstrateChange1_Load(object sender, EventArgs e) { LoadEvents(); }
        private void OKButton_Click(object sender, EventArgs e) { OpenSubstrateChangeWindow(); }
    }
}
