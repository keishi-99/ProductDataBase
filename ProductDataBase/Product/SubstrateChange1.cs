using System.Data;
using System.Data.SQLite;

namespace ProductDatabase {
    public partial class SubstrateChange1 : Form {

        public string StrFontName { get; set; } = "Meiryo UI";
        public int IntFontSize { get; set; } = 9;

        public DataTable DtHistoryTable { get; set; } = new();

        private readonly List<string> ListColFilter = [];

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

        public SubstrateChange1() => InitializeComponent();

        // ロードイベント
        private void LoadEvents() {
            try {
                Font = new Font(StrFontName, IntFontSize);

                switch (IntRadioBtnFlg) {
                    case 4:
                        using (SQLiteConnection _con = new(MainWindow.GetConnectionString2())) {
                            DtHistoryTable.Clear();
                            using SQLiteDataAdapter _adapter = new($"SELECT _rowid_, * FROM Product_Reg_{StrProductName} WHERE col_Product_Model = '{StrProductModel}' ORDER BY _rowid_ DESC", _con);
                            _adapter.Fill(DtHistoryTable);

                            SubstrateChangeDataGridView.DataSource = DtHistoryTable;

                            ListColFilter.Add("");
                            for (int _i = 0; _i < SubstrateChangeDataGridView.ColumnCount; _i++) {
                                string _headerValue = SubstrateChangeDataGridView.Columns[_i].HeaderCell.Value?.ToString() ?? string.Empty;
                                if (_headerValue != null) { ListColFilter.Add(_headerValue); }
                            }

                            // 最大サイズをディスプレイサイズに合わせる
                            if (Screen.PrimaryScreen != null) {
                                int _h = Screen.PrimaryScreen.Bounds.Height;
                                int _w = Screen.PrimaryScreen.Bounds.Width;
                                SubstrateChangeDataGridView.MaximumSize = new Size(_w, _h);
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

            int _i = SubstrateChangeDataGridView.SelectedCells[0].RowIndex;

            string _orderNumber = SubstrateChangeDataGridView.Rows[_i].Cells[1].Value.ToString() ?? string.Empty;
            string _productNumber = SubstrateChangeDataGridView.Rows[_i].Cells[2].Value.ToString() ?? string.Empty;
            string _productType = SubstrateChangeDataGridView.Rows[_i].Cells[3].Value.ToString() ?? string.Empty;
            string _productModel = SubstrateChangeDataGridView.Rows[_i].Cells[4].Value.ToString() ?? string.Empty;
            string _revision = SubstrateChangeDataGridView.Rows[_i].Cells[8].Value.ToString() ?? string.Empty;
            string _strSerialFirstNumber = SubstrateChangeDataGridView.Rows[_i].Cells[9].Value.ToString() ?? string.Empty;
            string _strSerialLastNumber = SubstrateChangeDataGridView.Rows[_i].Cells[10].Value.ToString() ?? string.Empty;
            string _comment = SubstrateChangeDataGridView.Rows[_i].Cells[12].Value.ToString() ?? string.Empty;
            string _usedSubstrate = SubstrateChangeDataGridView.Rows[_i].Cells[13].Value.ToString() ?? string.Empty;

            int _intQuantity = Convert.ToInt32(SubstrateChangeDataGridView.Rows[_i].Cells[5].Value);
            int _intSerialLastNumber = Convert.ToInt32(SubstrateChangeDataGridView.Rows[_i].Cells[11].Value);

            using SubstrateChange2 _window = new();
            _window.StrFontName = StrFontName;
            _window.IntFontSize = IntFontSize;
            _window.IntPrintType = IntPrintType;
            _window.IntRegType = IntRegType;
            _window.StrProductName = StrProductName;
            _window.StrStockName = StrStockName;
            _window.StrOrderNumber = _orderNumber;
            _window.StrProductNumber = _productNumber;
            _window.StrProductType = _productType;
            _window.StrProductModel = _productModel;
            _window.IntQuantity = _intQuantity;
            _window.StrRevision = _revision;
            _window.StrComment = _comment;
            _window.StrUseSubstrate = StrUseSubstrate;
            _window.StrUsedSubstrate = _usedSubstrate;
            _window.StrSerialFirstNumber = _strSerialFirstNumber;
            _window.StrSerialLastNumber = _strSerialLastNumber;
            _window.IntSerialLastNumber = _intSerialLastNumber;
            _window.ShowDialog(this);
            Close();
        }

        private void SubstrateChange1_Load(object sender, EventArgs e) { LoadEvents(); }
        private void OKButton_Click(object sender, EventArgs e) { OpenSubstrateChangeWindow(); }
    }
}
