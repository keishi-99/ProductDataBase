using System.Data;
using System.Data.SQLite;
using static ProductDatabase.MainWindow;

namespace ProductDatabase {
    public partial class HistoryWindow : Form {

        public ProductInfomation ProductInfo { get; set; } = new ProductInfomation();

        private readonly DataTable _historyTable = new();

        private readonly List<string> _listColFilter = [];

        public HistoryWindow() {
            InitializeComponent();
        }

        private void LoadEvents() {
            try {
                Font = new Font(ProductInfo.FontName, ProductInfo.FontSize);

                switch (ProductInfo.RadioButtonFlg) {
                    case 1:
                        using (SQLiteConnection con = new(GetConnectionString2())) {
                            _historyTable.Clear();

                            var query = $"SELECT _rowid_, * FROM Substrate_Reg_{ProductInfo.ProductName} WHERE col_Substrate_Model = @col_Substrate_Model ORDER BY _rowid_ DESC";
                            using SQLiteCommand command = new(query, con);
                            command.Parameters.AddWithValue("@col_Substrate_Model", ProductInfo.SubstrateModel);
                            // SQLiteDataAdapterのインスタンス化
                            using SQLiteDataAdapter adapter = new(command);

                            // データの取得とDataTableへの格納
                            adapter.Fill(_historyTable);

                            DataBaseDataGridView.DataSource = _historyTable;

                            _listColFilter.Add("");
                            for (var i = 0; i < DataBaseDataGridView.ColumnCount; i++) {
                                var headerValue = DataBaseDataGridView.Columns[i].HeaderCell.Value?.ToString() ?? String.Empty;
                                if (headerValue != null) { _listColFilter.Add(headerValue); }
                            }

                            // 最大サイズをディスプレイサイズに合わせる
                            if (Screen.PrimaryScreen != null) {
                                var h = Screen.PrimaryScreen.Bounds.Height;
                                var w = Screen.PrimaryScreen.Bounds.Width;
                                DataBaseDataGridView.MaximumSize = new Size(w, h);
                            }
                            DataBaseDataGridView.ColumnHeadersDefaultCellStyle.Font = new Font(DataBaseDataGridView.Font, FontStyle.Bold);
                            DataBaseDataGridView.RowHeadersVisible = true;
                            DataBaseDataGridView.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                            DataBaseDataGridView.AlternatingRowsDefaultCellStyle.BackColor = Color.Gainsboro;
                            //ヘッダーとすべてのセルの内容に合わせて、行の高さを自動調整する
                            DataBaseDataGridView.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
                            DataBaseDataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;

                            DataBaseDataGridView.Columns[0].HeaderCell.Value = "ID";
                            DataBaseDataGridView.Columns[0].Width = 40;
                            DataBaseDataGridView.Columns[1].HeaderCell.Value = "基板名";
                            DataBaseDataGridView.Columns[2].HeaderCell.Value = "基板型式";
                            DataBaseDataGridView.Columns[3].HeaderCell.Value = "製造番号";
                            DataBaseDataGridView.Columns[3].Width = 130;
                            DataBaseDataGridView.Columns[4].HeaderCell.Value = "注文番号";
                            DataBaseDataGridView.Columns[5].HeaderCell.Value = "追加量";
                            DataBaseDataGridView.Columns[5].Width = 70;
                            DataBaseDataGridView.Columns[6].HeaderCell.Value = "減少量";
                            DataBaseDataGridView.Columns[6].Width = 70;
                            DataBaseDataGridView.Columns[7].HeaderCell.Value = "不良";
                            DataBaseDataGridView.Columns[7].Width = 40;
                            DataBaseDataGridView.Columns[8].HeaderCell.Value = "使用製品名";
                            DataBaseDataGridView.Columns[9].HeaderCell.Value = "使用製番";
                            DataBaseDataGridView.Columns[9].Width = 130;
                            DataBaseDataGridView.Columns[10].HeaderCell.Value = "使用注番";
                            DataBaseDataGridView.Columns[11].HeaderCell.Value = "Rev";
                            DataBaseDataGridView.Columns[11].Width = 40;
                            DataBaseDataGridView.Columns[12].HeaderCell.Value = "担当者";
                            DataBaseDataGridView.Columns[12].Width = 70;
                            DataBaseDataGridView.Columns[13].HeaderCell.Value = "登録日";
                            DataBaseDataGridView.Columns[13].Width = 80;
                            DataBaseDataGridView.Columns[14].HeaderCell.Value = "コメント";

                            CategoryComboBox.Items.Add("");
                            for (var i = 0; i < DataBaseDataGridView.ColumnCount; i++) {
                                CategoryComboBox.Items.Add(DataBaseDataGridView.Columns[i].HeaderCell.Value.ToString() ?? String.Empty);
                            }
                        }
                        break;
                    case 2:
                        using (SQLiteConnection con = new(GetConnectionString2())) {
                            _historyTable.Clear();

                            var query = $"SELECT _rowid_, * FROM Product_Reg_{ProductInfo.ProductName} WHERE col_Product_Model = @col_Product_Model ORDER BY _rowid_ DESC";
                            using SQLiteCommand command = new(query, con);
                            command.Parameters.AddWithValue("@col_Product_Model", ProductInfo.ProductModel);
                            // SQLiteDataAdapterのインスタンス化
                            using SQLiteDataAdapter adapter = new(command);

                            // データの取得とDataTableへの格納
                            adapter.Fill(_historyTable);

                            DataBaseDataGridView.DataSource = _historyTable;

                            _listColFilter.Add("");
                            for (var i = 0; i < DataBaseDataGridView.ColumnCount; i++) {
                                var headerValue = DataBaseDataGridView.Columns[i].HeaderCell.Value?.ToString() ?? String.Empty;
                                if (headerValue != null) { _listColFilter.Add(headerValue); }
                            }

                            // 最大サイズをディスプレイサイズに合わせる
                            if (Screen.PrimaryScreen != null) {
                                var h = Screen.PrimaryScreen.Bounds.Height;
                                var w = Screen.PrimaryScreen.Bounds.Width;
                                DataBaseDataGridView.MaximumSize = new Size(w, h);
                            }
                            DataBaseDataGridView.ColumnHeadersDefaultCellStyle.Font = new Font(DataBaseDataGridView.Font, FontStyle.Bold);
                            DataBaseDataGridView.RowHeadersVisible = true;
                            DataBaseDataGridView.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                            DataBaseDataGridView.AlternatingRowsDefaultCellStyle.BackColor = Color.Gainsboro;
                            //ヘッダーとすべてのセルの内容に合わせて、行の高さを自動調整する
                            DataBaseDataGridView.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
                            DataBaseDataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;

                            DataBaseDataGridView.Columns[0].HeaderCell.Value = "ID";
                            DataBaseDataGridView.Columns[0].Width = 40;
                            DataBaseDataGridView.Columns[1].HeaderCell.Value = "注文番号";
                            DataBaseDataGridView.Columns[2].HeaderCell.Value = "製造番号";
                            DataBaseDataGridView.Columns[2].Width = 130;
                            DataBaseDataGridView.Columns[3].HeaderCell.Value = "製品名";
                            DataBaseDataGridView.Columns[4].HeaderCell.Value = "製品型式";
                            DataBaseDataGridView.Columns[5].HeaderCell.Value = "数量";
                            DataBaseDataGridView.Columns[5].Width = 40;
                            DataBaseDataGridView.Columns[6].HeaderCell.Value = "担当者";
                            DataBaseDataGridView.Columns[6].Width = 70;
                            DataBaseDataGridView.Columns[7].HeaderCell.Value = "登録日";
                            DataBaseDataGridView.Columns[7].Width = 80;
                            DataBaseDataGridView.Columns[8].HeaderCell.Value = "Rev";
                            DataBaseDataGridView.Columns[8].Width = 40;
                            DataBaseDataGridView.Columns[9].HeaderCell.Value = "シリアル先頭";
                            DataBaseDataGridView.Columns[10].HeaderCell.Value = "シリアル末尾";
                            DataBaseDataGridView.Columns[11].HeaderCell.Value = "末番";
                            DataBaseDataGridView.Columns[11].Width = 40;
                            DataBaseDataGridView.Columns[12].HeaderCell.Value = "コメント";
                            DataBaseDataGridView.Columns[13].HeaderCell.Value = "使用基板";
                            DataBaseDataGridView.Columns[13].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                            //DataBaseDataGridView.Columns[13].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;

                            CategoryComboBox.Items.Add("");
                            for (var i = 0; i < DataBaseDataGridView.ColumnCount; i++) {
                                CategoryComboBox.Items.Add(DataBaseDataGridView.Columns[i].HeaderCell.Value.ToString() ?? String.Empty);
                            }
                        }
                        break;
                    case 3:
                        using (SQLiteConnection con = new(GetConnectionString2())) {
                            _historyTable.Clear();

                            var query = $"SELECT _rowid_, * FROM Reprint WHERE col_Product_Model = @col_Product_Model ORDER BY _rowid_ DESC";
                            using SQLiteCommand command = new(query, con);
                            command.Parameters.AddWithValue("@col_Product_Model", ProductInfo.ProductModel);
                            // SQLiteDataAdapterのインスタンス化
                            using SQLiteDataAdapter adapter = new(command);

                            // データの取得とDataTableへの格納
                            adapter.Fill(_historyTable);

                            DataBaseDataGridView.DataSource = _historyTable;

                            _listColFilter.Add("");
                            for (var i = 0; i < DataBaseDataGridView.ColumnCount; i++) {
                                var headerValue = DataBaseDataGridView.Columns[i].HeaderCell.Value?.ToString() ?? String.Empty;
                                if (headerValue != null) { _listColFilter.Add(headerValue); }
                            }

                            // 最大サイズをディスプレイサイズに合わせる
                            if (Screen.PrimaryScreen != null) {
                                var h = Screen.PrimaryScreen.Bounds.Height;
                                var w = Screen.PrimaryScreen.Bounds.Width;
                                DataBaseDataGridView.MaximumSize = new Size(w, h);
                            }
                            DataBaseDataGridView.ColumnHeadersDefaultCellStyle.Font = new Font(DataBaseDataGridView.Font, FontStyle.Bold);
                            DataBaseDataGridView.RowHeadersVisible = true;
                            DataBaseDataGridView.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                            DataBaseDataGridView.AlternatingRowsDefaultCellStyle.BackColor = Color.Gainsboro;
                            //ヘッダーとすべてのセルの内容に合わせて、行の高さを自動調整する
                            DataBaseDataGridView.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
                            DataBaseDataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;

                            DataBaseDataGridView.Columns[0].HeaderCell.Value = "ID";
                            DataBaseDataGridView.Columns[0].Width = 40;
                            DataBaseDataGridView.Columns[1].HeaderCell.Value = "印刷対象";
                            DataBaseDataGridView.Columns[2].HeaderCell.Value = "注文番号";
                            DataBaseDataGridView.Columns[3].HeaderCell.Value = "製造番号";
                            DataBaseDataGridView.Columns[3].Width = 130;
                            DataBaseDataGridView.Columns[4].HeaderCell.Value = "製品名";
                            DataBaseDataGridView.Columns[5].HeaderCell.Value = "製品型式";
                            DataBaseDataGridView.Columns[6].HeaderCell.Value = "数量";
                            DataBaseDataGridView.Columns[6].Width = 70;
                            DataBaseDataGridView.Columns[7].HeaderCell.Value = "担当者";
                            DataBaseDataGridView.Columns[7].Width = 70;
                            DataBaseDataGridView.Columns[8].HeaderCell.Value = "登録日";
                            DataBaseDataGridView.Columns[8].Width = 80;
                            DataBaseDataGridView.Columns[9].HeaderCell.Value = "Rev";
                            DataBaseDataGridView.Columns[9].Width = 40;
                            DataBaseDataGridView.Columns[10].HeaderCell.Value = "シリアル末尾";
                            DataBaseDataGridView.Columns[11].HeaderCell.Value = "末番";
                            DataBaseDataGridView.Columns[12].HeaderCell.Value = "コメント";

                            CategoryComboBox.Items.Add("");
                            for (var i = 0; i < DataBaseDataGridView.ColumnCount; i++) {
                                CategoryComboBox.Items.Add(DataBaseDataGridView.Columns[i].HeaderCell.Value.ToString() ?? String.Empty);
                            }
                        }
                        break;
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message);
            }
        }

        private void HistoryTableFilter(object sender, EventArgs e) {
            try {
                if (String.IsNullOrEmpty(FilterStringTextBox.Text) || CategoryComboBox.SelectedIndex == 0) {
                    _historyTable.DefaultView.RowFilter = null;
                }
                else if (CategoryComboBox.Text == "ID") {
                    if (Int32.TryParse(FilterStringTextBox.Text, out var id)) { _historyTable.DefaultView.RowFilter = $"{_listColFilter[CategoryComboBox.SelectedIndex]} = '{id}'"; }
                }
                else if (CategoryComboBox.Text != "") {
                    _historyTable.DefaultView.RowFilter = $"{_listColFilter[CategoryComboBox.SelectedIndex]} LIKE '*{FilterStringTextBox.Text}*'";
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message);
            }
        }

        private void HistoryWindow_Load(object sender, EventArgs e) { LoadEvents(); }
        private void FilterStringTextBox_TextChanged(object sender, EventArgs e) { HistoryTableFilter(sender, e); }
    }
}
