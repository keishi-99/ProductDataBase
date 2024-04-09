using System.Data;
using System.Data.SQLite;

namespace ProductDatabase {
    public partial class HistoryWindow : Form {

        public string StrFontName { get; set; } = "Meiryo UI";
        public int IntFontSize { get; set; } = 9;

        public DataTable DtHistoryTable { get; set; } = new();

        private readonly List<string> ListColFilter = new();

        public int IntRadioBtnFlg { get; set; }

        public string StrProductName { get; set; } = string.Empty;
        public string StrProductModel { get; set; } = string.Empty;
        public string StrProductType { get; set; } = string.Empty;
        public string StrSubstrateName { get; set; } = string.Empty;
        public string StrSubstrateModel { get; set; } = string.Empty;

        public HistoryWindow() {
            InitializeComponent();
        }

        private void LoadEvents() {
            Font = new Font(StrFontName, IntFontSize);

            switch (IntRadioBtnFlg) {
                case 1:
                    using (SQLiteConnection _con = new(MainWindow.GetConnectionString2())) {
                        DtHistoryTable.Clear();
                        using SQLiteDataAdapter _adapter = new($"SELECT _rowid_, * FROM 'Substrate_Reg_{StrProductName}' WHERE col_Substrate_Model = '{StrSubstrateModel}' ORDER BY _rowid_ DESC", _con);
                        _adapter.Fill(DtHistoryTable);

                        DataBaseDataGridView.DataSource = DtHistoryTable;

                        ListColFilter.Add("");
                        for (int _i = 0; _i < DataBaseDataGridView.ColumnCount; _i++) {
                            string _headerValue = DataBaseDataGridView.Columns[_i].HeaderCell.Value?.ToString() ?? string.Empty;
                            if (_headerValue != null) { ListColFilter.Add(_headerValue); }
                        }

                        // 最大サイズをディスプレイサイズに合わせる
                        if (Screen.PrimaryScreen != null) {
                            int h = Screen.PrimaryScreen.Bounds.Height;
                            int w = Screen.PrimaryScreen.Bounds.Width;
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
                        for (int _i = 0; _i < DataBaseDataGridView.ColumnCount; _i++) {
                            CategoryComboBox.Items.Add(DataBaseDataGridView.Columns[_i].HeaderCell.Value.ToString());
                        }
                    }
                    break;
                case 2:
                    using (SQLiteConnection _con = new(MainWindow.GetConnectionString2())) {
                        DtHistoryTable.Clear();
                        using SQLiteDataAdapter _adapter = new($"SELECT _rowid_, * FROM 'Product_Reg_{StrProductName}' WHERE col_Product_Model = '{StrProductModel}' ORDER BY _rowid_ DESC", _con);
                        _adapter.Fill(DtHistoryTable);

                        DataBaseDataGridView.DataSource = DtHistoryTable;

                        ListColFilter.Add("");
                        for (int _i = 0; _i < DataBaseDataGridView.ColumnCount; _i++) {
                            string _headerValue = DataBaseDataGridView.Columns[_i].HeaderCell.Value?.ToString() ?? string.Empty;
                            if (_headerValue != null) { ListColFilter.Add(_headerValue); }
                        }

                        // 最大サイズをディスプレイサイズに合わせる
                        if (Screen.PrimaryScreen != null) {
                            int h = Screen.PrimaryScreen.Bounds.Height;
                            int w = Screen.PrimaryScreen.Bounds.Width;
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
                        for (int _i = 0; _i < DataBaseDataGridView.ColumnCount; _i++) {
                            CategoryComboBox.Items.Add(DataBaseDataGridView.Columns[_i].HeaderCell.Value.ToString());
                        }
                    }
                    break;
                case 3:
                    using (SQLiteConnection _con = new(MainWindow.GetConnectionString2())) {
                        DtHistoryTable.Clear();
                        using SQLiteDataAdapter _adapter = new($"SELECT _rowid_, * FROM Reprint WHERE col_Product_Model = '{StrProductModel}' ORDER BY _rowid_ DESC", _con);
                        _adapter.Fill(DtHistoryTable);

                        DataBaseDataGridView.DataSource = DtHistoryTable;

                        ListColFilter.Add("");
                        for (int _i = 0; _i < DataBaseDataGridView.ColumnCount; _i++) {
                            string _headerValue = DataBaseDataGridView.Columns[_i].HeaderCell.Value?.ToString() ?? string.Empty;
                            if (_headerValue != null) { ListColFilter.Add(_headerValue); }
                        }

                        // 最大サイズをディスプレイサイズに合わせる
                        if (Screen.PrimaryScreen != null) {
                            int h = Screen.PrimaryScreen.Bounds.Height;
                            int w = Screen.PrimaryScreen.Bounds.Width;
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
                        for (int _i = 0; _i < DataBaseDataGridView.ColumnCount; _i++) {
                            CategoryComboBox.Items.Add(DataBaseDataGridView.Columns[_i].HeaderCell.Value.ToString());
                        }
                    }
                    break;
            }
        }

        private void HistoryTableFilter(object sender, EventArgs e) {
            try {
                if (string.IsNullOrEmpty(FilterStringTextBox.Text) || CategoryComboBox.SelectedIndex == 0) {
                    DtHistoryTable.DefaultView.RowFilter = null;
                }
                else if (CategoryComboBox.Text == "ID") {
                    if (int.TryParse(FilterStringTextBox.Text, out int id)) { DtHistoryTable.DefaultView.RowFilter = $"{ListColFilter[CategoryComboBox.SelectedIndex]} = '{id}'"; }
                }
                else if (CategoryComboBox.Text != "") {
                    DtHistoryTable.DefaultView.RowFilter = $"{ListColFilter[CategoryComboBox.SelectedIndex]} LIKE '*{FilterStringTextBox.Text}*'";
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message);
            }
        }

        private void HistoryWindow_Load(object sender, EventArgs e) { LoadEvents(); }
        private void FilterStringTextBox_TextChanged(object sender, EventArgs e) { HistoryTableFilter(sender, e); }
    }
}
