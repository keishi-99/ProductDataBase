using System.Data;
using System.Data.SQLite;
using System.Windows.Forms;

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
                        DataBaseDataGridView.Columns[13].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;

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
                    ExportCsvButton.Enabled = false;
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

        private void ExportCsv() {
            string _name = string.Empty;
            switch (IntRadioBtnFlg) {
                case 1:
                    _name = $"{StrProductName} - {StrSubstrateName}";
                    break;
                case 2:
                    _name = $"{StrProductName} - {StrProductModel}";
                    break;
                case 3:
                    return;
            }

            using SaveFileDialog _sfd = new() {
                FileName = $"{_name}.csv",
                Filter = "すべてのファイル(*.*)|*.*",
                Title = "名前を付けて保存",
                RestoreDirectory = true
            };

            if (_sfd.ShowDialog() == DialogResult.OK) {
                CnvDataTableToCsv(DtHistoryTable, _sfd.FileName, true);
            }

        }
        public void CnvDataTableToCsv(DataTable dt, string csvPath, bool WriteHeader) {
            System.Text.Encoding _encoding = System.Text.Encoding.GetEncoding("Shift_JIS");

            // 書き込むファイルを開く
            using StreamWriter _sr = new(csvPath, false, _encoding);
            string[] _arr = Array.Empty<string>();
            int _colCount = 0;
            int _lastColIndex = 0;

            // ヘッダを書き込む
            if (WriteHeader) {
                switch (IntRadioBtnFlg) {
                    case 1:
                        _arr = new string[] { "ID", "基板名", "基板型式", "製造番号", "注文番号", "追加量", "減少量", "不良", "使用製品名", "使用製番", "使用注番", "Revision", "担当者", "登録日", "コメント" };
                        break;
                    case 2:
                        _arr = new string[] { "ID", "注文番号", "製造番号", "製品名", "製品型式", "数量", "担当者", "登録日", "Revision", "シリアル先頭", "シリアル末尾", "末番", "コメント", "使用基板" };
                        break;
                }
                List<string> _colName = new();
                _colName.AddRange(_arr);
                _colCount = _colName.Count;
                _lastColIndex = _colCount - 1;
                for (int _i = 0; _i < _colCount; _i++) {
                    string _field = _colName[_i];
                    _field = EncloseDoubleQuotes(_field);
                    _sr.Write(_field);
                    // カンマ付与
                    if (_lastColIndex > _i) _sr.Write(',');
                }
                // 改行
                _sr.Write(Environment.NewLine);
            }

            // レコードを書き込む
            foreach (DataRow _row in dt.Rows) {
                for (int _i = 0; _i < _colCount; _i++) {
                    // フィールドの取得
                    string _field = _row[_i]?.ToString() ?? string.Empty; // nullチェックを追加
                    if (_field != null) {
                        // "で囲み書き込む
                        _field = EncloseDoubleQuotes(_field);
                        _sr.Write(_field);
                        // カンマ付与
                        if (_lastColIndex > _i) {
                            _sr.Write(',');
                        }
                    }
                }
                _sr.Write(Environment.NewLine);
            }

        }
        private static string EncloseDoubleQuotes(string field) { return $"\"{field}\""; }

        private void HistoryWindow_Load(object sender, EventArgs e) { LoadEvents(); }
        private void ExportCsvButton_Click(object sender, EventArgs e) { ExportCsv(); }
        private void FilterStringTextBox_TextChanged(object sender, EventArgs e) { HistoryTableFilter(sender, e); }
    }
}
