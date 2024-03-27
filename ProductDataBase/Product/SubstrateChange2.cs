using DocumentFormat.OpenXml.Office.Word;
using System.Data.SQLite;

namespace ProductDatabase {
    public partial class SubstrateChange2 : Form {

        public string StrFontName { get; set; } = "Meiryo UI";
        public int IntFontSize { get; set; } = 9;

        public string? StrProductName { get; set; } = string.Empty;
        public string StrStockName { get; set; } = string.Empty;
        public string StrProductType { get; set; } = string.Empty;
        public string StrProductModel { get; set; } = string.Empty;
        public string StrUseSubstrate { get; set; } = string.Empty;
        public string[] ArrUseSubstrate = Array.Empty<string>();
        public string StrUsedSubstrate { get; set; } = string.Empty;
        public string[] ArrUsedSubstrate = Array.Empty<string>();
        public string StrInitial { get; set; } = string.Empty;
        public string StrOrderNumber { get; set; } = string.Empty;
        public string StrProductNumber { get; set; } = string.Empty;
        public string StrRegDate { get; set; } = string.Empty;
        public string StrPerson { get; set; } = string.Empty;
        public string StrRevision { get; set; } = string.Empty;
        public string StrComment { get; set; } = string.Empty;

        public int IntQuantity { get; set; }
        public int IntRegType { get; set; }
        public int IntPrintType { get; set; }
        public int IntCheckBin { get; set; }
        public int IntSerialDigit { get; set; }
        public int IntSerialFirstNumber { get; set; }

        readonly List<string> CheckBoxNames = new() {
                        "Substrate1CheckBox", "Substrate2CheckBox", "Substrate3CheckBox", "Substrate4CheckBox","Substrate5CheckBox",
                        "Substrate6CheckBox", "Substrate7CheckBox", "Substrate8CheckBox", "Substrate9CheckBox","Substrate10CheckBox"
                        };
        readonly List<string> DataGridViewNames = new() {
                        "Substrate1DataGridView", "Substrate2DataGridView", "Substrate3DataGridView", "Substrate4DataGridView","Substrate5DataGridView",
                        "Substrate6DataGridView", "Substrate7DataGridView", "Substrate8DataGridView", "Substrate9DataGridView","Substrate10DataGridView"
                        };

        public SubstrateChange2() => InitializeComponent();

        // ロードイベント
        private void LoadEvents() {
            try {
                Font = new Font(StrFontName, IntFontSize);

                // テキストボックスに入力
                OrderNumberTextBox.Text = StrOrderNumber;
                ManufacturingNumberTextBox.Text = StrProductNumber;
                QuantityTextBox.Text = IntQuantity.ToString();
                DateTime _dtNow = DateTime.Now;
                RegistrationDateMaskedTextBox.Text = _dtNow.ToShortDateString();
                RevisionTextBox.Text = StrRevision;
                CommentTextBox.Text = StrComment;

                // DB1へ接続し担当者取得
                using (SQLiteConnection _con = new(MainWindow.GetConnectionString1())) {
                    _con.Open();
                    using SQLiteCommand _cmd = _con.CreateCommand();
                    // テーブル検索SQL - 担当者をComboboxへ追加
                    _cmd.CommandText = "SELECT * FROM Person ORDER BY _rowid_ ASC";
                    using SQLiteDataReader _dr = _cmd.ExecuteReader();
                    while (_dr.Read()) {
                        PersonComboBox.Items.Add($"{_dr["col_Person_Name"]}");
                    }
                }

                ArrUseSubstrate = StrUseSubstrate.Split(",");
                ArrUsedSubstrate = StrUsedSubstrate.Split(",");
                string _strQuantity = string.Empty;
                string _strSubstrateName = string.Empty;

                switch (IntRegType) {
                    case 2:
                        for (int _i = 0; _i <= ArrUseSubstrate.GetUpperBound(0); _i++) {
                            int _quantity = IntQuantity;

                            CheckBox? _objCbx = Controls[CheckBoxNames[_i]] as CheckBox;
                            if (_objCbx != null) {
                                _objCbx.Enabled = true;
                                _objCbx.Checked = true;
                            }

                            DataGridView? _objDgv = Controls[DataGridViewNames[_i]] as DataGridView;
                            if (_objDgv != null) {
                                _objDgv.RowHeadersWidth = 30;
                                _objDgv.Columns[2].DefaultCellStyle.BackColor = Color.LightGray;
                                _objDgv.Columns[2].ReadOnly = true;
                                _objDgv.Columns[3].ReadOnly = false;
                                _objDgv.Columns[4].ReadOnly = false;
                            }

                            using SQLiteConnection _con = new(MainWindow.GetConnectionString2());
                            _con.Open();

                            using SQLiteCommand _cmd = _con.CreateCommand();
                            // テーブル検索SQL - [Product_Name]_Stockテーブルを在庫数フラグ有&基板型式[Model]で抽出して取得
                            _cmd.CommandText = $"SELECT * FROM 'Stock_{StrStockName}' WHERE col_Substrate_Model = '{ArrUseSubstrate[_i]}' ORDER BY _rowid_ ASC";
                            using SQLiteDataReader _dr = _cmd.ExecuteReader();
                            int _j = 0;
                            while (_dr.Read()) {
                                string _strUsedSubNum = string.Empty;
                                string _strUsedQuantity = string.Empty;
                                int _intUsedQuantity = 0;

                                // 抽出した行から製造番号,在庫取得
                                string _strSubstrateNum = $"{_dr["col_Substrate_num"]}";
                                int _intStock = int.Parse($"{_dr["col_Stock"]}");
                                int _colFlg = int.Parse($"{_dr["col_Flg"]}");
                                _strSubstrateName = $"{_dr["col_Substrate_Name"]}";
                                if (_objCbx != null) { _objCbx.Text = $"{_strSubstrateName} - {ArrUseSubstrate[_i]}"; }

                                int _subIndex = 0;
                                bool _b1 = false;

                                foreach (var _d in ArrUsedSubstrate) {
                                    _b1 = _d.Contains(_strSubstrateNum);
                                    if (_b1 == true)
                                        break;
                                    _subIndex++;
                                }
                                if (_b1 == true) {
                                    _strUsedSubNum = ArrUsedSubstrate[_subIndex];
                                    _strUsedSubNum = _strUsedSubNum[(_strUsedSubNum.IndexOf("]") + 1)..];
                                    _strUsedSubNum = _strUsedSubNum[.._strUsedSubNum.IndexOf("(")];

                                    _strUsedQuantity = ArrUsedSubstrate[_subIndex];
                                    _strUsedQuantity = _strUsedQuantity[(_strUsedQuantity.IndexOf("(") + 1)..];
                                    _strUsedQuantity = _strUsedQuantity[.._strUsedQuantity.IndexOf(")")];
                                    _intUsedQuantity = int.Parse(_strUsedQuantity);
                                }

                                if (_colFlg == 1) {
                                    if (_objDgv == null) { break; }
                                    _objDgv.Rows.Add();
                                    _objDgv.Rows[_j].Cells[0].Value = _strSubstrateNum;
                                    _objDgv.Rows[_j].Cells[1].Value = _intStock;

                                    if (_intUsedQuantity != 0) {
                                        _objDgv.Rows[_j].Cells[2].Value = _intUsedQuantity;
                                        _objDgv.Rows[_j].Cells[3].Value = _intUsedQuantity;
                                        _objDgv.Rows[_j].Cells[4].Value = true;
                                    }
                                    else {
                                        _objDgv.Rows[_j].Cells[2].Value = 0;
                                        _objDgv.Rows[_j].Cells[3].Value = 0;
                                    }
                                    _j++;
                                }
                                else if (_strUsedSubNum == _strSubstrateNum) {
                                    if (_objDgv == null) { break; }
                                    _objDgv.Rows.Add();
                                    _objDgv.Rows[_j].Cells[0].Value = _strSubstrateNum;
                                    _objDgv.Rows[_j].Cells[1].Value = _intStock;

                                    if (_intUsedQuantity != 0) {
                                        _objDgv.Rows[_j].Cells[2].Value = _intUsedQuantity;
                                        _objDgv.Rows[_j].Cells[3].Value = _intUsedQuantity;
                                        _objDgv.Rows[_j].Cells[4].Value = true;
                                    }
                                    else {
                                        _objDgv.Rows[_j].Cells[2].Value = 0;
                                        _objDgv.Rows[_j].Cells[3].Value = 0;
                                    }
                                    _j++;
                                }
                            }
                        }
                        break;
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } finally {
            }
        }
        // 変更登録
        private void ChangeRegistration() {

        }
        // リスト印刷
        private void ListPrint() {
        }

        // チェックボックスイベント
        private void CheckBox_CheckedChanged(object sender, EventArgs e) {
            CheckBox _checkBox = (CheckBox)sender;
            DataGridView _dataGridView = new();

            switch (_checkBox.Name) {
                case "Substrate1CheckBox":
                    _dataGridView = Substrate1DataGridView;
                    break;
                case "Substrate2CheckBox":
                    _dataGridView = Substrate2DataGridView;
                    break;
                case "Substrate3CheckBox":
                    _dataGridView = Substrate3DataGridView;
                    break;
                case "Substrate4CheckBox":
                    _dataGridView = Substrate4DataGridView;
                    break;
                case "Substrate5CheckBox":
                    _dataGridView = Substrate5DataGridView;
                    break;
                case "Substrate6CheckBox":
                    _dataGridView = Substrate6DataGridView;
                    break;
                case "Substrate7CheckBox":
                    _dataGridView = Substrate7DataGridView;
                    break;
                case "Substrate8CheckBox":
                    _dataGridView = Substrate8DataGridView;
                    break;
                case "Substrate9CheckBox":
                    _dataGridView = Substrate9DataGridView;
                    break;
                case "Substrate10CheckBox":
                    _dataGridView = Substrate10DataGridView;
                    break;
                default:
                    break;
            }

            _dataGridView.Enabled = _checkBox.Checked;
            _checkBox.ForeColor = _checkBox.Checked ? Color.Black : Color.Red;

            if (!_checkBox.Checked) {
                MessageBox.Show("チェックがない場合在庫から引き落とされなくなります。", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }


        private void SubstrateChange2_Load(object sender, EventArgs e) { LoadEvents(); }
        private void RegisterButton_Click(object sender, EventArgs e) { ChangeRegistration(); }
        private void SubstrateListPrintButton_Click(object sender, EventArgs e) { ListPrint(); }
        private void CloseButton_Click(object sender, EventArgs e) { Close(); }
    }
}
