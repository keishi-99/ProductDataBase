using ClosedXML.Excel;
using System.Data;
using System.Data.SQLite;
using Excel = Microsoft.Office.Interop.Excel;

namespace ProductDatabase {
    public partial class SubstrateChange2 : Form {

        public string StrFontName { get; set; } = "Meiryo UI";
        public int IntFontSize { get; set; } = 9;

        public string? StrProductName { get; set; } = string.Empty;
        public string StrStockName { get; set; } = string.Empty;
        public string StrProductType { get; set; } = string.Empty;
        public string StrProductModel { get; set; } = string.Empty;
        public string StrUseSubstrate { get; set; } = string.Empty;
        public string[] ArrUseSubstrate = [];
        public string StrUsedSubstrate { get; set; } = string.Empty;
        public string[] ArrUsedSubstrate = [];
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
        public int IntSerialLastNumber { get; set; }

        public string StrSerialFirstNumber { get; set; } = string.Empty;
        public string StrSerialLastNumber { get; set; } = string.Empty;
        public string StrTotalSubstrate { get; set; } = string.Empty;

        private readonly List<string> ListUsedSubstrate = [];
        private readonly List<string> ListUsedProductNumber = [];
        private readonly List<int> ListUsedQuantity = [];
        private readonly List<string> CheckBoxNames = [
                        "Substrate1CheckBox", "Substrate2CheckBox", "Substrate3CheckBox", "Substrate4CheckBox","Substrate5CheckBox",
                        "Substrate6CheckBox", "Substrate7CheckBox", "Substrate8CheckBox", "Substrate9CheckBox","Substrate10CheckBox"
                        ];
        private readonly List<string> DataGridViewNames = [
                        "Substrate1DataGridView", "Substrate2DataGridView", "Substrate3DataGridView", "Substrate4DataGridView","Substrate5DataGridView",
                        "Substrate6DataGridView", "Substrate7DataGridView", "Substrate8DataGridView", "Substrate9DataGridView","Substrate10DataGridView"
                        ];
        private CheckBox? ObjCbx;
        private DataGridView? ObjDgv;

        public SubstrateChange2() => InitializeComponent();

        // ロードイベント
        private void LoadEvents() {
            try {
                Font = new Font(StrFontName, IntFontSize);
                CloseButton.Enabled = true;

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

                            // チェックボックスとDgvを有効に
                            ObjCbx = Controls[CheckBoxNames[_i]] as CheckBox;
                            if (ObjCbx != null) {
                                ObjCbx.Enabled = true;
                                ObjCbx.Checked = true;
                            }

                            ObjDgv = Controls[DataGridViewNames[_i]] as DataGridView;
                            if (ObjDgv != null) {
                                ObjDgv.RowHeadersWidth = 30;
                                ObjDgv.Columns[2].DefaultCellStyle.BackColor = Color.LightGray;
                                ObjDgv.Columns[2].ReadOnly = true;
                                ObjDgv.Columns[3].ReadOnly = false;
                                ObjDgv.Columns[4].ReadOnly = false;
                            }

                            using SQLiteConnection _con = new(MainWindow.GetConnectionString2());
                            _con.Open();

                            using SQLiteCommand _cmd = _con.CreateCommand();
                            // テーブル検索SQL - [Product_Name]_Stockテーブルを在庫数フラグ有&基板型式[Model]で抽出して取得
                            _cmd.CommandText = $"SELECT * FROM Stock_{StrStockName} WHERE col_Substrate_Model = @col_Substrate_Model ORDER BY _rowid_ ASC";
                            _cmd.Parameters.Add("@col_Substrate_Model", DbType.String).Value = ArrUseSubstrate[_i];
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
                                if (ObjCbx != null) { ObjCbx.Text = $"{_strSubstrateName} - {ArrUseSubstrate[_i]}"; }

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
                                    _strUsedSubNum = _strUsedSubNum[(_strUsedSubNum.IndexOf(']') + 1)..];
                                    _strUsedSubNum = _strUsedSubNum[.._strUsedSubNum.IndexOf('(')];

                                    _strUsedQuantity = ArrUsedSubstrate[_subIndex];
                                    _strUsedQuantity = _strUsedQuantity[(_strUsedQuantity.IndexOf('(') + 1)..];
                                    _strUsedQuantity = _strUsedQuantity[.._strUsedQuantity.IndexOf(')')];
                                    _intUsedQuantity = int.Parse(_strUsedQuantity);
                                }

                                if (_colFlg == 1) {
                                    if (ObjDgv == null) { break; }
                                    ObjDgv.Rows.Add();
                                    ObjDgv.Rows[_j].Cells[0].Value = _strSubstrateNum;
                                    ObjDgv.Rows[_j].Cells[1].Value = _intStock;

                                    if (_intUsedQuantity != 0) {
                                        ObjDgv.Rows[_j].Cells[2].Value = _intUsedQuantity;
                                        ObjDgv.Rows[_j].Cells[3].Value = _intUsedQuantity;
                                        ObjDgv.Rows[_j].Cells[4].Value = true;
                                    }
                                    else {
                                        ObjDgv.Rows[_j].Cells[2].Value = 0;
                                        ObjDgv.Rows[_j].Cells[3].Value = 0;
                                    }
                                    _j++;
                                }
                                else if (_strUsedSubNum == _strSubstrateNum) {
                                    if (ObjDgv == null) { break; }
                                    ObjDgv.Rows.Add();
                                    ObjDgv.Rows[_j].Cells[0].Value = _strSubstrateNum;
                                    ObjDgv.Rows[_j].Cells[1].Value = _intStock;

                                    if (_intUsedQuantity != 0) {
                                        ObjDgv.Rows[_j].Cells[2].Value = _intUsedQuantity;
                                        ObjDgv.Rows[_j].Cells[3].Value = _intUsedQuantity;
                                        ObjDgv.Rows[_j].Cells[4].Value = true;
                                    }
                                    else {
                                        ObjDgv.Rows[_j].Cells[2].Value = 0;
                                        ObjDgv.Rows[_j].Cells[3].Value = 0;
                                    }
                                    _j++;
                                }
                            }
                        }
                        break;
                }

                RegisterButton.Enabled = true;

            } catch (Exception ex) {
                MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } finally {
            }
        }
        // 変更登録
        private void ChangeRegistration() {
            try {
                if (!QuantityCheck()) { return; }
                if (!Registration()) { throw new Exception("登録失敗しました。"); }

                MessageBox.Show("登録完了");

                // フォームを編集不可にする
                RegistrationDateMaskedTextBox.Enabled = false;
                PersonComboBox.Enabled = false;
                RegisterButton.Enabled = false;
                for (int _i = 0; _i <= 9; _i++) {
                    ObjCbx = Controls[CheckBoxNames[_i]] as CheckBox;
                    if (ObjCbx != null) {
                        ObjCbx.Enabled = false;
                    }

                    ObjDgv = Controls[DataGridViewNames[_i]] as DataGridView;
                    if (ObjDgv != null) {
                        ObjDgv.Enabled = false;
                    }
                }
                // リスト印刷ボタンを有効に
                if (IntPrintType == 5 || IntPrintType == 6) {
                    SubstrateListPrintButton.Enabled = true;
                }

            } catch (Exception ex) {
                MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private bool QuantityCheck() {
            try {
                switch (IntRegType) {
                    case 2:
                    case 3:
                        if (ArrUseSubstrate == null) { throw new Exception("ArrUseSubstrateが空です"); }
                        for (int _i = 0; _i <= ArrUseSubstrate.GetUpperBound(0); _i++) {

                            CheckBox? _objCbx = Controls[CheckBoxNames[_i]] as CheckBox ?? throw new Exception("objCbxがnullです。");
                            _objCbx.Enabled = true;
                            _objCbx.Checked = true;

                            DataGridView? _objDgv = Controls[DataGridViewNames[_i]] as DataGridView ?? throw new Exception("objDgvがnullです。");
                            _objDgv.Columns[3].ReadOnly = false;
                            _objDgv.Columns[4].ReadOnly = false;

                            if (_objCbx.Checked) {
                                int _intQuantityCheck = IntQuantity;
                                int _dgvRowCnt = _objDgv.Rows.Count;

                                for (int _j = 0; _j < _dgvRowCnt; _j++) {
                                    bool _boolCbx = _objDgv.Rows[_j].Cells[4].Value != null && (bool)_objDgv.Rows[_j].Cells[4].Value;
                                    if (_boolCbx) {
                                        int _stockValue = Convert.ToInt32(_objDgv.Rows[_j].Cells[1].Value.ToString());
                                        int _usedValue = Convert.ToInt32(_objDgv.Rows[_j].Cells[2].Value.ToString());
                                        int _useValue = Convert.ToInt32(_objDgv.Rows[_j].Cells[3].Value.ToString());

                                        if (_stockValue + _usedValue < _useValue) {
                                            throw new Exception("在庫より多い数量が入力されています。");
                                        }
                                        _intQuantityCheck -= _useValue;
                                    }
                                }

                                if (_intQuantityCheck != 0) {
                                    throw new Exception("入力された数量の合計が必要数と一致しません。");
                                }
                            }
                        }
                        break;
                    default:
                        break;
                }
                return true;
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }
        private bool Registration() {
            try {
                StrRegDate = RegistrationDateMaskedTextBox.Text;
                StrPerson = PersonComboBox.Text;
                if (string.IsNullOrEmpty(StrPerson)) { throw new Exception("担当者を選択してください。"); }

                switch (IntRegType) {
                    case 2:
                    case 3:
                        using (SQLiteConnection _con = new(MainWindow.GetConnectionString2())) {
                            _con.Open();
                            if (ArrUseSubstrate == null) { throw new Exception("ArrUseSubstrateがnullです。"); }
                            for (int _i = 0; _i <= ArrUseSubstrate.Length; _i++) {

                                CheckBox? _objCbx = Controls[CheckBoxNames[_i]] as CheckBox ?? throw new Exception("objCbxがnullです。");

                                if (_objCbx.Checked) {
                                    DataGridView? _objDgv = Controls[DataGridViewNames[_i]] as DataGridView ?? throw new Exception("objCbxがnullです。");
                                    int _dgvRowCnt = _objDgv.Rows.Count;
                                    string _subTotalTemp = string.Empty;

                                    for (int _j = 0; _j <= _dgvRowCnt - 1; _j++) {
                                        bool _boolCbx = Convert.ToBoolean(_objDgv.Rows[_j].Cells[4].Value);
                                        if (_boolCbx) {
                                            string _substrateName = string.Empty;
                                            string _substrateModel = string.Empty;
                                            string _substrateNum = _objDgv.Rows[_j].Cells[0].Value.ToString() ?? string.Empty;
                                            int _stockValue = Convert.ToInt32(_objDgv.Rows[_j].Cells[1].Value);
                                            int _usedValue = Convert.ToInt32(_objDgv.Rows[_j].Cells[2].Value.ToString());
                                            int _useValue = Convert.ToInt32(_objDgv.Rows[_j].Cells[3].Value);

                                            using (SQLiteCommand _cmd = _con.CreateCommand()) {
                                                _cmd.CommandText =
                                                    $@"
                                                    UPDATE Stock_{StrStockName}
                                                    SET
                                                        col_Flg = @col_Flg,
                                                        col_Stock = @col_Stock,
                                                        col_History = ifnull(col_History,'')|| @col_History
                                                    WHERE
                                                        col_Substrate_Num = @col_Substrate_Num
                                                    ";
                                                _cmd.Parameters.Add("@col_Substrate_Num", DbType.String).Value = _substrateNum;

                                                switch (_useValue - _usedValue) {
                                                    case int diff when diff > 0:
                                                        _cmd.Parameters.Add("@col_Flg", DbType.String).Value = (_stockValue - diff == 0) ? 0 : 1;
                                                        break;
                                                    case int diff when diff < 0:
                                                        _cmd.Parameters.Add("@col_Flg", DbType.String).Value = 1;
                                                        break;
                                                    case int diff when diff == 0:
                                                        _cmd.Parameters.Add("@col_Flg", DbType.String).Value = (_stockValue == 0) ? 0 : 1;
                                                        break;
                                                }

                                                _cmd.Parameters.Add("@col_Stock", DbType.String).Value = _stockValue + _usedValue - _useValue;
                                                _cmd.Parameters.Add("@col_History", DbType.String).Value = $"{StrProductNumber}({_useValue}),";

                                                _cmd.ExecuteNonQuery();

                                                _cmd.CommandText = $@"SELECT * FROM Stock_{StrStockName} WHERE col_Substrate_Model = @col_Substrate_Model ORDER BY _rowid_ ASC";
                                                _cmd.Parameters.Add("@col_Substrate_Model", DbType.String).Value = ArrUseSubstrate[_i];
                                                using SQLiteDataReader _dr = _cmd.ExecuteReader();
                                                while (_dr.Read()) {
                                                    _substrateName = $"{_dr["col_Substrate_Name"]}";
                                                    _substrateModel = $"{_dr["col_Substrate_Model"]}";
                                                }
                                            }

                                            if (_useValue != 0) {
                                                if (string.IsNullOrEmpty(_subTotalTemp)) {
                                                    _subTotalTemp = $"{_substrateNum}({_useValue})";
                                                }
                                                else {
                                                    _subTotalTemp = $"{_subTotalTemp},{_substrateNum}({_useValue})";
                                                }
                                            }

                                            // 基板テーブルへ追加
                                            using (SQLiteCommand _cmd = _con.CreateCommand()) {
                                                _cmd.CommandText =
                                                    $@"
                                                    INSERT INTO Substrate_Reg_{StrStockName}
                                                        (col_Substrate_Name,
                                                        col_Substrate_Model,
                                                        col_Substrate_Num,
                                                        col_Decrease,
                                                        col_Use_P_Type,
                                                        col_Use_P_Num,
                                                        col_Use_O_Num,
                                                        col_Person,
                                                        col_RegDate,
                                                        col_Comment)
                                                    VALUES
                                                        (@col_Substrate_Name,
                                                        @col_Substrate_Model,
                                                        @col_Substrate_Num,
                                                        @col_Decrease,
                                                        @col_Use_P_Type,
                                                        @col_Use_P_Num,
                                                        @col_Use_O_Num,
                                                        @col_Person,
                                                        @col_RegDate,
                                                        @col_Comment)
                                                    ";

                                                _cmd.Parameters.Add("@col_Substrate_Name", DbType.String).Value = _substrateName;
                                                _cmd.Parameters.Add("@col_Substrate_Model", DbType.String).Value = _substrateModel;
                                                _cmd.Parameters.Add("@col_Substrate_Num", DbType.String).Value = _substrateNum;
                                                _cmd.Parameters.Add("@col_Decrease", DbType.String).Value = 0 - _useValue;
                                                _cmd.Parameters.Add("@col_Use_P_Type", DbType.String).Value = StrProductType;
                                                _cmd.Parameters.Add("@col_Use_P_Num", DbType.String).Value = StrProductNumber;
                                                _cmd.Parameters.Add("@col_Use_O_Num", DbType.String).Value = StrOrderNumber;
                                                _cmd.Parameters.Add("@col_Person", DbType.String).Value = StrPerson;
                                                _cmd.Parameters.Add("@col_RegDate", DbType.String).Value = StrRegDate;
                                                _cmd.Parameters.Add("@col_Comment", DbType.String).Value = StrComment;

                                                _cmd.ExecuteNonQuery();
                                            }

                                            switch (IntPrintType) {
                                                case 5:
                                                case 6:
                                                    ListUsedSubstrate.Add(ArrUseSubstrate[_i]);
                                                    if (_substrateNum != null) { ListUsedProductNumber.Add(_substrateNum); }
                                                    ListUsedQuantity.Add(_useValue);
                                                    break;
                                                default:
                                                    break;
                                            }
                                        }
                                    }
                                    if (string.IsNullOrEmpty(StrTotalSubstrate)) {
                                        StrTotalSubstrate = $"[{ArrUseSubstrate[_i]}]{_subTotalTemp}";
                                    }
                                    else {
                                        StrTotalSubstrate = $"{StrTotalSubstrate},[{ArrUseSubstrate[_i]}]{_subTotalTemp}";
                                    }
                                    _subTotalTemp = string.Empty;
                                }
                            }

                            // 製品テーブルへ追加
                            using (SQLiteCommand _cmd = _con.CreateCommand()) {
                                _cmd.CommandText =
                                    $@"
                                    UPDATE Product_Reg_{StrProductName}
                                    SET
                                        col_Quantity = @col_Quantity,
                                        col_Person = @col_Person,
                                        col_RegDate = @col_RegDate,
                                        col_Revision = @col_Revision,
                                        col_Serial_Last = @col_Serial_Last,
                                        col_Serial_LastNum = @col_Serial_LastNum,
                                        col_Comment = @col_Comment
                                        col_Use_Substrate = @col_Use_Substrate
                                    WHERE
                                        col_Product_Num = @col_Product_Num
                                    AND
                                        col_Serial_First = @col_Serial_First
                                    ";

                                _cmd.Parameters.Add("@col_Product_Type", DbType.String).Value = StrProductType;
                                _cmd.Parameters.Add("@col_Product_Model", DbType.String).Value = StrProductModel;
                                _cmd.Parameters.Add("@col_Order_Num", DbType.String).Value = StrOrderNumber;
                                _cmd.Parameters.Add("@col_Product_Num", DbType.String).Value = StrProductNumber;
                                _cmd.Parameters.Add("@col_Quantity", DbType.String).Value = IntQuantity;
                                _cmd.Parameters.Add("@col_Person", DbType.String).Value = StrPerson;
                                _cmd.Parameters.Add("@col_RegDate", DbType.String).Value = StrRegDate;
                                _cmd.Parameters.Add("@col_Revision", DbType.String).Value = StrRevision;
                                _cmd.Parameters.Add("@col_Serial_First", DbType.String).Value = StrSerialFirstNumber;
                                _cmd.Parameters.Add("@col_Serial_Last", DbType.String).Value = StrSerialLastNumber;
                                _cmd.Parameters.Add("@col_Serial_LastNum", DbType.String).Value = IntSerialLastNumber;
                                _cmd.Parameters.Add("@col_Use_Substrate", DbType.String).Value = StrTotalSubstrate;
                                _cmd.Parameters.Add("@col_Comment", DbType.String).Value = StrComment;

                                _cmd.ExecuteNonQuery();
                            }
                        }
                        break;
                }
                return true;
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        // リスト印刷
        private void ListPrint() {
            try {
                string _sheetName = string.Empty;
                string _productName = string.Empty;
                string _productModel = string.Empty;

                string _productNameRange = string.Empty;
                string _productNumberRange = string.Empty;
                string _orderNumberRange = string.Empty;
                string _regDateRange = string.Empty;
                string _productModelRange = string.Empty;
                string _quantityRange = string.Empty;
                string _serialFirstRange = string.Empty;
                string _serialLastRange = string.Empty;
                string _commentRange = string.Empty;

                using FileStream _fileStream = new($@"{Environment.CurrentDirectory}./config/Excel/ConfigList.xlsx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using XLWorkbook _workBook = new(_fileStream);
                IXLWorksheet _workSheetMain = _workBook.Worksheet("Sheet1");

                int _findRow = 0;
                // セル検索
                foreach (IXLCell _cell in _workSheetMain.Search(StrProductModel)) {
                    _findRow = _cell.Address.RowNumber;
                }

                if (_findRow == 0) {
                    throw new Exception($"Configに品目番号:{StrProductModel}が見つかりません。");
                }

                // ワークシートのセルから値を取得
                _sheetName = StrProductModel;
                _productName = _workSheetMain.Cell(_findRow, 2).Value.ToString();
                _productNameRange = _workSheetMain.Cell(_findRow, 3).Value.ToString();
                _productNumberRange = _workSheetMain.Cell(_findRow, 4).Value.ToString();
                _orderNumberRange = _workSheetMain.Cell(_findRow, 5).Value.ToString();
                _regDateRange = _workSheetMain.Cell(_findRow, 6).Value.ToString();
                _productModel = _workSheetMain.Cell(_findRow, 7).Value.ToString();
                _productModelRange = _workSheetMain.Cell(_findRow, 8).Value.ToString();
                _quantityRange = _workSheetMain.Cell(_findRow, 9).Value.ToString();
                _serialFirstRange = _workSheetMain.Cell(_findRow, 10).Value.ToString();
                _serialLastRange = _workSheetMain.Cell(_findRow, 11).Value.ToString();
                _commentRange = _workSheetMain.Cell(_findRow, 12).Value.ToString();

                IXLWorksheet _workSheetTemp = _workBook.Worksheet(_sheetName);
                _workSheetTemp.Cell(_productNameRange).Value = _productName;
                _workSheetTemp.Cell(_productNumberRange).Value = StrProductNumber;
                _workSheetTemp.Cell(_orderNumberRange).Value = StrOrderNumber;
                _workSheetTemp.Cell(_regDateRange).Value = StrRegDate;
                _workSheetTemp.Cell(_productModelRange).Value = _productModel;
                _workSheetTemp.Cell(_quantityRange).Value = IntQuantity;
                _workSheetTemp.Cell(_serialFirstRange).Value = StrSerialFirstNumber;
                _workSheetTemp.Cell(_serialLastRange).Value = StrSerialLastNumber;
                _workSheetTemp.Cell(_commentRange).Value = StrComment;

                int _i = 0;
                int _findColumn = 0;
                for (_i = 0; _i <= ListUsedSubstrate.Count - 1; _i++) {

                    IXLRange _searchRange = _workSheetMain.Range(_findRow, 1, _findRow, 28);
                    string _searchValue = $"{ListUsedSubstrate[_i]}";
                    IXLCell? _foundCell = _searchRange.CellsUsed(c => c.Value.ToString() == _searchValue).FirstOrDefault();

                    if (_foundCell != null) {
                        // セルが見つかった場合の処理
                        int foundRow = _foundCell.Address.RowNumber;
                        int foundColumn = _foundCell.Address.ColumnNumber;
                    }

                    foreach (IXLCell _cell in _workSheetMain.Search(ListUsedSubstrate[_i])) {
                        if (_cell.Address.RowNumber == _findRow) {
                            _findColumn = _cell.Address.ColumnNumber;
                            break;
                        }
                    }

                    if (_findColumn == 0) {
                        throw new Exception($"{ListUsedSubstrate[_i]}が見つかりません。");
                    }

                    var _mainCellValue = _workSheetMain.Cell(_findRow, _findColumn + 1).Value.ToString();
                    var _tempCellValue = _workSheetTemp.Cell(_mainCellValue).Value.ToString();

                    if (_mainCellValue != string.Empty) {
                        if (_tempCellValue == string.Empty) {
                            _workSheetTemp.Cell(_mainCellValue).Value = $"{ListUsedProductNumber[_i]}({ListUsedQuantity[_i]})";
                        }
                        else {
                            _workSheetTemp.Cell(_mainCellValue).Value += $"    {ListUsedProductNumber[_i]}({ListUsedQuantity[_i]})";
                        }
                    }
                }

                //引数に保存先パスを指定
                _workBook.SaveAs($@"{Environment.CurrentDirectory}./config/Excel/temporarily.xlsx");

                // 印刷
                Excel.Application _xlApp = new() {
                    Visible = true // Excelウィンドウを表示します。
                };

                // ワークブック開く
                Excel.Workbooks _xlBooks = _xlApp.Workbooks;
                Excel.Workbook _xlBook = _xlBooks.Open($@"{Environment.CurrentDirectory}./config/Excel/temporarily.xlsx");

                // ワークシート選択
                Excel.Sheets _xlSheets = _xlBook.Sheets;
                Excel.Worksheet _xlSheet = _xlSheets[_sheetName];

                // ワークシート印刷
                _xlSheet.PrintOut(Preview: true);

                // ワークブックを閉じてExcelを終了
                _xlBook.Close(false);
                _xlApp.Quit();

                _ = System.Runtime.InteropServices.Marshal.ReleaseComObject(_xlSheet);
                _ = System.Runtime.InteropServices.Marshal.ReleaseComObject(_xlSheets);
                _ = System.Runtime.InteropServices.Marshal.ReleaseComObject(_xlBook);
                _ = System.Runtime.InteropServices.Marshal.ReleaseComObject(_xlBooks);
                _ = System.Runtime.InteropServices.Marshal.ReleaseComObject(_xlApp);

            } catch (Exception ex) {
                MessageBox.Show($"{ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
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
