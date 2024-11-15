using ClosedXML.Excel;
using System.Data;
using System.Data.SQLite;
using static ProductDatabase.MainWindow;
using Excel = Microsoft.Office.Interop.Excel;

namespace ProductDatabase {
    public partial class SubstrateChange2 : Form {

        public ProductInfomation ProductInfo { get; set; } = new ProductInfomation();

        private string _totalSubstrate = String.Empty;

        private string[] _useSubstrate = [];
        private string[] _usedSubstrate = [];

        private readonly List<string> _listUsedSubstrate = [];
        private readonly List<string> _listUsedProductNumber = [];
        private readonly List<int> _listUsedQuantity = [];
        private readonly List<string> _checkBoxNames = [
                        "Substrate1CheckBox", "Substrate2CheckBox", "Substrate3CheckBox", "Substrate4CheckBox","Substrate5CheckBox",
                        "Substrate6CheckBox", "Substrate7CheckBox", "Substrate8CheckBox", "Substrate9CheckBox","Substrate10CheckBox"
                        ];
        private readonly List<string> _dataGridViewNames = [
                        "Substrate1DataGridView", "Substrate2DataGridView", "Substrate3DataGridView", "Substrate4DataGridView","Substrate5DataGridView",
                        "Substrate6DataGridView", "Substrate7DataGridView", "Substrate8DataGridView", "Substrate9DataGridView","Substrate10DataGridView"
                        ];
        private CheckBox? _objCbx;
        private DataGridView? _objDgv;

        public SubstrateChange2() {
            InitializeComponent();
        }

        // ロードイベント
        private void LoadEvents() {
            try {
                Font = new Font(ProductInfo.FontName, ProductInfo.FontSize);
                CloseButton.Enabled = true;

                // テキストボックスに入力
                OrderNumberTextBox.Text = ProductInfo.OrderNumber;
                ManufacturingNumberTextBox.Text = ProductInfo.ProductNumber;
                QuantityTextBox.Text = ProductInfo.Quantity.ToString();
                var dtNow = DateTime.Now;
                RegistrationDateMaskedTextBox.Text = dtNow.ToShortDateString();
                RevisionTextBox.Text = ProductInfo.Revision;
                CommentTextBox.Text = ProductInfo.Comment;

                // DB1へ接続し担当者取得
                using (SQLiteConnection con = new(GetConnectionString1())) {
                    con.Open();
                    using var cmd = con.CreateCommand();
                    // テーブル検索SQL - 担当者をComboboxへ追加
                    cmd.CommandText = "SELECT * FROM Person ORDER BY _rowid_ ASC";
                    using var dr = cmd.ExecuteReader();
                    while (dr.Read()) {
                        PersonComboBox.Items.Add($"{dr["PersonName"]}");
                    }
                }

                _useSubstrate = ProductInfo.UseSubstrate.Split(",");
                _usedSubstrate = ProductInfo.UsedSubstrate.Split(",");
                var strQuantity = String.Empty;
                var strSubstrateName = String.Empty;

                switch (ProductInfo.RegType) {
                    case 2:
                        for (var i = 0; i <= _useSubstrate.GetUpperBound(0); i++) {
                            var quantity = ProductInfo.Quantity;

                            // チェックボックスとDgvを有効に
                            _objCbx = Controls[_checkBoxNames[i]] as CheckBox;
                            if (_objCbx != null) {
                                _objCbx.Enabled = true;
                                _objCbx.Checked = true;
                            }

                            _objDgv = Controls[_dataGridViewNames[i]] as DataGridView;
                            if (_objDgv != null) {
                                _objDgv.RowHeadersWidth = 30;
                                _objDgv.Columns[2].DefaultCellStyle.BackColor = Color.LightGray;
                                _objDgv.Columns[2].ReadOnly = true;
                                _objDgv.Columns[3].ReadOnly = false;
                                _objDgv.Columns[4].ReadOnly = false;
                            }

                            using SQLiteConnection con = new(GetConnectionString2());
                            con.Open();

                            using var cmd = con.CreateCommand();
                            // テーブル検索SQL - [Product_Name]_Stockテーブルを在庫数フラグ有&基板型式[Model]で抽出して取得
                            cmd.CommandText = $"SELECT * FROM Stock_{ProductInfo.StockName} WHERE SubstrateModel = @SubstrateModel ORDER BY _rowid_ ASC";
                            cmd.Parameters.Add("@SubstrateModel", DbType.String).Value = _useSubstrate[i];
                            using var dr = cmd.ExecuteReader();
                            var j = 0;
                            while (dr.Read()) {
                                var strUsedSubNum = String.Empty;
                                var strUsedQuantity = String.Empty;
                                var intUsedQuantity = 0;

                                // 抽出した行から製造番号,在庫取得
                                var strSubstrateNum = $"{dr["SubstrateNumber"]}";
                                var intStock = Int32.Parse($"{dr["Stock"]}");
                                //var colFlg = Int32.Parse($"{dr["col_Flg"]}");
                                strSubstrateName = $"{dr["SubstrateName"]}";
                                if (_objCbx != null) { _objCbx.Text = $"{strSubstrateName} - {_useSubstrate[i]}"; }

                                var subIndex = 0;
                                var b1 = false;

                                for (var i1 = 0; i1 < _usedSubstrate.Length; i1++) {
                                    var d = _usedSubstrate[i1];
                                    b1 = d.Contains(strSubstrateNum);
                                    if (b1 == true) {
                                        break;
                                    }

                                    subIndex++;
                                }
                                if (b1 == true) {
                                    strUsedSubNum = _usedSubstrate[subIndex];
                                    strUsedSubNum = strUsedSubNum[(strUsedSubNum.IndexOf(']') + 1)..];
                                    strUsedSubNum = strUsedSubNum[..strUsedSubNum.IndexOf('(')];

                                    strUsedQuantity = _usedSubstrate[subIndex];
                                    strUsedQuantity = strUsedQuantity[(strUsedQuantity.IndexOf('(') + 1)..];
                                    strUsedQuantity = strUsedQuantity[..strUsedQuantity.IndexOf(')')];
                                    intUsedQuantity = Int32.Parse(strUsedQuantity);
                                }

                                if (intStock > 0) {
                                    if (_objDgv == null) { break; }
                                    _objDgv.Rows.Add();
                                    _objDgv.Rows[j].Cells[0].Value = strSubstrateNum;
                                    _objDgv.Rows[j].Cells[1].Value = intStock;

                                    if (intUsedQuantity != 0) {
                                        _objDgv.Rows[j].Cells[2].Value = intUsedQuantity;
                                        _objDgv.Rows[j].Cells[3].Value = intUsedQuantity;
                                        _objDgv.Rows[j].Cells[4].Value = true;
                                    }
                                    else {
                                        _objDgv.Rows[j].Cells[2].Value = 0;
                                        _objDgv.Rows[j].Cells[3].Value = 0;
                                    }
                                    j++;
                                }
                                else if (strUsedSubNum == strSubstrateNum) {
                                    if (_objDgv == null) { break; }
                                    _objDgv.Rows.Add();
                                    _objDgv.Rows[j].Cells[0].Value = strSubstrateNum;
                                    _objDgv.Rows[j].Cells[1].Value = intStock;

                                    if (intUsedQuantity != 0) {
                                        _objDgv.Rows[j].Cells[2].Value = intUsedQuantity;
                                        _objDgv.Rows[j].Cells[3].Value = intUsedQuantity;
                                        _objDgv.Rows[j].Cells[4].Value = true;
                                    }
                                    else {
                                        _objDgv.Rows[j].Cells[2].Value = 0;
                                        _objDgv.Rows[j].Cells[3].Value = 0;
                                    }
                                    j++;
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
                for (var i = 0; i <= 9; i++) {
                    _objCbx = Controls[_checkBoxNames[i]] as CheckBox;
                    if (_objCbx != null) {
                        _objCbx.Enabled = false;
                    }

                    _objDgv = Controls[_dataGridViewNames[i]] as DataGridView;
                    if (_objDgv != null) {
                        _objDgv.Enabled = false;
                    }
                }
                // リスト印刷ボタンを有効に
                if (ProductInfo.PrintType is 5 or 6) {
                    SubstrateListPrintButton.Enabled = true;
                }

            } catch (Exception ex) {
                MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private bool QuantityCheck() {
            try {
                switch (ProductInfo.RegType) {
                    case 2:
                    case 3:
                        if (_useSubstrate == null) { throw new Exception("ArrUseSubstrateが空です"); }
                        for (var i = 0; i <= _useSubstrate.GetUpperBound(0); i++) {

                            var objCbx = Controls[_checkBoxNames[i]] as CheckBox ?? throw new Exception("objCbxがnullです。");
                            objCbx.Enabled = true;
                            objCbx.Checked = true;

                            var objDgv = Controls[_dataGridViewNames[i]] as DataGridView ?? throw new Exception("objDgvがnullです。");
                            objDgv.Columns[3].ReadOnly = false;
                            objDgv.Columns[4].ReadOnly = false;

                            if (objCbx.Checked) {
                                var intQuantityCheck = ProductInfo.Quantity;
                                var dgvRowCnt = objDgv.Rows.Count;

                                for (var j = 0; j < dgvRowCnt; j++) {
                                    var boolCbx = objDgv.Rows[j].Cells[4].Value != null && (bool)objDgv.Rows[j].Cells[4].Value;
                                    if (boolCbx) {
                                        var stockValue = Convert.ToInt32(objDgv.Rows[j].Cells[1].Value.ToString());
                                        var usedValue = Convert.ToInt32(objDgv.Rows[j].Cells[2].Value.ToString());
                                        var useValue = Convert.ToInt32(objDgv.Rows[j].Cells[3].Value.ToString());

                                        if (stockValue + usedValue < useValue) {
                                            throw new Exception("在庫より多い数量が入力されています。");
                                        }
                                        intQuantityCheck -= useValue;
                                    }
                                }

                                if (intQuantityCheck != 0) {
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
                ProductInfo.RegDate = RegistrationDateMaskedTextBox.Text;
                ProductInfo.Person = PersonComboBox.Text;
                if (String.IsNullOrEmpty(ProductInfo.Person)) { throw new Exception("担当者を選択してください。"); }

                switch (ProductInfo.RegType) {
                    case 2:
                    case 3:
                        using (SQLiteConnection con = new(GetConnectionString2())) {
                            con.Open();
                            if (_useSubstrate == null) { throw new Exception("ArrUseSubstrateがnullです。"); }
                            for (var i = 0; i <= _useSubstrate.Length; i++) {

                                var objCbx = Controls[_checkBoxNames[i]] as CheckBox ?? throw new Exception("objCbxがnullです。");

                                if (objCbx.Checked) {
                                    var objDgv = Controls[_dataGridViewNames[i]] as DataGridView ?? throw new Exception("objCbxがnullです。");
                                    var dgvRowCnt = objDgv.Rows.Count;
                                    var subTotalTemp = String.Empty;

                                    for (var j = 0; j <= dgvRowCnt - 1; j++) {
                                        var boolCbx = Convert.ToBoolean(objDgv.Rows[j].Cells[4].Value);
                                        if (boolCbx) {
                                            var substrateName = String.Empty;
                                            var substrateModel = String.Empty;
                                            var substrateNum = objDgv.Rows[j].Cells[0].Value.ToString() ?? String.Empty;
                                            var stockValue = Convert.ToInt32(objDgv.Rows[j].Cells[1].Value);
                                            var usedValue = Convert.ToInt32(objDgv.Rows[j].Cells[2].Value.ToString());
                                            var useValue = Convert.ToInt32(objDgv.Rows[j].Cells[3].Value);

                                            using (var cmd = con.CreateCommand()) {
                                                cmd.CommandText =
                                                    $"""
                                                    UPDATE Stock_{ProductInfo.StockName}
                                                    SET
                                                        Stock = @Stock,
                                                        col_History = ifnull(col_History,'')|| @col_History
                                                    WHERE
                                                        SubstrateNumber = @SubstrateNumber
                                                    """;
                                                //cmd.CommandText =
                                                //    $"""
                                                //    UPDATE Stock_{ProductInfo.StockName}
                                                //    SET
                                                //        col_Flg = @col_Flg,
                                                //        Stock = @Stock,
                                                //        col_History = ifnull(col_History,'')|| @col_History
                                                //    WHERE
                                                //        SubstrateNumber = @SubstrateNumber
                                                //    """;
                                                cmd.Parameters.Add("@SubstrateNumber", DbType.String).Value = substrateNum;

                                                //switch (useValue - usedValue) {
                                                //    case int diff when diff > 0:
                                                //        cmd.Parameters.Add("@col_Flg", DbType.String).Value = (stockValue - diff == 0) ? 0 : 1;
                                                //        break;
                                                //    case int diff when diff < 0:
                                                //        cmd.Parameters.Add("@col_Flg", DbType.String).Value = 1;
                                                //        break;
                                                //    case int diff when diff == 0:
                                                //        cmd.Parameters.Add("@col_Flg", DbType.String).Value = (stockValue == 0) ? 0 : 1;
                                                //        break;
                                                //}

                                                cmd.Parameters.Add("@Stock", DbType.String).Value = stockValue + usedValue - useValue;
                                                cmd.Parameters.Add("@col_History", DbType.String).Value = $"{ProductInfo.ProductNumber}({useValue}),";

                                                cmd.ExecuteNonQuery();

                                                cmd.CommandText = $@"SELECT * FROM Stock_{ProductInfo.StockName} WHERE SubstrateModel = @SubstrateModel ORDER BY _rowid_ ASC";
                                                cmd.Parameters.Add("@SubstrateModel", DbType.String).Value = _useSubstrate[i];
                                                using var dr = cmd.ExecuteReader();
                                                while (dr.Read()) {
                                                    substrateName = $"{dr["SubstrateName"]}";
                                                    substrateModel = $"{dr["SubstrateModel"]}";
                                                }
                                            }

                                            if (useValue != 0) {
                                                subTotalTemp = String.IsNullOrEmpty(subTotalTemp) ? $"{substrateNum}({useValue})" : $"{subTotalTemp},{substrateNum}({useValue})";
                                            }

                                            // 基板テーブルへ追加
                                            using (var cmd = con.CreateCommand()) {
                                                cmd.CommandText =
                                                    $"""
                                                    INSERT INTO Substrate_{ProductInfo.StockName}
                                                        (
                                                        SubstrateName,
                                                        SubstrateModel,
                                                        SubstrateNumber,
                                                        col_Decrease,
                                                        col_Use_P_Type,
                                                        col_Use_P_Num,
                                                        col_Use_O_Num,
                                                        Person,
                                                        RegDate,
                                                        Comment
                                                        )
                                                    VALUES
                                                        (
                                                        @SubstrateName,
                                                        @SubstrateModel,
                                                        @SubstrateNumber,
                                                        @col_Decrease,
                                                        @col_Use_P_Type,
                                                        @col_Use_P_Num,
                                                        @col_Use_O_Num,
                                                        @Person,
                                                        @RegDate,
                                                        @Comment
                                                        )
                                                    """;

                                                cmd.Parameters.Add("@SubstrateName", DbType.String).Value = String.IsNullOrWhiteSpace(substrateName) ? DBNull.Value : substrateName;
                                                cmd.Parameters.Add("@SubstrateModel", DbType.String).Value = String.IsNullOrWhiteSpace(substrateModel) ? DBNull.Value : substrateModel;
                                                cmd.Parameters.Add("@SubstrateNumber", DbType.String).Value = String.IsNullOrWhiteSpace(substrateNum) ? DBNull.Value : substrateNum;
                                                cmd.Parameters.Add("@col_Decrease", DbType.String).Value = 0 - useValue;
                                                cmd.Parameters.Add("@col_Use_P_Type", DbType.String).Value = String.IsNullOrWhiteSpace(ProductInfo.ProductType) ? DBNull.Value : ProductInfo.ProductType;
                                                cmd.Parameters.Add("@col_Use_P_Num", DbType.String).Value = String.IsNullOrWhiteSpace(ProductInfo.ProductNumber) ? DBNull.Value : ProductInfo.ProductNumber;
                                                cmd.Parameters.Add("@col_Use_O_Num", DbType.String).Value = String.IsNullOrWhiteSpace(ProductInfo.OrderNumber) ? DBNull.Value : ProductInfo.OrderNumber;
                                                cmd.Parameters.Add("@Person", DbType.String).Value = String.IsNullOrWhiteSpace(ProductInfo.Person) ? DBNull.Value : ProductInfo.Person;
                                                cmd.Parameters.Add("@RegDate", DbType.String).Value = String.IsNullOrWhiteSpace(ProductInfo.RegDate) ? DBNull.Value : ProductInfo.RegDate;
                                                cmd.Parameters.Add("@Comment", DbType.String).Value = String.IsNullOrWhiteSpace(ProductInfo.Comment) ? DBNull.Value : ProductInfo.Comment;

                                                cmd.ExecuteNonQuery();
                                            }

                                            switch (ProductInfo.PrintType) {
                                                case 5:
                                                case 6:
                                                    _listUsedSubstrate.Add(_useSubstrate[i]);
                                                    if (substrateNum != null) { _listUsedProductNumber.Add(substrateNum); }
                                                    _listUsedQuantity.Add(useValue);
                                                    break;
                                                default:
                                                    break;
                                            }
                                        }
                                    }
                                    _totalSubstrate = String.IsNullOrEmpty(_totalSubstrate)
                                        ? $"[{_useSubstrate[i]}]{subTotalTemp}"
                                        : $"{_totalSubstrate},[{_useSubstrate[i]}]{subTotalTemp}";
                                    subTotalTemp = String.Empty;
                                }
                            }

                            // 製品テーブルへ追加
                            using (var cmd = con.CreateCommand()) {
                                cmd.CommandText =
                                    $"""
                                    UPDATE Product_{ProductInfo.ProductName}
                                    SET
                                        Quantity = @Quantity,
                                        Person = @Person,
                                        RegDate = @RegDate,
                                        Revision = @Revision,
                                        SerialLast = @SerialLast,
                                        SerialLastNum = @SerialLastNum,
                                        Comment = @Comment
                                        col_Use_Substrate = @col_Use_Substrate
                                    WHERE
                                        ProductNumber = @ProductNumber
                                    AND
                                        SerialFirst = @SerialFirst
                                    """;

                                cmd.Parameters.Add("@ProductType", DbType.String).Value = String.IsNullOrWhiteSpace(ProductInfo.ProductType) ? DBNull.Value : ProductInfo.ProductType;
                                cmd.Parameters.Add("@ProductModel", DbType.String).Value = String.IsNullOrWhiteSpace(ProductInfo.ProductModel) ? DBNull.Value : ProductInfo.ProductModel;
                                cmd.Parameters.Add("@OrderNumber", DbType.String).Value = String.IsNullOrWhiteSpace(ProductInfo.OrderNumber) ? DBNull.Value : ProductInfo.OrderNumber;
                                cmd.Parameters.Add("@ProductNumber", DbType.String).Value = String.IsNullOrWhiteSpace(ProductInfo.ProductNumber) ? DBNull.Value : ProductInfo.ProductNumber;
                                cmd.Parameters.Add("@Quantity", DbType.String).Value = ProductInfo.Quantity;
                                cmd.Parameters.Add("@Person", DbType.String).Value = String.IsNullOrWhiteSpace(ProductInfo.Person) ? DBNull.Value : ProductInfo.Person;
                                cmd.Parameters.Add("@RegDate", DbType.String).Value = String.IsNullOrWhiteSpace(ProductInfo.RegDate) ? DBNull.Value : ProductInfo.RegDate;
                                cmd.Parameters.Add("@Revision", DbType.String).Value = String.IsNullOrWhiteSpace(ProductInfo.Revision) ? DBNull.Value : ProductInfo.Revision;
                                cmd.Parameters.Add("@SerialFirst", DbType.String).Value = String.IsNullOrWhiteSpace(ProductInfo.SerialFirst) ? DBNull.Value : ProductInfo.SerialFirst;
                                cmd.Parameters.Add("@SerialLast", DbType.String).Value = String.IsNullOrWhiteSpace(ProductInfo.SerialLast) ? DBNull.Value : ProductInfo.SerialLast;
                                cmd.Parameters.Add("@SerialLastNum", DbType.String).Value = ProductInfo.SerialLastNumber;
                                cmd.Parameters.Add("@col_Use_Substrate", DbType.String).Value = String.IsNullOrWhiteSpace(_totalSubstrate) ? DBNull.Value : _totalSubstrate;
                                cmd.Parameters.Add("@Comment", DbType.String).Value = String.IsNullOrWhiteSpace(ProductInfo.Comment) ? DBNull.Value : ProductInfo.Comment;

                                cmd.ExecuteNonQuery();
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
                var sheetName = String.Empty;
                var productName = String.Empty;
                var productModel = String.Empty;

                var productNameRange = String.Empty;
                var productNumberRange = String.Empty;
                var orderNumberRange = String.Empty;
                var regDateRange = String.Empty;
                var productModelRange = String.Empty;
                var quantityRange = String.Empty;
                var serialFirstRange = String.Empty;
                var serialLastRange = String.Empty;
                var commentRange = String.Empty;

                using FileStream fileStream = new($"{Environment.CurrentDirectory}./config/Excel/ConfigList.xlsx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using XLWorkbook workBook = new(fileStream);
                var workSheetMain = workBook.Worksheet("Sheet1");

                var findRow = 0;
                // セル検索
                foreach (var cell in workSheetMain.Search(ProductInfo.ProductModel)) {
                    findRow = cell.Address.RowNumber;
                }

                if (findRow == 0) {
                    throw new Exception($"Configに品目番号:{ProductInfo.ProductModel}が見つかりません。");
                }

                // ワークシートのセルから値を取得
                sheetName = ProductInfo.ProductModel;
                productName = workSheetMain.Cell(findRow, 2).Value.ToString();
                productNameRange = workSheetMain.Cell(findRow, 3).Value.ToString();
                productNumberRange = workSheetMain.Cell(findRow, 4).Value.ToString();
                orderNumberRange = workSheetMain.Cell(findRow, 5).Value.ToString();
                regDateRange = workSheetMain.Cell(findRow, 6).Value.ToString();
                productModel = workSheetMain.Cell(findRow, 7).Value.ToString();
                productModelRange = workSheetMain.Cell(findRow, 8).Value.ToString();
                quantityRange = workSheetMain.Cell(findRow, 9).Value.ToString();
                serialFirstRange = workSheetMain.Cell(findRow, 10).Value.ToString();
                serialLastRange = workSheetMain.Cell(findRow, 11).Value.ToString();
                commentRange = workSheetMain.Cell(findRow, 12).Value.ToString();

                var workSheetTemp = workBook.Worksheet(sheetName);
                workSheetTemp.Cell(productNameRange).Value = productName;
                workSheetTemp.Cell(productNumberRange).Value = ProductInfo.ProductNumber;
                workSheetTemp.Cell(orderNumberRange).Value = ProductInfo.OrderNumber;
                workSheetTemp.Cell(regDateRange).Value = ProductInfo.RegDate;
                workSheetTemp.Cell(productModelRange).Value = productModel;
                workSheetTemp.Cell(quantityRange).Value = ProductInfo.Quantity;
                workSheetTemp.Cell(serialFirstRange).Value = ProductInfo.SerialFirst;
                workSheetTemp.Cell(serialLastRange).Value = ProductInfo.SerialLast;
                workSheetTemp.Cell(commentRange).Value = ProductInfo.Comment;

                var i = 0;
                var findColumn = 0;
                for (i = 0; i <= _listUsedSubstrate.Count - 1; i++) {

                    var searchRange = workSheetMain.Range(findRow, 1, findRow, 28);
                    var searchValue = $"{_listUsedSubstrate[i]}";
                    var foundCell = searchRange.CellsUsed(c => c.Value.ToString() == searchValue).FirstOrDefault();

                    if (foundCell != null) {
                        // セルが見つかった場合の処理
                        var foundRow = foundCell.Address.RowNumber;
                        var foundColumn = foundCell.Address.ColumnNumber;
                    }

                    foreach (var cell in workSheetMain.Search(_listUsedSubstrate[i])) {
                        if (cell.Address.RowNumber == findRow) {
                            findColumn = cell.Address.ColumnNumber;
                            break;
                        }
                    }

                    if (findColumn == 0) {
                        throw new Exception($"{_listUsedSubstrate[i]}が見つかりません。");
                    }

                    var mainCellValue = workSheetMain.Cell(findRow, findColumn + 1).Value.ToString();
                    var tempCellValue = workSheetTemp.Cell(mainCellValue).Value.ToString();

                    if (mainCellValue != String.Empty) {
                        if (tempCellValue == String.Empty) {
                            workSheetTemp.Cell(mainCellValue).Value = $"{_listUsedProductNumber[i]}({_listUsedQuantity[i]})";
                        }
                        else {
                            workSheetTemp.Cell(mainCellValue).Value += $"    {_listUsedProductNumber[i]}({_listUsedQuantity[i]})";
                        }
                    }
                }

                //引数に保存先パスを指定
                workBook.SaveAs($"{Environment.CurrentDirectory}./config/Excel/temporarily.xlsx");

                // 印刷
                Excel.Application xlApp = new() {
                    Visible = true // Excelウィンドウを表示します。
                };

                // ワークブック開く
                var xlBooks = xlApp.Workbooks;
                var xlBook = xlBooks.Open($"{Environment.CurrentDirectory}./config/Excel/temporarily.xlsx");

                // ワークシート選択
                var xlSheets = xlBook.Sheets;
                Excel.Worksheet xlSheet = xlSheets[sheetName];

                // ワークシート印刷
                xlSheet.PrintOut(Preview: true);

                // ワークブックを閉じてExcelを終了
                xlBook.Close(false);
                xlApp.Quit();

                _ = System.Runtime.InteropServices.Marshal.ReleaseComObject(xlSheet);
                _ = System.Runtime.InteropServices.Marshal.ReleaseComObject(xlSheets);
                _ = System.Runtime.InteropServices.Marshal.ReleaseComObject(xlBook);
                _ = System.Runtime.InteropServices.Marshal.ReleaseComObject(xlBooks);
                _ = System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);

            } catch (Exception ex) {
                MessageBox.Show($"{ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // チェックボックスイベント
        private void CheckBox_CheckedChanged(object sender, EventArgs e) {
            var checkBox = (CheckBox)sender;
            DataGridView dataGridView = new();

            switch (checkBox.Name) {
                case "Substrate1CheckBox":
                    dataGridView = Substrate1DataGridView;
                    break;
                case "Substrate2CheckBox":
                    dataGridView = Substrate2DataGridView;
                    break;
                case "Substrate3CheckBox":
                    dataGridView = Substrate3DataGridView;
                    break;
                case "Substrate4CheckBox":
                    dataGridView = Substrate4DataGridView;
                    break;
                case "Substrate5CheckBox":
                    dataGridView = Substrate5DataGridView;
                    break;
                case "Substrate6CheckBox":
                    dataGridView = Substrate6DataGridView;
                    break;
                case "Substrate7CheckBox":
                    dataGridView = Substrate7DataGridView;
                    break;
                case "Substrate8CheckBox":
                    dataGridView = Substrate8DataGridView;
                    break;
                case "Substrate9CheckBox":
                    dataGridView = Substrate9DataGridView;
                    break;
                case "Substrate10CheckBox":
                    dataGridView = Substrate10DataGridView;
                    break;
                default:
                    break;
            }

            dataGridView.Enabled = checkBox.Checked;
            checkBox.ForeColor = checkBox.Checked ? Color.Black : Color.Red;

            if (!checkBox.Checked) {
                MessageBox.Show("チェックがない場合在庫から引き落とされなくなります。", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void SubstrateChange2_Load(object sender, EventArgs e) { LoadEvents(); }
        private void RegisterButton_Click(object sender, EventArgs e) { ChangeRegistration(); }
        private void SubstrateListPrintButton_Click(object sender, EventArgs e) { ListPrint(); }
        private void CloseButton_Click(object sender, EventArgs e) { Close(); }
    }
}
