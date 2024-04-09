using System.Data;
using System.Data.Odbc;
using System.Data.SQLite;

namespace ProductDatabase {
    public partial class MainWindow : Form {

        public DataTable ProductDataTable { get; } = new();
        public string FontName { get; set; } = "Meiryo UI";

        public int IntFontSize { get; set; } = 9;
        public int IntRadioBtnFlg { get; set; }
        public int IntRegType { get; set; }
        public int IntPrintType { get; set; }
        public int IntSerialDigit { get; set; }

        public string StrProductName { get; set; } = string.Empty;
        public string StrStockName { get; set; } = string.Empty;
        public string StrProductType { get; set; } = string.Empty;
        public string StrProductModel { get; set; } = string.Empty;
        public string StrSubstrateName { get; set; } = string.Empty;
        public string StrSubstrateModel { get; set; } = string.Empty;
        public string StrUseSubstrate { get; set; } = string.Empty;
        public string StrInitial { get; set; } = string.Empty;

        public int IntCheckBin { get; set; }
        public string StrProness1 { get; set; } = string.Empty;
        public string StrProness2 { get; set; } = string.Empty;
        public string StrProness3 { get; set; } = string.Empty;
        public int StrProness4 { get; set; }
        public string StrProness5 { get; set; } = string.Empty;

        private int ListIndex { get; set; }

        public string StrCategory11 { get; set; } = string.Empty;
        public string StrCategory12 { get; set; } = string.Empty;
        public string StrCategory13 { get; set; } = string.Empty;
        public string StrCategory14 { get; set; } = string.Empty;
        public List<string> ListCategory11 { get; } = new();
        public List<string> ListCategory12 { get; } = new();
        public List<string> ListCategory13 { get; } = new();
        public List<string> ListCategory14 { get; } = new();

        public MainWindow() => InitializeComponent();

        public static string GetConnectionString1() {
            return new SQLiteConnectionStringBuilder() { DataSource = "./db/information.db" }.ToString();
        }
        public static string GetConnectionString2() {
            return new SQLiteConnectionStringBuilder() { DataSource = "./db/registration.db" }.ToString();
        }
        // ロードイベント
        private void LoadEvents() {
            try {
                // その日のbakファイルがない場合バックアップ作成
                DateTime _d = DateTime.Now;
                string _bakFilepath = $".\\bak\\{_d.Year}\\_bak_{_d.Year}-{_d.Month:00}-{_d.Day:00}.db";

                if (!File.Exists(_bakFilepath)) {
                    string _filepath = "./db/registration.db";

                    //コピー先のフォルダーが存在するか確認し、なければ作成します。
                    string _targetFolder = Path.GetDirectoryName(_bakFilepath) ?? throw new ArgumentNullException(nameof(_bakFilepath));
                    if (!Directory.Exists(_targetFolder)) {
                        Directory.CreateDirectory(_targetFolder);
                    }
                    File.Copy(_filepath, _bakFilepath, true);
                }

                ActiveControl = QRCodeTextBox;
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void ResetFields() {
            IntRegType = IntPrintType = IntSerialDigit = ListIndex = 0;
            StrProductName = StrProductType = StrProductModel = StrSubstrateName = StrSubstrateModel = StrUseSubstrate = StrInitial = string.Empty;
            IntCheckBin = 0;
            StrProness1 = StrProness2 = StrProness3 = StrProness5 = string.Empty;
            StrProness4 = 0;
        }
        // 登録ボタン処理
        private void Registration() {
            try {
                ResetFields();

                switch (IntRadioBtnFlg) {
                    case 1:
                        HandleSubstrateRegistration();
                        break;
                    case 2:
                        HandleProductRegistration1();
                        break;
                    case 3:
                        HandleRePrint();
                        break;
                    case 4:
                        HandleSubstrateChange1();
                        break;
                    default:
                        break;
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void HandleSubstrateRegistration() {
            DataRow[] selectedRows = ProductDataTable.Select($"class001 = '{CategoryListBox1.SelectedItem}' and col_Product_Name = '{CategoryListBox2.SelectedItem}' and col_Substrate_Name = '{CategoryListBox3.SelectedItem}'");

            if (selectedRows.Length > 0) {
                using SubstrateRegistrationWindow window = new();
                window.StrFontName = FontName;
                window.IntFontSize = IntFontSize;
                window.StrProductName = selectedRows[0]["col_Product_Name"].ToString() ?? string.Empty;
                window.StrStockName = selectedRows[0]["col_Stock_Name"].ToString() ?? string.Empty;
                window.StrSubstrateName = selectedRows[0]["col_Substrate_Name"].ToString() ?? string.Empty;
                window.StrSubstrateModel = selectedRows[0]["col_Substrate_Model"].ToString() ?? string.Empty;
                window.IntRegType = Convert.ToInt32(selectedRows[0]["col_Reg_Type"]);
                window.IntPrintType = Convert.ToInt32(selectedRows[0]["col_Print_Type"]);
                window.IntCheckBin = Convert.ToInt32(selectedRows[0]["col_Checkbox"].ToString(), 2);
                window.ShowDialog(this);
            }
        }
        private void HandleProductRegistration1() {
            DataRow[] selectedRows = ProductDataTable.Select($"class001 = '{CategoryListBox1.SelectedItem}' and col_Product_Name = '{CategoryListBox2.SelectedItem}' and col_Product_Type = '{CategoryListBox3.SelectedItem}'");

            if (selectedRows.Length > 0) {
                using ProductRegistration1Window window = new();
                window.StrFontName = FontName;
                window.IntFontSize = IntFontSize;
                window.StrProductName = selectedRows[0]["col_Product_Name"].ToString() ?? string.Empty;
                window.StrStockName = selectedRows[0]["col_Stock_Name"].ToString() ?? string.Empty;
                window.StrProductType = selectedRows[0]["col_Product_Type"].ToString() ?? string.Empty;
                window.IntRegType = Convert.ToInt32(selectedRows[0]["col_Reg_Type"]);
                window.IntPrintType = Convert.ToInt32(selectedRows[0]["col_Print_Type"]);
                window.IntSerialDigit = Convert.ToInt32(selectedRows[0]["col_Serial_Digit"]);
                window.StrProductModel = selectedRows[0]["col_Product_Model"].ToString() ?? string.Empty;
                window.IntCheckBin = Convert.ToInt32(selectedRows[0]["col_Checkbox"].ToString(), 2);
                window.StrUseSubstrate = selectedRows[0]["col_Use_Substrate"].ToString() ?? string.Empty;
                window.StrInitial = selectedRows[0]["col_Initial"].ToString() ?? string.Empty;
                window.ShowDialog(this);
            }
        }
        private void HandleRePrint() {
            DataRow[] selectedRows = ProductDataTable.Select($"class001 = '{CategoryListBox1.SelectedItem}' and col_Product_Name = '{CategoryListBox2.SelectedItem}' and col_Product_Type = '{CategoryListBox3.SelectedItem}'");

            if (selectedRows.Length > 0) {
                using RePrintWindow window = new();
                window.StrFontName = FontName;
                window.IntFontSize = IntFontSize;
                window.StrProductName = selectedRows[0]["col_Product_Name"].ToString() ?? string.Empty;
                window.StrProductType = selectedRows[0]["col_Product_Type"].ToString() ?? string.Empty;
                window.IntRegType = Convert.ToInt32(selectedRows[0]["col_Reg_Type"]);
                window.IntPrintType = Convert.ToInt32(selectedRows[0]["col_Print_Type"]);
                window.IntSerialDigit = Convert.ToInt32(selectedRows[0]["col_Serial_Digit"]);
                window.StrProductModel = selectedRows[0]["col_Product_Model"].ToString() ?? string.Empty;
                window.IntCheckBin = Convert.ToInt32(selectedRows[0]["col_Checkbox"].ToString(), 2);
                window.StrInitial = selectedRows[0]["col_Initial"].ToString() ?? string.Empty;
                window.ShowDialog(this);
            }
        }
        private void HandleSubstrateChange1() {
            DataRow[] selectedRows = ProductDataTable.Select($"class001 = '{CategoryListBox1.SelectedItem}' and col_Product_Name = '{CategoryListBox2.SelectedItem}' and col_Product_Type = '{CategoryListBox3.SelectedItem}'");

            if (selectedRows.Length > 0) {
                using SubstrateChange1 window = new();
                window.StrFontName = FontName;
                window.IntFontSize = IntFontSize;
                window.IntRadioBtnFlg = IntRadioBtnFlg;
                window.IntPrintType = Convert.ToInt32(selectedRows[0]["col_Print_Type"]);
                window.IntRegType = Convert.ToInt32(selectedRows[0]["col_Reg_Type"]);
                window.StrProductName = selectedRows[0]["col_Product_Name"].ToString() ?? string.Empty;
                window.StrStockName = selectedRows[0]["col_Stock_Name"].ToString() ?? string.Empty;
                window.StrProductType = selectedRows[0]["col_Product_Type"].ToString() ?? string.Empty;
                window.StrProductModel = selectedRows[0]["col_Product_Model"].ToString() ?? string.Empty;
                window.StrUseSubstrate = selectedRows[0]["col_Use_Substrate"].ToString() ?? string.Empty;
                window.ShowDialog(this);
            }
        }
        // 履歴ボタン処理
        private void History() {
            ResetFields();
            try {
                DataRow[]? selectedRow = null;

                switch (IntRadioBtnFlg) {
                    case 1:
                        selectedRow = ProductDataTable.Select($"class001 = '{CategoryListBox1.SelectedItem}' and col_Product_Name = '{CategoryListBox2.SelectedItem}' and col_Substrate_Name = '{CategoryListBox3.SelectedItem}'");
                        break;
                    case 2:
                    case 3:
                    case 4:
                        selectedRow = ProductDataTable.Select($"class001 = '{CategoryListBox1.SelectedItem}' and col_Product_Name = '{CategoryListBox2.SelectedItem}' and col_Product_Type = '{CategoryListBox3.SelectedItem}'");
                        break;
                }

                if (selectedRow != null && selectedRow.Length > 0) {
                    using HistoryWindow window = new();
                    window.StrFontName = FontName;
                    window.IntFontSize = IntFontSize;
                    window.IntRadioBtnFlg = IntRadioBtnFlg;

                    if (IntRadioBtnFlg == 1) {
                        window.StrProductName = selectedRow[0]["col_Product_Name"].ToString() ?? string.Empty;
                        window.StrSubstrateName = selectedRow[0]["col_Substrate_Name"].ToString() ?? string.Empty;
                        window.StrSubstrateModel = selectedRow[0]["col_Substrate_Model"].ToString() ?? string.Empty;
                    }
                    else {
                        window.StrProductName = selectedRow[0]["col_Product_Name"].ToString() ?? string.Empty;
                        window.StrProductType = selectedRow[0]["col_Product_Type"].ToString() ?? string.Empty;
                        window.StrProductModel = selectedRow[0]["col_Product_Model"].ToString() ?? string.Empty;
                    }

                    window.ShowDialog(this);
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // 処理カテゴリセレクト
        private void CategorySelect(object sender) {
            try {
                RegisterButton.Enabled = false;
                HistoryButton.Enabled = false;
                IntRadioBtnFlg = 0;
                CategoryListBox1.Items.Clear();
                CategoryListBox2.Items.Clear();
                CategoryListBox3.Items.Clear();
                ProductDataTable.Clear();

                RadioButton selectedRadioButton = (RadioButton)sender;
                string strSqlQuery = string.Empty;

                switch (selectedRadioButton.Name) {
                    case "CategoryRadioButton1":
                        IntRadioBtnFlg = 1;
                        strSqlQuery = "SELECT * FROM Substrate WHERE Visual = '1';";
                        break;
                    case "CategoryRadioButton2":
                    case "CategoryRadioButton3":
                        IntRadioBtnFlg = selectedRadioButton.Name == "CategoryRadioButton2" ? 2 : 3;
                        strSqlQuery = "SELECT * FROM Product WHERE Visual = '1';";
                        break;
                    case "CategoryRadioButton4":
                        IntRadioBtnFlg = 4;
                        strSqlQuery = "SELECT * FROM Product WHERE col_Print_Type = '5' OR col_Print_Type = '6';";
                        break;
                    default:
                        break;
                }

                using (SQLiteConnection con = new(GetConnectionString1()))
                using (SQLiteDataAdapter adapter = new(strSqlQuery, con)) {
                    adapter.Fill(ProductDataTable);
                }

                var class001Set = new SortedSet<string>(ProductDataTable.AsEnumerable()
                                                            .Select(row => row.Field<string?>("class001"))
                                                            .Where(classVal => classVal != null)
                                                            .Select(classVal => classVal!));

                foreach (string classVal in class001Set) {
                    CategoryListBox1.Items.Add(classVal);
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // 製品カテゴリセレクト
        private void CategoryListBox1Select() {
            try {
                RegisterButton.Enabled = false;
                HistoryButton.Enabled = false;
                CategoryListBox2.Items.Clear();
                CategoryListBox3.Items.Clear();

                HashSet<string> productNames = new();

                DataRow[] selectedRows = ProductDataTable.Select($"class001 = '{CategoryListBox1.SelectedItem}'", "col_Product_Name ASC");

                foreach (DataRow row in selectedRows) {
                    string productName = row["col_Product_Name"].ToString() ?? string.Empty;
                    if (!string.IsNullOrEmpty(productName)) {
                        productNames.Add(productName);
                    }
                }

                CategoryListBox2.Items.AddRange(productNames.ToArray());

            } catch (Exception ex) {
                MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void CategoryListBox2Select() {
            try {
                RegisterButton.Enabled = false;
                HistoryButton.Enabled = false;
                CategoryListBox3.Items.Clear();

                DataRow[] selectedRows;

                switch (IntRadioBtnFlg) {
                    case 1:
                        selectedRows = ProductDataTable.Select($"class001 = '{CategoryListBox1.SelectedItem}' and col_Product_Name = '{CategoryListBox2.SelectedItem}'", "col_Substrate_Model ASC");
                        HashSet<string> substrateNames = new(selectedRows.AsEnumerable()
                                                                                .Select(x => x.Field<string>("col_Substrate_Name"))
                                                                                .Where(x => x != null)
                                                                                .Select(x => x!));

                        CategoryListBox3.Items.AddRange(substrateNames.ToArray());
                        break;

                    case 2:
                    case 3:
                    case 4:
                        selectedRows = ProductDataTable.Select($"class001 = '{CategoryListBox1.SelectedItem}' and col_Product_Name = '{CategoryListBox2.SelectedItem}'", "col_Product_Model ASC");
                        HashSet<string> productTypes = new(selectedRows.AsEnumerable()
                                                                                .Select(x => x.Field<string>("col_Product_Type"))
                                                                                .Where(x => x != null)
                                                                                .Select(x => x!));

                        CategoryListBox3.Items.AddRange(productTypes.ToArray());
                        break;

                    default:
                        break;
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void CategoryListBox3Select() {
            try {
                RegisterButton.Enabled = true;
                HistoryButton.Enabled = true;

                switch (IntRadioBtnFlg) {
                    case 1:
                        StrSubstrateName = CategoryListBox3.SelectedItem?.ToString() ?? string.Empty;
                        break;

                    case 2:
                    case 3:
                        StrProductType = CategoryListBox3.SelectedItem?.ToString() ?? string.Empty;
                        break;

                    case 4:
                        StrProductType = CategoryListBox3.SelectedItem?.ToString() ?? string.Empty;
                        HistoryButton.Enabled = false;
                        break;
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // QRコード読み取り
        private void CodeScan() {
            try {
                ResetFieldsForCodeScan();

                string[]? arr = ParseQRCodeInput() ?? throw new Exception("QRコード入力に誤りがあります。");
                FetchDataFromDatabase(arr);

                if (ListCategory11.Count >= 2) {
                    ShowDialogWindowForMultipleItems();
                }

                HandleSelectedItem();
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } finally {
                CleanupAfterScan();
            }
        }
        private void ResetFieldsForCodeScan() {
            IntRadioBtnFlg = 0;
            CategoryRadioButton1.Checked = CategoryRadioButton2.Checked = CategoryRadioButton3.Checked = CategoryRadioButton4.Checked = false;
            CategoryListBox1.Items.Clear();
            CategoryListBox2.Items.Clear();
            CategoryListBox3.Items.Clear();
            ListCategory11.Clear();
            ListCategory12.Clear();
            ListCategory13.Clear();
            ListCategory14.Clear();
            Enabled = false;
        }
        private string[]? ParseQRCodeInput() {
            if (QRCodeCheckBox.Checked) {
                return QRCodeTextBox.Text.Split(new string[] { "//" }, StringSplitOptions.None);
            }
            else {
                using (OdbcConnection con = new("DSN=DrSum_PRONES_YD; UID=YD00; PWD=YD00")) {
                    con.Open();
                    using OdbcCommand cmd = new($"SELECT * FROM V_宮崎手配情報 WHERE 手配管理番号 = '{QRCodeTextBox.Text}'", con);
                    using OdbcDataReader dr = cmd.ExecuteReader();
                    while (dr.Read()) {
                        StrProness1 = dr["手配製番"].ToString() ?? string.Empty;
                        StrProness2 = dr["品目番号"].ToString() ?? string.Empty;
                        StrProness3 = dr["品目名称"].ToString() ?? string.Empty;
                        StrProness4 = Convert.ToInt32(dr["手配数"]);
                        StrProness5 = dr["請求先注番"].ToString() ?? string.Empty;
                    }
                }

                if (string.IsNullOrEmpty(StrProness1)) { throw new Exception($"一致する情報がありません。{Environment.NewLine}手配製番:{StrProness1}"); }
                return null;
            }
        }
        private void FetchDataFromDatabase(string[] arr) {
            if (arr != null) {
                ProcessQRCodeInput(arr);
            }

            ProcessCategoryItemData();
        }
        private void ProcessQRCodeInput(string[] arr) {
            StrProness1 = arr[0];
            StrProness2 = arr[1];
            StrProness4 = Convert.ToInt32(arr[2]);
            StrProness5 = arr[3];
        }
        private void ProcessCategoryItemData() {
            StrProness2 = StrProness2.Replace("-SMT", "")
                                     .Replace("-H", "")
                                     .Replace("-GH", "")
                                     .Replace("-ACGH", "-AC")
                                     .Replace("-DCGH", "-DC");

            FetchDataFromSQLite();
        }
        private void FetchDataFromSQLite() {
            using SQLiteConnection con = new(GetConnectionString1());
            con.Open();
            using SQLiteCommand cmd = con.CreateCommand();
            cmd.CommandText = $"SELECT * FROM V_ItemList WHERE col_ItemNumber = '{StrProness2}' OR 'col_ItemNumber:1' = '{StrProness2}'";
            using SQLiteDataReader dr = cmd.ExecuteReader();
            while (dr.Read()) {
                string colItemNumber = dr["col_ItemNumber"].ToString() ?? string.Empty;
                string colItemNumber1 = dr["col_ItemNumber:1"].ToString() ?? string.Empty;

                if (!string.IsNullOrEmpty(colItemNumber)) {
                    string _substrateName = dr["col_Substrate_Name"]?.ToString() ?? string.Empty;
                    string _productName = dr["col_Product_Name"]?.ToString() ?? string.Empty;
                    AddToLists(colItemNumber, _substrateName, _productName, "1");
                }

                if (!string.IsNullOrEmpty(colItemNumber1)) {
                    string _productType = dr["col_Product_Type"]?.ToString() ?? string.Empty;
                    string _productName = dr["col_Product_Name:1"]?.ToString() ?? string.Empty;
                    AddToLists(colItemNumber1, _productType, _productName, "2");
                }
            }
        }
        private void AddToLists(string itemNumber, string category12, string category13, string category14) {
            ListCategory11.Add(itemNumber);
            ListCategory12.Add(category12);
            ListCategory13.Add(category13);
            ListCategory14.Add(category14);
        }
        private void ShowDialogWindowForMultipleItems() {
            using SeveralDialogWindow dlg = new() {
                StrFontName = FontName,
                IntFontSize = IntFontSize,
                ListCategory11 = ListCategory11,
                ListCategory12 = ListCategory12,
                ListCategory13 = ListCategory13,
                ListCategory14 = ListCategory14
            };

            if (dlg.ShowDialog(this) == DialogResult.OK) {
                ListIndex = dlg.SelectedIndex;
            }
        }
        private void HandleSelectedItem() {
            StrCategory11 = ListCategory11[ListIndex];
            StrCategory12 = ListCategory12[ListIndex];
            StrCategory13 = ListCategory13[ListIndex];
            StrCategory14 = ListCategory14[ListIndex];

            switch (StrCategory14) {
                case "1":
                    HandleSubstrateSelection();
                    break;
                case "2":
                    HandleProductSelection();
                    break;
                default:
                    throw new Exception($"一致する情報がありません。{Environment.NewLine}品目番号:{StrProness2}{Environment.NewLine}");
            }
        }
        private void HandleSubstrateSelection() {
            using (SQLiteConnection con = new(GetConnectionString1())) {
                using SQLiteDataAdapter adapter = new("SELECT * FROM Substrate;", con);
                adapter.Fill(ProductDataTable);
            }

            DataRow[] substrateRet = ProductDataTable.Select($"col_Product_Name = '{StrCategory13}' and col_Substrate_Name = '{StrCategory12}'");
            OpenSubstrateRegistrationWindow(substrateRet);
        }
        private void OpenSubstrateRegistrationWindow(DataRow[] substrateRet) {
            using SubstrateRegistrationWindow window = new();
            window.StrFontName = FontName;
            window.IntFontSize = IntFontSize;
            window.StrProductName = substrateRet[0]["col_Product_Name"].ToString() ?? string.Empty;
            window.StrStockName = substrateRet[0]["col_Stock_Name"].ToString() ?? string.Empty;
            window.StrSubstrateName = substrateRet[0]["col_Substrate_Name"].ToString() ?? string.Empty;
            window.StrSubstrateModel = substrateRet[0]["col_Substrate_Model"].ToString() ?? string.Empty;
            window.StrInitial = string.Empty;
            window.IntRegType = Convert.ToInt32(substrateRet[0]["col_Reg_Type"]);
            window.IntPrintType = Convert.ToInt32(substrateRet[0]["col_Print_Type"]);
            window.IntCheckBin = Convert.ToInt32(substrateRet[0]["col_Checkbox"].ToString(), 2);
            window.ShowDialog(this);
        }
        private void HandleProductSelection() {
            using (SQLiteConnection con = new(GetConnectionString1())) {
                using SQLiteDataAdapter adapter = new("SELECT * FROM Product;", con);
                adapter.Fill(ProductDataTable);
            }

            DataRow[] productRet = ProductDataTable.Select($"col_Product_Name = '{StrCategory13}' and col_Product_Type = '{StrCategory12}'");
            OpenProductRegistrationWindow(productRet);
        }
        private void OpenProductRegistrationWindow(DataRow[] productRet) {
            using ProductRegistration1Window window = new();
            window.StrFontName = FontName;
            window.IntFontSize = IntFontSize;
            window.StrProductName = productRet[0]["col_Product_Name"].ToString() ?? string.Empty;
            window.StrStockName = productRet[0]["col_Stock_Name"].ToString() ?? string.Empty;
            window.StrProductType = productRet[0]["col_Product_Type"].ToString() ?? string.Empty;
            window.StrProductModel = productRet[0]["col_Product_Model"].ToString() ?? string.Empty;
            window.StrUseSubstrate = productRet[0]["col_Use_Substrate"].ToString() ?? string.Empty;
            window.StrInitial = productRet[0]["col_Initial"].ToString() ?? string.Empty;
            window.IntRegType = Convert.ToInt32(productRet[0]["col_Reg_Type"]);
            window.IntPrintType = Convert.ToInt32(productRet[0]["col_Print_Type"]);
            window.IntCheckBin = Convert.ToInt32(productRet[0]["col_Checkbox"].ToString(), 2);
            window.IntSerialDigit = Convert.ToInt32(productRet[0]["col_Serial_Digit"]);
            window.ShowDialog(this);
        }
        private void CleanupAfterScan() {
            QRCodeTextBox.Text = string.Empty;
            Enabled = true;
            QRCodeTextBox.Focus();
        }
        // フォント変更
        private void FontChange(object sender) {
            RadioButton radioButton = (RadioButton)sender;
            int fontSize;

            switch (radioButton.Name) {
                case "FontSize9RadioButton":
                    fontSize = 9;
                    break;
                case "FontSize12RadioButton":
                    fontSize = 12;
                    break;
                case "FontSize16RadioButton":
                    fontSize = 16;
                    break;
                default:
                    return;
            }

            IntFontSize = fontSize;
            Font = new Font(FontName, IntFontSize);
        }

        private void MainWindow_Load(object sender, EventArgs e) { LoadEvents(); }
        private void 終了ToolStripMenuItem_Click(object sender, EventArgs e) { Close(); }
        private void RegisterButton_Click(object sender, EventArgs e) { Registration(); }
        private void HistoryButton_Click(object sender, EventArgs e) { History(); }
        private void CategoryListBox1_SelectedIndexChanged(object sender, EventArgs e) { CategoryListBox1Select(); }
        private void CategoryListBox2_SelectedIndexChanged(object sender, EventArgs e) { CategoryListBox2Select(); }
        private void CategoryListBox3_SelectedIndexChanged(object sender, EventArgs e) { CategoryListBox3Select(); }
        private void CategoryRadioButton_CheckedChanged(object sender, EventArgs e) { CategorySelect(sender); }
        private void FontSize_CheckedChanged(object sender, EventArgs e) { FontChange(sender); }
        private void QRCodeTextBox_KeyDown(object sender, KeyEventArgs e) {
            if (e.KeyCode != Keys.Enter) { return; }
            if (string.IsNullOrWhiteSpace(QRCodeTextBox.Text)) { return; }
            CodeScan();
        }
        private void QRCodeButton_Click(object sender, EventArgs e) { CodeScan(); }
    }
}
