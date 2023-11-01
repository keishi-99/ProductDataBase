using System.Data;
using System.Data.Odbc;
using System.Data.SQLite;

namespace ProductDataBase {
    public partial class MainWindow : Form {

        public DataTable Dt { get; } = new();
        public string StrFontName { get; set; } = "Meiryo UI";
        public int IntFontSize { get; set; } = 9;
        public int IntRadioBtnFlg { get; set; }
        public int IntRegType { get; set; }
        public int IntPrintType { get; set; }
        public int IntSerialDigit { get; set; }
        public string? StrProductName { get; set; }
        public string? StrStockName { get; set; }
        public string? StrProductType { get; set; }
        public string? StrProductModel { get; set; }
        public string? StrSubstrateName { get; set; }
        public string? StrSubstrateModel { get; set; }
        public string? StrUseSubstrate { get; set; }
        public string? StrInitial { get; set; }
        public int IntCheckBin { get; set; }
        public string? StrProness1 { get; set; }
        public string? StrProness2 { get; set; }
        public string? StrProness3 { get; set; }
        public int StrProness4 { get; set; }
        public string? StrProness5 { get; set; }
        private int ListIndex { get; set; }
        public string? StrCategory11 { get; set; }
        public string? StrCategory12 { get; set; }
        public string? StrCategory13 { get; set; }
        public string? StrCategory14 { get; set; }
        public List<string> ListCategory11 { get; } = new();
        public List<string> ListCategory12 { get; } = new();
        public List<string> ListCategory13 { get; } = new();
        public List<string> ListCategory14 { get; } = new();

        public MainWindow() {
            InitializeComponent();
        }

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
                DateTime d = DateTime.Now;
                string filepath = "./db/registration.db";
                string bk_filepath = $".\\bak\\{d.Year}\\{d.Month:00}_bak_{d.Year}-{d.Month:00}-{d.Day:00}.db";

                if (!System.IO.File.Exists(bk_filepath)) {
                    System.IO.File.Copy(filepath, bk_filepath, true);
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
                        DataRow[] ret1 = Dt.Select($"class001 = '{CategoryListBox1.SelectedItem}' and col_Product_Name = '{CategoryListBox2.SelectedItem}' and col_Substrate_Name = '{CategoryListBox3.SelectedItem}'");

                        if (ret1.Length > 0) {
                            SubstrateRegistrationWindow substrateRegistrationWindow = new() {
                                StrProductName = ret1[0]["col_Product_Name"].ToString(),
                                StrStockName = ret1[0]["col_Stock_Name"].ToString(),
                                StrSubstrateName = ret1[0]["col_Substrate_Name"].ToString(),
                                StrSubstrateModel = ret1[0]["col_Substrate_Model"].ToString(),
                                IntRegType = Convert.ToInt32(ret1[0]["col_Reg_Type"]),
                                IntPrintType = Convert.ToInt32(ret1[0]["col_Print_Type"]),
                                IntCheckBin = Convert.ToInt32(ret1[0]["col_Checkbox"].ToString(), 2)
                            };

                            substrateRegistrationWindow.ShowDialog(this);
                        }
                        break;
                    case 2:
                    case 3:
                    case 4:
                        DataRow[] ret2 = Dt.Select($"class001 = '{CategoryListBox1.SelectedItem}' and col_Product_Name = '{CategoryListBox2.SelectedItem}' and col_Product_Type = '{CategoryListBox3.SelectedItem}'");

                        if (ret2.Length > 0) {

                            switch (IntRadioBtnFlg) {
                                case 2:
                                    ProductRegistration1Window productRegistration1Window = new() {
                                        StrProductName = ret2[0]["col_Product_Name"].ToString(),
                                        StrStockName = ret2[0]["col_Stock_Name"].ToString(),
                                        StrProductType = ret2[0]["col_Product_Type"].ToString(),
                                        IntRegType = Convert.ToInt32(ret2[0]["col_Reg_Type"]),
                                        IntPrintType = Convert.ToInt32(ret2[0]["col_Print_Type"]),
                                        IntSerialDigit = Convert.ToInt32(ret2[0]["col_Serial_Digit"]),
                                        StrProductModel = ret2[0]["col_Product_Model"].ToString(),
                                        IntCheckBin = Convert.ToInt32(ret2[0]["col_Checkbox"].ToString(), 2),
                                        StrUseSubstrate = ret2[0]["col_Use_Substrate"].ToString(),
                                        StrInitial = ret2[0]["col_Initial"].ToString()
                                    };
                                    productRegistration1Window.ShowDialog(this);
                                    break;
                                case 3:
                                    RePrint rePrint = new();
                                    rePrint.ShowDialog(this);
                                    break;
                                case 4:
                                    SubstrateChange1 substrateChange1 = new();
                                    substrateChange1.ShowDialog(this);
                                    break;
                            }
                        }
                        break;
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } finally {
            }
        }
        // 履歴ボタン処理
        private void History() {
            ResetFields();
            try {
                switch (IntRadioBtnFlg) {
                    case 1:
                        DataRow[] ret1 = Dt.Select($"class001 = '{CategoryListBox1.SelectedItem}' and col_Product_Name = '{CategoryListBox2.SelectedItem}' and col_Substrate_Name = '{CategoryListBox3.SelectedItem}'");
                        StrProductName = ret1[0]["col_Product_Name"].ToString();
                        StrSubstrateName = ret1[0]["col_Substrate_Name"].ToString();
                        StrSubstrateModel = ret1[0]["col_Substrate_Model"].ToString();
                        break;

                    case 2:
                    case 3:
                    case 4:
                        DataRow[] ret2 = Dt.Select($"class001 = '{CategoryListBox1.SelectedItem}' and col_Product_Name = '{CategoryListBox2.SelectedItem}' and col_Product_Type = '{CategoryListBox3.SelectedItem}'");
                        StrProductName = ret2[0]["col_Product_Name"].ToString();
                        StrProductType = ret2[0]["col_Product_Type"].ToString();
                        StrProductModel = ret2[0]["col_Product_Model"].ToString();
                        break;
                }

                History history = new();
                history.ShowDialog(this);
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
                Dt.Clear();

                var radioButton = (RadioButton)sender;
                string strSqlQuery = string.Empty;

                switch (radioButton.Name) {
                    case "CategoryRadioButton1":
                        IntRadioBtnFlg = 1;
                        strSqlQuery = "SELECT * FROM Substrate;";
                        break;

                    case "CategoryRadioButton2":
                        IntRadioBtnFlg = 2;
                        strSqlQuery = "SELECT * FROM Product;";
                        break;

                    case "CategoryRadioButton3":
                        IntRadioBtnFlg = 3;
                        strSqlQuery = "SELECT * FROM Product;";
                        break;

                    case "CategoryRadioButton4":
                        IntRadioBtnFlg = 4;
                        strSqlQuery = "SELECT * FROM Product WHERE col_Print_Type = '5';";
                        break;
                }

                using SQLiteConnection con = new(GetConnectionString1());
                using SQLiteDataAdapter adapter = new(strSqlQuery, con);
                adapter.Fill(Dt);

                SortedSet<string> class001 = new(Dt.AsEnumerable().Select(x => x.Field<string?>("class001")).Where(x => x != null).Select(x => x!));

                foreach (string b in class001) {
                    CategoryListBox1.Items.Add(b);
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

                DataRow[] dr1 = Dt.Select($"class001 = '{CategoryListBox1.SelectedItem}'", "col_Product_Name ASC");

                foreach (DataRow row in dr1) {
                    object productNameObj = row["col_Product_Name"];
                    string? productName = productNameObj != DBNull.Value ? productNameObj.ToString() : null;

                    if (!string.IsNullOrEmpty(productName)) {
                        productNames.Add(productName);
                    }
                }

                CategoryListBox2.Items.AddRange(productNames.ToArray());

            } catch (Exception ex) {
                MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } finally {
            }
        }
        private void CategoryListBox2Select() {
            try {
                RegisterButton.Enabled = false;
                HistoryButton.Enabled = false;
                CategoryListBox3.Items.Clear();

                DataRow[] dr2;

                switch (IntRadioBtnFlg) {
                    case 1:
                        dr2 = Dt.Select($"class001 = '{CategoryListBox1.SelectedItem}' and col_Product_Name = '{CategoryListBox2.SelectedItem}'", "col_Substrate_Model ASC");
                        HashSet<string> col_Substrate_Name = new(dr2.AsEnumerable().Select(x => x.Field<string>("col_Substrate_Name")).Where(x => x != null).Select(x => x!));

                        foreach (string d in col_Substrate_Name) {
                            CategoryListBox3.Items.Add(d);
                        }
                        break;

                    case 2:
                    case 3:
                    case 4:
                        dr2 = Dt.Select($"class001 = '{CategoryListBox1.SelectedItem}' and col_Product_Name = '{CategoryListBox2.SelectedItem}'", "col_Product_Model ASC");
                        HashSet<string> col_Product_Type = new(dr2.AsEnumerable().Select(x => x.Field<string>("col_Product_Type")).Where(x => x != null).Select(x => x!));

                        foreach (string d in col_Product_Type) {
                            CategoryListBox3.Items.Add(d);
                        }
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
                        StrSubstrateName = CategoryListBox3.SelectedItem?.ToString();
                        break;

                    case 2:
                    case 3:
                        StrProductType = CategoryListBox3.SelectedItem?.ToString();
                        break;

                    case 4:
                        StrProductType = CategoryListBox3.SelectedItem?.ToString();
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

                if (QRCodeCheckBox.Checked) {
                    string[] arr1 = QRCodeTextBox.Text.Split(new string[] { "//" }, StringSplitOptions.None);

                    StrProness1 = arr1[0];
                    StrProness2 = arr1[1];
                    StrProness4 = Convert.ToInt32(arr1[2]);
                    StrProness5 = arr1[3];
                }
                else {
                    using (OdbcConnection con = new("DSN=DrSum_PRONES_YD; UID=YD00; PWD=YD00")) {
                        con.Open();
                        using OdbcCommand cmd = new($"SELECT * FROM V_宮崎手配情報 WHERE 手配管理番号 = '{QRCodeTextBox.Text}'", con);
                        using OdbcDataReader dr = cmd.ExecuteReader();
                        while (dr.Read()) {
                            StrProness1 = dr["手配製番"].ToString();
                            StrProness2 = dr["品目番号"].ToString();
                            StrProness3 = dr["品目名称"].ToString();
                            StrProness4 = Convert.ToInt32(dr["手配数"].ToString());
                            StrProness5 = dr["請求先注番"].ToString();
                        }
                    }

                    if (string.IsNullOrEmpty(StrProness1)) { throw new Exception($"一致する情報がありません。{Environment.NewLine}手配製番:{StrProness1}"); }
                }

                // 文字列の置換
                if (StrProness2 != null) { StrProness2 = StrProness2.Replace("-SMT", "").Replace("-H", "").Replace("-GH", "").Replace("-ACGH", "-AC").Replace("-DCGH", "-DC"); }

                using (var con = new SQLiteConnection(GetConnectionString1())) {
                    con.Open();
                    using var cmd = con.CreateCommand();
                    cmd.CommandText = $"SELECT * FROM V_ItemList WHERE col_ItemNumber = '{StrProness2}' OR col_ItemNumber:1 = '{StrProness2}'";
                    using var dr = cmd.ExecuteReader();
                    while (dr.Read()) {
                        string colItemNumber = dr["col_ItemNumber"].ToString() ?? string.Empty;
                        string colItemNumber1 = dr["col_ItemNumber:1"].ToString() ?? string.Empty;

                        if (!string.IsNullOrEmpty(colItemNumber)) {
                            ListCategory11.Add(colItemNumber);
                            ListCategory12.Add(dr["col_Substrate_Name"].ToString() ?? string.Empty);
                            ListCategory13.Add(dr["col_Product_Name"].ToString() ?? string.Empty);
                            ListCategory14.Add("1");
                        }

                        if (!string.IsNullOrEmpty(colItemNumber1)) {
                            ListCategory11.Add(colItemNumber1);
                            ListCategory12.Add(dr["col_Product_Type"].ToString() ?? string.Empty);
                            ListCategory13.Add(dr["col_Product_Name:1"].ToString() ?? string.Empty);
                            ListCategory14.Add("2");
                        }
                    }
                }

                if (ListCategory11.Count == 0) { throw new Exception($"品目番号:[{StrProness2}] と一致するデータがありませんでした。"); }

                if (ListCategory11.Count >= 2) {
                    using var dlg = new SeveralDialog();
                    if (dlg.ShowDialog(this) == DialogResult.OK) {
                        ListIndex = dlg.SelectedIndex;
                    }
                }

                StrCategory11 = ListCategory11[ListIndex];
                StrCategory12 = ListCategory12[ListIndex];
                StrCategory13 = ListCategory13[ListIndex];
                StrCategory14 = ListCategory14[ListIndex];

                switch (StrCategory14) {
                    case "1":
                        using (var con = new SQLiteConnection(GetConnectionString1())) {
                            using var adapter = new SQLiteDataAdapter("SELECT * FROM Substrate;", con);
                            adapter.Fill(Dt);
                        }

                        DataRow[] substrateRet = Dt.Select($"col_Product_Name = '{StrCategory13}' and col_Substrate_Name = '{StrCategory12}'");

                        SubstrateRegistrationWindow substrateRegistrationWindow = new() {

                            StrProductName = substrateRet[0]["col_Product_Name"].ToString(),
                            StrStockName = substrateRet[0]["col_Stock_Name"].ToString(),
                            StrSubstrateName = substrateRet[0]["col_Substrate_Name"].ToString(),
                            StrSubstrateModel = substrateRet[0]["col_Substrate_Model"].ToString(),
                            IntRegType = (int)substrateRet[0]["col_Reg_Type"],
                            IntPrintType = (int)substrateRet[0]["col_Print_Type"],
                            IntCheckBin = Convert.ToInt32(substrateRet[0]["col_Checkbox"].ToString(), 2),
                            StrInitial = string.Empty,
                        };

                        substrateRegistrationWindow.ShowDialog(this);
                        break;

                    case "2":
                        using (var con = new SQLiteConnection(GetConnectionString1())) {
                            using var adapter = new SQLiteDataAdapter("SELECT * FROM Product;", con);
                            adapter.Fill(Dt);
                        }

                        DataRow[] productRet = Dt.Select($"col_Product_Name = '{StrCategory13}' and col_Product_Type = '{StrCategory12}'");

                        IntRegType = (int)productRet[0]["col_Reg_Type"];
                        StrProductName = productRet[0]["col_Product_Name"].ToString();
                        StrStockName = productRet[0]["col_Stock_Name"].ToString();
                        StrProductType = productRet[0]["col_Product_Type"].ToString();
                        IntPrintType = (int)productRet[0]["col_Print_Type"];
                        IntSerialDigit = (int)productRet[0]["col_Serial_Digit"];
                        StrProductModel = productRet[0]["col_Product_Model"].ToString();
                        IntCheckBin = Convert.ToInt32(productRet[0]["col_Checkbox"].ToString(), 2);
                        StrUseSubstrate = productRet[0]["col_Use_Substrate"].ToString();
                        StrInitial = productRet[0]["col_Initial"].ToString();

                        ProductRegistration1Window productRegistration1Window = new();
                        productRegistration1Window.ShowDialog(this);
                        break;

                    default: throw new Exception($"一致する情報がありません。{Environment.NewLine}品目番号:{StrProness2}{Environment.NewLine}");
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } finally {
                QRCodeTextBox.Text = string.Empty;
                Enabled = true;
                QRCodeTextBox.Focus();
            }
        }
        // フォント変更
        private void FontChange(object sender) {
            var radioButton = (RadioButton)sender;
            switch (radioButton.Name) {
                case "FontSize9RadioButton":
                    IntFontSize = 9;
                    break;
                case "FontSize12RadioButton":
                    IntFontSize = 12;
                    break;
                case "FontSize16RadioButton":
                    IntFontSize = 16;
                    break;
            }
            Font = new Font(StrFontName, IntFontSize);
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
    }
}
