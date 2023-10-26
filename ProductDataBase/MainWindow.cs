using System.Data;
using System.Data.Odbc;
using System.Data.SQLite;

namespace ProductDataBase {
    public partial class MainWindow : Form {
        public MainWindow() {
            InitializeComponent();
        }

        private readonly DataTable dt = new();
        public string strFontName = "Meiryo UI";
        public int intFontSize = 9;
        public int intRadioBtnFlg = 0;
        public int intRegType = 0;
        public int intPrintType = 0;
        public int intSerialDigit = 0;
        public string? strProductName = string.Empty;
        public string? strStockName = string.Empty;
        public string? strProductType = string.Empty;
        public string? strProductModel = string.Empty;
        public string? strSubstrateName = string.Empty;
        public string? strSubstrateModel = string.Empty;
        public string? strUseSubstrate = string.Empty;
        public string? strInitial = string.Empty;
        public int intCheckBin = 0;
        public string? strProness1 = string.Empty;
        public string? strProness2 = string.Empty;
        public string? strProness3 = string.Empty;
        public int strProness4 = 0;
        public string? strProness5 = string.Empty;
        private int ListIndex = 0;
        public string strCategory11 = string.Empty;
        public string strCategory12 = string.Empty;
        public string strCategory13 = string.Empty;
        public string strCategory14 = string.Empty;
        public List<string> ListCategory11 = new();
        public List<string> ListCategory12 = new();
        public List<string> ListCategory13 = new();
        public List<string> ListCategory14 = new();

        public string GetConnectionString1() {
            return new SQLiteConnectionStringBuilder() { DataSource = "./db/information.db" }.ToString();
        }
        public string GetConnectionString2() {
            return new SQLiteConnectionStringBuilder() { DataSource = "./db/registration.db" }.ToString();
        }

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
                MessageBox.Show(ex.Message);
            }
        }
        private void ResetFields() {
            intRegType = intPrintType = intSerialDigit = ListIndex = 0;
            strProductName = strProductType = strProductModel = strSubstrateName = strSubstrateModel = strUseSubstrate = strInitial = string.Empty;
            intCheckBin = 0;
            strProness1 = strProness2 = strProness3 = strProness5 = string.Empty;
            strProness4 = 0;
        }
        private void Registration() {
            try {
                ResetFields();

                switch (intRadioBtnFlg) {
                    case 1:
                        DataRow[] ret1 = dt.Select($"class001 = '{CategoryListBox1.SelectedItem}' and col_Product_Name = '{CategoryListBox2.SelectedItem}' and col_Substrate_Name = '{CategoryListBox3.SelectedItem}'");

                        if (ret1.Length > 0) {
                            strProductName = ret1[0]["col_Product_Name"].ToString();
                            strStockName = ret1[0]["col_Stock_Name"].ToString();
                            strSubstrateName = ret1[0]["col_Substrate_Name"].ToString();
                            strSubstrateModel = ret1[0]["col_Substrate_Model"].ToString();
                            intRegType = Convert.ToInt32(ret1[0]["col_Reg_Type"]);
                            intPrintType = Convert.ToInt32(ret1[0]["col_Print_Type"]);
                            intCheckBin = Convert.ToInt32(ret1[0]["col_Checkbox"].ToString(), 2);

                            SubstrateRegistrationWindow substrateRegistrationWindow = new();
                            substrateRegistrationWindow.ShowDialog(this);
                        }
                        break;
                    case 2:
                    case 3:
                    case 4:
                        DataRow[] ret2 = dt.Select($"class001 = '{CategoryListBox1.SelectedItem}' and col_Product_Name = '{CategoryListBox2.SelectedItem}' and col_Product_Type = '{CategoryListBox3.SelectedItem}'");

                        if (ret2.Length > 0) {
                            strProductName = ret2[0]["col_Product_Name"].ToString();
                            strStockName = ret2[0]["col_Stock_Name"].ToString();
                            strProductType = ret2[0]["col_Product_Type"].ToString();
                            intRegType = Convert.ToInt32(ret2[0]["col_Reg_Type"]);
                            intPrintType = Convert.ToInt32(ret2[0]["col_Print_Type"]);
                            intSerialDigit = Convert.ToInt32(ret2[0]["col_Serial_Digit"]);
                            strProductModel = ret2[0]["col_Product_Model"].ToString();
                            intCheckBin = Convert.ToInt32(ret2[0]["col_Checkbox"].ToString(), 2);
                            strUseSubstrate = ret2[0]["col_Use_Substrate"].ToString();
                            strInitial = ret2[0]["col_Initial"].ToString();

                            switch (intRadioBtnFlg) {
                                case 2:
                                    ProductRegistration1Window productRegistration1Window = new();
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
                MessageBox.Show(ex.Message);
            } finally {
            }
        }
        private void History() {
            ResetFields();
            try {
                switch (intRadioBtnFlg) {
                    case 1:
                        DataRow[] ret1 = dt.Select($"class001 = '{CategoryListBox1.SelectedItem}' and col_Product_Name = '{CategoryListBox2.SelectedItem}' and col_Substrate_Name = '{CategoryListBox3.SelectedItem}'");
                        strProductName = ret1[0]["col_Product_Name"].ToString();
                        strSubstrateName = ret1[0]["col_Substrate_Name"].ToString();
                        strSubstrateModel = ret1[0]["col_Substrate_Model"].ToString();
                        break;

                    case 2:
                    case 3:
                    case 4:
                        DataRow[] ret2 = dt.Select($"class001 = '{CategoryListBox1.SelectedItem}' and col_Product_Name = '{CategoryListBox2.SelectedItem}' and col_Product_Type = '{CategoryListBox3.SelectedItem}'");
                        strProductName = ret2[0]["col_Product_Name"].ToString();
                        strProductType = ret2[0]["col_Product_Type"].ToString();
                        strProductModel = ret2[0]["col_Product_Model"].ToString();
                        break;
                }

                History history = new();
                history.ShowDialog(this);
            } catch (Exception ex) {
                MessageBox.Show(ex.Message);
            }

        }
        private void CategoryListBox1Select() {
            try {
                RegisterButton.Enabled = false;
                HistoryButton.Enabled = false;
                CategoryListBox2.Items.Clear();
                CategoryListBox3.Items.Clear();

                HashSet<string> productNames = new();

                DataRow[] dr1 = dt.Select($"class001 = '{CategoryListBox1.SelectedItem}'", "col_Product_Name ASC");

                foreach (DataRow row in dr1) {
                    object productNameObj = row["col_Product_Name"];
                    string? productName = productNameObj != DBNull.Value ? productNameObj.ToString() : null;

                    if (!string.IsNullOrEmpty(productName)) {
                        productNames.Add(productName);
                    }
                }

                CategoryListBox2.Items.AddRange(productNames.ToArray());

            } catch (Exception ex) {
                MessageBox.Show(ex.Message);
            } finally {
            }
        }
        private void CategoryListBox2Select() {
            try {
                RegisterButton.Enabled = false;
                HistoryButton.Enabled = false;
                CategoryListBox3.Items.Clear();

                DataRow[] dr2;

                switch (intRadioBtnFlg) {
                    case 1:
                        dr2 = dt.Select($"class001 = '{CategoryListBox1.SelectedItem}' and col_Product_Name = '{CategoryListBox2.SelectedItem}'", "col_Substrate_Model ASC");
                        HashSet<string> col_Substrate_Name = new(dr2.AsEnumerable().Select(x => x.Field<string>("col_Substrate_Name")).Where(x => x != null).Select(x => x!));

                        foreach (string d in col_Substrate_Name) {
                            CategoryListBox3.Items.Add(d);
                        }
                        break;

                    case 2:
                    case 3:
                    case 4:
                        dr2 = dt.Select($"class001 = '{CategoryListBox1.SelectedItem}' and col_Product_Name = '{CategoryListBox2.SelectedItem}'", "col_Product_Model ASC");
                        HashSet<string> col_Product_Type = new(dr2.AsEnumerable().Select(x => x.Field<string>("col_Product_Type")).Where(x => x != null).Select(x => x!));

                        foreach (string d in col_Product_Type) {
                            CategoryListBox3.Items.Add(d);
                        }
                        break;
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message);
            }
        }
        private void CategoryListBox3Select() {
            try {
                RegisterButton.Enabled = true;
                HistoryButton.Enabled = true;

                switch (intRadioBtnFlg) {
                    case 1:
                        strSubstrateName = CategoryListBox3.SelectedItem?.ToString();
                        break;

                    case 2:
                    case 3:
                        strProductType = CategoryListBox3.SelectedItem?.ToString();
                        break;

                    case 4:
                        strProductType = CategoryListBox3.SelectedItem?.ToString();
                        HistoryButton.Enabled = false;
                        break;
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message);
            }
        }
        private void CodeScan() {
            try {
                intRadioBtnFlg = 0;
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

                    strProness1 = arr1[0];
                    strProness2 = arr1[1];
                    strProness4 = Convert.ToInt32(arr1[2]);
                    strProness5 = arr1[3];
                }
                else {
                    using (OdbcConnection con = new("DSN=DrSum_PRONES_YD; UID=YD00; PWD=YD00")) {
                        con.Open();
                        using OdbcCommand cmd = new($"SELECT * FROM V_宮崎手配情報 WHERE 手配管理番号 = '{QRCodeTextBox.Text}'", con);
                        using OdbcDataReader dr = cmd.ExecuteReader();
                        while (dr.Read()) {
                            strProness1 = dr["手配製番"].ToString();
                            strProness2 = dr["品目番号"].ToString();
                            strProness3 = dr["品目名称"].ToString();
                            strProness4 = Convert.ToInt32(dr["手配数"].ToString());
                            strProness5 = dr["請求先注番"].ToString();
                        }
                    }

                    if (string.IsNullOrEmpty(strProness1)) {
                        MessageBox.Show($"一致する情報がありません。{Environment.NewLine}手配製番:{strProness1}");
                        return;
                    }
                }

                // 文字列の置換
                if (strProness2 != null) { strProness2 = strProness2.Replace("-SMT", "").Replace("-H", "").Replace("-GH", "").Replace("-ACGH", "-AC").Replace("-DCGH", "-DC"); }

                using (var con = new SQLiteConnection(GetConnectionString1())) {
                    con.Open();
                    using var cmd = con.CreateCommand();
                    cmd.CommandText = $"SELECT * FROM V_ItemList WHERE col_ItemNumber = '{strProness2}' OR col_ItemNumber:1 = '{strProness2}'";
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

                if (ListCategory11.Count == 0) {
                    MessageBox.Show($"品目番号:[{strProness2}] と一致するデータがありませんでした。");
                    return;
                }

                if (ListCategory11.Count >= 2) {
                    using var dlg = new SeveralDialog();
                    if (dlg.ShowDialog(this) == DialogResult.OK) {
                        ListIndex = dlg.SelectedIndex;
                    }
                }

                strCategory11 = ListCategory11[ListIndex];
                strCategory12 = ListCategory12[ListIndex];
                strCategory13 = ListCategory13[ListIndex];
                strCategory14 = ListCategory14[ListIndex];

                switch (strCategory14) {
                    case "1":
                        using (var con = new SQLiteConnection(GetConnectionString1())) {
                            using var adapter = new SQLiteDataAdapter("SELECT * FROM Substrate;", con);
                            adapter.Fill(dt);
                        }

                        DataRow[] substrateRet = dt.Select($"col_Product_Name = '{strCategory13}' and col_Substrate_Name = '{strCategory12}'");

                        strProductName = substrateRet[0]["col_Product_Name"].ToString();
                        strStockName = substrateRet[0]["col_Stock_Name"].ToString();
                        strSubstrateName = substrateRet[0]["col_Substrate_Name"].ToString();
                        strSubstrateModel = substrateRet[0]["col_Substrate_Model"].ToString();
                        intRegType = (int)substrateRet[0]["col_Reg_Type"];
                        intPrintType = (int)substrateRet[0]["col_Print_Type"];
                        intCheckBin = Convert.ToInt32(substrateRet[0]["col_Checkbox"].ToString(), 2);
                        strInitial = string.Empty;

                        SubstrateRegistrationWindow substrateRegistrationWindow = new();
                        substrateRegistrationWindow.ShowDialog(this);
                        break;

                    case "2":
                        using (var con = new SQLiteConnection(GetConnectionString1())) {
                            using var adapter = new SQLiteDataAdapter("SELECT * FROM Product;", con);
                            adapter.Fill(dt);
                        }

                        DataRow[] productRet = dt.Select($"col_Product_Name = '{strCategory13}' and col_Product_Type = '{strCategory12}'");

                        intRegType = (int)productRet[0]["col_Reg_Type"];
                        strProductName = productRet[0]["col_Product_Name"].ToString();
                        strStockName = productRet[0]["col_Stock_Name"].ToString();
                        strProductType = productRet[0]["col_Product_Type"].ToString();
                        intPrintType = (int)productRet[0]["col_Print_Type"];
                        intSerialDigit = (int)productRet[0]["col_Serial_Digit"];
                        strProductModel = productRet[0]["col_Product_Model"].ToString();
                        intCheckBin = Convert.ToInt32(productRet[0]["col_Checkbox"].ToString(), 2);
                        strUseSubstrate = productRet[0]["col_Use_Substrate"].ToString();
                        strInitial = productRet[0]["col_Initial"].ToString();

                        ProductRegistration1Window productRegistration1Window = new();
                        productRegistration1Window.ShowDialog(this);
                        break;

                    default:
                        MessageBox.Show($"一致する情報がありません。{Environment.NewLine}品目番号:{strProness2}{Environment.NewLine}");
                        break;
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message);
            } finally {
                QRCodeTextBox.Text = string.Empty;
                Enabled = true;
                QRCodeTextBox.Focus();
            }
        }
        private void CategorySelect(object sender) {
            try {
                RegisterButton.Enabled = false;
                HistoryButton.Enabled = false;
                intRadioBtnFlg = 0;
                CategoryListBox1.Items.Clear();
                CategoryListBox2.Items.Clear();
                CategoryListBox3.Items.Clear();
                dt.Clear();

                var radioButton = (RadioButton)sender;
                string strSqlQuery = string.Empty;

                switch (radioButton.Name) {
                    case "CategoryRadioButton1":
                        intRadioBtnFlg = 1;
                        strSqlQuery = "SELECT * FROM Substrate;";
                        break;

                    case "CategoryRadioButton2":
                        intRadioBtnFlg = 2;
                        strSqlQuery = "SELECT * FROM Product;";
                        break;

                    case "CategoryRadioButton3":
                        intRadioBtnFlg = 3;
                        strSqlQuery = "SELECT * FROM Product;";
                        break;

                    case "CategoryRadioButton4":
                        intRadioBtnFlg = 4;
                        strSqlQuery = "SELECT * FROM Product WHERE col_Print_Type = '5';";
                        break;
                }

                using SQLiteConnection con = new(GetConnectionString1());
                using SQLiteDataAdapter adapter = new(strSqlQuery, con);
                adapter.Fill(dt);

                SortedSet<string> class001 = new(dt.AsEnumerable().Select(x => x.Field<string?>("class001")).Where(x => x != null).Select(x => x!));

                foreach (string b in class001) {
                    CategoryListBox1.Items.Add(b);
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message);
            }
        }
        private void FontChange(object sender) {
            var radioButton = (RadioButton)sender;
            switch (radioButton.Name) {
                case "FontSize9RadioButton":
                    intFontSize = 9;
                    break;
                case "FontSize12RadioButton":
                    intFontSize = 12;
                    break;
                case "FontSize16RadioButton":
                    intFontSize = 16;
                    break;
            }
            Font = new Font(strFontName, intFontSize);
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
