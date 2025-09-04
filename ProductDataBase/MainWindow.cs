using Microsoft.Extensions.Configuration;
using ProductDatabase.Other;
using System.Data;
using System.Data.Odbc;
using System.Data.SQLite;
using System.Text.RegularExpressions;

namespace ProductDatabase {
    public partial class MainWindow : Form {

        public int RadioButtonNumber { get; set; }
        private DataTable MainDataTable { get; } = new();

        public class ProductInformation {
            public int ProductID { get; set; }
            public string CategoryName { get; set; } = string.Empty;
            public string ProductName { get; set; } = string.Empty;
            public string StockName { get; set; } = string.Empty;
            public string ProductType { get; set; } = string.Empty;
            public string ProductModel { get; set; } = string.Empty;
            public string SubstrateName { get; set; } = string.Empty;
            public string SubstrateModel { get; set; } = string.Empty;
            public string UseSubstrate { get; set; } = string.Empty;
            public string Initial { get; set; } = string.Empty;
            public int SerialType { get; set; }
            public int RevisionGroup { get; set; }
            public int CheckBin { get; set; }
            public string Proness1 { get; set; } = string.Empty;
            public string Proness2 { get; set; } = string.Empty;
            public string Proness3 { get; set; } = string.Empty;
            public int Proness4 { get; set; }
            public string Proness5 { get; set; } = string.Empty;
            public List<string> CategoryItemNumber { get; set; } = [];
            public List<string> CategoryProductName { get; set; } = [];
            public List<string> CategorySubstrateName { get; set; } = [];
            public List<string> CategoryProductType { get; set; } = [];
            public List<string> CategoryType { get; set; } = [];
            public List<string> PersonList { get; set; } = [];
            public string FontName { get; } = "Meiryo UI";
            public int FontSize { get; set; } = 9;

            public string OrderNumber { get; set; } = string.Empty;
            public string ProductNumber { get; set; } = string.Empty;
            public string RegDate { get; set; } = string.Empty;
            public string Person { get; set; } = string.Empty;
            public string Revision { get; set; } = string.Empty;
            public string Comment { get; set; } = string.Empty;
            public string SerialFirst { get; set; } = string.Empty;
            public string SerialLast { get; set; } = string.Empty;

            private int _regType;
            private int _printType;

            public int RegType {
                get => _regType;
                set {
                    _regType = value;
                    UpdatePrintFlags();
                }
            }

            public int PrintType {
                get => _printType;
                set {
                    _printType = value;
                    UpdatePrintFlags();
                }
            }

            public bool IsLabelPrint { get; private set; }
            public bool IsBarcodePrint { get; private set; }
            public bool IsListPrint { get; private set; }
            public bool IsCheckSheetPrint { get; private set; }
            public bool IsLast4Digits { get; private set; }
            public bool IsRegType9 { get; private set; }
            public bool IsSerialGeneration { get; private set; }
            public bool IsUnderlinePrint { get; private set; }

            private void UpdatePrintFlags() {
                UpdateRegTypeFlags();
                UpdatePrintTypeFlags();
            }

            private void UpdateRegTypeFlags() {
                IsRegType9 = RegType == 9;
                IsSerialGeneration = RegType is 1 or 2 or 3 or 9;
            }

            private void UpdatePrintTypeFlags() {
                IsLabelPrint = PrintType is 1 or 3 or 4 or 5 or 6 or 7 or 9;
                IsBarcodePrint = PrintType is 2 or 3 or 8;
                IsListPrint = (PrintType is 5 or 6) && !IsRegType9;
                IsCheckSheetPrint = (PrintType is 6 or 7 or 8) && !IsRegType9;
                IsLast4Digits = (PrintType == 9) && !IsRegType9;
                IsUnderlinePrint = (PrintType == 4) && !IsRegType9;
            }

            public int Quantity { get; set; }
            public int SerialFirstNumber { get; set; }
            public int SerialLastNumber { get; set; }

            public DataTable ProductDataTable { get; } = new();
            public DataTable SubstrateDataTable { get; } = new();

            public int SerialDigit => SerialType switch {
                3 or 101 or 102 => 3,
                4 => 4,
                _ => 0
            };
        }
        public static class Auth {
            private static bool s_isAdministrator = false;
            public static bool IsAdministrator {
                get => s_isAdministrator; set => s_isAdministrator = value;
            }

            private static bool s_isAuthorizedUser = false;
            public static bool IsAuthorizedUser {
                get => s_isAuthorizedUser; set => s_isAuthorizedUser = value;
            }
        }

        public ProductInformation ProductInfo { get; set; } = new();

        public MainWindow() {
            InitializeComponent();
        }

        public static string GetConnectionInformation() {
            var informationPath = Path.Combine(Environment.CurrentDirectory, "db", "information.db");
            if (!File.Exists(informationPath)) { throw new FileNotFoundException("ファイルが見つかりません。", informationPath); }
            var u = new Uri(informationPath);
            if (u.IsUnc) {
                informationPath = @"\" + informationPath; // UNCパス
            }
            return new SQLiteConnectionStringBuilder() { DataSource = informationPath }.ToString();
        }
        public static string GetConnectionRegistration() {
            var registrationPath = Path.Combine(Environment.CurrentDirectory, "db", "registration.db");
            if (!File.Exists(registrationPath)) { throw new FileNotFoundException("ファイルが見つかりません。", registrationPath); }
            var u = new Uri(registrationPath);
            if (u.IsUnc) {
                registrationPath = @"\" + registrationPath; // UNCパス
            }
            return new SQLiteConnectionStringBuilder() { DataSource = registrationPath }.ToString();
        }

        // ロードイベント
        private void LoadEvents() {
            try {
                // JSONファイルのパス
                var jsonFilePath = Path.Combine(Environment.CurrentDirectory, "Config", "General", "appsettings.json");

                // パスのディレクトリ部分を取得
                var basePath = Path.GetDirectoryName(jsonFilePath);
                if (string.IsNullOrEmpty(basePath) || !Directory.Exists(basePath)) {
                    throw new DirectoryNotFoundException($"The directory '{basePath}' does not exist.");
                }
                // JSONファイルを読み込む
                var config = new ConfigurationBuilder()
                    .SetBasePath(basePath)
                    .AddJsonFile(Path.GetFileName(jsonFilePath), optional: false, reloadOnChange: true)
                    .Build();
                // CloneFolderPathを取得
                CommonUtils.s_networkPath = config["NetworkFolderPath"] ?? throw new Exception("フォルダが設定されてません。");
                if (string.IsNullOrEmpty(CommonUtils.s_networkPath)) { throw new Exception("フォルダが設定されてません。"); }
                if (!Directory.Exists(CommonUtils.s_networkPath)) {
                    throw new DirectoryNotFoundException($"フォルダ '{CommonUtils.s_networkPath}' が見つかりません。");
                }

                // その日の backupファイルがない場合バックアップ作成
                var d = DateTime.Now;
                var backupDir = Path.Combine(CommonUtils.s_networkPath, "db", "backup", $"{d.Year}", $"{d.Month:00}");
                var backupFilePath = Path.Combine(CommonUtils.s_networkPath, "db", "backup", $"{d.Year}", $"{d.Month:00}", $"_bak_{d.Year}-{d.Month:00}-{d.Day:00}.db");
                var registrationPath = Path.Combine(Environment.CurrentDirectory, "db", "registration.db");

                if (!File.Exists(backupFilePath)) {
                    Directory.CreateDirectory(backupDir);  // ディレクトリが存在しない場合に作成
                    File.Copy(registrationPath, backupFilePath, false);
                }

                using SQLiteConnection con = new(GetConnectionInformation());
                using (SQLiteDataAdapter adapter = new("SELECT * FROM Product;", con)) { adapter.Fill(ProductInfo.ProductDataTable); }
                using (SQLiteDataAdapter adapter = new("SELECT * FROM Substrate;", con)) { adapter.Fill(ProductInfo.SubstrateDataTable); }

                // DB1へ接続し担当者取得
                ProductInfo.PersonList.Clear();
                con.Open();
                using var cmd = con.CreateCommand();
                // テーブル検索SQL - 担当者をPersonListへ追加
                cmd.CommandText = "SELECT * FROM Person ORDER BY _rowid_ ASC";
                using var dr = cmd.ExecuteReader();
                while (dr.Read()) {
                    ProductInfo.PersonList.Add($"{dr["PersonName"]}");
                }

                // 認証ユーザー名を取得
                var userNames = config.GetSection("AuthorizedUsers").Get<string[]>() ?? [];
                // 現在のユーザー名がリストに含まれるかチェック
                Auth.IsAuthorizedUser = userNames?.Any(name => name.Equals(Environment.UserName, StringComparison.OrdinalIgnoreCase)) ?? false;

                // 管理者ユーザー名を取得
                var adminUserNames = config.GetSection("Administrator").Get<string[]>() ?? [];
                // 現在のユーザー名がリストに含まれるかチェック
                Auth.IsAdministrator = adminUserNames?.Any(name => name.Equals(Environment.UserName, StringComparison.OrdinalIgnoreCase)) ?? false;

                QRCodePanel.Enabled = Auth.IsAuthorizedUser;

                RegisterButton.Enabled = false;
                HistoryButton.Enabled = false;

                CategoryRadioButton1.Checked = false;
                CategoryRadioButton2.Checked = false;
                CategoryRadioButton3.Checked = false;
                CategoryRadioButton4.Checked = false;

                CategoryListBox1.Items.Clear();
                CategoryListBox2.Items.Clear();
                CategoryListBox3.Items.Clear();
                MainDataTable.Clear();

                this.Activate();
                QRCodeTextBox.Select();
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void ResetFields() {
            ProductInfo = new ProductInformation();

            using SQLiteConnection con = new(GetConnectionInformation());
            using (SQLiteDataAdapter adapter = new("SELECT * FROM Product;", con)) { adapter.Fill(ProductInfo.ProductDataTable); }
            using (SQLiteDataAdapter adapter = new("SELECT * FROM Substrate;", con)) { adapter.Fill(ProductInfo.SubstrateDataTable); }

            // DB1へ接続し担当者取得
            ProductInfo.PersonList.Clear();
            con.Open();
            using var cmd = con.CreateCommand();
            // テーブル検索SQL - 担当者をPersonListへ追加
            cmd.CommandText = "SELECT * FROM Person ORDER BY _rowid_ ASC";
            using var dr = cmd.ExecuteReader();
            while (dr.Read()) {
                ProductInfo.PersonList.Add($"{dr["PersonName"]}");
            }
        }
        // 登録ボタン処理
        private void Registration() {
            try {
                ResetFields();

                if (CategoryListBox3.SelectedIndex == -1) { return; }
                switch (RadioButtonNumber) {
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

                QRCodeTextBox.Text = string.Empty;
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void HandleSubstrateRegistration() {
            var selectedRows = MainDataTable.Select($"CategoryName = '{CategoryListBox1.SelectedItem}' AND ProductName = '{CategoryListBox2.SelectedItem}' AND SubstrateName = '{CategoryListBox3.SelectedItem}'");

            if (selectedRows.Length > 0) {
                ProductInfo.CategoryName = selectedRows[0]["CategoryName"].ToString() ?? string.Empty;
                ProductInfo.ProductName = selectedRows[0]["ProductName"].ToString() ?? string.Empty;
                ProductInfo.StockName = selectedRows[0]["StockName"].ToString() ?? string.Empty;
                ProductInfo.SubstrateName = selectedRows[0]["SubstrateName"].ToString() ?? string.Empty;
                ProductInfo.SubstrateModel = selectedRows[0]["SubstrateModel"].ToString() ?? string.Empty;
                ProductInfo.RegType = Convert.ToInt32(selectedRows[0]["RegType"] ?? throw new Exception("RegType is null"));
                ProductInfo.PrintType = Convert.ToInt32(selectedRows[0]["PrintType"] ?? throw new Exception("PrintType is null"));
                ProductInfo.CheckBin = Convert.ToInt32(selectedRows[0]["Checkbox"].ToString() ?? throw new Exception("Checkbox is null"), 2);
                using SubstrateRegistrationWindow window = new(ProductInfo);
                window.ShowDialog(this);
            }
        }
        private void HandleProductRegistration1() {
            var selectedRows = MainDataTable.Select($"CategoryName = '{CategoryListBox1.SelectedItem}' AND ProductName = '{CategoryListBox2.SelectedItem}' AND ProductType = '{CategoryListBox3.SelectedItem}'");

            if (selectedRows.Length > 0) {
                ProductInfo.CategoryName = selectedRows[0]["CategoryName"].ToString() ?? string.Empty;
                ProductInfo.ProductName = selectedRows[0]["ProductName"].ToString() ?? string.Empty;
                ProductInfo.StockName = selectedRows[0]["StockName"].ToString() ?? string.Empty;
                ProductInfo.ProductType = selectedRows[0]["ProductType"].ToString() ?? string.Empty;
                ProductInfo.RegType = Convert.ToInt32(selectedRows[0]["RegType"] ?? throw new Exception("RegType is null"));
                ProductInfo.PrintType = Convert.ToInt32(selectedRows[0]["PrintType"] ?? throw new Exception("PrintType is null"));
                ProductInfo.SerialType = Convert.ToInt32(selectedRows[0]["SerialType"] ?? throw new Exception("SerialType is null"));
                ProductInfo.ProductModel = selectedRows[0]["ProductModel"].ToString() ?? string.Empty;
                ProductInfo.CheckBin = Convert.ToInt32(selectedRows[0]["Checkbox"].ToString() ?? throw new Exception("Checkbox is null"), 2);
                ProductInfo.UseSubstrate = selectedRows[0]["UseSubstrate"].ToString() ?? string.Empty;
                ProductInfo.Initial = selectedRows[0]["Initial"].ToString() ?? string.Empty;
                ProductInfo.RevisionGroup = Convert.ToInt32(selectedRows[0]["RevisionGroup"] ?? throw new Exception("RevisionGroup is null"));
                using ProductRegistration1Window window = new(ProductInfo);
                window.ShowDialog(this);
            }
        }
        private void HandleRePrint() {
            var selectedRows = MainDataTable.Select($"CategoryName = '{CategoryListBox1.SelectedItem}' AND ProductName = '{CategoryListBox2.SelectedItem}' AND ProductType = '{CategoryListBox3.SelectedItem}'");

            if (selectedRows.Length > 0) {
                ProductInfo.CategoryName = selectedRows[0]["CategoryName"].ToString() ?? string.Empty;
                ProductInfo.ProductName = selectedRows[0]["ProductName"].ToString() ?? string.Empty;
                ProductInfo.ProductType = selectedRows[0]["ProductType"].ToString() ?? string.Empty;
                ProductInfo.RegType = Convert.ToInt32(selectedRows[0]["RegType"] ?? throw new Exception("RegType is null"));
                ProductInfo.PrintType = Convert.ToInt32(selectedRows[0]["PrintType"] ?? throw new Exception("PrintType is null"));
                ProductInfo.SerialType = Convert.ToInt32(selectedRows[0]["SerialType"] ?? throw new Exception("SerialType is null"));
                ProductInfo.ProductModel = selectedRows[0]["ProductModel"].ToString() ?? string.Empty;
                ProductInfo.CheckBin = Convert.ToInt32(selectedRows[0]["Checkbox"].ToString() ?? throw new Exception("Checkbox is null"), 2);
                ProductInfo.Initial = selectedRows[0]["Initial"].ToString() ?? string.Empty;
                ProductInfo.RevisionGroup = Convert.ToInt32(selectedRows[0]["RevisionGroup"] ?? throw new Exception("RevisionGroup is null"));
                using RePrintWindow window = new(ProductInfo);
                window.ShowDialog(this);
            }
        }
        private void HandleSubstrateChange1() {
            var selectedRows = MainDataTable.Select($"CategoryName = '{CategoryListBox1.SelectedItem}' AND ProductName = '{CategoryListBox2.SelectedItem}' AND ProductType = '{CategoryListBox3.SelectedItem}'");

            if (selectedRows.Length > 0) {
                ProductInfo.CategoryName = selectedRows[0]["CategoryName"].ToString() ?? string.Empty;
                ProductInfo.PrintType = Convert.ToInt32(selectedRows[0]["PrintType"] ?? throw new Exception("PrintType is null"));
                ProductInfo.RegType = Convert.ToInt32(selectedRows[0]["RegType"] ?? throw new Exception("RegType is null"));
                ProductInfo.ProductName = selectedRows[0]["ProductName"].ToString() ?? string.Empty;
                ProductInfo.StockName = selectedRows[0]["StockName"].ToString() ?? string.Empty;
                ProductInfo.ProductType = selectedRows[0]["ProductType"].ToString() ?? string.Empty;
                ProductInfo.ProductModel = selectedRows[0]["ProductModel"].ToString() ?? string.Empty;
                ProductInfo.RevisionGroup = Convert.ToInt32(selectedRows[0]["RevisionGroup"] ?? throw new Exception("RevisionGroup is null"));
                ProductInfo.UseSubstrate = selectedRows[0]["UseSubstrate"].ToString() ?? string.Empty;
                using SubstrateChange1 window = new(ProductInfo);
                window.ShowDialog(this);
            }
        }
        // 履歴ボタン処理
        private void History() {
            ResetFields();
            try {
                DataRow[]? selectedRow = null;

                var listBox2 = CategoryListBox2.SelectedIndex == -1 ? string.Empty : $"AND ProductName = '{CategoryListBox2.SelectedItem}'";

                var listBox3 = RadioButtonNumber switch {
                    1 => CategoryListBox3.SelectedIndex == -1 ? string.Empty : $"AND SubstrateName = '{CategoryListBox3.SelectedItem}'",
                    2 or 3 or 4 => CategoryListBox3.SelectedIndex == -1 ? string.Empty : $"AND ProductType = '{CategoryListBox3.SelectedItem}'",
                    _ => string.Empty
                };
                selectedRow = MainDataTable.Select($"CategoryName = '{CategoryListBox1.SelectedItem}' {listBox2} {listBox3}");

                if (selectedRow != null && selectedRow.Length > 0) {
                    switch (RadioButtonNumber) {
                        case 1:
                            ProductInfo.CategoryName = selectedRow[0]["CategoryName"].ToString() ?? string.Empty;

                            if (!string.IsNullOrEmpty(listBox2)) {
                                ProductInfo.ProductName = selectedRow[0]["ProductName"].ToString() ?? string.Empty;
                                ProductInfo.StockName = selectedRow[0]["StockName"].ToString() ?? string.Empty;
                            }

                            if (!string.IsNullOrEmpty(listBox3)) {
                                ProductInfo.SubstrateName = selectedRow[0]["SubstrateName"].ToString() ?? string.Empty;
                                ProductInfo.SubstrateModel = selectedRow[0]["SubstrateModel"].ToString() ?? string.Empty;
                                ProductInfo.PrintType = Convert.ToInt32(selectedRow[0]["PrintType"] ?? throw new Exception("RegType is null"));
                                ProductInfo.RegType = Convert.ToInt32(selectedRow[0]["RegType"] ?? throw new Exception("RegType is null"));
                            }
                            break;
                        case 2:
                        case 3:
                            ProductInfo.CategoryName = selectedRow[0]["CategoryName"].ToString() ?? string.Empty;

                            if (!string.IsNullOrEmpty(listBox2)) {
                                ProductInfo.ProductName = selectedRow[0]["ProductName"].ToString() ?? string.Empty;
                                ProductInfo.StockName = selectedRow[0]["StockName"].ToString() ?? string.Empty;
                            }

                            if (!string.IsNullOrEmpty(listBox3)) {
                                ProductInfo.ProductType = selectedRow[0]["ProductType"].ToString() ?? string.Empty;
                                ProductInfo.ProductModel = selectedRow[0]["ProductModel"].ToString() ?? string.Empty;
                                ProductInfo.PrintType = Convert.ToInt32(selectedRow[0]["PrintType"] ?? throw new Exception("RegType is null"));
                                ProductInfo.RegType = Convert.ToInt32(selectedRow[0]["RegType"] ?? throw new Exception("RegType is null"));
                            }
                            break;
                    }

                    using HistoryWindow window = new(ProductInfo, RadioButtonNumber);
                    window.ShowDialog(this);
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // 処理カテゴリセレクト
        private void CategorySelect(object sender) {
            // Tagプロパティを使用してラジオボタンの識別子を取得
            Dictionary<string, (int Number, string Sql)> radioButtonMap = new() {
                ["1"] = (1, "SELECT * FROM Substrate WHERE Visible = 1 ORDER BY SortNumber ASC;"),
                ["2"] = (2, "SELECT * FROM Product WHERE Visible = 1 ORDER BY SortNumber ASC;"),
                ["3"] = (3, "SELECT * FROM Product WHERE Visible = 1 AND PrintType != 0 ORDER BY SortNumber ASC;"),
                ["4"] = (4, "SELECT * FROM Product WHERE Visible = 1 AND (PrintType = 5 OR PrintType = 6) ORDER BY SortNumber ASC;")
            };

            try {
                RegisterButton.Enabled = false;
                HistoryButton.Enabled = false;
                CategoryListBox1.Items.Clear();
                CategoryListBox2.Items.Clear();
                CategoryListBox3.Items.Clear();
                MainDataTable.Clear();

                if (sender is RadioButton selectedRadioButton && radioButtonMap.TryGetValue(selectedRadioButton.Tag?.ToString() ?? "", out var map)) {
                    RadioButtonNumber = map.Number;
                    using var con = new SQLiteConnection(GetConnectionInformation());
                    using var adapter = new SQLiteDataAdapter(map.Sql, con);
                    adapter.Fill(MainDataTable);
                }

                // CategoryName 列の重複を削除し、ソートする
                var categoryNames = MainDataTable.AsEnumerable()
                    .Select(row => row.Field<string?>("CategoryName") ?? string.Empty)
                    .Where(name => !string.IsNullOrWhiteSpace(name))
                    .Distinct()
                    .ToList();

                // リストボックスにアイテムを追加する
                CategoryListBox1.Items.AddRange([.. categoryNames]);

            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // 製品カテゴリセレクト
        private void CategoryListBox1Select() {
            try {
                RegisterButton.Enabled = false;
                if (RadioButtonNumber == 4) { HistoryButton.Enabled = false; }
                else { HistoryButton.Enabled = true; }
                CategoryListBox2.Items.Clear();
                CategoryListBox3.Items.Clear();

                HashSet<string> productNames = [];

                var selectedRows = MainDataTable.Select($"CategoryName = '{CategoryListBox1.SelectedItem}'", "ProductName ASC");

                foreach (var row in selectedRows) {
                    var productName = row["ProductName"].ToString() ?? throw new Exception("ProductName is null");
                    productNames.Add(productName);
                }

                CategoryListBox2.Items.AddRange([.. productNames]);

            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void CategoryListBox2Select() {
            try {
                RegisterButton.Enabled = false;
                if (RadioButtonNumber == 4) { HistoryButton.Enabled = false; }
                else { HistoryButton.Enabled = true; }
                CategoryListBox3.Items.Clear();

                DataRow[] selectedRows;

                switch (RadioButtonNumber) {
                    case 1:
                        selectedRows = MainDataTable.Select($"CategoryName = '{CategoryListBox1.SelectedItem}' AND ProductName = '{CategoryListBox2.SelectedItem}'", "SubstrateName ASC");
                        HashSet<string> substrateNames = [.. selectedRows.AsEnumerable()
                            .Select(x => x.Field<string>("SubstrateName"))
                            .Where(x => x != null)
                            .Select(x => x!)];

                        CategoryListBox3.Items.AddRange([.. substrateNames]);
                        break;

                    case 2:
                    case 3:
                    case 4:
                        selectedRows = MainDataTable.Select($"CategoryName = '{CategoryListBox1.SelectedItem}' AND ProductName = '{CategoryListBox2.SelectedItem}'", "ProductType ASC");
                        HashSet<string> productTypes = [.. selectedRows.AsEnumerable()
                            .Select(x => x.Field<string>("ProductType"))
                            .Where(x => x != null)
                            .Select(x => x!)];

                        CategoryListBox3.Items.AddRange([.. productTypes]);
                        break;

                    default:
                        break;
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void CategoryListBox3Select() {
            try {
                if (CategoryListBox3.SelectedIndex == -1) { return; }
                RegisterButton.Enabled = Auth.IsAuthorizedUser;
                HistoryButton.Enabled = true;

                switch (RadioButtonNumber) {
                    case 1:
                        ProductInfo.SubstrateName = CategoryListBox3.SelectedItem?.ToString() ?? string.Empty;
                        break;

                    case 2:
                    case 3:
                        ProductInfo.ProductType = CategoryListBox3.SelectedItem?.ToString() ?? string.Empty;
                        break;

                    case 4:
                        ProductInfo.ProductType = CategoryListBox3.SelectedItem?.ToString() ?? string.Empty;
                        HistoryButton.Enabled = false;
                        break;
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // QRコード読み取り
        private void CodeScan() {
            try {
                if (string.IsNullOrWhiteSpace(QRCodeTextBox.Text)) { return; }
                ResetFieldsForCodeScan();

                if (RadioButtonQR.Checked) {
                    if (textToUpperCheckBox.Checked) { QRCodeTextBox.Text = QRCodeTextBox.Text.ToUpper(); }
                    ParseQRCodeInput();
                }
                else if (RadioButtonBarcode.Checked) {
                    BarcodeInput();
                }

                ProcessCategoryItemData();
                FetchDataFromSQLite();

                var listIndex = 0;
                if (ProductInfo.CategoryItemNumber.Count >= 2) {
                    listIndex = ShowDialogWindowForMultipleItems();
                }

                if (listIndex == -1) { return; }
                HandleSelectedItem(listIndex);
                QRCodeTextBox.Text = string.Empty;
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } finally {
                CleanupAfterScan();
            }
        }
        private void ResetFieldsForCodeScan() {
            CategoryRadioButton1.Checked = CategoryRadioButton2.Checked = CategoryRadioButton3.Checked = CategoryRadioButton4.Checked = false;
            CategoryListBox1.Items.Clear();
            CategoryListBox2.Items.Clear();
            CategoryListBox3.Items.Clear();
            ProductInfo.CategoryItemNumber.Clear();
            ProductInfo.CategoryProductType.Clear();
            ProductInfo.CategoryProductName.Clear();
            ProductInfo.CategoryType.Clear();
            ResetFields();
            Enabled = false;
        }
        private void ParseQRCodeInput() {
            try {
                string[] separator = ["//"];
                var arr = QRCodeTextBox.Text.Split(separator, StringSplitOptions.None);
                if (arr.Length == 1) {
                    ProductInfo.Proness2 = QRCodeTextBox.Text;
                    return;
                }
                if (arr.Length != 4) { throw new Exception("QRコードが正しくありません。"); }
                ProductInfo.Proness1 = arr[0];
                ProductInfo.Proness2 = arr[1];
                ProductInfo.Proness4 = int.TryParse(arr[2], out var result) ? result : throw new Exception("数量に数値以外が入力されています。");
                ProductInfo.Proness5 = arr[3];
            } catch (Exception ex) {
                throw new Exception($"{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー{Environment.NewLine}{ex.Message}");
            }
        }
        private void BarcodeInput() {
            using (OdbcConnection con = new("DSN=DrSum_PRONES_YD; UID=YD00; PWD=YD00")) {
                con.Open();
                using OdbcCommand cmd = new($"SELECT * FROM V_宮崎手配情報 WHERE 手配管理番号 = '{QRCodeTextBox.Text}';", con);
                using var dr = cmd.ExecuteReader();
                while (dr.Read()) {
                    ProductInfo.Proness1 = dr["手配製番"].ToString() ?? string.Empty;
                    ProductInfo.Proness2 = dr["品目番号"].ToString() ?? string.Empty;
                    ProductInfo.Proness3 = dr["品目名称"].ToString() ?? string.Empty;
                    ProductInfo.Proness4 = Convert.ToInt32(dr["手配数"] ?? throw new Exception("手配数 is null"));
                    ProductInfo.Proness5 = dr["請求先注番"].ToString() ?? string.Empty;
                }
            }
            if (string.IsNullOrEmpty(ProductInfo.Proness1)) { throw new Exception($"一致する情報がありません。{Environment.NewLine}手配管理番号:{QRCodeTextBox.Text}"); }
        }
        private void ProcessCategoryItemData() {
            var pattern = @"-(?:SMT|H|GH).*";
            var result = Regex.Replace(ProductInfo.Proness2, pattern, "");
            ProductInfo.Proness2 = result
                .Replace("-ACGH", "-AC")
                .Replace("-DCGH", "-DC");
        }
        private void FetchDataFromSQLite() {
            using SQLiteConnection con = new(GetConnectionInformation());
            con.Open();
            using var cmd = con.CreateCommand();

            cmd.CommandText =
                $"""
                SELECT
                    s.SubItemNumber,
                    s.SubstrateName,
                    s.ProductName AS sName,
                    p.ProductName AS pName,
                    p.ProductType,
                    p.ProItemNumber
                FROM
                    Substrate AS s
                FULL JOIN
                    Product AS p
                ON
                    s.SubItemNumber = p.ProItemNumber
                WHERE
                    s.SubItemNumber LIKE '%'|| @StrProness2 ||'%'
                OR
                    p.ProItemNumber LIKE '%'|| @StrProness2 ||'%'
                ;
                """;

            cmd.Parameters.Add("@StrProness2", DbType.String).Value = ProductInfo.Proness2;
            using var dr = cmd.ExecuteReader();
            if (!dr.HasRows) { throw new Exception($"品目番号が見つかりません。\n品目番号:[{ProductInfo.Proness2}]"); }
            while (dr.Read()) {
                var colSubItemNumber = dr["SubItemNumber"].ToString() ?? string.Empty;
                var colProItemNumber = dr["ProItemNumber"].ToString() ?? string.Empty;

                if (!string.IsNullOrWhiteSpace(colSubItemNumber)) {
                    var productName = dr["sName"]?.ToString() ?? string.Empty;
                    var substrateName = dr["SubstrateName"]?.ToString() ?? string.Empty;
                    AddToLists(colSubItemNumber, productName, string.Empty, substrateName, "1");
                }

                if (!string.IsNullOrWhiteSpace(colProItemNumber)) {
                    var productName = dr["pName"]?.ToString() ?? string.Empty;
                    var productType = dr["ProductType"]?.ToString() ?? string.Empty;
                    AddToLists(colProItemNumber, productName, productType, string.Empty, "2");
                }
            }
        }
        private void AddToLists(string itemNumber, string productName, string productType, string substrateName, string type) {
            ProductInfo.CategoryItemNumber.Add(itemNumber);
            ProductInfo.CategoryProductName.Add(productName);
            ProductInfo.CategoryProductType.Add(productType);
            ProductInfo.CategorySubstrateName.Add(substrateName);
            ProductInfo.CategoryType.Add(type);
        }
        private int ShowDialogWindowForMultipleItems() {
            using SeveralDialogWindow window = new(ProductInfo);
            window.ShowDialog(this);
            return window.SelectedIndex;
        }
        private void HandleSelectedItem(int listIndex) {
            var productName = ProductInfo.CategoryProductName[listIndex];
            var productType = ProductInfo.CategoryProductType[listIndex];
            var substrateName = ProductInfo.CategorySubstrateName[listIndex];
            var type = ProductInfo.CategoryType[listIndex];

            switch (type) {
                case "1":
                    HandleSubstrateSelection(productName, substrateName);
                    break;
                case "2":
                    HandleProductSelection(productName, productType);
                    break;
                default:
                    throw new Exception($"一致する情報がありません。{Environment.NewLine}品目番号:{ProductInfo.Proness2}{Environment.NewLine}");
            }
        }
        private void HandleSubstrateSelection(string productName, string substrateName) {
            using (SQLiteConnection con = new(GetConnectionInformation())) {
                using SQLiteDataAdapter adapter = new("SELECT * FROM Substrate;", con);
                adapter.Fill(MainDataTable);
            }

            var substrateRet = MainDataTable.Select($"ProductName = '{productName}' AND SubstrateName = '{substrateName}'");
            OpenSubstrateRegistrationWindow(substrateRet);
        }
        private void OpenSubstrateRegistrationWindow(DataRow[] substrateRet) {
            ProductInfo.CategoryName = substrateRet[0]["CategoryName"].ToString() ?? string.Empty;
            ProductInfo.ProductName = substrateRet[0]["ProductName"].ToString() ?? string.Empty;
            ProductInfo.StockName = substrateRet[0]["StockName"].ToString() ?? string.Empty;
            ProductInfo.SubstrateName = substrateRet[0]["SubstrateName"].ToString() ?? string.Empty;
            ProductInfo.SubstrateModel = substrateRet[0]["SubstrateModel"].ToString() ?? string.Empty;
            ProductInfo.Initial = string.Empty;
            ProductInfo.RegType = Convert.ToInt32(substrateRet[0]["RegType"] ?? throw new Exception("RegType is null"));
            ProductInfo.PrintType = Convert.ToInt32(substrateRet[0]["PrintType"] ?? throw new Exception("PrintType is null"));
            ProductInfo.CheckBin = Convert.ToInt32(substrateRet[0]["Checkbox"].ToString() ?? throw new Exception("Checkbox is null"), 2);
            using SubstrateRegistrationWindow window = new(ProductInfo);
            window.ShowDialog(this);
        }
        private void HandleProductSelection(string productName, string productType) {
            using (SQLiteConnection con = new(GetConnectionInformation())) {
                using SQLiteDataAdapter adapter = new("SELECT * FROM Product;", con);
                adapter.Fill(MainDataTable);
            }

            var productRet = MainDataTable.Select($"ProductName = '{productName}' AND ProductType = '{productType}'");
            OpenProductRegistrationWindow(productRet);
        }
        private void OpenProductRegistrationWindow(DataRow[] productRet) {
            ProductInfo.CategoryName = productRet[0]["CategoryName"].ToString() ?? string.Empty;
            ProductInfo.ProductName = productRet[0]["ProductName"].ToString() ?? string.Empty;
            ProductInfo.StockName = productRet[0]["StockName"].ToString() ?? string.Empty;
            ProductInfo.ProductType = productRet[0]["ProductType"].ToString() ?? string.Empty;
            ProductInfo.ProductModel = productRet[0]["ProductModel"].ToString() ?? string.Empty;
            ProductInfo.UseSubstrate = productRet[0]["UseSubstrate"].ToString() ?? string.Empty;
            ProductInfo.Initial = productRet[0]["Initial"].ToString() ?? string.Empty;
            ProductInfo.RevisionGroup = Convert.ToInt32(productRet[0]["RevisionGroup"] ?? throw new Exception("RevisionGroup is null"));
            ProductInfo.RegType = Convert.ToInt32(productRet[0]["RegType"] ?? throw new Exception("RegType is null"));
            ProductInfo.PrintType = Convert.ToInt32(productRet[0]["PrintType"] ?? throw new Exception("PrintType is null"));
            ProductInfo.CheckBin = Convert.ToInt32(productRet[0]["Checkbox"].ToString() ?? throw new Exception("Checkbox is null"), 2);
            ProductInfo.SerialType = Convert.ToInt32(productRet[0]["SerialType"] ?? throw new Exception("SerialType is null"));
            using ProductRegistration1Window window = new(ProductInfo);
            window.ShowDialog(this);
        }
        private void CleanupAfterScan() {
            Enabled = true;
            QRCodeTextBox.Focus();
        }
        // フォント変更
        private void FontChange(object sender) {
            var radioButton = (RadioButton)sender;
            int fontSize;

            switch (radioButton.Name) {
                case "FontSize9RadioButton":
                    fontSize = 9;
                    break;
                case "FontSize12RadioButton":
                    fontSize = 12;
                    break;
                case "FontSize14RadioButton":
                    fontSize = 14;
                    break;
                default:
                    return;
            }

            ProductInfo.FontSize = fontSize;
            Font = new System.Drawing.Font(ProductInfo.FontName, ProductInfo.FontSize);
        }
        // excel
        private void OpenExcel(string filePath) {
            //Excel実行ファイルの場所を取得
            var xlApp = new Microsoft.Office.Interop.Excel.Application();
            var excelFullPath = System.IO.Path.Combine(xlApp.Path, "EXCEL.EXE"); //フルパス作成
            xlApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
            // ファイルを実行
            System.Diagnostics.Process.Start(excelFullPath, filePath);
        }

        private void MainWindow_Load(object sender, EventArgs e) { LoadEvents(); }
        private void ReloadToolStripMenuItem_Click(object sender, EventArgs e) { LoadEvents(); }
        private void ConfigReportToolStripMenuItem_Click(object sender, EventArgs e) {
            var reportConfigPath = Path.Combine(Environment.CurrentDirectory, "config", "General", "Excel", "ConfigReport.xlsx");
            OpenExcel(reportConfigPath);
        }
        private void ConfigListToolStripMenuItem_Click(object sender, EventArgs e) {
            var listConfigPath = Path.Combine(Environment.CurrentDirectory, "config", "General", "Excel", "ConfigList.xlsx");
            OpenExcel(listConfigPath);
        }
        private void ConfigCheckSheetToolStripMenuItem_Click(object sender, EventArgs e) {
            var checkSheetConfigPath = Path.Combine(Environment.CurrentDirectory, "config", "General", "Excel", "ConfigCheckSheet.xlsx");
            OpenExcel(checkSheetConfigPath);
        }
        private void ConfigSubstrateInformationToolStripMenuItem_Click(object sender, EventArgs e) {
            var checkSheetConfigPath = Path.Combine(Environment.CurrentDirectory, "config", "General", "Excel", "ConfigSubstrateInformation.xlsx");
            OpenExcel(checkSheetConfigPath);
        }
        private void 終了ToolStripMenuItem_Click(object sender, EventArgs e) { Close(); }
        private void RegisterButton_Click(object sender, EventArgs e) { Registration(); }
        private void HistoryButton_Click(object sender, EventArgs e) { History(); }
        private void CategoryListBox1_SelectedIndexChanged(object sender, EventArgs e) { CategoryListBox1Select(); }
        private void CategoryListBox2_SelectedIndexChanged(object sender, EventArgs e) { CategoryListBox2Select(); }
        private void CategoryListBox3_SelectedIndexChanged(object sender, EventArgs e) { CategoryListBox3Select(); }
        private void CategoryRadioButton_CheckedChanged(object sender, EventArgs e) { CategorySelect(sender); }
        private void FontSize_CheckedChanged(object sender, EventArgs e) { FontChange(sender); }
        private void CategoryListBox3_KeyDown(object sender, KeyEventArgs e) {
            if (e.KeyCode != Keys.Enter) { return; }
            Registration();
        }
        private void QRCodeTextBox_KeyDown(object sender, KeyEventArgs e) {
            if (e.KeyCode != Keys.Enter) { return; }
            CodeScan();
        }
        private void QRCodeButton_Click(object sender, EventArgs e) { CodeScan(); }
        private void QRCodeTextBox_Enter(object sender, EventArgs e) { CommonUtils.Keyboard.CapsDisable(); }

    }
}
