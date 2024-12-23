using Microsoft.Extensions.Configuration;
using System.Data;
using System.Data.Odbc;
using System.Data.SQLite;

namespace ProductDatabase {
    public partial class MainWindow : Form {

        private static readonly string s_networkPath = string.Empty; // ClonePathを保持する静的変数
        private static readonly string[]? s_userNames = []; // ClonePathを保持する静的変数
        // 静的コンストラクタでClonePathを読み込む
        static MainWindow() {
            try {
                // JSONファイルのパス
                var jsonFilePath = Path.Combine(Environment.CurrentDirectory, "Config", "general", "appsettings.json");

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
                s_networkPath = config["CloneFolderPath"] ?? throw new Exception("フォルダが設定されてません。");
                if (string.IsNullOrEmpty(s_networkPath)) { throw new Exception("フォルダが設定されてません。"); }
                if (!Directory.Exists(s_networkPath)) {
                    throw new DirectoryNotFoundException($"フォルダ '{s_networkPath}' が見つかりません。");
                }

                s_userNames = config.GetSection("AuthorizedUsers").Get<string[]>() ?? [];

            } catch (Exception ex) {
                MessageBox.Show($"設定の読み込み中にエラーが発生しました: {ex.Message}", "設定エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ログ作成
        public static class Logger {
            private static readonly string s_logDirectory = Path.Combine(s_databasePath, "db", "logs"); // ログを保存するディレクトリ
            private static readonly object s_lockObject = new();

            /// <summary>
            /// 作業ログを追記します。
            /// </summary>
            /// <param name="message">記録する作業内容</param>
            public static void AppendLog(string message) {
                try {
                    lock (s_lockObject) {
                        // ディレクトリが存在しない場合は作成
                        if (!Directory.Exists(s_logDirectory)) {
                            Directory.CreateDirectory(s_logDirectory);
                        }

                        //// 年と月を含むログファイル名を生成
                        var logFileName = $"log_{DateTime.Now:yyyyMM}.txt";
                        var logFilePath = Path.Combine(s_logDirectory, logFileName);

                        var logEntry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} {message}";
                        //// ログ内容をファイルの末尾に追記
                        File.AppendAllText(logFilePath, logEntry + Environment.NewLine);

                        if (!string.IsNullOrEmpty(s_networkPath)) {
                            var cloneFilePath = Path.Combine(s_networkPath, "db", "logs", logFileName);
                            if (cloneFilePath != logFilePath) {
                                File.Copy(logFilePath, cloneFilePath, true);
                            }
                        }
                    }
                } catch (Exception ex) {
                    MessageBox.Show($"ログの書き込み中にエラーが発生しました: {ex.Message}", $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        // バックアップ作成
        public static class BackupManager {
            private static readonly string s_backupDirectory = Path.Combine(s_databasePath, "db", "backup"); // バックアップを保存するディレクトリ
            private static readonly string s_originalFilePath = Path.Combine(s_databasePath, "db", "registration.db"); // 元ファイルパス
            private static readonly int s_maxBackupFiles = 10; // 最大バックアップファイル数
            private static readonly object s_lockObject = new();

            /// <summary>
            /// バックアップを作成します。
            /// </summary>
            public static void CreateBackup() {
                try {
                    lock (s_lockObject) {
                        // バックアップ用ディレクトリが存在しない場合は作成
                        if (!Directory.Exists(s_backupDirectory)) {
                            Directory.CreateDirectory(s_backupDirectory);
                        }

                        // 日付と時間をファイル名に付加
                        var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                        var backupFileName = $"registration_{timestamp}.db";
                        var backupFilePath = Path.Combine(s_backupDirectory, backupFileName);

                        // 元ファイルをバックアップにコピー
                        File.Copy(s_originalFilePath, backupFilePath, true);
                        if (!string.IsNullOrEmpty(s_networkPath)) {
                            var cloneFilePath = Path.Combine(s_networkPath, "db", "registration.db");
                            if (Path.Combine(Environment.CurrentDirectory, "db") != s_networkPath) {
                                File.Copy(s_originalFilePath, cloneFilePath, true);
                            }
                        }

                        // バックアップファイルを管理
                        ManageBackupFiles();
                    }
                } catch (Exception ex) {
                    MessageBox.Show($"バックアップの作成中にエラーが発生しました: {ex.Message}", $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            /// <summary>
            /// 古いバックアップファイルを削除します。
            /// </summary>
            private static void ManageBackupFiles() {
                try {
                    var backupFiles = Directory.GetFiles(s_backupDirectory, "registration_*.db")
                                               .OrderBy(File.GetCreationTime) // 作成日時順に並べる
                                               .ToList();

                    while (backupFiles.Count > s_maxBackupFiles) {
                        var oldestFile = backupFiles.First();
                        File.Delete(oldestFile);
                        backupFiles.RemoveAt(0);
                    }
                } catch (Exception ex) {
                    MessageBox.Show($"バックアップの作成中にエラーが発生しました: {ex.Message}", $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        public class ProductInformation {
            public int ProductID { get; set; }
            public string ProductName { get; set; } = string.Empty;
            public string StockName { get; set; } = string.Empty;
            public string ProductType { get; set; } = string.Empty;
            public string ProductModel { get; set; } = string.Empty;
            public string SubstrateName { get; set; } = string.Empty;
            public string SubstrateModel { get; set; } = string.Empty;
            public string UseSubstrate { get; set; } = string.Empty;
            public string Initial { get; set; } = string.Empty;
            public int SerialDigit { get; set; }
            public int RegType { get; set; }
            public int PrintType { get; set; }
            public int RevisionGroup { get; set; }
            public int CheckBin { get; set; }
            public string Proness1 { get; set; } = string.Empty;
            public string Proness2 { get; set; } = string.Empty;
            public string Proness3 { get; set; } = string.Empty;
            public int Proness4 { get; set; }
            public string Proness5 { get; set; } = string.Empty;
            public List<string> Category11 { get; set; } = [];
            public List<string> Category12 { get; set; } = [];
            public List<string> Category13 { get; set; } = [];
            public List<string> Category14 { get; set; } = [];
            public int RadioButtonFlg { get; set; }
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
            public string UsedSubstrate { get; set; } = string.Empty;

            public int Quantity { get; set; }
            public int SerialFirstNumber { get; set; }
            public int SerialLastNumber { get; set; }

            public DataTable ProductDataTable { get; } = new();
            public DataTable SubstrateDataTable { get; } = new();

            public void Reset() {
                ProductID = 0;
                ProductName = string.Empty;
                StockName = string.Empty;
                ProductType = string.Empty;
                ProductModel = string.Empty;
                SubstrateName = string.Empty;
                SubstrateModel = string.Empty;
                UseSubstrate = string.Empty;
                Initial = string.Empty;
                SerialDigit = 0;
                RegType = 0;
                PrintType = 0;
                RevisionGroup = 0;
                CheckBin = 0;
                Proness1 = string.Empty;
                Proness2 = string.Empty;
                Proness3 = string.Empty;
                Proness4 = 0;
                Proness5 = string.Empty;
                Category11 = [];
                Category12 = [];
                Category13 = [];
                Category14 = [];
                OrderNumber = string.Empty;
                ProductNumber = string.Empty;
                RegDate = string.Empty;
                Person = string.Empty;
                Revision = string.Empty;
                Comment = string.Empty;
                SerialFirst = string.Empty;
                SerialLast = string.Empty;
                UsedSubstrate = string.Empty;
                SerialLastNumber = 0;
                Quantity = 0;
                SerialFirstNumber = 0;
            }
        }

        public ProductInformation ProductInfo { get; set; } = new();

        private string _strCategory12 = string.Empty;
        private string _strCategory13 = string.Empty;
        private string _strCategory14 = string.Empty;

        private bool _isAuthorizedUser = false;

        public MainWindow() {
            InitializeComponent();
        }

        public static string GetConnectionInformation() {
            var informationPath = Path.Combine("db", "information.db");

            if (!NetworkInterface.GetIsNetworkAvailable()) {
                throw new Exception("ネットワーク接続がありません。処理を中断します。");
            }

            // ファイルが存在するか確認
            return !File.Exists(informationPath)
                ? throw new FileNotFoundException("データベースファイルが見つかりません。", informationPath)
                : new SQLiteConnectionStringBuilder() { DataSource = informationPath }.ToString();
        }
        public static string GetConnectionRegistration() {
            var registrationPath = Path.Combine("db", "registration.db");

            if (!NetworkInterface.GetIsNetworkAvailable()) {
                throw new Exception("ネットワーク接続がありません。処理を中断します。");
            }

            // ファイルが存在するか確認
            return !File.Exists(registrationPath)
                ? throw new FileNotFoundException("データベースファイルが見つかりません。", registrationPath)
                : new SQLiteConnectionStringBuilder() { DataSource = registrationPath }.ToString();
        }

        // ロードイベント
        private void LoadEvents() {
            try {

                // その日のbackupファイルがない場合バックアップ作成
                var d = DateTime.Now;
                if (string.IsNullOrEmpty(s_clonePath)) { throw new InvalidOperationException("s_clonePath is null"); }
                var backupDir = Path.Combine(s_clonePath, "backup", $"{d.Year}", $"{d.Month:00}");
                var backupFilepath = Path.Combine(s_clonePath, "backup", $"{d.Year}", $"{d.Month:00}", $"_bak_{d.Year}-{d.Month:00}-{d.Day:00}.db");
                var registrationPath = Path.Combine(Environment.CurrentDirectory, "db", "registration.db");

                Directory.CreateDirectory(backupDir);  // ディレクトリが存在しない場合に作成
                File.Copy(registrationPath, backupFilepath, true);

                using SQLiteConnection con = new(GetConnectionInformation());
                using (SQLiteDataAdapter adapter = new("SELECT * FROM Product;", con)) { adapter.Fill(ProductInfo.ProductDataTable); }
                using (SQLiteDataAdapter adapter = new("SELECT * FROM Substrate;", con)) { adapter.Fill(ProductInfo.SubstrateDataTable); }

                // 現在のユーザー名がリストに含まれるかチェック
                _isAuthorizedUser = s_userNames?.Any(name => name.Equals(Environment.UserName, StringComparison.OrdinalIgnoreCase)) ?? false;

                QRCodePanel.Enabled = _isAuthorizedUser;

                RegisterButton.Enabled = false;
                HistoryButton.Enabled = false;

                CategoryListBox1.Items.Clear();
                CategoryListBox2.Items.Clear();
                CategoryListBox3.Items.Clear();
                ProductInfo.ProductDataTable.Clear();

                ActiveControl = QRCodeTextBox;
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void ResetFields() {
            ProductInfo.Reset();
        }
        // 登録ボタン処理
        private void Registration() {
            try {
                ResetFields();

                if (CategoryListBox3.SelectedIndex == -1) { return; }
                switch (ProductInfo.RadioButtonFlg) {
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
            var selectedRows = ProductInfo.ProductDataTable.Select($"class001 = '{CategoryListBox1.SelectedItem}' AND ProductName = '{CategoryListBox2.SelectedItem}' AND SubstrateName = '{CategoryListBox3.SelectedItem}'");

            if (selectedRows.Length > 0) {
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
            var selectedRows = ProductInfo.ProductDataTable.Select($"class001 = '{CategoryListBox1.SelectedItem}' AND ProductName = '{CategoryListBox2.SelectedItem}' AND ProductType = '{CategoryListBox3.SelectedItem}'");

            if (selectedRows.Length > 0) {
                ProductInfo.ProductName = selectedRows[0]["ProductName"].ToString() ?? string.Empty;
                ProductInfo.StockName = selectedRows[0]["StockName"].ToString() ?? string.Empty;
                ProductInfo.ProductType = selectedRows[0]["ProductType"].ToString() ?? string.Empty;
                ProductInfo.RegType = Convert.ToInt32(selectedRows[0]["RegType"] ?? throw new Exception("RegType is null"));
                ProductInfo.PrintType = Convert.ToInt32(selectedRows[0]["PrintType"] ?? throw new Exception("PrintType is null"));
                ProductInfo.SerialDigit = Convert.ToInt32(selectedRows[0]["SerialDigit"] ?? throw new Exception("SerialDigit is null"));
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
            var selectedRows = ProductInfo.ProductDataTable.Select($"class001 = '{CategoryListBox1.SelectedItem}' AND ProductName = '{CategoryListBox2.SelectedItem}' AND ProductType = '{CategoryListBox3.SelectedItem}'");

            if (selectedRows.Length > 0) {
                ProductInfo.ProductName = selectedRows[0]["ProductName"].ToString() ?? string.Empty;
                ProductInfo.ProductType = selectedRows[0]["ProductType"].ToString() ?? string.Empty;
                ProductInfo.RegType = Convert.ToInt32(selectedRows[0]["RegType"] ?? throw new Exception("RegType is null"));
                ProductInfo.PrintType = Convert.ToInt32(selectedRows[0]["PrintType"] ?? throw new Exception("PrintType is null"));
                ProductInfo.SerialDigit = Convert.ToInt32(selectedRows[0]["SerialDigit"] ?? throw new Exception("SerialDigit is null"));
                ProductInfo.ProductModel = selectedRows[0]["ProductModel"].ToString() ?? string.Empty;
                ProductInfo.CheckBin = Convert.ToInt32(selectedRows[0]["Checkbox"].ToString() ?? throw new Exception("Checkbox is null"), 2);
                ProductInfo.Initial = selectedRows[0]["Initial"].ToString() ?? string.Empty;
                ProductInfo.RevisionGroup = Convert.ToInt32(selectedRows[0]["RevisionGroup"] ?? throw new Exception("RevisionGroup is null"));
                using RePrintWindow window = new(ProductInfo);
                window.ShowDialog(this);
            }
        }
        private void HandleSubstrateChange1() {
            var selectedRows = ProductInfo.ProductDataTable.Select($"class001 = '{CategoryListBox1.SelectedItem}' AND ProductName = '{CategoryListBox2.SelectedItem}' AND ProductType = '{CategoryListBox3.SelectedItem}'");

            if (selectedRows.Length > 0) {
                ProductInfo.PrintType = Convert.ToInt32(selectedRows[0]["PrintType"] ?? throw new Exception("PrintType is null"));
                ProductInfo.RegType = Convert.ToInt32(selectedRows[0]["RegType"] ?? throw new Exception("RegType is null"));
                ProductInfo.ProductName = selectedRows[0]["ProductName"].ToString() ?? string.Empty;
                ProductInfo.StockName = selectedRows[0]["StockName"].ToString() ?? string.Empty;
                ProductInfo.ProductType = selectedRows[0]["ProductType"].ToString() ?? string.Empty;
                ProductInfo.ProductModel = selectedRows[0]["ProductModel"].ToString() ?? string.Empty;
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

                switch (ProductInfo.RadioButtonFlg) {
                    case 1:
                        selectedRow = ProductInfo.ProductDataTable.Select($"class001 = '{CategoryListBox1.SelectedItem}' AND ProductName = '{CategoryListBox2.SelectedItem}' AND SubstrateName = '{CategoryListBox3.SelectedItem}'");
                        break;
                    case 2:
                    case 3:
                    case 4:
                        selectedRow = ProductInfo.ProductDataTable.Select($"class001 = '{CategoryListBox1.SelectedItem}' AND ProductName = '{CategoryListBox2.SelectedItem}' AND ProductType = '{CategoryListBox3.SelectedItem}'");
                        break;
                }

                if (selectedRow != null && selectedRow.Length > 0) {

                    switch (ProductInfo.RadioButtonFlg) {
                        case 1:
                            ProductInfo.ProductName = selectedRow[0]["ProductName"].ToString() ?? string.Empty;
                            ProductInfo.SubstrateName = selectedRow[0]["SubstrateName"].ToString() ?? string.Empty;
                            ProductInfo.SubstrateModel = selectedRow[0]["SubstrateModel"].ToString() ?? string.Empty;
                            ProductInfo.RegType = Convert.ToInt32(selectedRow[0]["RegType"] ?? throw new Exception("RegType is null"));
                            break;
                        case 2:
                        case 3:
                            ProductInfo.ProductName = selectedRow[0]["ProductName"].ToString() ?? string.Empty;
                            ProductInfo.ProductType = selectedRow[0]["ProductType"].ToString() ?? string.Empty;
                            ProductInfo.ProductModel = selectedRow[0]["ProductModel"].ToString() ?? string.Empty;
                            ProductInfo.RegType = Convert.ToInt32(selectedRow[0]["RegType"] ?? throw new Exception("RegType is null"));
                            break;
                    }

                    using HistoryWindow window = new(ProductInfo);
                    window.ShowDialog(this);
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // 処理カテゴリセレクト
        private void CategorySelect(object sender) {
            try {
                RegisterButton.Enabled = false;
                HistoryButton.Enabled = false;
                CategoryListBox1.Items.Clear();
                CategoryListBox2.Items.Clear();
                CategoryListBox3.Items.Clear();
                ProductInfo.ProductDataTable.Clear();

                var selectedRadioButton = (RadioButton)sender;
                var strSqlQuery = string.Empty;

                switch (selectedRadioButton.Tag) {
                    case "1":
                        ProductInfo.RadioButtonFlg = 1;
                        strSqlQuery = """SELECT * FROM Substrate WHERE Visible = 1;""";
                        break;
                    case "2":
                        ProductInfo.RadioButtonFlg = 2;
                        strSqlQuery = """SELECT * FROM Product WHERE Visible = 1;""";
                        break;
                    case "3":
                        ProductInfo.RadioButtonFlg = 3;
                        strSqlQuery = """SELECT * FROM Product WHERE Visible = 1;""";
                        break;
                    case "4":
                        ProductInfo.RadioButtonFlg = 4;
                        strSqlQuery = """SELECT * FROM Product WHERE Visible = 1 AND (PrintType = 5 OR PrintType = 6);""";
                        break;
                    default:
                        break;
                }

                using (SQLiteConnection con = new(GetConnectionInformation()))
                using (SQLiteDataAdapter adapter = new(strSqlQuery, con)) {
                    adapter.Fill(ProductInfo.ProductDataTable);
                }

                var class001Set = new SortedSet<string>(ProductInfo.ProductDataTable.AsEnumerable()
                    .Select(row => row.Field<string?>("class001"))
                    .Where(classVal => classVal != null)
                    .Cast<string>()
                );

                CategoryListBox1.Items.AddRange([.. class001Set]);
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // 製品カテゴリセレクト
        private void CategoryListBox1Select() {
            try {
                RegisterButton.Enabled = false;
                HistoryButton.Enabled = false;
                CategoryListBox2.Items.Clear();
                CategoryListBox3.Items.Clear();

                HashSet<string> productNames = [];

                var selectedRows = ProductInfo.ProductDataTable.Select($"class001 = '{CategoryListBox1.SelectedItem}'", "ProductName ASC");

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
                HistoryButton.Enabled = false;
                CategoryListBox3.Items.Clear();

                DataRow[] selectedRows;

                switch (ProductInfo.RadioButtonFlg) {
                    case 1:
                        selectedRows = ProductInfo.ProductDataTable.Select($"class001 = '{CategoryListBox1.SelectedItem}' AND ProductName = '{CategoryListBox2.SelectedItem}'", "SubstrateName ASC");
                        HashSet<string> substrateNames = new(selectedRows.AsEnumerable()
                                                                                .Select(x => x.Field<string>("SubstrateName"))
                                                                                .Where(x => x != null)
                                                                                .Select(x => x!));

                        CategoryListBox3.Items.AddRange([.. substrateNames]);
                        break;

                    case 2:
                    case 3:
                    case 4:
                        selectedRows = ProductInfo.ProductDataTable.Select($"class001 = '{CategoryListBox1.SelectedItem}' AND ProductName = '{CategoryListBox2.SelectedItem}'", "ProductType ASC");
                        HashSet<string> productTypes = new(selectedRows.AsEnumerable()
                                                                                .Select(x => x.Field<string>("ProductType"))
                                                                                .Where(x => x != null)
                                                                                .Select(x => x!));

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
                RegisterButton.Enabled = _isAuthorizedUser;
                HistoryButton.Enabled = true;

                switch (ProductInfo.RadioButtonFlg) {
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
                if (ProductInfo.Category11.Count >= 2) {
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
            ProductInfo.Category11.Clear();
            ProductInfo.Category12.Clear();
            ProductInfo.Category13.Clear();
            ProductInfo.Category14.Clear();
            Enabled = false;
        }
        private void ParseQRCodeInput() {
            try {
                string[] separator = ["//"];
                var arr = QRCodeTextBox.Text.Split(separator, StringSplitOptions.None);
                if (arr.Length != 4) { throw new Exception("QRコードが正しくありません。"); }
                if (arr != null) {
                    ProductInfo.Proness1 = arr[0];
                    ProductInfo.Proness2 = arr[1];
                    ProductInfo.Proness4 = Convert.ToInt32(arr[2] ?? throw new Exception("arr[2] is null"));
                    ProductInfo.Proness5 = arr[3];
                }
            } catch (Exception ex) {
                throw new Exception($"{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー{Environment.NewLine}{ex.Message}");
            }
        }
        private void BarcodeInput() {
            using (OdbcConnection con = new("DSN=DrSum_PRONES_YD; UID=YD00; PWD=YD00")) {
                con.Open();
                using OdbcCommand cmd = new($"""SELECT * FROM V_宮崎手配情報 WHERE 手配管理番号 = "{QRCodeTextBox.Text}" """, con);
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
            ProductInfo.Proness2 = ProductInfo.Proness2
                                                .Replace("-SMT", "")
                                                .Replace("-H", "")
                                                .Replace("-GH", "")
                                                .Replace("-ACGH", "-AC")
                                                .Replace("-DCGH", "-DC");
        }
        private void FetchDataFromSQLite() {
            using SQLiteConnection con = new(GetConnectionInformation());
            con.Open();
            using var cmd = con.CreateCommand();
            cmd.CommandText = $"""SELECT * FROM V_ItemList WHERE SubItemNumber = @StrProness2 OR ProItemNumber = @StrProness2""";
            cmd.Parameters.Add("@StrProness2", DbType.String).Value = ProductInfo.Proness2;
            using var dr = cmd.ExecuteReader();
            if (!dr.HasRows) { throw new Exception($"品目番号が見つかりません。\n品目番号:[{ProductInfo.Proness2}]"); }
            while (dr.Read()) {
                var colSubItemNumber = dr["SubItemNumber"].ToString() ?? string.Empty;
                var colProItemNumber = dr["ProItemNumber"].ToString() ?? string.Empty;

                if (!string.IsNullOrEmpty(colSubItemNumber)) {
                    var substrateName = dr["SubstrateName"]?.ToString() ?? string.Empty;
                    var productName = dr["ProductName"]?.ToString() ?? string.Empty;
                    AddToLists(colSubItemNumber, substrateName, productName, "1");
                }

                if (!string.IsNullOrEmpty(colProItemNumber)) {
                    var productType = dr["ProductType"]?.ToString() ?? string.Empty;
                    var productName = dr["ProductName:1"]?.ToString() ?? string.Empty;
                    AddToLists(colProItemNumber, productType, productName, "2");
                }
            }
        }
        private void AddToLists(string itemNumber, string category12, string category13, string category14) {
            ProductInfo.Category11.Add(itemNumber);
            ProductInfo.Category12.Add(category12);
            ProductInfo.Category13.Add(category13);
            ProductInfo.Category14.Add(category14);
        }
        private int ShowDialogWindowForMultipleItems() {
            using SeveralDialogWindow window = new(ProductInfo);
            window.ShowDialog(this);
            return window.SelectedIndex;
        }
        private void HandleSelectedItem(int listIndex) {
            _strCategory12 = ProductInfo.Category12[listIndex];
            _strCategory13 = ProductInfo.Category13[listIndex];
            _strCategory14 = ProductInfo.Category14[listIndex];

            switch (_strCategory14) {
                case "1":
                    HandleSubstrateSelection();
                    break;
                case "2":
                    HandleProductSelection();
                    break;
                default:
                    throw new Exception($"一致する情報がありません。{Environment.NewLine}品目番号:{ProductInfo.Proness2}{Environment.NewLine}");
            }
        }
        private void HandleSubstrateSelection() {
            using (SQLiteConnection con = new(GetConnectionInformation())) {
                using SQLiteDataAdapter adapter = new("""SELECT * FROM Substrate;""", con);
                adapter.Fill(ProductInfo.ProductDataTable);
            }

            var substrateRet = ProductInfo.ProductDataTable.Select($"ProductName = '{_strCategory13}' AND SubstrateName = '{_strCategory12}'");
            OpenSubstrateRegistrationWindow(substrateRet);
        }
        private void OpenSubstrateRegistrationWindow(DataRow[] substrateRet) {
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
        private void HandleProductSelection() {
            using (SQLiteConnection con = new(GetConnectionInformation())) {
                using SQLiteDataAdapter adapter = new("""SELECT * FROM Product;""", con);
                adapter.Fill(ProductInfo.ProductDataTable);
            }

            var productRet = ProductInfo.ProductDataTable.Select($"ProductName = '{_strCategory13}' AND ProductType = '{_strCategory12}'");
            OpenProductRegistrationWindow(productRet);
        }
        private void OpenProductRegistrationWindow(DataRow[] productRet) {
            ProductInfo.ProductName = productRet[0]["ProductName"].ToString() ?? string.Empty;
            ProductInfo.StockName = productRet[0]["StockName"].ToString() ?? string.Empty;
            ProductInfo.ProductType = productRet[0]["ProductType"].ToString() ?? string.Empty;
            ProductInfo.ProductModel = productRet[0]["ProductModel"].ToString() ?? string.Empty;
            ProductInfo.UseSubstrate = productRet[0]["UseSubstrate"].ToString() ?? string.Empty;
            ProductInfo.Initial = productRet[0]["Initial"].ToString() ?? string.Empty;
            ProductInfo.RegType = Convert.ToInt32(productRet[0]["RegType"] ?? throw new Exception("RegType is null"));
            ProductInfo.PrintType = Convert.ToInt32(productRet[0]["PrintType"] ?? throw new Exception("PrintType is null"));
            ProductInfo.CheckBin = Convert.ToInt32(productRet[0]["Checkbox"].ToString() ?? throw new Exception("Checkbox is null"), 2);
            ProductInfo.SerialDigit = Convert.ToInt32(productRet[0]["SerialDigit"] ?? throw new Exception("SerialDigit is null"));
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
            var excelFullPath = System.IO.Path.Combine(xlApp.Path, "EXCEL.EXE");   //フルパス作成
            xlApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
            // ファイルを実行
            System.Diagnostics.Process.Start(excelFullPath, filePath);
        }

        private void MainWindow_Load(object sender, EventArgs e) { LoadEvents(); }
        private void ReloadToolStripMenuItem_Click(object sender, EventArgs e) { LoadEvents(); }
        private void ReportConfigToolStripMenuItem_Click(object sender, EventArgs e) {
            var reportConfigPath = Path.Combine(Environment.CurrentDirectory, "config", "Excel", "ConfigReport.xlsx");
            OpenExcel(reportConfigPath);
        }
        private void ListConfigToolStripMenuItem_Click(object sender, EventArgs e) {
            var listConfigPath = Path.Combine(Environment.CurrentDirectory, "config", "Excel", "ConfigList.xlsx");
            OpenExcel(listConfigPath);
        }
        private void CheckSheetConfigToolStripMenuItem_Click(object sender, EventArgs e) {
            var checkSheetConfigPath = Path.Combine(Environment.CurrentDirectory, "config", "Excel", "ConfigCheckSheet.xlsx");
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
    }
}
