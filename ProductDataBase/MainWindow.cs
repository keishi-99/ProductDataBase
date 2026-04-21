using ProductDatabase.Common;
using ProductDatabase.Data;
using ProductDatabase.History;
using ProductDatabase.LogViewer;
using ProductDatabase.Models;
using ProductDatabase.Services;
using System.Data;

namespace ProductDatabase {
    public partial class MainWindow : Form {

        public int RadioButtonNumber { get; set; }
        private float _fontSize = SystemFonts.DefaultFont.Size;
        private IEnumerable<DataRow> _currentTargetRows = [];

        private BarcodeService? _barcodeService;

        readonly string _jsonFilePath = Path.Combine(Environment.CurrentDirectory, "Config", "General", "appsettings.json");

        private readonly ProductRepository _productRepository;

        private readonly ProductMaster _productMaster;
        private readonly ProductRegisterWork _productRegisterWork;

        private readonly SubstrateMaster _substrateMaster;
        private readonly SubstrateRegisterWork _substrateRegisterWork;

        private readonly QrSettings _qrSettings;
        private readonly AppSettings _appSettings;

        public MainWindow() {
            InitializeComponent();

            _productRepository = new ProductRepository();
            _productMaster = new ProductMaster();
            _productRegisterWork = new ProductRegisterWork();
            _substrateMaster = new SubstrateMaster();
            _substrateRegisterWork = new SubstrateRegisterWork();
            _qrSettings = new QrSettings();
            _appSettings = new AppSettings();
        }

        private static FileStream? _lockStream;

        // フォームロード時にファイルロック・設定読み込み・DBデータ取得・日次バックアップ作成を行いUIを初期化する
        private void LoadEvents() {
            try {
                // ファイルロック
                LockSelf();

                RegisterButton.Enabled = false;
                HistoryButton.Enabled = false;

                CategoryRadioButton1.Checked = false;
                CategoryRadioButton2.Checked = false;
                CategoryRadioButton3.Checked = false;
                CategoryRadioButton4.Checked = false;

                CategoryListBox1.Items.Clear();
                CategoryListBox2.Items.Clear();
                CategoryListBox3.Items.Clear();

                // 設定ファイル読み込み
                var generalSettings = SettingsLoader.Load(_jsonFilePath);

                // バックアップパス取得
                FileUtils.BackupPath = generalSettings.BackupFolderPath;

                // バックアップ作成
                BackupManager.CreateDailyBackup();

                // 担当者取得
                _appSettings.PersonList = [.. generalSettings.Persons];

                // 認証ユーザー名を取得
                _appSettings.IsAuthorizedUser = generalSettings.AuthorizedUsers
                    .Contains(Environment.UserName, StringComparer.OrdinalIgnoreCase);

                // 管理者ユーザー名を取得
                _appSettings.IsAdministrator = generalSettings.Administrators
                    .Contains(Environment.UserName, StringComparer.OrdinalIgnoreCase);

                _barcodeService = new BarcodeService(generalSettings.DSN, generalSettings.UID, generalSettings.PWD);

                _productRepository.LoadAll();

                this.Activate();
                QRCodePanel.Enabled = _appSettings.IsAuthorizedUser;

                // 管理者のみマスター管理メニューを有効にする
                //MasterManagementToolStripMenuItem.Enabled = _appSettings.IsAdministrator;
                MasterManagementToolStripMenuItem.Enabled = true;
                QRCodeTextBox.Select();
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // 各マスター・作業データをリセットしDBを再読み込みする
        private void ResetFields() {
            _productMaster.Reset();
            _productRegisterWork.Reset();
            _productRepository.Clear();
            _substrateMaster.Reset();
            _substrateRegisterWork.Reset();

            _productRepository.LoadAll();

            _appSettings.FontSize = _fontSize;
        }
        // 実行中のEXEファイルをロックして二重起動を防止する
        private static void LockSelf() {
            try {
                string exePath = Application.ExecutablePath;
                _lockStream = new FileStream(exePath, FileMode.Open, FileAccess.Read, FileShare.Read);
            } catch {
                // 失敗しても無視
            }
        }
        // ラジオボタンのモードに応じて選択品目の基板登録・製品登録・再印刷・基板変更ウィンドウを開く
        private void Registration() {
            ResetFields();

            if (CategoryListBox3.SelectedItem is not ListItem<long> item) { return; }

            switch (RadioButtonNumber) {
                case 1:
                    HandleSubstrateRegistration(item.Id);
                    break;

                case 2:
                    HandleProductRegistration(item.Id, ProductOperationMode.Register);
                    break;

                case 3:
                    HandleProductRegistration(item.Id, ProductOperationMode.RePrint);
                    break;

                case 4:
                    HandleProductRegistration(item.Id, ProductOperationMode.SubstrateChange);
                    break;
            }

            QRCodeTextBox.Text = string.Empty;
        }
        // 指定基板IDのマスターを読み込み基板登録ウィンドウを開く
        private void HandleSubstrateRegistration(long substrateId) {

            var row = _productRepository.GetSubstrateById(substrateId);

            _substrateMaster.LoadFrom(row);

            using SubstrateRegistrationWindow window = new(_substrateMaster, _substrateRegisterWork, _appSettings);
            window.ShowDialog(this);
        }
        // 指定製品IDのマスターを読み込みモードに応じた製品操作ウィンドウを開く
        private void HandleProductRegistration(long productId, ProductOperationMode mode) {

            var row = _productRepository.GetProductById(productId);

            _productMaster.LoadFrom(row);
            _productMaster.UseSubstrates = ProductRepository.GetUseSubstrates(_productMaster.ProductID);

            switch (mode) {
                case ProductOperationMode.Register:
                    using (var window = new ProductRegistration1Window(_productMaster, _productRegisterWork, _appSettings)) {
                        window.ShowDialog(this);
                    }
                    break;

                case ProductOperationMode.RePrint:
                    using (var window = new RePrintWindow(_productMaster, _productRegisterWork, _appSettings)) {
                        window.ShowDialog(this);
                    }
                    break;

                case ProductOperationMode.SubstrateChange:
                    using (var window = new SubstrateChange1(_productMaster, _productRegisterWork, _appSettings)) {
                        window.ShowDialog(this);
                    }
                    break;
            }
        }
        // 品目選択有無に応じてマスターをセットし履歴ウィンドウをダイアログで開く
        private void History() {

            ResetFields();

            if (CategoryListBox3.SelectedItem is not ListItem<long> item) {
                LoadHistoryWithoutSelection();
            }
            else {
                LoadHistoryWithSelection(item.Id);
            }

            using var window = new HistoryWindow(
                _productMaster,
                _productRegisterWork,
                _substrateMaster,
                _substrateRegisterWork,
                RadioButtonNumber,
                _appSettings);

            window.ShowDialog(this);
        }
        // 品目未選択時にカテゴリ名・製品名のみマスターにセットする
        private void LoadHistoryWithoutSelection() {
            switch (RadioButtonNumber) {
                case 1:
                    _substrateMaster.CategoryName = CategoryListBox1.SelectedItem?.ToString() ?? string.Empty;
                    _substrateMaster.ProductName = CategoryListBox2.SelectedItem?.ToString() ?? string.Empty;
                    break;
                case 2 or 3 or 4:
                    _productMaster.CategoryName = CategoryListBox1.SelectedItem?.ToString() ?? string.Empty;
                    _productMaster.ProductName = CategoryListBox2.SelectedItem?.ToString() ?? string.Empty;
                    break;
            }
        }
        // 品目選択時に該当マスターデータをDBから読み込む
        private void LoadHistoryWithSelection(long itemId) {
            switch (RadioButtonNumber) {
                case 1:
                    _substrateMaster.LoadFrom(_productRepository.GetSubstrateById(itemId));
                    break;
                case 2 or 3 or 4:
                    _productMaster.LoadFrom(_productRepository.GetProductById(itemId));
                    break;
                default:
                    throw new InvalidOperationException("不正なモードです");
            }
        }

        private record CategoryConfig(string OrderKey, string IdKey, string NameKey);
        private readonly Dictionary<int, CategoryConfig> _categoryConfigs = new() {
            { 1, new CategoryConfig("SubstrateName", "SubstrateID", "SubstrateName") }, // 基板登録
            { 2, new CategoryConfig("ProductType",   "ProductID",   "ProductType")   }, // 製品登録
            { 3, new CategoryConfig("ProductType",   "ProductID",   "ProductType")   }, // 再印刷
            { 4, new CategoryConfig("ProductType",   "ProductID",   "ProductType")   }  // 基板変更
        };
        // ラジオボタン選択時にモードに応じたマスターデータをフィルタしてCategoryListBox1にカテゴリ一覧を表示する
        private void CategorySelect(object sender) {

            RegisterButton.Enabled = false;
            HistoryButton.Enabled = false;
            CategoryListBox1.Items.Clear();
            CategoryListBox2.Items.Clear();
            CategoryListBox3.Items.Clear();

            // RadioButton の Tag 取得
            if (sender is not RadioButton rb ||
                !int.TryParse(rb.Tag?.ToString(), out int number)) {
                return;
            }

            RadioButtonNumber = number;

            // 未定義モード防止
            if (!_categoryConfigs.ContainsKey(RadioButtonNumber)) {
                _currentTargetRows = [];
                return;
            }

            // データソース切替
            bool isSubstrateMode = RadioButtonNumber == 1;

            var sourceTable = isSubstrateMode
                ? _productRepository.SubstrateDataTable
                : _productRepository.ProductDataTable;

            // フィルタ済み行を保持
            _currentTargetRows = sourceTable
                .AsEnumerable()
                .Where(r => r.Field<long?>("Visible") == 1)
                .Where(r => RadioButtonNumber switch {
                    1 => true,
                    2 => true,
                    3 => r.Field<long?>("SerialPrintType") is long spt && spt != 0,
                    4 => r.Field<long?>("SheetPrintType") is long shp && (shp == 2 || shp == 3),
                    _ => false
                });

            // CategoryName 列の重複除外＋名前順（"Other" は末尾）
            var categoryNames = _currentTargetRows
                .Where(r => r["CategoryName"] != DBNull.Value)
                .Select(r => r["CategoryName"]!.ToString()!)
                .Where(name => !string.IsNullOrWhiteSpace(name))
                .Distinct()
                .OrderBy(name => name == "Other" ? 1 : 0)
                .ThenBy(name => name)
                .ToList();

            CategoryListBox1.Items.AddRange([.. categoryNames]);
            HistoryButton.Enabled = RadioButtonNumber != 4;
        }
        // カテゴリ選択時に一致する製品名またはSubstrateNameの一覧をListBox2に表示する
        private void CategoryListBox1Select() {
            RegisterButton.Enabled = false;
            HistoryButton.Enabled = RadioButtonNumber != 4;
            CategoryListBox2.Items.Clear();
            CategoryListBox3.Items.Clear();

            if (CategoryListBox1.SelectedItem is null) {
                return;
            }

            string selectedCategory = CategoryListBox1.SelectedItem.ToString()!;

            // _currentTargetRows から ProductName を取得
            var productNames = _currentTargetRows
                .Where(r =>
                    r["CategoryName"]?.ToString() == selectedCategory &&
                    r["ProductName"] != DBNull.Value)
                .Select(r => r["ProductName"]!.ToString()!)
                .Where(name => !string.IsNullOrWhiteSpace(name))
                .Distinct()
                .OrderBy(name => name)
                .ToList();

            CategoryListBox2.Items.AddRange([.. productNames]);
        }
        // 製品名セレクト：カテゴリ・製品名に一致する品目一覧をListBox3に表示する
        private void CategoryListBox2Select() {
            RegisterButton.Enabled = false;
            HistoryButton.Enabled = RadioButtonNumber != 4;
            CategoryListBox3.Items.Clear();

            if (CategoryListBox1.SelectedItem is null ||
                CategoryListBox2.SelectedItem is null) {
                return;
            }

            if (!_categoryConfigs.TryGetValue(RadioButtonNumber, out var config)) {
                return;
            }

            var selectedRows = _currentTargetRows
                .Where(r =>
                    r["CategoryName"]?.ToString() == CategoryListBox1.SelectedItem?.ToString() &&
                    r["ProductName"]?.ToString() == CategoryListBox2.SelectedItem?.ToString())
                .OrderBy(r => r[config.OrderKey]?.ToString())
                .ToArray();

            var items = selectedRows
                .Select(r => new ListItem<long> {
                    Id = r.Field<long>(config.IdKey),
                    Name = r[config.NameKey]?.ToString() ?? string.Empty
                })
                .GroupBy(x => x.Id)
                .Select(g => g.First())
                .OrderBy(x => x.Name)
                .ToList();

            CategoryListBox3.Items.AddRange([.. items.Cast<object>()]);
        }
        // 品目セレクト：選択確定時に登録ボタンを有効化する
        private void CategoryListBox3Select() {
            RegisterButton.Enabled = true;
        }

        // QR/バーコード入力を解析してSQLiteを検索し品目候補に応じた登録ウィンドウを開く
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
                if (_qrSettings.CategoryItemNumber.Count >= 2) {
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
        // コードスキャン前にUIと各設定データをリセットしフォームを無効化する
        private void ResetFieldsForCodeScan() {
            CategoryRadioButton1.Checked = CategoryRadioButton2.Checked = CategoryRadioButton3.Checked = CategoryRadioButton4.Checked = false;
            CategoryListBox1.Items.Clear();
            CategoryListBox2.Items.Clear();
            CategoryListBox3.Items.Clear();
            _qrSettings.CategoryItemNumber.Clear();
            _qrSettings.CategoryProductType.Clear();
            _qrSettings.CategoryProductName.Clear();
            _qrSettings.CategorySubstrateName.Clear();
            _qrSettings.CategoryType.Clear();
            ResetFields();
            Enabled = false;
        }
        // QRコードテキストを解析して製番・品目番号・数量・注番を取得する
        private void ParseQRCodeInput() {
            try {
                var parsed = QrCodeParser.Parse(QRCodeTextBox.Text);
                _productMaster.ProductModel = parsed.ProductModel;
                if (!string.IsNullOrEmpty(parsed.ProductNumber)) {
                    _productRegisterWork.ProductNumber = parsed.ProductNumber;
                    _productRegisterWork.Quantity = parsed.Quantity;
                    _productRegisterWork.OrderNumber = parsed.OrderNumber;
                }
            } catch (Exception ex) {
                throw new Exception($"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー{Environment.NewLine}{ex.Message}", ex);
            }
        }
        // バーコードの手配管理番号からODBCで手配情報を取得して各フィールドにセットする
        private void BarcodeInput() {
            var result = (_barcodeService ?? throw new InvalidOperationException("BarcodeService が初期化されていません。"))
                .Query(QRCodeTextBox.Text);
            _productRegisterWork.ProductNumber = result.ProductNumber;
            _productMaster.ProductModel = result.ProductModel;
            _productMaster.ProductName = result.ProductName;
            _productRegisterWork.Quantity = result.Quantity;
            _productRegisterWork.OrderNumber = result.OrderNumber;
        }
        // 品目番号から不要なサフィックスを除去して正規化する
        private void ProcessCategoryItemData() {
            _productMaster.ProductModel = QrCodeParser.NormalizeProductModel(_productMaster.ProductModel);
        }
        // 品目番号でSQLiteを検索し一致する基板・製品情報をリストに追加する
        private void FetchDataFromSQLite() {
            var items = ProductRepository.SearchByModel(_productMaster.ProductModel);
            foreach (var item in items) {
                _qrSettings.CategoryItemNumber.Add(item.ItemNumber);
                _qrSettings.CategoryProductName.Add(item.ProductName);
                _qrSettings.CategoryProductType.Add(item.ProductType);
                _qrSettings.CategorySubstrateName.Add(item.SubstrateName);
                _qrSettings.CategoryType.Add(item.Type);
            }
        }
        // 複数候補がある場合に選択ダイアログを表示し選択インデックスを返す
        private int ShowDialogWindowForMultipleItems() {
            using SeveralDialogWindow window = new(_qrSettings, _appSettings);
            window.ShowDialog(this);
            return window.SelectedIndex;
        }
        // 選択された品目のタイプに応じて基板または製品の登録処理を実行する
        private void HandleSelectedItem(int listIndex) {
            var productName = _qrSettings.CategoryProductName[listIndex];
            var productType = _qrSettings.CategoryProductType[listIndex];
            var substrateName = _qrSettings.CategorySubstrateName[listIndex];
            var type = _qrSettings.CategoryType[listIndex];

            switch (type) {
                case "1":
                    HandleSubstrateSelection(productName, substrateName);
                    break;
                case "2":
                    HandleProductSelection(productName, productType);
                    break;
                default:
                    throw new Exception($"一致する情報がありません。{Environment.NewLine}品目番号:{_productMaster.ProductModel}{Environment.NewLine}");
            }
        }
        // 製品名と基板名で基板マスターを検索し基板登録ウィンドウを開く
        private void HandleSubstrateSelection(string productName, string substrateName) {
            var substrateRet = _productRepository.SubstrateDataTable
                .AsEnumerable()
                .Where(r => r["ProductName"]?.ToString() == productName &&
                            r["SubstrateName"]?.ToString() == substrateName)
                .ToArray();
            OpenSubstrateRegistrationWindow(substrateRet);
        }
        // 基板マスターと作業データをセットして基板登録ウィンドウを表示する
        private void OpenSubstrateRegistrationWindow(DataRow[] substrateRet) {
            _substrateMaster.LoadFrom(substrateRet[0]);
            _substrateRegisterWork.ProductNumber = _productRegisterWork.ProductNumber;
            _substrateRegisterWork.OrderNumber = _productRegisterWork.OrderNumber;
            _substrateRegisterWork.AddQuantity = _productRegisterWork.Quantity;
            using SubstrateRegistrationWindow window = new(_substrateMaster, _substrateRegisterWork, _appSettings);
            window.ShowDialog(this);
        }
        // 製品名と型式で製品マスターを検索し製品登録ウィンドウを開く
        private void HandleProductSelection(string productName, string productType) {
            var productRet = _productRepository.ProductDataTable
                .AsEnumerable()
                .Where(r => r["ProductName"]?.ToString() == productName &&
                            r["ProductType"]?.ToString() == productType)
                .ToArray();
            OpenProductRegistrationWindow(productRet);
        }
        // 製品マスターと使用基板情報をセットして製品登録ウィンドウを表示する
        private void OpenProductRegistrationWindow(DataRow[] productRet) {
            _productMaster.LoadFrom(productRet[0]);
            _productMaster.UseSubstrates = ProductRepository.GetUseSubstrates(_productMaster.ProductID);
            using ProductRegistration1Window window = new(_productMaster, _productRegisterWork, _appSettings);
            window.ShowDialog(this);
        }
        // スキャン後にフォームを再有効化してQRコード入力欄にフォーカスを戻す
        private void CleanupAfterScan() {
            Enabled = true;
            QRCodeTextBox.Focus();
        }

        // ラジオボタン選択に応じてアプリ全体のフォントサイズを変更する
        private void FontChange(object sender) {
            var radioButton = (RadioButton)sender;

            _fontSize = radioButton.Name switch {
                "FontSize9RadioButton" => 9,
                "FontSize12RadioButton" => 12,
                "FontSize14RadioButton" => 14,
                _ => _fontSize
            };

            _appSettings.FontSize = _fontSize;
            Font = new System.Drawing.Font(_appSettings.FontName, _appSettings.FontSize);
        }
        private void MainWindow_Load(object sender, EventArgs e) { LoadEvents(); }
        private void ReloadToolStripMenuItem_Click(object sender, EventArgs e) { LoadEvents(); }
        private void ConfigReportToolStripMenuItem_Click(object sender, EventArgs e) {
            var reportConfigPath = Path.Combine(Environment.CurrentDirectory, "config", "General", "Excel", "ConfigReport.xlsm");
            ExcelLauncher.Open(reportConfigPath);
        }
        private void ConfigListToolStripMenuItem_Click(object sender, EventArgs e) {
            var listConfigPath = Path.Combine(Environment.CurrentDirectory, "config", "General", "Excel", "ConfigList.xlsm");
            ExcelLauncher.Open(listConfigPath);
        }
        private void ConfigCheckSheetToolStripMenuItem_Click(object sender, EventArgs e) {
            var checkSheetConfigPath = Path.Combine(Environment.CurrentDirectory, "config", "General", "Excel", "ConfigCheckSheet.xlsm");
            ExcelLauncher.Open(checkSheetConfigPath);
        }
        private void ConfigSubstrateInformationToolStripMenuItem_Click(object sender, EventArgs e) {
            var checkSheetConfigPath = Path.Combine(Environment.CurrentDirectory, "config", "General", "Excel", "ConfigSubstrateInformation.xlsm");
            ExcelLauncher.Open(checkSheetConfigPath);
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

        private void LogViewerToolStripMenuItem_Click(object sender, EventArgs e) {
            using var window = new LogViewerWindow(_appSettings);
            window.ShowDialog(this);
        }

        // マスター管理画面を管理者専用ダイアログで開く
        private void MasterManagementToolStripMenuItem_Click(object sender, EventArgs e) {
            try {
                using var window = new MasterManagement.MasterManagementWindow(_productRepository, _appSettings);
                window.ShowDialog(this);
                // マスターデータが変更されている可能性があるためキャッシュを更新する
                _productRepository.LoadAll();
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

    }
}
