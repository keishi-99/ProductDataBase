using ProductDatabase.Common;
using ProductDatabase.Data;
using ProductDatabase.History;
using ProductDatabase.LogViewer;
using ProductDatabase.Models;
using ProductDatabase.Services;
using System.ComponentModel;
using System.Data;

namespace ProductDatabase {
    public partial class MainWindow : Form {

        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public RadioButtonMode RadioButtonNumber { get; set; }
        private IEnumerable<DataRow> _currentTargetRows = [];

        private BarcodeService? _barcodeService;

        readonly string _jsonFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Config", "General", "appsettings.json");

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

                // 初期化の実行
                var initializer = new ApplicationInitializer(_jsonFilePath);

                // 設定ファイル読み込み
                GeneralSettings generalSettings;
                try {
                    generalSettings = initializer.LoadSettings();
                } catch (Exception ex) {
                    MessageBox.Show($"設定ファイルの読み込みに失敗しました:\n{ex.Message}", "致命的エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Close();
                    return;
                }

                // バックアップ作成
                initializer.CreateDailyBackup(generalSettings.BackupFolderPath);

                // アプリ設定
                try {
                    var appSettings = initializer.ConfigureAppSettings(generalSettings);
                    _appSettings.PersonList = appSettings.PersonList;
                    _appSettings.IsAuthorizedUser = appSettings.IsAuthorizedUser;
                    _appSettings.IsAdministrator = appSettings.IsAdministrator;
                } catch (Exception ex) {
                    MessageBox.Show($"アプリケーション設定に失敗しました:\n{ex.Message}", "致命的エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Close();
                    return;
                }

                // バーコード・QRサービス
                try {
                    _barcodeService = initializer.CreateBarcodeService(generalSettings);
                } catch (Exception ex) {
                    MessageBox.Show($"バーコード・QRサービスの初期化に失敗しました:\n{ex.Message}", "致命的エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Close();
                    return;
                }

                // DB読み込み
                try {
                    initializer.LoadDatabase(_productRepository);
                } catch (Exception ex) {
                    MessageBox.Show($"データベースの読み込みに失敗しました:\n{ex.Message}", "致命的エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Close();
                    return;
                }

                this.Activate();
                QRCodePanel.Enabled = _appSettings.IsAuthorizedUser;

                // 管理者のみマスター管理メニューを有効にする
                //MasterManagementToolStripMenuItem.Enabled = _appSettings.IsAdministrator;
                MasterManagementToolStripMenuItem.Enabled = true;
                QRCodeTextBox.Select();
            } catch (Exception ex) {
                Logger.AppendErrorLog(nameof(LoadEvents), ex, "予期しないエラーが発生しました");
                MessageBox.Show($"予期しないエラーが発生しました:\n{ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
        }
        // 実行中のEXEファイルをロックして二重起動を防止する
        private static void LockSelf() {
            try {
                string exePath = Application.ExecutablePath;
                _lockStream = new FileStream(exePath, FileMode.Open, FileAccess.Read, FileShare.Read);
            } catch (Exception ex) {
                Logger.AppendErrorLog(nameof(LockSelf), ex, "ファイルロック失敗。二重起動検出が機能しない可能性があります");
            }
        }
        // ラジオボタンのモードに応じて選択品目の基板登録・製品登録・再印刷・基板変更ウィンドウを開く
        private void Registration() {
            ResetFields();

            if (CategoryListBox3.SelectedItem is not ListItem<long> item) { return; }

            switch (RadioButtonNumber) {
                case RadioButtonMode.Substrate:
                    HandleSubstrateRegistration(item.Id);
                    break;

                case RadioButtonMode.ProductRegister:
                    HandleProductRegistration(item.Id, ProductOperationMode.Register);
                    break;

                case RadioButtonMode.RePrint:
                    HandleProductRegistration(item.Id, ProductOperationMode.RePrint);
                    break;

                case RadioButtonMode.SubstrateChange:
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
                case RadioButtonMode.Substrate:
                    _substrateMaster.CategoryName = CategoryListBox1.SelectedItem?.ToString() ?? string.Empty;
                    _substrateMaster.ProductName = CategoryListBox2.SelectedItem?.ToString() ?? string.Empty;
                    break;
                case RadioButtonMode.ProductRegister or RadioButtonMode.RePrint or RadioButtonMode.SubstrateChange:
                    _productMaster.CategoryName = CategoryListBox1.SelectedItem?.ToString() ?? string.Empty;
                    _productMaster.ProductName = CategoryListBox2.SelectedItem?.ToString() ?? string.Empty;
                    break;
            }
        }
        // 品目選択時に該当マスターデータをDBから読み込む
        private void LoadHistoryWithSelection(long itemId) {
            switch (RadioButtonNumber) {
                case RadioButtonMode.Substrate:
                    _substrateMaster.LoadFrom(_productRepository.GetSubstrateById(itemId));
                    break;
                case RadioButtonMode.ProductRegister or RadioButtonMode.RePrint or RadioButtonMode.SubstrateChange:
                    _productMaster.LoadFrom(_productRepository.GetProductById(itemId));
                    break;
                default:
                    throw new InvalidOperationException("不正なモードです");
            }
        }

        private record CategoryConfig(string OrderKey, string IdKey, string NameKey);
        private readonly Dictionary<RadioButtonMode, CategoryConfig> _categoryConfigs = new() {
            { RadioButtonMode.Substrate,        new CategoryConfig("SubstrateName", "SubstrateID", "SubstrateName") }, // 基板登録
            { RadioButtonMode.ProductRegister,  new CategoryConfig("ProductType",   "ProductID",   "ProductType")   }, // 製品登録
            { RadioButtonMode.RePrint,          new CategoryConfig("ProductType",   "ProductID",   "ProductType")   }, // 再印刷
            { RadioButtonMode.SubstrateChange,  new CategoryConfig("ProductType",   "ProductID",   "ProductType")   }  // 基板変更
        };
        // ラジオボタン選択時にモードに応じたマスターデータをフィルタしてCategoryListBox1にカテゴリ一覧を表示する
        private void CategorySelect(RadioButtonMode mode) {

            RegisterButton.Enabled = false;
            HistoryButton.Enabled = false;
            CategoryListBox1.Items.Clear();
            CategoryListBox2.Items.Clear();
            CategoryListBox3.Items.Clear();

            RadioButtonNumber = mode;

            // 未定義モード防止
            if (!_categoryConfigs.ContainsKey(RadioButtonNumber)) {
                _currentTargetRows = [];
                return;
            }

            // データソース切替
            bool isSubstrateMode = RadioButtonNumber == RadioButtonMode.Substrate;

            var sourceTable = isSubstrateMode
                ? _productRepository.SubstrateDataTable
                : _productRepository.ProductDataTable;

            // フィルタ済み行を保持
            _currentTargetRows = sourceTable
                .AsEnumerable()
                .Where(r => r.Field<long?>("Visible") == 1)
                .Where(r => RadioButtonNumber switch {
                    RadioButtonMode.Substrate => true,
                    RadioButtonMode.ProductRegister => true,
                    RadioButtonMode.RePrint => r.Field<long?>("SerialPrintType") is long spt && spt != 0,
                    RadioButtonMode.SubstrateChange => r.Field<long?>("SheetPrintType") is long shp && (shp == 2 || shp == 3),
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
            HistoryButton.Enabled = RadioButtonNumber != RadioButtonMode.SubstrateChange;
        }
        // カテゴリ選択時に一致する製品名またはSubstrateNameの一覧をListBox2に表示する
        private void CategoryListBox1Select() {
            RegisterButton.Enabled = false;
            HistoryButton.Enabled = RadioButtonNumber != RadioButtonMode.SubstrateChange;
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
            HistoryButton.Enabled = RadioButtonNumber != RadioButtonMode.SubstrateChange;
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
                ErrorMessageHelper.Show(ex);
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

        private void MainWindow_Load(object sender, EventArgs e) { LoadEvents(); }
        private void ReloadToolStripMenuItem_Click(object sender, EventArgs e) { LoadEvents(); }
        private void ConfigReportToolStripMenuItem_Click(object sender, EventArgs e) {
            var reportConfigPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config", "General", "Excel", "ConfigReport.xlsm");
            ExcelLauncher.Open(reportConfigPath);
        }
        private void ConfigListToolStripMenuItem_Click(object sender, EventArgs e) {
            var listConfigPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config", "General", "Excel", "ConfigList.xlsm");
            ExcelLauncher.Open(listConfigPath);
        }
        private void ConfigCheckSheetToolStripMenuItem_Click(object sender, EventArgs e) {
            var checkSheetConfigPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config", "General", "Excel", "ConfigCheckSheet.xlsm");
            ExcelLauncher.Open(checkSheetConfigPath);
        }
        private void ConfigSubstrateInformationToolStripMenuItem_Click(object sender, EventArgs e) {
            var checkSheetConfigPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config", "General", "Excel", "ConfigSubstrateInformation.xlsm");
            ExcelLauncher.Open(checkSheetConfigPath);
        }
        private void 終了ToolStripMenuItem_Click(object sender, EventArgs e) { Close(); }
        private void RegisterButton_Click(object sender, EventArgs e) { Registration(); }
        private void HistoryButton_Click(object sender, EventArgs e) { History(); }
        private void CategoryListBox1_SelectedIndexChanged(object sender, EventArgs e) { CategoryListBox1Select(); }
        private void CategoryListBox2_SelectedIndexChanged(object sender, EventArgs e) { CategoryListBox2Select(); }
        private void CategoryListBox3_SelectedIndexChanged(object sender, EventArgs e) { CategoryListBox3Select(); }
        private void CategoryRadioButton1_CheckedChanged(object sender, EventArgs e) { if (CategoryRadioButton1.Checked) { CategorySelect(RadioButtonMode.Substrate); } }
        private void CategoryRadioButton2_CheckedChanged(object sender, EventArgs e) { if (CategoryRadioButton2.Checked) { CategorySelect(RadioButtonMode.ProductRegister); } }
        private void CategoryRadioButton3_CheckedChanged(object sender, EventArgs e) { if (CategoryRadioButton3.Checked) { CategorySelect(RadioButtonMode.RePrint); } }
        private void CategoryRadioButton4_CheckedChanged(object sender, EventArgs e) { if (CategoryRadioButton4.Checked) { CategorySelect(RadioButtonMode.SubstrateChange); } }
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
            using var window = new LogViewerWindow();
            window.ShowDialog(this);
        }

        // マスター管理画面を管理者専用ダイアログで開く
        private void MasterManagementToolStripMenuItem_Click(object sender, EventArgs e) {
            try {
                using var window = new MasterManagement.MasterManagementWindow(_productRepository, _appSettings);
                window.ShowDialog(this);
                // マスターデータが変更されている可能性があるためキャッシュをクリアして再読み込みする
                _productRepository.Clear();
                _productRepository.LoadAll();
            } catch (Exception ex) {
                ErrorMessageHelper.Show(ex);
            }
        }

        private void PersonManagementToolStripMenuItem_Click(object sender, EventArgs e) {
            try {
                using var window = new MasterManagement.PersonManagementWindow();
                window.ShowDialog(this);
            } catch (Exception ex) {
                ErrorMessageHelper.Show(ex);
            }
        }

    }
}
