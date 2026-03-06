using static ProductDatabase.ProductRepository;

namespace ProductDatabase {
    public partial class SeveralDialogWindow : Form {

        private readonly QrSettings _qrSettings;
        private readonly AppSettings _appSettings;

        public int SelectedIndex { get; private set; } = -1;

        public SeveralDialogWindow(QrSettings qrSettings, AppSettings appSettings) {
            InitializeComponent();
            _qrSettings = qrSettings;
            _appSettings = appSettings;
        }

        // ロード時に品目候補リストを重複除去してListBoxに表示し先頭を選択する
        private void LoadEvents() {
            SeveralListBox.Items.Clear();
            Font = new Font(_appSettings.FontName, _appSettings.FontSize);
            var j = 0;

            foreach (var category in _qrSettings.CategoryItemNumber) {
                var isSubstrate = _qrSettings.CategoryType[j] == "1";
                var type = isSubstrate ? "基板登録" : "製品登録";
                var textString = isSubstrate ? _qrSettings.CategorySubstrateName[j] : _qrSettings.CategoryProductType[j];

                if (!string.IsNullOrEmpty(_qrSettings.CategoryItemNumber[j])) {
                    var itemText = $"[{type}]  [{_qrSettings.CategoryProductName[j]}]  [{textString}]  [{category}]";

                    // 重複チェック
                    if (!SeveralListBox.Items.Contains(itemText)) {
                        SeveralListBox.Items.Add(itemText);
                    }
                }
                j++;
            }
            SeveralListBox.SelectedIndex = 0;
        }

        // ListBoxの選択インデックスをSelectedIndexプロパティに保存する
        private void ReturnedIndex() { SelectedIndex = SeveralListBox.SelectedIndex; }

        private void SeveralDialogWindow_Load(object sender, EventArgs e) { LoadEvents(); }
        private void OkButton_Click(object sender, EventArgs e) {
            ReturnedIndex();
            Close();
        }
        private void SeveralListBox_KeyPress(object sender, KeyPressEventArgs e) {
            if (e.KeyChar == (char)Keys.Enter) { OKButton.PerformClick(); }
        }

        private void CloseButton_Click(object sender, EventArgs e) {
            Close();
        }
    }
}
