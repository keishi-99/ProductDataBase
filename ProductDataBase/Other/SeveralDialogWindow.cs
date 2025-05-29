using static ProductDatabase.MainWindow;

namespace ProductDatabase {
    public partial class SeveralDialogWindow : Form {

        public ProductInformation ProductInfo { get; }

        public int SelectedIndex { get; private set; } = -1;

        public SeveralDialogWindow(ProductInformation productInfo) {
            InitializeComponent();
            ProductInfo = productInfo;
        }

        // ロードイベント
        private void LoadEvents() {
            SeveralListBox.Items.Clear();
            Font = new Font(ProductInfo.FontName, ProductInfo.FontSize);
            var j = 0;

            foreach (var category in ProductInfo.CategoryItemNumber) {
                var isSubstrate = ProductInfo.CategoryType[j] == "1";
                var type = isSubstrate ? "基板登録" : "製品登録";
                var textString = isSubstrate ? ProductInfo.CategorySubstrateName[j] : ProductInfo.CategoryProductType[j];

                if (!string.IsNullOrEmpty(ProductInfo.CategoryItemNumber[j])) {
                    var itemText = $"[{type}]  [{ProductInfo.CategoryProductName[j]}]  [{textString}]  [{category}]";

                    // 重複チェック
                    if (!SeveralListBox.Items.Contains(itemText)) {
                        SeveralListBox.Items.Add(itemText);
                    }
                }
                j++;
            }
            SeveralListBox.SelectedIndex = 0;
        }

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
