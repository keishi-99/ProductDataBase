namespace ProductDatabase.Other {
    public partial class InputDialog1 : Form {

        public string Temperature { get; private set; } = string.Empty;
        public string Humidity { get; private set; } = string.Empty;

        public InputDialog1() {
            InitializeComponent();
        }

        // ロード時にフォーカスを温度欄にセットしEnterキーで次欄・OKへ移動するキーハンドラを登録する
        private void LoadEvents() {
            ActiveControl = TemperatureTextBox;
            TemperatureTextBox.KeyPress += (sender, e) => {
                if (e.KeyChar == (char)Keys.Enter) {
                    HumidityTextBox.Focus();
                }
            };
            HumidityTextBox.KeyPress += (sender, e) => {
                if (e.KeyChar == (char)Keys.Enter) {
                    OkButton();
                }
            };
        }

        // 温度・湿度の入力値をプロパティに保存してダイアログをOKで閉じる
        private void OkButton() {
            Temperature = TemperatureTextBox.Text;
            Humidity = HumidityTextBox.Text;
            DialogResult = DialogResult.OK;
            return;
        }

        private void OKButton_Click(object sender, EventArgs e) { OkButton(); }
        private void InputDialog1_Load(object sender, EventArgs e) { LoadEvents(); }
    }
}
