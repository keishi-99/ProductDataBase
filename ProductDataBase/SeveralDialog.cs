namespace ProductDataBase {
    public partial class SeveralDialog : Form {
        public SeveralDialog() {
            InitializeComponent();
        }

        //// 別のフォームで選択された値を格納するプロパティ
        public int SelectedIndex = 0;

        private void OkButton_Click(object sender, EventArgs e) {
            SelectedIndex = SeveralListBox.SelectedIndex;
            DialogResult = DialogResult.OK;
            Close();
        }
    }
}
