namespace ProductDatabase.LogViewer {
    partial class LogViewerWindow {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing) {
            if (disposing && (components != null)) {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent() {
            this.TopPanel = new Panel();
            this.YearMonthLabel = new Label();
            this.YearMonthComboBox = new ComboBox();
            this.OperationTypeLabel = new Label();
            this.OperationTypeComboBox = new ComboBox();
            this.SearchLabel = new Label();
            this.SearchTextBox = new TextBox();
            this.SearchButton = new Button();
            this.LogDataGridView = new DataGridView();
            this.BottomStatusStrip = new StatusStrip();
            this.CountLabel = new ToolStripStatusLabel();

            this.TopPanel.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)this.LogDataGridView).BeginInit();
            this.BottomStatusStrip.SuspendLayout();
            this.SuspendLayout();

            // TopPanel
            this.TopPanel.Controls.Add(this.YearMonthLabel);
            this.TopPanel.Controls.Add(this.YearMonthComboBox);
            this.TopPanel.Controls.Add(this.OperationTypeLabel);
            this.TopPanel.Controls.Add(this.OperationTypeComboBox);
            this.TopPanel.Controls.Add(this.SearchLabel);
            this.TopPanel.Controls.Add(this.SearchTextBox);
            this.TopPanel.Controls.Add(this.SearchButton);
            this.TopPanel.Dock = DockStyle.Top;
            this.TopPanel.Height = 40;
            this.TopPanel.Name = "TopPanel";
            this.TopPanel.TabIndex = 0;

            // YearMonthLabel
            this.YearMonthLabel.AutoSize = true;
            this.YearMonthLabel.Location = new Point(10, 12);
            this.YearMonthLabel.Name = "YearMonthLabel";
            this.YearMonthLabel.Text = "年月：";

            // YearMonthComboBox
            this.YearMonthComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            this.YearMonthComboBox.Location = new Point(52, 8);
            this.YearMonthComboBox.Name = "YearMonthComboBox";
            this.YearMonthComboBox.Size = new Size(130, 23);
            this.YearMonthComboBox.TabIndex = 1;
            this.YearMonthComboBox.SelectedIndexChanged += this.YearMonthComboBox_SelectedIndexChanged;

            // OperationTypeLabel
            this.OperationTypeLabel.AutoSize = true;
            this.OperationTypeLabel.Location = new Point(200, 12);
            this.OperationTypeLabel.Name = "OperationTypeLabel";
            this.OperationTypeLabel.Text = "操作種別：";

            // OperationTypeComboBox
            this.OperationTypeComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            this.OperationTypeComboBox.Location = new Point(268, 8);
            this.OperationTypeComboBox.Name = "OperationTypeComboBox";
            this.OperationTypeComboBox.Size = new Size(150, 23);
            this.OperationTypeComboBox.TabIndex = 2;
            this.OperationTypeComboBox.SelectedIndexChanged += this.OperationTypeComboBox_SelectedIndexChanged;

            // SearchLabel
            this.SearchLabel.AutoSize = true;
            this.SearchLabel.Location = new Point(435, 12);
            this.SearchLabel.Name = "SearchLabel";
            this.SearchLabel.Text = "検索：";

            // SearchTextBox
            this.SearchTextBox.Location = new Point(472, 8);
            this.SearchTextBox.Name = "SearchTextBox";
            this.SearchTextBox.Size = new Size(200, 23);
            this.SearchTextBox.TabIndex = 3;
            this.SearchTextBox.KeyDown += this.SearchTextBox_KeyDown;

            // SearchButton
            this.SearchButton.Location = new Point(680, 7);
            this.SearchButton.Name = "SearchButton";
            this.SearchButton.Size = new Size(60, 25);
            this.SearchButton.TabIndex = 4;
            this.SearchButton.Text = "検索";
            this.SearchButton.UseVisualStyleBackColor = true;
            this.SearchButton.Click += this.SearchButton_Click;

            // LogDataGridView
            this.LogDataGridView.BackgroundColor = Color.White;
            this.LogDataGridView.BorderStyle = BorderStyle.None;
            this.LogDataGridView.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.LogDataGridView.Dock = DockStyle.Fill;
            this.LogDataGridView.Name = "LogDataGridView";
            this.LogDataGridView.TabIndex = 5;

            // BottomStatusStrip
            this.BottomStatusStrip.Items.AddRange(new ToolStripItem[] { this.CountLabel });
            this.BottomStatusStrip.Location = new Point(0, 628);
            this.BottomStatusStrip.Name = "BottomStatusStrip";
            this.BottomStatusStrip.SizingGrip = false;
            this.BottomStatusStrip.Size = new Size(1150, 22);
            this.BottomStatusStrip.TabIndex = 6;

            // CountLabel
            this.CountLabel.Name = "CountLabel";
            this.CountLabel.Size = new Size(89, 17);
            this.CountLabel.Text = "0 件表示中";

            // LogViewerWindow
            this.AutoScaleDimensions = new SizeF(7F, 15F);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.ClientSize = new Size(1150, 650);
            this.Controls.Add(this.LogDataGridView);
            this.Controls.Add(this.BottomStatusStrip);
            this.Controls.Add(this.TopPanel);
            this.Font = new Font("Meiryo UI", 9F);
            this.Name = "LogViewerWindow";
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Text = "ログ閲覧";

            this.TopPanel.ResumeLayout(false);
            this.TopPanel.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)this.LogDataGridView).EndInit();
            this.BottomStatusStrip.ResumeLayout(false);
            this.BottomStatusStrip.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        private Panel TopPanel;
        private Label YearMonthLabel;
        private ComboBox YearMonthComboBox;
        private Label OperationTypeLabel;
        private ComboBox OperationTypeComboBox;
        private Label SearchLabel;
        private TextBox SearchTextBox;
        private Button SearchButton;
        private DataGridView LogDataGridView;
        private StatusStrip BottomStatusStrip;
        private ToolStripStatusLabel CountLabel;
    }
}
