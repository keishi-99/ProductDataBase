namespace ProductDatabase {
    partial class HistoryWindow {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing) {
            if (disposing && (components != null)) {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent() {
            this.DataBaseDataGridView = new DataGridView();
            this.CategoryComboBox = new ComboBox();
            this.FilterStringTextBox = new TextBox();
            this.CategoryRadioButton1 = new RadioButton();
            this.CategoryRadioButton2 = new RadioButton();
            this.StockCheckBox = new CheckBox();
            this.AllSubstrateCheckBox = new CheckBox();
            ((System.ComponentModel.ISupportInitialize)this.DataBaseDataGridView).BeginInit();
            this.SuspendLayout();
            // 
            // DataBaseDataGridView
            // 
            this.DataBaseDataGridView.AllowUserToAddRows = false;
            this.DataBaseDataGridView.AllowUserToDeleteRows = false;
            this.DataBaseDataGridView.AllowUserToResizeRows = false;
            this.DataBaseDataGridView.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            this.DataBaseDataGridView.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DataBaseDataGridView.Location = new Point(0, 0);
            this.DataBaseDataGridView.Name = "DataBaseDataGridView";
            this.DataBaseDataGridView.ReadOnly = true;
            this.DataBaseDataGridView.RowHeadersVisible = false;
            this.DataBaseDataGridView.RowTemplate.Height = 21;
            this.DataBaseDataGridView.Size = new Size(1184, 468);
            this.DataBaseDataGridView.TabIndex = 0;
            // 
            // CategoryComboBox
            // 
            this.CategoryComboBox.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            this.CategoryComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            this.CategoryComboBox.FormattingEnabled = true;
            this.CategoryComboBox.Location = new Point(12, 476);
            this.CategoryComboBox.Name = "CategoryComboBox";
            this.CategoryComboBox.Size = new Size(121, 23);
            this.CategoryComboBox.TabIndex = 1;
            // 
            // FilterStringTextBox
            // 
            this.FilterStringTextBox.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            this.FilterStringTextBox.Location = new Point(139, 477);
            this.FilterStringTextBox.MaxLength = 50;
            this.FilterStringTextBox.Name = "FilterStringTextBox";
            this.FilterStringTextBox.Size = new Size(100, 23);
            this.FilterStringTextBox.TabIndex = 2;
            this.FilterStringTextBox.TextChanged += this.FilterStringTextBox_TextChanged;
            // 
            // CategoryRadioButton1
            // 
            this.CategoryRadioButton1.Appearance = Appearance.Button;
            this.CategoryRadioButton1.Location = new Point(261, 474);
            this.CategoryRadioButton1.Name = "CategoryRadioButton1";
            this.CategoryRadioButton1.Size = new Size(92, 25);
            this.CategoryRadioButton1.TabIndex = 3;
            this.CategoryRadioButton1.TabStop = true;
            this.CategoryRadioButton1.Tag = "1";
            this.CategoryRadioButton1.Text = "登録履歴";
            this.CategoryRadioButton1.TextAlign = ContentAlignment.MiddleCenter;
            this.CategoryRadioButton1.UseVisualStyleBackColor = true;
            this.CategoryRadioButton1.CheckedChanged += this.CategoryRadioButton_CheckedChanged;
            // 
            // CategoryRadioButton2
            // 
            this.CategoryRadioButton2.Appearance = Appearance.Button;
            this.CategoryRadioButton2.Location = new Point(359, 474);
            this.CategoryRadioButton2.Name = "CategoryRadioButton2";
            this.CategoryRadioButton2.Size = new Size(92, 25);
            this.CategoryRadioButton2.TabIndex = 4;
            this.CategoryRadioButton2.TabStop = true;
            this.CategoryRadioButton2.Tag = "2";
            this.CategoryRadioButton2.Text = "Button2";
            this.CategoryRadioButton2.TextAlign = ContentAlignment.MiddleCenter;
            this.CategoryRadioButton2.UseVisualStyleBackColor = true;
            this.CategoryRadioButton2.CheckedChanged += this.CategoryRadioButton_CheckedChanged;
            // 
            // StockCheckBox
            // 
            this.StockCheckBox.AutoSize = true;
            this.StockCheckBox.Location = new Point(457, 479);
            this.StockCheckBox.Name = "StockCheckBox";
            this.StockCheckBox.Size = new Size(83, 19);
            this.StockCheckBox.TabIndex = 5;
            this.StockCheckBox.Text = "在庫有のみ";
            this.StockCheckBox.UseVisualStyleBackColor = true;
            this.StockCheckBox.CheckedChanged += this.StockCheckBox_CheckedChanged;
            // 
            // AllSubstrateCheckBox
            // 
            this.AllSubstrateCheckBox.AutoSize = true;
            this.AllSubstrateCheckBox.Location = new Point(546, 479);
            this.AllSubstrateCheckBox.Name = "AllSubstrateCheckBox";
            this.AllSubstrateCheckBox.Size = new Size(62, 19);
            this.AllSubstrateCheckBox.TabIndex = 6;
            this.AllSubstrateCheckBox.Text = "他基板";
            this.AllSubstrateCheckBox.UseVisualStyleBackColor = true;
            this.AllSubstrateCheckBox.CheckedChanged += this.AllSubstrateCheckBox_CheckedChanged;
            // 
            // HistoryWindow
            // 
            this.AutoScaleDimensions = new SizeF(7F, 15F);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.ClientSize = new Size(1184, 511);
            this.Controls.Add(this.AllSubstrateCheckBox);
            this.Controls.Add(this.StockCheckBox);
            this.Controls.Add(this.CategoryRadioButton2);
            this.Controls.Add(this.CategoryRadioButton1);
            this.Controls.Add(this.FilterStringTextBox);
            this.Controls.Add(this.CategoryComboBox);
            this.Controls.Add(this.DataBaseDataGridView);
            this.Font = new Font("Meiryo UI", 9F);
            this.MinimizeBox = false;
            this.Name = "HistoryWindow";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = FormStartPosition.CenterParent;
            this.Text = "データベース";
            this.Load += this.HistoryWindow_Load;
            ((System.ComponentModel.ISupportInitialize)this.DataBaseDataGridView).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion

        private DataGridView DataBaseDataGridView;
        private ComboBox CategoryComboBox;
        private TextBox FilterStringTextBox;
        private RadioButton CategoryRadioButton1;
        private RadioButton CategoryRadioButton2;
        private CheckBox StockCheckBox;
        private CheckBox AllSubstrateCheckBox;
    }
}