namespace ProductDatabase.MasterManagement {
    partial class SubstrateMasterEditDialog {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing) {
            if (disposing && (components != null)) {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        private void InitializeComponent() {
            this.BasicInfoGroupBox = new GroupBox();
            this.CategoryNameLabel = new Label();
            this.CategoryNameTextBox = new TextBox();
            this.ProductNameLabel = new Label();
            this.ProductNameTextBox = new TextBox();
            this.SubstrateNameLabel = new Label();
            this.SubstrateNameTextBox = new TextBox();
            this.SubstrateModelLabel = new Label();
            this.SubstrateModelTextBox = new TextBox();
            this.VisibleLabel = new Label();
            this.VisibleCheckBox = new CheckBox();
            this.SerialGroupBox = new GroupBox();
            this.RegTypeLabel = new Label();
            this.RegTypeComboBox = new ComboBox();
            this.PrintGroupBox = new GroupBox();
            this.LabelPrintCheckBox = new CheckBox();
            this.CheckBinGroupBox = new GroupBox();
            this.CheckBinCheckBox0 = new CheckBox();
            this.CheckBinCheckBox1 = new CheckBox();
            this.CheckBinCheckBox2 = new CheckBox();
            this.CheckBinCheckBox3 = new CheckBox();
            this.CheckBinCheckBox4 = new CheckBox();
            this.CheckBinCheckBox5 = new CheckBox();
            this.CheckBinCheckBox6 = new CheckBox();
            this.CheckBinCheckBox7 = new CheckBox();
            this.CheckBinCheckBox8 = new CheckBox();
            this.CheckBinCheckBox9 = new CheckBox();
            this.CheckBinCheckBox10 = new CheckBox();
            this.SaveButton = new Button();
            this.DialogCancelButton = new Button();
            this.BasicInfoGroupBox.SuspendLayout();
            this.SerialGroupBox.SuspendLayout();
            this.PrintGroupBox.SuspendLayout();
            this.CheckBinGroupBox.SuspendLayout();
            this.SuspendLayout();
            // 
            // BasicInfoGroupBox
            // 
            this.BasicInfoGroupBox.Controls.Add(this.CategoryNameLabel);
            this.BasicInfoGroupBox.Controls.Add(this.CategoryNameTextBox);
            this.BasicInfoGroupBox.Controls.Add(this.ProductNameLabel);
            this.BasicInfoGroupBox.Controls.Add(this.ProductNameTextBox);
            this.BasicInfoGroupBox.Controls.Add(this.SubstrateNameLabel);
            this.BasicInfoGroupBox.Controls.Add(this.SubstrateNameTextBox);
            this.BasicInfoGroupBox.Controls.Add(this.SubstrateModelLabel);
            this.BasicInfoGroupBox.Controls.Add(this.SubstrateModelTextBox);
            this.BasicInfoGroupBox.Controls.Add(this.VisibleLabel);
            this.BasicInfoGroupBox.Controls.Add(this.VisibleCheckBox);
            this.BasicInfoGroupBox.Location = new Point(12, 12);
            this.BasicInfoGroupBox.Name = "BasicInfoGroupBox";
            this.BasicInfoGroupBox.Size = new Size(395, 195);
            this.BasicInfoGroupBox.TabIndex = 0;
            this.BasicInfoGroupBox.TabStop = false;
            this.BasicInfoGroupBox.Text = "基本情報";
            // 
            // CategoryNameLabel
            // 
            this.CategoryNameLabel.AutoSize = true;
            this.CategoryNameLabel.Location = new Point(8, 30);
            this.CategoryNameLabel.Name = "CategoryNameLabel";
            this.CategoryNameLabel.Size = new Size(69, 15);
            this.CategoryNameLabel.TabIndex = 0;
            this.CategoryNameLabel.Text = "カテゴリ名 *:";
            // 
            // CategoryNameTextBox
            // 
            this.CategoryNameTextBox.Location = new Point(130, 27);
            this.CategoryNameTextBox.Name = "CategoryNameTextBox";
            this.CategoryNameTextBox.Size = new Size(250, 23);
            this.CategoryNameTextBox.TabIndex = 0;
            // 
            // ProductNameLabel
            // 
            this.ProductNameLabel.AutoSize = true;
            this.ProductNameLabel.Location = new Point(8, 62);
            this.ProductNameLabel.Name = "ProductNameLabel";
            this.ProductNameLabel.Size = new Size(59, 15);
            this.ProductNameLabel.TabIndex = 1;
            this.ProductNameLabel.Text = "製品名 *:";
            // 
            // ProductNameTextBox
            // 
            this.ProductNameTextBox.Location = new Point(130, 59);
            this.ProductNameTextBox.Name = "ProductNameTextBox";
            this.ProductNameTextBox.Size = new Size(250, 23);
            this.ProductNameTextBox.TabIndex = 1;
            // 
            // SubstrateNameLabel
            // 
            this.SubstrateNameLabel.AutoSize = true;
            this.SubstrateNameLabel.Location = new Point(8, 94);
            this.SubstrateNameLabel.Name = "SubstrateNameLabel";
            this.SubstrateNameLabel.Size = new Size(71, 15);
            this.SubstrateNameLabel.TabIndex = 2;
            this.SubstrateNameLabel.Text = "基板名称 *:";
            // 
            // SubstrateNameTextBox
            // 
            this.SubstrateNameTextBox.Location = new Point(130, 91);
            this.SubstrateNameTextBox.Name = "SubstrateNameTextBox";
            this.SubstrateNameTextBox.Size = new Size(250, 23);
            this.SubstrateNameTextBox.TabIndex = 2;
            // 
            // SubstrateModelLabel
            // 
            this.SubstrateModelLabel.AutoSize = true;
            this.SubstrateModelLabel.Location = new Point(8, 126);
            this.SubstrateModelLabel.Name = "SubstrateModelLabel";
            this.SubstrateModelLabel.Size = new Size(71, 15);
            this.SubstrateModelLabel.TabIndex = 3;
            this.SubstrateModelLabel.Text = "基板型式 *:";
            // 
            // SubstrateModelTextBox
            // 
            this.SubstrateModelTextBox.Location = new Point(130, 123);
            this.SubstrateModelTextBox.Name = "SubstrateModelTextBox";
            this.SubstrateModelTextBox.Size = new Size(250, 23);
            this.SubstrateModelTextBox.TabIndex = 3;
            // 
            // VisibleLabel
            // 
            this.VisibleLabel.AutoSize = true;
            this.VisibleLabel.Location = new Point(8, 160);
            this.VisibleLabel.Name = "VisibleLabel";
            this.VisibleLabel.Size = new Size(36, 15);
            this.VisibleLabel.TabIndex = 4;
            this.VisibleLabel.Text = "表示:";
            // 
            // VisibleCheckBox
            // 
            this.VisibleCheckBox.Location = new Point(130, 158);
            this.VisibleCheckBox.Name = "VisibleCheckBox";
            this.VisibleCheckBox.Size = new Size(120, 22);
            this.VisibleCheckBox.TabIndex = 4;
            this.VisibleCheckBox.Text = "表示する";
            // 
            // SerialGroupBox
            // 
            this.SerialGroupBox.Controls.Add(this.RegTypeLabel);
            this.SerialGroupBox.Controls.Add(this.RegTypeComboBox);
            this.SerialGroupBox.Location = new Point(12, 217);
            this.SerialGroupBox.Name = "SerialGroupBox";
            this.SerialGroupBox.Size = new Size(395, 62);
            this.SerialGroupBox.TabIndex = 1;
            this.SerialGroupBox.TabStop = false;
            this.SerialGroupBox.Text = "在庫登録設定";
            // 
            // RegTypeLabel
            // 
            this.RegTypeLabel.AutoSize = true;
            this.RegTypeLabel.Location = new Point(8, 27);
            this.RegTypeLabel.Name = "RegTypeLabel";
            this.RegTypeLabel.Size = new Size(62, 15);
            this.RegTypeLabel.TabIndex = 0;
            this.RegTypeLabel.Text = "登録タイプ:";
            // 
            // RegTypeComboBox
            // 
            this.RegTypeComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            this.RegTypeComboBox.Location = new Point(130, 24);
            this.RegTypeComboBox.Name = "RegTypeComboBox";
            this.RegTypeComboBox.Size = new Size(250, 23);
            this.RegTypeComboBox.TabIndex = 4;
            // 
            // PrintGroupBox
            // 
            this.PrintGroupBox.Controls.Add(this.LabelPrintCheckBox);
            this.PrintGroupBox.Location = new Point(12, 289);
            this.PrintGroupBox.Name = "PrintGroupBox";
            this.PrintGroupBox.Size = new Size(395, 62);
            this.PrintGroupBox.TabIndex = 2;
            this.PrintGroupBox.TabStop = false;
            this.PrintGroupBox.Text = "印刷設定";
            // 
            // LabelPrintCheckBox
            // 
            this.LabelPrintCheckBox.Location = new Point(8, 27);
            this.LabelPrintCheckBox.Name = "LabelPrintCheckBox";
            this.LabelPrintCheckBox.Size = new Size(120, 22);
            this.LabelPrintCheckBox.TabIndex = 5;
            this.LabelPrintCheckBox.Text = "ラベル印刷";
            // 
            // CheckBinGroupBox
            // 
            this.CheckBinGroupBox.Controls.Add(this.CheckBinCheckBox0);
            this.CheckBinGroupBox.Controls.Add(this.CheckBinCheckBox1);
            this.CheckBinGroupBox.Controls.Add(this.CheckBinCheckBox2);
            this.CheckBinGroupBox.Controls.Add(this.CheckBinCheckBox3);
            this.CheckBinGroupBox.Controls.Add(this.CheckBinCheckBox4);
            this.CheckBinGroupBox.Controls.Add(this.CheckBinCheckBox5);
            this.CheckBinGroupBox.Controls.Add(this.CheckBinCheckBox6);
            this.CheckBinGroupBox.Controls.Add(this.CheckBinCheckBox7);
            this.CheckBinGroupBox.Controls.Add(this.CheckBinCheckBox8);
            this.CheckBinGroupBox.Controls.Add(this.CheckBinCheckBox9);
            this.CheckBinGroupBox.Controls.Add(this.CheckBinCheckBox10);
            this.CheckBinGroupBox.Location = new Point(12, 361);
            this.CheckBinGroupBox.Name = "CheckBinGroupBox";
            this.CheckBinGroupBox.Size = new Size(395, 164);
            this.CheckBinGroupBox.TabIndex = 3;
            this.CheckBinGroupBox.TabStop = false;
            this.CheckBinGroupBox.Text = "登録フォーム表示項目";
            // 
            // CheckBinCheckBox0
            // 
            this.CheckBinCheckBox0.Location = new Point(8, 22);
            this.CheckBinCheckBox0.Name = "CheckBinCheckBox0";
            this.CheckBinCheckBox0.Size = new Size(100, 22);
            this.CheckBinCheckBox0.TabIndex = 10;
            this.CheckBinCheckBox0.Text = "注文番号";
            // 
            // CheckBinCheckBox1
            // 
            this.CheckBinCheckBox1.Location = new Point(8, 50);
            this.CheckBinCheckBox1.Name = "CheckBinCheckBox1";
            this.CheckBinCheckBox1.Size = new Size(100, 22);
            this.CheckBinCheckBox1.TabIndex = 11;
            this.CheckBinCheckBox1.Text = "製造番号";
            // 
            // CheckBinCheckBox2
            // 
            this.CheckBinCheckBox2.Location = new Point(8, 78);
            this.CheckBinCheckBox2.Name = "CheckBinCheckBox2";
            this.CheckBinCheckBox2.Size = new Size(100, 22);
            this.CheckBinCheckBox2.TabIndex = 12;
            this.CheckBinCheckBox2.Text = "追加量";
            // 
            // CheckBinCheckBox3
            // 
            this.CheckBinCheckBox3.Location = new Point(8, 106);
            this.CheckBinCheckBox3.Name = "CheckBinCheckBox3";
            this.CheckBinCheckBox3.Size = new Size(100, 22);
            this.CheckBinCheckBox3.TabIndex = 13;
            this.CheckBinCheckBox3.Text = "減少量";
            // 
            // CheckBinCheckBox4
            // 
            this.CheckBinCheckBox4.Location = new Point(8, 134);
            this.CheckBinCheckBox4.Name = "CheckBinCheckBox4";
            this.CheckBinCheckBox4.Size = new Size(100, 22);
            this.CheckBinCheckBox4.TabIndex = 14;
            this.CheckBinCheckBox4.Text = "予備";
            // 
            // CheckBinCheckBox5
            // 
            this.CheckBinCheckBox5.Location = new Point(138, 22);
            this.CheckBinCheckBox5.Name = "CheckBinCheckBox5";
            this.CheckBinCheckBox5.Size = new Size(100, 22);
            this.CheckBinCheckBox5.TabIndex = 15;
            this.CheckBinCheckBox5.Text = "予備";
            // 
            // CheckBinCheckBox6
            // 
            this.CheckBinCheckBox6.Location = new Point(138, 50);
            this.CheckBinCheckBox6.Name = "CheckBinCheckBox6";
            this.CheckBinCheckBox6.Size = new Size(100, 22);
            this.CheckBinCheckBox6.TabIndex = 16;
            this.CheckBinCheckBox6.Text = "予備";
            // 
            // CheckBinCheckBox7
            // 
            this.CheckBinCheckBox7.Location = new Point(138, 78);
            this.CheckBinCheckBox7.Name = "CheckBinCheckBox7";
            this.CheckBinCheckBox7.Size = new Size(100, 22);
            this.CheckBinCheckBox7.TabIndex = 17;
            this.CheckBinCheckBox7.Text = "予備";
            // 
            // CheckBinCheckBox8
            // 
            this.CheckBinCheckBox8.Location = new Point(138, 106);
            this.CheckBinCheckBox8.Name = "CheckBinCheckBox8";
            this.CheckBinCheckBox8.Size = new Size(100, 22);
            this.CheckBinCheckBox8.TabIndex = 18;
            this.CheckBinCheckBox8.Text = "登録日";
            // 
            // CheckBinCheckBox9
            // 
            this.CheckBinCheckBox9.Location = new Point(138, 134);
            this.CheckBinCheckBox9.Name = "CheckBinCheckBox9";
            this.CheckBinCheckBox9.Size = new Size(100, 22);
            this.CheckBinCheckBox9.TabIndex = 19;
            this.CheckBinCheckBox9.Text = "担当者";
            // 
            // CheckBinCheckBox10
            // 
            this.CheckBinCheckBox10.Location = new Point(268, 22);
            this.CheckBinCheckBox10.Name = "CheckBinCheckBox10";
            this.CheckBinCheckBox10.Size = new Size(100, 22);
            this.CheckBinCheckBox10.TabIndex = 20;
            this.CheckBinCheckBox10.Text = "コメント";
            // 
            // SaveButton
            // 
            this.SaveButton.Location = new Point(198, 531);
            this.SaveButton.Name = "SaveButton";
            this.SaveButton.Size = new Size(100, 35);
            this.SaveButton.TabIndex = 21;
            this.SaveButton.Text = "保存";
            this.SaveButton.UseVisualStyleBackColor = true;
            this.SaveButton.Click += this.SaveButton_Click;
            // 
            // DialogCancelButton
            // 
            this.DialogCancelButton.Location = new Point(308, 531);
            this.DialogCancelButton.Name = "DialogCancelButton";
            this.DialogCancelButton.Size = new Size(100, 35);
            this.DialogCancelButton.TabIndex = 22;
            this.DialogCancelButton.Text = "キャンセル";
            this.DialogCancelButton.UseVisualStyleBackColor = true;
            this.DialogCancelButton.Click += this.CancelButton_Click;
            // 
            // SubstrateMasterEditDialog
            // 
            this.AutoScaleDimensions = new SizeF(7F, 15F);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.ClientSize = new Size(420, 576);
            this.Controls.Add(this.BasicInfoGroupBox);
            this.Controls.Add(this.SerialGroupBox);
            this.Controls.Add(this.PrintGroupBox);
            this.Controls.Add(this.CheckBinGroupBox);
            this.Controls.Add(this.SaveButton);
            this.Controls.Add(this.DialogCancelButton);
            this.Font = new Font("Meiryo UI", 9F);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "SubstrateMasterEditDialog";
            this.StartPosition = FormStartPosition.CenterParent;
            this.Text = "基板マスター編集";
            this.Load += this.SubstrateMasterEditDialog_Load;
            this.BasicInfoGroupBox.ResumeLayout(false);
            this.BasicInfoGroupBox.PerformLayout();
            this.SerialGroupBox.ResumeLayout(false);
            this.SerialGroupBox.PerformLayout();
            this.PrintGroupBox.ResumeLayout(false);
            this.CheckBinGroupBox.ResumeLayout(false);
            this.ResumeLayout(false);
        }

        #endregion

        // セクション1
        private GroupBox BasicInfoGroupBox;
        private Label CategoryNameLabel;
        private TextBox CategoryNameTextBox;
        private Label ProductNameLabel;
        private TextBox ProductNameTextBox;
        private Label SubstrateNameLabel;
        private TextBox SubstrateNameTextBox;
        private Label SubstrateModelLabel;
        private TextBox SubstrateModelTextBox;
        private Label VisibleLabel;
        private CheckBox VisibleCheckBox;

        // セクション2
        private GroupBox SerialGroupBox;
        private Label RegTypeLabel;
        private ComboBox RegTypeComboBox;

        // セクション3
        private GroupBox PrintGroupBox;
        private CheckBox LabelPrintCheckBox;

        // セクション4
        private GroupBox CheckBinGroupBox;
        private CheckBox CheckBinCheckBox0;
        private CheckBox CheckBinCheckBox1;
        private CheckBox CheckBinCheckBox2;
        private CheckBox CheckBinCheckBox3;
        private CheckBox CheckBinCheckBox4;
        private CheckBox CheckBinCheckBox5;
        private CheckBox CheckBinCheckBox6;
        private CheckBox CheckBinCheckBox7;
        private CheckBox CheckBinCheckBox8;
        private CheckBox CheckBinCheckBox9;
        private CheckBox CheckBinCheckBox10;

        // ボタン
        private Button SaveButton;
        private Button DialogCancelButton;
    }
}
