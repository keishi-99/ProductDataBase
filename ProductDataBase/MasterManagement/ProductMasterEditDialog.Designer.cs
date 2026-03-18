namespace ProductDatabase.MasterManagement {
    partial class ProductMasterEditDialog {
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
            this.ProductModelLabel = new Label();
            this.ProductModelTextBox = new TextBox();
            this.ProductTypeLabel = new Label();
            this.ProductTypeTextBox = new TextBox();
            this.InitialLabel = new Label();
            this.InitialTextBox = new TextBox();
            this.VisibleLabel = new Label();
            this.VisibleCheckBox = new CheckBox();
            this.SerialGroupBox = new GroupBox();
            this.RegTypeLabel = new Label();
            this.RegTypeComboBox = new ComboBox();
            this.SerialTypeLabel = new Label();
            this.SerialTypeComboBox = new ComboBox();
            this.RevisionGroupLabel = new Label();
            this.RevisionGroupNumericUpDown = new NumericUpDown();
            this.PrintGroupBox = new GroupBox();
            this.LabelPrintCheckBox = new CheckBox();
            this.BarcodePrintCheckBox = new CheckBox();
            this.NameplatePrintCheckBox = new CheckBox();
            this.UnderlinePrintCheckBox = new CheckBox();
            this.Last4DigitsPrintCheckBox = new CheckBox();
            this.CheckSheetPrintCheckBox = new CheckBox();
            this.ListPrintCheckBox = new CheckBox();
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
            this.UseSubstrateGroupBox = new GroupBox();
            this.SubstrateProductNameFilterLabel = new Label();
            this.SubstrateProductNameFilterTextBox = new TextBox();
            this.SubstrateModelFilterLabel = new Label();
            this.SubstrateModelFilterTextBox = new TextBox();
            this.SubstrateCheckedListBox = new CheckedListBox();
            this.SaveButton = new Button();
            this.DialogCancelButton = new Button();
            this.BasicInfoGroupBox.SuspendLayout();
            this.SerialGroupBox.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)this.RevisionGroupNumericUpDown).BeginInit();
            this.PrintGroupBox.SuspendLayout();
            this.CheckBinGroupBox.SuspendLayout();
            this.UseSubstrateGroupBox.SuspendLayout();
            this.SuspendLayout();
            // 
            // BasicInfoGroupBox
            // 
            this.BasicInfoGroupBox.Controls.Add(this.CategoryNameLabel);
            this.BasicInfoGroupBox.Controls.Add(this.CategoryNameTextBox);
            this.BasicInfoGroupBox.Controls.Add(this.ProductNameLabel);
            this.BasicInfoGroupBox.Controls.Add(this.ProductNameTextBox);
            this.BasicInfoGroupBox.Controls.Add(this.ProductModelLabel);
            this.BasicInfoGroupBox.Controls.Add(this.ProductModelTextBox);
            this.BasicInfoGroupBox.Controls.Add(this.ProductTypeLabel);
            this.BasicInfoGroupBox.Controls.Add(this.ProductTypeTextBox);
            this.BasicInfoGroupBox.Controls.Add(this.InitialLabel);
            this.BasicInfoGroupBox.Controls.Add(this.InitialTextBox);
            this.BasicInfoGroupBox.Controls.Add(this.VisibleLabel);
            this.BasicInfoGroupBox.Controls.Add(this.VisibleCheckBox);
            this.BasicInfoGroupBox.Location = new Point(12, 12);
            this.BasicInfoGroupBox.Name = "BasicInfoGroupBox";
            this.BasicInfoGroupBox.Size = new Size(736, 197);
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
            this.CategoryNameTextBox.Size = new Size(590, 23);
            this.CategoryNameTextBox.TabIndex = 0;
            // 
            // ProductNameLabel
            // 
            this.ProductNameLabel.AutoSize = true;
            this.ProductNameLabel.Location = new Point(8, 60);
            this.ProductNameLabel.Name = "ProductNameLabel";
            this.ProductNameLabel.Size = new Size(59, 15);
            this.ProductNameLabel.TabIndex = 1;
            this.ProductNameLabel.Text = "製品名 *:";
            // 
            // ProductNameTextBox
            // 
            this.ProductNameTextBox.Location = new Point(130, 57);
            this.ProductNameTextBox.Name = "ProductNameTextBox";
            this.ProductNameTextBox.Size = new Size(590, 23);
            this.ProductNameTextBox.TabIndex = 1;
            // 
            // ProductModelLabel
            // 
            this.ProductModelLabel.AutoSize = true;
            this.ProductModelLabel.Location = new Point(8, 90);
            this.ProductModelLabel.Name = "ProductModelLabel";
            this.ProductModelLabel.Size = new Size(71, 15);
            this.ProductModelLabel.TabIndex = 2;
            this.ProductModelLabel.Text = "製品型式 *:";
            // 
            // ProductModelTextBox
            // 
            this.ProductModelTextBox.Location = new Point(130, 87);
            this.ProductModelTextBox.Name = "ProductModelTextBox";
            this.ProductModelTextBox.Size = new Size(590, 23);
            this.ProductModelTextBox.TabIndex = 2;
            // 
            // ProductTypeLabel
            // 
            this.ProductTypeLabel.AutoSize = true;
            this.ProductTypeLabel.Location = new Point(8, 120);
            this.ProductTypeLabel.Name = "ProductTypeLabel";
            this.ProductTypeLabel.Size = new Size(62, 15);
            this.ProductTypeLabel.TabIndex = 3;
            this.ProductTypeLabel.Text = "製品タイプ *:";
            // 
            // ProductTypeTextBox
            // 
            this.ProductTypeTextBox.Location = new Point(130, 117);
            this.ProductTypeTextBox.Name = "ProductTypeTextBox";
            this.ProductTypeTextBox.Size = new Size(590, 23);
            this.ProductTypeTextBox.TabIndex = 3;
            // 
            // InitialLabel
            // 
            this.InitialLabel.AutoSize = true;
            this.InitialLabel.Location = new Point(8, 150);
            this.InitialLabel.Name = "InitialLabel";
            this.InitialLabel.Size = new Size(59, 15);
            this.InitialLabel.TabIndex = 4;
            this.InitialLabel.Text = "シリアル接頭文字:";
            // 
            // InitialTextBox
            // 
            this.InitialTextBox.Location = new Point(130, 147);
            this.InitialTextBox.Name = "InitialTextBox";
            this.InitialTextBox.Size = new Size(200, 23);
            this.InitialTextBox.TabIndex = 4;
            // 
            // VisibleLabel
            // 
            this.VisibleLabel.AutoSize = true;
            this.VisibleLabel.Location = new Point(8, 180);
            this.VisibleLabel.Name = "VisibleLabel";
            this.VisibleLabel.Size = new Size(61, 15);
            this.VisibleLabel.TabIndex = 5;
            this.VisibleLabel.Text = "表示フラグ:";
            // 
            // VisibleCheckBox
            // 
            this.VisibleCheckBox.AutoSize = true;
            this.VisibleCheckBox.Location = new Point(130, 178);
            this.VisibleCheckBox.Name = "VisibleCheckBox";
            this.VisibleCheckBox.Size = new Size(69, 19);
            this.VisibleCheckBox.TabIndex = 5;
            this.VisibleCheckBox.Text = "表示する";
            // 
            // SerialGroupBox
            // 
            this.SerialGroupBox.Controls.Add(this.RegTypeLabel);
            this.SerialGroupBox.Controls.Add(this.RegTypeComboBox);
            this.SerialGroupBox.Controls.Add(this.SerialTypeLabel);
            this.SerialGroupBox.Controls.Add(this.SerialTypeComboBox);
            this.SerialGroupBox.Controls.Add(this.RevisionGroupLabel);
            this.SerialGroupBox.Controls.Add(this.RevisionGroupNumericUpDown);
            this.SerialGroupBox.Location = new Point(12, 219);
            this.SerialGroupBox.Name = "SerialGroupBox";
            this.SerialGroupBox.Size = new Size(736, 128);
            this.SerialGroupBox.TabIndex = 1;
            this.SerialGroupBox.TabStop = false;
            this.SerialGroupBox.Text = "シリアル設定";
            // 
            // RegTypeLabel
            // 
            this.RegTypeLabel.AutoSize = true;
            this.RegTypeLabel.Location = new Point(8, 30);
            this.RegTypeLabel.Name = "RegTypeLabel";
            this.RegTypeLabel.Size = new Size(62, 15);
            this.RegTypeLabel.TabIndex = 0;
            this.RegTypeLabel.Text = "登録タイプ:";
            // 
            // RegTypeComboBox
            // 
            this.RegTypeComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            this.RegTypeComboBox.Location = new Point(130, 27);
            this.RegTypeComboBox.Name = "RegTypeComboBox";
            this.RegTypeComboBox.Size = new Size(250, 23);
            this.RegTypeComboBox.TabIndex = 7;
            // 
            // SerialTypeLabel
            // 
            this.SerialTypeLabel.AutoSize = true;
            this.SerialTypeLabel.Location = new Point(8, 62);
            this.SerialTypeLabel.Name = "SerialTypeLabel";
            this.SerialTypeLabel.Size = new Size(98, 15);
            this.SerialTypeLabel.TabIndex = 8;
            this.SerialTypeLabel.Text = "シリアル桁数タイプ:";
            // 
            // SerialTypeComboBox
            // 
            this.SerialTypeComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            this.SerialTypeComboBox.Location = new Point(130, 59);
            this.SerialTypeComboBox.Name = "SerialTypeComboBox";
            this.SerialTypeComboBox.Size = new Size(250, 23);
            this.SerialTypeComboBox.TabIndex = 8;
            // 
            // RevisionGroupLabel
            // 
            this.RevisionGroupLabel.AutoSize = true;
            this.RevisionGroupLabel.Location = new Point(8, 94);
            this.RevisionGroupLabel.Name = "RevisionGroupLabel";
            this.RevisionGroupLabel.Size = new Size(90, 15);
            this.RevisionGroupLabel.TabIndex = 9;
            this.RevisionGroupLabel.Text = "リビジョングループ:";
            // 
            // RevisionGroupNumericUpDown
            // 
            this.RevisionGroupNumericUpDown.Location = new Point(130, 91);
            this.RevisionGroupNumericUpDown.Maximum = new decimal(new int[] { 99, 0, 0, 0 });
            this.RevisionGroupNumericUpDown.Name = "RevisionGroupNumericUpDown";
            this.RevisionGroupNumericUpDown.Size = new Size(80, 23);
            this.RevisionGroupNumericUpDown.TabIndex = 9;
            // 
            // PrintGroupBox
            // 
            this.PrintGroupBox.Controls.Add(this.LabelPrintCheckBox);
            this.PrintGroupBox.Controls.Add(this.BarcodePrintCheckBox);
            this.PrintGroupBox.Controls.Add(this.NameplatePrintCheckBox);
            this.PrintGroupBox.Controls.Add(this.UnderlinePrintCheckBox);
            this.PrintGroupBox.Controls.Add(this.Last4DigitsPrintCheckBox);
            this.PrintGroupBox.Controls.Add(this.CheckSheetPrintCheckBox);
            this.PrintGroupBox.Controls.Add(this.ListPrintCheckBox);
            this.PrintGroupBox.Location = new Point(12, 357);
            this.PrintGroupBox.Name = "PrintGroupBox";
            this.PrintGroupBox.Size = new Size(736, 98);
            this.PrintGroupBox.TabIndex = 2;
            this.PrintGroupBox.TabStop = false;
            this.PrintGroupBox.Text = "印刷設定";
            // 
            // LabelPrintCheckBox
            //
            this.LabelPrintCheckBox.Location = new Point(8, 24);
            this.LabelPrintCheckBox.Name = "LabelPrintCheckBox";
            this.LabelPrintCheckBox.Size = new Size(140, 22);
            this.LabelPrintCheckBox.TabIndex = 10;
            this.LabelPrintCheckBox.Text = "ラベル印刷";
            //
            // BarcodePrintCheckBox
            //
            this.BarcodePrintCheckBox.Location = new Point(152, 24);
            this.BarcodePrintCheckBox.Name = "BarcodePrintCheckBox";
            this.BarcodePrintCheckBox.Size = new Size(140, 22);
            this.BarcodePrintCheckBox.TabIndex = 11;
            this.BarcodePrintCheckBox.Text = "バーコード印刷";
            //
            // NameplatePrintCheckBox
            //
            this.NameplatePrintCheckBox.Location = new Point(296, 24);
            this.NameplatePrintCheckBox.Name = "NameplatePrintCheckBox";
            this.NameplatePrintCheckBox.Size = new Size(140, 22);
            this.NameplatePrintCheckBox.TabIndex = 12;
            this.NameplatePrintCheckBox.Text = "銘板印刷";
            //
            // UnderlinePrintCheckBox
            //
            this.UnderlinePrintCheckBox.Location = new Point(440, 24);
            this.UnderlinePrintCheckBox.Name = "UnderlinePrintCheckBox";
            this.UnderlinePrintCheckBox.Size = new Size(140, 22);
            this.UnderlinePrintCheckBox.TabIndex = 13;
            this.UnderlinePrintCheckBox.Text = "下線印刷";
            //
            // Last4DigitsPrintCheckBox
            //
            this.Last4DigitsPrintCheckBox.Location = new Point(584, 24);
            this.Last4DigitsPrintCheckBox.Name = "Last4DigitsPrintCheckBox";
            this.Last4DigitsPrintCheckBox.Size = new Size(140, 22);
            this.Last4DigitsPrintCheckBox.TabIndex = 14;
            this.Last4DigitsPrintCheckBox.Text = "末尾4桁印刷";
            //
            // CheckSheetPrintCheckBox
            //
            this.CheckSheetPrintCheckBox.Location = new Point(8, 50);
            this.CheckSheetPrintCheckBox.Name = "CheckSheetPrintCheckBox";
            this.CheckSheetPrintCheckBox.Size = new Size(140, 22);
            this.CheckSheetPrintCheckBox.TabIndex = 15;
            this.CheckSheetPrintCheckBox.Text = "チェックシート印刷";
            //
            // ListPrintCheckBox
            //
            this.ListPrintCheckBox.Location = new Point(152, 50);
            this.ListPrintCheckBox.Name = "ListPrintCheckBox";
            this.ListPrintCheckBox.Size = new Size(140, 22);
            this.ListPrintCheckBox.TabIndex = 16;
            this.ListPrintCheckBox.Text = "リスト印刷";
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
            this.CheckBinGroupBox.Location = new Point(12, 465);
            this.CheckBinGroupBox.Name = "CheckBinGroupBox";
            this.CheckBinGroupBox.Size = new Size(736, 98);
            this.CheckBinGroupBox.TabIndex = 3;
            this.CheckBinGroupBox.TabStop = false;
            this.CheckBinGroupBox.Text = "登録フォーム表示項目";
            // 
            // CheckBinCheckBox0
            //
            this.CheckBinCheckBox0.Location = new Point(8, 24);
            this.CheckBinCheckBox0.Name = "CheckBinCheckBox0";
            this.CheckBinCheckBox0.Size = new Size(118, 22);
            this.CheckBinCheckBox0.TabIndex = 17;
            this.CheckBinCheckBox0.Text = "注文番号";
            //
            // CheckBinCheckBox1
            //
            this.CheckBinCheckBox1.Location = new Point(128, 24);
            this.CheckBinCheckBox1.Name = "CheckBinCheckBox1";
            this.CheckBinCheckBox1.Size = new Size(118, 22);
            this.CheckBinCheckBox1.TabIndex = 18;
            this.CheckBinCheckBox1.Text = "製造番号";
            //
            // CheckBinCheckBox2
            //
            this.CheckBinCheckBox2.Location = new Point(248, 24);
            this.CheckBinCheckBox2.Name = "CheckBinCheckBox2";
            this.CheckBinCheckBox2.Size = new Size(118, 22);
            this.CheckBinCheckBox2.TabIndex = 19;
            this.CheckBinCheckBox2.Text = "数量";
            //
            // CheckBinCheckBox3
            //
            this.CheckBinCheckBox3.Location = new Point(368, 24);
            this.CheckBinCheckBox3.Name = "CheckBinCheckBox3";
            this.CheckBinCheckBox3.Size = new Size(118, 22);
            this.CheckBinCheckBox3.TabIndex = 20;
            this.CheckBinCheckBox3.Text = "先頭シリアル番号";
            //
            // CheckBinCheckBox4
            //
            this.CheckBinCheckBox4.Location = new Point(488, 24);
            this.CheckBinCheckBox4.Name = "CheckBinCheckBox4";
            this.CheckBinCheckBox4.Size = new Size(118, 22);
            this.CheckBinCheckBox4.TabIndex = 21;
            this.CheckBinCheckBox4.Text = "リビジョン";
            //
            // CheckBinCheckBox5
            //
            this.CheckBinCheckBox5.Location = new Point(608, 24);
            this.CheckBinCheckBox5.Name = "CheckBinCheckBox5";
            this.CheckBinCheckBox5.Size = new Size(118, 22);
            this.CheckBinCheckBox5.TabIndex = 22;
            this.CheckBinCheckBox5.Text = "予備";
            //
            // CheckBinCheckBox6
            //
            this.CheckBinCheckBox6.Location = new Point(8, 50);
            this.CheckBinCheckBox6.Name = "CheckBinCheckBox6";
            this.CheckBinCheckBox6.Size = new Size(118, 22);
            this.CheckBinCheckBox6.TabIndex = 23;
            this.CheckBinCheckBox6.Text = "OLes番号";
            //
            // CheckBinCheckBox7
            //
            this.CheckBinCheckBox7.Location = new Point(128, 50);
            this.CheckBinCheckBox7.Name = "CheckBinCheckBox7";
            this.CheckBinCheckBox7.Size = new Size(118, 22);
            this.CheckBinCheckBox7.TabIndex = 24;
            this.CheckBinCheckBox7.Text = "予備";
            //
            // CheckBinCheckBox8
            //
            this.CheckBinCheckBox8.Location = new Point(248, 50);
            this.CheckBinCheckBox8.Name = "CheckBinCheckBox8";
            this.CheckBinCheckBox8.Size = new Size(118, 22);
            this.CheckBinCheckBox8.TabIndex = 25;
            this.CheckBinCheckBox8.Text = "登録日";
            //
            // CheckBinCheckBox9
            //
            this.CheckBinCheckBox9.Location = new Point(368, 50);
            this.CheckBinCheckBox9.Name = "CheckBinCheckBox9";
            this.CheckBinCheckBox9.Size = new Size(118, 22);
            this.CheckBinCheckBox9.TabIndex = 26;
            this.CheckBinCheckBox9.Text = "担当者";
            //
            // CheckBinCheckBox10
            //
            this.CheckBinCheckBox10.Location = new Point(488, 50);
            this.CheckBinCheckBox10.Name = "CheckBinCheckBox10";
            this.CheckBinCheckBox10.Size = new Size(118, 22);
            this.CheckBinCheckBox10.TabIndex = 27;
            this.CheckBinCheckBox10.Text = "コメント";
            // 
            // UseSubstrateGroupBox
            // 
            this.UseSubstrateGroupBox.Controls.Add(this.SubstrateProductNameFilterLabel);
            this.UseSubstrateGroupBox.Controls.Add(this.SubstrateProductNameFilterTextBox);
            this.UseSubstrateGroupBox.Controls.Add(this.SubstrateModelFilterLabel);
            this.UseSubstrateGroupBox.Controls.Add(this.SubstrateModelFilterTextBox);
            this.UseSubstrateGroupBox.Controls.Add(this.SubstrateCheckedListBox);
            this.UseSubstrateGroupBox.Location = new Point(12, 573);
            this.UseSubstrateGroupBox.Name = "UseSubstrateGroupBox";
            this.UseSubstrateGroupBox.Size = new Size(736, 250);
            this.UseSubstrateGroupBox.TabIndex = 4;
            this.UseSubstrateGroupBox.TabStop = false;
            this.UseSubstrateGroupBox.Text = "使用基板（M_ProductUseSubstrate）";
            //
            // SubstrateProductNameFilterLabel
            //
            this.SubstrateProductNameFilterLabel.AutoSize = true;
            this.SubstrateProductNameFilterLabel.Location = new Point(8, 27);
            this.SubstrateProductNameFilterLabel.Name = "SubstrateProductNameFilterLabel";
            this.SubstrateProductNameFilterLabel.Text = "製品名:";
            //
            // SubstrateProductNameFilterTextBox
            //
            this.SubstrateProductNameFilterTextBox.Location = new Point(56, 23);
            this.SubstrateProductNameFilterTextBox.Name = "SubstrateProductNameFilterTextBox";
            this.SubstrateProductNameFilterTextBox.Size = new Size(200, 23);
            this.SubstrateProductNameFilterTextBox.TabIndex = 0;
            this.SubstrateProductNameFilterTextBox.TextChanged += this.SubstrateProductNameFilterTextBox_TextChanged;
            //
            // SubstrateModelFilterLabel
            //
            this.SubstrateModelFilterLabel.AutoSize = true;
            this.SubstrateModelFilterLabel.Location = new Point(268, 27);
            this.SubstrateModelFilterLabel.Name = "SubstrateModelFilterLabel";
            this.SubstrateModelFilterLabel.Text = "基板型式:";
            //
            // SubstrateModelFilterTextBox
            //
            this.SubstrateModelFilterTextBox.Location = new Point(332, 23);
            this.SubstrateModelFilterTextBox.Name = "SubstrateModelFilterTextBox";
            this.SubstrateModelFilterTextBox.Size = new Size(200, 23);
            this.SubstrateModelFilterTextBox.TabIndex = 1;
            this.SubstrateModelFilterTextBox.TextChanged += this.SubstrateModelFilterTextBox_TextChanged;
            //
            // SubstrateCheckedListBox
            //
            this.SubstrateCheckedListBox.CheckOnClick = true;
            this.SubstrateCheckedListBox.Location = new Point(8, 77);
            this.SubstrateCheckedListBox.Name = "SubstrateCheckedListBox";
            this.SubstrateCheckedListBox.Size = new Size(718, 161);
            this.SubstrateCheckedListBox.TabIndex = 28;
            this.SubstrateCheckedListBox.ItemCheck += this.SubstrateCheckedListBox_ItemCheck;
            // 
            // SaveButton
            // 
            this.SaveButton.Location = new Point(538, 835);
            this.SaveButton.Name = "SaveButton";
            this.SaveButton.Size = new Size(100, 35);
            this.SaveButton.TabIndex = 29;
            this.SaveButton.Text = "保存";
            this.SaveButton.UseVisualStyleBackColor = true;
            this.SaveButton.Click += this.SaveButton_Click;
            // 
            // DialogCancelButton
            // 
            this.DialogCancelButton.Location = new Point(648, 835);
            this.DialogCancelButton.Name = "DialogCancelButton";
            this.DialogCancelButton.Size = new Size(100, 35);
            this.DialogCancelButton.TabIndex = 30;
            this.DialogCancelButton.Text = "キャンセル";
            this.DialogCancelButton.UseVisualStyleBackColor = true;
            this.DialogCancelButton.Click += this.CancelButton_Click;
            // 
            // ProductMasterEditDialog
            // 
            this.AutoScaleDimensions = new SizeF(7F, 15F);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.ClientSize = new Size(760, 882);
            this.Controls.Add(this.BasicInfoGroupBox);
            this.Controls.Add(this.SerialGroupBox);
            this.Controls.Add(this.PrintGroupBox);
            this.Controls.Add(this.CheckBinGroupBox);
            this.Controls.Add(this.UseSubstrateGroupBox);
            this.Controls.Add(this.SaveButton);
            this.Controls.Add(this.DialogCancelButton);
            this.Font = new Font("Meiryo UI", 9F);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ProductMasterEditDialog";
            this.StartPosition = FormStartPosition.CenterParent;
            this.Text = "製品マスター編集";
            this.Load += this.ProductMasterEditDialog_Load;
            this.BasicInfoGroupBox.ResumeLayout(false);
            this.BasicInfoGroupBox.PerformLayout();
            this.SerialGroupBox.ResumeLayout(false);
            this.SerialGroupBox.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)this.RevisionGroupNumericUpDown).EndInit();
            this.PrintGroupBox.ResumeLayout(false);
            this.CheckBinGroupBox.ResumeLayout(false);
            this.UseSubstrateGroupBox.ResumeLayout(false);
            this.ResumeLayout(false);
        }

        #endregion

        // セクション1
        private GroupBox        BasicInfoGroupBox;
        private Label           CategoryNameLabel;
        private TextBox         CategoryNameTextBox;
        private Label           ProductNameLabel;
        private TextBox         ProductNameTextBox;
        private Label           ProductModelLabel;
        private TextBox         ProductModelTextBox;
        private Label           ProductTypeLabel;
        private TextBox         ProductTypeTextBox;
        private Label           InitialLabel;
        private TextBox         InitialTextBox;
        private Label           VisibleLabel;
        private CheckBox        VisibleCheckBox;

        // セクション2
        private GroupBox        SerialGroupBox;
        private Label           RegTypeLabel;
        private ComboBox        RegTypeComboBox;
        private Label           SerialTypeLabel;
        private ComboBox        SerialTypeComboBox;
        private Label           RevisionGroupLabel;
        private NumericUpDown   RevisionGroupNumericUpDown;

        // セクション3
        private GroupBox        PrintGroupBox;
        private CheckBox        LabelPrintCheckBox;
        private CheckBox        BarcodePrintCheckBox;
        private CheckBox        NameplatePrintCheckBox;
        private CheckBox        UnderlinePrintCheckBox;
        private CheckBox        Last4DigitsPrintCheckBox;
        private CheckBox        CheckSheetPrintCheckBox;
        private CheckBox        ListPrintCheckBox;

        // セクション4
        private GroupBox        CheckBinGroupBox;
        private CheckBox        CheckBinCheckBox0;
        private CheckBox        CheckBinCheckBox1;
        private CheckBox        CheckBinCheckBox2;
        private CheckBox        CheckBinCheckBox3;
        private CheckBox        CheckBinCheckBox4;
        private CheckBox        CheckBinCheckBox5;
        private CheckBox        CheckBinCheckBox6;
        private CheckBox        CheckBinCheckBox7;
        private CheckBox        CheckBinCheckBox8;
        private CheckBox        CheckBinCheckBox9;
        private CheckBox        CheckBinCheckBox10;

        // セクション5
        private GroupBox        UseSubstrateGroupBox;
        private Label           SubstrateProductNameFilterLabel;
        private TextBox         SubstrateProductNameFilterTextBox;
        private Label           SubstrateModelFilterLabel;
        private TextBox         SubstrateModelFilterTextBox;
        private CheckedListBox  SubstrateCheckedListBox;

        // ボタン
        private Button          SaveButton;
        private Button          DialogCancelButton;
    }
}
