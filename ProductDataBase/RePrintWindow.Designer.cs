namespace ProductDataBase {
    partial class RePrintWindow {
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(RePrintWindow));
            TemplateButton = new Button();
            CommentComboBox = new ComboBox();
            CommentTextBox = new TextBox();
            CommentCheckBox = new CheckBox();
            ExtraTextBox6 = new TextBox();
            ExtraCheckBox6 = new CheckBox();
            ExtraTextBox5 = new TextBox();
            ExtraCheckBox5 = new CheckBox();
            ExtraTextBox4 = new TextBox();
            ExtraCheckBox4 = new CheckBox();
            PersonComboBox = new ComboBox();
            PersonCheckBox = new CheckBox();
            RegistrationDateMaskedTextBox = new MaskedTextBox();
            RegistrationDateCheckBox = new CheckBox();
            FirstSerialNumberTextBox = new TextBox();
            FirstSerialNumberCheckBox = new CheckBox();
            ExtraTextBox3 = new TextBox();
            ExtraCheckBox3 = new CheckBox();
            ExtraTextBox2 = new TextBox();
            ExtraCheckBox2 = new CheckBox();
            RevisionTextBox = new TextBox();
            RevisionCheckBox = new CheckBox();
            ExtraTextBox1 = new TextBox();
            ExtraCheckBox1 = new CheckBox();
            QuantityTextBox = new TextBox();
            QuantityCheckBox = new CheckBox();
            ManufacturingNumberMaskedTextBox = new MaskedTextBox();
            ManufacturingNumberCheckBox = new CheckBox();
            OrderNumberTextBox = new TextBox();
            OrderNumberCheckBox = new CheckBox();
            SubstrateModelLabel2 = new Label();
            SubstrateModelLabel1 = new Label();
            ProductNameLabel2 = new Label();
            ProductNameLabel1 = new Label();
            RePrintMenuStrip = new MenuStrip();
            ファイルToolStripMenuItem = new ToolStripMenuItem();
            終了ToolStripMenuItem = new ToolStripMenuItem();
            ヘルプToolStripMenuItem = new ToolStripMenuItem();
            取得情報ToolStripMenuItem = new ToolStripMenuItem();
            BarcodePrintButton = new Button();
            LabelPrintButton = new Button();
            PrintPostionNumericUpDown = new NumericUpDown();
            PrintRowLabel = new Label();
            RePrintPrintDialog = new PrintDialog();
            RePrintPrintDocument = new System.Drawing.Printing.PrintDocument();
            RePrintPrintPreviewDialog = new PrintPreviewDialog();
            RePrintMenuStrip.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)PrintPostionNumericUpDown).BeginInit();
            SuspendLayout();
            // 
            // TemplateButton
            // 
            TemplateButton.Enabled = false;
            TemplateButton.Location = new Point(539, 349);
            TemplateButton.Name = "TemplateButton";
            TemplateButton.Size = new Size(75, 25);
            TemplateButton.TabIndex = 116;
            TemplateButton.Text = "定型文";
            TemplateButton.UseVisualStyleBackColor = true;
            // 
            // CommentComboBox
            // 
            CommentComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            CommentComboBox.DropDownWidth = 150;
            CommentComboBox.Enabled = false;
            CommentComboBox.FormattingEnabled = true;
            CommentComboBox.Location = new Point(464, 322);
            CommentComboBox.Name = "CommentComboBox";
            CommentComboBox.Size = new Size(150, 23);
            CommentComboBox.TabIndex = 115;
            // 
            // CommentTextBox
            // 
            CommentTextBox.Enabled = false;
            CommentTextBox.Location = new Point(464, 143);
            CommentTextBox.Margin = new Padding(0);
            CommentTextBox.MaxLength = 500;
            CommentTextBox.Multiline = true;
            CommentTextBox.Name = "CommentTextBox";
            CommentTextBox.ScrollBars = ScrollBars.Vertical;
            CommentTextBox.ShortcutsEnabled = false;
            CommentTextBox.Size = new Size(150, 175);
            CommentTextBox.TabIndex = 114;
            // 
            // CommentCheckBox
            // 
            CommentCheckBox.Location = new Point(464, 120);
            CommentCheckBox.Margin = new Padding(0);
            CommentCheckBox.Name = "CommentCheckBox";
            CommentCheckBox.Size = new Size(59, 19);
            CommentCheckBox.TabIndex = 113;
            CommentCheckBox.TabStop = false;
            CommentCheckBox.Text = "コメント";
            CommentCheckBox.UseVisualStyleBackColor = true;
            CommentCheckBox.CheckedChanged += CheckBoxChecked;
            // 
            // ExtraTextBox6
            // 
            ExtraTextBox6.Enabled = false;
            ExtraTextBox6.Location = new Point(313, 243);
            ExtraTextBox6.Margin = new Padding(0);
            ExtraTextBox6.MaxLength = 20;
            ExtraTextBox6.Name = "ExtraTextBox6";
            ExtraTextBox6.ShortcutsEnabled = false;
            ExtraTextBox6.Size = new Size(120, 23);
            ExtraTextBox6.TabIndex = 112;
            ExtraTextBox6.TextAlign = HorizontalAlignment.Right;
            // 
            // ExtraCheckBox6
            // 
            ExtraCheckBox6.Location = new Point(313, 220);
            ExtraCheckBox6.Margin = new Padding(0);
            ExtraCheckBox6.Name = "ExtraCheckBox6";
            ExtraCheckBox6.Size = new Size(50, 19);
            ExtraCheckBox6.TabIndex = 111;
            ExtraCheckBox6.TabStop = false;
            ExtraCheckBox6.Text = "予備";
            ExtraCheckBox6.UseVisualStyleBackColor = true;
            ExtraCheckBox6.CheckedChanged += CheckBoxChecked;
            // 
            // ExtraTextBox5
            // 
            ExtraTextBox5.Enabled = false;
            ExtraTextBox5.Location = new Point(313, 193);
            ExtraTextBox5.Margin = new Padding(0);
            ExtraTextBox5.MaxLength = 20;
            ExtraTextBox5.Name = "ExtraTextBox5";
            ExtraTextBox5.ShortcutsEnabled = false;
            ExtraTextBox5.Size = new Size(120, 23);
            ExtraTextBox5.TabIndex = 110;
            ExtraTextBox5.TextAlign = HorizontalAlignment.Right;
            // 
            // ExtraCheckBox5
            // 
            ExtraCheckBox5.Location = new Point(313, 170);
            ExtraCheckBox5.Margin = new Padding(0);
            ExtraCheckBox5.Name = "ExtraCheckBox5";
            ExtraCheckBox5.Size = new Size(50, 19);
            ExtraCheckBox5.TabIndex = 109;
            ExtraCheckBox5.TabStop = false;
            ExtraCheckBox5.Text = "予備";
            ExtraCheckBox5.UseVisualStyleBackColor = true;
            ExtraCheckBox5.CheckedChanged += CheckBoxChecked;
            // 
            // ExtraTextBox4
            // 
            ExtraTextBox4.Enabled = false;
            ExtraTextBox4.Location = new Point(313, 143);
            ExtraTextBox4.Margin = new Padding(0);
            ExtraTextBox4.MaxLength = 20;
            ExtraTextBox4.Name = "ExtraTextBox4";
            ExtraTextBox4.ShortcutsEnabled = false;
            ExtraTextBox4.Size = new Size(120, 23);
            ExtraTextBox4.TabIndex = 108;
            ExtraTextBox4.TextAlign = HorizontalAlignment.Right;
            // 
            // ExtraCheckBox4
            // 
            ExtraCheckBox4.Location = new Point(313, 120);
            ExtraCheckBox4.Margin = new Padding(0);
            ExtraCheckBox4.Name = "ExtraCheckBox4";
            ExtraCheckBox4.Size = new Size(50, 19);
            ExtraCheckBox4.TabIndex = 107;
            ExtraCheckBox4.TabStop = false;
            ExtraCheckBox4.Text = "予備";
            ExtraCheckBox4.UseVisualStyleBackColor = true;
            ExtraCheckBox4.CheckedChanged += CheckBoxChecked;
            // 
            // PersonComboBox
            // 
            PersonComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            PersonComboBox.Enabled = false;
            PersonComboBox.FormattingEnabled = true;
            PersonComboBox.Location = new Point(166, 343);
            PersonComboBox.Name = "PersonComboBox";
            PersonComboBox.Size = new Size(121, 23);
            PersonComboBox.TabIndex = 106;
            // 
            // PersonCheckBox
            // 
            PersonCheckBox.Location = new Point(166, 320);
            PersonCheckBox.Margin = new Padding(0);
            PersonCheckBox.Name = "PersonCheckBox";
            PersonCheckBox.Size = new Size(62, 19);
            PersonCheckBox.TabIndex = 105;
            PersonCheckBox.TabStop = false;
            PersonCheckBox.Text = "担当者";
            PersonCheckBox.UseVisualStyleBackColor = true;
            PersonCheckBox.CheckedChanged += CheckBoxChecked;
            // 
            // RegistrationDateMaskedTextBox
            // 
            RegistrationDateMaskedTextBox.Location = new Point(166, 293);
            RegistrationDateMaskedTextBox.Mask = "0000/00/00";
            RegistrationDateMaskedTextBox.Name = "RegistrationDateMaskedTextBox";
            RegistrationDateMaskedTextBox.PromptChar = '*';
            RegistrationDateMaskedTextBox.Size = new Size(120, 23);
            RegistrationDateMaskedTextBox.TabIndex = 104;
            RegistrationDateMaskedTextBox.TextAlign = HorizontalAlignment.Right;
            RegistrationDateMaskedTextBox.TypeValidationCompleted += RegistrationDateCheck;
            // 
            // RegistrationDateCheckBox
            // 
            RegistrationDateCheckBox.Location = new Point(166, 270);
            RegistrationDateCheckBox.Margin = new Padding(0);
            RegistrationDateCheckBox.Name = "RegistrationDateCheckBox";
            RegistrationDateCheckBox.Size = new Size(62, 19);
            RegistrationDateCheckBox.TabIndex = 103;
            RegistrationDateCheckBox.TabStop = false;
            RegistrationDateCheckBox.Text = "登録日";
            RegistrationDateCheckBox.UseVisualStyleBackColor = true;
            RegistrationDateCheckBox.CheckedChanged += CheckBoxChecked;
            // 
            // FirstSerialNumberTextBox
            // 
            FirstSerialNumberTextBox.Enabled = false;
            FirstSerialNumberTextBox.Location = new Point(166, 243);
            FirstSerialNumberTextBox.Margin = new Padding(0);
            FirstSerialNumberTextBox.MaxLength = 20;
            FirstSerialNumberTextBox.Name = "FirstSerialNumberTextBox";
            FirstSerialNumberTextBox.ShortcutsEnabled = false;
            FirstSerialNumberTextBox.Size = new Size(120, 23);
            FirstSerialNumberTextBox.TabIndex = 102;
            FirstSerialNumberTextBox.TextAlign = HorizontalAlignment.Right;
            // 
            // FirstSerialNumberCheckBox
            // 
            FirstSerialNumberCheckBox.Location = new Point(166, 220);
            FirstSerialNumberCheckBox.Margin = new Padding(0);
            FirstSerialNumberCheckBox.Name = "FirstSerialNumberCheckBox";
            FirstSerialNumberCheckBox.Size = new Size(110, 19);
            FirstSerialNumberCheckBox.TabIndex = 101;
            FirstSerialNumberCheckBox.TabStop = false;
            FirstSerialNumberCheckBox.Text = "先頭シリアル番号";
            FirstSerialNumberCheckBox.UseVisualStyleBackColor = true;
            FirstSerialNumberCheckBox.CheckedChanged += CheckBoxChecked;
            // 
            // ExtraTextBox3
            // 
            ExtraTextBox3.Enabled = false;
            ExtraTextBox3.Location = new Point(166, 193);
            ExtraTextBox3.Margin = new Padding(0);
            ExtraTextBox3.MaxLength = 20;
            ExtraTextBox3.Name = "ExtraTextBox3";
            ExtraTextBox3.ShortcutsEnabled = false;
            ExtraTextBox3.Size = new Size(120, 23);
            ExtraTextBox3.TabIndex = 100;
            ExtraTextBox3.TextAlign = HorizontalAlignment.Right;
            // 
            // ExtraCheckBox3
            // 
            ExtraCheckBox3.Location = new Point(166, 170);
            ExtraCheckBox3.Margin = new Padding(0);
            ExtraCheckBox3.Name = "ExtraCheckBox3";
            ExtraCheckBox3.Size = new Size(50, 19);
            ExtraCheckBox3.TabIndex = 99;
            ExtraCheckBox3.TabStop = false;
            ExtraCheckBox3.Text = "予備";
            ExtraCheckBox3.UseVisualStyleBackColor = true;
            ExtraCheckBox3.CheckedChanged += CheckBoxChecked;
            // 
            // ExtraTextBox2
            // 
            ExtraTextBox2.Enabled = false;
            ExtraTextBox2.Location = new Point(166, 143);
            ExtraTextBox2.Margin = new Padding(0);
            ExtraTextBox2.MaxLength = 20;
            ExtraTextBox2.Name = "ExtraTextBox2";
            ExtraTextBox2.ShortcutsEnabled = false;
            ExtraTextBox2.Size = new Size(120, 23);
            ExtraTextBox2.TabIndex = 98;
            ExtraTextBox2.TextAlign = HorizontalAlignment.Right;
            // 
            // ExtraCheckBox2
            // 
            ExtraCheckBox2.Location = new Point(166, 120);
            ExtraCheckBox2.Margin = new Padding(0);
            ExtraCheckBox2.Name = "ExtraCheckBox2";
            ExtraCheckBox2.Size = new Size(50, 19);
            ExtraCheckBox2.TabIndex = 97;
            ExtraCheckBox2.TabStop = false;
            ExtraCheckBox2.Text = "予備";
            ExtraCheckBox2.UseVisualStyleBackColor = true;
            ExtraCheckBox2.CheckedChanged += CheckBoxChecked;
            // 
            // RevisionTextBox
            // 
            RevisionTextBox.Enabled = false;
            RevisionTextBox.Location = new Point(19, 343);
            RevisionTextBox.Margin = new Padding(0);
            RevisionTextBox.MaxLength = 4;
            RevisionTextBox.Name = "RevisionTextBox";
            RevisionTextBox.ShortcutsEnabled = false;
            RevisionTextBox.Size = new Size(120, 23);
            RevisionTextBox.TabIndex = 96;
            RevisionTextBox.TextAlign = HorizontalAlignment.Right;
            // 
            // RevisionCheckBox
            // 
            RevisionCheckBox.Location = new Point(19, 320);
            RevisionCheckBox.Margin = new Padding(0);
            RevisionCheckBox.Name = "RevisionCheckBox";
            RevisionCheckBox.Size = new Size(68, 19);
            RevisionCheckBox.TabIndex = 95;
            RevisionCheckBox.TabStop = false;
            RevisionCheckBox.Text = "レビジョン";
            RevisionCheckBox.UseVisualStyleBackColor = true;
            RevisionCheckBox.CheckedChanged += CheckBoxChecked;
            // 
            // ExtraTextBox1
            // 
            ExtraTextBox1.Enabled = false;
            ExtraTextBox1.Location = new Point(19, 293);
            ExtraTextBox1.Margin = new Padding(0);
            ExtraTextBox1.MaxLength = 4;
            ExtraTextBox1.Name = "ExtraTextBox1";
            ExtraTextBox1.ShortcutsEnabled = false;
            ExtraTextBox1.Size = new Size(120, 23);
            ExtraTextBox1.TabIndex = 94;
            ExtraTextBox1.TextAlign = HorizontalAlignment.Right;
            // 
            // ExtraCheckBox1
            // 
            ExtraCheckBox1.Location = new Point(19, 270);
            ExtraCheckBox1.Margin = new Padding(0);
            ExtraCheckBox1.Name = "ExtraCheckBox1";
            ExtraCheckBox1.Size = new Size(62, 19);
            ExtraCheckBox1.TabIndex = 93;
            ExtraCheckBox1.TabStop = false;
            ExtraCheckBox1.Text = "予備";
            ExtraCheckBox1.UseVisualStyleBackColor = true;
            ExtraCheckBox1.CheckedChanged += CheckBoxChecked;
            // 
            // QuantityTextBox
            // 
            QuantityTextBox.Enabled = false;
            QuantityTextBox.Location = new Point(19, 243);
            QuantityTextBox.Margin = new Padding(0);
            QuantityTextBox.MaxLength = 4;
            QuantityTextBox.Name = "QuantityTextBox";
            QuantityTextBox.ShortcutsEnabled = false;
            QuantityTextBox.Size = new Size(120, 23);
            QuantityTextBox.TabIndex = 92;
            QuantityTextBox.TextAlign = HorizontalAlignment.Right;
            QuantityTextBox.KeyPress += NumericOnly;
            // 
            // QuantityCheckBox
            // 
            QuantityCheckBox.Location = new Point(19, 220);
            QuantityCheckBox.Margin = new Padding(0);
            QuantityCheckBox.Name = "QuantityCheckBox";
            QuantityCheckBox.Size = new Size(50, 19);
            QuantityCheckBox.TabIndex = 91;
            QuantityCheckBox.TabStop = false;
            QuantityCheckBox.Text = "数量";
            QuantityCheckBox.UseVisualStyleBackColor = true;
            QuantityCheckBox.CheckedChanged += CheckBoxChecked;
            // 
            // ManufacturingNumberMaskedTextBox
            // 
            ManufacturingNumberMaskedTextBox.Location = new Point(19, 193);
            ManufacturingNumberMaskedTextBox.Mask = "LA00L00000-0000";
            ManufacturingNumberMaskedTextBox.Name = "ManufacturingNumberMaskedTextBox";
            ManufacturingNumberMaskedTextBox.PromptChar = '*';
            ManufacturingNumberMaskedTextBox.Size = new Size(120, 23);
            ManufacturingNumberMaskedTextBox.TabIndex = 90;
            ManufacturingNumberMaskedTextBox.Text = "H";
            ManufacturingNumberMaskedTextBox.TextAlign = HorizontalAlignment.Right;
            // 
            // ManufacturingNumberCheckBox
            // 
            ManufacturingNumberCheckBox.Location = new Point(19, 170);
            ManufacturingNumberCheckBox.Margin = new Padding(0);
            ManufacturingNumberCheckBox.Name = "ManufacturingNumberCheckBox";
            ManufacturingNumberCheckBox.Size = new Size(74, 19);
            ManufacturingNumberCheckBox.TabIndex = 89;
            ManufacturingNumberCheckBox.TabStop = false;
            ManufacturingNumberCheckBox.Text = "製造番号";
            ManufacturingNumberCheckBox.UseVisualStyleBackColor = true;
            ManufacturingNumberCheckBox.CheckedChanged += CheckBoxChecked;
            // 
            // OrderNumberTextBox
            // 
            OrderNumberTextBox.Enabled = false;
            OrderNumberTextBox.Location = new Point(19, 143);
            OrderNumberTextBox.Margin = new Padding(0);
            OrderNumberTextBox.MaxLength = 20;
            OrderNumberTextBox.Name = "OrderNumberTextBox";
            OrderNumberTextBox.ShortcutsEnabled = false;
            OrderNumberTextBox.Size = new Size(120, 23);
            OrderNumberTextBox.TabIndex = 88;
            OrderNumberTextBox.TextAlign = HorizontalAlignment.Right;
            // 
            // OrderNumberCheckBox
            // 
            OrderNumberCheckBox.Location = new Point(19, 120);
            OrderNumberCheckBox.Margin = new Padding(0);
            OrderNumberCheckBox.Name = "OrderNumberCheckBox";
            OrderNumberCheckBox.Size = new Size(74, 19);
            OrderNumberCheckBox.TabIndex = 87;
            OrderNumberCheckBox.TabStop = false;
            OrderNumberCheckBox.Text = "注文番号";
            OrderNumberCheckBox.UseVisualStyleBackColor = true;
            OrderNumberCheckBox.CheckedChanged += CheckBoxChecked;
            // 
            // SubstrateModelLabel2
            // 
            SubstrateModelLabel2.AutoSize = true;
            SubstrateModelLabel2.BorderStyle = BorderStyle.Fixed3D;
            SubstrateModelLabel2.Location = new Point(84, 67);
            SubstrateModelLabel2.Margin = new Padding(4, 0, 4, 0);
            SubstrateModelLabel2.Name = "SubstrateModelLabel2";
            SubstrateModelLabel2.Size = new Size(24, 17);
            SubstrateModelLabel2.TabIndex = 86;
            SubstrateModelLabel2.Text = "---";
            SubstrateModelLabel2.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // SubstrateModelLabel1
            // 
            SubstrateModelLabel1.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            SubstrateModelLabel1.Location = new Point(16, 67);
            SubstrateModelLabel1.Margin = new Padding(4, 0, 4, 0);
            SubstrateModelLabel1.Name = "SubstrateModelLabel1";
            SubstrateModelLabel1.Size = new Size(60, 15);
            SubstrateModelLabel1.TabIndex = 85;
            SubstrateModelLabel1.Text = "基盤型式:";
            SubstrateModelLabel1.TextAlign = ContentAlignment.MiddleRight;
            // 
            // ProductNameLabel2
            // 
            ProductNameLabel2.AutoSize = true;
            ProductNameLabel2.BorderStyle = BorderStyle.Fixed3D;
            ProductNameLabel2.Location = new Point(84, 45);
            ProductNameLabel2.Margin = new Padding(4, 0, 4, 0);
            ProductNameLabel2.Name = "ProductNameLabel2";
            ProductNameLabel2.Size = new Size(24, 17);
            ProductNameLabel2.TabIndex = 84;
            ProductNameLabel2.Text = "---";
            ProductNameLabel2.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // ProductNameLabel1
            // 
            ProductNameLabel1.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            ProductNameLabel1.Location = new Point(28, 45);
            ProductNameLabel1.Margin = new Padding(4, 0, 4, 0);
            ProductNameLabel1.Name = "ProductNameLabel1";
            ProductNameLabel1.Size = new Size(48, 15);
            ProductNameLabel1.TabIndex = 83;
            ProductNameLabel1.Text = "製品名:";
            ProductNameLabel1.TextAlign = ContentAlignment.MiddleRight;
            // 
            // RePrintMenuStrip
            // 
            RePrintMenuStrip.Font = new Font("Meiryo UI", 9F, FontStyle.Regular, GraphicsUnit.Point);
            RePrintMenuStrip.Items.AddRange(new ToolStripItem[] { ファイルToolStripMenuItem, ヘルプToolStripMenuItem });
            RePrintMenuStrip.Location = new Point(0, 0);
            RePrintMenuStrip.Name = "RePrintMenuStrip";
            RePrintMenuStrip.Size = new Size(630, 24);
            RePrintMenuStrip.TabIndex = 82;
            RePrintMenuStrip.Text = "menuStrip";
            // 
            // ファイルToolStripMenuItem
            // 
            ファイルToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { 終了ToolStripMenuItem });
            ファイルToolStripMenuItem.Name = "ファイルToolStripMenuItem";
            ファイルToolStripMenuItem.Size = new Size(53, 20);
            ファイルToolStripMenuItem.Text = "ファイル";
            // 
            // 終了ToolStripMenuItem
            // 
            終了ToolStripMenuItem.Name = "終了ToolStripMenuItem";
            終了ToolStripMenuItem.Size = new Size(98, 22);
            終了ToolStripMenuItem.Text = "終了";
            // 
            // ヘルプToolStripMenuItem
            // 
            ヘルプToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { 取得情報ToolStripMenuItem });
            ヘルプToolStripMenuItem.Name = "ヘルプToolStripMenuItem";
            ヘルプToolStripMenuItem.Size = new Size(48, 20);
            ヘルプToolStripMenuItem.Text = "ヘルプ";
            // 
            // 取得情報ToolStripMenuItem
            // 
            取得情報ToolStripMenuItem.Name = "取得情報ToolStripMenuItem";
            取得情報ToolStripMenuItem.Size = new Size(122, 22);
            取得情報ToolStripMenuItem.Text = "取得情報";
            取得情報ToolStripMenuItem.Click += 取得情報ToolStripMenuItem_Click;
            // 
            // BarcodePrintButton
            // 
            BarcodePrintButton.Enabled = false;
            BarcodePrintButton.Location = new Point(450, 404);
            BarcodePrintButton.Margin = new Padding(0);
            BarcodePrintButton.Name = "BarcodePrintButton";
            BarcodePrintButton.Size = new Size(100, 25);
            BarcodePrintButton.TabIndex = 121;
            BarcodePrintButton.Text = "バーコード印刷";
            BarcodePrintButton.UseVisualStyleBackColor = true;
            BarcodePrintButton.Click += BarcodePrintButton_Click;
            // 
            // LabelPrintButton
            // 
            LabelPrintButton.Location = new Point(313, 404);
            LabelPrintButton.Margin = new Padding(0);
            LabelPrintButton.Name = "LabelPrintButton";
            LabelPrintButton.Size = new Size(100, 25);
            LabelPrintButton.TabIndex = 119;
            LabelPrintButton.Text = "ラベル印刷";
            LabelPrintButton.UseVisualStyleBackColor = true;
            LabelPrintButton.Click += LabelPrintButton_Click;
            // 
            // PrintPostionNumericUpDown
            // 
            PrintPostionNumericUpDown.Location = new Point(166, 406);
            PrintPostionNumericUpDown.Margin = new Padding(0);
            PrintPostionNumericUpDown.Maximum = new decimal(new int[] { 50, 0, 0, 0 });
            PrintPostionNumericUpDown.Minimum = new decimal(new int[] { 1, 0, 0, 0 });
            PrintPostionNumericUpDown.Name = "PrintPostionNumericUpDown";
            PrintPostionNumericUpDown.Size = new Size(110, 23);
            PrintPostionNumericUpDown.TabIndex = 118;
            PrintPostionNumericUpDown.TextAlign = HorizontalAlignment.Right;
            PrintPostionNumericUpDown.Value = new decimal(new int[] { 1, 0, 0, 0 });
            // 
            // PrintRowLabel
            // 
            PrintRowLabel.Location = new Point(178, 391);
            PrintRowLabel.Margin = new Padding(0);
            PrintRowLabel.Name = "PrintRowLabel";
            PrintRowLabel.Size = new Size(79, 15);
            PrintRowLabel.TabIndex = 117;
            PrintRowLabel.Text = "印刷開始位置";
            // 
            // RePrintPrintDialog
            // 
            RePrintPrintDialog.UseEXDialog = true;
            // 
            // RePrintPrintPreviewDialog
            // 
            RePrintPrintPreviewDialog.AutoScrollMargin = new Size(0, 0);
            RePrintPrintPreviewDialog.AutoScrollMinSize = new Size(0, 0);
            RePrintPrintPreviewDialog.ClientSize = new Size(400, 300);
            RePrintPrintPreviewDialog.Enabled = true;
            RePrintPrintPreviewDialog.Icon = (Icon)resources.GetObject("RePrintPrintPreviewDialog.Icon");
            RePrintPrintPreviewDialog.Name = "RePrintPrintPreviewDialog";
            RePrintPrintPreviewDialog.Visible = false;
            // 
            // RePrintWindow
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(630, 441);
            Controls.Add(BarcodePrintButton);
            Controls.Add(LabelPrintButton);
            Controls.Add(PrintPostionNumericUpDown);
            Controls.Add(PrintRowLabel);
            Controls.Add(TemplateButton);
            Controls.Add(CommentComboBox);
            Controls.Add(CommentTextBox);
            Controls.Add(CommentCheckBox);
            Controls.Add(ExtraTextBox6);
            Controls.Add(ExtraCheckBox6);
            Controls.Add(ExtraTextBox5);
            Controls.Add(ExtraCheckBox5);
            Controls.Add(ExtraTextBox4);
            Controls.Add(ExtraCheckBox4);
            Controls.Add(PersonComboBox);
            Controls.Add(PersonCheckBox);
            Controls.Add(RegistrationDateMaskedTextBox);
            Controls.Add(RegistrationDateCheckBox);
            Controls.Add(FirstSerialNumberTextBox);
            Controls.Add(FirstSerialNumberCheckBox);
            Controls.Add(ExtraTextBox3);
            Controls.Add(ExtraCheckBox3);
            Controls.Add(ExtraTextBox2);
            Controls.Add(ExtraCheckBox2);
            Controls.Add(RevisionTextBox);
            Controls.Add(RevisionCheckBox);
            Controls.Add(ExtraTextBox1);
            Controls.Add(ExtraCheckBox1);
            Controls.Add(QuantityTextBox);
            Controls.Add(QuantityCheckBox);
            Controls.Add(ManufacturingNumberMaskedTextBox);
            Controls.Add(ManufacturingNumberCheckBox);
            Controls.Add(OrderNumberTextBox);
            Controls.Add(OrderNumberCheckBox);
            Controls.Add(SubstrateModelLabel2);
            Controls.Add(SubstrateModelLabel1);
            Controls.Add(ProductNameLabel2);
            Controls.Add(ProductNameLabel1);
            Controls.Add(RePrintMenuStrip);
            Font = new Font("Meiryo UI", 9F, FontStyle.Regular, GraphicsUnit.Point);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "RePrintWindow";
            ShowIcon = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterParent;
            Text = "RePrint";
            Load += RePrintWindow_Load;
            RePrintMenuStrip.ResumeLayout(false);
            RePrintMenuStrip.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)PrintPostionNumericUpDown).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button TemplateButton;
        private ComboBox CommentComboBox;
        private TextBox CommentTextBox;
        private CheckBox CommentCheckBox;
        private TextBox ExtraTextBox6;
        private CheckBox ExtraCheckBox6;
        private TextBox ExtraTextBox5;
        private CheckBox ExtraCheckBox5;
        private TextBox ExtraTextBox4;
        private CheckBox ExtraCheckBox4;
        private ComboBox PersonComboBox;
        private CheckBox PersonCheckBox;
        private MaskedTextBox RegistrationDateMaskedTextBox;
        private CheckBox RegistrationDateCheckBox;
        private TextBox FirstSerialNumberTextBox;
        private CheckBox FirstSerialNumberCheckBox;
        private TextBox ExtraTextBox3;
        private CheckBox ExtraCheckBox3;
        private TextBox ExtraTextBox2;
        private CheckBox ExtraCheckBox2;
        private TextBox RevisionTextBox;
        private CheckBox RevisionCheckBox;
        private TextBox ExtraTextBox1;
        private CheckBox ExtraCheckBox1;
        private TextBox QuantityTextBox;
        private CheckBox QuantityCheckBox;
        private MaskedTextBox ManufacturingNumberMaskedTextBox;
        private CheckBox ManufacturingNumberCheckBox;
        private TextBox OrderNumberTextBox;
        private CheckBox OrderNumberCheckBox;
        private Label SubstrateModelLabel2;
        private Label SubstrateModelLabel1;
        private Label ProductNameLabel2;
        private Label ProductNameLabel1;
        private MenuStrip RePrintMenuStrip;
        private ToolStripMenuItem ファイルToolStripMenuItem;
        private ToolStripMenuItem 終了ToolStripMenuItem;
        private ToolStripMenuItem ヘルプToolStripMenuItem;
        private ToolStripMenuItem 取得情報ToolStripMenuItem;
        private Button BarcodePrintButton;
        private Button LabelPrintButton;
        private NumericUpDown PrintPostionNumericUpDown;
        private Label PrintRowLabel;
        private PrintDialog RePrintPrintDialog;
        private System.Drawing.Printing.PrintDocument RePrintPrintDocument;
        private PrintPreviewDialog RePrintPrintPreviewDialog;
    }
}