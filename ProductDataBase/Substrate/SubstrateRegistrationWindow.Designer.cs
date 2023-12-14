namespace ProductDatabase {
    partial class SubstrateRegistrationWindow {
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SubstrateRegistrationWindow));
            SubstrateRegistrationMenuStrip = new MenuStrip();
            ファイルToolStripMenuItem = new ToolStripMenuItem();
            印刷ToolStripMenuItem = new ToolStripMenuItem();
            印刷プレビューToolStripMenuItem = new ToolStripMenuItem();
            終了ToolStripMenuItem = new ToolStripMenuItem();
            設定ToolStripMenuItem = new ToolStripMenuItem();
            印刷設定ToolStripMenuItem = new ToolStripMenuItem();
            ヘルプToolStripMenuItem = new ToolStripMenuItem();
            取得情報ToolStripMenuItem = new ToolStripMenuItem();
            ProductNameLabel1 = new Label();
            ProductNameLabel2 = new Label();
            SubstrateModelLabel1 = new Label();
            SubstrateModelLabel2 = new Label();
            StockLabel1 = new Label();
            StockLabel2 = new Label();
            OrderNumberCheckBox = new CheckBox();
            OrderNumberTextBox = new TextBox();
            ManufacturingNumberCheckBox = new CheckBox();
            ManufacturingNumberMaskedTextBox = new MaskedTextBox();
            QuantityTextBox = new TextBox();
            QuantityCheckBox = new CheckBox();
            DefectNumberTextBox = new TextBox();
            DefectNumberCheckBox = new CheckBox();
            RevisionTextBox = new TextBox();
            RevisionCheckBox = new CheckBox();
            ExtraTextBox1 = new TextBox();
            ExtraCheckBox1 = new CheckBox();
            ExtraTextBox2 = new TextBox();
            ExtraCheckBox2 = new CheckBox();
            ExtraTextBox3 = new TextBox();
            ExtraCheckBox3 = new CheckBox();
            RegistrationDateCheckBox = new CheckBox();
            RegistrationDateMaskedTextBox = new MaskedTextBox();
            PersonCheckBox = new CheckBox();
            PersonComboBox = new ComboBox();
            ExtraTextBox4 = new TextBox();
            ExtraCheckBox4 = new CheckBox();
            ExtraTextBox5 = new TextBox();
            ExtraCheckBox5 = new CheckBox();
            ExtraTextBox6 = new TextBox();
            ExtraCheckBox6 = new CheckBox();
            CommentCheckBox = new CheckBox();
            CommentTextBox = new TextBox();
            CommentComboBox = new ComboBox();
            TemplateButton = new Button();
            PrintRowLabel = new Label();
            PrintPostionNumericUpDown = new NumericUpDown();
            RegisterButton = new Button();
            PrintOnlyCheckBox = new CheckBox();
            PrintButton = new Button();
            SubstrateRegistrationPrintDialog = new PrintDialog();
            SubstrateRegistrationPrintDocument = new System.Drawing.Printing.PrintDocument();
            SubstrateRegistrationPrintPreviewDialog = new PrintPreviewDialog();
            SubstrateRegistrationMenuStrip.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)PrintPostionNumericUpDown).BeginInit();
            SuspendLayout();
            // 
            // SubstrateRegistrationMenuStrip
            // 
            SubstrateRegistrationMenuStrip.Font = new Font("Meiryo UI", 9F, FontStyle.Regular, GraphicsUnit.Point);
            SubstrateRegistrationMenuStrip.Items.AddRange(new ToolStripItem[] { ファイルToolStripMenuItem, 設定ToolStripMenuItem, ヘルプToolStripMenuItem });
            SubstrateRegistrationMenuStrip.Location = new Point(0, 0);
            SubstrateRegistrationMenuStrip.Name = "SubstrateRegistrationMenuStrip";
            SubstrateRegistrationMenuStrip.Size = new Size(630, 24);
            SubstrateRegistrationMenuStrip.TabIndex = 0;
            SubstrateRegistrationMenuStrip.Text = "menuStrip";
            // 
            // ファイルToolStripMenuItem
            // 
            ファイルToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { 印刷ToolStripMenuItem, 印刷プレビューToolStripMenuItem, 終了ToolStripMenuItem });
            ファイルToolStripMenuItem.Name = "ファイルToolStripMenuItem";
            ファイルToolStripMenuItem.Size = new Size(53, 20);
            ファイルToolStripMenuItem.Text = "ファイル";
            // 
            // 印刷ToolStripMenuItem
            // 
            印刷ToolStripMenuItem.Name = "印刷ToolStripMenuItem";
            印刷ToolStripMenuItem.Size = new Size(142, 22);
            印刷ToolStripMenuItem.Text = "印刷";
            印刷ToolStripMenuItem.Click += 印刷ToolStripMenuItem_Click;
            // 
            // 印刷プレビューToolStripMenuItem
            // 
            印刷プレビューToolStripMenuItem.Name = "印刷プレビューToolStripMenuItem";
            印刷プレビューToolStripMenuItem.Size = new Size(142, 22);
            印刷プレビューToolStripMenuItem.Text = "印刷プレビュー";
            印刷プレビューToolStripMenuItem.Click += 印刷プレビューToolStripMenuItem_Click;
            // 
            // 終了ToolStripMenuItem
            // 
            終了ToolStripMenuItem.Name = "終了ToolStripMenuItem";
            終了ToolStripMenuItem.Size = new Size(142, 22);
            終了ToolStripMenuItem.Text = "終了";
            // 
            // 設定ToolStripMenuItem
            // 
            設定ToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { 印刷設定ToolStripMenuItem });
            設定ToolStripMenuItem.Name = "設定ToolStripMenuItem";
            設定ToolStripMenuItem.Size = new Size(43, 20);
            設定ToolStripMenuItem.Text = "設定";
            // 
            // 印刷設定ToolStripMenuItem
            // 
            印刷設定ToolStripMenuItem.Name = "印刷設定ToolStripMenuItem";
            印刷設定ToolStripMenuItem.Size = new Size(122, 22);
            印刷設定ToolStripMenuItem.Text = "印刷設定";
            印刷設定ToolStripMenuItem.Click += 印刷設定ToolStripMenuItem_Click;
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
            // ProductNameLabel1
            // 
            ProductNameLabel1.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            ProductNameLabel1.Location = new Point(28, 39);
            ProductNameLabel1.Margin = new Padding(4, 0, 4, 0);
            ProductNameLabel1.Name = "ProductNameLabel1";
            ProductNameLabel1.Size = new Size(48, 15);
            ProductNameLabel1.TabIndex = 1;
            ProductNameLabel1.Text = "製品名:";
            ProductNameLabel1.TextAlign = ContentAlignment.MiddleRight;
            // 
            // ProductNameLabel2
            // 
            ProductNameLabel2.AutoSize = true;
            ProductNameLabel2.BorderStyle = BorderStyle.Fixed3D;
            ProductNameLabel2.Location = new Point(84, 39);
            ProductNameLabel2.Margin = new Padding(4, 0, 4, 0);
            ProductNameLabel2.Name = "ProductNameLabel2";
            ProductNameLabel2.Size = new Size(24, 17);
            ProductNameLabel2.TabIndex = 2;
            ProductNameLabel2.Text = "---";
            ProductNameLabel2.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // SubstrateModelLabel1
            // 
            SubstrateModelLabel1.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            SubstrateModelLabel1.Location = new Point(16, 61);
            SubstrateModelLabel1.Margin = new Padding(4, 0, 4, 0);
            SubstrateModelLabel1.Name = "SubstrateModelLabel1";
            SubstrateModelLabel1.Size = new Size(60, 15);
            SubstrateModelLabel1.TabIndex = 3;
            SubstrateModelLabel1.Text = "基盤型式:";
            SubstrateModelLabel1.TextAlign = ContentAlignment.MiddleRight;
            // 
            // SubstrateModelLabel2
            // 
            SubstrateModelLabel2.AutoSize = true;
            SubstrateModelLabel2.BorderStyle = BorderStyle.Fixed3D;
            SubstrateModelLabel2.Location = new Point(84, 61);
            SubstrateModelLabel2.Margin = new Padding(4, 0, 4, 0);
            SubstrateModelLabel2.Name = "SubstrateModelLabel2";
            SubstrateModelLabel2.Size = new Size(24, 17);
            SubstrateModelLabel2.TabIndex = 4;
            SubstrateModelLabel2.Text = "---";
            SubstrateModelLabel2.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // StockLabel1
            // 
            StockLabel1.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            StockLabel1.Location = new Point(40, 83);
            StockLabel1.Margin = new Padding(4, 0, 4, 0);
            StockLabel1.Name = "StockLabel1";
            StockLabel1.Size = new Size(36, 15);
            StockLabel1.TabIndex = 5;
            StockLabel1.Text = "在庫:";
            StockLabel1.TextAlign = ContentAlignment.MiddleRight;
            // 
            // StockLabel2
            // 
            StockLabel2.AutoSize = true;
            StockLabel2.BorderStyle = BorderStyle.Fixed3D;
            StockLabel2.Location = new Point(84, 83);
            StockLabel2.Margin = new Padding(4, 0, 4, 0);
            StockLabel2.Name = "StockLabel2";
            StockLabel2.Size = new Size(24, 17);
            StockLabel2.TabIndex = 6;
            StockLabel2.Text = "---";
            StockLabel2.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // OrderNumberCheckBox
            // 
            OrderNumberCheckBox.Location = new Point(19, 118);
            OrderNumberCheckBox.Margin = new Padding(0);
            OrderNumberCheckBox.Name = "OrderNumberCheckBox";
            OrderNumberCheckBox.Size = new Size(74, 19);
            OrderNumberCheckBox.TabIndex = 7;
            OrderNumberCheckBox.TabStop = false;
            OrderNumberCheckBox.Text = "注文番号";
            OrderNumberCheckBox.UseVisualStyleBackColor = true;
            OrderNumberCheckBox.CheckedChanged += NumberCheckBox_CheckedChanged;
            // 
            // OrderNumberTextBox
            // 
            OrderNumberTextBox.Enabled = false;
            OrderNumberTextBox.Location = new Point(19, 141);
            OrderNumberTextBox.Margin = new Padding(0);
            OrderNumberTextBox.MaxLength = 20;
            OrderNumberTextBox.Name = "OrderNumberTextBox";
            OrderNumberTextBox.ShortcutsEnabled = false;
            OrderNumberTextBox.Size = new Size(120, 23);
            OrderNumberTextBox.TabIndex = 8;
            OrderNumberTextBox.TextAlign = HorizontalAlignment.Right;
            // 
            // ManufacturingNumberCheckBox
            // 
            ManufacturingNumberCheckBox.Location = new Point(19, 168);
            ManufacturingNumberCheckBox.Margin = new Padding(0);
            ManufacturingNumberCheckBox.Name = "ManufacturingNumberCheckBox";
            ManufacturingNumberCheckBox.Size = new Size(74, 19);
            ManufacturingNumberCheckBox.TabIndex = 9;
            ManufacturingNumberCheckBox.TabStop = false;
            ManufacturingNumberCheckBox.Text = "製造番号";
            ManufacturingNumberCheckBox.UseVisualStyleBackColor = true;
            ManufacturingNumberCheckBox.CheckedChanged += CheckBoxChecked;
            // 
            // ManufacturingNumberMaskedTextBox
            // 
            ManufacturingNumberMaskedTextBox.Enabled = false;
            ManufacturingNumberMaskedTextBox.Location = new Point(19, 191);
            ManufacturingNumberMaskedTextBox.Mask = "LA00L00000-0000";
            ManufacturingNumberMaskedTextBox.Name = "ManufacturingNumberMaskedTextBox";
            ManufacturingNumberMaskedTextBox.PromptChar = '*';
            ManufacturingNumberMaskedTextBox.Size = new Size(120, 23);
            ManufacturingNumberMaskedTextBox.TabIndex = 10;
            ManufacturingNumberMaskedTextBox.Text = "H";
            ManufacturingNumberMaskedTextBox.TextAlign = HorizontalAlignment.Right;
            // 
            // QuantityTextBox
            // 
            QuantityTextBox.Enabled = false;
            QuantityTextBox.Location = new Point(19, 241);
            QuantityTextBox.Margin = new Padding(0);
            QuantityTextBox.MaxLength = 4;
            QuantityTextBox.Name = "QuantityTextBox";
            QuantityTextBox.ShortcutsEnabled = false;
            QuantityTextBox.Size = new Size(120, 23);
            QuantityTextBox.TabIndex = 12;
            QuantityTextBox.TextAlign = HorizontalAlignment.Right;
            QuantityTextBox.KeyPress += QuantityTextBox_KeyPress;
            // 
            // QuantityCheckBox
            // 
            QuantityCheckBox.Location = new Point(19, 218);
            QuantityCheckBox.Margin = new Padding(0);
            QuantityCheckBox.Name = "QuantityCheckBox";
            QuantityCheckBox.Size = new Size(50, 19);
            QuantityCheckBox.TabIndex = 11;
            QuantityCheckBox.TabStop = false;
            QuantityCheckBox.Text = "数量";
            QuantityCheckBox.UseVisualStyleBackColor = true;
            QuantityCheckBox.CheckedChanged += CheckBoxChecked;
            // 
            // DefectNumberTextBox
            // 
            DefectNumberTextBox.Enabled = false;
            DefectNumberTextBox.Location = new Point(19, 291);
            DefectNumberTextBox.Margin = new Padding(0);
            DefectNumberTextBox.MaxLength = 4;
            DefectNumberTextBox.Name = "DefectNumberTextBox";
            DefectNumberTextBox.ShortcutsEnabled = false;
            DefectNumberTextBox.Size = new Size(120, 23);
            DefectNumberTextBox.TabIndex = 14;
            DefectNumberTextBox.TextAlign = HorizontalAlignment.Right;
            DefectNumberTextBox.KeyPress += DefectNumberTextBox_KeyPress;
            // 
            // DefectNumberCheckBox
            // 
            DefectNumberCheckBox.Location = new Point(19, 268);
            DefectNumberCheckBox.Margin = new Padding(0);
            DefectNumberCheckBox.Name = "DefectNumberCheckBox";
            DefectNumberCheckBox.Size = new Size(62, 19);
            DefectNumberCheckBox.TabIndex = 13;
            DefectNumberCheckBox.TabStop = false;
            DefectNumberCheckBox.Text = "不良量";
            DefectNumberCheckBox.UseVisualStyleBackColor = true;
            DefectNumberCheckBox.CheckedChanged += CheckBoxChecked;
            // 
            // RevisionTextBox
            // 
            RevisionTextBox.Enabled = false;
            RevisionTextBox.Location = new Point(19, 341);
            RevisionTextBox.Margin = new Padding(0);
            RevisionTextBox.MaxLength = 4;
            RevisionTextBox.Name = "RevisionTextBox";
            RevisionTextBox.ShortcutsEnabled = false;
            RevisionTextBox.Size = new Size(120, 23);
            RevisionTextBox.TabIndex = 16;
            RevisionTextBox.TextAlign = HorizontalAlignment.Right;
            // 
            // RevisionCheckBox
            // 
            RevisionCheckBox.Location = new Point(19, 318);
            RevisionCheckBox.Margin = new Padding(0);
            RevisionCheckBox.Name = "RevisionCheckBox";
            RevisionCheckBox.Size = new Size(68, 19);
            RevisionCheckBox.TabIndex = 15;
            RevisionCheckBox.TabStop = false;
            RevisionCheckBox.Text = "レビジョン";
            RevisionCheckBox.UseVisualStyleBackColor = true;
            RevisionCheckBox.CheckedChanged += CheckBoxChecked;
            // 
            // ExtraTextBox1
            // 
            ExtraTextBox1.Enabled = false;
            ExtraTextBox1.Location = new Point(166, 141);
            ExtraTextBox1.Margin = new Padding(0);
            ExtraTextBox1.MaxLength = 20;
            ExtraTextBox1.Name = "ExtraTextBox1";
            ExtraTextBox1.ShortcutsEnabled = false;
            ExtraTextBox1.Size = new Size(120, 23);
            ExtraTextBox1.TabIndex = 18;
            ExtraTextBox1.TextAlign = HorizontalAlignment.Right;
            // 
            // ExtraCheckBox1
            // 
            ExtraCheckBox1.Location = new Point(166, 118);
            ExtraCheckBox1.Margin = new Padding(0);
            ExtraCheckBox1.Name = "ExtraCheckBox1";
            ExtraCheckBox1.Size = new Size(50, 19);
            ExtraCheckBox1.TabIndex = 17;
            ExtraCheckBox1.TabStop = false;
            ExtraCheckBox1.Text = "予備";
            ExtraCheckBox1.UseVisualStyleBackColor = true;
            ExtraCheckBox1.CheckedChanged += CheckBoxChecked;
            // 
            // ExtraTextBox2
            // 
            ExtraTextBox2.Enabled = false;
            ExtraTextBox2.Location = new Point(166, 191);
            ExtraTextBox2.Margin = new Padding(0);
            ExtraTextBox2.MaxLength = 20;
            ExtraTextBox2.Name = "ExtraTextBox2";
            ExtraTextBox2.ShortcutsEnabled = false;
            ExtraTextBox2.Size = new Size(120, 23);
            ExtraTextBox2.TabIndex = 20;
            ExtraTextBox2.TextAlign = HorizontalAlignment.Right;
            // 
            // ExtraCheckBox2
            // 
            ExtraCheckBox2.Location = new Point(166, 168);
            ExtraCheckBox2.Margin = new Padding(0);
            ExtraCheckBox2.Name = "ExtraCheckBox2";
            ExtraCheckBox2.Size = new Size(50, 19);
            ExtraCheckBox2.TabIndex = 19;
            ExtraCheckBox2.TabStop = false;
            ExtraCheckBox2.Text = "予備";
            ExtraCheckBox2.UseVisualStyleBackColor = true;
            ExtraCheckBox2.CheckedChanged += CheckBoxChecked;
            // 
            // ExtraTextBox3
            // 
            ExtraTextBox3.Enabled = false;
            ExtraTextBox3.Location = new Point(166, 241);
            ExtraTextBox3.Margin = new Padding(0);
            ExtraTextBox3.MaxLength = 20;
            ExtraTextBox3.Name = "ExtraTextBox3";
            ExtraTextBox3.ShortcutsEnabled = false;
            ExtraTextBox3.Size = new Size(120, 23);
            ExtraTextBox3.TabIndex = 22;
            ExtraTextBox3.TextAlign = HorizontalAlignment.Right;
            // 
            // ExtraCheckBox3
            // 
            ExtraCheckBox3.Location = new Point(166, 218);
            ExtraCheckBox3.Margin = new Padding(0);
            ExtraCheckBox3.Name = "ExtraCheckBox3";
            ExtraCheckBox3.Size = new Size(50, 19);
            ExtraCheckBox3.TabIndex = 21;
            ExtraCheckBox3.TabStop = false;
            ExtraCheckBox3.Text = "予備";
            ExtraCheckBox3.UseVisualStyleBackColor = true;
            ExtraCheckBox3.CheckedChanged += CheckBoxChecked;
            // 
            // RegistrationDateCheckBox
            // 
            RegistrationDateCheckBox.Location = new Point(166, 268);
            RegistrationDateCheckBox.Margin = new Padding(0);
            RegistrationDateCheckBox.Name = "RegistrationDateCheckBox";
            RegistrationDateCheckBox.Size = new Size(62, 19);
            RegistrationDateCheckBox.TabIndex = 23;
            RegistrationDateCheckBox.TabStop = false;
            RegistrationDateCheckBox.Text = "登録日";
            RegistrationDateCheckBox.UseVisualStyleBackColor = true;
            RegistrationDateCheckBox.CheckedChanged += CheckBoxChecked;
            // 
            // RegistrationDateMaskedTextBox
            // 
            RegistrationDateMaskedTextBox.Enabled = false;
            RegistrationDateMaskedTextBox.Location = new Point(166, 291);
            RegistrationDateMaskedTextBox.Mask = "0000/00/00";
            RegistrationDateMaskedTextBox.Name = "RegistrationDateMaskedTextBox";
            RegistrationDateMaskedTextBox.PromptChar = '*';
            RegistrationDateMaskedTextBox.Size = new Size(120, 23);
            RegistrationDateMaskedTextBox.TabIndex = 24;
            RegistrationDateMaskedTextBox.TextAlign = HorizontalAlignment.Right;
            RegistrationDateMaskedTextBox.TypeValidationCompleted += RegistrationDateMaskedTextBox_TypeValidationCompleted;
            // 
            // PersonCheckBox
            // 
            PersonCheckBox.Location = new Point(166, 318);
            PersonCheckBox.Margin = new Padding(0);
            PersonCheckBox.Name = "PersonCheckBox";
            PersonCheckBox.Size = new Size(62, 19);
            PersonCheckBox.TabIndex = 25;
            PersonCheckBox.TabStop = false;
            PersonCheckBox.Text = "担当者";
            PersonCheckBox.UseVisualStyleBackColor = true;
            PersonCheckBox.CheckedChanged += CheckBoxChecked;
            // 
            // PersonComboBox
            // 
            PersonComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            PersonComboBox.Enabled = false;
            PersonComboBox.FormattingEnabled = true;
            PersonComboBox.Location = new Point(166, 341);
            PersonComboBox.Name = "PersonComboBox";
            PersonComboBox.Size = new Size(121, 23);
            PersonComboBox.TabIndex = 26;
            // 
            // ExtraTextBox4
            // 
            ExtraTextBox4.Enabled = false;
            ExtraTextBox4.Location = new Point(313, 141);
            ExtraTextBox4.Margin = new Padding(0);
            ExtraTextBox4.MaxLength = 20;
            ExtraTextBox4.Name = "ExtraTextBox4";
            ExtraTextBox4.ShortcutsEnabled = false;
            ExtraTextBox4.Size = new Size(120, 23);
            ExtraTextBox4.TabIndex = 28;
            ExtraTextBox4.TextAlign = HorizontalAlignment.Right;
            // 
            // ExtraCheckBox4
            // 
            ExtraCheckBox4.Location = new Point(313, 118);
            ExtraCheckBox4.Margin = new Padding(0);
            ExtraCheckBox4.Name = "ExtraCheckBox4";
            ExtraCheckBox4.Size = new Size(50, 19);
            ExtraCheckBox4.TabIndex = 27;
            ExtraCheckBox4.TabStop = false;
            ExtraCheckBox4.Text = "予備";
            ExtraCheckBox4.UseVisualStyleBackColor = true;
            ExtraCheckBox4.CheckedChanged += CheckBoxChecked;
            // 
            // ExtraTextBox5
            // 
            ExtraTextBox5.Enabled = false;
            ExtraTextBox5.Location = new Point(313, 191);
            ExtraTextBox5.Margin = new Padding(0);
            ExtraTextBox5.MaxLength = 20;
            ExtraTextBox5.Name = "ExtraTextBox5";
            ExtraTextBox5.ShortcutsEnabled = false;
            ExtraTextBox5.Size = new Size(120, 23);
            ExtraTextBox5.TabIndex = 30;
            ExtraTextBox5.TextAlign = HorizontalAlignment.Right;
            // 
            // ExtraCheckBox5
            // 
            ExtraCheckBox5.Location = new Point(313, 168);
            ExtraCheckBox5.Margin = new Padding(0);
            ExtraCheckBox5.Name = "ExtraCheckBox5";
            ExtraCheckBox5.Size = new Size(50, 19);
            ExtraCheckBox5.TabIndex = 29;
            ExtraCheckBox5.TabStop = false;
            ExtraCheckBox5.Text = "予備";
            ExtraCheckBox5.UseVisualStyleBackColor = true;
            ExtraCheckBox5.CheckedChanged += CheckBoxChecked;
            // 
            // ExtraTextBox6
            // 
            ExtraTextBox6.Enabled = false;
            ExtraTextBox6.Location = new Point(313, 241);
            ExtraTextBox6.Margin = new Padding(0);
            ExtraTextBox6.MaxLength = 20;
            ExtraTextBox6.Name = "ExtraTextBox6";
            ExtraTextBox6.ShortcutsEnabled = false;
            ExtraTextBox6.Size = new Size(120, 23);
            ExtraTextBox6.TabIndex = 32;
            ExtraTextBox6.TextAlign = HorizontalAlignment.Right;
            // 
            // ExtraCheckBox6
            // 
            ExtraCheckBox6.Location = new Point(313, 218);
            ExtraCheckBox6.Margin = new Padding(0);
            ExtraCheckBox6.Name = "ExtraCheckBox6";
            ExtraCheckBox6.Size = new Size(50, 19);
            ExtraCheckBox6.TabIndex = 31;
            ExtraCheckBox6.TabStop = false;
            ExtraCheckBox6.Text = "予備";
            ExtraCheckBox6.UseVisualStyleBackColor = true;
            ExtraCheckBox6.CheckedChanged += CheckBoxChecked;
            // 
            // CommentCheckBox
            // 
            CommentCheckBox.Location = new Point(464, 118);
            CommentCheckBox.Margin = new Padding(0);
            CommentCheckBox.Name = "CommentCheckBox";
            CommentCheckBox.Size = new Size(59, 19);
            CommentCheckBox.TabIndex = 33;
            CommentCheckBox.TabStop = false;
            CommentCheckBox.Text = "コメント";
            CommentCheckBox.UseVisualStyleBackColor = true;
            CommentCheckBox.CheckedChanged += CheckBoxChecked;
            // 
            // CommentTextBox
            // 
            CommentTextBox.Enabled = false;
            CommentTextBox.Location = new Point(464, 141);
            CommentTextBox.Margin = new Padding(0);
            CommentTextBox.MaxLength = 500;
            CommentTextBox.Multiline = true;
            CommentTextBox.Name = "CommentTextBox";
            CommentTextBox.ScrollBars = ScrollBars.Vertical;
            CommentTextBox.ShortcutsEnabled = false;
            CommentTextBox.Size = new Size(150, 175);
            CommentTextBox.TabIndex = 34;
            // 
            // CommentComboBox
            // 
            CommentComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            CommentComboBox.DropDownWidth = 150;
            CommentComboBox.Enabled = false;
            CommentComboBox.FormattingEnabled = true;
            CommentComboBox.Location = new Point(464, 320);
            CommentComboBox.Name = "CommentComboBox";
            CommentComboBox.Size = new Size(150, 23);
            CommentComboBox.TabIndex = 35;
            // 
            // TemplateButton
            // 
            TemplateButton.Enabled = false;
            TemplateButton.Location = new Point(539, 347);
            TemplateButton.Name = "TemplateButton";
            TemplateButton.Size = new Size(75, 25);
            TemplateButton.TabIndex = 36;
            TemplateButton.Text = "定型文";
            TemplateButton.UseVisualStyleBackColor = true;
            TemplateButton.Click += TemplateButton_Click;
            // 
            // PrintRowLabel
            // 
            PrintRowLabel.Location = new Point(177, 392);
            PrintRowLabel.Margin = new Padding(0);
            PrintRowLabel.Name = "PrintRowLabel";
            PrintRowLabel.Size = new Size(79, 15);
            PrintRowLabel.TabIndex = 37;
            PrintRowLabel.Text = "印刷開始位置";
            // 
            // PrintPostionNumericUpDown
            // 
            PrintPostionNumericUpDown.Location = new Point(166, 410);
            PrintPostionNumericUpDown.Margin = new Padding(0);
            PrintPostionNumericUpDown.Maximum = new decimal(new int[] { 50, 0, 0, 0 });
            PrintPostionNumericUpDown.Minimum = new decimal(new int[] { 1, 0, 0, 0 });
            PrintPostionNumericUpDown.Name = "PrintPostionNumericUpDown";
            PrintPostionNumericUpDown.Size = new Size(110, 23);
            PrintPostionNumericUpDown.TabIndex = 38;
            PrintPostionNumericUpDown.TextAlign = HorizontalAlignment.Right;
            PrintPostionNumericUpDown.Value = new decimal(new int[] { 1, 0, 0, 0 });
            // 
            // RegisterButton
            // 
            RegisterButton.Location = new Point(313, 407);
            RegisterButton.Margin = new Padding(0);
            RegisterButton.Name = "RegisterButton";
            RegisterButton.Size = new Size(75, 25);
            RegisterButton.TabIndex = 39;
            RegisterButton.Text = "OK";
            RegisterButton.UseVisualStyleBackColor = true;
            RegisterButton.Click += RegisterButton_Click;
            // 
            // PrintOnlyCheckBox
            // 
            PrintOnlyCheckBox.Location = new Point(426, 382);
            PrintOnlyCheckBox.Name = "PrintOnlyCheckBox";
            PrintOnlyCheckBox.Size = new Size(71, 19);
            PrintOnlyCheckBox.TabIndex = 40;
            PrintOnlyCheckBox.TabStop = false;
            PrintOnlyCheckBox.Text = "印刷のみ";
            PrintOnlyCheckBox.UseVisualStyleBackColor = true;
            PrintOnlyCheckBox.CheckedChanged += CheckBoxChecked;
            // 
            // PrintButton
            // 
            PrintButton.Enabled = false;
            PrintButton.Location = new Point(426, 407);
            PrintButton.Margin = new Padding(0);
            PrintButton.Name = "PrintButton";
            PrintButton.Size = new Size(75, 25);
            PrintButton.TabIndex = 41;
            PrintButton.Text = "印刷";
            PrintButton.UseVisualStyleBackColor = true;
            PrintButton.Click += PrintButton_Click;
            // 
            // SubstrateRegistrationPrintDialog
            // 
            SubstrateRegistrationPrintDialog.UseEXDialog = true;
            // 
            // SubstrateRegistrationPrintDocument
            // 
            SubstrateRegistrationPrintDocument.PrintPage += SubstrateRegistrationPrintDocument_PrintPage;
            // 
            // SubstrateRegistrationPrintPreviewDialog
            // 
            SubstrateRegistrationPrintPreviewDialog.AutoScrollMargin = new Size(0, 0);
            SubstrateRegistrationPrintPreviewDialog.AutoScrollMinSize = new Size(0, 0);
            SubstrateRegistrationPrintPreviewDialog.ClientSize = new Size(400, 300);
            SubstrateRegistrationPrintPreviewDialog.Enabled = true;
            SubstrateRegistrationPrintPreviewDialog.Icon = (Icon)resources.GetObject("SubstrateRegistrationPrintPreviewDialog.Icon");
            SubstrateRegistrationPrintPreviewDialog.Name = "SubstrateRegistrationPrintPreviewDialog";
            SubstrateRegistrationPrintPreviewDialog.Visible = false;
            SubstrateRegistrationPrintPreviewDialog.Load += SubstrateRegistrationPrintPreviewDialog_Load;
            // 
            // SubstrateRegistrationWindow
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(630, 441);
            Controls.Add(PrintButton);
            Controls.Add(PrintOnlyCheckBox);
            Controls.Add(RegisterButton);
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
            Controls.Add(ExtraTextBox3);
            Controls.Add(ExtraCheckBox3);
            Controls.Add(ExtraTextBox2);
            Controls.Add(ExtraCheckBox2);
            Controls.Add(ExtraTextBox1);
            Controls.Add(ExtraCheckBox1);
            Controls.Add(RevisionTextBox);
            Controls.Add(RevisionCheckBox);
            Controls.Add(DefectNumberTextBox);
            Controls.Add(DefectNumberCheckBox);
            Controls.Add(QuantityTextBox);
            Controls.Add(QuantityCheckBox);
            Controls.Add(ManufacturingNumberMaskedTextBox);
            Controls.Add(ManufacturingNumberCheckBox);
            Controls.Add(OrderNumberTextBox);
            Controls.Add(OrderNumberCheckBox);
            Controls.Add(StockLabel2);
            Controls.Add(StockLabel1);
            Controls.Add(SubstrateModelLabel2);
            Controls.Add(SubstrateModelLabel1);
            Controls.Add(ProductNameLabel2);
            Controls.Add(ProductNameLabel1);
            Controls.Add(SubstrateRegistrationMenuStrip);
            Font = new Font("Meiryo UI", 9F, FontStyle.Regular, GraphicsUnit.Point);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MainMenuStrip = SubstrateRegistrationMenuStrip;
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "SubstrateRegistrationWindow";
            ShowIcon = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterParent;
            Text = "基盤登録";
            FormClosing += SubstrateRegistrationWindow_FormClosing;
            Load += SubstrateRegistrationWindow_Load;
            SubstrateRegistrationMenuStrip.ResumeLayout(false);
            SubstrateRegistrationMenuStrip.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)PrintPostionNumericUpDown).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private MenuStrip SubstrateRegistrationMenuStrip;
        private ToolStripMenuItem ファイルToolStripMenuItem;
        private ToolStripMenuItem 印刷ToolStripMenuItem;
        private ToolStripMenuItem 印刷プレビューToolStripMenuItem;
        private ToolStripMenuItem 終了ToolStripMenuItem;
        private ToolStripMenuItem 設定ToolStripMenuItem;
        private ToolStripMenuItem 印刷設定ToolStripMenuItem;
        private ToolStripMenuItem ヘルプToolStripMenuItem;
        private ToolStripMenuItem 取得情報ToolStripMenuItem;
        private Label ProductNameLabel1;
        private Label ProductNameLabel2;
        private Label SubstrateModelLabel1;
        private Label SubstrateModelLabel2;
        private Label StockLabel1;
        private Label StockLabel2;
        private CheckBox OrderNumberCheckBox;
        private TextBox OrderNumberTextBox;
        private CheckBox ManufacturingNumberCheckBox;
        private MaskedTextBox ManufacturingNumberMaskedTextBox;
        private TextBox QuantityTextBox;
        private CheckBox QuantityCheckBox;
        private TextBox DefectNumberTextBox;
        private CheckBox DefectNumberCheckBox;
        private TextBox RevisionTextBox;
        private CheckBox RevisionCheckBox;
        private TextBox ExtraTextBox1;
        private CheckBox ExtraCheckBox1;
        private TextBox ExtraTextBox2;
        private CheckBox ExtraCheckBox2;
        private TextBox ExtraTextBox3;
        private CheckBox ExtraCheckBox3;
        private CheckBox RegistrationDateCheckBox;
        private MaskedTextBox RegistrationDateMaskedTextBox;
        private CheckBox PersonCheckBox;
        private ComboBox PersonComboBox;
        private TextBox ExtraTextBox4;
        private CheckBox ExtraCheckBox4;
        private TextBox ExtraTextBox5;
        private CheckBox ExtraCheckBox5;
        private TextBox ExtraTextBox6;
        private CheckBox ExtraCheckBox6;
        private CheckBox CommentCheckBox;
        private TextBox CommentTextBox;
        private ComboBox CommentComboBox;
        private Button TemplateButton;
        private Label PrintRowLabel;
        private NumericUpDown PrintPostionNumericUpDown;
        private Button RegisterButton;
        private CheckBox PrintOnlyCheckBox;
        private Button PrintButton;
        private PrintDialog SubstrateRegistrationPrintDialog;
        private System.Drawing.Printing.PrintDocument SubstrateRegistrationPrintDocument;
        private PrintPreviewDialog SubstrateRegistrationPrintPreviewDialog;
    }
}