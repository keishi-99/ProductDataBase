namespace ProductDatabase {
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
            this.TemplateButton = new Button();
            this.CommentComboBox = new ComboBox();
            this.CommentTextBox = new TextBox();
            this.CommentCheckBox = new CheckBox();
            this.ExtraTextBox6 = new TextBox();
            this.ExtraCheckBox6 = new CheckBox();
            this.ExtraTextBox5 = new TextBox();
            this.ExtraCheckBox5 = new CheckBox();
            this.ExtraTextBox4 = new TextBox();
            this.ExtraCheckBox4 = new CheckBox();
            this.PersonComboBox = new ComboBox();
            this.PersonCheckBox = new CheckBox();
            this.RegistrationDateCheckBox = new CheckBox();
            this.FirstSerialNumberTextBox = new TextBox();
            this.FirstSerialNumberCheckBox = new CheckBox();
            this.ExtraTextBox3 = new TextBox();
            this.ExtraCheckBox3 = new CheckBox();
            this.ExtraTextBox2 = new TextBox();
            this.ExtraCheckBox2 = new CheckBox();
            this.RevisionTextBox = new TextBox();
            this.RevisionCheckBox = new CheckBox();
            this.ExtraTextBox1 = new TextBox();
            this.ExtraCheckBox1 = new CheckBox();
            this.QuantityTextBox = new TextBox();
            this.QuantityCheckBox = new CheckBox();
            this.ManufacturingNumberMaskedTextBox = new MaskedTextBox();
            this.ManufacturingNumberCheckBox = new CheckBox();
            this.OrderNumberTextBox = new TextBox();
            this.OrderNumberCheckBox = new CheckBox();
            this.SubstrateModelLabel2 = new Label();
            this.SubstrateModelLabel1 = new Label();
            this.ProductNameLabel2 = new Label();
            this.ProductNameLabel1 = new Label();
            this.RePrintMenuStrip = new MenuStrip();
            this.ファイルToolStripMenuItem = new ToolStripMenuItem();
            this.シリアルラベル印刷プレビューToolStripMenuItem = new ToolStripMenuItem();
            this.バーコード印刷プレビューToolStripMenuItem = new ToolStripMenuItem();
            this.終了ToolStripMenuItem = new ToolStripMenuItem();
            this.設定ToolStripMenuItem = new ToolStripMenuItem();
            this.シリアルラベル印刷設定ToolStripMenuItem = new ToolStripMenuItem();
            this.バーコード印刷設定ToolStripMenuItem = new ToolStripMenuItem();
            this.ヘルプToolStripMenuItem = new ToolStripMenuItem();
            this.取得情報ToolStripMenuItem = new ToolStripMenuItem();
            this.BarcodePrintButton = new Button();
            this.LabelPrintButton = new Button();
            this.PrintPositionNumericUpDown = new NumericUpDown();
            this.PrintRowLabel = new Label();
            this.panel1 = new Panel();
            this.QrCodeTextBox = new TextBox();
            this.ProductTypeLabel2 = new Label();
            this.ProductTypeLabel1 = new Label();
            this.label1 = new Label();
            this.QrCodeButton = new Button();
            this.textToUpperCheckBox = new CheckBox();
            this.RegistrationDateTimePicker = new DateTimePicker();
            this.RePrintMenuStrip.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)this.PrintPositionNumericUpDown).BeginInit();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // TemplateButton
            // 
            this.TemplateButton.Enabled = false;
            this.TemplateButton.Location = new Point(539, 347);
            this.TemplateButton.Name = "TemplateButton";
            this.TemplateButton.Size = new Size(75, 25);
            this.TemplateButton.TabIndex = 116;
            this.TemplateButton.Text = "定型文";
            this.TemplateButton.UseVisualStyleBackColor = true;
            // 
            // CommentComboBox
            // 
            this.CommentComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            this.CommentComboBox.DropDownWidth = 150;
            this.CommentComboBox.Enabled = false;
            this.CommentComboBox.FormattingEnabled = true;
            this.CommentComboBox.Location = new Point(464, 320);
            this.CommentComboBox.Name = "CommentComboBox";
            this.CommentComboBox.Size = new Size(150, 23);
            this.CommentComboBox.TabIndex = 115;
            // 
            // CommentTextBox
            // 
            this.CommentTextBox.Enabled = false;
            this.CommentTextBox.Location = new Point(0, 0);
            this.CommentTextBox.Margin = new Padding(0);
            this.CommentTextBox.MaxLength = 500;
            this.CommentTextBox.Multiline = true;
            this.CommentTextBox.Name = "CommentTextBox";
            this.CommentTextBox.ScrollBars = ScrollBars.Vertical;
            this.CommentTextBox.Size = new Size(150, 175);
            this.CommentTextBox.TabIndex = 114;
            // 
            // CommentCheckBox
            // 
            this.CommentCheckBox.Location = new Point(464, 118);
            this.CommentCheckBox.Margin = new Padding(0);
            this.CommentCheckBox.Name = "CommentCheckBox";
            this.CommentCheckBox.Size = new Size(59, 19);
            this.CommentCheckBox.TabIndex = 113;
            this.CommentCheckBox.TabStop = false;
            this.CommentCheckBox.Text = "コメント";
            this.CommentCheckBox.UseVisualStyleBackColor = true;
            this.CommentCheckBox.CheckedChanged += this.CheckBoxChecked;
            // 
            // ExtraTextBox6
            // 
            this.ExtraTextBox6.Enabled = false;
            this.ExtraTextBox6.Location = new Point(313, 241);
            this.ExtraTextBox6.Margin = new Padding(0);
            this.ExtraTextBox6.MaxLength = 20;
            this.ExtraTextBox6.Name = "ExtraTextBox6";
            this.ExtraTextBox6.Size = new Size(120, 23);
            this.ExtraTextBox6.TabIndex = 112;
            this.ExtraTextBox6.TextAlign = HorizontalAlignment.Right;
            // 
            // ExtraCheckBox6
            // 
            this.ExtraCheckBox6.Location = new Point(313, 218);
            this.ExtraCheckBox6.Margin = new Padding(0);
            this.ExtraCheckBox6.Name = "ExtraCheckBox6";
            this.ExtraCheckBox6.Size = new Size(50, 19);
            this.ExtraCheckBox6.TabIndex = 111;
            this.ExtraCheckBox6.TabStop = false;
            this.ExtraCheckBox6.Text = "予備";
            this.ExtraCheckBox6.UseVisualStyleBackColor = true;
            this.ExtraCheckBox6.Visible = false;
            this.ExtraCheckBox6.CheckedChanged += this.CheckBoxChecked;
            // 
            // ExtraTextBox5
            // 
            this.ExtraTextBox5.Enabled = false;
            this.ExtraTextBox5.Location = new Point(313, 191);
            this.ExtraTextBox5.Margin = new Padding(0);
            this.ExtraTextBox5.MaxLength = 20;
            this.ExtraTextBox5.Name = "ExtraTextBox5";
            this.ExtraTextBox5.Size = new Size(120, 23);
            this.ExtraTextBox5.TabIndex = 110;
            this.ExtraTextBox5.TextAlign = HorizontalAlignment.Right;
            // 
            // ExtraCheckBox5
            // 
            this.ExtraCheckBox5.Location = new Point(313, 168);
            this.ExtraCheckBox5.Margin = new Padding(0);
            this.ExtraCheckBox5.Name = "ExtraCheckBox5";
            this.ExtraCheckBox5.Size = new Size(50, 19);
            this.ExtraCheckBox5.TabIndex = 109;
            this.ExtraCheckBox5.TabStop = false;
            this.ExtraCheckBox5.Text = "予備";
            this.ExtraCheckBox5.UseVisualStyleBackColor = true;
            this.ExtraCheckBox5.Visible = false;
            this.ExtraCheckBox5.CheckedChanged += this.CheckBoxChecked;
            // 
            // ExtraTextBox4
            // 
            this.ExtraTextBox4.Enabled = false;
            this.ExtraTextBox4.Location = new Point(313, 141);
            this.ExtraTextBox4.Margin = new Padding(0);
            this.ExtraTextBox4.MaxLength = 20;
            this.ExtraTextBox4.Name = "ExtraTextBox4";
            this.ExtraTextBox4.Size = new Size(120, 23);
            this.ExtraTextBox4.TabIndex = 108;
            this.ExtraTextBox4.TextAlign = HorizontalAlignment.Right;
            // 
            // ExtraCheckBox4
            // 
            this.ExtraCheckBox4.Location = new Point(313, 118);
            this.ExtraCheckBox4.Margin = new Padding(0);
            this.ExtraCheckBox4.Name = "ExtraCheckBox4";
            this.ExtraCheckBox4.Size = new Size(50, 19);
            this.ExtraCheckBox4.TabIndex = 107;
            this.ExtraCheckBox4.TabStop = false;
            this.ExtraCheckBox4.Text = "予備";
            this.ExtraCheckBox4.UseVisualStyleBackColor = true;
            this.ExtraCheckBox4.Visible = false;
            this.ExtraCheckBox4.CheckedChanged += this.CheckBoxChecked;
            // 
            // PersonComboBox
            // 
            this.PersonComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            this.PersonComboBox.Enabled = false;
            this.PersonComboBox.FormattingEnabled = true;
            this.PersonComboBox.Location = new Point(166, 341);
            this.PersonComboBox.Name = "PersonComboBox";
            this.PersonComboBox.Size = new Size(121, 23);
            this.PersonComboBox.TabIndex = 106;
            // 
            // PersonCheckBox
            // 
            this.PersonCheckBox.Location = new Point(166, 318);
            this.PersonCheckBox.Margin = new Padding(0);
            this.PersonCheckBox.Name = "PersonCheckBox";
            this.PersonCheckBox.Size = new Size(62, 19);
            this.PersonCheckBox.TabIndex = 105;
            this.PersonCheckBox.TabStop = false;
            this.PersonCheckBox.Text = "担当者";
            this.PersonCheckBox.UseVisualStyleBackColor = true;
            this.PersonCheckBox.CheckedChanged += this.CheckBoxChecked;
            // 
            // RegistrationDateCheckBox
            // 
            this.RegistrationDateCheckBox.Location = new Point(166, 268);
            this.RegistrationDateCheckBox.Margin = new Padding(0);
            this.RegistrationDateCheckBox.Name = "RegistrationDateCheckBox";
            this.RegistrationDateCheckBox.Size = new Size(120, 19);
            this.RegistrationDateCheckBox.TabIndex = 103;
            this.RegistrationDateCheckBox.TabStop = false;
            this.RegistrationDateCheckBox.Text = "シリアル発行年月";
            this.RegistrationDateCheckBox.UseVisualStyleBackColor = true;
            this.RegistrationDateCheckBox.CheckedChanged += this.CheckBoxChecked;
            // 
            // FirstSerialNumberTextBox
            // 
            this.FirstSerialNumberTextBox.Enabled = false;
            this.FirstSerialNumberTextBox.Location = new Point(166, 241);
            this.FirstSerialNumberTextBox.Margin = new Padding(0);
            this.FirstSerialNumberTextBox.MaxLength = 20;
            this.FirstSerialNumberTextBox.Name = "FirstSerialNumberTextBox";
            this.FirstSerialNumberTextBox.Size = new Size(120, 23);
            this.FirstSerialNumberTextBox.TabIndex = 102;
            this.FirstSerialNumberTextBox.TextAlign = HorizontalAlignment.Right;
            // 
            // FirstSerialNumberCheckBox
            // 
            this.FirstSerialNumberCheckBox.Location = new Point(166, 218);
            this.FirstSerialNumberCheckBox.Margin = new Padding(0);
            this.FirstSerialNumberCheckBox.Name = "FirstSerialNumberCheckBox";
            this.FirstSerialNumberCheckBox.Size = new Size(110, 19);
            this.FirstSerialNumberCheckBox.TabIndex = 101;
            this.FirstSerialNumberCheckBox.TabStop = false;
            this.FirstSerialNumberCheckBox.Text = "先頭シリアル番号";
            this.FirstSerialNumberCheckBox.UseVisualStyleBackColor = true;
            this.FirstSerialNumberCheckBox.CheckedChanged += this.CheckBoxChecked;
            // 
            // ExtraTextBox3
            // 
            this.ExtraTextBox3.Enabled = false;
            this.ExtraTextBox3.Location = new Point(166, 191);
            this.ExtraTextBox3.Margin = new Padding(0);
            this.ExtraTextBox3.MaxLength = 20;
            this.ExtraTextBox3.Name = "ExtraTextBox3";
            this.ExtraTextBox3.Size = new Size(120, 23);
            this.ExtraTextBox3.TabIndex = 100;
            this.ExtraTextBox3.TextAlign = HorizontalAlignment.Right;
            // 
            // ExtraCheckBox3
            // 
            this.ExtraCheckBox3.Location = new Point(166, 168);
            this.ExtraCheckBox3.Margin = new Padding(0);
            this.ExtraCheckBox3.Name = "ExtraCheckBox3";
            this.ExtraCheckBox3.Size = new Size(50, 19);
            this.ExtraCheckBox3.TabIndex = 99;
            this.ExtraCheckBox3.TabStop = false;
            this.ExtraCheckBox3.Text = "予備";
            this.ExtraCheckBox3.UseVisualStyleBackColor = true;
            this.ExtraCheckBox3.Visible = false;
            this.ExtraCheckBox3.CheckedChanged += this.CheckBoxChecked;
            // 
            // ExtraTextBox2
            // 
            this.ExtraTextBox2.Enabled = false;
            this.ExtraTextBox2.Location = new Point(166, 141);
            this.ExtraTextBox2.Margin = new Padding(0);
            this.ExtraTextBox2.MaxLength = 20;
            this.ExtraTextBox2.Name = "ExtraTextBox2";
            this.ExtraTextBox2.Size = new Size(120, 23);
            this.ExtraTextBox2.TabIndex = 98;
            this.ExtraTextBox2.TextAlign = HorizontalAlignment.Right;
            // 
            // ExtraCheckBox2
            // 
            this.ExtraCheckBox2.Location = new Point(166, 118);
            this.ExtraCheckBox2.Margin = new Padding(0);
            this.ExtraCheckBox2.Name = "ExtraCheckBox2";
            this.ExtraCheckBox2.Size = new Size(50, 19);
            this.ExtraCheckBox2.TabIndex = 97;
            this.ExtraCheckBox2.TabStop = false;
            this.ExtraCheckBox2.Text = "予備";
            this.ExtraCheckBox2.UseVisualStyleBackColor = true;
            this.ExtraCheckBox2.Visible = false;
            this.ExtraCheckBox2.CheckedChanged += this.CheckBoxChecked;
            // 
            // RevisionTextBox
            // 
            this.RevisionTextBox.Enabled = false;
            this.RevisionTextBox.Location = new Point(19, 341);
            this.RevisionTextBox.Margin = new Padding(0);
            this.RevisionTextBox.MaxLength = 4;
            this.RevisionTextBox.Name = "RevisionTextBox";
            this.RevisionTextBox.Size = new Size(120, 23);
            this.RevisionTextBox.TabIndex = 96;
            this.RevisionTextBox.TextAlign = HorizontalAlignment.Right;
            // 
            // RevisionCheckBox
            // 
            this.RevisionCheckBox.Location = new Point(19, 318);
            this.RevisionCheckBox.Margin = new Padding(0);
            this.RevisionCheckBox.Name = "RevisionCheckBox";
            this.RevisionCheckBox.Size = new Size(68, 19);
            this.RevisionCheckBox.TabIndex = 95;
            this.RevisionCheckBox.TabStop = false;
            this.RevisionCheckBox.Text = "レビジョン";
            this.RevisionCheckBox.UseVisualStyleBackColor = true;
            this.RevisionCheckBox.CheckedChanged += this.CheckBoxChecked;
            // 
            // ExtraTextBox1
            // 
            this.ExtraTextBox1.Enabled = false;
            this.ExtraTextBox1.Location = new Point(19, 291);
            this.ExtraTextBox1.Margin = new Padding(0);
            this.ExtraTextBox1.MaxLength = 4;
            this.ExtraTextBox1.Name = "ExtraTextBox1";
            this.ExtraTextBox1.Size = new Size(120, 23);
            this.ExtraTextBox1.TabIndex = 94;
            this.ExtraTextBox1.TextAlign = HorizontalAlignment.Right;
            // 
            // ExtraCheckBox1
            // 
            this.ExtraCheckBox1.Location = new Point(19, 268);
            this.ExtraCheckBox1.Margin = new Padding(0);
            this.ExtraCheckBox1.Name = "ExtraCheckBox1";
            this.ExtraCheckBox1.Size = new Size(62, 19);
            this.ExtraCheckBox1.TabIndex = 93;
            this.ExtraCheckBox1.TabStop = false;
            this.ExtraCheckBox1.Text = "予備";
            this.ExtraCheckBox1.UseVisualStyleBackColor = true;
            this.ExtraCheckBox1.Visible = false;
            this.ExtraCheckBox1.CheckedChanged += this.CheckBoxChecked;
            // 
            // QuantityTextBox
            // 
            this.QuantityTextBox.Enabled = false;
            this.QuantityTextBox.Location = new Point(19, 241);
            this.QuantityTextBox.Margin = new Padding(0);
            this.QuantityTextBox.MaxLength = 4;
            this.QuantityTextBox.Name = "QuantityTextBox";
            this.QuantityTextBox.Size = new Size(120, 23);
            this.QuantityTextBox.TabIndex = 92;
            this.QuantityTextBox.TextAlign = HorizontalAlignment.Right;
            this.QuantityTextBox.KeyPress += this.NumericOnly;
            // 
            // QuantityCheckBox
            // 
            this.QuantityCheckBox.Location = new Point(19, 218);
            this.QuantityCheckBox.Margin = new Padding(0);
            this.QuantityCheckBox.Name = "QuantityCheckBox";
            this.QuantityCheckBox.Size = new Size(50, 19);
            this.QuantityCheckBox.TabIndex = 91;
            this.QuantityCheckBox.TabStop = false;
            this.QuantityCheckBox.Text = "数量";
            this.QuantityCheckBox.UseVisualStyleBackColor = true;
            this.QuantityCheckBox.CheckedChanged += this.CheckBoxChecked;
            // 
            // ManufacturingNumberMaskedTextBox
            // 
            this.ManufacturingNumberMaskedTextBox.Location = new Point(19, 191);
            this.ManufacturingNumberMaskedTextBox.Mask = ">LA00L00000-0000";
            this.ManufacturingNumberMaskedTextBox.Name = "ManufacturingNumberMaskedTextBox";
            this.ManufacturingNumberMaskedTextBox.PromptChar = '*';
            this.ManufacturingNumberMaskedTextBox.Size = new Size(120, 23);
            this.ManufacturingNumberMaskedTextBox.TabIndex = 90;
            this.ManufacturingNumberMaskedTextBox.Text = "H";
            this.ManufacturingNumberMaskedTextBox.TextAlign = HorizontalAlignment.Right;
            // 
            // ManufacturingNumberCheckBox
            // 
            this.ManufacturingNumberCheckBox.Location = new Point(19, 168);
            this.ManufacturingNumberCheckBox.Margin = new Padding(0);
            this.ManufacturingNumberCheckBox.Name = "ManufacturingNumberCheckBox";
            this.ManufacturingNumberCheckBox.Size = new Size(74, 19);
            this.ManufacturingNumberCheckBox.TabIndex = 89;
            this.ManufacturingNumberCheckBox.TabStop = false;
            this.ManufacturingNumberCheckBox.Text = "製造番号";
            this.ManufacturingNumberCheckBox.UseVisualStyleBackColor = true;
            this.ManufacturingNumberCheckBox.CheckedChanged += this.CheckBoxChecked;
            // 
            // OrderNumberTextBox
            // 
            this.OrderNumberTextBox.Enabled = false;
            this.OrderNumberTextBox.Location = new Point(19, 141);
            this.OrderNumberTextBox.Margin = new Padding(0);
            this.OrderNumberTextBox.MaxLength = 20;
            this.OrderNumberTextBox.Name = "OrderNumberTextBox";
            this.OrderNumberTextBox.Size = new Size(120, 23);
            this.OrderNumberTextBox.TabIndex = 88;
            this.OrderNumberTextBox.TextAlign = HorizontalAlignment.Right;
            // 
            // OrderNumberCheckBox
            // 
            this.OrderNumberCheckBox.Location = new Point(19, 118);
            this.OrderNumberCheckBox.Margin = new Padding(0);
            this.OrderNumberCheckBox.Name = "OrderNumberCheckBox";
            this.OrderNumberCheckBox.Size = new Size(74, 19);
            this.OrderNumberCheckBox.TabIndex = 87;
            this.OrderNumberCheckBox.TabStop = false;
            this.OrderNumberCheckBox.Text = "注文番号";
            this.OrderNumberCheckBox.UseVisualStyleBackColor = true;
            this.OrderNumberCheckBox.CheckedChanged += this.CheckBoxChecked;
            // 
            // SubstrateModelLabel2
            // 
            this.SubstrateModelLabel2.AutoSize = true;
            this.SubstrateModelLabel2.BorderStyle = BorderStyle.Fixed3D;
            this.SubstrateModelLabel2.Location = new Point(90, 85);
            this.SubstrateModelLabel2.Margin = new Padding(4, 0, 4, 0);
            this.SubstrateModelLabel2.Name = "SubstrateModelLabel2";
            this.SubstrateModelLabel2.Size = new Size(24, 17);
            this.SubstrateModelLabel2.TabIndex = 86;
            this.SubstrateModelLabel2.Text = "---";
            this.SubstrateModelLabel2.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // SubstrateModelLabel1
            // 
            this.SubstrateModelLabel1.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            this.SubstrateModelLabel1.AutoSize = true;
            this.SubstrateModelLabel1.BorderStyle = BorderStyle.Fixed3D;
            this.SubstrateModelLabel1.Location = new Point(16, 85);
            this.SubstrateModelLabel1.Margin = new Padding(4, 0, 4, 0);
            this.SubstrateModelLabel1.Name = "SubstrateModelLabel1";
            this.SubstrateModelLabel1.Size = new Size(66, 17);
            this.SubstrateModelLabel1.TabIndex = 85;
            this.SubstrateModelLabel1.Text = "製品型式 :";
            this.SubstrateModelLabel1.TextAlign = ContentAlignment.MiddleRight;
            // 
            // ProductNameLabel2
            // 
            this.ProductNameLabel2.AutoSize = true;
            this.ProductNameLabel2.BorderStyle = BorderStyle.Fixed3D;
            this.ProductNameLabel2.Location = new Point(90, 39);
            this.ProductNameLabel2.Margin = new Padding(4, 0, 4, 0);
            this.ProductNameLabel2.Name = "ProductNameLabel2";
            this.ProductNameLabel2.Size = new Size(24, 17);
            this.ProductNameLabel2.TabIndex = 84;
            this.ProductNameLabel2.Text = "---";
            this.ProductNameLabel2.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // ProductNameLabel1
            // 
            this.ProductNameLabel1.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            this.ProductNameLabel1.AutoSize = true;
            this.ProductNameLabel1.BorderStyle = BorderStyle.Fixed3D;
            this.ProductNameLabel1.Location = new Point(28, 39);
            this.ProductNameLabel1.Margin = new Padding(4, 0, 4, 0);
            this.ProductNameLabel1.Name = "ProductNameLabel1";
            this.ProductNameLabel1.Size = new Size(54, 17);
            this.ProductNameLabel1.TabIndex = 83;
            this.ProductNameLabel1.Text = "製品名 :";
            this.ProductNameLabel1.TextAlign = ContentAlignment.MiddleRight;
            // 
            // RePrintMenuStrip
            // 
            this.RePrintMenuStrip.Font = new Font("Meiryo UI", 9F);
            this.RePrintMenuStrip.Items.AddRange(new ToolStripItem[] { this.ファイルToolStripMenuItem, this.設定ToolStripMenuItem, this.ヘルプToolStripMenuItem });
            this.RePrintMenuStrip.Location = new Point(0, 0);
            this.RePrintMenuStrip.Name = "RePrintMenuStrip";
            this.RePrintMenuStrip.Size = new Size(630, 24);
            this.RePrintMenuStrip.TabIndex = 82;
            this.RePrintMenuStrip.Text = "menuStrip";
            // 
            // ファイルToolStripMenuItem
            // 
            this.ファイルToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { this.シリアルラベル印刷プレビューToolStripMenuItem, this.バーコード印刷プレビューToolStripMenuItem, this.終了ToolStripMenuItem });
            this.ファイルToolStripMenuItem.Name = "ファイルToolStripMenuItem";
            this.ファイルToolStripMenuItem.Size = new Size(53, 20);
            this.ファイルToolStripMenuItem.Text = "ファイル";
            // 
            // シリアルラベル印刷プレビューToolStripMenuItem
            // 
            this.シリアルラベル印刷プレビューToolStripMenuItem.Name = "シリアルラベル印刷プレビューToolStripMenuItem";
            this.シリアルラベル印刷プレビューToolStripMenuItem.Size = new Size(206, 22);
            this.シリアルラベル印刷プレビューToolStripMenuItem.Text = "シリアルラベル印刷プレビュー";
            this.シリアルラベル印刷プレビューToolStripMenuItem.Click += this.シリアルラベル印刷プレビューToolStripMenuItem_Click;
            // 
            // バーコード印刷プレビューToolStripMenuItem
            // 
            this.バーコード印刷プレビューToolStripMenuItem.Name = "バーコード印刷プレビューToolStripMenuItem";
            this.バーコード印刷プレビューToolStripMenuItem.Size = new Size(206, 22);
            this.バーコード印刷プレビューToolStripMenuItem.Text = "バーコード印刷プレビュー";
            this.バーコード印刷プレビューToolStripMenuItem.Click += this.バーコード印刷プレビューToolStripMenuItem_Click;
            // 
            // 終了ToolStripMenuItem
            // 
            this.終了ToolStripMenuItem.Name = "終了ToolStripMenuItem";
            this.終了ToolStripMenuItem.Size = new Size(206, 22);
            this.終了ToolStripMenuItem.Text = "終了";
            // 
            // 設定ToolStripMenuItem
            // 
            this.設定ToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { this.シリアルラベル印刷設定ToolStripMenuItem, this.バーコード印刷設定ToolStripMenuItem });
            this.設定ToolStripMenuItem.Name = "設定ToolStripMenuItem";
            this.設定ToolStripMenuItem.Size = new Size(43, 20);
            this.設定ToolStripMenuItem.Text = "設定";
            // 
            // シリアルラベル印刷設定ToolStripMenuItem
            // 
            this.シリアルラベル印刷設定ToolStripMenuItem.Name = "シリアルラベル印刷設定ToolStripMenuItem";
            this.シリアルラベル印刷設定ToolStripMenuItem.Size = new Size(186, 22);
            this.シリアルラベル印刷設定ToolStripMenuItem.Text = "シリアルラベル印刷設定";
            this.シリアルラベル印刷設定ToolStripMenuItem.Click += this.シリアルラベル印刷設定ToolStripMenuItem_Click;
            // 
            // バーコード印刷設定ToolStripMenuItem
            // 
            this.バーコード印刷設定ToolStripMenuItem.Name = "バーコード印刷設定ToolStripMenuItem";
            this.バーコード印刷設定ToolStripMenuItem.Size = new Size(186, 22);
            this.バーコード印刷設定ToolStripMenuItem.Text = "バーコード印刷設定";
            this.バーコード印刷設定ToolStripMenuItem.Click += this.バーコード印刷設定ToolStripMenuItem_Click;
            // 
            // ヘルプToolStripMenuItem
            // 
            this.ヘルプToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { this.取得情報ToolStripMenuItem });
            this.ヘルプToolStripMenuItem.Name = "ヘルプToolStripMenuItem";
            this.ヘルプToolStripMenuItem.Size = new Size(48, 20);
            this.ヘルプToolStripMenuItem.Text = "ヘルプ";
            // 
            // 取得情報ToolStripMenuItem
            // 
            this.取得情報ToolStripMenuItem.Name = "取得情報ToolStripMenuItem";
            this.取得情報ToolStripMenuItem.Size = new Size(122, 22);
            this.取得情報ToolStripMenuItem.Text = "取得情報";
            this.取得情報ToolStripMenuItem.Click += this.取得情報ToolStripMenuItem_Click;
            // 
            // BarcodePrintButton
            // 
            this.BarcodePrintButton.Location = new Point(450, 404);
            this.BarcodePrintButton.Margin = new Padding(0);
            this.BarcodePrintButton.Name = "BarcodePrintButton";
            this.BarcodePrintButton.Size = new Size(100, 25);
            this.BarcodePrintButton.TabIndex = 121;
            this.BarcodePrintButton.Text = "バーコード印刷";
            this.BarcodePrintButton.UseVisualStyleBackColor = true;
            this.BarcodePrintButton.Click += this.BarcodePrintButton_Click;
            // 
            // LabelPrintButton
            // 
            this.LabelPrintButton.Location = new Point(313, 404);
            this.LabelPrintButton.Margin = new Padding(0);
            this.LabelPrintButton.Name = "LabelPrintButton";
            this.LabelPrintButton.Size = new Size(100, 25);
            this.LabelPrintButton.TabIndex = 119;
            this.LabelPrintButton.Text = "ラベル印刷";
            this.LabelPrintButton.UseVisualStyleBackColor = true;
            this.LabelPrintButton.Click += this.LabelPrintButton_Click;
            // 
            // PrintPositionNumericUpDown
            // 
            this.PrintPositionNumericUpDown.Location = new Point(166, 406);
            this.PrintPositionNumericUpDown.Margin = new Padding(0);
            this.PrintPositionNumericUpDown.Maximum = new decimal(new int[] { 50, 0, 0, 0 });
            this.PrintPositionNumericUpDown.Minimum = new decimal(new int[] { 1, 0, 0, 0 });
            this.PrintPositionNumericUpDown.Name = "PrintPositionNumericUpDown";
            this.PrintPositionNumericUpDown.Size = new Size(110, 23);
            this.PrintPositionNumericUpDown.TabIndex = 118;
            this.PrintPositionNumericUpDown.TextAlign = HorizontalAlignment.Right;
            this.PrintPositionNumericUpDown.Value = new decimal(new int[] { 1, 0, 0, 0 });
            // 
            // PrintRowLabel
            // 
            this.PrintRowLabel.Location = new Point(178, 391);
            this.PrintRowLabel.Margin = new Padding(0);
            this.PrintRowLabel.Name = "PrintRowLabel";
            this.PrintRowLabel.Size = new Size(79, 15);
            this.PrintRowLabel.TabIndex = 117;
            this.PrintRowLabel.Text = "印刷開始位置";
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.CommentTextBox);
            this.panel1.Location = new Point(464, 141);
            this.panel1.Name = "panel1";
            this.panel1.Size = new Size(150, 175);
            this.panel1.TabIndex = 122;
            // 
            // QrCodeTextBox
            // 
            this.QrCodeTextBox.Location = new Point(298, 39);
            this.QrCodeTextBox.MaxLength = 100;
            this.QrCodeTextBox.Name = "QrCodeTextBox";
            this.QrCodeTextBox.Size = new Size(320, 23);
            this.QrCodeTextBox.TabIndex = 123;
            this.QrCodeTextBox.Enter += this.QrCodeTextBox_Enter;
            // 
            // ProductTypeLabel2
            // 
            this.ProductTypeLabel2.AutoSize = true;
            this.ProductTypeLabel2.BorderStyle = BorderStyle.Fixed3D;
            this.ProductTypeLabel2.Location = new Point(90, 62);
            this.ProductTypeLabel2.Margin = new Padding(4, 0, 4, 0);
            this.ProductTypeLabel2.Name = "ProductTypeLabel2";
            this.ProductTypeLabel2.Size = new Size(24, 17);
            this.ProductTypeLabel2.TabIndex = 125;
            this.ProductTypeLabel2.Text = "---";
            this.ProductTypeLabel2.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // ProductTypeLabel1
            // 
            this.ProductTypeLabel1.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            this.ProductTypeLabel1.AutoSize = true;
            this.ProductTypeLabel1.BorderStyle = BorderStyle.Fixed3D;
            this.ProductTypeLabel1.Location = new Point(38, 62);
            this.ProductTypeLabel1.Margin = new Padding(4, 0, 4, 0);
            this.ProductTypeLabel1.Name = "ProductTypeLabel1";
            this.ProductTypeLabel1.Size = new Size(44, 17);
            this.ProductTypeLabel1.TabIndex = 124;
            this.ProductTypeLabel1.Text = "タイプ :";
            this.ProductTypeLabel1.TextAlign = ContentAlignment.MiddleRight;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new Point(298, 24);
            this.label1.Name = "label1";
            this.label1.Size = new Size(86, 15);
            this.label1.TabIndex = 126;
            this.label1.Text = "QRコード入力用";
            // 
            // QrCodeButton
            // 
            this.QrCodeButton.Location = new Point(543, 68);
            this.QrCodeButton.Name = "QrCodeButton";
            this.QrCodeButton.Size = new Size(75, 23);
            this.QrCodeButton.TabIndex = 127;
            this.QrCodeButton.Text = "入力";
            this.QrCodeButton.UseVisualStyleBackColor = true;
            this.QrCodeButton.Click += this.QrCodeButton_Click;
            // 
            // textToUpperCheckBox
            // 
            this.textToUpperCheckBox.AutoSize = true;
            this.textToUpperCheckBox.Checked = true;
            this.textToUpperCheckBox.CheckState = CheckState.Checked;
            this.textToUpperCheckBox.Location = new Point(397, 71);
            this.textToUpperCheckBox.Name = "textToUpperCheckBox";
            this.textToUpperCheckBox.Size = new Size(140, 19);
            this.textToUpperCheckBox.TabIndex = 609;
            this.textToUpperCheckBox.Text = "小文字を大文字に変換";
            this.textToUpperCheckBox.UseVisualStyleBackColor = true;
            // 
            // RegistrationDateTimePicker
            // 
            this.RegistrationDateTimePicker.Format = DateTimePickerFormat.Short;
            this.RegistrationDateTimePicker.Location = new Point(166, 291);
            this.RegistrationDateTimePicker.Name = "RegistrationDateTimePicker";
            this.RegistrationDateTimePicker.Size = new Size(120, 23);
            this.RegistrationDateTimePicker.TabIndex = 611;
            // 
            // RePrintWindow
            // 
            this.AutoScaleDimensions = new SizeF(7F, 15F);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.ClientSize = new Size(630, 441);
            this.Controls.Add(this.RegistrationDateTimePicker);
            this.Controls.Add(this.textToUpperCheckBox);
            this.Controls.Add(this.QrCodeButton);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.ProductTypeLabel2);
            this.Controls.Add(this.ProductTypeLabel1);
            this.Controls.Add(this.QrCodeTextBox);
            this.Controls.Add(this.BarcodePrintButton);
            this.Controls.Add(this.LabelPrintButton);
            this.Controls.Add(this.PrintPositionNumericUpDown);
            this.Controls.Add(this.PrintRowLabel);
            this.Controls.Add(this.TemplateButton);
            this.Controls.Add(this.CommentComboBox);
            this.Controls.Add(this.CommentCheckBox);
            this.Controls.Add(this.ExtraTextBox6);
            this.Controls.Add(this.ExtraCheckBox6);
            this.Controls.Add(this.ExtraTextBox5);
            this.Controls.Add(this.ExtraCheckBox5);
            this.Controls.Add(this.ExtraTextBox4);
            this.Controls.Add(this.ExtraCheckBox4);
            this.Controls.Add(this.PersonComboBox);
            this.Controls.Add(this.PersonCheckBox);
            this.Controls.Add(this.RegistrationDateCheckBox);
            this.Controls.Add(this.FirstSerialNumberTextBox);
            this.Controls.Add(this.FirstSerialNumberCheckBox);
            this.Controls.Add(this.ExtraTextBox3);
            this.Controls.Add(this.ExtraCheckBox3);
            this.Controls.Add(this.ExtraTextBox2);
            this.Controls.Add(this.ExtraCheckBox2);
            this.Controls.Add(this.RevisionTextBox);
            this.Controls.Add(this.RevisionCheckBox);
            this.Controls.Add(this.ExtraTextBox1);
            this.Controls.Add(this.ExtraCheckBox1);
            this.Controls.Add(this.QuantityTextBox);
            this.Controls.Add(this.QuantityCheckBox);
            this.Controls.Add(this.ManufacturingNumberMaskedTextBox);
            this.Controls.Add(this.ManufacturingNumberCheckBox);
            this.Controls.Add(this.OrderNumberTextBox);
            this.Controls.Add(this.OrderNumberCheckBox);
            this.Controls.Add(this.SubstrateModelLabel2);
            this.Controls.Add(this.SubstrateModelLabel1);
            this.Controls.Add(this.ProductNameLabel2);
            this.Controls.Add(this.ProductNameLabel1);
            this.Controls.Add(this.RePrintMenuStrip);
            this.Controls.Add(this.panel1);
            this.Font = new Font("Meiryo UI", 9F);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "RePrintWindow";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = FormStartPosition.CenterParent;
            this.Text = "RePrint";
            this.Load += this.RePrintWindow_Load;
            this.RePrintMenuStrip.ResumeLayout(false);
            this.RePrintMenuStrip.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)this.PrintPositionNumericUpDown).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();
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
        private NumericUpDown PrintPositionNumericUpDown;
        private Label PrintRowLabel;
        private Panel panel1;
        private ToolStripMenuItem 設定ToolStripMenuItem;
        private ToolStripMenuItem シリアルラベル印刷設定ToolStripMenuItem;
        private ToolStripMenuItem バーコード印刷設定ToolStripMenuItem;
        private ToolStripMenuItem シリアルラベル印刷プレビューToolStripMenuItem;
        private ToolStripMenuItem バーコード印刷プレビューToolStripMenuItem;
        private TextBox QrCodeTextBox;
        private Label ProductTypeLabel2;
        private Label ProductTypeLabel1;
        private Label label1;
        private Button QrCodeButton;
        private CheckBox textToUpperCheckBox;
        private DateTimePicker RegistrationDateTimePicker;
    }
}