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
            var resources = new System.ComponentModel.ComponentResourceManager(typeof(SubstrateRegistrationWindow));
            this.SubstrateRegistrationMenuStrip = new MenuStrip();
            this.ファイルToolStripMenuItem = new ToolStripMenuItem();
            this.印刷ToolStripMenuItem = new ToolStripMenuItem();
            this.印刷プレビューToolStripMenuItem = new ToolStripMenuItem();
            this.終了ToolStripMenuItem = new ToolStripMenuItem();
            this.設定ToolStripMenuItem = new ToolStripMenuItem();
            this.印刷設定ToolStripMenuItem = new ToolStripMenuItem();
            this.ヘルプToolStripMenuItem = new ToolStripMenuItem();
            this.取得情報ToolStripMenuItem = new ToolStripMenuItem();
            this.ProductNameLabel1 = new Label();
            this.ProductNameLabel2 = new Label();
            this.SubstrateModelLabel1 = new Label();
            this.SubstrateModelLabel2 = new Label();
            this.OrderNumberCheckBox = new CheckBox();
            this.OrderNumberTextBox = new TextBox();
            this.ManufacturingNumberCheckBox = new CheckBox();
            this.ManufacturingNumberMaskedTextBox = new MaskedTextBox();
            this.QuantityTextBox = new TextBox();
            this.QuantityCheckBox = new CheckBox();
            this.DefectNumberTextBox = new TextBox();
            this.DefectNumberCheckBox = new CheckBox();
            this.RevisionTextBox = new TextBox();
            this.ExtraCheckBox7 = new CheckBox();
            this.ExtraTextBox1 = new TextBox();
            this.ExtraCheckBox1 = new CheckBox();
            this.ExtraTextBox2 = new TextBox();
            this.ExtraCheckBox2 = new CheckBox();
            this.ExtraTextBox3 = new TextBox();
            this.ExtraCheckBox3 = new CheckBox();
            this.RegistrationDateCheckBox = new CheckBox();
            this.RegistrationDateMaskedTextBox = new MaskedTextBox();
            this.PersonCheckBox = new CheckBox();
            this.PersonComboBox = new ComboBox();
            this.ExtraTextBox4 = new TextBox();
            this.ExtraCheckBox4 = new CheckBox();
            this.ExtraTextBox5 = new TextBox();
            this.ExtraCheckBox5 = new CheckBox();
            this.ExtraTextBox6 = new TextBox();
            this.ExtraCheckBox6 = new CheckBox();
            this.CommentCheckBox = new CheckBox();
            this.CommentTextBox = new TextBox();
            this.CommentComboBox = new ComboBox();
            this.TemplateButton = new Button();
            this.PrintRowLabel = new Label();
            this.PrintPostionNumericUpDown = new NumericUpDown();
            this.RegisterButton = new Button();
            this.PrintOnlyCheckBox = new CheckBox();
            this.PrintButton = new Button();
            this.SubstrateRegistrationPrintDialog = new PrintDialog();
            this.SubstrateRegistrationPrintDocument = new System.Drawing.Printing.PrintDocument();
            this.SubstrateRegistrationPrintPreviewDialog = new PrintPreviewDialog();
            this.panelCommentTexrBox = new Panel();
            this.label1 = new Label();
            this.QrCodeButton = new Button();
            this.QrCodeTextBox = new TextBox();
            this.textToUpperCheckBox = new CheckBox();
            this.SubstrateRegistrationMenuStrip.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)this.PrintPostionNumericUpDown).BeginInit();
            this.panelCommentTexrBox.SuspendLayout();
            this.SuspendLayout();
            // 
            // SubstrateRegistrationMenuStrip
            // 
            this.SubstrateRegistrationMenuStrip.Font = new Font("Meiryo UI", 9F);
            this.SubstrateRegistrationMenuStrip.Items.AddRange(new ToolStripItem[] { this.ファイルToolStripMenuItem, this.設定ToolStripMenuItem, this.ヘルプToolStripMenuItem });
            this.SubstrateRegistrationMenuStrip.Location = new Point(0, 0);
            this.SubstrateRegistrationMenuStrip.Name = "SubstrateRegistrationMenuStrip";
            this.SubstrateRegistrationMenuStrip.Size = new Size(630, 24);
            this.SubstrateRegistrationMenuStrip.TabIndex = 0;
            this.SubstrateRegistrationMenuStrip.Text = "menuStrip";
            // 
            // ファイルToolStripMenuItem
            // 
            this.ファイルToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { this.印刷ToolStripMenuItem, this.印刷プレビューToolStripMenuItem, this.終了ToolStripMenuItem });
            this.ファイルToolStripMenuItem.Name = "ファイルToolStripMenuItem";
            this.ファイルToolStripMenuItem.Size = new Size(53, 20);
            this.ファイルToolStripMenuItem.Text = "ファイル";
            // 
            // 印刷ToolStripMenuItem
            // 
            this.印刷ToolStripMenuItem.Name = "印刷ToolStripMenuItem";
            this.印刷ToolStripMenuItem.Size = new Size(142, 22);
            this.印刷ToolStripMenuItem.Text = "印刷";
            this.印刷ToolStripMenuItem.Click += this.印刷ToolStripMenuItem_Click;
            // 
            // 印刷プレビューToolStripMenuItem
            // 
            this.印刷プレビューToolStripMenuItem.Name = "印刷プレビューToolStripMenuItem";
            this.印刷プレビューToolStripMenuItem.Size = new Size(142, 22);
            this.印刷プレビューToolStripMenuItem.Text = "印刷プレビュー";
            this.印刷プレビューToolStripMenuItem.Click += this.印刷プレビューToolStripMenuItem_Click;
            // 
            // 終了ToolStripMenuItem
            // 
            this.終了ToolStripMenuItem.Name = "終了ToolStripMenuItem";
            this.終了ToolStripMenuItem.Size = new Size(142, 22);
            this.終了ToolStripMenuItem.Text = "終了";
            // 
            // 設定ToolStripMenuItem
            // 
            this.設定ToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { this.印刷設定ToolStripMenuItem });
            this.設定ToolStripMenuItem.Name = "設定ToolStripMenuItem";
            this.設定ToolStripMenuItem.Size = new Size(43, 20);
            this.設定ToolStripMenuItem.Text = "設定";
            // 
            // 印刷設定ToolStripMenuItem
            // 
            this.印刷設定ToolStripMenuItem.Name = "印刷設定ToolStripMenuItem";
            this.印刷設定ToolStripMenuItem.Size = new Size(122, 22);
            this.印刷設定ToolStripMenuItem.Text = "印刷設定";
            this.印刷設定ToolStripMenuItem.Click += this.印刷設定ToolStripMenuItem_Click;
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
            // ProductNameLabel1
            // 
            this.ProductNameLabel1.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            this.ProductNameLabel1.AutoSize = true;
            this.ProductNameLabel1.BorderStyle = BorderStyle.Fixed3D;
            this.ProductNameLabel1.Location = new Point(28, 39);
            this.ProductNameLabel1.Margin = new Padding(4, 0, 4, 0);
            this.ProductNameLabel1.Name = "ProductNameLabel1";
            this.ProductNameLabel1.Size = new Size(54, 17);
            this.ProductNameLabel1.TabIndex = 1;
            this.ProductNameLabel1.Text = "製品名 :";
            this.ProductNameLabel1.TextAlign = ContentAlignment.MiddleRight;
            // 
            // ProductNameLabel2
            // 
            this.ProductNameLabel2.AutoSize = true;
            this.ProductNameLabel2.BorderStyle = BorderStyle.Fixed3D;
            this.ProductNameLabel2.Location = new Point(90, 39);
            this.ProductNameLabel2.Margin = new Padding(4, 0, 4, 0);
            this.ProductNameLabel2.Name = "ProductNameLabel2";
            this.ProductNameLabel2.Size = new Size(24, 17);
            this.ProductNameLabel2.TabIndex = 2;
            this.ProductNameLabel2.Text = "---";
            this.ProductNameLabel2.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // SubstrateModelLabel1
            // 
            this.SubstrateModelLabel1.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            this.SubstrateModelLabel1.AutoSize = true;
            this.SubstrateModelLabel1.BorderStyle = BorderStyle.Fixed3D;
            this.SubstrateModelLabel1.Location = new Point(16, 62);
            this.SubstrateModelLabel1.Margin = new Padding(4, 0, 4, 0);
            this.SubstrateModelLabel1.Name = "SubstrateModelLabel1";
            this.SubstrateModelLabel1.Size = new Size(66, 17);
            this.SubstrateModelLabel1.TabIndex = 3;
            this.SubstrateModelLabel1.Text = "基板型式 :";
            this.SubstrateModelLabel1.TextAlign = ContentAlignment.MiddleRight;
            // 
            // SubstrateModelLabel2
            // 
            this.SubstrateModelLabel2.AutoSize = true;
            this.SubstrateModelLabel2.BorderStyle = BorderStyle.Fixed3D;
            this.SubstrateModelLabel2.Location = new Point(90, 62);
            this.SubstrateModelLabel2.Margin = new Padding(4, 0, 4, 0);
            this.SubstrateModelLabel2.Name = "SubstrateModelLabel2";
            this.SubstrateModelLabel2.Size = new Size(24, 17);
            this.SubstrateModelLabel2.TabIndex = 4;
            this.SubstrateModelLabel2.Text = "---";
            this.SubstrateModelLabel2.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // OrderNumberCheckBox
            // 
            this.OrderNumberCheckBox.Location = new Point(19, 118);
            this.OrderNumberCheckBox.Margin = new Padding(0);
            this.OrderNumberCheckBox.Name = "OrderNumberCheckBox";
            this.OrderNumberCheckBox.Size = new Size(74, 19);
            this.OrderNumberCheckBox.TabIndex = 7;
            this.OrderNumberCheckBox.TabStop = false;
            this.OrderNumberCheckBox.Text = "注文番号";
            this.OrderNumberCheckBox.UseVisualStyleBackColor = true;
            this.OrderNumberCheckBox.CheckedChanged += this.NumberCheckBox_CheckedChanged;
            // 
            // OrderNumberTextBox
            // 
            this.OrderNumberTextBox.Enabled = false;
            this.OrderNumberTextBox.Location = new Point(19, 141);
            this.OrderNumberTextBox.Margin = new Padding(0);
            this.OrderNumberTextBox.MaxLength = 20;
            this.OrderNumberTextBox.Name = "OrderNumberTextBox";
            this.OrderNumberTextBox.Size = new Size(120, 23);
            this.OrderNumberTextBox.TabIndex = 8;
            this.OrderNumberTextBox.TextAlign = HorizontalAlignment.Right;
            // 
            // ManufacturingNumberCheckBox
            // 
            this.ManufacturingNumberCheckBox.Location = new Point(19, 168);
            this.ManufacturingNumberCheckBox.Margin = new Padding(0);
            this.ManufacturingNumberCheckBox.Name = "ManufacturingNumberCheckBox";
            this.ManufacturingNumberCheckBox.Size = new Size(74, 19);
            this.ManufacturingNumberCheckBox.TabIndex = 9;
            this.ManufacturingNumberCheckBox.TabStop = false;
            this.ManufacturingNumberCheckBox.Text = "製造番号";
            this.ManufacturingNumberCheckBox.UseVisualStyleBackColor = true;
            this.ManufacturingNumberCheckBox.CheckedChanged += this.CheckBoxChecked;
            // 
            // ManufacturingNumberMaskedTextBox
            // 
            this.ManufacturingNumberMaskedTextBox.Enabled = false;
            this.ManufacturingNumberMaskedTextBox.Location = new Point(19, 191);
            this.ManufacturingNumberMaskedTextBox.Mask = ">LA00L00000-0000";
            this.ManufacturingNumberMaskedTextBox.Name = "ManufacturingNumberMaskedTextBox";
            this.ManufacturingNumberMaskedTextBox.PromptChar = '*';
            this.ManufacturingNumberMaskedTextBox.Size = new Size(120, 23);
            this.ManufacturingNumberMaskedTextBox.TabIndex = 10;
            this.ManufacturingNumberMaskedTextBox.Text = "H";
            this.ManufacturingNumberMaskedTextBox.TextAlign = HorizontalAlignment.Right;
            // 
            // QuantityTextBox
            // 
            this.QuantityTextBox.Enabled = false;
            this.QuantityTextBox.Location = new Point(19, 241);
            this.QuantityTextBox.Margin = new Padding(0);
            this.QuantityTextBox.MaxLength = 4;
            this.QuantityTextBox.Name = "QuantityTextBox";
            this.QuantityTextBox.Size = new Size(120, 23);
            this.QuantityTextBox.TabIndex = 12;
            this.QuantityTextBox.TextAlign = HorizontalAlignment.Right;
            this.QuantityTextBox.KeyPress += this.QuantityTextBox_KeyPress;
            // 
            // QuantityCheckBox
            // 
            this.QuantityCheckBox.Location = new Point(19, 218);
            this.QuantityCheckBox.Margin = new Padding(0);
            this.QuantityCheckBox.Name = "QuantityCheckBox";
            this.QuantityCheckBox.Size = new Size(62, 19);
            this.QuantityCheckBox.TabIndex = 11;
            this.QuantityCheckBox.TabStop = false;
            this.QuantityCheckBox.Text = "追加量";
            this.QuantityCheckBox.UseVisualStyleBackColor = true;
            this.QuantityCheckBox.CheckedChanged += this.CheckBoxChecked;
            // 
            // DefectNumberTextBox
            // 
            this.DefectNumberTextBox.Enabled = false;
            this.DefectNumberTextBox.Location = new Point(19, 291);
            this.DefectNumberTextBox.Margin = new Padding(0);
            this.DefectNumberTextBox.MaxLength = 4;
            this.DefectNumberTextBox.Name = "DefectNumberTextBox";
            this.DefectNumberTextBox.Size = new Size(120, 23);
            this.DefectNumberTextBox.TabIndex = 14;
            this.DefectNumberTextBox.TextAlign = HorizontalAlignment.Right;
            this.DefectNumberTextBox.KeyPress += this.DefectNumberTextBox_KeyPress;
            // 
            // DefectNumberCheckBox
            // 
            this.DefectNumberCheckBox.Location = new Point(19, 268);
            this.DefectNumberCheckBox.Margin = new Padding(0);
            this.DefectNumberCheckBox.Name = "DefectNumberCheckBox";
            this.DefectNumberCheckBox.Size = new Size(62, 19);
            this.DefectNumberCheckBox.TabIndex = 13;
            this.DefectNumberCheckBox.TabStop = false;
            this.DefectNumberCheckBox.Text = "減少量";
            this.DefectNumberCheckBox.UseVisualStyleBackColor = true;
            this.DefectNumberCheckBox.CheckedChanged += this.CheckBoxChecked;
            // 
            // RevisionTextBox
            // 
            this.RevisionTextBox.Enabled = false;
            this.RevisionTextBox.Location = new Point(19, 341);
            this.RevisionTextBox.Margin = new Padding(0);
            this.RevisionTextBox.MaxLength = 4;
            this.RevisionTextBox.Name = "RevisionTextBox";
            this.RevisionTextBox.Size = new Size(120, 23);
            this.RevisionTextBox.TabIndex = 16;
            this.RevisionTextBox.TextAlign = HorizontalAlignment.Right;
            // 
            // ExtraCheckBox7
            // 
            this.ExtraCheckBox7.Location = new Point(19, 318);
            this.ExtraCheckBox7.Margin = new Padding(0);
            this.ExtraCheckBox7.Name = "ExtraCheckBox7";
            this.ExtraCheckBox7.Size = new Size(68, 19);
            this.ExtraCheckBox7.TabIndex = 15;
            this.ExtraCheckBox7.TabStop = false;
            this.ExtraCheckBox7.Text = "予備";
            this.ExtraCheckBox7.UseVisualStyleBackColor = true;
            this.ExtraCheckBox7.CheckedChanged += this.CheckBoxChecked;
            // 
            // ExtraTextBox1
            // 
            this.ExtraTextBox1.Enabled = false;
            this.ExtraTextBox1.Location = new Point(166, 141);
            this.ExtraTextBox1.Margin = new Padding(0);
            this.ExtraTextBox1.MaxLength = 20;
            this.ExtraTextBox1.Name = "ExtraTextBox1";
            this.ExtraTextBox1.Size = new Size(120, 23);
            this.ExtraTextBox1.TabIndex = 18;
            this.ExtraTextBox1.TextAlign = HorizontalAlignment.Right;
            // 
            // ExtraCheckBox1
            // 
            this.ExtraCheckBox1.Location = new Point(166, 118);
            this.ExtraCheckBox1.Margin = new Padding(0);
            this.ExtraCheckBox1.Name = "ExtraCheckBox1";
            this.ExtraCheckBox1.Size = new Size(50, 19);
            this.ExtraCheckBox1.TabIndex = 17;
            this.ExtraCheckBox1.TabStop = false;
            this.ExtraCheckBox1.Text = "予備";
            this.ExtraCheckBox1.UseVisualStyleBackColor = true;
            this.ExtraCheckBox1.Visible = false;
            this.ExtraCheckBox1.CheckedChanged += this.CheckBoxChecked;
            // 
            // ExtraTextBox2
            // 
            this.ExtraTextBox2.Enabled = false;
            this.ExtraTextBox2.Location = new Point(166, 191);
            this.ExtraTextBox2.Margin = new Padding(0);
            this.ExtraTextBox2.MaxLength = 20;
            this.ExtraTextBox2.Name = "ExtraTextBox2";
            this.ExtraTextBox2.Size = new Size(120, 23);
            this.ExtraTextBox2.TabIndex = 20;
            this.ExtraTextBox2.TextAlign = HorizontalAlignment.Right;
            // 
            // ExtraCheckBox2
            // 
            this.ExtraCheckBox2.Location = new Point(166, 168);
            this.ExtraCheckBox2.Margin = new Padding(0);
            this.ExtraCheckBox2.Name = "ExtraCheckBox2";
            this.ExtraCheckBox2.Size = new Size(50, 19);
            this.ExtraCheckBox2.TabIndex = 19;
            this.ExtraCheckBox2.TabStop = false;
            this.ExtraCheckBox2.Text = "予備";
            this.ExtraCheckBox2.UseVisualStyleBackColor = true;
            this.ExtraCheckBox2.Visible = false;
            this.ExtraCheckBox2.CheckedChanged += this.CheckBoxChecked;
            // 
            // ExtraTextBox3
            // 
            this.ExtraTextBox3.Enabled = false;
            this.ExtraTextBox3.Location = new Point(166, 241);
            this.ExtraTextBox3.Margin = new Padding(0);
            this.ExtraTextBox3.MaxLength = 20;
            this.ExtraTextBox3.Name = "ExtraTextBox3";
            this.ExtraTextBox3.Size = new Size(120, 23);
            this.ExtraTextBox3.TabIndex = 22;
            this.ExtraTextBox3.TextAlign = HorizontalAlignment.Right;
            // 
            // ExtraCheckBox3
            // 
            this.ExtraCheckBox3.Location = new Point(166, 218);
            this.ExtraCheckBox3.Margin = new Padding(0);
            this.ExtraCheckBox3.Name = "ExtraCheckBox3";
            this.ExtraCheckBox3.Size = new Size(50, 19);
            this.ExtraCheckBox3.TabIndex = 21;
            this.ExtraCheckBox3.TabStop = false;
            this.ExtraCheckBox3.Text = "予備";
            this.ExtraCheckBox3.UseVisualStyleBackColor = true;
            this.ExtraCheckBox3.Visible = false;
            this.ExtraCheckBox3.CheckedChanged += this.CheckBoxChecked;
            // 
            // RegistrationDateCheckBox
            // 
            this.RegistrationDateCheckBox.Location = new Point(166, 268);
            this.RegistrationDateCheckBox.Margin = new Padding(0);
            this.RegistrationDateCheckBox.Name = "RegistrationDateCheckBox";
            this.RegistrationDateCheckBox.Size = new Size(62, 19);
            this.RegistrationDateCheckBox.TabIndex = 23;
            this.RegistrationDateCheckBox.TabStop = false;
            this.RegistrationDateCheckBox.Text = "登録日";
            this.RegistrationDateCheckBox.UseVisualStyleBackColor = true;
            this.RegistrationDateCheckBox.CheckedChanged += this.CheckBoxChecked;
            // 
            // RegistrationDateMaskedTextBox
            // 
            this.RegistrationDateMaskedTextBox.Enabled = false;
            this.RegistrationDateMaskedTextBox.Location = new Point(166, 291);
            this.RegistrationDateMaskedTextBox.Mask = "0000/00/00";
            this.RegistrationDateMaskedTextBox.Name = "RegistrationDateMaskedTextBox";
            this.RegistrationDateMaskedTextBox.PromptChar = '*';
            this.RegistrationDateMaskedTextBox.Size = new Size(120, 23);
            this.RegistrationDateMaskedTextBox.TabIndex = 24;
            this.RegistrationDateMaskedTextBox.TextAlign = HorizontalAlignment.Right;
            this.RegistrationDateMaskedTextBox.TypeValidationCompleted += this.RegistrationDateMaskedTextBox_TypeValidationCompleted;
            // 
            // PersonCheckBox
            // 
            this.PersonCheckBox.Location = new Point(166, 318);
            this.PersonCheckBox.Margin = new Padding(0);
            this.PersonCheckBox.Name = "PersonCheckBox";
            this.PersonCheckBox.Size = new Size(62, 19);
            this.PersonCheckBox.TabIndex = 25;
            this.PersonCheckBox.TabStop = false;
            this.PersonCheckBox.Text = "担当者";
            this.PersonCheckBox.UseVisualStyleBackColor = true;
            this.PersonCheckBox.CheckedChanged += this.CheckBoxChecked;
            // 
            // PersonComboBox
            // 
            this.PersonComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            this.PersonComboBox.Enabled = false;
            this.PersonComboBox.FormattingEnabled = true;
            this.PersonComboBox.Location = new Point(166, 341);
            this.PersonComboBox.Name = "PersonComboBox";
            this.PersonComboBox.Size = new Size(121, 23);
            this.PersonComboBox.TabIndex = 26;
            // 
            // ExtraTextBox4
            // 
            this.ExtraTextBox4.Enabled = false;
            this.ExtraTextBox4.Location = new Point(313, 141);
            this.ExtraTextBox4.Margin = new Padding(0);
            this.ExtraTextBox4.MaxLength = 20;
            this.ExtraTextBox4.Name = "ExtraTextBox4";
            this.ExtraTextBox4.Size = new Size(120, 23);
            this.ExtraTextBox4.TabIndex = 28;
            this.ExtraTextBox4.TextAlign = HorizontalAlignment.Right;
            // 
            // ExtraCheckBox4
            // 
            this.ExtraCheckBox4.Location = new Point(313, 118);
            this.ExtraCheckBox4.Margin = new Padding(0);
            this.ExtraCheckBox4.Name = "ExtraCheckBox4";
            this.ExtraCheckBox4.Size = new Size(50, 19);
            this.ExtraCheckBox4.TabIndex = 27;
            this.ExtraCheckBox4.TabStop = false;
            this.ExtraCheckBox4.Text = "予備";
            this.ExtraCheckBox4.UseVisualStyleBackColor = true;
            this.ExtraCheckBox4.Visible = false;
            this.ExtraCheckBox4.CheckedChanged += this.CheckBoxChecked;
            // 
            // ExtraTextBox5
            // 
            this.ExtraTextBox5.Enabled = false;
            this.ExtraTextBox5.Location = new Point(313, 191);
            this.ExtraTextBox5.Margin = new Padding(0);
            this.ExtraTextBox5.MaxLength = 20;
            this.ExtraTextBox5.Name = "ExtraTextBox5";
            this.ExtraTextBox5.Size = new Size(120, 23);
            this.ExtraTextBox5.TabIndex = 30;
            this.ExtraTextBox5.TextAlign = HorizontalAlignment.Right;
            // 
            // ExtraCheckBox5
            // 
            this.ExtraCheckBox5.Location = new Point(313, 168);
            this.ExtraCheckBox5.Margin = new Padding(0);
            this.ExtraCheckBox5.Name = "ExtraCheckBox5";
            this.ExtraCheckBox5.Size = new Size(50, 19);
            this.ExtraCheckBox5.TabIndex = 29;
            this.ExtraCheckBox5.TabStop = false;
            this.ExtraCheckBox5.Text = "予備";
            this.ExtraCheckBox5.UseVisualStyleBackColor = true;
            this.ExtraCheckBox5.Visible = false;
            this.ExtraCheckBox5.CheckedChanged += this.CheckBoxChecked;
            // 
            // ExtraTextBox6
            // 
            this.ExtraTextBox6.Enabled = false;
            this.ExtraTextBox6.Location = new Point(313, 241);
            this.ExtraTextBox6.Margin = new Padding(0);
            this.ExtraTextBox6.MaxLength = 20;
            this.ExtraTextBox6.Name = "ExtraTextBox6";
            this.ExtraTextBox6.Size = new Size(120, 23);
            this.ExtraTextBox6.TabIndex = 32;
            this.ExtraTextBox6.TextAlign = HorizontalAlignment.Right;
            // 
            // ExtraCheckBox6
            // 
            this.ExtraCheckBox6.Location = new Point(313, 218);
            this.ExtraCheckBox6.Margin = new Padding(0);
            this.ExtraCheckBox6.Name = "ExtraCheckBox6";
            this.ExtraCheckBox6.Size = new Size(50, 19);
            this.ExtraCheckBox6.TabIndex = 31;
            this.ExtraCheckBox6.TabStop = false;
            this.ExtraCheckBox6.Text = "予備";
            this.ExtraCheckBox6.UseVisualStyleBackColor = true;
            this.ExtraCheckBox6.Visible = false;
            this.ExtraCheckBox6.CheckedChanged += this.CheckBoxChecked;
            // 
            // CommentCheckBox
            // 
            this.CommentCheckBox.Location = new Point(464, 118);
            this.CommentCheckBox.Margin = new Padding(0);
            this.CommentCheckBox.Name = "CommentCheckBox";
            this.CommentCheckBox.Size = new Size(59, 19);
            this.CommentCheckBox.TabIndex = 33;
            this.CommentCheckBox.TabStop = false;
            this.CommentCheckBox.Text = "コメント";
            this.CommentCheckBox.UseVisualStyleBackColor = true;
            this.CommentCheckBox.CheckedChanged += this.CheckBoxChecked;
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
            this.CommentTextBox.TabIndex = 34;
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
            this.CommentComboBox.TabIndex = 35;
            // 
            // TemplateButton
            // 
            this.TemplateButton.Enabled = false;
            this.TemplateButton.Location = new Point(539, 347);
            this.TemplateButton.Name = "TemplateButton";
            this.TemplateButton.Size = new Size(75, 25);
            this.TemplateButton.TabIndex = 36;
            this.TemplateButton.Text = "定型文";
            this.TemplateButton.UseVisualStyleBackColor = true;
            this.TemplateButton.Click += this.TemplateButton_Click;
            // 
            // PrintRowLabel
            // 
            this.PrintRowLabel.Location = new Point(177, 392);
            this.PrintRowLabel.Margin = new Padding(0);
            this.PrintRowLabel.Name = "PrintRowLabel";
            this.PrintRowLabel.Size = new Size(79, 15);
            this.PrintRowLabel.TabIndex = 37;
            this.PrintRowLabel.Text = "印刷開始位置";
            // 
            // PrintPostionNumericUpDown
            // 
            this.PrintPostionNumericUpDown.Location = new Point(166, 410);
            this.PrintPostionNumericUpDown.Margin = new Padding(0);
            this.PrintPostionNumericUpDown.Maximum = new decimal(new int[] { 50, 0, 0, 0 });
            this.PrintPostionNumericUpDown.Minimum = new decimal(new int[] { 1, 0, 0, 0 });
            this.PrintPostionNumericUpDown.Name = "PrintPostionNumericUpDown";
            this.PrintPostionNumericUpDown.Size = new Size(110, 23);
            this.PrintPostionNumericUpDown.TabIndex = 38;
            this.PrintPostionNumericUpDown.TextAlign = HorizontalAlignment.Right;
            this.PrintPostionNumericUpDown.Value = new decimal(new int[] { 1, 0, 0, 0 });
            // 
            // RegisterButton
            // 
            this.RegisterButton.Location = new Point(313, 407);
            this.RegisterButton.Margin = new Padding(0);
            this.RegisterButton.Name = "RegisterButton";
            this.RegisterButton.Size = new Size(75, 25);
            this.RegisterButton.TabIndex = 39;
            this.RegisterButton.Text = "登録";
            this.RegisterButton.UseVisualStyleBackColor = true;
            this.RegisterButton.Click += this.RegisterButton_Click;
            // 
            // PrintOnlyCheckBox
            // 
            this.PrintOnlyCheckBox.Location = new Point(426, 382);
            this.PrintOnlyCheckBox.Name = "PrintOnlyCheckBox";
            this.PrintOnlyCheckBox.Size = new Size(71, 19);
            this.PrintOnlyCheckBox.TabIndex = 40;
            this.PrintOnlyCheckBox.TabStop = false;
            this.PrintOnlyCheckBox.Text = "印刷のみ";
            this.PrintOnlyCheckBox.UseVisualStyleBackColor = true;
            this.PrintOnlyCheckBox.CheckedChanged += this.CheckBoxChecked;
            // 
            // PrintButton
            // 
            this.PrintButton.Enabled = false;
            this.PrintButton.Location = new Point(426, 407);
            this.PrintButton.Margin = new Padding(0);
            this.PrintButton.Name = "PrintButton";
            this.PrintButton.Size = new Size(75, 25);
            this.PrintButton.TabIndex = 41;
            this.PrintButton.Text = "印刷";
            this.PrintButton.UseVisualStyleBackColor = true;
            this.PrintButton.Click += this.PrintButton_Click;
            // 
            // SubstrateRegistrationPrintDialog
            // 
            this.SubstrateRegistrationPrintDialog.UseEXDialog = true;
            // 
            // SubstrateRegistrationPrintPreviewDialog
            // 
            this.SubstrateRegistrationPrintPreviewDialog.AutoScrollMargin = new Size(0, 0);
            this.SubstrateRegistrationPrintPreviewDialog.AutoScrollMinSize = new Size(0, 0);
            this.SubstrateRegistrationPrintPreviewDialog.ClientSize = new Size(400, 300);
            this.SubstrateRegistrationPrintPreviewDialog.Enabled = true;
            this.SubstrateRegistrationPrintPreviewDialog.Icon = (Icon)resources.GetObject("SubstrateRegistrationPrintPreviewDialog.Icon");
            this.SubstrateRegistrationPrintPreviewDialog.Name = "SubstrateRegistrationPrintPreviewDialog";
            this.SubstrateRegistrationPrintPreviewDialog.Visible = false;
            this.SubstrateRegistrationPrintPreviewDialog.Load += this.SubstrateRegistrationPrintPreviewDialog_Load;
            // 
            // panelCommentTexrBox
            // 
            this.panelCommentTexrBox.Controls.Add(this.CommentTextBox);
            this.panelCommentTexrBox.Location = new Point(464, 141);
            this.panelCommentTexrBox.Name = "panelCommentTexrBox";
            this.panelCommentTexrBox.Size = new Size(150, 175);
            this.panelCommentTexrBox.TabIndex = 42;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new Point(298, 24);
            this.label1.Name = "label1";
            this.label1.Size = new Size(86, 15);
            this.label1.TabIndex = 90;
            this.label1.Text = "QRコード入力用";
            // 
            // QrCodeButton
            // 
            this.QrCodeButton.Location = new Point(543, 68);
            this.QrCodeButton.Name = "QrCodeButton";
            this.QrCodeButton.Size = new Size(75, 23);
            this.QrCodeButton.TabIndex = 89;
            this.QrCodeButton.Text = "入力";
            this.QrCodeButton.UseVisualStyleBackColor = true;
            this.QrCodeButton.Click += this.QrCodeButton_Click;
            // 
            // QrCodeTextBox
            // 
            this.QrCodeTextBox.Location = new Point(298, 39);
            this.QrCodeTextBox.MaxLength = 100;
            this.QrCodeTextBox.Name = "QrCodeTextBox";
            this.QrCodeTextBox.Size = new Size(320, 23);
            this.QrCodeTextBox.TabIndex = 88;
            this.QrCodeTextBox.Enter += this.QrCodeTextBox_Enter;
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
            // SubstrateRegistrationWindow
            // 
            this.AutoScaleDimensions = new SizeF(7F, 15F);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.ClientSize = new Size(630, 441);
            this.Controls.Add(this.textToUpperCheckBox);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.QrCodeButton);
            this.Controls.Add(this.QrCodeTextBox);
            this.Controls.Add(this.PrintButton);
            this.Controls.Add(this.PrintOnlyCheckBox);
            this.Controls.Add(this.RegisterButton);
            this.Controls.Add(this.PrintPostionNumericUpDown);
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
            this.Controls.Add(this.RegistrationDateMaskedTextBox);
            this.Controls.Add(this.RegistrationDateCheckBox);
            this.Controls.Add(this.ExtraTextBox3);
            this.Controls.Add(this.ExtraCheckBox3);
            this.Controls.Add(this.ExtraTextBox2);
            this.Controls.Add(this.ExtraCheckBox2);
            this.Controls.Add(this.ExtraTextBox1);
            this.Controls.Add(this.ExtraCheckBox1);
            this.Controls.Add(this.RevisionTextBox);
            this.Controls.Add(this.ExtraCheckBox7);
            this.Controls.Add(this.DefectNumberTextBox);
            this.Controls.Add(this.DefectNumberCheckBox);
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
            this.Controls.Add(this.SubstrateRegistrationMenuStrip);
            this.Controls.Add(this.panelCommentTexrBox);
            this.Font = new Font("Meiryo UI", 9F);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MainMenuStrip = this.SubstrateRegistrationMenuStrip;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "SubstrateRegistrationWindow";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = FormStartPosition.CenterParent;
            this.Text = "基板登録";
            this.Load += this.SubstrateRegistrationWindow_Load;
            this.SubstrateRegistrationMenuStrip.ResumeLayout(false);
            this.SubstrateRegistrationMenuStrip.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)this.PrintPostionNumericUpDown).EndInit();
            this.panelCommentTexrBox.ResumeLayout(false);
            this.panelCommentTexrBox.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();
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
        private CheckBox OrderNumberCheckBox;
        private TextBox OrderNumberTextBox;
        private CheckBox ManufacturingNumberCheckBox;
        private MaskedTextBox ManufacturingNumberMaskedTextBox;
        private TextBox QuantityTextBox;
        private CheckBox QuantityCheckBox;
        private TextBox DefectNumberTextBox;
        private CheckBox DefectNumberCheckBox;
        private TextBox RevisionTextBox;
        private CheckBox ExtraCheckBox7;
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
        private Panel panelCommentTexrBox;
        private Label label1;
        private Button QrCodeButton;
        private TextBox QrCodeTextBox;
        private CheckBox textToUpperCheckBox;
    }
}