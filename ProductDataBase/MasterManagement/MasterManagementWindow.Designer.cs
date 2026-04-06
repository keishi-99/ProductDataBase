namespace ProductDatabase.MasterManagement {
    partial class MasterManagementWindow {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing) {
            if (disposing && (components != null)) {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        private void InitializeComponent() {
            this.MasterTabControl = new TabControl();
            this.ProductTabPage = new TabPage();
            this.ProductDataGridView = new DataGridView();
            this.ProductButtonPanel = new Panel();
            this.ProductAddButton = new Button();
            this.ProductEditButton = new Button();
            this.ProductDeleteButton = new Button();
            this.ProductPrintDetailContextMenuStrip = new ContextMenuStrip();
            this.ProductLabelSettingsMenuItem = new ToolStripMenuItem();
            this.ProductBarcodeSettingsMenuItem = new ToolStripMenuItem();
            this.ProductNameplateSettingsMenuItem = new ToolStripMenuItem();
            this.ProductPrintDetailSettingsButton = new Button();
            this.ProductSearchPanel = new Panel();
            this.ProductSearchLabel = new Label();
            this.ProductSearchBox = new TextBox();
            this.SubstrateTabPage = new TabPage();
            this.SubstrateDataGridView = new DataGridView();
            this.SubstrateButtonPanel = new Panel();
            this.SubstrateAddButton = new Button();
            this.SubstrateEditButton = new Button();
            this.SubstrateDeleteButton = new Button();
            this.SubstrateSearchPanel = new Panel();
            this.SubstrateSearchLabel = new Label();
            this.SubstrateSearchBox = new TextBox();
            this.MasterTabControl.SuspendLayout();
            this.ProductTabPage.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)this.ProductDataGridView).BeginInit();
            this.ProductButtonPanel.SuspendLayout();
            this.ProductSearchPanel.SuspendLayout();
            this.SubstrateTabPage.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)this.SubstrateDataGridView).BeginInit();
            this.SubstrateButtonPanel.SuspendLayout();
            this.SubstrateSearchPanel.SuspendLayout();
            this.SuspendLayout();
            // 
            // MasterTabControl
            // 
            this.MasterTabControl.Controls.Add(this.SubstrateTabPage);
            this.MasterTabControl.Controls.Add(this.ProductTabPage);
            this.MasterTabControl.Dock = DockStyle.Fill;
            this.MasterTabControl.Location = new Point(0, 0);
            this.MasterTabControl.Name = "MasterTabControl";
            this.MasterTabControl.SelectedIndex = 0;
            this.MasterTabControl.Size = new Size(960, 640);
            this.MasterTabControl.TabIndex = 0;
            // 
            // ProductTabPage
            // 
            this.ProductTabPage.Controls.Add(this.ProductDataGridView);
            this.ProductTabPage.Controls.Add(this.ProductButtonPanel);
            this.ProductTabPage.Controls.Add(this.ProductSearchPanel);
            this.ProductTabPage.Location = new Point(4, 24);
            this.ProductTabPage.Name = "ProductTabPage";
            this.ProductTabPage.Padding = new Padding(3);
            this.ProductTabPage.Size = new Size(952, 612);
            this.ProductTabPage.TabIndex = 1;
            this.ProductTabPage.Text = "製品マスター";
            this.ProductTabPage.UseVisualStyleBackColor = true;
            // 
            // ProductDataGridView
            // 
            this.ProductDataGridView.AllowUserToAddRows = false;
            this.ProductDataGridView.AllowUserToDeleteRows = false;
            this.ProductDataGridView.AllowUserToResizeRows = false;
            this.ProductDataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            this.ProductDataGridView.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.ProductDataGridView.Dock = DockStyle.Fill;
            this.ProductDataGridView.Location = new Point(3, 39);
            this.ProductDataGridView.MultiSelect = false;
            this.ProductDataGridView.Name = "ProductDataGridView";
            this.ProductDataGridView.ReadOnly = true;
            this.ProductDataGridView.RowHeadersVisible = false;
            this.ProductDataGridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            this.ProductDataGridView.Size = new Size(946, 536);
            this.ProductDataGridView.TabIndex = 1;
            this.ProductDataGridView.ColumnHeaderMouseClick += this.ProductDataGridView_ColumnHeaderMouseClick;
            // 
            // ProductButtonPanel
            // 
            this.ProductButtonPanel.Controls.Add(this.ProductAddButton);
            this.ProductButtonPanel.Controls.Add(this.ProductEditButton);
            this.ProductButtonPanel.Controls.Add(this.ProductDeleteButton);
            this.ProductButtonPanel.Controls.Add(this.ProductPrintDetailSettingsButton);
            this.ProductButtonPanel.Dock = DockStyle.Bottom;
            this.ProductButtonPanel.Location = new Point(3, 575);
            this.ProductButtonPanel.Name = "ProductButtonPanel";
            this.ProductButtonPanel.Size = new Size(946, 34);
            this.ProductButtonPanel.TabIndex = 2;
            // 
            // ProductAddButton
            // 
            this.ProductAddButton.Location = new Point(3, 3);
            this.ProductAddButton.Name = "ProductAddButton";
            this.ProductAddButton.Size = new Size(80, 28);
            this.ProductAddButton.TabIndex = 0;
            this.ProductAddButton.Text = "追加";
            this.ProductAddButton.UseVisualStyleBackColor = true;
            this.ProductAddButton.Click += this.ProductAddButton_Click;
            // 
            // ProductEditButton
            // 
            this.ProductEditButton.Location = new Point(90, 3);
            this.ProductEditButton.Name = "ProductEditButton";
            this.ProductEditButton.Size = new Size(80, 28);
            this.ProductEditButton.TabIndex = 1;
            this.ProductEditButton.Text = "編集";
            this.ProductEditButton.UseVisualStyleBackColor = true;
            this.ProductEditButton.Click += this.ProductEditButton_Click;
            // 
            // ProductDeleteButton
            // 
            this.ProductDeleteButton.Location = new Point(177, 3);
            this.ProductDeleteButton.Name = "ProductDeleteButton";
            this.ProductDeleteButton.Size = new Size(80, 28);
            this.ProductDeleteButton.TabIndex = 2;
            this.ProductDeleteButton.Text = "削除";
            this.ProductDeleteButton.UseVisualStyleBackColor = true;
            this.ProductDeleteButton.Click += this.ProductDeleteButton_Click;
            //
            // ProductPrintDetailContextMenuStrip
            //
            this.ProductPrintDetailContextMenuStrip.Items.AddRange(new ToolStripItem[] {
                this.ProductLabelSettingsMenuItem,
                this.ProductBarcodeSettingsMenuItem,
                this.ProductNameplateSettingsMenuItem
            });
            this.ProductPrintDetailContextMenuStrip.Name = "ProductPrintDetailContextMenuStrip";
            //
            // ProductLabelSettingsMenuItem
            //
            this.ProductLabelSettingsMenuItem.Name = "ProductLabelSettingsMenuItem";
            this.ProductLabelSettingsMenuItem.Text = "ラベル設定";
            this.ProductLabelSettingsMenuItem.Click += this.ProductLabelSettingsMenuItem_Click;
            //
            // ProductBarcodeSettingsMenuItem
            //
            this.ProductBarcodeSettingsMenuItem.Name = "ProductBarcodeSettingsMenuItem";
            this.ProductBarcodeSettingsMenuItem.Text = "バーコード設定";
            this.ProductBarcodeSettingsMenuItem.Click += this.ProductBarcodeSettingsMenuItem_Click;
            //
            // ProductNameplateSettingsMenuItem
            //
            this.ProductNameplateSettingsMenuItem.Name = "ProductNameplateSettingsMenuItem";
            this.ProductNameplateSettingsMenuItem.Text = "銘版設定";
            this.ProductNameplateSettingsMenuItem.Click += this.ProductNameplateSettingsMenuItem_Click;
            //
            // ProductPrintDetailSettingsButton
            //
            this.ProductPrintDetailSettingsButton.Location = new Point(264, 3);
            this.ProductPrintDetailSettingsButton.Name = "ProductPrintDetailSettingsButton";
            this.ProductPrintDetailSettingsButton.Size = new Size(120, 28);
            this.ProductPrintDetailSettingsButton.TabIndex = 3;
            this.ProductPrintDetailSettingsButton.Text = "印刷詳細設定 ▼";
            this.ProductPrintDetailSettingsButton.UseVisualStyleBackColor = true;
            this.ProductPrintDetailSettingsButton.Click += this.ProductPrintDetailSettingsButton_Click;
            //
            // ProductSearchPanel
            // 
            this.ProductSearchPanel.Controls.Add(this.ProductSearchLabel);
            this.ProductSearchPanel.Controls.Add(this.ProductSearchBox);
            this.ProductSearchPanel.Dock = DockStyle.Top;
            this.ProductSearchPanel.Location = new Point(3, 3);
            this.ProductSearchPanel.Name = "ProductSearchPanel";
            this.ProductSearchPanel.Size = new Size(946, 36);
            this.ProductSearchPanel.TabIndex = 0;
            // 
            // ProductSearchLabel
            // 
            this.ProductSearchLabel.AutoSize = true;
            this.ProductSearchLabel.Location = new Point(3, 9);
            this.ProductSearchLabel.Name = "ProductSearchLabel";
            this.ProductSearchLabel.Size = new Size(126, 15);
            this.ProductSearchLabel.TabIndex = 0;
            this.ProductSearchLabel.Text = "検索（製品名・型式）:";
            // 
            // ProductSearchBox
            // 
            this.ProductSearchBox.Location = new Point(190, 6);
            this.ProductSearchBox.Name = "ProductSearchBox";
            this.ProductSearchBox.Size = new Size(300, 23);
            this.ProductSearchBox.TabIndex = 1;
            this.ProductSearchBox.TextChanged += this.ProductSearchBox_TextChanged;
            // 
            // SubstrateTabPage
            // 
            this.SubstrateTabPage.Controls.Add(this.SubstrateDataGridView);
            this.SubstrateTabPage.Controls.Add(this.SubstrateButtonPanel);
            this.SubstrateTabPage.Controls.Add(this.SubstrateSearchPanel);
            this.SubstrateTabPage.Location = new Point(4, 24);
            this.SubstrateTabPage.Name = "SubstrateTabPage";
            this.SubstrateTabPage.Padding = new Padding(3);
            this.SubstrateTabPage.Size = new Size(952, 612);
            this.SubstrateTabPage.TabIndex = 0;
            this.SubstrateTabPage.Text = "基板マスター";
            this.SubstrateTabPage.UseVisualStyleBackColor = true;
            // 
            // SubstrateDataGridView
            // 
            this.SubstrateDataGridView.AllowUserToAddRows = false;
            this.SubstrateDataGridView.AllowUserToDeleteRows = false;
            this.SubstrateDataGridView.AllowUserToResizeRows = false;
            this.SubstrateDataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            this.SubstrateDataGridView.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.SubstrateDataGridView.Dock = DockStyle.Fill;
            this.SubstrateDataGridView.Location = new Point(3, 39);
            this.SubstrateDataGridView.MultiSelect = false;
            this.SubstrateDataGridView.Name = "SubstrateDataGridView";
            this.SubstrateDataGridView.ReadOnly = true;
            this.SubstrateDataGridView.RowHeadersVisible = false;
            this.SubstrateDataGridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            this.SubstrateDataGridView.Size = new Size(946, 536);
            this.SubstrateDataGridView.TabIndex = 1;
            this.SubstrateDataGridView.ColumnHeaderMouseClick += this.SubstrateDataGridView_ColumnHeaderMouseClick;
            // 
            // SubstrateButtonPanel
            // 
            this.SubstrateButtonPanel.Controls.Add(this.SubstrateAddButton);
            this.SubstrateButtonPanel.Controls.Add(this.SubstrateEditButton);
            this.SubstrateButtonPanel.Controls.Add(this.SubstrateDeleteButton);
            this.SubstrateButtonPanel.Dock = DockStyle.Bottom;
            this.SubstrateButtonPanel.Location = new Point(3, 575);
            this.SubstrateButtonPanel.Name = "SubstrateButtonPanel";
            this.SubstrateButtonPanel.Size = new Size(946, 34);
            this.SubstrateButtonPanel.TabIndex = 2;
            // 
            // SubstrateAddButton
            // 
            this.SubstrateAddButton.Location = new Point(3, 3);
            this.SubstrateAddButton.Name = "SubstrateAddButton";
            this.SubstrateAddButton.Size = new Size(80, 28);
            this.SubstrateAddButton.TabIndex = 0;
            this.SubstrateAddButton.Text = "追加";
            this.SubstrateAddButton.UseVisualStyleBackColor = true;
            this.SubstrateAddButton.Click += this.SubstrateAddButton_Click;
            // 
            // SubstrateEditButton
            // 
            this.SubstrateEditButton.Location = new Point(90, 3);
            this.SubstrateEditButton.Name = "SubstrateEditButton";
            this.SubstrateEditButton.Size = new Size(80, 28);
            this.SubstrateEditButton.TabIndex = 1;
            this.SubstrateEditButton.Text = "編集";
            this.SubstrateEditButton.UseVisualStyleBackColor = true;
            this.SubstrateEditButton.Click += this.SubstrateEditButton_Click;
            // 
            // SubstrateDeleteButton
            // 
            this.SubstrateDeleteButton.Location = new Point(177, 3);
            this.SubstrateDeleteButton.Name = "SubstrateDeleteButton";
            this.SubstrateDeleteButton.Size = new Size(80, 28);
            this.SubstrateDeleteButton.TabIndex = 2;
            this.SubstrateDeleteButton.Text = "削除";
            this.SubstrateDeleteButton.UseVisualStyleBackColor = true;
            this.SubstrateDeleteButton.Click += this.SubstrateDeleteButton_Click;
            // 
            // SubstrateSearchPanel
            // 
            this.SubstrateSearchPanel.Controls.Add(this.SubstrateSearchLabel);
            this.SubstrateSearchPanel.Controls.Add(this.SubstrateSearchBox);
            this.SubstrateSearchPanel.Dock = DockStyle.Top;
            this.SubstrateSearchPanel.Location = new Point(3, 3);
            this.SubstrateSearchPanel.Name = "SubstrateSearchPanel";
            this.SubstrateSearchPanel.Size = new Size(946, 36);
            this.SubstrateSearchPanel.TabIndex = 0;
            // 
            // SubstrateSearchLabel
            // 
            this.SubstrateSearchLabel.AutoSize = true;
            this.SubstrateSearchLabel.Location = new Point(3, 9);
            this.SubstrateSearchLabel.Name = "SubstrateSearchLabel";
            this.SubstrateSearchLabel.Size = new Size(150, 15);
            this.SubstrateSearchLabel.TabIndex = 0;
            this.SubstrateSearchLabel.Text = "検索（製品名・基板型式）:";
            // 
            // SubstrateSearchBox
            // 
            this.SubstrateSearchBox.Location = new Point(190, 6);
            this.SubstrateSearchBox.Name = "SubstrateSearchBox";
            this.SubstrateSearchBox.Size = new Size(300, 23);
            this.SubstrateSearchBox.TabIndex = 1;
            this.SubstrateSearchBox.TextChanged += this.SubstrateSearchBox_TextChanged;
            // 
            // MasterManagementWindow
            // 
            this.AutoScaleDimensions = new SizeF(7F, 15F);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.ClientSize = new Size(960, 640);
            this.Controls.Add(this.MasterTabControl);
            this.Font = new Font("Meiryo UI", 9F);
            this.MinimumSize = new Size(800, 500);
            this.Name = "MasterManagementWindow";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = FormStartPosition.CenterParent;
            this.Text = "マスター管理";
            this.Load += this.MasterManagementWindow_Load;
            this.MasterTabControl.ResumeLayout(false);
            this.ProductTabPage.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)this.ProductDataGridView).EndInit();
            this.ProductButtonPanel.ResumeLayout(false);
            this.ProductSearchPanel.ResumeLayout(false);
            this.ProductSearchPanel.PerformLayout();
            this.SubstrateTabPage.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)this.SubstrateDataGridView).EndInit();
            this.SubstrateButtonPanel.ResumeLayout(false);
            this.SubstrateSearchPanel.ResumeLayout(false);
            this.SubstrateSearchPanel.PerformLayout();
            this.ResumeLayout(false);
        }

        #endregion

        private TabControl MasterTabControl;
        private TabPage ProductTabPage;
        private Panel ProductSearchPanel;
        private Label ProductSearchLabel;
        private TextBox ProductSearchBox;
        private DataGridView ProductDataGridView;
        private Panel ProductButtonPanel;
        private Button ProductAddButton;
        private Button ProductEditButton;
        private Button ProductDeleteButton;
        private ContextMenuStrip ProductPrintDetailContextMenuStrip;
        private ToolStripMenuItem ProductLabelSettingsMenuItem;
        private ToolStripMenuItem ProductBarcodeSettingsMenuItem;
        private ToolStripMenuItem ProductNameplateSettingsMenuItem;
        private Button ProductPrintDetailSettingsButton;
        private TabPage SubstrateTabPage;
        private Panel SubstrateSearchPanel;
        private Label SubstrateSearchLabel;
        private TextBox SubstrateSearchBox;
        private DataGridView SubstrateDataGridView;
        private Panel SubstrateButtonPanel;
        private Button SubstrateAddButton;
        private Button SubstrateEditButton;
        private Button SubstrateDeleteButton;
    }
}
