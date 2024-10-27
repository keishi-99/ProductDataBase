namespace ProductDatabase {
    partial class SeveralDialogWindow {
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
            this.SeveralListBox = new ListBox();
            this.OKButton = new Button();
            this.panelSeveralList = new Panel();
            this.panelSeveralList.SuspendLayout();
            this.SuspendLayout();
            // 
            // SeveralListBox
            // 
            this.SeveralListBox.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            this.SeveralListBox.FormattingEnabled = true;
            this.SeveralListBox.ItemHeight = 15;
            this.SeveralListBox.Location = new Point(0, 0);
            this.SeveralListBox.Name = "SeveralListBox";
            this.SeveralListBox.Size = new Size(482, 214);
            this.SeveralListBox.TabIndex = 0;
            this.SeveralListBox.KeyPress += this.SeveralListBox_KeyPress;
            // 
            // OKButton
            // 
            this.OKButton.Location = new Point(216, 247);
            this.OKButton.Name = "OKButton";
            this.OKButton.Size = new Size(75, 23);
            this.OKButton.TabIndex = 1;
            this.OKButton.Text = "OK";
            this.OKButton.UseVisualStyleBackColor = true;
            this.OKButton.Click += this.OkButton_Click;
            // 
            // panelSeveralList
            // 
            this.panelSeveralList.Controls.Add(this.SeveralListBox);
            this.panelSeveralList.Location = new Point(12, 12);
            this.panelSeveralList.Name = "panelSeveralList";
            this.panelSeveralList.Size = new Size(482, 214);
            this.panelSeveralList.TabIndex = 2;
            // 
            // SeveralDialogWindow
            // 
            this.AutoScaleDimensions = new SizeF(7F, 15F);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.ClientSize = new Size(506, 282);
            this.ControlBox = false;
            this.Controls.Add(this.OKButton);
            this.Controls.Add(this.panelSeveralList);
            this.Font = new Font("Meiryo UI", 9F);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "SeveralDialogWindow";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = FormStartPosition.CenterParent;
            this.TopMost = true;
            this.Load += this.SeveralDialogWindow_Load;
            this.panelSeveralList.ResumeLayout(false);
            this.ResumeLayout(false);
        }

        #endregion

        private ListBox SeveralListBox;
        private Button OKButton;
        private Panel panelSeveralList;
    }
}