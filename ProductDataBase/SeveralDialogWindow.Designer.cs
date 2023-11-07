namespace ProductDataBase {
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
            SeveralListBox = new ListBox();
            OKButton = new Button();
            SuspendLayout();
            // 
            // SeveralListBox
            // 
            SeveralListBox.FormattingEnabled = true;
            SeveralListBox.ItemHeight = 15;
            SeveralListBox.Location = new Point(37, 12);
            SeveralListBox.Name = "SeveralListBox";
            SeveralListBox.Size = new Size(320, 94);
            SeveralListBox.TabIndex = 0;
            SeveralListBox.KeyPress += SeveralListBox_KeyPress;
            // 
            // OKButton
            // 
            OKButton.Location = new Point(160, 119);
            OKButton.Name = "OKButton";
            OKButton.Size = new Size(75, 23);
            OKButton.TabIndex = 1;
            OKButton.Text = "OK";
            OKButton.UseVisualStyleBackColor = true;
            OKButton.Click += OkButton_Click;
            // 
            // SeveralDialogWindow
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(394, 154);
            ControlBox = false;
            Controls.Add(OKButton);
            Controls.Add(SeveralListBox);
            Font = new Font("Meiryo UI", 9F, FontStyle.Regular, GraphicsUnit.Point);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "SeveralDialogWindow";
            ShowIcon = false;
            ShowInTaskbar = false;
            TopMost = true;
            Load += SeveralDialogWindow_Load;
            ResumeLayout(false);
        }

        #endregion

        private ListBox SeveralListBox;
        private Button OKButton;
    }
}