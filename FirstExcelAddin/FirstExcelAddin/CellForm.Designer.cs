namespace FirstExcelAddin
{
    partial class CellForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.General_DescriptionCF = new System.Windows.Forms.Label();
            this.OutputBox = new System.Windows.Forms.TextBox();
            this.Output = new System.Windows.Forms.Label();
            this.okButtonCell = new System.Windows.Forms.Button();
            this.cancelButtonCell = new System.Windows.Forms.Button();
            this.clear_buttonCell = new System.Windows.Forms.Button();
            this.General_DescriptionBox = new System.Windows.Forms.TextBox();
            this.inputBox = new System.Windows.Forms.TextBox();
            this.inputCell = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.panel2 = new System.Windows.Forms.Panel();
            this.remove_buttonCell = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // General_DescriptionCF
            // 
            this.General_DescriptionCF.AutoSize = true;
            this.General_DescriptionCF.Location = new System.Drawing.Point(12, 26);
            this.General_DescriptionCF.Name = "General_DescriptionCF";
            this.General_DescriptionCF.Size = new System.Drawing.Size(100, 13);
            this.General_DescriptionCF.TabIndex = 1;
            this.General_DescriptionCF.Text = "General Description";
            // 
            // OutputBox
            // 
            this.OutputBox.Location = new System.Drawing.Point(118, 182);
            this.OutputBox.Multiline = true;
            this.OutputBox.Name = "OutputBox";
            this.OutputBox.Size = new System.Drawing.Size(378, 50);
            this.OutputBox.TabIndex = 4;
            // 
            // Output
            // 
            this.Output.AutoSize = true;
            this.Output.Location = new System.Drawing.Point(12, 185);
            this.Output.Name = "Output";
            this.Output.Size = new System.Drawing.Size(39, 13);
            this.Output.TabIndex = 5;
            this.Output.Text = "Output";
            // 
            // okButtonCell
            // 
            this.okButtonCell.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.okButtonCell.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.okButtonCell.Location = new System.Drawing.Point(352, 28);
            this.okButtonCell.Name = "okButtonCell";
            this.okButtonCell.Size = new System.Drawing.Size(75, 23);
            this.okButtonCell.TabIndex = 25;
            this.okButtonCell.Text = "OK";
            this.okButtonCell.Click += new System.EventHandler(this.okButton_Click);
            // 
            // cancelButtonCell
            // 
            this.cancelButtonCell.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.cancelButtonCell.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.cancelButtonCell.Location = new System.Drawing.Point(433, 28);
            this.cancelButtonCell.Name = "cancelButtonCell";
            this.cancelButtonCell.Size = new System.Drawing.Size(75, 23);
            this.cancelButtonCell.TabIndex = 26;
            this.cancelButtonCell.Text = "CANCEL";
            this.cancelButtonCell.Click += new System.EventHandler(this.cancelButton_Click);
            // 
            // clear_buttonCell
            // 
            this.clear_buttonCell.Location = new System.Drawing.Point(3, 28);
            this.clear_buttonCell.Name = "clear_buttonCell";
            this.clear_buttonCell.Size = new System.Drawing.Size(75, 23);
            this.clear_buttonCell.TabIndex = 27;
            this.clear_buttonCell.Text = "CLEAR";
            this.clear_buttonCell.Click += new System.EventHandler(this.clear_button_Click);
            // 
            // General_DescriptionBox
            // 
            this.General_DescriptionBox.Location = new System.Drawing.Point(118, 26);
            this.General_DescriptionBox.Multiline = true;
            this.General_DescriptionBox.Name = "General_DescriptionBox";
            this.General_DescriptionBox.Size = new System.Drawing.Size(378, 71);
            this.General_DescriptionBox.TabIndex = 0;
            // 
            // inputBox
            // 
            this.inputBox.Location = new System.Drawing.Point(118, 115);
            this.inputBox.Multiline = true;
            this.inputBox.Name = "inputBox";
            this.inputBox.Size = new System.Drawing.Size(378, 50);
            this.inputBox.TabIndex = 29;
            // 
            // inputCell
            // 
            this.inputCell.AutoSize = true;
            this.inputCell.Location = new System.Drawing.Point(12, 118);
            this.inputCell.MaximumSize = new System.Drawing.Size(97, 0);
            this.inputCell.Name = "inputCell";
            this.inputCell.Size = new System.Drawing.Size(31, 13);
            this.inputCell.TabIndex = 30;
            this.inputCell.Text = "Input";
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.okButtonCell);
            this.panel1.Controls.Add(this.cancelButtonCell);
            this.panel1.Controls.Add(this.panel2);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel1.Location = new System.Drawing.Point(0, 238);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(520, 54);
            this.panel1.TabIndex = 31;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.clear_buttonCell);
            this.panel2.Controls.Add(this.remove_buttonCell);
            this.panel2.Location = new System.Drawing.Point(30, 0);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(82, 54);
            this.panel2.TabIndex = 32;
            // 
            // remove_buttonCell
            // 
            this.remove_buttonCell.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.remove_buttonCell.Enabled = false;
            this.remove_buttonCell.Location = new System.Drawing.Point(3, 3);
            this.remove_buttonCell.Name = "remove_buttonCell";
            this.remove_buttonCell.Size = new System.Drawing.Size(75, 23);
            this.remove_buttonCell.TabIndex = 28;
            this.remove_buttonCell.Text = "REMOVE";
            this.remove_buttonCell.UseVisualStyleBackColor = true;
            this.remove_buttonCell.Click += new System.EventHandler(this.remove_button_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 94);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(31, 13);
            this.label1.TabIndex = 32;
            this.label1.Text = "Input";
            // 
            // CellForm
            // 
            this.AcceptButton = this.okButtonCell;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.CancelButton = this.cancelButtonCell;
            this.ClientSize = new System.Drawing.Size(520, 292);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.inputCell);
            this.Controls.Add(this.inputBox);
            this.Controls.Add(this.Output);
            this.Controls.Add(this.OutputBox);
            this.Controls.Add(this.General_DescriptionCF);
            this.Controls.Add(this.General_DescriptionBox);
            this.MaximizeBox = false;
            this.Name = "CellForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Cell";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.panel1.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label General_DescriptionCF;
        public System.Windows.Forms.Label Output;
        public System.Windows.Forms.TextBox General_DescriptionBox;
        private System.Windows.Forms.Panel panel1;
        public System.Windows.Forms.Button remove_buttonCell;
        private System.Windows.Forms.Panel panel2;
        public System.Windows.Forms.Button cancelButtonCell;
        public System.Windows.Forms.Button clear_buttonCell;
        public System.Windows.Forms.Button okButtonCell;
        public System.Windows.Forms.TextBox OutputBox;
        public System.Windows.Forms.TextBox inputBox;
        public System.Windows.Forms.Label inputCell;
        private System.Windows.Forms.Label label1;
    }
}