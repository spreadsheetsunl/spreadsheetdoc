namespace FirstExcelAddin
{
    partial class GeneralForm
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
            this.General_DescriptionGF = new System.Windows.Forms.Label();
            this.General_DescriptionGFBox = new System.Windows.Forms.TextBox();
            this.okButtonGF = new System.Windows.Forms.Button();
            this.cancelButtonGF = new System.Windows.Forms.Button();
            this.clear_buttonGF = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.remove_button_GF = new System.Windows.Forms.Button();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // General_DescriptionGF
            // 
            this.General_DescriptionGF.Dock = System.Windows.Forms.DockStyle.Fill;
            this.General_DescriptionGF.Location = new System.Drawing.Point(0, 0);
            this.General_DescriptionGF.Name = "General_DescriptionGF";
            this.General_DescriptionGF.Padding = new System.Windows.Forms.Padding(0, 15, 0, 0);
            this.General_DescriptionGF.Size = new System.Drawing.Size(356, 290);
            this.General_DescriptionGF.TabIndex = 0;
            this.General_DescriptionGF.Text = "General Description";
            this.General_DescriptionGF.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // General_DescriptionGFBox
            // 
            this.General_DescriptionGFBox.Location = new System.Drawing.Point(59, 48);
            this.General_DescriptionGFBox.Multiline = true;
            this.General_DescriptionGFBox.Name = "General_DescriptionGFBox";
            this.General_DescriptionGFBox.Size = new System.Drawing.Size(236, 152);
            this.General_DescriptionGFBox.TabIndex = 1;
            // 
            // okButtonGF
            // 
            this.okButtonGF.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.okButtonGF.Location = new System.Drawing.Point(188, 28);
            this.okButtonGF.Name = "okButtonGF";
            this.okButtonGF.Size = new System.Drawing.Size(75, 23);
            this.okButtonGF.TabIndex = 2;
            this.okButtonGF.Text = "OK";
            this.okButtonGF.UseVisualStyleBackColor = true;
            this.okButtonGF.Click += new System.EventHandler(this.okButton_Click);
            // 
            // cancelButtonGF
            // 
            this.cancelButtonGF.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.cancelButtonGF.Location = new System.Drawing.Point(269, 28);
            this.cancelButtonGF.Name = "cancelButtonGF";
            this.cancelButtonGF.Size = new System.Drawing.Size(75, 23);
            this.cancelButtonGF.TabIndex = 3;
            this.cancelButtonGF.Text = "CANCEL";
            this.cancelButtonGF.UseVisualStyleBackColor = true;
            this.cancelButtonGF.Click += new System.EventHandler(this.cancelButton_Click);
            // 
            // clear_buttonGF
            // 
            this.clear_buttonGF.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.clear_buttonGF.Location = new System.Drawing.Point(12, 28);
            this.clear_buttonGF.Name = "clear_buttonGF";
            this.clear_buttonGF.Size = new System.Drawing.Size(75, 23);
            this.clear_buttonGF.TabIndex = 28;
            this.clear_buttonGF.Text = "CLEAR";
            this.clear_buttonGF.Click += new System.EventHandler(this.clear_button_Click);
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.remove_button_GF);
            this.panel1.Controls.Add(this.clear_buttonGF);
            this.panel1.Controls.Add(this.cancelButtonGF);
            this.panel1.Controls.Add(this.okButtonGF);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel1.Location = new System.Drawing.Point(0, 236);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(356, 54);
            this.panel1.TabIndex = 29;
            // 
            // remove_button_GF
            // 
            this.remove_button_GF.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.remove_button_GF.Enabled = false;
            this.remove_button_GF.Location = new System.Drawing.Point(12, 3);
            this.remove_button_GF.Name = "remove_button_GF";
            this.remove_button_GF.Size = new System.Drawing.Size(75, 23);
            this.remove_button_GF.TabIndex = 29;
            this.remove_button_GF.Text = "REMOVE";
            this.remove_button_GF.UseVisualStyleBackColor = true;
            this.remove_button_GF.Visible = false;
            this.remove_button_GF.Click += new System.EventHandler(this.remove_button_Click);
            // 
            // GeneralForm
            // 
            this.AcceptButton = this.okButtonGF;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.cancelButtonGF;
            this.ClientSize = new System.Drawing.Size(356, 290);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.General_DescriptionGFBox);
            this.Controls.Add(this.General_DescriptionGF);
            this.MaximizeBox = false;
            this.Name = "GeneralForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "GeneralForm";
            this.Load += new System.EventHandler(this.GeneralForm_Load);
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label General_DescriptionGF;
        public System.Windows.Forms.TextBox General_DescriptionGFBox;
        private System.Windows.Forms.Panel panel1;
        public System.Windows.Forms.Button remove_button_GF;
        public System.Windows.Forms.Button clear_buttonGF;
        public System.Windows.Forms.Button cancelButtonGF;
        public System.Windows.Forms.Button okButtonGF;

    }
}