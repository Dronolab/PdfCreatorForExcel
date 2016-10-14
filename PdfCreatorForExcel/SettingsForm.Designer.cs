namespace PdfCreatorForExcel
{
    partial class SettingsForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SettingsForm));
            this.BtnSave = new System.Windows.Forms.Button();
            this.BtnChooseTemplatePath = new System.Windows.Forms.Button();
            this.TxbTemplatePath = new System.Windows.Forms.TextBox();
            this.BtnChooseOutputPath = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.TxbOutputPath = new System.Windows.Forms.TextBox();
            this.OFDTemplatePath = new System.Windows.Forms.OpenFileDialog();
            this.FBDOutputPath = new System.Windows.Forms.FolderBrowserDialog();
            this.label2 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // BtnSave
            // 
            this.BtnSave.Location = new System.Drawing.Point(260, 145);
            this.BtnSave.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.BtnSave.Name = "BtnSave";
            this.BtnSave.Size = new System.Drawing.Size(107, 39);
            this.BtnSave.TabIndex = 14;
            this.BtnSave.Text = "Sauvegarder";
            this.BtnSave.UseVisualStyleBackColor = true;
            this.BtnSave.Click += new System.EventHandler(this.SaveSettings);
            // 
            // BtnChooseTemplatePath
            // 
            this.BtnChooseTemplatePath.Location = new System.Drawing.Point(549, 91);
            this.BtnChooseTemplatePath.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.BtnChooseTemplatePath.Name = "BtnChooseTemplatePath";
            this.BtnChooseTemplatePath.Size = new System.Drawing.Size(107, 39);
            this.BtnChooseTemplatePath.TabIndex = 13;
            this.BtnChooseTemplatePath.Text = "Choisir";
            this.BtnChooseTemplatePath.UseVisualStyleBackColor = true;
            this.BtnChooseTemplatePath.Click += new System.EventHandler(this.ChooseTemplatePath);
            // 
            // TxbTemplatePath
            // 
            this.TxbTemplatePath.Location = new System.Drawing.Point(15, 107);
            this.TxbTemplatePath.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.TxbTemplatePath.Name = "TxbTemplatePath";
            this.TxbTemplatePath.Size = new System.Drawing.Size(529, 22);
            this.TxbTemplatePath.TabIndex = 11;
            // 
            // BtnChooseOutputPath
            // 
            this.BtnChooseOutputPath.Location = new System.Drawing.Point(549, 14);
            this.BtnChooseOutputPath.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.BtnChooseOutputPath.Name = "BtnChooseOutputPath";
            this.BtnChooseOutputPath.Size = new System.Drawing.Size(107, 39);
            this.BtnChooseOutputPath.TabIndex = 10;
            this.BtnChooseOutputPath.Text = "Choisir";
            this.BtnChooseOutputPath.UseVisualStyleBackColor = true;
            this.BtnChooseOutputPath.Click += new System.EventHandler(this.ChooseOutputPath);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 11);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(166, 17);
            this.label1.TabIndex = 9;
            this.label1.Text = "Dossier de sortie des pdf";
            // 
            // TxbOutputPath
            // 
            this.TxbOutputPath.Location = new System.Drawing.Point(15, 32);
            this.TxbOutputPath.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.TxbOutputPath.Name = "TxbOutputPath";
            this.TxbOutputPath.Size = new System.Drawing.Size(529, 22);
            this.TxbOutputPath.TabIndex = 8;
            // 
            // OFDTemplatePath
            // 
            this.OFDTemplatePath.FileName = "Template.pdf";
            this.OFDTemplatePath.Filter = "Pdf Files|*.pdf";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(15, 83);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(157, 17);
            this.label2.TabIndex = 12;
            this.label2.Text = "Chemin du template pdf";
            // 
            // SettingsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(668, 194);
            this.Controls.Add(this.BtnSave);
            this.Controls.Add(this.BtnChooseTemplatePath);
            this.Controls.Add(this.TxbTemplatePath);
            this.Controls.Add(this.BtnChooseOutputPath);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.TxbOutputPath);
            this.Controls.Add(this.label2);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "SettingsForm";
            this.Text = "Settings";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button BtnSave;
        private System.Windows.Forms.Button BtnChooseTemplatePath;
        private System.Windows.Forms.TextBox TxbTemplatePath;
        private System.Windows.Forms.Button BtnChooseOutputPath;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox TxbOutputPath;
        private System.Windows.Forms.OpenFileDialog OFDTemplatePath;
        private System.Windows.Forms.FolderBrowserDialog FBDOutputPath;
        private System.Windows.Forms.Label label2;
    }
}