using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using PdfCreatorForExcel.Properties;

namespace PdfCreatorForExcel
{
    public partial class SettingsForm : Form
    {
        public SettingsForm()
        {
            InitializeComponent();
            TxbOutputPath.Text = Settings.Default.FolderOutputPath;
            TxbTemplatePath.Text = Settings.Default.TemplatePath;
        }

        private void ChooseOutputPath(object sender, EventArgs e)
        {
            if (FBDOutputPath.ShowDialog() == DialogResult.OK)
            {
                TxbOutputPath.Text = FBDOutputPath.SelectedPath;
            }
        }

        private void ChooseTemplatePath(object sender, EventArgs e)
        {
            if (OFDTemplatePath.ShowDialog() == DialogResult.OK)
            {
                TxbTemplatePath.Text = OFDTemplatePath.InitialDirectory + OFDTemplatePath.FileName;
            }
        }

        private void SaveSettings(object sender, EventArgs e)
        {
            Settings.Default.FolderOutputPath = TxbOutputPath.Text;
            Settings.Default.TemplatePath = TxbTemplatePath.Text;
            Settings.Default.Save();
            Close();
        }
    }
}
