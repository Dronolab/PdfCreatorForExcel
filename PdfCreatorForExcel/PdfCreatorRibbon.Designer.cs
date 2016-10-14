namespace PdfCreatorForExcel
{
    partial class PdfCreatorRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public PdfCreatorRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

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

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.PdfTab = this.Factory.CreateRibbonTab();
            this.PdfCreation = this.Factory.CreateRibbonGroup();
            this.BtnCreatePdf = this.Factory.CreateRibbonButton();
            this.BtnSettings = this.Factory.CreateRibbonButton();
            this.PdfTab.SuspendLayout();
            this.PdfCreation.SuspendLayout();
            this.SuspendLayout();
            // 
            // PdfTab
            // 
            this.PdfTab.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.PdfTab.Groups.Add(this.PdfCreation);
            this.PdfTab.Label = "Pdf";
            this.PdfTab.Name = "PdfTab";
            // 
            // PdfCreation
            // 
            this.PdfCreation.Items.Add(this.BtnCreatePdf);
            this.PdfCreation.Items.Add(this.BtnSettings);
            this.PdfCreation.Label = "Création de Pdf";
            this.PdfCreation.Name = "PdfCreation";
            // 
            // BtnCreatePdf
            // 
            this.BtnCreatePdf.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.BtnCreatePdf.Image = global::PdfCreatorForExcel.Properties.Resources.pdf_icon_2;
            this.BtnCreatePdf.Label = "Créer le Pdf";
            this.BtnCreatePdf.Name = "BtnCreatePdf";
            this.BtnCreatePdf.ShowImage = true;
            // 
            // BtnSettings
            // 
            this.BtnSettings.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.BtnSettings.Image = global::PdfCreatorForExcel.Properties.Resources.settings;
            this.BtnSettings.Label = "Settings";
            this.BtnSettings.Name = "BtnSettings";
            this.BtnSettings.ShowImage = true;
            // 
            // PdfCreatorRibbon
            // 
            this.Name = "PdfCreatorRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.PdfTab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.PdfCreatorRibbon_Load);
            this.PdfTab.ResumeLayout(false);
            this.PdfTab.PerformLayout();
            this.PdfCreation.ResumeLayout(false);
            this.PdfCreation.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab PdfTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup PdfCreation;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnCreatePdf;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnSettings;
    }

    partial class ThisRibbonCollection
    {
        internal PdfCreatorRibbon PdfCreatorRibbon
        {
            get { return this.GetRibbon<PdfCreatorRibbon>(); }
        }
    }
}
