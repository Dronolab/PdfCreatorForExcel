using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.IO;
using PdfSharp.Pdf.IO;
using PdfSharp.Pdf.AcroForms;
using PdfSharp.Pdf;
using Microsoft.Office.Tools.Ribbon;
using PdfCreatorForExcel.Properties;

namespace PdfCreatorForExcel
{
    public partial class ThisAddIn
    {
        private PdfAcroForm _form;
        private Excel.Worksheet _currentWorkSheet;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            Globals.Ribbons.PdfCreatorRibbon.BtnCreatePdf.Click += Validate;
            Globals.Ribbons.PdfCreatorRibbon.BtnCreatePdf.Click += CreatePdf;
            Globals.Ribbons.PdfCreatorRibbon.BtnSettings.Click += OpenSettingsBox;
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
        }

        #region PDF creation

        private void CreatePdf(object sender, RibbonControlEventArgs e)
        {
            if (Settings.Default.FolderOutputPath.Length == 0 || Settings.Default.TemplatePath.Length == 0)
            {
                OpenSettingsBox(null, null);
                return;
            }

            // Get Worksheet
            _currentWorkSheet = Application.ActiveSheet;

            var pdfDocument = LoadPdf();

            if (pdfDocument == null)
                return;

            WriteAllCells();
            WriteAllCheckBoxes();
            WriteAllGroup();

            //Save pdf
            try
            {
                var path = Path.Combine(Settings.Default.FolderOutputPath, DateTime.Today.ToString("dd-MM-yyyy") + GetCodePermanent() + ".pdf");

                pdfDocument.Save(path);

                if (GetCodePermanent().Equals(" Undefined"))
                    MessageBox.Show(Resources.CreatePdf_CodePermanentEmpty_MessageBoxText);
            }
            catch (Exception ex)
            {
                MessageBox.Show(Resources.CreatePdf_SavePdfDocument_Exception_MessageBoxText + ex.StackTrace);
            }
            // Put values in pdf
            MessageBox.Show(Resources.PdfCreated);
        }

        private PdfDocument LoadPdf()
        {
            // Load PDF
            try
            {
                File.Copy(Settings.Default.TemplatePath,
                    Path.Combine(Settings.Default.FolderOutputPath,
                        DateTime.Today.ToString("dd-MM-yyyy") + GetCodePermanent() + ".pdf"), true);
            }
            catch (UnauthorizedAccessException)
            {
                MessageBox.Show(Resources.LoadPdf_CopyTemplate_UnauthorizedAccessException_MessageBoxText);
                return null;
            }
            catch (ArgumentNullException)
            {
                MessageBox.Show(Resources.LoadPdf_CopyTemplate_ArgumentNullException_MessageBoxText);
                OpenSettingsBox(null, null);
                return null;
            }
            catch (DirectoryNotFoundException)
            {
                MessageBox.Show(Resources.LoadPdf_CopyTemplate_DirectoryNotFoundException_MessageBoxText);
                OpenSettingsBox(null, null);
                return null;
            }
            catch (FileNotFoundException)
            {
                MessageBox.Show(Resources.LoadPdf_CopyTemplate_FileNotFoundException_MessageBoxText);
                OpenSettingsBox(null, null);
                return null;
            }
            catch (IOException)
            {
                MessageBox.Show(Resources.LoadPdf_CopyTemplate_IOException_MessageBoxText);
                return null;
            }
            catch (Exception ex)
            {
                MessageBox.Show(Resources.LoadPdf_CopyTemplate_Exception_MessageBoxText + ex.StackTrace);
                return null;
            }

            PdfDocument pdfDocument = null;
            // Open pdf
            try
            {
                pdfDocument = PdfReader.Open(Settings.Default.TemplatePath, PdfDocumentOpenMode.Modify);
                _form = pdfDocument.AcroForm;
            }
            catch (Exception ex)
            {
                MessageBox.Show(Resources.LoadPdf_OpenTemplate_Exception_MessageBoxText + ex.StackTrace);
            }

            if (_form.Elements.ContainsKey("/NeedAppearances"))
            {
                _form.Elements["/NeedAppearances"] = new PdfBoolean(true);
            }
            else
            {
                _form.Elements.Add("/NeedAppearances", new PdfBoolean(true));
            }

            return pdfDocument;
        }

        private string GetCodePermanent()
        {
            if (!string.IsNullOrWhiteSpace(GetCellValue(4, "K")))
                return " " + GetCellValue(4, "K");

            return " Undefined";
        }

        private void WriteAllCells()
        {
            WriteStringCell(1, "B", "NoRef");
            WriteStringCell(3, "B", "Description");
            WriteStringCell(4, "B", "Avance");

            if (GetCellValue(31, "K").Equals("0"))
            {
                WriteStringCell(7, "A", "Détails1");
                WriteStringCell(8, "A", "Détails2");
                WriteStringCell(9, "A", "Détails3");
                WriteStringCell(10, "A", "Détails4");
                WriteStringCell(11, "A", "Détails5");
                WriteStringCell(12, "A", "Détails6");
                WriteStringCell(13, "A", "Détails7");
                WriteStringCell(14, "A", "Détails8");

                WriteStringCell(7, "B", "Montant$1");
                WriteStringCell(8, "B", "Montant$2");
                WriteStringCell(9, "B", "Montant$3");
                WriteStringCell(10, "B", "Montant$4");
                WriteStringCell(11, "B", "Montant$5");
                WriteStringCell(12, "B", "Montant$6");
                WriteStringCell(13, "B", "Montant$7");
                WriteStringCell(14, "B", "Montant$8");
            }
            else
            {
                // Get the field from the PDF
                var currentField = (PdfTextField)(_form.Fields["Détails1"]);
                currentField.ReadOnly = false;
                currentField.Value = new PdfString(Settings.Default.Plus8Article);
            }

            var typeRemb = (PdfRadioButtonField)(_form.Fields["Group1"]); // Values  
            typeRemb.ReadOnly = false;
            if (!string.IsNullOrWhiteSpace(GetCellValue(1, "G")))
            {
                WriteStringCell(1, "H", "Destination");
                WriteDateCell(2, "H", "DébutActivité"); // DateTime
                WriteDateCell(3, "H", "FinActivité"); // DateTime
                WriteDateCell(4, "H", "Départ"); // DateTime
                WriteDateCell(5, "H", "Retour"); // DateTime
                WriteStringCell(6, "H", "Présentation");
                WriteStringCell(7, "H", "Conférence");
                WriteStringCell(8, "H", "Autres");
                WriteStringCell(9, "H", "KM");

                typeRemb.Value = new PdfName("/Choix1");
            }
            else
            {
                typeRemb.Value = new PdfName("/Choix2");
            }

            WriteStringCell(1, "K", "Prenom");
            WriteStringCell(2, "K", "Nom");
            WriteStringCell(3, "K", "NomUA");
            WriteStringCell(4, "K", "CodePermanent");
            WriteDateCell(5, "K", "Date"); // DateTime
            WriteStringCell(6, "K", "Tél. demandeur");
            WriteStringCell(7, "K", "Tel requérant");
            WriteStringCell(8, "K", "Tel Sup hiérarchique");
            WriteStringCell(9, "K", "Réclamation$");
            WriteStringCell(13, "K", "Adresse");
            WriteStringCell(14, "K", "Province");
            WriteStringCell(15, "K", "CodePostal");
            WriteStringCell(16, "K", "Ville");
            WriteStringCell(17, "K", "Courriel");
            WriteStringCell(18, "K", "TotalMontant");
            WriteStringCell(19, "K", "ccMontant$1");
            WriteStringCell(20, "K", "ccMontant$2");
            WriteStringCell(21, "K", "ccMontant$3");
            WriteStringCell(22, "K", "ccMontant$4");
            WriteStringCell(23, "K", "ccMontant$5");
            WriteStringCell(18, "K", "TotalccMontant$");
            WriteStringCell(24, "K", "CC1");
            WriteStringCell(25, "K", "CC2");
            WriteStringCell(26, "K", "CC3");
            WriteStringCell(27, "K", "CC4");
            WriteStringCell(28, "K", "CC5");

            if (!string.IsNullOrWhiteSpace(GetCellValue(24, "K")))
                WriteStringCell(29, "K", "UBR1");

            if (!string.IsNullOrWhiteSpace(GetCellValue(25, "K")))
                WriteStringCell(29, "K", "UBR2");

            if (!string.IsNullOrWhiteSpace(GetCellValue(26, "K")))
                WriteStringCell(29, "K", "UBR3");

            if (!string.IsNullOrWhiteSpace(GetCellValue(27, "K")))
                WriteStringCell(29, "K", "UBR4");

            if (!string.IsNullOrWhiteSpace(GetCellValue(28, "K")))
                WriteStringCell(29, "K", "UBR5");
        }

        private void WriteAllGroup()
        {
            // Group
            var emplEtu = (PdfRadioButtonField)(_form.Fields["Group2"]);
            emplEtu.ReadOnly = false;
            emplEtu.Value = new PdfName("/Choix2");
            var depotcheque = (PdfRadioButtonField)(_form.Fields["Group4"]); //Values cheque/depot
            depotcheque.ReadOnly = false;

            var type = GetCellValue(30, "K");
            if (!string.IsNullOrWhiteSpace(type))
                depotcheque.Value = string.CompareOrdinal(type, "depot") == 0 ? new PdfName("/Dépôt") : new PdfName("/Chèque");
        }

        private void WriteAllCheckBoxes()
        {
            // Box
            var b1 = !string.IsNullOrWhiteSpace(GetCellValue(6, "H"));
            var b2 = !string.IsNullOrWhiteSpace(GetCellValue(7, "H"));
            var b3 = !string.IsNullOrWhiteSpace(GetCellValue(8, "H"));
            FillCheckBox(b1, "Boite1");
            FillCheckBox(b2, "Boite2");
            FillCheckBox(b3, "Boite4");
            if(!string.IsNullOrWhiteSpace(GetCellValue(1, "G")))
                FillCheckBox(b1 == b2 == b3 == false, "Boite3");
        }

        private void FillCheckBox(bool check, string fieldName)
        {
            // Get the field from the PDF
            var currentField = (PdfCheckBoxField)(_form.Fields[fieldName]);
            currentField.ReadOnly = false;
            currentField.Checked = check;
        }

        private void WriteStringCell(int nombre, string lettre, string fieldName)
        {
            // Get the field from the PDF
            var currentField = (PdfTextField)(_form.Fields[fieldName]);
            currentField.ReadOnly = false;
            // Get the value of the field from Excel
            var textFromExcel = GetCellValue(nombre, lettre);
            // Write the Excel value to the PDF field
            var textToPdf = new PdfString(textFromExcel);
            currentField.Value = textToPdf;
        }

        private void WriteDateCell(int nombre, string lettre, string fieldName)
        {
            // Get the field from the PDF
            var currentField = (PdfTextField)(_form.Fields[fieldName]);
            currentField.ReadOnly = false;
            // Get the value of the field from Excel
            var textFromExcel = GetCellValue(nombre, lettre);
            textFromExcel = DateTime.FromOADate(Convert.ToDouble(textFromExcel)).ToString("yyyy-MM-dd");

            // Write the Excel value to the PDF field
            var textToPdf = new PdfString(textFromExcel);
            currentField.Value = textToPdf;
        }

        private string GetCellValue(int number, string letter)
        {
            return ((Excel.Range)_currentWorkSheet.Cells[number, letter]).Value2 + "";
        }

        #endregion

        #region PDF settings

        private void OpenSettingsBox(object sender, RibbonControlEventArgs e)
        {
            new SettingsForm().Show();
        }

        #endregion

        #region Excel

        private void Validate(object sender, RibbonControlEventArgs e)
        {
            Application.Run("calculsstotaux");
        }

        #endregion

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }

        #endregion
    }
}
