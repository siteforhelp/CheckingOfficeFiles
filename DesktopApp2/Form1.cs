using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

// This is the code for your desktop app.
// Press Ctrl+F5 (or go to Debug > Start Without Debugging) to run your app.

namespace DesktopApp2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // Click on the link below to continue learning how to build a desktop app using WinForms!
            System.Diagnostics.Process.Start("http://aka.ms/dotnet-get-started-desktop");

        }

        private void button1_Click(object sender, EventArgs e)
        {
            //MessageBox.Show("Thanks!");
            RunOpenXmlValidation("D:\\Временное\\СО - 2.1 Центральный ПУЭС Дагестана.xlsx", "office2007");

        }

        public void RunOpenXmlValidation(string filePath, string openXmlFormatVersion)
        {
            string fileName = Path.GetFileName(filePath);
            StageName = String.Format("RUNNING OpenXmlValidation for FILE {0}", fileName);
            outputManager.BeginWriteInfoLine(String.Format("Running OpenXmlValidation for saved file '{0}'", fileName));
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, false))
            {
                DocumentFormat.OpenXml.FileFormatVersions formatVersion = DocumentFormat.OpenXml.FileFormatVersions.Office2010;
                if (openXmlFormatVersion == "office2007")
                    formatVersion = DocumentFormat.OpenXml.FileFormatVersions.Office2007;
                else if (openXmlFormatVersion == "office2013")
                    formatVersion = DocumentFormat.OpenXml.FileFormatVersions.Office2013;
                OpenXmlValidator validator = new OpenXmlValidator(formatVersion);
                var errors = validator.Validate(wordDoc);
                StringBuilder builder = new StringBuilder();
                foreach (ValidationErrorInfo error in errors)
                {
                    string errorMsg = string.Format("{0}: {1}, {2}, {3}", error.ErrorType.ToString(), error.Part.Uri, error.Path.XPath, error.Node.LocalName);
                    builder.AppendLine(errorMsg);
                    builder.AppendLine(error.Description);
                }
                string logContent = builder.ToString();
                if (!string.IsNullOrEmpty(logContent))
                    throw new FileFormatValidationFailedException(logContent);
            }
        }
    }
}
