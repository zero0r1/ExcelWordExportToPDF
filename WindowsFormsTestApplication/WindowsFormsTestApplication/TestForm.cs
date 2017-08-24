using Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;


namespace WindowsFormsTestApplication
{
    public partial class TestForm : Form
    {
        public TestForm()
        {
            InitializeComponent();
        }

        List<string> lstAddress = new List<string>();

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            openFileDialog1.FileName = string.Empty;
            Excel.Sheets oSheets = null;
            Excel.Workbook excelWorkBook = null;
            Excel.ApplicationClass application = null;
            object paramMissing = Type.Missing;

            try
            {
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    this.labSourceExcel.Text = openFileDialog1.FileName;
                }

                application = new Excel.ApplicationClass();
                excelWorkBook = application.Workbooks.Open(openFileDialog1.FileName);
                oSheets = excelWorkBook.Worksheets;

                this.checkedListBox1.Items.Clear();
                foreach (Excel.Worksheet oSheet in oSheets)
                {
                    this.checkedListBox1.Items.Add(oSheet.Name);
                }
            }
            catch
            {

            }
            finally
            {
                if (excelWorkBook != null)
                {
                    excelWorkBook.Close(false, paramMissing, paramMissing);
                    excelWorkBook = null;
                }
                if (application != null)
                {
                    application.Quit();
                    application = null;
                }

                //ms solution is like this
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

        }

        private void btnWord_Click(object sender, EventArgs e)
        {
            //"Excel Files|*.xls;*.xlsx;*.xlsm|Word Files|*.docx"
            openFileDialog1.Filter = "Word Files|*.docx";
            openFileDialog1.FileName = string.Empty;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                this.labSourceWord.Text = openFileDialog1.FileName;
            }
        }

        private void button1_Click_2(object sender, EventArgs e)
        {
            XLSConvertToPDFEach(this.labSourceExcel.Text);
        }

        private void btnConvertPDF_Click(object sender, EventArgs e)
        {
            XLSConvertToPDF(this.labSourceExcel.Text);
        }

        private void button1_Click_1(object sender, EventArgs e)
        {

            DOCXConvertToPDF(this.labSourceWord.Text);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            XLSConvertToPDFMarge(this.labSourceExcel.Text);

        }

        bool XLSConvertToPDF(string sourceFile)
        {
            string sourcePath = sourceFile;
            string targetPath = GetFileName(sourceFile);
            bool result = false;
            Excel.XlFixedFormatType targetType = Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF;// Excel.XlFixedFormatType.xlTypePDF;
            object paramMissing = Type.Missing;
            Excel.ApplicationClass application = null;
            Excel.Workbook excelWorkBook = null;
            try
            {
                application = new Excel.ApplicationClass();
                object target = targetPath;
                object type = targetType;
                excelWorkBook = application.Workbooks.Open(sourcePath);

                Excel.Worksheet xlSheet = (Excel.Worksheet)excelWorkBook.Worksheets[1];

                if (excelWorkBook != null)
                {
                    excelWorkBook.ExportAsFixedFormat(targetType, target, Excel.XlFixedFormatQuality.xlQualityStandard, true, true, paramMissing, paramMissing, false, paramMissing);
                    result = true;
                }
            }
            catch
            {
                result = false;
            }
            finally
            {
                if (excelWorkBook != null)
                {
                    excelWorkBook.Close(false, paramMissing, paramMissing);
                    excelWorkBook = null;
                }
                if (application != null)
                {
                    application.Quit();
                    application = null;
                }

                //ms solution is like this
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            return result;
        }

        bool XLSConvertToPDFEach(string sourceFile)
        {
            string sourcePath = sourceFile;
            string targetPath = GetFileName(sourceFile);
            bool result = false;
            Excel.XlFixedFormatType targetType = Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF;// Excel.XlFixedFormatType.xlTypePDF;
            object paramMissing = Type.Missing;
            Excel.ApplicationClass application = null;
            Excel.Workbook excelWorkBook = null;
            Excel.Worksheet excelWorksheet = null;
            try
            {
                application = new Excel.ApplicationClass();
                object target = targetPath;
                object type = targetType;
                excelWorkBook = application.Workbooks.Open(sourcePath);
                List<string> lstSheets = new List<string>();

                foreach (string item in this.checkedListBox1.CheckedItems)
                {
                    excelWorksheet = (Excel.Worksheet)excelWorkBook.Worksheets[item];
                    if (excelWorksheet != null)
                    {
                        excelWorksheet.ExportAsFixedFormat(targetType, target + "_" + excelWorksheet.Name, Excel.XlFixedFormatQuality.xlQualityStandard, true, paramMissing, paramMissing, paramMissing, false, paramMissing);
                        result = true;
                    }
                }
            }
            catch
            {
                result = false;
            }
            finally
            {
                if (excelWorkBook != null)
                {
                    excelWorkBook.Close(false, paramMissing, paramMissing);
                    excelWorkBook = null;
                }
                if (application != null)
                {
                    application.Quit();
                    application = null;
                }

                //ms solution is like this
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            return result;

        }

        bool XLSConvertToPDFMarge(string sourceFile)
        {
            string sourcePath = sourceFile;
            string targetPath = GetFileName(sourceFile);
            bool result = false;
            Excel.XlFixedFormatType targetType = Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF;// Excel.XlFixedFormatType.xlTypePDF;
            object paramMissing = Type.Missing;
            Excel.ApplicationClass application = null;
            Excel.Workbook excelWorkBook = null;
            Excel.Worksheet excelWorksheet = null;
            Excel.Worksheet sheetDelete = null;
            try
            {
                application = new Excel.ApplicationClass();
                object target = targetPath;
                object type = targetType;
                excelWorkBook = application.Workbooks.Open(sourcePath);
                List<string> lstSheets = new List<string>();

                for (int i = 0; i < this.checkedListBox1.Items.Count; i++)
                {
                    if (this.checkedListBox1.GetItemCheckState(i) == CheckState.Unchecked)
                    {
                        string deleteSheetName = this.checkedListBox1.Items[i].ToString();
                        sheetDelete = (Excel.Worksheet)excelWorkBook.Worksheets[deleteSheetName];

                        if (sheetDelete != null)
                        {
                            sheetDelete.Delete();
                        }

                        sheetDelete = null;
                    }
                }

                excelWorkBook.ExportAsFixedFormat(targetType, target, Excel.XlFixedFormatQuality.xlQualityStandard, true, true, paramMissing, paramMissing, false, paramMissing);
                result = true;
            }
            catch
            {
                result = false;
            }
            finally
            {
                if (excelWorkBook != null)
                {
                    excelWorkBook.Close(false, paramMissing, paramMissing);
                    excelWorkBook = null;
                }
                if (application != null)
                {
                    application.Quit();
                    application = null;
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            return result;
        }

        bool DOCXConvertToPDF(string sourceFile)
        {
            bool result = false;
            object sourcePath = sourceFile;
            object targetPath = GetFileName(sourceFile);
            object targetType = Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF; //Word.WdExportFormat.wdExportFormatPDF;

            object paramMissing = Type.Missing;
            Word.ApplicationClass wordApplication = null;
            Word.Document wordDocument = null;
            try
            {
                wordApplication = new Word.ApplicationClass();
                wordDocument = wordApplication.Documents.Open(ref sourcePath);

                if (wordDocument != null)
                {
                    wordDocument.SaveAs(targetPath, targetType);
                    result = true;
                }
            }
            catch
            {
                result = false;
            }
            finally
            {
                if (wordDocument != null)
                {
                    wordDocument.Close(ref paramMissing, ref paramMissing, ref paramMissing);
                    wordDocument = null;
                }
                if (wordApplication != null)
                {
                    wordApplication.Quit(ref paramMissing, ref paramMissing, ref paramMissing);
                    wordApplication = null;
                }

                //ms solution is like this
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            return result;
        }



        string GetFileName(string pathString)
        {
            return Path.ChangeExtension(pathString, @".pdf");
        }
    }
}
