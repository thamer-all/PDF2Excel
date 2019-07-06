using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using iTextSharp.text.pdf;
using iTextSharp.text;
using iTextSharp.text.pdf.parser;
using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace PDF2Excel
{
    public partial class Form1 : Form
    {
        System.String file_name = "";
        System.String excel_name = "";
        Int32 nCnt;
        public Form1()
        {
            InitializeComponent();
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog
            {
                InitialDirectory = @"D:\",
                Title = "Browse PDF Files",

                CheckFileExists = true,
                CheckPathExists = true,

                DefaultExt = "pdf",
                Filter = "pdf files (*.pdf)|*.pdf",
                FilterIndex = 2,
                RestoreDirectory = true,

                ReadOnlyChecked = true,
                ShowReadOnly = true
            };

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = openFileDialog1.FileName;
                file_name = textBox1.Text;
                System.String[] nameList = file_name.Split('\\');
                nCnt = nameList.Length;
                //textBox2.Text = nameList[nCnt - 1];
                System.String pdf_name = nameList[nCnt - 1];
                excel_name = pdf_name.Replace("pdf", "xls");
                textBox2.Text = excel_name;
            }
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            Int32 nCount = 0;
            System.String str_result_text = "";
            nCount = SplitPDF(file_name);
            System.String str_dir_path = GetDirPath(file_name);
            str_dir_path = str_dir_path + "result\\";
            str_result_text = ExtractTextFromPdf(str_dir_path, nCount, excel_name);
            //SaveData2Excel(str_result_text);
            file_name = file_name + excel_name;
            textBox2.Text = GetDirPath(file_name);
        }
        public static string ExtractTextFromPdf(string path, Int32 nCount, string excel_name)
        {
            System.String str_fname = "";
            List<string> ContentList = new List<string>();
            for (int i = 1; i <= nCount; i++)
            {
                str_fname = path + i + ".pdf";
                if (!File.Exists(str_fname))
                {
                    continue;
                }
                    
                    using (PdfReader reader = new PdfReader(str_fname))
                {
                    StringBuilder text = new StringBuilder();
                    text.Append(PdfTextExtractor.GetTextFromPage(reader, 1));
                    System.String content = text.ToString();
                    ContentList.Add(content);
                }
            }

            SaveData2Excel(ContentList, path, nCount, excel_name);

            for (int i = 1; i <= 200; i++)
            {
                if (File.Exists(path + i + ".pdf"))
                {
                    // If file found, delete it    
                    File.Delete(path + i + ".pdf");
                }
            }
        
            return "success";
        }

        public static void SaveData2Excel(List<string> content_list, string dest_path, Int32 idx, string excel_name)
        {
            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            System.String xls_fname = dest_path + excel_name;
            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
                return;
            }

            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            xlWorkSheet.Cells[1, 1] = "County";
            xlWorkSheet.Cells[1, 2] = "Assessed to";
            xlWorkSheet.Cells[1, 3] = "C/S #";
            xlWorkSheet.Cells[1, 4] = "Date of Sale to State";
            xlWorkSheet.Cells[1, 5] = "Parcel #";
            xlWorkSheet.Cells[1, 6] = "Price #";
            xlWorkSheet.Cells[1, 7] = "valid through date";

            for (int i = 0; i < idx; i++)
            {
                xlWorkSheet.Cells[i + 2, 1] = getCounty(content_list[i]);
                xlWorkSheet.Cells[i + 2, 2] = getAssesed2Name(content_list[i]);
                xlWorkSheet.Cells[i + 2, 3] = getCS(content_list[i]);
                xlWorkSheet.Cells[i + 2, 4] = getDateofSale(content_list[i]);
                xlWorkSheet.Cells[i + 2, 5] = getParcel(content_list[i]);
                xlWorkSheet.Cells[i + 2, 6] = getPrice(content_list[i]);
                xlWorkSheet.Cells[i + 2, 7] = getValidThroughDate(content_list[i]);
            }


            xlWorkBook.SaveAs(xls_fname, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

            MessageBox.Show("Excel file created , you can find the file" + xls_fname);
        }

        public static Int32 SplitPDF(String str_src_path)
        {
            //variables
            String str_dir_path = "";
            String result = "";
            PdfCopy copy;

            str_dir_path = GetDirPath(str_src_path);
            str_dir_path = str_dir_path + "result\\";

            if (!Directory.Exists(str_dir_path))
            {
                Directory.CreateDirectory(str_dir_path);
            }

            //create PdfReader object
            PdfReader reader = new PdfReader(str_src_path);

            for (int i = 1; i <= reader.NumberOfPages; i++)
            {
                //create Document object
                Document document = new Document();
                copy = new PdfCopy(document, new FileStream(str_dir_path + i + ".pdf", FileMode.Create));
                //open the document 
                document.Open();
                //add page to PdfCopy 
                copy.AddPage(copy.GetImportedPage(reader, i));
                //close the document object
                document.Close();
            }

            return reader.NumberOfPages;
        }

        public static String GetDirPath(String src)
        {
            String rst = "";
            String filename;
            Int32 nCount;

            System.String[] nameList = src.Split('\\');
            nCount = nameList.Length;
            //filename = nCount.ToString();
            filename = nameList[nCount - 1];
            rst = src.Replace(filename, "");

            return rst;
        }

        public static String getCounty(string content)
        {
            System.String str_result = "";

            int firstStringPosition = content.IndexOf("County:");
            int secondStringPosition = content.IndexOf("Assessed to:");
            str_result = content.Substring(firstStringPosition, secondStringPosition - firstStringPosition);
            Int32 nlen = str_result.Length;
            str_result = str_result.Substring(7, nlen - 7);
            return str_result;
        }
        public static String getAssesed2Name(string content)
        {
            System.String str_result = "";

            int firstStringPosition = content.IndexOf("Assessed to:");
            int secondStringPosition = content.IndexOf("C/S #:");
            str_result = content.Substring(firstStringPosition, secondStringPosition - firstStringPosition);
            Int32 nlen = str_result.Length;
            str_result = str_result.Substring(12, nlen - 12);
            return str_result;
        }
        public static String getCS(string content)
        {
            System.String str_result = "";

            int firstStringPosition = content.IndexOf("C/S #:");
            int secondStringPosition = content.IndexOf("Date of Sale to State:");
            str_result = content.Substring(firstStringPosition, secondStringPosition - firstStringPosition);
            Int32 nlen = str_result.Length;
            str_result = str_result.Substring(6, nlen - 6);
            return str_result;
        }
        public static String getDateofSale(string content)
        {
            System.String str_result = "";

            int firstStringPosition = content.IndexOf("Date of Sale to State:");
            int secondStringPosition = content.IndexOf("Dear Applicant:");
            str_result = content.Substring(firstStringPosition, secondStringPosition - firstStringPosition);
            Int32 nlen = str_result.Length;
            str_result = str_result.Substring(22, nlen - 22);
            return str_result;
        }
        public static String getParcel(string content)
        {
            System.String str_result = "";

            int firstStringPosition = content.IndexOf("as follows:");
            int secondStringPosition = content.IndexOf("Parcel #:");
            str_result = content.Substring(firstStringPosition, secondStringPosition - firstStringPosition);
            Int32 nlen = str_result.Length;
            str_result = str_result.Substring(11, nlen - 11);
            return str_result;
        }
        public static String getPrice(string content)
        {
            System.String str_result = "";

            int firstStringPosition = content.IndexOf("$");
            int secondStringPosition = content.IndexOf("The above");
            str_result = content.Substring(firstStringPosition, secondStringPosition - firstStringPosition);
            Int32 nlen = str_result.Length;
            str_result = str_result.Substring(1, nlen - 1);
            return str_result;
        }
        public static String getValidThroughDate(string content)
        {
            System.String str_result = "";

            int firstStringPosition = content.IndexOf(" valid through");
            int secondStringPosition = content.IndexOf("Sincerely,");
            str_result = content.Substring(firstStringPosition, secondStringPosition - firstStringPosition);
            Int32 nlen = str_result.Length;
            str_result = str_result.Substring(14, 11);
            return str_result;
        }
    }
}