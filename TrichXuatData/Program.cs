using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using System.Text;
using A = DocumentFormat.OpenXml.Drawing;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using DocumentFormat.OpenXml.Spreadsheet;
using EPocalipse.IFilter;
using LinqToExcel;
using TikaOnDotNet.TextExtraction;
using System.Data;
using System.Data.Entity;
using System.Data.SqlClient;


namespace TrichXuatData
{
    public class MyData : DbContext
    {
        public MyData()
            : base("Name=DACKContext")
        {

        }
    }
    class Program
    {

        private static string ReadPdfFile(string fileName)
        {
            var text = new StringBuilder();

            if (File.Exists(fileName))
            {
                var pdfReader = new PdfReader(fileName);

                for (int page = 1; page <= pdfReader.NumberOfPages; page++)
                {
                    var strategy = new SimpleTextExtractionStrategy();
                    string currentText = PdfTextExtractor.GetTextFromPage(pdfReader, page, strategy);

                    currentText = Encoding.UTF8.GetString(ASCIIEncoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.Default.GetBytes(currentText)));
                    text.Append(currentText);
                }
                pdfReader.Close();
            }
            return text.ToString();
        }

        static void Main(string[] args)
        {
            string[] files = Directory.GetFiles(@"E:\DATA");
            var textExtractor = new TextExtractor();
            MyData db = new MyData();
            foreach (var item in files)
            {
                switch (System.IO.Path.GetExtension(item))
                {
                    case ".pdf":
                        {
                            var text = ReadPdfFile(item);
                            SqlParameter param1 = new SqlParameter("@name", System.IO.Path.GetFileName(item));
                            SqlParameter param2 = new SqlParameter("@extension", System.IO.Path.GetExtension(item));
                            SqlParameter param3 = new SqlParameter("@content", text);
                            db.Database.ExecuteSqlCommand("SP_InsertFile @name, @extension, @content", param1, param2, param3);
                        }
                        break;
                    case ".doc":
                    case ".docx":
                        {
                            WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(item, true);
                            string text = wordprocessingDocument.MainDocumentPart.Document.InnerText;

                            SqlParameter param1 = new SqlParameter("@name", System.IO.Path.GetFileName(item));
                            SqlParameter param2 = new SqlParameter("@extension", System.IO.Path.GetExtension(item));
                            SqlParameter param3 = new SqlParameter("@content", text);
                            db.Database.ExecuteSqlCommand("SP_InsertFile @name, @extension, @content", param1, param2, param3);
                        }
                        break;
                    case "xls":
                    case "xlsx":
                        {
                            var text = textExtractor.Extract(item);
                            SqlParameter param1 = new SqlParameter("@name", System.IO.Path.GetFileName(item));
                            SqlParameter param2 = new SqlParameter("@extension", System.IO.Path.GetExtension(item));
                            SqlParameter param3 = new SqlParameter("@content", text);
                            db.Database.ExecuteSqlCommand("SP_InsertFile @name, @extension, @content", param1, param2, param3);
                        }
                        break;
                    case "ppt":
                    case "pptx":
                        {
                            var text = textExtractor.Extract(item);
                            SqlParameter param1 = new SqlParameter("@name", System.IO.Path.GetFileName(item));
                            SqlParameter param2 = new SqlParameter("@extension", System.IO.Path.GetExtension(item));
                            SqlParameter param3 = new SqlParameter("@content", text);
                            db.Database.ExecuteSqlCommand("SP_InsertFile @name, @extension, @content", param1, param2, param3);
                        }
                        break;
                    default:
                        break;
                }
            }
        }
    }
}