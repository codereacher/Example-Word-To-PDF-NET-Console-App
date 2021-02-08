using Microsoft.Office.Interop.Word;
using System;
using System.Configuration;
using System.IO;

namespace WordToPdfConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            string sourceFolderPath = ConfigurationManager.AppSettings["WordFolderPath"];
            string fileExtension = ConfigurationManager.AppSettings["WordFileExtension"];
            string outputFolderPath = ConfigurationManager.AppSettings["PdfFolderPath"];

            PdfConversion myPdfConversion = new PdfConversion();
            myPdfConversion.ConvertWordToPdf(sourceFolderPath, fileExtension, outputFolderPath);
        }
    }
}
