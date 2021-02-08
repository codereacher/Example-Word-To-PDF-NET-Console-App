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
            string sourceFolderPath = ConfigurationManager.AppSettings["Word_Folder_Path"];
            string outputFolderPath = ConfigurationManager.AppSettings["Pdf_Folder_Path"];

            ConvertWordToPdf(sourceFolderPath, "docx", outputFolderPath);
        }

        /// <summary>
        /// Converts Microsoft Word files to PDF format.
        /// </summary>
        /// <param name="sourceFolderPath">Path to source folder containing Microsoft Word files.</param>
        /// <param name="fileExtension">File extension of Microsoft Word files (e.g., doc, docx) in source folder.</param>
        /// <param name="outputFolderPath">Path to output folder where PDF files will be saved.</param>
        private static void ConvertWordToPdf(string sourceFolderPath, string fileExtension, string outputFolderPath)
        {
            // Create a new instance of Microsoft Word application object
            Application wordApp = new Application();
            FileInfo[] wordFiles = null;

            // Use dummy value since C# does not have optional arguments
            object oMissing = System.Reflection.Missing.Value;

            try
            {
                // Get list of Word files having the specified file extension
                DirectoryInfo dirInfo = new DirectoryInfo(sourceFolderPath);
                wordFiles = dirInfo.GetFiles("*." + fileExtension);

                wordApp.Visible = false;
                wordApp.ScreenUpdating = false;

                for (int i = 0; i < wordFiles.Length; i++)
                {
                    ConsoleWriteLine("Doing " + Convert.ToInt32(i + 1) + "/" + wordFiles.Length + " | File: " + wordFiles[i].Name);

                    FileInfo wordFile = wordFiles[i];

                    if (wordFile.Name.Contains("~"))
                    {
                        ConsoleWriteLine("Skipping file since it contains '~' in the file name");
                        continue;
                    }

                    object wordFileName = wordFile.FullName;

                    // Use dummy value as a placeholder for optional arguments
                    Document wordDoc = wordApp.Documents.Open(ref wordFileName, ref oMissing,
                        ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                        ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                        ref oMissing, ref oMissing, ref oMissing, ref oMissing);

                    // Use same file name except replace file extension with "pdf"
                    object outputFilePath = outputFolderPath + wordFile.Name.Replace("." + fileExtension, ".pdf");
                    object fileFormat = WdSaveFormat.wdFormatPDF;

                    // Save as PDF file
                    wordDoc.SaveAs(ref outputFilePath, ref fileFormat, ref oMissing, ref oMissing,
                        ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                        ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                        ref oMissing, ref oMissing, ref oMissing, ref oMissing);

                    ConsoleWriteLine("Saved as PDF at " + outputFilePath);

                    // Close Word file without saving changes
                    object saveChanges = WdSaveOptions.wdDoNotSaveChanges;
                    wordDoc.Close(ref saveChanges, ref oMissing, ref oMissing);
                    wordDoc = null;
                }

                wordApp.Quit(ref oMissing, ref oMissing, ref oMissing);
                wordApp = null;
                wordFiles = null;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);

                wordApp.Quit(ref oMissing, ref oMissing, ref oMissing);
                wordApp = null;
                wordFiles = null;
            }
        }

        private static void ConsoleWriteLine(string message)
        {
            Console.WriteLine(FormatDateTime(DateTime.Now, "MM/dd/yyyy", "HH:mm:ss.fff") + ": " + message);
        }

        private static string FormatDateTime(DateTime currentDateTime, string dateFormat, string timeFormat)
        {
            return currentDateTime.ToString(dateFormat) + " " + currentDateTime.ToString(timeFormat);
        }
    }
}
