using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DsofileDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            string filePath = Console.ReadLine();// @"c:\temp\MyFile.pdf";
            printFileSummaryInfo(filePath);
            Console.ReadLine();
        }
        /// <summary>
        /// 打印文件摘要信息
        /// </summary>
        /// <param name="filePath"></param>
        private static void printFileSummaryInfo(string filePath)
        {
            try
            {
                DSOFile.OleDocumentProperties dso = new DSOFile.OleDocumentProperties();
                dso.Open(filePath, true, DSOFile.dsoFileOpenOptions.dsoOptionOpenReadOnlyIfNoWriteAccess);
                Console.WriteLine(dso.SummaryProperties.Title);
                Console.WriteLine(dso.SummaryProperties.Author);
                Console.WriteLine(dso.SummaryProperties.ByteCount);
                Console.WriteLine(dso.SummaryProperties.CharacterCount);
                Console.WriteLine(dso.SummaryProperties.Version);
                Console.WriteLine(dso.SummaryProperties.CharacterCountWithSpaces);
                Console.WriteLine(dso.SummaryProperties.Comments);
                Console.WriteLine(dso.SummaryProperties.Company);

                //Console.WriteLine(Convert.ToDateTime(dso.SummaryProperties.DateCreated));
                //Console.WriteLine(Convert.ToDateTime(dso.SummaryProperties.DateLastSaved));
                Console.WriteLine(dso.SummaryProperties.DateCreated);
                Console.WriteLine(dso.SummaryProperties.DateLastSaved);
                Console.WriteLine(dso.SummaryProperties.LastSavedBy);
                Console.WriteLine(dso.SummaryProperties.LineCount);
                Console.WriteLine(dso.SummaryProperties.PageCount);
                Console.WriteLine(dso.SummaryProperties.ParagraphCount);
                Console.WriteLine(dso.SummaryProperties.RevisionNumber);
                Console.WriteLine(dso.SummaryProperties.Subject);
                Console.WriteLine(dso.SummaryProperties.WordCount);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            
        }
        /// <summary>
        /// 修改文件摘要信息
        /// </summary>
        /// <param name="filePath"></param>
        private static bool changFileSummaryInfo(string filePath)
        {
            try
            {
                //This is the PDF file we want to update.
                string filename = filePath;// @"c:\temp\MyFile.pdf";
                //Create the OleDocumentProperties object.
                DSOFile.OleDocumentProperties dso = new DSOFile.OleDocumentProperties();
                //Open the file for writing if we can. If not we will get an exception.
                dso.Open(filename, false, DSOFile.dsoFileOpenOptions.dsoOptionOpenReadOnlyIfNoWriteAccess);

                //Set the summary properties that you want.
                dso.SummaryProperties.Title = "This is the Title";
                dso.SummaryProperties.Subject = "This is the Subject";
                dso.SummaryProperties.Company = "RTDev";
                dso.SummaryProperties.Author = "Ron T.";
                //Save the Summary information.
                dso.Save();
                //Close the file.
                dso.Close(false);
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
