using System;

namespace ConsoleApplication1
{
    class Program
    {
        static void Main(string[] args)
        {
            CombineWord("E:/2.docx", "E:/3.docx", "E:/4.docx");
        }
        public static void CombineWord(string fileToMerge1, string fileToMerge2, string outputFilename)
        {
            object missing = System.Type.Missing;
            object pageBreak = Word.WdBreakType.wdPageBreak;
            object outputFile = outputFilename;
            // Create  a new Word application
            Word._Application wordApplication = new Word.Application();
            try
            {
                
                Word._Document wordDocument = wordApplication.Documents.Open(
                                            fileToMerge1
                                            , ref missing
                                            , ref missing
                                            , ref missing
                                            , ref missing);
                
                Word.Selection selection = wordApplication.Selection;
                wordDocument.Merge(fileToMerge2, ref missing, ref missing, ref missing, ref missing);
                // Save the document to it's output file.
                wordDocument.SaveAs(
                            ref outputFile
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing);
                
                wordDocument = null;
            }
            catch (Exception ex)
            {
                //I didn't include a default error handler so i'm just throwing the error
                throw ex;
            }
            finally
            {
                
                wordApplication.Quit(ref missing, ref missing, ref missing);
            }
        }    }
}
