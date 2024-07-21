using ECMWordGenerator.Models;
using System.Collections.Generic;
using System.IO;
using Word = Microsoft.Office.Interop.Word;

namespace ECMWordGenerator.Services
{
    /// <summary>
    /// Class responsible for processing Word templates.
    /// </summary>
    public class WordTemplateProcessor
    {
        /// <summary>
        /// Replaces placeholders in the Word document with specified values and saves the result.
        /// </summary>
        /// <param name="documentPath">The path to the original Word document.</param>
        /// <param name="data">A list containing the placeholders and their replacement values.</param>
        /// <returns>The path to the generated Word document.</returns>
        public string ProcessTemplate(string documentPath, List<Item> data)
        {
            // Declare variables for the Word application and document
            Word.Application wordApp = null;
            Word.Document doc = null;

            try
            {
                // Open the Word application and document
                wordApp = new Word.Application();
                doc = wordApp.Documents.Open(documentPath);

                // Replace each placeholder with the corresponding value
                foreach (var item in data)
                {
                    foreach (Word.ContentControl contentControl in doc.ContentControls)
                    {
                        if (contentControl.Title == item.Placeholder)
                        {
                            contentControl.Range.Text = item.Value;
                        }
                    }
                }

                // Define the path for the result document and save it
                string resultDocumentPath = Path.Combine(Path.GetDirectoryName(documentPath), "result_" + Path.GetFileName(documentPath));
                if (File.Exists(resultDocumentPath))
                {
                    File.Delete(resultDocumentPath);
                }

                doc.SaveAs2(resultDocumentPath);
                return resultDocumentPath;
            }
            finally
            {
                // Ensure the document and Word application are closed
                if (doc != null)
                {
                    doc.Close(false);
                }
                if (wordApp != null)
                {
                    wordApp.Quit();
                }
            }
        }
    }
}
