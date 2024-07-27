using ECMWordGenerator.Logging;
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
                wordApp.Visible = false; // Ensure the application is not visible
                doc = wordApp.Documents.Open(documentPath);

                // Replace each placeholder with the corresponding value
                foreach (var item in data)
                {
                    foreach (Word.ContentControl contentControl in doc.ContentControls)
                    {
                        if (contentControl.Title == item.Placeholder)
                        {
                            // Clear the current content
                            contentControl.Range.Text = string.Empty;

                            // Log the value to be inserted
                            Logger.Log($"Inserting value for placeholder {item.Placeholder}: {item.Value}");

                            // Split the value by lines and insert each line
                            var lines = item.Value.Split(new[] { '\r', '\n' }, System.StringSplitOptions.RemoveEmptyEntries);
                            Word.Range range = contentControl.Range;
                            object style = range.get_Style();

                            for (int i = 0; i < lines.Length; i++)
                            {
                                if (i > 0)
                                {
                                    range.InsertParagraphAfter();
                                    range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                                }
                                range.InsertAfter(lines[i]);
                                range.set_Style(style);
                                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                            }
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
            catch (System.Exception ex)
            {
                throw new System.Exception($"Error processing the template: {ex.Message}, StackTrace: {ex.StackTrace}", ex);
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
