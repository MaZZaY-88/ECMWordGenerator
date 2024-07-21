using ECMWordGenerator.Contracts;
using ECMWordGenerator.Logging;
using ECMWordGenerator.Models;
using System.Collections.Generic;
using System.IO;
using Word = Microsoft.Office.Interop.Word;

namespace ECMWordGenerator.Services
{
    public class WordGeneratorService : IWordGeneratorService
    {
        public string GenerateWordFile(RequestData requestData)
        {
            try
            {
                Logger.Log($"User {requestData.UserName} initiated document generation for {requestData.Document}");

                GenerateWordFileInternal(requestData.Document, requestData.Data);

                Logger.Log($"Document generation completed successfully for {requestData.Document}");
                return "Document generated successfully.";
            }
            catch (System.Exception ex)
            {
                Logger.Log($"Error generating document: {ex.Message}", true);
                return $"Error: {ex.Message}";
            }
        }

        private void GenerateWordFileInternal(string documentPath, Dictionary<string, string> data)
        {
            var wordApp = new Word.Application();
            var doc = wordApp.Documents.Open(documentPath);

            foreach (Word.ContentControl contentControl in doc.ContentControls)
            {
                if (data.ContainsKey(contentControl.Title))
                {
                    contentControl.Range.Text = data[contentControl.Title];
                }
            }

            doc.Save();
            doc.Close();
            wordApp.Quit();
        }
    }
}
