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
                // Authentication logic here using requestData.AuthToken
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

            foreach (var item in data)
            {
                Word.Find findObject = wordApp.Selection.Find;
                findObject.ClearFormatting();
                findObject.Text = item.Key;
                findObject.Replacement.ClearFormatting();
                findObject.Replacement.Text = item.Value;
                findObject.Execute(Replace: Word.WdReplace.wdReplaceAll);
            }

            doc.Save();
            doc.Close();
            wordApp.Quit();
        }
    }
}
