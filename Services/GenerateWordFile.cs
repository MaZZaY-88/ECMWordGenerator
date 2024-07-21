using ECMWordGenerator.Contracts;
using ECMWordGenerator.Logging;
using ECMWordGenerator.Models;
using System.Collections.Generic;
using System.IO;
using System.ServiceModel;
using Word = Microsoft.Office.Interop.Word;

namespace ECMWordGenerator.Services
{
    [ServiceBehavior(AddressFilterMode = AddressFilterMode.Any)]
    public class WordGeneratorService : IWordGeneratorService
    {
        /// <summary>
        /// Generates a Word file by replacing placeholders with provided values.
        /// </summary>
        /// <param name="requestData">The request data containing user details, authentication token, document path, and data for replacements.</param>
        /// <returns>A string message indicating the result of the operation.</returns>
        public string GenerateWordFile(RequestData requestData)
        {
            // Check if requestData is null
            if (requestData == null)
            {
                Logger.Log("Error: requestData is null.", true);
                return "Error: requestData is null.";
            }

            // Check if any required fields in requestData are null or empty
            if (string.IsNullOrEmpty(requestData.UserName) || string.IsNullOrEmpty(requestData.AuthToken) ||
                string.IsNullOrEmpty(requestData.Document) || requestData.Data == null)
            {
                Logger.Log("Error: One or more required fields are null or empty.", true);
                return "Error: One or more required fields are null or empty.";
            }

            try
            {
                // Log the initiation of the document generation
                Logger.Log($"User {requestData.UserName} initiated document generation for {requestData.Document}");
                Logger.Log($"Incoming Params: UserName: {requestData.UserName}, AuthToken: {requestData.AuthToken}, Document: {requestData.Document}, Data: {string.Join(", ", requestData.Data)}");

                // Generate the Word file with the provided data
                string resultDocumentPath = GenerateWordFileInternal(requestData.Document, requestData.Data);

                // Log the successful completion of the document generation
                Logger.Log($"Document generation completed successfully for {requestData.Document}");
                return $"Document generated successfully. Result file: {resultDocumentPath}";
            }
            catch (System.Exception ex)
            {
                // Log any errors that occur during the document generation process
                Logger.Log($"Error generating document: {ex.Message}", true);
                return $"Error: {ex.Message}";
            }
        }

        /// <summary>
        /// Replaces placeholders in the Word document with specified values and saves the result.
        /// </summary>
        /// <param name="documentPath">The path to the original Word document.</param>
        /// <param name="data">A dictionary containing the placeholders and their replacement values.</param>
        /// <returns>The path to the generated Word document.</returns>
        private string GenerateWordFileInternal(string documentPath, Dictionary<string, string> data)
        {
            // Open the Word application and document
            var wordApp = new Word.Application();
            var doc = wordApp.Documents.Open(documentPath);

            // Replace each placeholder with the corresponding value
            foreach (Word.ContentControl contentControl in doc.ContentControls)
            {
                if (data.ContainsKey(contentControl.Title))
                {
                    contentControl.Range.Text = data[contentControl.Title];
                }
            }

            // Define the path for the result document and save it
            string resultDocumentPath = Path.Combine(Path.GetDirectoryName(documentPath), "result_" + Path.GetFileName(documentPath));
            if (File.Exists(resultDocumentPath))
            {
                File.Delete(resultDocumentPath);
            }

            doc.SaveAs2(resultDocumentPath);
            doc.Close();
            wordApp.Quit();

            return resultDocumentPath;
        }
    }
}
