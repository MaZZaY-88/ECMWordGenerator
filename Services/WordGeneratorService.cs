using ECMWordGenerator.Contracts;
using ECMWordGenerator.Logging;
using ECMWordGenerator.Models;
using System.Collections.Generic;
using System.ServiceModel;

namespace ECMWordGenerator.Services
{
    [ServiceBehavior(AddressFilterMode = AddressFilterMode.Any)]
    public class WordGeneratorService : IWordGeneratorService
    {
        private readonly WordTemplateProcessor _wordTemplateProcessor;

        /// <summary>
        /// Initializes a new instance of the <see cref="WordGeneratorService"/> class.
        /// </summary>
        public WordGeneratorService()
        {
            _wordTemplateProcessor = new WordTemplateProcessor();
        }

        /// <summary>
        /// Generates a Word file by replacing placeholders with provided values.
        /// </summary>
        /// <param name="requestData">The request data containing user details, authentication token, document path, and data for replacements.</param>
        /// <returns>A response containing the status, filename, and details of the operation.</returns>
        public GenerateWordFileResponse GenerateWordFile(RequestData requestData)
        {
            // Log the start of the method
            Logger.Log("GenerateWordFile method called.");

            // Check if requestData is null
            if (requestData == null)
            {
                Logger.Log("Error: requestData is null.", true);
                return new GenerateWordFileResponse
                {
                    Status = "error",
                    Filename = null,
                    Details = "Error: requestData is null."
                };
            }

            // Check if any required fields in requestData are null or empty
            if (string.IsNullOrEmpty(requestData.UserName) || string.IsNullOrEmpty(requestData.AuthToken) ||
                string.IsNullOrEmpty(requestData.Document) || requestData.Data == null)
            {
                Logger.Log("Error: One or more required fields are null or empty.", true);
                return new GenerateWordFileResponse
                {
                    Status = "error",
                    Filename = null,
                    Details = "Error: One or more required fields are null or empty."
                };
            }

            try
            {
                // Log the initiation of the document generation
                Logger.Log($"User {requestData.UserName} initiated document generation for {requestData.Document}");
                Logger.Log($"Incoming Params: UserName: {requestData.UserName}, AuthToken: {requestData.AuthToken}, Document: {requestData.Document}, Data: {string.Join(", ", requestData.Data)}");

                // Generate the Word file with the provided data
                string resultDocumentPath = _wordTemplateProcessor.ProcessTemplate(requestData.Document, requestData.Data);

                // Log the successful completion of the document generation
                Logger.Log($"Document generation completed successfully for {requestData.Document}");
                return new GenerateWordFileResponse
                {
                    Status = "ok",
                    Filename = resultDocumentPath,
                    Details = "Document generated successfully."
                };
            }
            catch (System.Exception ex)
            {
                // Log any errors that occur during the document generation process
                Logger.Log($"Error generating document: {ex.Message}", true);
                return new GenerateWordFileResponse
                {
                    Status = "error",
                    Filename = null,
                    Details = $"Error: {ex.Message}"
                };
            }
        }
    }
}
