using ECMWordGenerator.Models;
using System.ServiceModel;

namespace ECMWordGenerator.Contracts
{
    [ServiceContract]
    public interface IWordGeneratorService
    {
        [OperationContract]
        GenerateWordFileResponse GenerateWordFile(RequestData requestData);
    }
}
