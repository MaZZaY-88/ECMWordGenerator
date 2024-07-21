using System.Runtime.Serialization;

namespace ECMWordGenerator.Models
{
    [DataContract]
    public class GenerateWordFileResponse
    {
        [DataMember]
        public string Status { get; set; }

        [DataMember]
        public string Filename { get; set; }

        [DataMember]
        public string Details { get; set; }
    }
}
