using System.Collections.Generic;
using System.Runtime.Serialization;

namespace ECMWordGenerator.Models
{
    [DataContract]
    public class RequestData
    {
        [DataMember]
        public string UserName { get; set; }

        [DataMember]
        public string AuthToken { get; set; }

        [DataMember]
        public string Document { get; set; }

        [DataMember]
        public Dictionary<string, string> Data { get; set; }
    }
}
