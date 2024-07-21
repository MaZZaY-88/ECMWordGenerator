using System.Collections.Generic;
using System.Runtime.Serialization;

namespace ECMWordGenerator.Models
{
    /// <summary>
    /// Represents the data required for generating a Word file.
    /// </summary>
    [DataContract]
    public class RequestData
    {
        /// <summary>
        /// Gets or sets the user name of the requestor.
        /// </summary>
        [DataMember]
        public string UserName { get; set; }

        /// <summary>
        /// Gets or sets the authentication token.
        /// </summary>
        [DataMember]
        public string AuthToken { get; set; }

        /// <summary>
        /// Gets or sets the local path to the Word document template.
        /// </summary>
        [DataMember]
        public string Document { get; set; }

        /// <summary>
        /// Gets or sets the list of placeholder-value pairs for replacements in the document.
        /// </summary>
        [DataMember]
        public List<Item> Data { get; set; }
    }
}
