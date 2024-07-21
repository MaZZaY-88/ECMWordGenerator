using System.Runtime.Serialization;

namespace ECMWordGenerator.Models
{
    /// <summary>
    /// Represents a key-value pair for placeholders in the document.
    /// </summary>
    [DataContract]
    public class Item
    {
        /// <summary>
        /// Gets or sets the placeholder name.
        /// </summary>
        [DataMember]
        public string Placeholder { get; set; }

        /// <summary>
        /// Gets or sets the value to replace the placeholder with.
        /// </summary>
        [DataMember]
        public string Value { get; set; }
    }
}
