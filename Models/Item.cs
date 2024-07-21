using System.Runtime.Serialization;

namespace ECMWordGenerator.Models
{
    /// <summary>
    /// Represents a key-value pair.
    /// </summary>
    [DataContract]
    public class Item
    {
        /// <summary>
        /// Gets or sets the key.
        /// </summary>
        [DataMember]
        public string Placeholder { get; set; }

        /// <summary>
        /// Gets or sets the value.
        /// </summary>
        [DataMember]
        public string Value { get; set; }
    }
}
