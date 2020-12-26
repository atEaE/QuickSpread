using System;

namespace Dotnet.Core.Sample.Model
{
    /// <summary>
    /// Sample model class.
    /// </summary>
    public class SampleModel
    {
        /// <summary>
        /// User id.
        /// </summary>
        public int ID { get; set; }

        /// <summary>
        /// Username.
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Created at time.
        /// </summary>
        public DateTime CreatedAt { get; set; }

        /// <summary>
        /// Updated at time.
        /// </summary>
        public DateTime UpdatedAt { get; set; }
    }
}
