using System;

namespace Dotnet.Framework.Sample.Model
{
    /// <summary>
    /// Sample model class.
    /// </summary>
    public class SampleModel
    {
        /// <summary>
        /// User id.
        /// </summary>
        public int Prop_ID { get; set; }

        /// <summary>
        /// Username.
        /// </summary>
        public string Prop_Name { get; set; }

        /// <summary>
        /// Created at time.
        /// </summary>
        public DateTime Prop_CreatedAt { get; set; }

        /// <summary>
        /// Updated at time.
        /// </summary>
        public DateTime Prop_UpdatedAt { get; set; }

        /// <summary>
        /// User id.
        /// </summary>
        public int Field_ID;

        /// <summary>
        /// Username.
        /// </summary>
        public string Field_Name;

        /// <summary>
        /// Created at time.
        /// </summary>
        public DateTime Field_CreatedAt;

        /// <summary>
        /// Updated at time.
        /// </summary>
        public DateTime Field_UpdatedAt;
    }
}
