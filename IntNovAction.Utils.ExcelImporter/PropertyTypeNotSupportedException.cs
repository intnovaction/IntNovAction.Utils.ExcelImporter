using System;
using System.Runtime.Serialization;

namespace IntNovAction.Utils.ExcelImporter
{
    [Serializable]
    internal class PropertyTypeNotSupportedException : Exception
    {
        public PropertyTypeNotSupportedException()
        {
        }

        public PropertyTypeNotSupportedException(string message) : base(message)
        {
        }

        public PropertyTypeNotSupportedException(string message, Exception innerException) : base(message, innerException)
        {
        }

        protected PropertyTypeNotSupportedException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
        }
    }
}