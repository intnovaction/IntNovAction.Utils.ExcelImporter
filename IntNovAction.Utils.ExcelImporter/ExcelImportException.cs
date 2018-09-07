using System;
using System.Runtime.Serialization;

namespace IntNovAction.Utils.Importer
{
    [Serializable]
    internal class ExcelImportException : Exception
    {
        public ExcelImportException()
        {
        }

        public ExcelImportException(string message) : base(message)
        {
        }

        public ExcelImportException(string message, Exception innerException) : base(message, innerException)
        {
        }

        protected ExcelImportException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
        }
    }
}