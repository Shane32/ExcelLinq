using System;

namespace Shane32.ExcelLinq.Exceptions
{
    public class InvalidDataException : Exception
    {
        public InvalidDataException(string message) : base(message) { }
        public InvalidDataException(string message, Exception innerException) : base(message, innerException) { }
    }
}
