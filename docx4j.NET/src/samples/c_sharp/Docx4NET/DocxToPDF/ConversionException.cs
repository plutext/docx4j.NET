using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace com.plutext.conversion
{
    /// 

    /// Wraps any exception returned by the conversion process,
    /// for example, System.Net.WebException representing an HTTP error code. 
    /// 

    public class ConversionException : Exception
    {
        public ConversionException()
        {
        }

        public ConversionException(string message)
            : base(message)
        {
        }

        public ConversionException(string message, Exception inner)
            : base(message, inner)
        {
        }
    }
}