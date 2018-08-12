using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace com.plutext.conversion
{
    public class Converter
    {

        private string URL = "http://127.0.0.1:9016/v1/00000000-0000-0000-0000-000000000000/convert";

        /**
         * Constructor which uses endpoint: "http://127.0.0.1:9016/v1/00000000-0000-0000-0000-000000000000/convert"; 
         * First you'll need to download/install, from https://converter-eval.plutext.com/
         */
        public Converter()
        {
        }

        public Converter(string endpointURL)
        {
            if (endpointURL != null)
            {
                this.URL = endpointURL;
            }

        }


        private string map(Format f)
        {
            if (Format.DOCX.Equals(f))
            {
                return "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
            }
            else if (Format.DOC.Equals(f))
            {
                return "application/msword";
            }

            return null;
        }

        public void convert(byte[] bytesIn, Format fromFormat, Stream output, Format toFormat) 
        {
		
		    if ( Format.DOCX.Equals(fromFormat) ||  Format.DOC.Equals(fromFormat) ) {
			    // OK
		    } else {
			    throw new ConversionException("Conversion from format " + fromFormat + " not supported");
		    }
		
		    if (!Format.PDF.Equals(toFormat)) {
			    throw new ConversionException("Conversion to format " + toFormat + " not supported");			
		    }

            try
            {
                using (Stream responseStream = HttpPost(map(fromFormat), bytesIn))
                {
                    CopyStream(responseStream, output);
                }
            }
            catch (Exception we)
            {
                throw new ConversionException(we.Message, we);
            }
		
	    }

        public System.IO.Stream HttpPost(string contentType, byte[] bytes)
        {
            System.Net.WebRequest req = System.Net.WebRequest.Create(URL);

            req.ContentType = contentType;
            req.Method = "POST";

            //We need to count how many bytes we're sending. 
            req.ContentLength = bytes.Length;

            System.IO.Stream os = req.GetRequestStream();
            os.Write(bytes, 0, bytes.Length); //Push it out there
            os.Close();

            System.Net.WebResponse resp = req.GetResponse();
            if (resp == null) return null;
            
            return resp.GetResponseStream(); 
        }

        private static void CopyStream(Stream input, Stream output)
        {
            byte[] buffer = new byte[8 * 1024];
            int len;
            while ((len = input.Read(buffer, 0, buffer.Length)) > 0)
            {
                output.Write(buffer, 0, len);
            }
        }

    }

    public enum Format {
	
	    DOCX, DOC, PDF

    };
}
