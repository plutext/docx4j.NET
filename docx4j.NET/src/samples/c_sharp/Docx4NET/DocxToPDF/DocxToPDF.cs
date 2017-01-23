using System;
//using System.Collections.Generic;
//using System.Linq;
using System.Text;
using System.IO;

using com.plutext.conversion;

namespace docx4j.NET.samples
{
    /// <summary>
    /// This example doesn't actually use docx4j, since from v3.3, docx4j's
    /// default PDF converter changed from XSL FO + Apache FOP 
    /// to Plutext's commercial PDF Converter. 
    /// 
    /// Per http://www.docx4java.org/forums/announces/docx4j-3-3-0-released-t2381.html
    /// 
    ///   XSL FO based PDF output moved to new/separate project docx4j-export-fo; 
    ///   We made this change after careful consideration, since the quality/performance 
    ///   is so much better, and it removes various dependencies from docx4j itself. 
    /// 
    /// So this example just invokes the converter directly.
    /// 
    /// It would be possible to IKVM docx4j-export-fo, but we haven't done that yet.
    /// </summary>
    class DocxToPDF
    {
        static void Main(string[] args)
        {

            string projectDir = System.IO.Directory.GetParent(
                System.IO.Directory.GetParent(
                Environment.CurrentDirectory.ToString()).ToString()).ToString() + "\\";

            string fileIN = projectDir + @"src\samples\resources\sample-docx.docx";
            string fileOUT = projectDir + @"OUT\sample-docx.pdf";

            Console.WriteLine(fileIN);

            try
            {
                using (Stream file = File.Create(fileOUT))
                {
                    Converter converter = new Converter("http://converter-eval.plutext.com/v1/00000000-0000-0000-0000-000000000000/convert");
                    // you should install your own local instance, and point to that
                    converter.convert(File.ReadAllBytes(fileIN), Format.DOCX, file, Format.PDF);
                }
                Console.WriteLine("done!");
            }
            catch (ConversionException e)
            {
                Console.WriteLine(e);
            }

        }


    }
}
