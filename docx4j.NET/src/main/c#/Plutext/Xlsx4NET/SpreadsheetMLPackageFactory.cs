using System;
using org.docx4j.openpackaging.packages;
using System.IO;
using DocumentFormat.OpenXml.Packaging;


namespace Plutext.Xlsx4NET
{
    /// <summary>
    /// Create a docx4j SpreadsheetMLPackage object
    /// </summary>
    public class SpreadsheetMLPackageFactory
    {
        /// <summary>
        /// Create a SpreadsheetMLPackage from the file at path
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        /// <exception cref="Docx4JException">something went wrong</exception>
        public static SpreadsheetMLPackage createSpreadsheetMLPackage(string path) 
        {
            return SpreadsheetMLPackage
                    .load(new java.io.File(path));
        }


        /// <summary>
        /// Create a SpreadsheetMLPackage from an IO stream
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        /// <exception cref="Docx4JException">something went wrong</exception>
        public static SpreadsheetMLPackage createSpreadsheetMLPackage(Stream stream)
        {
            return SpreadsheetMLPackage
                    .load(new ikvm.io.InputStreamWrapper(stream)) as SpreadsheetMLPackage;
        }


        /// <summary>
        /// Create a SpreadsheetMLPackage from a byte array
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        /// <exception cref="Docx4JException">something went wrong</exception>
        public static SpreadsheetMLPackage createSpreadsheetMLPackage(byte[] bytes)
        {
            return createSpreadsheetMLPackage(new MemoryStream(bytes));
        }

        /// <summary>
        /// Create a SpreadsheetMLPackage from an OpenXML SDK SpreadsheetDocument
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        /// <exception cref="Docx4JException">something went wrong</exception>
        public static SpreadsheetMLPackage createSpreadsheetMLPackage(SpreadsheetDocument xlsxDoc)
        {
            return createSpreadsheetMLPackage(SpreadsheetDocumentToStream(xlsxDoc));
        }



        private static Stream SpreadsheetDocumentToStream(SpreadsheetDocument pptxDoc)
        {
            MemoryStream mem = new MemoryStream();

            using (var resultDoc = SpreadsheetDocument.Create(mem, pptxDoc.DocumentType))
            {

                // copy parts from source document to new document
                foreach (var part in pptxDoc.Parts)
                {
                    OpenXmlPart targetPart = resultDoc.AddPart(part.OpenXmlPart, part.RelationshipId); // that's recursive :-)
                }

                resultDoc.Package.Flush();
            }
            //resultDoc.Package.Close(); // must do this (or using), or the zip won't get created properly

            mem.Position = 0;

            return mem;
        }

    }
}
