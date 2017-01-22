using System;
using org.docx4j.openpackaging.packages;
using System.IO;


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


    }
}
