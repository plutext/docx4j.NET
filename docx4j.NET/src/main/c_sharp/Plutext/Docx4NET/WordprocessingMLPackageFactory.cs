using System;
using org.docx4j.openpackaging.packages;
using System.IO;


namespace Plutext.Docx4NET
{
    /// <summary>
    /// Create a docx4j WordprocessingMLPackage object
    /// </summary>
    public class WordprocessingMLPackageFactory
    {
        /// <summary>
        /// Create a docx4j WordprocessingMLPackage from the file at path
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        /// <exception cref="Docx4JException">something went wrong</exception>
        public static WordprocessingMLPackage createWordprocessingMLPackage(string path) 
        {
            return WordprocessingMLPackage
                    .load(new java.io.File(path));
        }


        /// <summary>
        /// Create a docx4j WordprocessingMLPackage from an IO stream
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        /// <exception cref="Docx4JException">something went wrong</exception>
        public static WordprocessingMLPackage createWordprocessingMLPackage(Stream stream)
        {
            return WordprocessingMLPackage
                    .load(new ikvm.io.InputStreamWrapper(stream));
        }


        /// <summary>
        /// Create a docx4j WordprocessingMLPackage from a byte array
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        /// <exception cref="Docx4JException">something went wrong</exception>
        public static WordprocessingMLPackage createWordprocessingMLPackage(byte[] bytes)
        {
            return createWordprocessingMLPackage(new MemoryStream(bytes));
        }


    }
}
