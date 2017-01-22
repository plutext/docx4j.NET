using System;
using org.docx4j.openpackaging.packages;
using System.IO;


namespace Plutext.Pptx4NET
{
    /// <summary>
    /// Create a docx4j PresentationMLPackage object
    /// </summary>
    public class PresentationMLPackageFactory
    {
        /// <summary>
        /// Create a PresentationMLPackage from the file at path
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        /// <exception cref="Docx4JException">something went wrong</exception>
        public static PresentationMLPackage createPresentationMLPackage(string path) 
        {
            return PresentationMLPackage
                    .load(new java.io.File(path));
        }


        /// <summary>
        /// Create a PresentationMLPackage from an IO stream
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        /// <exception cref="Docx4JException">something went wrong</exception>
        public static PresentationMLPackage createPresentationMLPackage(Stream stream)
        {
            return PresentationMLPackage
                    .load(new ikvm.io.InputStreamWrapper(stream)) as PresentationMLPackage;
        }


        /// <summary>
        /// Create a PresentationMLPackage from a byte array
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        /// <exception cref="Docx4JException">something went wrong</exception>
        public static PresentationMLPackage createPresentationMLPackage(byte[] bytes)
        {
            return createPresentationMLPackage(new MemoryStream(bytes));
        }

    }
}
