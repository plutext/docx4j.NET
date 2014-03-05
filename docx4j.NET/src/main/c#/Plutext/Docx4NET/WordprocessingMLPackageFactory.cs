using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
using org.docx4j.openpackaging.packages;
using System.IO;
//using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;


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

        /// <summary>
        /// Create a docx4j WordprocessingMLPackage from an OpenXML SDK WordprocessingDocument
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        /// <exception cref="Docx4JException">something went wrong</exception>
        public static WordprocessingMLPackage createWordprocessingMLPackage(WordprocessingDocument wordDoc)
        {
            return createWordprocessingMLPackage(WordprocessingDocumentToStream(wordDoc));
        }



        private static Stream WordprocessingDocumentToStream(WordprocessingDocument wordDoc)
        {
            MemoryStream mem = new MemoryStream();

            using (var resultDoc = WordprocessingDocument.Create(mem, wordDoc.DocumentType))
            {

                // copy parts from source document to new document
                foreach (var part in wordDoc.Parts)
                {
                    OpenXmlPart targetPart = resultDoc.AddPart(part.OpenXmlPart, part.RelationshipId); // that's recursive :-)
                }

                resultDoc.Package.Flush();
            }
            //resultDoc.Package.Close(); // must do this (or using), or the zip won't get created properly

            mem.Position = 0;

            return mem;

            //    byte[] bytes = new byte[mem.Length];
            //    mem.Read(bytes, 0, (int)mem.Length);
            //    mem.Close();

            //    //FileStream file = new FileStream(@"C:\Users\jharrop\Documents\tmp-test-docx\outs.docx", FileMode.Create, System.IO.FileAccess.Write);
            //    //byte[] bytes = new byte[mem.Length];
            //    //mem.Read(bytes, 0, (int)mem.Length);
            //    //file.Write(bytes, 0, bytes.Length);
            //    //file.Close();
            //    //mem.Close();

            //return bytes;

        }

    }
}
