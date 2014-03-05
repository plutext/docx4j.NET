using System;
using System.Collections.Generic;
using org.docx4j.openpackaging.packages;
using org.docx4j.openpackaging.io;
using System.IO;
//using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;

namespace Plutext.Docx4NET
{
    /// <summary>
    /// Create an OpenXML SDK WordprocessingDocument object
    /// </summary>
    public class WordprocessingDocumentFactory
    {
        /// <summary>
        /// Create an OpenXML SDK WordprocessingDocument object
        /// from a docx4j WordprocessingMLPackage
        /// </summary>
        /// <param name="wordPkg"></param>
        /// <param name="isEditable"></param>
        /// <param name="openSettings"></param>
        /// <returns></returns>
        public static WordprocessingDocument createWordprocessingDocument(
            WordprocessingMLPackage wordPkg, 
            bool isEditable, OpenSettings openSettings)
        {
            return WordprocessingDocument.Open(
                new MemoryStream(SaveFromJavaUtils.toBytes(wordPkg)), 
                isEditable, openSettings);
        }

        /// <summary>
        /// Create an OpenXML SDK WordprocessingDocument object
        /// from a docx4j WordprocessingMLPackage
        /// </summary>
        /// <param name="wordPkg"></param>
        /// <param name="isEditable"></param>
        /// <returns></returns>
        public static WordprocessingDocument createWordprocessingDocument(
            WordprocessingMLPackage wordPkg,
            bool isEditable)
        {
            return createWordprocessingDocument(wordPkg, isEditable, new OpenSettings());
        }


    }
}
