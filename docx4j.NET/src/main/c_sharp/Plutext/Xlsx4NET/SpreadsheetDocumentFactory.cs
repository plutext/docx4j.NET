using System;
using System.Collections.Generic;
using org.docx4j.openpackaging.packages;
using System.IO;
using DocumentFormat.OpenXml.Packaging;

namespace Plutext.Xlsx4NET
{
    /// <summary>
    /// Create an OpenXML SDK SpreadsheetDocument object
    /// </summary>
    public class SpreadsheetDocumentFactory
    {
        /// <summary>
        /// Create an OpenXML SDK SpreadsheetDocument object from a
        /// docx4j SpreadsheetMLPackage
        /// </summary>
        /// <param name="xlsxPkg"></param>
        /// <param name="isEditable"></param>
        /// <param name="openSettings"></param>
        /// <returns></returns>
        public static SpreadsheetDocument createPresentationDocument(
            SpreadsheetMLPackage xlsxPkg, 
            bool isEditable, OpenSettings openSettings)
        {
            return SpreadsheetDocument.Open(
                new MemoryStream(SaveFromJavaUtils.toBytes(xlsxPkg)), 
                isEditable, openSettings);
        }

        /// <summary>
        /// Create an OpenXML SDK SpreadsheetDocument object from a
        /// docx4j SpreadsheetMLPackage
        /// </summary>
        /// <param name="xlsxPkg"></param>
        /// <param name="isEditable"></param>
        /// <returns></returns>
        public static SpreadsheetDocument createPresentationDocument(
            SpreadsheetMLPackage xlsxPkg,
            bool isEditable)
        {
            return createPresentationDocument(xlsxPkg, isEditable, new OpenSettings());
        }


    }
}
