using System;
using System.Collections.Generic;
using org.docx4j.openpackaging.packages;
using System.IO;
using DocumentFormat.OpenXml.Packaging;

namespace Plutext.Pptx4NET
{
    /// <summary>
    /// Create an OpenXML SDK WordprocessingDocument object
    /// </summary>
    public class PresentationDocumentFactory
    {
        /// <summary>
        /// Create an OpenXML SDK PresentationDocument object
        /// from a docx4j PresentationMLPackage
        /// </summary>
        /// <param name="pptxPkg"></param>
        /// <param name="isEditable"></param>
        /// <param name="openSettings"></param>
        /// <returns></returns>
        public static PresentationDocument createPresentationDocument(
            PresentationMLPackage pptxPkg, 
            bool isEditable, OpenSettings openSettings)
        {
            return PresentationDocument.Open(
                new MemoryStream(SaveFromJavaUtils.toBytes(pptxPkg)), 
                isEditable, openSettings);
        }

        /// <summary>
        /// Create an OpenXML SDK PresentationDocument object
        /// from a docx4j PresentationMLPackage
        /// </summary>
        /// <param name="pptxPkg"></param>
        /// <param name="isEditable"></param>
        /// <returns></returns>
        public static PresentationDocument createPresentationDocument(
            PresentationMLPackage pptxPkg,
            bool isEditable)
        {
            return createPresentationDocument(pptxPkg, isEditable, new OpenSettings());
        }


    }
}
