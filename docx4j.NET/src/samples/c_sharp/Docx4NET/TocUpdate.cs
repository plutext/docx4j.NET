using System;
using System.Collections.Specialized;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Common.Logging;

using org.docx4j;
using org.docx4j.toc;
using org.docx4j.openpackaging.packages;
using org.docx4j.openpackaging.parts.WordprocessingML;

using org.docx4j.wml;
using javax.xml.bind;

namespace docx4j.NET.samples
{
    /// <summary>
    /// This is an example of using docx4j.NET to update a ToC.
    /// Updating the page numbers without opening the docx in Word to do that, is a common challenge.
    /// Here's how...
    /// 
    /// For many other samples, see https://github.com/plutext/docx4j/tree/master/src/samples
    /// 
    /// If you are trying this in Visual Studio, it'll be faster if you "start without debugging" (Ctrl+F5)
    /// And first, remember to set this as the "startup object" in project properties.
    /// 
    /// </summary>
    class TocUpdate
    {
        static ILog clog;

        static bool update = true;  // update the ToC?

        public static String TOC_STYLE_MASK = "TOC%s";

        static void Main(string[] args)
        {
            // set up logging
            clog = LoggingConfigurator.configureLogging();
            clog.Info("Hello from Common Logging");

            // create a dir to save the output docx
            string projectDir = System.IO.Directory.GetParent(
                System.IO.Directory.GetParent(
                Environment.CurrentDirectory.ToString()).ToString()).ToString() + "\\";
            System.IO.Directory.CreateDirectory(projectDir + "OUT");

            // docx4j.properties .. add as URL the dir containing docx4j.properties
            Plutext.PropertiesConfigurator.setDocx4jPropertiesDir(projectDir + @"src\samples\resources\");


            // create WordprocessingMLPackage, representing the docx
            // and add some content to it
            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.createPackage();
            MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();
            populateWithContent(documentPart);

            // Now add a ToC
            TocGenerator tocGenerator = new TocGenerator(wordMLPackage);
            // you should install your own local instance, and point to that in docx4j.properties

            tocGenerator.generateToc(0, " TOC \\o \"1-3\" \\h \\z \\u ", false);

            // Save the docx
            string fileOUT = projectDir + @"OUT\TocSample_Generated.docx";
            Docx4J.save(wordMLPackage, new java.io.File(fileOUT), Docx4J.FLAG_SAVE_ZIP_FILE);


            if (update)
            {

                documentPart.addStyledParagraphOfText("Heading2", "Hello 12");
                fillPageWithContent(documentPart, "Hello 12");
                documentPart.addStyledParagraphOfText("Heading1", "Hello 21");
                fillPageWithContent(documentPart, "Hello 21");
                documentPart.addStyledParagraphOfText("Heading2", "Hello 22");
                fillPageWithContent(documentPart, "Hello 22");
                documentPart.addStyledParagraphOfText("Heading3", "Hello 23");
                fillPageWithContent(documentPart, "Hello 23");

                tocGenerator.updateToc(false);

                fileOUT = projectDir + @"OUT\TocSample_Updated.docx";
                Docx4J.save(wordMLPackage, new java.io.File(fileOUT), Docx4J.FLAG_SAVE_ZIP_FILE);
            }

        }

        static void populateWithContent(MainDocumentPart documentPart)
        {
            for (int i = 1; i < 10; i++)
            {
                documentPart.getPropertyResolver().activateStyle(java.lang.String.format(TOC_STYLE_MASK, i));
            }

            documentPart.addStyledParagraphOfText("Heading1", "Hello 1");
            fillPageWithContent(documentPart, "Hello 1");
            documentPart.addStyledParagraphOfText("Heading2", "Hello 2");
            fillPageWithContent(documentPart, "Hello 2");
            documentPart.addStyledParagraphOfText("Title", "Title lvl 3");
            fillPageWithContent(documentPart, "Title test");
            documentPart.addStyledParagraphOfText("Heading3", "Hello 3");
            fillPageWithContent(documentPart, "Hello 3");
            documentPart.addStyledParagraphOfText("Heading1", "Hello 11");
            fillPageWithContent(documentPart, "Hello 11");
            documentPart.addStyledParagraphOfText("Heading1", "Hello 1");
            fillPageWithContent(documentPart, "Hello 1");
            documentPart.addStyledParagraphOfText("Heading2", "Hello 2");
            fillPageWithContent(documentPart, "Hello 2");
            documentPart.addStyledParagraphOfText("Title", "Title lvl 3");
            fillPageWithContent(documentPart, "Title test");
            documentPart.addStyledParagraphOfText("Heading3", "Hello 3");
            fillPageWithContent(documentPart, "Hello 3");
            documentPart.addStyledParagraphOfText("Heading1", "Hello 11");
            fillPageWithContent(documentPart, "Hello 11");

            documentPart.addStyledParagraphOfText("Heading1", "Hello 1");
            fillPageWithContent(documentPart, "Hello 1");
            documentPart.addStyledParagraphOfText("Heading2", "Hello 2");
            fillPageWithContent(documentPart, "Hello 2");
            documentPart.addStyledParagraphOfText("Title", "Title lvl 3");
            fillPageWithContent(documentPart, "Title test");
            documentPart.addStyledParagraphOfText("Heading3", "Hello 3");
            fillPageWithContent(documentPart, "Hello 3");
            documentPart.addStyledParagraphOfText("Heading1", "Hello 11");
            fillPageWithContent(documentPart, "Hello 11");
            documentPart.addStyledParagraphOfText("Heading1", "Hello 1");
            fillPageWithContent(documentPart, "Hello 1");
            documentPart.addStyledParagraphOfText("Heading2", "Hello 2");
            fillPageWithContent(documentPart, "Hello 2");
            documentPart.addStyledParagraphOfText("Title", "Title lvl 3");
            fillPageWithContent(documentPart, "Title test");
            documentPart.addStyledParagraphOfText("Heading3", "Hello 3");
            fillPageWithContent(documentPart, "Hello 3");
            documentPart.addStyledParagraphOfText("Heading1", "Hello 11");
            fillPageWithContent(documentPart, "Hello 11");

        }

        private static void fillPageWithContent(MainDocumentPart documentPart, String content)
        {
            for (int i = 0; i < 10; i++)
            {
                documentPart.addStyledParagraphOfText("Normal", content + " paragraph " + i);
            }
        }

    }
}
