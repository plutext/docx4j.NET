using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using org.docx4j.openpackaging.packages;
using org.docx4j.openpackaging.parts;

using System.Collections.Specialized;
using Common.Logging;

namespace docx4j.NET.samples
{
    /// <summary>
    /// This is an example of using docx4j to produce HTML output.
    /// 
    /// If you are trying this in Visual Studio, it'll be faster if you "start without debugging" (Ctrl+F5)
    /// And first, remember to set this as the "startup object" in project properties.
    /// 
    /// Its an alternative to using OpenXML PowerTools to create HTML;
    /// see http://openxmldeveloper.org/blog/b/openxmldeveloper/archive/2014/01/30/transform-docx-to-html-css-with-high-fidelity-using-powertools-for-open-xml.aspx
    /// 
    /// For issues/feedback, you can post at http://www.docx4java.org/forums/docx-java-f6/
    /// 
    /// </summary>
    class DocxToHTML
    {
        static ILog clog;
        static void Main(string[] args)
        {
            // set up logging
            clog = LoggingConfigurator.configureLogging();
            clog.Info("Hello from Common Logging");

            string projectDir = System.IO.Directory.GetParent(
                System.IO.Directory.GetParent(
                Environment.CurrentDirectory.ToString()).ToString()).ToString() + "\\";

            string fileIN = projectDir + @"src\samples\resources\sample-docx.docx";

            string fileOUT = projectDir + @"OUT\sample-docx2.html";
            string imageDirPath = projectDir + @"OUT\sample-docx2_files";

            System.IO.Directory.CreateDirectory(projectDir + "OUT");
            System.IO.Directory.CreateDirectory(imageDirPath);

            string imageTargetUri = imageDirPath;

            // Configure to find docx4j.properties
            // .. add as URL the dir containing docx4j.properties (not the file itself!)
            Plutext.PropertiesConfigurator.setDocx4jPropertiesDir(projectDir +@"src\samples\resources\");

            // OK, do it..
            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage
                    .load(new java.io.File(fileIN));

            java.io.FileOutputStream fos = new java.io.FileOutputStream(new java.io.File(fileOUT));

            org.docx4j.Docx4J.toHTML(wordMLPackage, imageDirPath, imageTargetUri, fos);

            fos.close();

            Console.WriteLine("done!");

        }

    }
}
