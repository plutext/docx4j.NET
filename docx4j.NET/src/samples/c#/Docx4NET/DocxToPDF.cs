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
    /// This is an example of using docx4j to produce PDF output.
    /// 
    /// If you are trying this in Visual Studio, it'll be faster if you "start without debugging" (Ctrl+F5)
    /// And first, remember to set this as the "startup object" in project properties.
    /// 
    /// To create XSL FO output, docx4j generates XSL FO, and then
    /// uses a FO renderer to convert the FO to PDF.
    /// 
    /// The default XSL FO renderer is Apache FOP.
    /// 
    /// We're continually improving our PDF output, but if it
    /// isn't good enough for your purposes, you could try:
    /// 
    /// * LibreOffice or OpenOffice (via JODConverter), or
    /// 
    /// * using OpenXML PowerTools to create HTML, then wkhtmltopdf    
    /// 
    /// For issues/feedback, you can post at http://www.docx4java.org/forums/pdf-output-f27/
    /// 
    /// </summary>
    class DocxToPDF
    {
        static void Main(string[] args)
        {

            string projectDir = System.IO.Directory.GetParent(
                System.IO.Directory.GetParent(
                Environment.CurrentDirectory.ToString()).ToString()).ToString() + "\\";

            string fileIN = projectDir + @"src\samples\resources\sample-docx.docx";
            string fileOUT = projectDir + @"OUT_sample-docx.pdf";

            // Programmatically configure Common Logging
            // (alternatively, you could do it declaratively in app.config)
            NameValueCollection commonLoggingproperties = new NameValueCollection();
            commonLoggingproperties["showDateTime"] = "false";
            commonLoggingproperties["level"] = "INFO";
            LogManager.Adapter = new Common.Logging.Simple.ConsoleOutLoggerFactoryAdapter(commonLoggingproperties);


            ILog log = LogManager.GetCurrentClassLogger();
            log.Info("Hello from Common Logging");

            // Necessary, if slf4j-api and slf4j-NetCommonLogging are separate DLLs
            ikvm.runtime.Startup.addBootClassPathAssembly(
                System.Reflection.Assembly.GetAssembly(
                    typeof(org.slf4j.impl.StaticLoggerBinder)));

            // Configure to find docx4j.properties
            // .. add as URL the dir containing docx4j.properties (not the file itself!)
            Plutext.PropertiesConfigurator.setDocx4jPropertiesDir(projectDir + @"src\samples\resources\");

            // OK, do it..
            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage
                    .load(new java.io.File(fileIN));

            java.io.FileOutputStream fos = new java.io.FileOutputStream(new java.io.File(fileOUT));

            org.docx4j.Docx4J.toPDF(wordMLPackage, fos);

            fos.close();

            Console.WriteLine("done!");

        }

    }
}
