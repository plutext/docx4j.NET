using System;
using org.docx4j.openpackaging.packages;
using org.docx4j.openpackaging.parts;
using org.docx4j.openpackaging.parts.WordprocessingML;
using org.docx4j.wml;

using DocumentFormat.OpenXml.Packaging;

using System.Collections.Specialized;
using Common.Logging;

namespace docx4j.NET.samples
{
    /// <summary>
    /// Demonstrates converting docx4j representation of a docx
    /// to Open XML SDK representation, and vice versa.
    /// 
    /// This is a simple way to take advantage of features
    /// unique to each library.
    /// 
    /// If you are trying this in Visual Studio, it'll be faster if you "start without debugging" (Ctrl+F5)
    /// And first, remember to set this as the "startup object" in project properties.
    /// 
    /// </summary>
    class InteropDocx
    {
        static void Main(string[] args)
        {
            string projectDir = System.IO.Directory.GetParent(
                System.IO.Directory.GetParent(
                Environment.CurrentDirectory.ToString()).ToString()).ToString() + "\\";

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


            // Create a docx4j docx
            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.createPackage();
            org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();
            documentPart.addParagraphOfText("Hello world");

            // .. or alternatively, load existing
            //string template = @"C:\Users\jharrop\Documents\tmp-test-docx\HelloWorld.docx";
            //WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage
            //        .load(new java.io.File(template));

            Console.WriteLine(documentPart.getJaxbElement().GetType().FullName); // could use logging here and below, but this is better for demo purposes
            Console.WriteLine(documentPart.getXML());

            // Convert to Open XML SDK object
            WordprocessingDocument openXmlSdkObj = Plutext.Docx4NET.WordprocessingDocumentFactory.createWordprocessingDocument(wordMLPackage, true);
            Console.WriteLine(openXmlSdkObj.GetType().FullName);

            // Convert back to docx4j object
            WordprocessingMLPackage wordMLPackage2 = Plutext.Docx4NET.WordprocessingMLPackageFactory.createWordprocessingMLPackage(openXmlSdkObj);
            Console.WriteLine(wordMLPackage2.GetType().FullName);
            Console.WriteLine(wordMLPackage2.getMainDocumentPart().getXML());

        }


    }
}