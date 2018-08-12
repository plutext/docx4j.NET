using System;
using org.docx4j;
using org.docx4j.openpackaging.packages;
using org.docx4j.openpackaging.parts;
using org.docx4j.openpackaging.parts.WordprocessingML;
using org.docx4j.wml;

using System.Collections.Specialized;
using Common.Logging;

namespace docx4j.NET.samples
{
    /// <summary>
    /// Demonstrates using docx4j to resolve the XPaths binding content controls to a CustomXML part.
    /// Ordinarily, Word does this when you open the docx in Word.  But sometimes, you want to do it
    /// yourself, for example, when producing PDF or XHTML output.
    /// 
    /// You might also want docx4j to do it, to take advantage of additional features, such as
    /// repeats, conditional content, or binding escaped XHTML or docx content.  See http://www.opendope.org
    /// 
    /// If you are trying this in Visual Studio, it'll be faster if you "start without debugging" (Ctrl+F5)
    /// And first, remember to set this as the "startup object" in project properties.
    /// 
    /// </summary>
    class ContentControlBind
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

            // the docx 'template'
            String input_DOCX = projectDir + @"src\samples\resources\ContentControlBind\binding-simple.docx";

            // the instance data
            String input_XML = projectDir + @"src\samples\resources\ContentControlBind\binding-simple-data.xml";

            // resulting docx
            String OUTPUT_DOCX = projectDir + @"OUT_ContentControlsMergeXML.docx";

            // Configure to find docx4j.properties
            // .. add as URL the dir containing docx4j.properties (not the file itself!)
            Plutext.PropertiesConfigurator.setDocx4jPropertiesDir(projectDir + @"src\samples\resources\");

		    // Load input_template.docx
		    WordprocessingMLPackage wordMLPackage = org.docx4j.Docx4J.load(new java.io.File(input_DOCX));

		    // Open the xml stream
		    java.io.FileInputStream xmlStream = new java.io.FileInputStream(new java.io.File(input_XML));

		    // Do the binding:
		    // FLAG_NONE means that all the steps of the binding will be done,
		    // otherwise you could pass a combination of the following flags:
		    // FLAG_BIND_INSERT_XML: inject the passed XML into the document
		    // FLAG_BIND_BIND_XML: bind the document and the xml (including any OpenDope handling)
		    // FLAG_BIND_REMOVE_SDT: remove the content controls from the document (only the content remains)
		    // FLAG_BIND_REMOVE_XML: remove the custom xml parts from the document 
			
		    //Docx4J.bind(wordMLPackage, xmlStream, Docx4J.FLAG_NONE);
		    //If a document doesn't include the Opendope definitions, eg. the XPathPart,
		    //then the only thing you can do is insert the xml
		    //the example document binding-simple.docx doesn't have an XPathPart....
		    Docx4J.bind(wordMLPackage, xmlStream, Docx4J.FLAG_BIND_INSERT_XML & Docx4J.FLAG_BIND_BIND_XML);
		
		    //Save the document 
		    Docx4J.save(wordMLPackage, new java.io.File(OUTPUT_DOCX), Docx4J.FLAG_NONE);
		    clog.Info("Saved: " + OUTPUT_DOCX);

        }

    }


}
