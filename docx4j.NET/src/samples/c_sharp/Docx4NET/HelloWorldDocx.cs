using System;
using System.Collections.Specialized;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Common.Logging;

using org.docx4j;
using org.docx4j.openpackaging.packages;
using org.docx4j.openpackaging.parts.WordprocessingML;

using org.docx4j.wml;
using javax.xml.bind;

namespace docx4j.NET.samples
{
    /// <summary>
    /// This is an example of using docx4j to create a docx from scratch.
    /// 
    /// For many more samples, see https://github.com/plutext/docx4j/tree/master/src/samples
    /// 
    /// To quickly generate code from a suitable sample docx you've created in Word, 
    /// use the docx4j webapp or the Docx4j Helper Word AddIn:
    /// 
    /// Webapp:  http://webapp.docx4java.org/OnlineDemo/PartsList.html
    /// 
    /// AddIn download:  http://www.plutext.com/dn/downloads/1472868282152/Docx4j_Helper_3-3-1-03.exe
    /// see further http://www.docx4java.org/forums/docx4jhelper-addin-f30/
    /// 
    /// If you are trying this in Visual Studio, it'll be faster if you "start without debugging" (Ctrl+F5)
    /// And first, remember to set this as the "startup object" in project properties.
    /// 
    /// </summary>
    class HelloWorldDocx
    {
        static ILog clog;
        static void Main(string[] args)
        {

            // set up logging
            clog = LoggingConfigurator.configureLogging();
            clog.Info("Hello from Common Logging");

            // create a dir to save the output docx
            string projectDir = System.IO.Directory.GetParent(
                System.IO.Directory.GetParent(
                Environment.CurrentDirectory.ToString()).ToString()).ToString() + "\\";
            string fileOUT = projectDir + @"OUT\HelloWorld.docx";
            System.IO.Directory.CreateDirectory(projectDir + "OUT");


            // docx4j.properties .. add as URL the dir containing docx4j.properties
            Plutext.PropertiesConfigurator.setDocx4jPropertiesDir(projectDir + @"src\samples\resources\");


            // create WordprocessingMLPackage, representing the docx
            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.createPackage();
            MainDocumentPart mdp = wordMLPackage.getMainDocumentPart();


            // add content
            mdp.addParagraphOfText("hello world");  // a convenience method

                // the more generic pattern is:

                org.docx4j.wml.ObjectFactory wmlObjectFactory = org.docx4j.jaxb.Context.getWmlObjectFactory();
                P p = wmlObjectFactory.createP(); // but you can just do = new P();
                mdp.getContent().add(p);
                // Create object for r
                R r = wmlObjectFactory.createR();
                p.getContent().add(r);

                // Create object for t (wrapped in JAXBElement) 
                Text text = wmlObjectFactory.createText();
                JAXBElement textWrapped = wmlObjectFactory.createRT(text); // instead of JAXBElement<org.docx4j.wml.Text>
                // here Text text = new Text() would actually have been fine.

                text.setValue("hello world 2");
                r.getContent().add(textWrapped);


            // save to file
            Docx4J.save(wordMLPackage, new java.io.File(fileOUT), Docx4J.FLAG_SAVE_ZIP_FILE);
        }


    }
}
