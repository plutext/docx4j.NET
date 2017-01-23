using System;
using System.Collections.Specialized;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Common.Logging;

using org.docx4j;
using org.docx4j.openpackaging.packages;
using org.docx4j.openpackaging.parts;

using org.docx4j.anon;

namespace docx4j.NET.samples
{
    /// <summary>
    /// This sample uses docx4j.NET to remove sensitive info from the docx.  
    /// 
    /// Text is converted to lorem ipsum strings, etc.
    /// 
    /// This is especially useful if you want to post the docx for technical support.
    /// 
    /// If you are trying this in Visual Studio, it'll be faster if you "start without debugging" (Ctrl+F5)
    /// And first, remember to set this as the "startup object" in project properties.
    /// 
    /// </summary>
    class Desensitize
    {
        static ILog clog;
        static void Main(string[] args)
        {

            // set up logging
            configureLogging();

            // create a dir to save the output docx
            string projectDir = System.IO.Directory.GetParent(
                System.IO.Directory.GetParent(
                Environment.CurrentDirectory.ToString()).ToString()).ToString() + "\\";

            System.IO.Directory.CreateDirectory(projectDir + "OUT");

            string fileIN = projectDir + @"src\samples\resources\sample-docx.docx";
            string fileOUT = projectDir + @"OUT\Anon.docx";

		    WordprocessingMLPackage pkg = Docx4J.load(new java.io.File(fileIN));
            // or 
            // WordprocessingMLPackage pkg = Plutext.Docx4NET.WordprocessingMLPackageFactory.createWordprocessingMLPackage(fileIN);


            // Anonymize/densensitize it
		    Anonymize anon = new Anonymize(pkg);
		    AnonymizeResult result = anon.go();
		

		    // Report
		    clog.Info("\n\n REPORT for " + fileIN + "\n\n");
		    if (result.isOK()) {
			    clog.Info("document successfully anonymised.");
		    } else {
			    clog.Info("document partially anonymised; please check " + fileOUT);
			
			    if (result.getUnsafeParts().size()>0) {
				    clog.Info("The following parts may leak info:");
				    foreach(Part p in result.getUnsafeParts()) {
					    clog.Info(p.getPartName().getName() + ", of type " + p.getClass().getName() );
				    }
			    }
			
			    // unsafe objects
			    clog.Info(result.reportUnsafeObjects());
		    }
		
		    clog.Info("\n\n .. end REPORT for " + fileIN  + "\n\n");
		

            // save result to file
            Docx4J.save(pkg, new java.io.File(fileOUT));
            clog.Info("\n\n saved " + fileOUT + "\n\n");
        }

        static void configureLogging()
        {
            ikvm.runtime.Startup.addBootClassPathAssembly(
                System.Reflection.Assembly.GetAssembly(
                    typeof(org.slf4j.impl.StaticLoggerBinder)));

            ikvm.runtime.Startup.addBootClassPathAssembly(
                System.Reflection.Assembly.GetAssembly(
                    typeof(org.slf4j.LoggerFactory)));

            NameValueCollection commonLoggingproperties = new NameValueCollection();
            commonLoggingproperties["showDateTime"] = "false";
            commonLoggingproperties["level"] = "INFO";

            Common.Logging.LogManager.Adapter = new Common.Logging.Simple.ConsoleOutLoggerFactoryAdapter(commonLoggingproperties);
            // In VS 2010 for output type console application, that shows up in a new console window,
            // whether you start with debugging, or without.

            //Common.Logging.LogManager.Adapter = new Common.Logging.Simple.TraceLoggerFactoryAdapter(commonLoggingproperties);
            // In VS 2010 for output type console application, that shows up in a "show output from debugging",
            // provided you are debugging!

            // In a real application, you might route Common.Logging to NLog
            // Common.Logging.LogManager.Adapter = new Common.Logging.NLog.NLogLoggerFactoryAdapter(commonLoggingproperties);

            clog = Common.Logging.LogManager.GetCurrentClassLogger();
            clog.Info("Hello from Common Logging");

        }

    }
}
