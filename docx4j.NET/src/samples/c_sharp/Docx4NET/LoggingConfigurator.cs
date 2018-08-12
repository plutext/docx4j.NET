using System;
using System.Collections.Specialized;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Common.Logging;


namespace docx4j.NET.samples
{
    class LoggingConfigurator
    {
        public static ILog configureLogging()
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

            return Common.Logging.LogManager.GetCurrentClassLogger();

        }
    }
}
