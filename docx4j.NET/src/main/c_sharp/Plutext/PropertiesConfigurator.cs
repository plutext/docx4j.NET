using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Common.Logging;

namespace Plutext
{
    public class PropertiesConfigurator
    {
        /// <summary>
        /// Specify the dir containing your docx4j properties file
        /// </summary>
        /// <param name="dir"></param>
        public static void setDocx4jPropertiesDir(string dir)
        {
            ILog log = LogManager.GetCurrentClassLogger();

            java.net.URL url = (new java.io.File(dir)).toURL();
            //java.net.URL url = new java.net.URL("file:///C:/Users/jharrop/Documents/Visual%20Studio%202010/Projects/docx4j.NET/docx4j.NET/src/samples/resources/"); // also OK

            java.lang.ClassLoader contextCL = java.lang.Thread.currentThread().getContextClassLoader();
            java.lang.ClassLoader urlCL = java.net.URLClassLoader.newInstance(new java.net.URL[] { url }, contextCL);
            java.lang.Thread.currentThread().setContextClassLoader(urlCL);

            //java.lang.Class clazz = typeof(org.docx4j.utils.ResourceUtils);
            //java.lang.ClassLoader contextCL = clazz.getClassLoader();
            //log.Warn(contextCL.getClass().getName()); // ikvm.runtime.AssemblyClassLoader
            //ikvm.runtime.AssemblyClassLoader acl = contextCL as ikvm.runtime.AssemblyClassLoader;

            // you can delete the below if the properties file is being found
            if (urlCL.getResource("docx4j.properties") == null)
            {
                log.Warn(url.toString() + " dir does not appear to contain docx4j.properties!");
            }
            else
            {
                log.Info(url.toString() + " dir contains docx4j.properties!");
            }

        }

    }
}
