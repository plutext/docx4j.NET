using System;
using System.Collections.Generic;
using org.docx4j.openpackaging.packages;
using org.docx4j.openpackaging.io;

namespace Plutext.Docx4NET
{
    public class WordprocessingMLPackageUtils
    {

        public static byte[] toBytes(WordprocessingMLPackage wordPkg)
        {
            SaveToZipFile saver = new SaveToZipFile(wordPkg);
            java.io.ByteArrayOutputStream baos = new java.io.ByteArrayOutputStream();
            saver.save(baos);

            return baos.toByteArray();
        }

        public static void save(WordprocessingMLPackage wordPkg, string path)
        {
            SaveToZipFile saver = new SaveToZipFile(wordPkg);
            saver.save(path);
        }
    }
}
