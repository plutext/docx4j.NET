using System;
using org.docx4j.openpackaging.packages;
using org.docx4j.openpackaging.parts;
using org.docx4j.openpackaging.parts.WordprocessingML;
using org.docx4j.wml;

using org.docx4j.model.fields.merge;
using Plutext;

using System.Collections.Specialized;
using Common.Logging;

namespace docx4j.NET.samples
{
    /// <summary>
    /// This is an example of using docx4j to perform a mail merge (injecting data into MERGEFIELD) in .NET.
    /// 
    /// If you are trying this in Visual Studio, it'll be faster if you "start without debugging" (Ctrl+F5)
    /// And first, remember to set this as the "startup object" in project properties.
    /// 
    /// For issues/feedback, you can post at http://www.docx4java.org/forums/docx-java-f6/
    /// </summary>
    class MailMergeField
    {
        static org.docx4j.wml.ObjectFactory wmlObjectFactory = new org.docx4j.wml.ObjectFactory();

        static ILog clog;
        static void Main(string[] args)
        {
            // set up logging
            clog = LoggingConfigurator.configureLogging();
            clog.Info("Hello from Common Logging");

            bool mergedOutput = true;

            string projectDir = System.IO.Directory.GetParent(
                System.IO.Directory.GetParent(
                Environment.CurrentDirectory.ToString()).ToString()).ToString() + "\\";

            string saveToPathPrefix = projectDir + @"OUT_MailMergeField";

            // Configure to find docx4j.properties
            // .. add as URL the dir containing docx4j.properties (not the file itself!)
            Plutext.PropertiesConfigurator.setDocx4jPropertiesDir(projectDir + @"src\samples\resources\");


            // Create a docx4j docx
            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.createPackage();
            org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();
            documentPart.addObject(addParagraphWithMergeField("Hallo, MERGEFORMAT: ", " MERGEFIELD  kundenname  \\* MERGEFORMAT ", "«Kundenname»"));
            documentPart.addObject(addParagraphWithMergeField("Hallo, lower: ", " MERGEFIELD  kundenname  \\* Lower ", "«Kundenname»"));
            documentPart.addObject(addParagraphWithMergeField("Hallo, firstcap: ", " MERGEFIELD  kundenname  \\* FirstCap MERGEFORMAT ", "«Kundenname»"));
            documentPart.addObject(addParagraphWithMergeField("Hallo, random case: ", " MERGEFIELD  KunDenName  \\* MERGEFORMAT ", "«Kundenname»"));
            documentPart.addObject(addParagraphWithMergeField("Hallo, all caps: ", " MERGEFIELD  KUNDENNAME  \\* Upper MERGEFORMAT ", "«Kundenname»"));
            // " MERGEFIELD  yourdate \@ &quot;dddd, MMMM dd, yyyy&quot; "
            //documentPart.addObject(addParagraphWithMergeField("Date example", " MERGEFIELD  yourdate \\@ 'dddd, MMMM dd, yyyy' ", "«Kundenname»")); // FIXME .. doesn't work via .NET.  Why?
            documentPart.addObject(addParagraphWithMergeField("Number example: ", " MERGEFIELD  yournumber \\# $#,###,###  ", "«Kundenname»"));

            // .. or alternatively, load existing
            //string template = @"C:\Users\jharrop\Documents\tmp-test-docx\HelloWorld.docx";
            //WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage
            //        .load(new java.io.File(template));

            //Console.WriteLine(documentPart.getXML());

		java.util.List data = new java.util.ArrayList();
            // TODO: make more .NET friendly

		// Instance 1
		java.util.Map map = new java.util.HashMap();
		map.put( new DataFieldName("KundenNAme"), "Daffy duck");
		map.put( new DataFieldName("Kundenname"), "Plutext");
		map.put(new DataFieldName("Kundenstrasse"), "Bourke Street");
		// To get dates right, make sure you have docx4j property docx4j.Fields.Dates.DateFormatInferencer.USA
		// set to true or false as appropriate.  It defaults to non-US.
		map.put(new DataFieldName("yourdate"), "15/4/2013");  
		map.put(new DataFieldName("yournumber"), "2456800");
		data.add(map);
				
		// Instance 2
		map = new java.util.HashMap();
		map.put( new DataFieldName("Kundenname"), "Jason");
		map.put(new DataFieldName("Kundenstrasse"), "Collins Street");
		map.put(new DataFieldName("yourdate"), "12/10/2013");
		map.put(new DataFieldName("yournumber"), "1234800");
		data.add(map);		
		
		
		if (mergedOutput) {
			/*
			 * This is a "poor man's" merge, which generates the mail merge  
			 * results as a single docx, and just hopes for the best.
			 * Images and hyperlinks should be ok. But numbering 
			 * will continue, as will footnotes/endnotes.
			 *  
			 * If your resulting documents aren't opening in Word, then
			 * you probably need MergeDocx to perform the merge.
			 */

			// How to treat the MERGEFIELD, in the output?
			org.docx4j.model.fields.merge.MailMerger.setMERGEFIELDInOutput(org.docx4j.model.fields.merge.MailMerger.OutputField.KEEP_MERGEFIELD);
			
//			log.Debug(XmlUtils.marshaltoString(wordMLPackage.getMainDocumentPart().getJaxbElement(), true, true));
			
			WordprocessingMLPackage output = org.docx4j.model.fields.merge.MailMerger.getConsolidatedResultCrude(wordMLPackage, data, true);
			
//			log.Info(XmlUtils.marshaltoString(output.getMainDocumentPart().getJaxbElement(), true, true));
			
            SaveFromJavaUtils.save(output, saveToPathPrefix + ".docx");
			
		} else {
			// Need to keep thane MERGEFIELDs. If you don't, you'd have to clone the docx, and perform the
			// merge on the clone.  For how to clone, see the MailMerger code, method getConsolidatedResultCrude
			org.docx4j.model.fields.merge.MailMerger.setMERGEFIELDInOutput(org.docx4j.model.fields.merge.MailMerger.OutputField.KEEP_MERGEFIELD);


            for (int i = 0; i < data.size(); i++ )
            {
                java.util.Map thismap = (java.util.Map)data.get(i);
                org.docx4j.model.fields.merge.MailMerger.performMerge(wordMLPackage, thismap, true);
                SaveFromJavaUtils.save(wordMLPackage, saveToPathPrefix +  "_" + i + ".docx");
            }			
		}
        clog.Info("Done! Saved to " + saveToPathPrefix);

        }

        static org.docx4j.wml.P addParagraphWithMergeField(string leadingText, string instr, string placeholder)
        {
            // Create object for p
            P p = wmlObjectFactory.createP();

            // Create object for r
            R r = wmlObjectFactory.createR();
            p.getContent().add(r);
            // Create object for t (wrapped in JAXBElement) 
            Text text = wmlObjectFactory.createText();
            javax.xml.bind.JAXBElement textWrapped = wmlObjectFactory.createRT(text);
            r.getContent().add(textWrapped);
            text.setValue(leadingText);
            text.setSpace("preserve");

            // Create object for fldSimple (wrapped in JAXBElement) 
            CTSimpleField simplefield = wmlObjectFactory.createCTSimpleField();
            javax.xml.bind.JAXBElement simplefieldWrapped = wmlObjectFactory.createPFldSimple(simplefield);
            p.getContent().add(simplefieldWrapped);
            simplefield.setInstr(instr);
            // Create object for r
            R r2 = wmlObjectFactory.createR();
            simplefield.getContent().add(r2);
            // Create object for t (wrapped in JAXBElement) 
            Text text2 = wmlObjectFactory.createText();
            javax.xml.bind.JAXBElement textWrapped2 = wmlObjectFactory.createRT(text2);
            r2.getContent().add(textWrapped2);
            text2.setValue(placeholder);

            return p;
        }
    }
}
