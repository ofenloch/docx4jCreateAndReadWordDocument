package de.ofenloch.docx4j.examples;

import java.io.File;
import java.io.StringWriter;
import java.util.List;

import org.docx4j.wml.Document;
import org.docx4j.TextUtils;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.Parts;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;

/**
 * @Author Oliver Ofenloch
 * 
 */

public class readWordFile {

    static public WordprocessingMLPackage readFromFile(final String fileName) {

        WordprocessingMLPackage wordMLPackage = null;

        try {
            wordMLPackage = WordprocessingMLPackage.load(new File(fileName));
        } catch (Docx4JException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
        return wordMLPackage;
    }

    static String extractText(WordprocessingMLPackage wordMLPackage) {
        StringWriter out = new StringWriter();

        MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();
        Document wmlDocumentEl = (org.docx4j.wml.Document) documentPart.getJaxbElement();
        try {
            TextUtils.extractText(wmlDocumentEl, out);
        } catch (Exception e) {
            out.append("caught exception in TextUtils.extractText");
            e.printStackTrace();
        }
        return out.toString();
    }
}