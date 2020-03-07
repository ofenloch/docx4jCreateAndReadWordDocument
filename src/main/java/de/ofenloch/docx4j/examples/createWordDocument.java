package de.ofenloch.docx4j.examples;

/**
 * @Author Oliver Ofenloch
 * 
 * This is taken from org.docx4j.samples.SampleDocument
 * in the docx4j distribution
 */
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Map.Entry;

import javax.xml.bind.JAXBException;

import org.docx4j.XmlUtils;
import org.docx4j.fonts.PhysicalFont;
import org.docx4j.fonts.PhysicalFonts;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;

/**
 * class for generating a simple Word document
 */
public class createWordDocument {

    public static WordprocessingMLPackage createSampleWordDocument() throws InvalidFormatException {
        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.createPackage();
        MainDocumentPart mdp = wordMLPackage.getMainDocumentPart();
        mdp.addParagraphOfText("Hello, World!");
        createContent(mdp);
        return wordMLPackage;
    }

    private static void createContent(MainDocumentPart wordDocumentPart) {
        /*
         * NB, this currently works nicely with viaIText, and viaXSLFO (provided you
         * view with Acrobat Reader .. it seems to overwhelm pdfviewer, which is weird,
         * since viaIText works in both).
         */

        try {
            wordDocumentPart.addParagraphOfText("List of System Fonts");
            // Do this explicitly, since we need
            // it in order to create our content
            PhysicalFonts.discoverPhysicalFonts();

            Map<String, PhysicalFont> physicalFontMap = PhysicalFonts.getPhysicalFonts();
            Iterator<Entry<String, PhysicalFont>> physicalFontMapIterator = physicalFontMap.entrySet().iterator();
            while (physicalFontMapIterator.hasNext()) {
                Map.Entry<String, PhysicalFont> pairs = (Map.Entry<String, PhysicalFont>) physicalFontMapIterator
                        .next();
                if (pairs.getKey() == null) {
                    pairs = (Map.Entry<String, PhysicalFont>) physicalFontMapIterator.next();
                }
                String fontName = (String) pairs.getKey();
                PhysicalFont pf = (PhysicalFont) pairs.getValue();

                System.out.print("Adding paragraph for font \"" + fontName + "\" ...");
                addObject(wordDocumentPart, sampleText, fontName);
                System.out.println("  ... done.");

                // bold, italic etc
                PhysicalFont pfVariation = PhysicalFonts.getBoldForm(pf);
                if (pfVariation != null) {
                    addObject(wordDocumentPart, sampleTextBold, pfVariation.getName());
                }
                pfVariation = PhysicalFonts.getBoldItalicForm(pf);
                if (pfVariation != null) {
                    addObject(wordDocumentPart, sampleTextBoldItalic, pfVariation.getName());
                }
                pfVariation = PhysicalFonts.getItalicForm(pf);
                if (pfVariation != null) {
                    addObject(wordDocumentPart, sampleTextItalic, pfVariation.getName());
                }

            }
        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    private static void addObject(MainDocumentPart wordDocumentPart, String template, String fontName)
            throws JAXBException {

        HashMap<String, String> substitution = new HashMap<String, String>();
        substitution.put("fontname", fontName);
        Object o = XmlUtils.unmarshallFromTemplate(template, substitution);
        wordDocumentPart.addObject(o);

    }

    private final static String sampleText = "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">"
            + "<w:r>" + "<w:rPr>"
            + "<w:rFonts w:ascii=\"${fontname}\" w:eastAsia=\"${fontname}\" w:hAnsi=\"${fontname}\" w:cs=\"${fontname}\" />"
            + "</w:rPr>" + "<w:t xml:space=\"preserve\">${fontname}</w:t>" + "</w:r>" + "</w:p>";
    private final static String sampleTextBold = "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">"
            + "<w:r>" + "<w:rPr>"
            + "<w:rFonts w:ascii=\"${fontname}\" w:eastAsia=\"${fontname}\" w:hAnsi=\"${fontname}\" w:cs=\"${fontname}\" />"
            + "<w:b />" + "</w:rPr>" + "<w:t>${fontname} bold;</w:t>" + "</w:r>" + "</w:p>";
    private final static String sampleTextItalic = "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">"
            + "<w:r>" + "<w:rPr>"
            + "<w:rFonts w:ascii=\"${fontname}\" w:eastAsia=\"${fontname}\" w:hAnsi=\"${fontname}\" w:cs=\"${fontname}\" />"
            + "<w:i />" + "</w:rPr>" + "<w:t>${fontname} italic; </w:t>" + "</w:r>" + "</w:p>";
    private final static String sampleTextBoldItalic = "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">"
            + "<w:r>" + "<w:rPr>"
            + "<w:rFonts w:ascii=\"${fontname}\" w:eastAsia=\"${fontname}\" w:hAnsi=\"${fontname}\" w:cs=\"${fontname}\" />"
            + "<w:b />" + "<w:i />" + "</w:rPr>" + "<w:t>${fontname} bold italic</w:t>" + "</w:r>" + "</w:p>";
} // public class createWordDocument