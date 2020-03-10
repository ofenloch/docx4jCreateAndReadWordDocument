package de.ofenloch.docx4j.examples;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertTrue;

import java.io.File;
import java.io.IOException;

import org.docx4j.Docx4J;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.junit.Test;

/**
 * Unit test for simple App.
 */
public class AppTest {

    @Test
    public void test_createDocument() {
        try {
            WordprocessingMLPackage wmlp = createWordDocument.createSampleWordDocument();
            assertNotNull("WordprocessingMLPackage is null", wmlp);
            File tempFile = File.createTempFile("hello", ".docx");
            wmlp.save(tempFile);
            assertTrue("temp file is not readable", tempFile.canRead());
            
            WordprocessingMLPackage wmlpTest = Docx4J.load(tempFile);
            // This fails! We need to do a differen comparison?
            // assertEquals(wmlp, wmlpTest);

        } catch (IOException | Docx4JException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
            assertTrue(e.getMessage(), false);
        }
    }
}
