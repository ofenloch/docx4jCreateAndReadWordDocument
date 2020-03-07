package de.ofenloch.docx4j.examples;

import java.io.File;

// We don't want the ConsoleAppender.
// import org.apache.log4j.ConsoleAppender;
import org.apache.log4j.FileAppender;
import org.apache.log4j.Level;
import org.apache.log4j.LogManager;
import org.apache.log4j.Logger;
import org.apache.log4j.SimpleLayout;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

/**
 * Hello world!
 *
 */
public class App {

    public static final Logger logger = LogManager.getLogger(App.class.getName());

    public static void main(String[] args) {
        System.out.println("Hello, World!");

        // you can use this method to initialize the logger if
        // you don't have a config file:
        // initLogger();

        String fileName = "HelloWorld.docx";
        try {
            WordprocessingMLPackage wmlp = createWordDocument.createSampleWordDocument();
            wmlp.save(new File(fileName));
        } catch (InvalidFormatException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        } catch (Docx4JException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    }

    /**
     * you can use this method to initialize the logger if you don't have a config
     * file
     */
    public static void initLogger() {
        try {
            SimpleLayout layout = new SimpleLayout();
            // We don't want the ConsoleAppender.
            // ConsoleAppender consoleAppender = new ConsoleAppender(layout);
            // logger.addAppender(consoleAppender);
            FileAppender fileAppender = new FileAppender(layout, "./App.log", false);
            logger.addAppender(fileAppender);
            // ALL | DEBUG | INFO | WARN | ERROR | FATAL | OFF:
            logger.setLevel(Level.DEBUG);
        } catch (Exception ex) {
            System.out.println("initLogger failed:" + ex);
        }
        logger.debug("test DEBUG message");
        logger.info("test INFO message");
        logger.warn("test WARN message");
        logger.error("test ERROR message");
        logger.fatal("test FATA error mesasge");
    }

}
