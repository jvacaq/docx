package com.advance.html;

import org.docx4j.convert.in.xhtml.XHTMLImporterImpl;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;

/**
 * Round-trip XHTML to docx and back to XHTML.
 */
public class XhtmlToDocxAndBack {

    private static Logger log = LoggerFactory.getLogger(XhtmlToDocxAndBack.class);


    public static void main(String[] args) throws Exception {
        InputStream resourceAsStream = XhtmlToDocxAndBack.class.getClassLoader().getResourceAsStream("content.html");

        String xhtml = HtmlTools.getXhtml(resourceAsStream);


        // To docx, with content controls
        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.createPackage();

        XHTMLImporterImpl XHTMLImporter = new XHTMLImporterImpl(wordMLPackage);
        //XHTMLImporter.setDivHandler(new DivToSdt());

        wordMLPackage.getMainDocumentPart().getContent().addAll(
                XHTMLImporter.convert(xhtml, null));
        try (OutputStream os = new FileOutputStream("out/from_html.docx")) {
            wordMLPackage.save(os);
        }

    }


}