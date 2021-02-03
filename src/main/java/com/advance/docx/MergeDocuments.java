package com.advance.docx;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.xmlbeans.XmlOptions;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;

import java.io.*;

public class MergeDocuments {
    public static void main(String[] args) throws Exception {
        InputStream fromJasper = MergeDocuments.class.getClassLoader().getResourceAsStream("FROM_JASPER.docx");
        InputStream fromHtml = MergeDocuments.class.getClassLoader().getResourceAsStream("FROM_HTML.docx");
        OutputStream outputStream = new FileOutputStream("out/result.docx");
        merge(fromJasper, fromHtml, outputStream);
    }

    public static void merge(InputStream src1, InputStream src2, OutputStream dest) throws Exception {

        OPCPackage src1Package = OPCPackage.open(src1);
        OPCPackage src2Package = OPCPackage.open(src2);
        XWPFDocument src1Document = new XWPFDocument(src1Package);
        CTBody src1Body = src1Document.getDocument().getBody();
        XWPFDocument src2Document = new XWPFDocument(src2Package);
        CTBody src2Body = src2Document.getDocument().getBody();
        appendBody(src1Body, src2Body);
        src1Document.write(dest);
    }

    private static void appendBody(CTBody src, CTBody append) throws Exception {
//        XmlOptions optionsOuter = new XmlOptions();
//        optionsOuter.setSaveOuter();
//        String appendString = append.xmlText(optionsOuter);
//        String srcString = src.xmlText();
//        String prefix = srcString.substring(0,srcString.indexOf(">")+1);
//        String mainPart = srcString.substring(srcString.indexOf(">")+1,srcString.lastIndexOf("<"));
//        String sufix = srcString.substring( srcString.lastIndexOf("<") );
//        String addPart = appendString.substring(appendString.indexOf(">") + 1, appendString.lastIndexOf("<"));
//        CTBody makeBody = CTBody.Factory.parse(prefix+mainPart+addPart+sufix);
//        src.set(makeBody);
    }
}
