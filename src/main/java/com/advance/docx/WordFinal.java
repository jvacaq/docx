package com.advance.docx;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlToken;
import org.openxmlformats.schemas.drawingml.x2006.main.CTNonVisualDrawingProps;
import org.openxmlformats.schemas.drawingml.x2006.main.CTPositiveSize2D;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTInline;

import java.io.ByteArrayInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

public class WordFinal {

    public static void main(String[] args) throws IOException, XmlException {
        XWPFDocument srcDoc = new XWPFDocument(WordFinal.class.getClassLoader().getResourceAsStream("FROM_HTML.docx"));

        XWPFDocument destDoc = new XWPFDocument(WordFinal.class.getClassLoader().getResourceAsStream("FROM_JASPER.docx"));

        OutputStream out = new FileOutputStream("out/Destination.docx");

        for (IBodyElement bodyElement : srcDoc.getBodyElements()) {

            BodyElementType elementType = bodyElement.getElementType();

            if (elementType == BodyElementType.PARAGRAPH) {

                XWPFParagraph srcPr = (XWPFParagraph) bodyElement;

                copyStyle(srcDoc, destDoc, srcDoc.getStyles().getStyle(srcPr.getStyleID()));

                boolean hasImage = false;

                XWPFParagraph dstPr = destDoc.createParagraph();

                // Extract image from source docx file and insert into destination docx file.
                for (XWPFRun srcRun : srcPr.getRuns()) {

                    // You need next code when you want to call XWPFParagraph.removeRun().
                    dstPr.createRun();

                    if (srcRun.getEmbeddedPictures().size() > 0)
                        hasImage = true;

                    for (XWPFPicture pic : srcRun.getEmbeddedPictures()) {

                        byte[] img = pic.getPictureData().getData();

                        long cx = pic.getCTPicture().getSpPr().getXfrm().getExt().getCx();
                        long cy = pic.getCTPicture().getSpPr().getXfrm().getExt().getCy();

                        try {
                            // Working addPicture Code below...
                            String blipId = dstPr.getDocument().addPictureData(new ByteArrayInputStream(img),
                                    Document.PICTURE_TYPE_PNG);
                            createPictureCxCy(destDoc, blipId, destDoc.getNextPicNameNumber(Document.PICTURE_TYPE_PNG),
                                    cx, cy);

                        } catch (InvalidFormatException e1) {
                            e1.printStackTrace();
                        }
                    }
                }

                if (hasImage == false) {
                    int pos = destDoc.getParagraphs().size() - 1;
                    destDoc.setParagraph(srcPr, pos);
                }

            } else if (elementType == BodyElementType.TABLE) {

                XWPFTable table = (XWPFTable) bodyElement;

                copyStyle(srcDoc, destDoc, srcDoc.getStyles().getStyle(table.getStyleID()));

                destDoc.createTable();

                int pos = destDoc.getTables().size() - 1;

                destDoc.setTable(pos, table);
            }
        }

        destDoc.write(out);
        out.close();
        System.out.println("Fin");
    }

    // Copy Styles of Table and Paragraph.
    private static void copyStyle(XWPFDocument srcDoc, XWPFDocument destDoc, XWPFStyle style) {
        if (destDoc == null || style == null)
            return;

        if (destDoc.getStyles() == null) {
            destDoc.createStyles();
        }

        List<XWPFStyle> usedStyleList = srcDoc.getStyles().getUsedStyleList(style);
        for (XWPFStyle xwpfStyle : usedStyleList) {
            destDoc.getStyles().addStyle(xwpfStyle);
        }
    }

    public static void createPictureCxCy(XWPFDocument document, String blipId, int id, long cx, long cy) {
        CTInline inline = document.createParagraph().createRun().getCTR().addNewDrawing().addNewInline();

        String picXml = "" +
                "<a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
                "   <a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
                "      <pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
                "         <pic:nvPicPr>" +
                "            <pic:cNvPr id=\"" + id + "\" name=\"Generated\"/>" +
                "            <pic:cNvPicPr/>" +
                "         </pic:nvPicPr>" +
                "         <pic:blipFill>" +
                "            <a:blip r:embed=\"" + blipId + "\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"/>" +
                "            <a:stretch>" +
                "               <a:fillRect/>" +
                "            </a:stretch>" +
                "         </pic:blipFill>" +
                "         <pic:spPr>" +
                "            <a:xfrm>" +
                "               <a:off x=\"0\" y=\"0\"/>" +
                "               <a:ext cx=\"" + cx + "\" cy=\"" + cy + "\"/>" +
                "            </a:xfrm>" +
                "            <a:prstGeom prst=\"rect\">" +
                "               <a:avLst/>" +
                "            </a:prstGeom>" +
                "         </pic:spPr>" +
                "      </pic:pic>" +
                "   </a:graphicData>" +
                "</a:graphic>";

        //CTGraphicalObjectData graphicData = inline.addNewGraphic().addNewGraphicData();
        XmlToken xmlToken = null;
        try {
            xmlToken = XmlToken.Factory.parse(picXml);
        } catch (XmlException xe) {
            xe.printStackTrace();
        }
        inline.set(xmlToken);
        //graphicData.set(xmlToken);

        inline.setDistT(0);
        inline.setDistB(0);
        inline.setDistL(0);
        inline.setDistR(0);

        CTPositiveSize2D extent = inline.addNewExtent();
        extent.setCx(cx);
        extent.setCy(cy);

        CTNonVisualDrawingProps docPr = inline.addNewDocPr();
        docPr.setId(id);
        docPr.setName("Picture " + id);
        docPr.setDescr("Generated");
    }

    public static void createPicture(XWPFDocument document, String blipId, int id, int width, int height) {
        final int EMU = 9525;
        width *= EMU;
        height *= EMU;

        createPictureCxCy(document, blipId, id, width, height);
    }

//    private static void copyLayout(XWPFDocument srcDoc, XWPFDocument destDoc) {
//        CTPageMar pgMar = srcDoc.getDocument().getBody().getSectPr().getPgMar();
//
//        BigInteger bottom = pgMar.getBottom();
//        BigInteger footer = pgMar.getFooter();
//        BigInteger gutter = pgMar.getGutter();
//        BigInteger header = pgMar.getHeader();
//        BigInteger left = pgMar.getLeft();
//        BigInteger right = pgMar.getRight();
//        BigInteger top = pgMar.getTop();
//
////        CTPageMargins addNewPgMar =
//        CTPageMar addNewPgMar = destDoc.getDocument().getBody().addNewSectPr().addNewPgMar();
//
//        addNewPgMar.setBottom(bottom);
//        addNewPgMar.setFooter(footer);
//        addNewPgMar.setGutter(gutter);
//        addNewPgMar.setHeader(header);
//        addNewPgMar.setLeft(left);
//        addNewPgMar.setRight(right);
//        addNewPgMar.setTop(top);
//
//        CTPageSz pgSzSrc = srcDoc.getDocument().getBody().getSectPr().getPgSz();
//
//        BigInteger code = pgSzSrc.getCode();
//        BigInteger h = pgSzSrc.getH();
////        Enum orient = pgSzSrc.getOrient();
//        STPageOrientation.Enum orient = pgSzSrc.getOrient();
//
//        BigInteger w = pgSzSrc.getW();
//
//        CTPageSz addNewPgSz = destDoc.getDocument().getBody().addNewSectPr().addNewPgSz();
//
//        addNewPgSz.setCode(code);
//        addNewPgSz.setH(h);
//        addNewPgSz.setOrient(orient);
//        addNewPgSz.setW(w);
//    }
}