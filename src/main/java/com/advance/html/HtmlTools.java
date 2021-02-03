package com.advance.html;

import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;

import java.io.IOException;
import java.io.InputStream;

public class HtmlTools {
    public static String getXhtml(String html) {
        Document document = Jsoup.parse(html);
        document.outputSettings().syntax(Document.OutputSettings.Syntax.xml);
        document.outputSettings().escapeMode(org.jsoup.nodes.Entities.EscapeMode.xhtml);
        return document.html();
    }

    public static String getXhtml(byte[] htmlBytes){
        return getXhtml(new String(htmlBytes));
    }

    public static String getXhtml(InputStream htmlStream) throws IOException {
        return getXhtml(htmlStream.readAllBytes());
    }}
