package org.excel.utils.service.LargeDataRead;

import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

import java.util.ArrayList;
import java.util.List;

public class SharedStringHandler extends DefaultHandler {
    private ArrayList<String> strings = new ArrayList<>();

    private String str;
    @Override
    public void startElement(String uri, String localName, String qName, Attributes attributes) throws SAXException {

    }

    @Override
    public void characters(char[] ch, int start, int length) throws SAXException {
        str = new String(ch,start,length);
    }

    @Override
    public void endElement(String uri, String localName, String qName) throws SAXException {
        if("t".equals(qName)) strings.add(str);
    }

    public ArrayList<String> getStrings() {
        return strings;
    }
}
