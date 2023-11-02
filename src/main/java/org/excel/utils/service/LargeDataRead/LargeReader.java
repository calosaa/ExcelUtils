package org.excel.utils.service.LargeDataRead;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.StylesTable;
import org.xml.sax.SAXException;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

public class LargeReader {
    private XSSFReader reader;

    private ArrayList<String> sharedStr;

    private StylesTable stylesTable;

    private SAXParserFactory factory = SAXParserFactory.newInstance();
    // 通过释器工厂创建一个sax解释器
    private SAXParser saxParser = null;
    public LargeReader(String filepath) throws ParserConfigurationException, SAXException {
        System.out.println("init...");
        OPCPackage opcPackage = null;
        try {
            opcPackage = OPCPackage.open(filepath);
            this.reader = new XSSFReader(opcPackage);
            setSharedStr();
            stylesTable = this.reader.getStylesTable();
        } catch (IOException | OpenXML4JException e) {
            throw new RuntimeException(e);
        }
    }

    private void setSharedStr(){
        System.out.println("parse shared strings");
        try {
            InputStream is = this.reader.getSharedStringsData();
            saxParser = factory.newSAXParser();
            SharedStringHandler handler = new SharedStringHandler();
            saxParser.parse(is,handler);
            sharedStr = handler.getStrings();
            for (String s : sharedStr) {
                System.out.println("shared str : " +s);
            }
        } catch (IOException | InvalidFormatException | ParserConfigurationException | SAXException e) {
            throw new RuntimeException(e);
        }
    }
    public InputStream selectSheet(String name){
        System.out.println("select sheet name : " + name);
        try {
            XSSFReader.SheetIterator sheetsData = (XSSFReader.SheetIterator) this.reader.getSheetsData();
            while (sheetsData.hasNext()) {
                InputStream is = sheetsData.next();
                if(sheetsData.getSheetName().equals(name)) return is;
            }
            return null;
        } catch (IOException | InvalidFormatException e) {
            throw new RuntimeException(e);
        }
    }

    public List<List<String>> parse(InputStream inputStream,boolean ignoreHead){
        System.out.println("parsing...");
        try {
            saxParser = factory.newSAXParser();
            // 创建事情处理器对象
            TableHandler handler = new TableHandler(sharedStr,stylesTable,ignoreHead);
            // 使用sax解释器解释xml文件
            saxParser.parse(inputStream, handler);
            System.out.println("parse end");
            return handler.getTable();
        } catch (ParserConfigurationException | SAXException | IOException e) {
            throw new RuntimeException(e);
        }
    }

    public <D> List<D> parse2obj(InputStream inputStream, IndividualObjConvertTableHandler.ConvertHandler<D> convertHandler){
        System.out.println("parsing...");
        try {
            saxParser = factory.newSAXParser();
            // 创建事情处理器对象
            IndividualObjConvertTableHandler<D> handler = new IndividualObjConvertTableHandler<>(sharedStr,stylesTable,convertHandler);
            // 使用sax解释器解释xml文件
            saxParser.parse(inputStream, handler);
            System.out.println("parse end");
            return handler.getTable();
        } catch (ParserConfigurationException | SAXException | IOException e) {
            throw new RuntimeException(e);
        }
    }
}
