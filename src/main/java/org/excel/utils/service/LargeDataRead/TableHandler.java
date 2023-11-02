package org.excel.utils.service.LargeDataRead;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

import java.util.ArrayList;
import java.util.List;

public class TableHandler extends DefaultHandler {

    private ArrayList<String> sharedStr;
    private List<List<String>> table;

    private StringBuilder sb;

    private ArrayList<String> currRow;

    private String type = "";

    private StylesTable stylesTable;

    private DataFormatter formatter = new DataFormatter();
    private int formatIndex;
    private String formatString;
    private boolean ignoreHead = true;
    private boolean firstRow = true;
    //private String dateFormat = "m/d/yy,yyyy\\-mm\\-dd;@,yyyy/m/d;@,yyyy/m/d\\ h:mm;@,mm/dd/yy;@,m/d;@,"
    //        + "yy/m/d;@,m/d/yy;@,[$-409]yyyy/m/d\\ h:mm\\ AM/PM;@,[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy";
    private String dateFormat = "m/d/yy,yyyy\\-mm\\-dd;@,yyyy/m/d;@,yyyy/m/d\\ h:mm;@,mm/dd/yy;@,m/d;@,"
            + "yy/m/d;@,m/d/yy;@,[$-409]yyyy/m/d\\ h:mm\\ AM/PM;@,[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy";
    public TableHandler(ArrayList<String> sharedStr,StylesTable stylesTable,boolean ignoreHead){
        this.sharedStr = sharedStr;
        this.stylesTable = stylesTable;
        this.ignoreHead = ignoreHead;
    }
    @Override
    public void startElement(String uri, String localName, String qName, Attributes attributes) throws SAXException {
        if (qName.equals("sheetData")){
            this.table = new ArrayList<>();
        }else if(qName.equals("row")){
            currRow = new ArrayList<>();
        }else if(qName.equals("c")){
            String t  = attributes.getValue("t");
            if(t == null) type = "";
            else if (t.equals("s")){
                type = "s";
            }else if (t.equals("b")){
                type = "b";
            }else if (t.equals("inlineStr")){
                type = "is";
            }
            String sstr  = attributes.getValue("s");
            if (sstr != null){
                int s = Integer.parseInt(sstr);
                XSSFCellStyle style = stylesTable.getStyleAt(s);
                formatIndex = style.getDataFormat();
                formatString = style.getDataFormatString();
                if (dateFormat.contains(formatString)){
                    formatString = "yyyy-MM-dd";
                }
                type = "k";
            }
        }

        sb = new StringBuilder();

    }

    @Override
    public void characters(char[] ch, int start, int length) throws SAXException {
        String text = new String(ch, start, length);
        if (sb != null) {
            sb.append(text);
        }
    }

    @Override
    public void endElement(String uri, String localName, String qName) throws SAXException {
        if (qName.equals("v")){
            if (type.equals("s")){
                int strIndex = Integer.parseInt(sb.toString());
                currRow.add(sharedStr.get(strIndex));
            }else if(type.equals("b")){
                int boolInt = Integer.parseInt(sb.toString());
                if (boolInt == 0) currRow.add("FALSE");
                else currRow.add("TRUE");
            }else if(type.equals("k")){
                currRow.add(formatter.formatRawCellContents(Double.parseDouble(sb.toString()), formatIndex, formatString));
            }
            else currRow.add(sb.toString());
            sb = null;
        }else if(qName.equals("row")){
            if (firstRow && ignoreHead) firstRow = false;
            else table.add(currRow);
        }else if(qName.equals("t") && type.equals("is")){
            currRow.add(sb.toString());
            sb = null;
        }
    }

    public List<List<String>> getTable() {
        return table;
    }
}
