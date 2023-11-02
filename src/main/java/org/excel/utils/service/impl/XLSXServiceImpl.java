package org.excel.utils.service.impl;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.*;
import org.excel.utils.service.BaseService;
import org.excel.utils.service.tools.Heap;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

public class XLSXServiceImpl implements BaseService {
    private XSSFWorkbook workbook;
    private XSSFSheet curSheet;


    @Override
    public Workbook getWorkBook() {
        return workbook;
    }

    @Override
    public void createWorkBook() {
        if(workbook==null) workbook = new XSSFWorkbook();
    }

    @Override
    public void createWorkBook(String[] sheets) {
        if (workbook==null){
            workbook = new XSSFWorkbook();
            for (String sheet : sheets) {
                workbook.createSheet(sheet);
            }
        }
    }

    @Override
    public void addSheet(String sheet) {
        workbook.createSheet(sheet);
    }

    @Override
    public Sheet getSheet(String sheet) {
        return workbook.getSheet(sheet);
    }

    @Override
    public void removeSheet(String sheet) {
        workbook.removeSheetAt(workbook.getSheetIndex(sheet));
    }

    @Override
    public void readFile(File file) {
        try {
            FileInputStream fileInputStream = new FileInputStream(file);
            workbook = new XSSFWorkbook(fileInputStream);
            fileInputStream.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Override
    public void readSheet(String sheet) {
            curSheet = workbook.getSheet(sheet);
    }

    @Override
    public void writeFile(File file) {

        try {
            FileOutputStream fos = new FileOutputStream(file);
            workbook.write(fos);
            fos.close();
        } catch (IOException e) {
            e.printStackTrace();
        }finally {

        }
    }

    @Override
    public XSSFRow addRow(){
        int lastRowNum = curSheet.getLastRowNum();
        return curSheet.createRow(lastRowNum + 1);
    }

    public XSSFCell addCell(XSSFRow row,int index){
        return row.createCell(index);
    }
    @Override
    public XSSFRow insertRow(int index) {
        curSheet.shiftRows(index,curSheet.getLastRowNum(),1);
        return curSheet.createRow(index);
    }

    @Override
    public List<Row> insertRows(int start, int rowN) {
        curSheet.shiftRows(start,curSheet.getLastRowNum(),rowN);
        List<Row> rows = new ArrayList<>();
        for (int i = 0; i < rowN; i++) {
            rows.add(curSheet.createRow(i+start));
        }
        return rows;
    }

    @Override
    public Cell getCell(int row, int cell) {
        return curSheet.getRow(row).getCell(cell);
    }

    @Override
    public XSSFRow getRow(int row) {
        return curSheet.getRow(row);
    }

    @Override
    public void removeCell(int row, int cell) {
        curSheet.getRow(row).removeCell(curSheet.getRow(row).getCell(cell));
    }

    @Override
    public void removeRow(int row) {
        curSheet.removeRow(curSheet.getRow(row));
    }

    @Override
    public void batchProcess(BatchHandler handler) {
        int first = curSheet.getFirstRowNum();
        int last = curSheet.getLastRowNum();
        for(int i=first;i<=last;i++){
            handler.doHandler(i,curSheet.getRow(i));
        }
    }

    /**
     *
     * @param fromArea 表源 例如 ： "Sheet1!A1:D12"
     * @param toCell 透视表生成位置 例如 ："Sheet2!A2"
     * @param rowTitle 头标签
     * @param colsI 列标签索引 从0开始
     * @param colsN 列标签在透视表中名称
     * @param fun 计算方法，默认求和
     */
    public void createPivotTable(String fromArea,String toCell,int rowTitle,int[] colsI,String[] colsN,DataConsolidateFunction fun){
        XSSFPivotTable pivotTable = curSheet.createPivotTable(new AreaReference(fromArea, null), new CellReference(toCell));
        pivotTable.addRowLabel(rowTitle);

        if(colsI.length!=colsN.length){
            System.err.println(" 列索引与列标数量签不一致");
            return;
        }
        if(fun==null) fun = DataConsolidateFunction.SUM;
        for (int i=0;i<colsI.length;i++) {
            pivotTable.addColumnLabel(fun, colsI[i], colsN[i]);
        }
    }

    /**
     *
     * @param t_x  左上角横坐标
     * @param t_y  左上角纵坐标
     * @param b_x 右下角横坐标
     * @param b_y 右下角纵坐标
     * @param index 排序索引
     * @param s true 正序 false 倒序
     */
    public void sortByNumber(int t_x,int t_y,int b_x,int b_y,int index,boolean s){
        addSheet("cur_cache");
        XSSFSheet cur_cache = (XSSFSheet)getSheet("cur_cache");
        double[] heads = new double[b_y-t_y+1];
        HashMap<Double,XSSFCell[]> map = new HashMap<>();
        System.out.println("获取源表内容");
        for (int r=t_y;r<=b_y;r++){
            XSSFCell[] row = new XSSFCell[b_x-t_x+1];
            XSSFRow row_cache = cur_cache.createRow(r);
            System.out.println("读取第"+r+"行");
            for(int c=t_x;c<=b_x;c++){
                XSSFCell cell_cache = row_cache.createCell(c);
                cell_cache.copyCellFrom((XSSFCell)getCell(r, c),new CellCopyPolicy());
                row[c-t_x] = cell_cache;
            }
            heads[r-t_y] =row[index].getNumericCellValue();
            map.put(heads[r-t_y],row);
        }
        System.out.println("排序中。。。");
        Heap.heapSort(heads);
        System.out.println("排序完成");
        if(s) {
            for (int r = t_y; r <= b_y; r++) {
                XSSFCell[] row = map.get(heads[r - t_y]);
                for (int c = t_x; c <= b_x; c++) {
                    XSSFCell cell = (XSSFCell) getCell(r, c);
                    cell.copyCellFrom(row[c-t_x], new CellCopyPolicy());
                }
            }
        }else{
            for (int r = t_y; r <= b_y; r++) {
                XSSFCell[] row = map.get(heads[b_y - r]);
                for (int c = t_x; c <= b_x; c++) {
                    XSSFCell cell = (XSSFCell) getCell(r, c);
                    cell.copyCellFrom(row[c-t_x], new CellCopyPolicy());
                }
            }
        }
        removeSheet("cur_cache");
        System.out.println("修改完成");
    }

    public void sortByStr(int t_x,int t_y,int b_x,int b_y,int index,boolean s){
        addSheet("cur_cache");
        XSSFSheet cur_cache = (XSSFSheet)getSheet("cur_cache");
        String[] heads = new String[b_y-t_y+1];
        HashMap<String,XSSFCell[]> map = new HashMap<>();
        System.out.println("获取源表内容");
        for (int r=t_y;r<=b_y;r++){
            XSSFCell[] row = new XSSFCell[b_x-t_x+1];
            XSSFRow row_cache = cur_cache.createRow(r);
            System.out.println("读取第"+r+"行");
            for(int c=t_x;c<=b_x;c++){
                XSSFCell cell_cache = row_cache.createCell(c);
                //注意此处可能为空，检查文件是否保存，以及源表范围是否正确,下标从0开始
                cell_cache.copyCellFrom((XSSFCell)getCell(r, c),new CellCopyPolicy());
                row[c-t_x] = cell_cache;
            }
            heads[r-t_y] =row[index].getStringCellValue();
            map.put(heads[r-t_y],row);
        }
        System.out.println("排序中。。。");
        Heap.strHeapSort(heads);
        System.out.println("排序完成");
        if(s) {
            for (int r = t_y; r <= b_y; r++) {
                XSSFCell[] row = map.get(heads[r - t_y]);
                for (int c = t_x; c <= b_x; c++) {
                    XSSFCell cell = (XSSFCell) getCell(r, c);
                    cell.copyCellFrom(row[c-t_x], new CellCopyPolicy());
                }
            }
        }else{
            for (int r = t_y; r <= b_y; r++) {
                XSSFCell[] row = map.get(heads[b_y - r]);
                for (int c = t_x; c <= b_x; c++) {
                    XSSFCell cell = (XSSFCell) getCell(r, c);
                    cell.copyCellFrom(row[c-t_x], new CellCopyPolicy());
                }
            }
        }
        removeSheet("cur_cache");
        System.out.println("修改完成");
    }

    @Override
    public void close() {
        try {
            workbook.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public void copyArea(int t_x,int t_y,int b_x,int b_y,int to_x,int to_y,XSSFSheet toSheet){
        if((to_x>=t_x && to_x<=b_x) && (to_y>=t_y && to_y<=b_y)){
            System.out.println("目标在源范围内，无法复制");
            return;
        }

        for(int r=t_y;r<=b_y;r++){
            XSSFRow to_row = toSheet.getRow(to_y + r - t_y);
            if(to_row==null) to_row = toSheet.createRow(to_y + r - t_y);
            for (int c=t_x;c<=b_x;c++){
                XSSFCell to_cell = to_row.getCell(to_x + c - t_x);
                if(to_cell==null) to_cell = to_row.createCell(to_x + c - t_x);
                to_cell.copyCellFrom(getCell(r,c),new CellCopyPolicy());
            }
        }
    }

    /**
     *
     * @param styleName e.g. 黑体 or 宋体 ...
     * @param bgColor e.g. IndexedColors.RED.getIndex()
     * @param fontColor e.g. IndexedColors.BLACK.getIndex()
     * @param cell XSSFCell
     */
    public void setCellStyle(String styleName,short bgColor,short fontColor,XSSFCell cell){
        XSSFCellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setFillForegroundColor(bgColor);
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        XSSFFont font = workbook.createFont();
        font.setFontName(styleName);
        font.setColor(fontColor);
        cellStyle.setFont(font);
        cell.setCellStyle(cellStyle);
    }

    public void areaBatchProcess(int t_x,int t_y,int b_x,int b_y,AreaBatchHandler handler){
        for(int r=t_y;r<=b_y;r++){
            for (int c=t_x;c<=b_x;c++){
                handler.doHandler(r,c,curSheet.getRow(r).getCell(c));
            }
        }
    }

    public interface AreaBatchHandler{
        void doHandler(int x,int y,XSSFCell cell);  // index/row start with 0
    }
}
