package org.excel.utils.service.impl;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.excel.utils.service.BaseService;

import java.io.*;
import java.util.List;

public class LargeDataWriteServiceImpl<D> implements BaseService {
    private SXSSFWorkbook workbook;
    private SXSSFSheet curSheet;

    @Override
    public void createWorkBook() {
        if(workbook==null)workbook = new SXSSFWorkbook(100);
    }

    @Override
    public void createWorkBook(String[] sheets) {
        if(workbook==null){
            workbook = new SXSSFWorkbook(100);
            for (String sheet : sheets) {
                workbook.createSheet(sheet);
            }
        }
    }

    @Override
    public Workbook getWorkBook() {
        return workbook;
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
            InputStream fileInputStream = new FileInputStream(file);
            XSSFWorkbook xssfWorkbook = new XSSFWorkbook(fileInputStream);
            workbook = new SXSSFWorkbook(xssfWorkbook);
            fileInputStream.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    @Override
    public void readSheet(String sheet) {
        curSheet = workbook.getSheet(sheet);
    }

    @Override
    public void writeFile(File file) {
        try {
            FileOutputStream fileOutputStream = new FileOutputStream(file);
            workbook.write(fileOutputStream);
            fileOutputStream.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    @Override
    public Row addRow() {
        return null;
    }

    @Override
    public Row insertRow(int index) {
        return null;
    }

    @Override
    public List<Row> insertRows(int start, int rows) {
        return null;
    }

    @Override
    public Cell getCell(int row, int cell) {
        return null;
    }

    @Override
    public Row getRow(int row) {
        return null;
    }

    @Override
    public void removeCell(int row, int cell) {

    }

    @Override
    public void removeRow(int row) {

    }

    @Override
    public void close() {

    }

    @Override
    public void batchProcess(BatchHandler handler) {

    }

    public void writeData(List<D> data,int start,LargeDataHandler<D> handler){
        if(start==0) return;

        for (D datum : data) {
            SXSSFRow row = curSheet.createRow(start++);
            handler.doWriteHandler(datum, row);
        }
    }

    public void modifyData(int row, D data, LargeDataHandler<D> handler){
        handler.doWriteHandler(data,curSheet.getRow(row));
    }

    public List<D> readData(List<D> data,int start,int end,LargeDataHandler<D> handler){
        int lastRowNum = curSheet.getLastRowNum();
        if(start>lastRowNum && end>lastRowNum && data==null) return null;
        for (int i = start; i <= end; i++) {
            SXSSFRow row = curSheet.getRow(i);
            if(row!=null) data.add(handler.doReadHandler(row));
        }
        return data;
    }
    public void writeHead(String[] heads){
        SXSSFRow row = curSheet.createRow(0);
        for (int i = 0; i < heads.length; i++) {
            SXSSFCell cell = row.createCell(i);
            cell.setCellValue(heads[i]);
        }
    }

    public interface LargeDataHandler<D>{
        void doWriteHandler(D data,SXSSFRow row);
        D doReadHandler(SXSSFRow row);
    }
}
