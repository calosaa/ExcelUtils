package org.excel.utils.service.impl;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.excel.utils.service.BaseService;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class XLSServiceImpl implements BaseService {
    private HSSFWorkbook workbook;
    private HSSFSheet curSheet;

    @Override
    public Workbook getWorkBook() {
        return workbook;
    }

    @Override
    public void createWorkBook() {
        workbook = new HSSFWorkbook();
    }

    @Override
    public void createWorkBook(String[] sheets) {
        workbook = new HSSFWorkbook();
        for (String sheet : sheets) {
            workbook.createSheet(sheet);
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
            workbook = new HSSFWorkbook(fileInputStream);
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
        }
    }

    @Override
    public HSSFRow addRow() {
        int index = curSheet.getLastRowNum();
        return curSheet.createRow(index + 1);
    }

    public HSSFCell addCell(HSSFRow row,int index){
        return row.createCell(index);
    }
    @Override
    public HSSFRow insertRow(int index) {
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
    public HSSFRow getRow(int row) {
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

    @Override
    public void close() {
        try {
            workbook.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}
