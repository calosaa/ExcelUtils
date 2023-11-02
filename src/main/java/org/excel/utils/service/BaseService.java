package org.excel.utils.service;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.util.List;

public interface BaseService {
    void createWorkBook();
    void createWorkBook(String[] sheets);

    Workbook getWorkBook();
    void addSheet(String sheet);
    Sheet getSheet(String sheet);
    void removeSheet(String sheet);

    void readFile(File file);
    void readSheet(String sheet);
    void writeFile(File file);

    Row addRow();
    Row insertRow(int index);
    List<Row> insertRows(int start,int rows);
    Cell getCell(int row,int cell);
    Row getRow(int row);
    void removeCell(int row,int cell);
    void removeRow(int row);

    void close();
    void batchProcess(BatchHandler handler);

    interface BatchHandler{
        void doHandler(int index,Row row);  // index/row start with 0
    }

}
