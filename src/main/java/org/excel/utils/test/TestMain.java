package org.excel.utils.test;

import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.excel.utils.service.impl.XLSXServiceImpl;

import java.io.File;

public class TestMain {
    public static void main(String[] args) {
        File file = new File("C:\\Users\\86136\\Desktop\\test1.xlsx");
        XLSXServiceImpl service = new XLSXServiceImpl();
        service.readFile(file);
        service.readSheet("Sheet2");
        //service.sortByNumber(7,1,7,8,0,true);
        //service.addSheet("sheet4");
        //service.copyArea(0,1,2,10,10,10,(XSSFSheet) service.getSheet("sheet4"));
        service.areaBatchProcess(0,0,2,0,(x,y,cell)->{
            service.setCellStyle("宋体", IndexedColors.GREY_50_PERCENT.getIndex(), IndexedColors.YELLOW.getIndex(), cell);
        });
        service.writeFile(file);
        service.close();
    }

    /*public static void main(String[] args) {
        HashMap<Integer,String> map = new HashMap<>();
        map.put(2,"cheng");
        map.put(1,"wang");
        map.put(5,"li");
        map.put(4,"fang");
        map.put(9,"yan");
        map.put(6,"dai");
        int[] list = new int[]{1,2,4,5,6,9};
        Heap.heapSort(list);
        for (int i : list) {
            System.out.println(map.get(i));
        }
    }*/
}
