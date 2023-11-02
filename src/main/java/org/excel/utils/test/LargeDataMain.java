package org.excel.utils.test;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.excel.utils.service.LargeDataRead.IndividualObjConvertTableHandler;
import org.excel.utils.service.LargeDataRead.LargeReader;
import org.excel.utils.service.impl.LargeDataWriteServiceImpl;
import org.xml.sax.SAXException;

import javax.xml.parsers.ParserConfigurationException;
import java.io.File;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.concurrent.TimeUnit;

public class LargeDataMain {
    /*public static void main(String[] args) {
        List<Student> students = new ArrayList<>();
        System.out.println("创建对象列表");
        for (int i = 0; i < 1048575; i++) {
            students.add(new Student("cheng",i%100, i % 2 == 0));
        }
        System.out.println("创建完成");
        LargeDataWriteServiceImpl<Student> service = new LargeDataWriteServiceImpl<>();
        service.createWorkBook();
        service.addSheet("sheet");
        service.readSheet("sheet");
        service.writeHead(new String[]{"name","age","sex"});
        System.out.println("开始写数据");
        service.writeData(students, 1, new LargeDataWriteServiceImpl.LargeDataHandler<Student>() {
            @Override
            public void doWriteHandler(Student data, SXSSFRow row) {
                SXSSFCell name = row.createCell(0);
                name.setCellValue(data.getName());
                SXSSFCell age = row.createCell(1);
                age.setCellValue(data.getAge());
                SXSSFCell sex = row.createCell(2);
                sex.setCellValue(data.isSex());
            }

            @Override
            public Student doReadHandler(SXSSFRow row) {
                return null;
            }
        });
        System.out.println("数据写完");
        File file = new File("D:\\large-data.xlsx");
        service.writeFile(file);
    }*/

    public static void main(String[] args) {
        try {
            LargeReader reader = new LargeReader("D:\\projects\\idea\\CTExcelUtils\\CTExcelUtils\\large-data.xlsx");
            System.out.println("start");
            System.out.println("parsing...");
            ArrayList<String> header = new ArrayList<>();
            //List<List<String>> table = reader.parse(reader.selectSheet("sheet"),true);
            long stime = System.currentTimeMillis();
            List<Student> table = reader.parse2obj(reader.selectSheet("sheet"), new IndividualObjConvertTableHandler.ConvertHandler<Student>() {
                @Override
                public Student convert(ArrayList<String> str,int index) {
                    if (index==1) {
                        header.addAll(str);
                        return null;
                    }
                    return new Student(str.get(0),(int) Double.parseDouble(str.get(1)),str.get(2).equals("TRUE"));
                }
            });
            long etime = System.currentTimeMillis();
            System.out.println("parsing time : " + (etime - stime));

            System.out.println("sorting...");
            /*table.sort(new Comparator<List<String>>() {
                @Override
                public int compare(List<String> o1, List<String> o2) {
                    return Integer.compare((int) Double.parseDouble(o2.get(1)),
                            (int) Double.parseDouble(o1.get(1)));

                }
            });*/
            table.sort(new Comparator<Student>() {
                @Override
                public int compare(Student o1, Student o2) {
                    return Integer.compare(o1.getAge(),o2.getAge());
                }
            });

            System.out.println("create write service");
            //LargeDataWriteServiceImpl<List<String>> service = new LargeDataWriteServiceImpl<>();
            LargeDataWriteServiceImpl<Student> service = new LargeDataWriteServiceImpl<>();
            service.createWorkBook();
            service.addSheet("sheet");
            service.readSheet("sheet");
            //service.writeHead(new String[]{"名称","年龄","性别"});
            service.writeHead(header.toArray(new String[header.size()]));
            System.out.println(header.toString());
            System.out.println("writing...");
            CellStyle cellStyle1 = service.getWorkBook().createCellStyle();
            cellStyle1.setFillForegroundColor(IndexedColors.BLUE.getIndex());
            cellStyle1.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            CellStyle cellStyle2 = service.getWorkBook().createCellStyle();
            cellStyle2.setFillForegroundColor(IndexedColors.RED.getIndex());
            cellStyle2.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            service.writeData(table, 1, new LargeDataWriteServiceImpl.LargeDataHandler<Student>() {
                @Override
                public void doWriteHandler(Student data, SXSSFRow row) {
                    SXSSFCell nameCell = row.createCell(0);
                    nameCell.setCellValue(data.getName());
                    SXSSFCell ageCell = row.createCell(1);
                    ageCell.setCellValue(data.getAge());
                    SXSSFCell sexCell = row.createCell(2);
                    sexCell.setCellValue(data.isSex());


                    if (data.isSex())sexCell.setCellStyle(cellStyle1);
                    else sexCell.setCellStyle(cellStyle2);


                }

                @Override
                public Student doReadHandler(SXSSFRow row) {
                    return null;
                }
            });
            System.out.println("write to file");
            service.writeFile(new File("D:\\large-data2.xlsx"));
            System.out.println("done");
        } catch (ParserConfigurationException | SAXException e) {
            throw new RuntimeException(e);
        }
    }
}
