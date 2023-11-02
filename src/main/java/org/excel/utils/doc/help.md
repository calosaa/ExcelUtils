# 服务(service)
 - impl
   - LargeDataServiceImpl 批量处理
   ```
   //当使用一种handler时，另外一个handler返回为null
   public interface LargeDataHandler<D>{
        void doWriteHandler(D data,SXSSFRow row); //write / modify function use
        D doReadHandler(SXSSFRow row);
    }
   ```
   - XLSServiceImpl .xls文件处理
   - XLSXServiceImpl .xlsx文件处理
 - tools
   - Heap 堆排序算法
 - BaseService 基础接口
# Cell
1.创建Cell
```java
Row.createCell(i) //i 表示该行索引，从0开始
```
<font color="red">下面方法弃用</font>
```java
#service.addCell(...) 
```

2.填写公式
```java
Cell.setCellFormula(formula) //formula 表示公式，不需要开头等号
```
3.透视表
```java
XSSFPivotTable pivotTable = sheet3.createPivotTable(new AreaReference("Sheet2!A1:C11", null), new CellReference("Sheet3!C3"));//创建透视表，null表示默认sheet版本
pivotTable.addRowLabel(index); //添加行标签
pivotTable.addColumnLabel(DataConsolidateFunction.SUM, index,colName); //添加列标签，colname设置列名,SUM表示求和

```
# ArrayList
 - 排序
    ```
   list.sort(Comparator.comparing(Data::method);
   list.sort(new Comparator<Data>(){
      @Override
      public int compare(Data o1, Data o2) {
         return Integer.compare(o1.getData(),o2.getData()); //此为升序
         //return Integer.compare(o1.getData(),o2.getData()) 此为降序
      }
   });
   ```
# 数据格式
 - 数字
   > 数据格式t="n"时，注意要先转为double再转为int（此为优先处理方式）