import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.Test;

import java.io.FileInputStream;
import java.util.Date;

public class ExcelReadTest {
    String PATH="F:\\IdeaProjects\\";

    @Test
    public  void testRead03() throws Exception{
        //获取文件流
        FileInputStream inputStream = new FileInputStream(PATH+"poiDemopoi测试03.xls");

        //1.创建一个工作簿，使用excel能操作的这边他都可以操作
        Workbook workbook = new HSSFWorkbook(inputStream);
        //2.得到表
        Sheet sheet = workbook.getSheetAt(0);
        //3.得到行
        Row row = sheet.getRow(1);
        //3.得到列
        Cell cell = row.getCell(1);

        //获取类型
        //getStringCellValue() 字符类型
       System.out.println(cell.getStringCellValue());
       // System.out.println( cell.getNumericCellValue());
        inputStream.close();
    }

    @Test
    public  void testRead07() throws Exception{
        //获取文件流
        FileInputStream inputStream = new FileInputStream(PATH+"poiDemopoi测试07.xlsx");

        //1.创建一个工作簿，使用excl能操作的这边他都可以操作
        Workbook workbook = new XSSFWorkbook(inputStream);
        //2.得到表
        Sheet sheet = workbook.getSheetAt(0);
        //3.得到行
        Row row = sheet.getRow(1);
        //3.得到列
        Cell cell = row.getCell(1);

        //获取类型
        //getStringCellValue() 字符类型
        System.out.println(cell.getStringCellValue());
        // System.out.println( cell.getNumericCellValue());
        inputStream.close();
    }

    @Test
    public void testCellType() throws Exception {
        //获取文件流
        FileInputStream inputStream = new FileInputStream(PATH+"03测试.xls");
        //1.创建一个工作簿，使用excl能操作的这边他都可以操作
        Workbook workbook = new HSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheetAt(0);
        //获取标题内容
        Row rowTitle = sheet.getRow(0);
        if (rowTitle != null){

            int cellCount = rowTitle.getPhysicalNumberOfCells();
            for (int cellNum = 0;cellNum < cellCount;cellNum++){
                Cell cell = rowTitle.getCell(cellNum);
                if (cell !=null){
                    int cellType = cell.getCellType();
                    String cellValue = cell.getStringCellValue();
                    System.out.print(cellValue+" | ");
                }
            }
            System.out.println();
        }
        //获取内容
        int rowCount = sheet.getPhysicalNumberOfRows();
        for (int rowNum = 1; rowNum < rowCount;rowNum++){
            Row rowData = sheet.getRow(rowNum);
            if (rowData !=null){
                //读取列
                int cellCount = rowTitle.getPhysicalNumberOfCells();
                for (int cellNum = 0;cellNum < cellCount;cellNum++){
                   // System.out.println("["+(rowNum+1)+"-"+(cellNum+1)+"]");
                    Cell cell = rowData.getCell(cellNum);
                    //匹配的数据类型
                    if (cell != null){
                        int cellType = cell.getCellType();
                        String cellValue = "";
                        switch (cellType){
                            case HSSFCell.CELL_TYPE_STRING: //字符串
                               // System.out.println("[String]");
                                cellValue = cell.getStringCellValue();
                                break;
                            case HSSFCell.CELL_TYPE_BOOLEAN: //布尔
                                //System.out.println("[Boolean]");
                                cellValue =String.valueOf( cell.getBooleanCellValue());
                                break;
                            case HSSFCell.CELL_TYPE_BLANK: //空
                              //  System.out.println("[Blank]");
                                break;

                            case HSSFCell.CELL_TYPE_NUMERIC: //数字
                               // System.out.println("[Number]");
                                if (HSSFDateUtil.isCellDateFormatted(cell)){  //日期
                                    System.out.println("日期");
                                    Date date = cell.getDateCellValue();
                                    cellValue = new DateTime(date).toString("yyyy-MM-dd");
                                }else {
                                    //不是日期格式，防止数字过长
                                   // System.out.println("转换成字符串输出");
                                    cell.setCellType(HSSFCell.CELL_TYPE_STRING);
                                    cellValue =cell.toString();
                                    break;
                                }
                            case HSSFCell.CELL_TYPE_ERROR: //Error
                                System.out.println("[数据类型错误]");
                                break;
                        }
                        System.out.println(cellValue);

                    }
                }
            }
        }
        inputStream.close();
    }

    @Test
    public  void testFormula() throws  Exception{
        //获取文件流
        FileInputStream inputStream = new FileInputStream(PATH+"03公式测试.xls");
        Workbook workbook = new HSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheetAt(0);

        Row row = sheet.getRow(5);
        Cell cell = row.getCell(0);

        //拿到计算公式
        FormulaEvaluator formulaEvaluator = new HSSFFormulaEvaluator((HSSFWorkbook) workbook);

        //输出单元格内容
        int cellType = cell.getCellType();

        switch (cellType){
            case Cell.CELL_TYPE_FORMULA: //公式
                String formula = cell.getCellFormula();
                System.out.println(formula);
                //计算
                CellValue evaluate = formulaEvaluator.evaluate(cell);
                String cellValue = evaluate.formatAsString();
                System.out.println(cellValue);
                break;

        }
    }
}


