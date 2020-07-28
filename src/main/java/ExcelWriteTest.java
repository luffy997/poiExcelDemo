import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.Test;


import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;


public class ExcelWriteTest {

    String PATH="F:\\IdeaProjects\\poiDemo";
    @Test
    public void testWrite03() throws Exception {
        //1.创建一个工作簿
        // 07使用 XSSFWorkbook（）创建对象
        // 03版 HSSFWorkbook()创建对象
        Workbook workbook = new XSSFWorkbook();
        //2.创建一个工作表
        Sheet sheet = workbook.createSheet("学生信息统计表");
        //3.创建一个行 (1,1)
        Row row1=sheet.createRow(0);
        //4.创建一个单元格
        Cell cell11 = row1.createCell(0);
        cell11.setCellValue("学号");
        //（1.2）
        Cell cell12 = row1.createCell(1);
        cell12.setCellValue("No.1");

        //第二行(2,1)
        Row row2=sheet.createRow(1);
        Cell cell21 = row2.createCell(0);
        cell21.setCellValue("日期");
        //(2,2)
        Cell cell22 = row2.createCell(1);
//        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd-Es HH-mm-ss");
//        String time =simpleDateFormat.format(new Date());
        String time = new DateTime().toString("yyyy-MM-dd Es HH:mm:ss");
        cell22.setCellValue(time);

        //生成表(IO流)  07版本使用xlsx结尾 03版本使用xsl结尾
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "poi测试07.xlsx");
        workbook.write(fileOutputStream);
        //关闭流
        fileOutputStream.close();
        System.out.println("文件生成完毕");

    }

    @Test
    //大量数据的写入 03
    public void testWrite03BigData() throws Exception {
        //时间
        long begin=System.currentTimeMillis();

        //创建一个工作簿
        Workbook workbook = new HSSFWorkbook();
        //创建一张表
        Sheet sheet = workbook.createSheet();
        //写入数据
        for (int rowNum = 0; rowNum < 65536; rowNum++) {
            Row row = sheet.createRow(rowNum);
            for (int cellNum = 0;cellNum < 10;cellNum ++){
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }
        System.out.println("运行完毕");
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "testWrite03BigData03.xls");
        workbook.write(fileOutputStream);
        fileOutputStream.close();
        long end=System.currentTimeMillis();
        System.out.println((double) (end-begin)/1000);
    }

    @Test
    //大量数据的写入 07 耗时较长
    public void testWrite07BigData() throws Exception {
        //时间
        long begin=System.currentTimeMillis();

        //创建一个工作簿
        Workbook workbook = new XSSFWorkbook();
        //创建一张表
        Sheet sheet = workbook.createSheet();
        //写入数据
        for (int rowNum = 0; rowNum < 10000; rowNum++) {
            Row row = sheet.createRow(rowNum);
            for (int cellNum = 0;cellNum < 10;cellNum ++){
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }
        System.out.println("运行完毕");
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "testWrite03BigData07.xlsx");
        workbook.write(fileOutputStream);
        fileOutputStream.close();
        long end=System.currentTimeMillis();
        System.out.println((double) (end-begin)/1000);
    }

    @Test
    //大量数据的写入 07  使用加速版 SXSSFWorkbook()
    public void testWrite07BigDataSuper() throws Exception {
        //时间
        long begin=System.currentTimeMillis();
        //创建一个工作簿
        Workbook workbook = new SXSSFWorkbook();
        //创建一张表
        Sheet sheet = workbook.createSheet();
        //写入数据
        for (int rowNum = 0; rowNum < 10000; rowNum++) {
            Row row = sheet.createRow(rowNum);
            for (int cellNum = 0;cellNum < 10;cellNum ++){
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }
        System.out.println("运行完毕");
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "testWrite03BigData07Super.xlsx");
        workbook.write(fileOutputStream);
        fileOutputStream.close();
        //清楚临时文件!
        ((SXSSFWorkbook)workbook).dispose();
        long end=System.currentTimeMillis();
        System.out.println((double) (end-begin)/1000);
    }
}
