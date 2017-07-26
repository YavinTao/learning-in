package cn.iamty.learn.in.excel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;

public class TestExcel {

    /**
     * Excel文档的构成
     *
     * 在工作簿(WorkBook)里面包含了工作表(Sheet) 在工作表里面包含了行(Row) 行里面包含了单元格(Cell)
     *
     *
     * 创建一个工作簿的基本步骤
     *
     * 第一步 创建一个 工作簿 第二步 创建一个 工作表 第三步 创建一行 第四步 创建单元格 第五步 写数据 第六步
     * 将内存中生成的workbook写到文件中 然后释放资源
     *
     */

    public static void testCreateFirstExcel97() throws Exception {
        Workbook wb = new HSSFWorkbook();
        FileOutputStream fileOut = new FileOutputStream("C:/workbook.xls");
        wb.write(fileOut);
        fileOut.close();
    }
    public static void testCreateFirstExcel07() throws Exception {
        Workbook wb = new XSSFWorkbook();
        FileOutputStream fileOut = new FileOutputStream("C:/workbook.xlsx");
        wb.write(fileOut);
        fileOut.close();
    }


    public static void createExcelOfData() throws Exception{
        Workbook wb = new HSSFWorkbook();

        //创建工作表
        Sheet sheet = wb.createSheet("测试Excel");

        //创建单元格   单元格是隶属于行

        Row row = sheet.createRow(0);   //起始从0开始

        Cell cell = row.createCell(0);

        cell.setCellValue("This is a test");
        FileOutputStream fileOut = new FileOutputStream("C:/test.xls");
        wb.write(fileOut);
        fileOut.close();
    }

    public static void createExcelOfUsers() throws Exception{
        Workbook wb = new XSSFWorkbook();

        //创建工作表
        Sheet sheet = wb.createSheet("用户信息");

        Object [][] data = new Object[][]{
                {1,"zhansgan",20,"zhangsan@zhangsan.com"},
                {2,"zhansgan1",30,"zhangsan1@zhangsan.com"},
                {3,"zhansgan2",40,"zhangsan2@zhangsan.com"},
                {4,"zhansgan3",20,"zhangsan3@zhangsan.com"},
                {5,"zhansgan4",43,"zhangsan4@zhangsan.com"},
                {6,"zhansgan5",24,"zhangsan5@zhangsan.com"},
                {7,"zhansgan6",55,"zhangsan6@zhangsan.com"},
                {8,"zhansgan7",12,"zhangsan7@zhangsan.com"},
        };


        //显示标题
        Row title_row = sheet.createRow(0);
        title_row.setHeight((short)(40*20));

        Cell title_cell = title_row.createCell(0);


        String headers[] = new String[]{"ID","用户名","年龄","邮箱"};

        Row header_row = sheet.createRow(1);
        header_row.setHeight((short)(20*24));

        //创建单元格的 显示样式
        CellStyle style = wb.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_CENTER); //水平方向上的对其方式
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);  //垂直方向上的对其方式


        title_cell.setCellStyle(style);
        title_cell.setCellValue("用户详细信息");

        sheet.addMergedRegion(new CellRangeAddress(0,0,0,headers.length-1));



        for(int i=0;i<headers.length;i++){
            //设置列宽   基数为256
            sheet.setColumnWidth(i, 30*256);
            Cell cell = header_row.createCell(i);
            //应用样式到  单元格上
            cell.setCellStyle(style);
            cell.setCellValue(headers[i]);
        }



        for(int i=0;i<data.length;i++){

            Row row = sheet.createRow(i+2);
            row.setHeight((short)(20*20)); //设置行高  基数为20
            for(int j=0;j<data[i].length;j++){
                Cell cell = row.createCell(j);
                cell.setCellValue(data[i][j].toString());
            }

        }
        FileOutputStream fileOut = new FileOutputStream("E:/users.xls");
        wb.write(fileOut);
        fileOut.close();
    }


    public static void main(String[] args) throws Exception {
//      testCreateFirstExcel97();
//      testCreateFirstExcel07();

//      createExcelOfData();
        createExcelOfUsers();
    }

}
