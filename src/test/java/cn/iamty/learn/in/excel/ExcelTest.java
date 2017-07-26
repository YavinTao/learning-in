package cn.iamty.learn.in.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

public class ExcelTest {

    @Test
    public void testGenerateExcel() throws Exception {

        XSSFWorkbook wb = new XSSFWorkbook();

        //创建工作表
        Sheet sheet = wb.createSheet("test-sheet");

        List<User> userList = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            User user = new User();
            user.setId(i + 1);
            user.setName("zhangsan" + i);
            user.setAge(i * 10);
            user.setMail("zhangsan" + i + "@gmail.com");
            userList.add(user);
        }
        String headers[] = new String[]{"ID", "用户名", "年龄", "邮箱"};

        Row title_row = sheet.createRow(0);
        title_row.setHeight((short) (40 * 20));

        Cell title_cell = title_row.createCell(0);

        XSSFCellStyle style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderLeft(BorderStyle.MEDIUM);
        style.setBorderRight(BorderStyle.MEDIUM);
        style.setBorderTop(BorderStyle.MEDIUM);
        style.setBorderBottom(BorderStyle.MEDIUM);



        XSSFCellStyle backGround = wb.createCellStyle();
        backGround.setAlignment(HorizontalAlignment.CENTER);
        backGround.setVerticalAlignment(VerticalAlignment.CENTER);
        backGround.setBorderLeft(BorderStyle.MEDIUM);
        backGround.setBorderRight(BorderStyle.MEDIUM);
        backGround.setBorderTop(BorderStyle.MEDIUM);
        backGround.setBorderBottom(BorderStyle.MEDIUM);
        backGround.setFillBackgroundColor(IndexedColors.RED.getIndex());
        backGround.setFillPattern(FillPatternType.SOLID_FOREGROUND);


        title_cell.setCellStyle(backGround);
        title_cell.setCellValue("用户详细信息");

        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, headers.length - 1));


        Row header_row = sheet.createRow(1);
        header_row.setHeight((short) (20 * 24));

        for (int i = 0; i < headers.length; i++) {
            sheet.setColumnWidth(i, 30 * 256);
            Cell cell = header_row.createCell(i);
            cell.setCellStyle(style);
            cell.setCellValue(headers[i]);
        }

        for (int i = 0; i < userList.size(); i++) {
            Row row = sheet.createRow(i + 2);
            row.setHeight((short) (20 * 20));
            Cell cell = row.createCell(0);
            cell.setCellStyle(style);
            cell.setCellValue(userList.get(i).getId());
            cell = row.createCell(1);
            cell.setCellStyle(style);
            cell.setCellValue(userList.get(i).getName());
            cell = row.createCell(2);
            cell.setCellStyle(style);
            cell.setCellValue(userList.get(i).getAge());
            cell = row.createCell(3);
            cell.setCellStyle(style);
            cell.setCellValue(userList.get(i).getMail());

        }


        FileOutputStream fileOut = new FileOutputStream("E:/test.xlsx");
        wb.write(fileOut);
        fileOut.close();

    }


    @Test
    public void testGenerateWorkData() {


    }


}
