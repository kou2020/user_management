package com.itheima.test;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;

//导出一个有样式的excel
//    1、边框线 2、合并单元格 3、行高列宽 4、对齐方式 5、字体
public class POIDemo4 {
    public static void main(String[] args) throws Exception {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("有样式的数据");
        //创建样式
        //需求：1、边框线：全边框  2、行高：42   3、合并单元格：第1行的第1个单元格到第5个单元格 4、对齐方式：水平垂直都要居中 5、字体：黑体18号字
        CellStyle bigTitleRowCellStyle = workbook.createCellStyle();
        bigTitleRowCellStyle.setBorderBottom(BorderStyle.THIN); //下边框  BorderStyle.THIN 细线
        bigTitleRowCellStyle.setBorderLeft(BorderStyle.THIN);  //左边框
        bigTitleRowCellStyle.setBorderRight(BorderStyle.THIN);  //右边框
        bigTitleRowCellStyle.setBorderTop(BorderStyle.THIN);  //上边框
        //对齐方式 水平对齐 垂直对齐
        bigTitleRowCellStyle.setAlignment(HorizontalAlignment.CENTER);//水平居中对齐
        bigTitleRowCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);//垂直居中对齐
        //创建字体
        Font bigTitleFont = workbook.createFont();
        bigTitleFont.setFontName("黑体");
        bigTitleFont.setFontHeightInPoints((short) 18);
        //把字体放入样式中
        bigTitleRowCellStyle.setFont(bigTitleFont);

        Row bigTitleRow = sheet.createRow(0);
        //设置行高
        bigTitleRow.setHeightInPoints(42);
        //设置列宽
        sheet.setColumnWidth(0, 5 * 256);
        sheet.setColumnWidth(1, 8 * 256);
        sheet.setColumnWidth(2, 15 * 256);
        sheet.setColumnWidth(3, 15 * 256);
        sheet.setColumnWidth(4, 30 * 256);
        //合并单元格
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 4));
        //添加单元格样式
        for (int i = 0; i < 5; i++) {
            Cell cell = bigTitleRow.createCell(i);
            cell.setCellStyle(bigTitleRowCellStyle);
        }
        //向单元格内放一句话
        sheet.getRow(0).getCell(0).setCellValue("用户信息数据");

        //小标题的样式
        CellStyle littleTitleRowCellStyle = workbook.createCellStyle();
        //样式的克隆
        littleTitleRowCellStyle.cloneStyleFrom(bigTitleRowCellStyle);
        //创建字体  宋体12号字加粗
        Font littleFont = workbook.createFont();
        littleFont.setFontName("宋体");
        littleFont.setFontHeightInPoints((short) 12);
        littleFont.setBold(true);
        //把字体放入到样式中
        littleTitleRowCellStyle.setFont(littleFont);

        //内容的样式
        CellStyle contentRowCellStyle = workbook.createCellStyle();
        //样式的克隆
        contentRowCellStyle.cloneStyleFrom(littleTitleRowCellStyle);
        contentRowCellStyle.setAlignment(HorizontalAlignment.LEFT);
        //创建字体  宋体12号字加粗
        Font contentFont = workbook.createFont();
        contentFont.setFontName("宋体");
        contentFont.setFontHeightInPoints((short) 11);
        contentFont.setBold(false);
        //把字体放入到样式中
        contentRowCellStyle.setFont(contentFont);

        Row titleRow = sheet.createRow(1);
        titleRow.setHeightInPoints(31.5F);
        String[] titles = new String[]{"编号", "姓名", "手机号", "入职日期", "现住址"};
        for (int i = 0; i < 5; i++) {
            Cell cell = titleRow.createCell(i);
            cell.setCellValue(titles[i]);
            cell.setCellStyle(littleTitleRowCellStyle);
        }
        String[] users = new String[]{"1", "大一", "13800000001", "2001-03-29", "北京市西城区宣武大街1号院"};
        Row row = sheet.createRow(2);
        for (int i = 0; i < 5; i++) {
            Cell cell = row.createCell(i);
            cell.setCellValue(users[i]);
            cell.setCellStyle(contentRowCellStyle);
        }
        workbook.write(new FileOutputStream("D://testStyle.xlsx"));
        //关闭资源
        workbook.close();
    }
}
