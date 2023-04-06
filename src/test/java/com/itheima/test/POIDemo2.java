package com.itheima.test;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;

//创建一个高版本的excel,并且向其中的一个单元格中随便写一句话
public class POIDemo2 {
    public static void main(String[] args) throws Exception {
        //创建一个全新的工作簿
        Workbook workbook = new XSSFWorkbook();
        //在工作簿中创建工作表
        Sheet sheet = workbook.createSheet("POI操作Excel");
        //在工作表中创建行
        Row row = sheet.createRow(0);
        //在行中创建单元
        Cell cell = row.createCell(0);
        //在单元格中写入内容
        cell.setCellValue("这是我第一次使用POI");

        workbook.write(new FileOutputStream("d:/poiTest.xlsx"));

        workbook.close();
    }
}
