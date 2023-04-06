package com.itheima.utils;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.util.Map;

public class ExcelExportEngine {
    public static Workbook writeToExcel(Object object, Workbook workbook, String imagePath) throws Exception {
        //把一个Bean转成Map
        Map<String, Object> map = EntityUtils.entityToMap(object);

        Sheet sheet = workbook.getSheetAt(0);
        //循环100行,每一行循环100个单位
        Row row = null;
        Cell cell = null;
        for (int i = 0; i < 100; i++) {
            row = sheet.getRow(i);
            if (row == null) {
                break;
            } else {
                for (int j = 0; j < 100; j++) {
                    cell = row.getCell(j);
                    if (cell != null) {
                        writeToCell(cell, map);
                    }

                }
            }

        }
        if (imagePath != null) {
            // 照片的位置
            //开始处理照片
            //先创建一个字节输出流
            ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
            //读取图片 放入了一个带有缓存区的图片类中
            BufferedImage bufferedImage = ImageIO.read(new File(imagePath));
            //把图片写入到了字节输出流中
            String extName = imagePath.substring(imagePath.lastIndexOf(".") + 1).toUpperCase();
            ImageIO.write(bufferedImage, extName, byteArrayOutputStream);
            //Patriarch 控制图片的写入 和ClientAnchor 指定图片的位置
            Drawing patriarch = sheet.createDrawingPatriarch();
            //指定图片的位置         开始列3 开始行2   结束列4  结束行5
            //偏移的单位：是一个英式公制的单位  1厘米=360000
            Sheet sheet1 = workbook.getSheetAt(1);
            int col1 = ((Double) sheet1.getRow(0).getCell(0).getNumericCellValue()).intValue();
            int row1 = ((Double) sheet1.getRow(0).getCell(1).getNumericCellValue()).intValue();
            int col2 = ((Double) sheet1.getRow(0).getCell(2).getNumericCellValue()).intValue();
            int row2 = ((Double) sheet1.getRow(0).getCell(3).getNumericCellValue()).intValue();
            ClientAnchor anchor = new XSSFClientAnchor(0, 0, 0, 0, col1, row1, col2, row2);
            //开始把图片写入到sheet指定的位置
            int format = 0;
            switch (extName) {
                case "JPG": {
                    format = XSSFWorkbook.PICTURE_TYPE_JPEG;
                }
                case "JPEG": {
                    format = XSSFWorkbook.PICTURE_TYPE_JPEG;
                }
                case "PNG": {
                    format = XSSFWorkbook.PICTURE_TYPE_PNG;
                }
            }

            patriarch.createPicture(anchor, workbook.addPicture(byteArrayOutputStream.toByteArray(), format));
            workbook.removeSheetAt(1);
        }

        return workbook;
    }

    /**
     * 比较单元格中的值是否和map中的key一致
     *
     * @param cell
     * @param map
     */
    private static void writeToCell(Cell cell, Map<String, Object> map) {
        CellType cellType = cell.getCellType();
        switch (cellType) {
            case FORMULA: {
                break;
            }
            default: {
                String cellValue = cell.getStringCellValue();
                for (String key : map.keySet()) {
                    if (key.equals(cellValue)) {
                        cell.setCellValue(map.get(key).toString());
                    }
                }
            }
        }
    }
}
