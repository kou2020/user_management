package com.itheima.test;

//测试百万数据的导入
public class POIDemo5 {

    public static void main(String[] args) throws Exception {
//        XSSFWorkbook workbook = new XSSFWorkbook("C:\\Users\\syl\\Desktop\\百万用户数据的导出.xlsx");
//        XSSFSheet sheetAt = workbook.getSheetAt(0);
//        String stringCellValue = sheetAt.getRow(0).getCell(0).getStringCellValue();
//        System.out.println(stringCellValue);
        new ExcelParse().parse("C:\\Users\\shang\\Downloads\\百万用户数据的导出.xlsx");

    }

}
