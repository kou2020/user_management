package com.itheima.test;

import net.sf.jasperreports.engine.JREmptyDataSource;
import net.sf.jasperreports.engine.JasperExportManager;
import net.sf.jasperreports.engine.JasperFillManager;
import net.sf.jasperreports.engine.JasperPrint;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.HashMap;
import java.util.Map;

/**
 * 使用JasperReport导出pdf
 */
public class PDFDemo {
    public static void main(String[] args) throws Exception {
        String filePath = "D:\\test01.jasper";//模板文件
        FileInputStream inputStream = new FileInputStream(filePath);
        Map params = new HashMap<>();
        params.put("userNameP", "张三");
        params.put("phoneP", "123456");
        JasperPrint jasperPrint = JasperFillManager.fillReport(inputStream, params, new JREmptyDataSource());
        JasperExportManager.exportReportToPdfStream(jasperPrint, new FileOutputStream("D:\\test01.pdf"));

    }
}
