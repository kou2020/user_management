package com.itheima.service;


import com.github.pagehelper.Page;
import com.github.pagehelper.PageHelper;
import com.itheima.mapper.UserMapper;
import com.itheima.pojo.User;
import com.itheima.utils.ExcelExportEngine;
import com.opencsv.CSVWriter;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.util.CollectionUtils;
import org.springframework.util.ResourceUtils;
import org.springframework.web.multipart.MultipartFile;

import javax.imageio.ImageIO;
import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.OutputStreamWriter;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

//import jxl.Workbook;
//import org.apache.poi.ss.usermodel.Workbook;

@Service
@Slf4j
public class UserService {

    @Autowired
    private UserMapper userMapper;

    private SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");

    public List<User> findAll() {
        return userMapper.selectAll();
    }

    public List<User> findPage(Integer page, Integer pageSize) {
        PageHelper.startPage(page, pageSize);  //开启分页
        Page<User> userPage = (Page<User>) userMapper.selectAll(); //实现查询
        return userPage.getResult();
    }

    public void downLoadXlsByJxl(HttpServletResponse response) throws Exception {
        //编号姓名手机号入职日期现住址
        //创建了一个全新的工作簿
        ServletOutputStream outputStream = response.getOutputStream();
        WritableWorkbook workbook = Workbook.createWorkbook(outputStream);
        //创建一个工作表
        WritableSheet sheet = workbook.createSheet("一个JXL入门", 0);
        //设置列宽
        sheet.setColumnView(0, 5); //  第一个参数：列的索引值  第二个参数：1代表一个标准字母的宽度
        sheet.setColumnView(1, 8); //  第一个参数：列的索引值  第二个参数：1代表一个标准字母的宽度
        sheet.setColumnView(2, 15); //  第一个参数：列的索引值  第二个参数：1代表一个标准字母的宽度
        sheet.setColumnView(3, 15); //  第一个参数：列的索引值  第二个参数：1代表一个标准字母的宽度
        sheet.setColumnView(4, 30); //  第一个参数：列的索引值  第二个参数：1代表一个标准字母的宽度
        //处理标题
        String[] titles = new String[]{"编号", "姓名", "手机号", "入职日期", "现住址"};
        Label label = null;
        for (int i = 0; i < titles.length; i++) {
            label = new Label(i, 0, titles[i]); //列脚标, 行脚标, 单元格中的内容
            sheet.addCell(label);
        }
        //查询所有用户数据
        List<User> userList = userMapper.selectAll();
        int rowIndex = 1;
        for (User user : userList) {
            label = new Label(0, rowIndex, user.getId().toString());
            sheet.addCell(label);
            label = new Label(1, rowIndex, user.getUserName());
            sheet.addCell(label);
            label = new Label(2, rowIndex, user.getPhone());
            sheet.addCell(label);
            label = new Label(3, rowIndex, simpleDateFormat.format(user.getHireDate()));
            sheet.addCell(label);
            label = new Label(4, rowIndex, user.getAddress());
            sheet.addCell(label);
            rowIndex++;
        }
        //文件导出一个流(outputStream)两个头(文件的打开方式 in-line attachment,文件下载时mime类型 application/vnd.ms-excel)
        String filename = "一个JXL入门.xls";
        response.setHeader("content-disposition", "attachment;filename=" + new String(filename.getBytes(), "ISO8859-1"));
        response.setContentType("application/vnd.ms-excel");
        // 写入数据
        workbook.write();

        //关闭资源
        workbook.close();
        outputStream.close();

    }

    public void uploadExcel(MultipartFile file) throws Exception {
        XSSFWorkbook workbook = new XSSFWorkbook(file.getInputStream());
        //获取到第一个工作表
        Sheet sheet = workbook.getSheetAt(0);
        //获取当前sheet的最后一行的角标
        int lastRowIndex = sheet.getLastRowNum();
        //读取工作表中的内容
        Row row = null;
        User user = null;
        for (int i = 1; i <= lastRowIndex; i++) {
            row = sheet.getRow(i);
            String userName = row.getCell(0).getStringCellValue();
            String phone = null; //手机号
            try {
                phone = row.getCell(1).getStringCellValue();
            } catch (Exception e) {
                phone = row.getCell(1).getNumericCellValue() + "";
            }
            String province = row.getCell(2).getStringCellValue(); //省份
            String city = row.getCell(3).getStringCellValue(); //城市
            Integer salary = ((Double) row.getCell(4).getNumericCellValue()).intValue(); //工资
            Date hireDate = simpleDateFormat.parse(row.getCell(5).getStringCellValue()); //入职日期
            Date birthDay = simpleDateFormat.parse(row.getCell(6).getStringCellValue()); //出生日期
            String address = row.getCell(7).getStringCellValue(); //现住地址
            user = new User();
            user.setUserName(userName);
            user.setPhone(phone);
            user.setProvince(province);
            user.setCity(city);
            user.setSalary(salary);
            user.setHireDate(hireDate);
            user.setBirthday(birthDay);
            user.setAddress(address);
            userMapper.insert(user);
        }
    }

    public void downLoadXlsxByPoi(HttpServletResponse response) throws Exception {
        ServletOutputStream outputStream = response.getOutputStream();
        //1.创建一个全新的工作簿
        XSSFWorkbook workbook = new XSSFWorkbook();
        //2.创建一个全新的工作表
        Sheet sheet = workbook.createSheet("用户数据");
        //设置列宽
        sheet.setColumnWidth(0, 5 * 256);
        sheet.setColumnWidth(1, 8 * 256);
        sheet.setColumnWidth(2, 15 * 256);
        sheet.setColumnWidth(3, 15 * 256);
        sheet.setColumnWidth(4, 30 * 256);
        //3.处理标题
        String[] titles = new String[]{"编号", "姓名", "手机号", "入职日期", "现住址"};
        Row titleRow = sheet.createRow(0);
        for (int i = 0; i < titles.length; i++) {
            Cell cell = titleRow.createCell(i);
            cell.setCellValue(titles[i]);
        }
        //4.从第二行开始遍历单元格内容数据
        List<User> userList = userMapper.selectAll();
        int rowIndex = 1;
        Row row = null;
        Cell cell = null;
        for (User user : userList) {
            row = sheet.createRow(rowIndex);
            cell = row.createCell(0);
            cell.setCellValue(user.getId());
            cell = row.createCell(1);
            cell.setCellValue(user.getUserName());
            cell = row.createCell(2);
            cell.setCellValue(user.getPhone());
            cell = row.createCell(3);
            cell.setCellValue(simpleDateFormat.format(user.getHireDate()));
            cell = row.createCell(4);
            cell.setCellValue(user.getAddress());
            rowIndex++;
        }
        //一个流两个头
        String filename = "员工数据.xlsx";
        response.setHeader("content-disposition", "attachment;filename=" + new String(filename.getBytes(), "ISO8859-1"));
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        workbook.write(outputStream);
        //关闭资源
        workbook.close();
        outputStream.close();

    }

    public void downLoadCSV(HttpServletResponse response) throws Exception {
        ServletOutputStream outputStream = response.getOutputStream();
        //一个流两个头
        String filename = "百万用户数据导出.csv";
        response.setHeader("content-disposition", "attachment;filename=" + new String(filename.getBytes(), "ISO8859-1"));
        response.setContentType("text/csv");
        CSVWriter csvWriter = new CSVWriter(new OutputStreamWriter(outputStream, "utf-8"));
        String[] titles = new String[]{"编号", "姓名", "手机号", "入职日期", "现住址"};
        // 写入表头
        csvWriter.writeNext(titles);
        int page = 1;
        while (true) {
            List<User> userList = this.findPage(page, 20000);
            if (CollectionUtils.isEmpty(userList)) {
                break;
            }
            for (User user : userList) {
                csvWriter.writeNext(new String[]{
                        user.getId().toString(),
                        user.getUserName(),
                        user.getPhone(),
                        simpleDateFormat.format(user.getHireDate()),
                        user.getAddress()
                });
            }
            page++;
            csvWriter.flush();
        }
        csvWriter.close();
        outputStream.close();
    }

    public void downLoadXlsxByPoiWithCellStyle(HttpServletResponse response) throws Exception {
        ServletOutputStream outputStream = response.getOutputStream();
        XSSFWorkbook workbook = new XSSFWorkbook();
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

        List<User> userList = userMapper.selectAll();
        int rowIndex = 2;
        Row row = null;
        Cell cell = null;
        for (User user : userList) {
            row = sheet.createRow(rowIndex);
            cell = row.createCell(0);
            cell.setCellStyle(contentRowCellStyle);
            cell.setCellValue(user.getId());

            cell = row.createCell(1);
            cell.setCellStyle(contentRowCellStyle);
            cell.setCellValue(user.getUserName());

            cell = row.createCell(2);
            cell.setCellStyle(contentRowCellStyle);
            cell.setCellValue(user.getPhone());

            cell = row.createCell(3);
            cell.setCellStyle(contentRowCellStyle);
            cell.setCellValue(simpleDateFormat.format(user.getHireDate()));

            cell = row.createCell(4);
            cell.setCellStyle(contentRowCellStyle);
            cell.setCellValue(user.getAddress());

            rowIndex++;
        }

        //一个流两个头
        String filename = "员工数据.xlsx";
        response.setHeader("content-disposition", "attachment;filename=" + new String(filename.getBytes(), "ISO8859-1"));
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        workbook.write(outputStream);
        //关闭资源
        workbook.close();
        outputStream.close();

    }

    public void downLoadXlsxByPoiWithTemplate(HttpServletResponse response) throws Exception {
        ServletOutputStream outputStream = response.getOutputStream();
        //1.获取模板
        File rootFile = new File(ResourceUtils.getURL("classpath:").getPath());//获取项目的根目录
        File templateFile = new File(rootFile, "/excel_template/userList.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(templateFile);
        //2.查询所有用户数据
        List<User> userList = userMapper.selectAll();
        //3.将数据放入模板中
        XSSFSheet sheet = workbook.getSheetAt(0);
        //获取准备好的单元格样式 在第二个sheet的第一个单元格中
        CellStyle contentRowCellStyle = workbook.getSheetAt(1).getRow(0).getCell(0).getCellStyle();
        int rowIndex = 2;
        Row row = null;
        Cell cell = null;
        for (User user : userList) {
            row = sheet.createRow(rowIndex);
            //行高
            row.setHeightInPoints(15);
            cell = row.createCell(0);
            cell.setCellStyle(contentRowCellStyle);
            cell.setCellValue(user.getId());

            cell = row.createCell(1);
            cell.setCellStyle(contentRowCellStyle);
            cell.setCellValue(user.getUserName());

            cell = row.createCell(2);
            cell.setCellStyle(contentRowCellStyle);
            cell.setCellValue(user.getPhone());

            cell = row.createCell(3);
            cell.setCellStyle(contentRowCellStyle);
            cell.setCellValue(simpleDateFormat.format(user.getHireDate()));

            cell = row.createCell(4);
            cell.setCellStyle(contentRowCellStyle);
            cell.setCellValue(user.getAddress());
            rowIndex++;
        }
        //删除第二个样式没用的sheet
        workbook.removeSheetAt(1);
        //4.导出文件
        //一个流两个头
        String filename = "员工数据.xlsx";
        response.setHeader("content-disposition", "attachment;filename=" + new String(filename.getBytes(), "ISO8859-1"));
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        workbook.write(outputStream);
        //关闭资源
        workbook.close();
        outputStream.close();
    }

    /**
     * 使用模板导出用户的详细信息
     *
     * @param id
     * @param response
     */
    public void downloadUserInfoByTemplate(Long id, HttpServletResponse response) throws Exception {
        ServletOutputStream outputStream = response.getOutputStream();
        //1.读取模板
        File rootFile = new File(ResourceUtils.getURL("classpath:").getPath());
        File templateFile = new File(rootFile, "/excel_template/userInfo.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(templateFile);
        Sheet sheet = workbook.getSheetAt(0);
        //2.根据ID获取某一个用户数据
        User user = userMapper.selectByPrimaryKey(id);
        //3.把用户数据放入模板
        //用户名  第2行第2列
        sheet.getRow(1).getCell(1).setCellValue(user.getUserName());
        //手机号  第3行第2列
        sheet.getRow(2).getCell(1).setCellValue(user.getPhone());
        //生日   第4行第2列
        sheet.getRow(3).getCell(1).setCellValue(simpleDateFormat.format(user.getBirthday()));
        //工资   第5行第2列
        sheet.getRow(4).getCell(1).setCellValue(user.getSalary());
        //入职日期  第6行第2列
        sheet.getRow(5).getCell(1).setCellValue(simpleDateFormat.format(user.getHireDate()));
        //省份   第7行第2列
        sheet.getRow(6).getCell(1).setCellValue(user.getProvince());
        //现住址  第8行第2列
        sheet.getRow(7).getCell(1).setCellValue(user.getAddress());
        //司龄  第6行第4列 使用公式稍后处理 =CONCATENATE(DATEDIF(B6,TODAY(),"Y"),"年",DATEDIF(B6,TODAY(),"YM"),"个月")
        sheet.getRow(5).getCell(3).setCellFormula("CONCATENATE(DATEDIF(B6,TODAY(),\"Y\"),\"年\",DATEDIF(B6,TODAY(),\"YM\"),\"个月\")");
        //城市  第7行第4列
        sheet.getRow(6).getCell(3).setCellValue(user.getCity());
        // 照片的位置
        //开始处理照片
        //先创建一个字节输出流
        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
        //读取图片 放入了一个带有缓存区的图片类中
        BufferedImage bufferedImage = ImageIO.read(new File(rootFile, user.getPhoto()));
        //把图片写入到了字节输出流中
        String extName = user.getPhoto().substring(user.getPhoto().lastIndexOf(".") + 1).toUpperCase();
        ImageIO.write(bufferedImage, extName, byteArrayOutputStream);
        //Patriarch 控制图片的写入 和ClientAnchor 指定图片的位置
        Drawing patriarch = sheet.createDrawingPatriarch();
        //指定图片的位置         开始列3 开始行2   结束列4  结束行5
        //偏移的单位：是一个英式公制的单位  1厘米=360000
        ClientAnchor anchor = new XSSFClientAnchor(0, 0, 0, 0, 2, 1, 4, 5);
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
        //4.导出文件
        //一个流两个头
        String filename = "员工数据(" + user.getUserName() + ").xlsx";
        response.setHeader("content-disposition", "attachment;filename=" + new String(filename.getBytes(), "ISO8859-1"));
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        workbook.write(outputStream);
        //关闭资源
        workbook.close();
        outputStream.close();
    }

    public void downloadUserInfoByTemplate2(Long id, HttpServletResponse response) throws Exception {
        ServletOutputStream outputStream = response.getOutputStream();
        //1.读取模板
        File rootFile = new File(ResourceUtils.getURL("classpath:").getPath());
        File templateFile = new File(rootFile, "/excel_template/userInfo2.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(templateFile);
        Sheet sheet = workbook.getSheetAt(0);
        //2.根据ID获取某一个用户数据
        User user = userMapper.selectByPrimaryKey(id);
        //3.通过自定义的引擎放数据
        workbook = (XSSFWorkbook) ExcelExportEngine.writeToExcel(user, workbook, rootFile.getPath() + user.getPhoto());
        //4.导出文件
        //一个流两个头
        String filename = "员工数据(" + user.getUserName() + ").xlsx";
        response.setHeader("content-disposition", "attachment;filename=" + new String(filename.getBytes(), "ISO8859-1"));
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        workbook.write(outputStream);
        //关闭资源
        workbook.close();
        outputStream.close();
    }
}
