package com.itheima.service;


import cn.afterturn.easypoi.csv.CsvExportUtil;
import cn.afterturn.easypoi.csv.entity.CsvExportParams;
import cn.afterturn.easypoi.entity.ImageEntity;
import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.ExcelImportUtil;
import cn.afterturn.easypoi.excel.entity.ExportParams;
import cn.afterturn.easypoi.excel.entity.ImportParams;
import cn.afterturn.easypoi.excel.entity.TemplateExportParams;
import cn.afterturn.easypoi.excel.entity.enmus.ExcelType;
import com.github.pagehelper.Page;
import com.github.pagehelper.PageHelper;
import com.itheima.mapper.ResourceMapper;
import com.itheima.mapper.UserMapper;
import com.itheima.pojo.Resource;
import com.itheima.pojo.User;
import com.itheima.utils.EntityUtils;
import com.itheima.utils.ExcelExportEngine;
import com.opencsv.CSVWriter;
import com.zaxxer.hikari.HikariDataSource;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import lombok.extern.slf4j.Slf4j;
import net.sf.jasperreports.engine.JasperExportManager;
import net.sf.jasperreports.engine.JasperFillManager;
import net.sf.jasperreports.engine.JasperPrint;
import net.sf.jasperreports.engine.data.JRBeanCollectionDataSource;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;
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
import java.io.FileInputStream;
import java.io.OutputStreamWriter;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

//import jxl.Workbook;
//import org.apache.poi.ss.usermodel.Workbook;

@Service
@Slf4j
public class UserService {

    @Autowired
    private UserMapper userMapper;

    @Autowired
    private ResourceMapper resourceMapper;

    @Autowired
    private HikariDataSource hikariDataSource;

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

    /**
     * 百万数据导出
     * 1.肯定使用的是高版本的Excel
     * 2.使用sax方式解析Excel(xml)
     * 3.限制:1.不能使用模板,2.不能使用太多的样式
     *
     * @param response
     */
    public void downLoadMillion(HttpServletResponse response) throws Exception {
//        指定使用的是sax方式解析
        SXSSFWorkbook workbook = new SXSSFWorkbook();  //sax方式就是逐行解析
//        Workbook workbook = new XSSFWorkbook(); //dom4j的方式
//        导出500W条数据 不可能放到同一个sheet中 规定：每个sheet不能超过100W条数据
        int page = 1;
        int num = 0;// 记录了处理数据的个数
        int rowIndex = 1; //记录的是每个sheet的行索引
        Row row = null;
        Sheet sheet = null;
        while (true) {
            List<User> userList = this.findPage(page, 100000);
            if (CollectionUtils.isEmpty(userList)) {
                break; //用户数据为空 跳出循环
            }
//           0   1000000  2000000  3000000  4000000  5000000
            if (num % 1000000 == 0) {  //表示应该创建新的标题
                sheet = workbook.createSheet("第" + ((num / 1000000) + 1) + "个工作表");
                rowIndex = 1; //每个sheet中的行索引重置
//            设置小标题
//            编号	姓名	手机号	入职日期	现住址
                String[] titles = new String[]{"编号", "姓名", "手机号", "入职日期", "现住址"};
                Row titleRow = sheet.createRow(0);
                for (int i = 0; i < 5; i++) {
                    titleRow.createCell(i).setCellValue(titles[i]);
                }
            }
            for (User user : userList) {
                row = sheet.createRow(rowIndex);
                row.createCell(0).setCellValue(user.getId());
                row.createCell(1).setCellValue(user.getUserName());
                row.createCell(2).setCellValue(user.getPhone());
                row.createCell(3).setCellValue(simpleDateFormat.format(user.getHireDate()));
                row.createCell(4).setCellValue(user.getAddress());

                rowIndex++;
                num++;
            }
            page++; //当前页码加1
        }

        String filename = "百万用户数据的导出.xlsx";
        response.setHeader("content-disposition", "attachment;filename=" + new String(filename.getBytes(), "ISO8859-1"));
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        workbook.write(response.getOutputStream());
    }

    public User findById(Long id) {
        //根据id查询用户，并且用户中附带公共用品数据
        User user = userMapper.selectByPrimaryKey(id);
        //再查询办公用品数据
        Resource resource = new Resource();
        resource.setUserId(id);
        List<Resource> resourceList = resourceMapper.select(resource);
        user.setResourceList(resourceList);
        return user;
    }

    //下载用户的合同文档
    public void downloadContract(Long id, HttpServletResponse response) throws Exception {
        //1、读取到模板
        File rootFile = new File(ResourceUtils.getURL("classpath:").getPath()); //获取项目的根目录
        File templateFile = new File(rootFile, "/word_template/contract_template.docx");
        XWPFDocument word = new XWPFDocument(new FileInputStream(templateFile));
        //2、查询当前用户User--->map
        User user = this.findById(id);
        Map<String, String> params = new HashMap<>();
        params.put("userName", user.getUserName());
        params.put("hireDate", simpleDateFormat.format(user.getHireDate()));
        params.put("address", user.getAddress());
        //3、替换数据
        //处理正文开始
        List<XWPFParagraph> paragraphs = word.getParagraphs();
        for (XWPFParagraph paragraph : paragraphs) {
            List<XWPFRun> runs = paragraph.getRuns();
            for (XWPFRun run : runs) {
                String text = run.getText(0);
                for (String key : params.keySet()) {
                    if (text.contains(key)) {
                        run.setText(text.replaceAll(key, params.get(key)), 0);
                    }
                }
            }
        }
//         处理正文结束

//      处理表格开始     名称	价值	是否需要归还	照片
        List<Resource> resourceList = user.getResourceList(); //表格中需要的数据
        XWPFTable xwpfTable = word.getTables().get(0);

        XWPFTableRow row = xwpfTable.getRow(0);
        int rowIndex = 1;
        for (Resource resource : resourceList) {
            //        添加行
//            xwpfTable.addRow(row);
            copyRow(xwpfTable, row, rowIndex);
            XWPFTableRow row1 = xwpfTable.getRow(rowIndex);
            row1.getCell(0).setText(resource.getName());
            row1.getCell(1).setText(resource.getPrice().toString());
            row1.getCell(2).setText(resource.getNeedReturn() ? "需求" : "不需要");

            File imageFile = new File(rootFile, "/static" + resource.getPhoto());
            setCellImage(row1.getCell(3), imageFile);
            rowIndex++;
        }
//     处理表格开始结束
//        4、导出word
        String filename = "员工(" + user.getUserName() + ")合同.docx";
        response.setHeader("content-disposition", "attachment;filename=" + new String(filename.getBytes(), "ISO8859-1"));
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        word.write(response.getOutputStream());
    }

    //    向单元格中写入图片
    private void setCellImage(XWPFTableCell cell, File imageFile) {

        XWPFRun run = cell.getParagraphs().get(0).createRun();
//        InputStream pictureData, int pictureType, String filename, int width, int height
        try (FileInputStream inputStream = new FileInputStream(imageFile)) {
            run.addPicture(inputStream, XWPFDocument.PICTURE_TYPE_JPEG, imageFile.getName(), Units.toEMU(100), Units.toEMU(50));
        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    //    用于深克隆行
    private void copyRow(XWPFTable xwpfTable, XWPFTableRow sourceRow, int rowIndex) {
        XWPFTableRow targetRow = xwpfTable.insertNewTableRow(rowIndex);
        targetRow.getCtRow().setTrPr(sourceRow.getCtRow().getTrPr());
//        获取源行的单元格
        List<XWPFTableCell> cells = sourceRow.getTableCells();
        if (CollectionUtils.isEmpty(cells)) {
            return;
        }
        XWPFTableCell targetCell = null;
        for (XWPFTableCell cell : cells) {
            targetCell = targetRow.addNewTableCell();
//            附上单元格的样式
//            单元格的属性
            targetCell.getCTTc().setTcPr(cell.getCTTc().getTcPr());
            targetCell.getParagraphs().get(0).getCTP().setPPr(cell.getParagraphs().get(0).getCTP().getPPr());
        }
    }

    /**
     * 使用EasyPOI方式导出Excel
     *
     * @param response
     */
    public void downLoadWithEasyPOI(HttpServletResponse response) throws Exception {
        ExportParams exportParams = new ExportParams("员工信息列表", "数据", ExcelType.XSSF);
        List<User> userList = userMapper.selectAll();
        org.apache.poi.ss.usermodel.Workbook workbook = ExcelExportUtil.exportExcel(exportParams, User.class, userList);
        String filename = "用户数据的导出.xlsx";
        response.setHeader("content-disposition", "attachment;filename=" + new String(filename.getBytes(), "ISO8859-1"));
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        workbook.write(response.getOutputStream());
    }

    public void uploadExcelWithEasyPOI(MultipartFile file) throws Exception {
        ImportParams importParams = new ImportParams();
        importParams.setNeedSave(false);
        importParams.setTitleRows(1);
        importParams.setHeadRows(1);
        List<User> userList = ExcelImportUtil.importExcel(file.getInputStream(), User.class, importParams);
        for (User user : userList) {
            user.setId(null);
            userMapper.insert(user);
        }

    }

    public void downloadUserInfoByEasyPOI(Long id, HttpServletResponse response) throws Exception {
        File rootFile = new File(ResourceUtils.getURL("classpath:").getPath()); //获取项目的根目录
        File templateFile = new File(rootFile, "/excel_template/userInfo3.xlsx");
//        Workbook workbook = new XSSFWorkbook(templateFile);
        TemplateExportParams exportParams = new TemplateExportParams(templateFile.getPath(), true);

        User user = userMapper.selectByPrimaryKey(id);
        Map<String, Object> map = EntityUtils.entityToMap(user);
        ImageEntity imageEntity = new ImageEntity();
        imageEntity.setUrl(user.getPhoto());
        imageEntity.setColspan(2); //占用多少列
        imageEntity.setRowspan(4); //占用多少行

        map.put("photo", imageEntity);
        org.apache.poi.ss.usermodel.Workbook workbook = ExcelExportUtil.exportExcel(exportParams, map);

        String filename = "用户数据.xlsx";
        response.setHeader("content-disposition", "attachment;filename=" + new String(filename.getBytes(), "ISO8859-1"));
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        workbook.write(response.getOutputStream());
    }

    public void downLoadCSVWithEasyPOI(HttpServletResponse response) throws Exception {
        ServletOutputStream outputStream = response.getOutputStream();
        String filename = "百万用户数据的导出.csv";
        response.setHeader("content-disposition", "attachment;filename=" + new String(filename.getBytes(), "ISO8859-1"));
        response.setContentType("text/csv");

        CsvExportParams csvExportParams = new CsvExportParams();
        csvExportParams.setExclusions(new String[]{"照片"});

        List<User> userList = userMapper.selectAll();
        CsvExportUtil.exportCsv(csvExportParams, User.class, userList, outputStream);

    }

    public void downLoadPDF(HttpServletResponse response) throws Exception {
        //1.获取模本文件
        File rootFile = new File(ResourceUtils.getURL("classpath:").getPath()); //获取项目的根目录
        File templateFile = new File(rootFile, "/pdf_template/userList_db.jasper");
        //2.准备数据库的连接
        FileInputStream inputStream = new FileInputStream(templateFile);
        Map params = new HashMap();
        //JasperPrint jasperPrint = JasperFillManager.fillReport(inputStream, params, getCon());
        JasperPrint jasperPrint = JasperFillManager.fillReport(inputStream, params, hikariDataSource.getConnection());
        ServletOutputStream outputStream = response.getOutputStream();
        String filename = "用户列表数据.pdf";
        response.setHeader("content-disposition", "attachment;filename=" + new String(filename.getBytes(), "ISO8859-1"));
        response.setContentType("application/pdf");
        JasperExportManager.exportReportToPdfStream(jasperPrint, outputStream);
    }

    public void downLoadPDF2(HttpServletResponse response) throws Exception {
        //1.获取模本文件
        File rootFile = new File(ResourceUtils.getURL("classpath:").getPath()); //获取项目的根目录
        File templateFile = new File(rootFile, "/pdf_template/userList.jasper");
        //2.准备显示数据
        FileInputStream inputStream = new FileInputStream(templateFile);
        List<User> userList = userMapper.selectAll();

        userList = userList.stream().map((user) -> {
            Date hireDate = user.getHireDate();
            String formatHireDate = simpleDateFormat.format(hireDate);
            user.setHireDateStr(formatHireDate);
            return user;
        }).collect(Collectors.toList());

        JRBeanCollectionDataSource jrBeanCollectionDataSource = new JRBeanCollectionDataSource(userList);
        Map params = new HashMap();
        JasperPrint jasperPrint = JasperFillManager.fillReport(inputStream, params, jrBeanCollectionDataSource);
        ServletOutputStream outputStream = response.getOutputStream();
        String filename = "用户列表数据.pdf";
        response.setHeader("content-disposition", "attachment;filename=" + new String(filename.getBytes(), "ISO8859-1"));
        response.setContentType("application/pdf");
        JasperExportManager.exportReportToPdfStream(jasperPrint, outputStream);
    }
/*
    private Connection getCon() throws Exception {
        Class.forName("com.mysql.jdbc.Driver");
        Connection connection = DriverManager.getConnection("jdbc:mysql://localhost:13306/report_manager_db", "root", "root");
        return connection;
    }
*/
}
