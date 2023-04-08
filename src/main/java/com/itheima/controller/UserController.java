package com.itheima.controller;

import com.itheima.pojo.User;
import com.itheima.service.UserService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.util.List;

@RestController
@RequestMapping("/user")
public class UserController {

    @Autowired
    private UserService userService;

    @GetMapping("/findPage")
    public List<User> findPage(
            @RequestParam(value = "page", defaultValue = "1") Integer page,
            @RequestParam(value = "rows", defaultValue = "10") Integer pageSize) {
        return userService.findPage(page, pageSize);
    }

    @GetMapping(value = "/downLoadXlsByJxl", name = "使用jxl导出excel")
    public void downLoadXlsByJxl(HttpServletResponse response) throws Exception {
        userService.downLoadXlsByJxl(response);
    }

    @PostMapping(value = "/uploadExcel", name = "上传用户文件")
    public void uploadExcel(MultipartFile file) throws Exception {
        //userService.uploadExcel(file);
        userService.uploadExcelWithEasyPOI(file);
    }

    @GetMapping(value = "/downLoadXlsxByPoi", name = "使用poi导出excel")
    public void downLoadXlsxByPoi(HttpServletResponse response) throws Exception {
        //userService.downLoadXlsxByPoi(response);
        //userService.downLoadXlsxByPoiWithCellStyle(response);
        userService.downLoadXlsxByPoiWithTemplate(response);
    }

    @GetMapping(value = "/download", name = "使用poi导出用户详细数据")
    public void downloadUserInfoByTemplate(Long id, HttpServletResponse response) throws Exception {
        //userService.downloadUserInfoByTemplate(id,response);
        //userService.downloadUserInfoByTemplate2(id, response);
        userService.downloadUserInfoByEasyPOI(id,response);
    }

    @GetMapping(value = "/downLoadMillion", name = "导出百万数据")
    public void downLoadMillion(Long id, HttpServletResponse response) throws Exception {
        userService.downLoadMillion(response);
    }

    @GetMapping(value = "/downLoadCSV", name = "使用CSV文件导出百万数据")
    public void downLoadCSV(HttpServletResponse response) throws Exception {
        //userService.downLoadCSV(response);
        userService.downLoadCSVWithEasyPOI(response);
    }

    @GetMapping(value = "/{id}", name = "根据id查询用户数据")
    public User findById(@PathVariable("id") Long id) throws Exception {
        return userService.findById(id);
    }

    @GetMapping(value = "/downloadContract", name = "下载用户的合同文档")
    public void downloadContract(Long id, HttpServletResponse response) throws Exception {
        userService.downloadContract(id, response);
    }

    @GetMapping(value = "/downLoadWithEasyPOI", name = "使用EasyPOI方式导出excel")
    public void downLoadWithEasyPOI(HttpServletResponse response) throws Exception {
        userService.downLoadWithEasyPOI(response);
    }

    @GetMapping(value = "/downLoadPDF", name = "导出用户数据到PDF")
    public void downLoadPDF(HttpServletResponse response) throws Exception {
        //userService.downLoadPDF(response);
        userService.downLoadPDF2(response);
    }
}
