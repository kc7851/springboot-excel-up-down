package com.skc.excel.controller;

import ch.qos.logback.core.util.FileUtil;
import com.skc.excel.service.ExcelService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.multipart.MultipartHttpServletRequest;
import org.springframework.web.servlet.ModelAndView;

import java.io.File;
import java.io.IOException;
import java.util.Objects;

@Controller
public class IndexController {

    @Autowired
    private ExcelService excelService;

    @GetMapping("/")
    public ModelAndView index() {
        ModelAndView modelAndView = new ModelAndView("index");

        return modelAndView;
    }

    @PostMapping("/excelUpload")
    public ModelAndView excelUploadAjax(@RequestParam MultipartFile excelFile)  throws Exception{
        ModelAndView modelAndView = new ModelAndView("index");

        String extension = Objects.requireNonNull(
                excelFile.getOriginalFilename()).substring(
                        excelFile.getOriginalFilename().lastIndexOf(".") + 1);

        if(excelFile==null || excelFile.isEmpty() || !extension.contains("xls")){
            throw new RuntimeException("엑셀파일을 선택 해 주세요.");
        }

        String userHomePath = System.getProperty("java.io.tmpdir");

        File destFile = new File(userHomePath+excelFile.getOriginalFilename());

        try{
            excelFile.transferTo(destFile);
        }catch(IllegalStateException | IOException e){
            throw new RuntimeException(e.getMessage(),e);
        }

        excelService.excelUpload(destFile);

        return modelAndView;
    }

}
