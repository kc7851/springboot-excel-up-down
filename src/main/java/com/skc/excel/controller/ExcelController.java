package com.skc.excel.controller;

import com.skc.excel.service.ExcelService;
import com.skc.excel.utils.ExcelMakeUtils;
import com.sun.org.apache.xpath.internal.operations.Mod;
import io.cogi.spring.constant.MediaType;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Objects;

@Controller
public class ExcelController {

    @Autowired
    private ExcelService excelService;

    @GetMapping("/")
    public ModelAndView index() {
        ModelAndView modelAndView = new ModelAndView("index");

        return modelAndView;
    }

    @GetMapping("/upload")
    public ModelAndView excelUpload() {
        ModelAndView modelAndView = new ModelAndView("upload");

        return modelAndView;
    }

    @GetMapping("/download")
    public ModelAndView excelDownload() {
        ModelAndView modelAndView = new ModelAndView("download");

        return modelAndView;
    }

    @PostMapping("/excelUpload")
    public String excelUploadPost(@RequestParam MultipartFile excelFile)  throws Exception{
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

        return "redirect:/upload";
    }

    @PostMapping(value = "/excelDownload")
    public void excelDownloadPost(
            HttpServletRequest request, HttpServletResponse response
    ) throws Exception {

        String[] strArrHeader = {"순번","회원 ID","회원명"};
        String filename = "엑셀파일이름"+ System.currentTimeMillis();

        //이 부분에 실제 엑셀에 넣을 데이터를 구성하면 됨.
        List<HashMap<String, String>> hashMapList = new ArrayList<>();
        HashMap<String, String> user1 = new HashMap<>();
        user1.put("순번", "1");
        user1.put("회원 ID", "id1");
        user1.put("회원명", "회원1");
        HashMap<String, String> user2 = new HashMap<>();
        user2.put("순번", "2");
        user2.put("회원 ID", null);
        user2.put("회원명", "회원2");
        hashMapList.add(user1);
        hashMapList.add(user2);

        ExcelMakeUtils.buildExcelDocument(hashMapList,filename,strArrHeader,request,response);

    }

    @PostMapping(value = "/excelDownloadView", produces = MediaType.APPLICATION_XLS_VALUE)
    public ModelAndView excelDownloadView(
            HttpServletRequest request, HttpServletResponse response
    ) throws Exception {

        String[] strArrHeader = {"순번","회원 ID","회원명"};
        String filename = "엑셀파일이름"+ System.currentTimeMillis();

        //이 부분에 실제 엑셀에 넣을 데이터를 구성하면 됨.
        List<HashMap<String, String>> hashMapList = new ArrayList<>();
        HashMap<String, String> user1 = new HashMap<>();
        user1.put("순번", "1");
        user1.put("id", "id1");
        user1.put("name", "회원1");
        HashMap<String, String> user2 = new HashMap<>();
        user2.put("순번", "2");
        user2.put("id", null);
        user2.put("name", "회원2");
        hashMapList.add(user1);
        hashMapList.add(user2);

        ModelAndView modelAndView = new ModelAndView("excelDownloadView한글");
        modelAndView.addObject("result", hashMapList);

        return modelAndView;
    }

}
