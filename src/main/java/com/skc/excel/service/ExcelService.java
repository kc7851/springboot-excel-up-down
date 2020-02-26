package com.skc.excel.service;

import com.skc.excel.utils.ExcelRead;
import com.skc.excel.utils.ExcelReadOption;
import org.springframework.stereotype.Service;

import java.io.File;
import java.util.List;
import java.util.Map;

@Service
public class ExcelService {
    public void excelUpload(File destFile) {
        ExcelReadOption excelReadOption = new ExcelReadOption();
        excelReadOption.setFilePath(destFile.getAbsolutePath());
        excelReadOption.setOutputColumns("A","B","C","D","E","F");
        excelReadOption.setStartRow(2);

        List<Map<String, String>> excelContent = ExcelRead.read(excelReadOption);

        for(Map<String, String> article: excelContent){
            System.out.println(article.get("A"));
            System.out.println(article.get("B"));
            System.out.println(article.get("C"));
            System.out.println(article.get("D"));
            System.out.println(article.get("E"));
            System.out.println(article.get("F"));
        }

    }
}
