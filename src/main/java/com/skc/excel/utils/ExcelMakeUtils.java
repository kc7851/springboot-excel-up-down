package com.skc.excel.utils;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;
import java.util.HashMap;
import java.util.List;

public class ExcelMakeUtils {
    private static final String CONTENT_TYPE_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

    public static Workbook buildExcelDocument(List<HashMap<String, String>> model, String fileName, String[] header, HttpServletRequest request, HttpServletResponse response) throws Exception {

        try {
            String browser = request.getHeader("User-Agent");
            //파일 인코딩
            if (browser.contains("MSIE") || browser.contains("Trident")
                    || browser.contains("Chrome")) {
                fileName = URLEncoder.encode(fileName, "UTF-8").replaceAll("\\+",
                        "%20");
            } else {
                fileName = new String(fileName.getBytes("UTF-8"), "ISO-8859-1");
            }
        } catch (UnsupportedEncodingException ex) {
            System.out.println("UnsupportedEncodingException");
        }
        response.setCharacterEncoding("UTF-8");
        response.setContentType("application/vnd.ms-excel");
        response.setHeader("Pragma", "public");
        response.setHeader("Expires", "0");
        response.setHeader("Content-Disposition", "attachment; filename = " + fileName + System.currentTimeMillis() + ".xlsx");

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet();

        Row row = null;
        int rowCount = 0;
        int cellCount = 0;

        row = sheet.createRow(rowCount++);

        for (int i = 0; i < header.length; i++)
            row.createCell(cellCount++).setCellValue(header[i]);

        for (HashMap<String, String> vo : model) {
            row = sheet.createRow(rowCount++);
            cellCount = 0;
            for (int idx = 0; idx < vo.size(); idx++) {
                String value = vo.get(header[idx]);
                if (value == null) value = "";

                if(isNumeric(value)) {
                    row.createCell(cellCount++, CellType.NUMERIC).setCellValue(Double.parseDouble(value));
                }else {
                    row.createCell(cellCount++).setCellValue(value);
                }
            }
        }

        ServletOutputStream out = response.getOutputStream();
        out.flush();
        workbook.write(out);
        out.flush();

        if (workbook instanceof SXSSFWorkbook) {
            ((SXSSFWorkbook) workbook).dispose();
        }

        return workbook;
    }

    private static boolean isNumeric(String s) {
        try {
            Double.parseDouble(s);
            return true;
        } catch (NumberFormatException e) {
            return false;
        }
    }
}
