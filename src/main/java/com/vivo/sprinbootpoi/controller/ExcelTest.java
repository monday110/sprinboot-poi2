package com.vivo.sprinbootpoi.controller;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.List;

@Controller
@RequestMapping("test")
public class ExcelTest {
    //创建一个list集合存cellStyle
    List<HSSFCellStyle> cellstyles=new ArrayList<>();

    @ResponseBody
    @RequestMapping("getExcelTest")
    public void getExcelTest(HttpServletResponse response) throws Exception {
       HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sheet = wb.createSheet("测试表");
        HSSFRow row=null;
        HSSFCellStyle cellStyle = wb.createCellStyle();
        HSSFFont font = wb.createFont();
        //1.生成表格
        for (int i = 1; i < 6; i++) {
             row = sheet.createRow(i);
            for (int i1 = 1; i1 <= 2; i1++) {
                row.createCell(i1);
            }
        }
        //A4_J5
        for (int i = 3;i < 5; i++) {
            for (int i1 = 0; i1 <= 9; i1++) {
                sheet.getRow(i).getCell(i1).setCellStyle(getStyle(wb,0));
            }
        }
        //A6_G98
        for (int i = 5;i < 98; i++) {
            for (int i1 = 0; i1 <= 6; i1++) {
                sheet.getRow(i).getCell(i1).setCellStyle(getStyle(wb,0));
            }
        }
        //2.填充字体
        sheet.getRow(1).getCell(1).setCellValue("水平居中");
        sheet.getRow(1).getCell(1).setCellValue("水平跨列居中");
        //3.设置特定样式
        //B2_C5
        for (int i = 1; i < 6; i++) {
            for (int i1 = 1; i1 <= 2; i1++) {
               sheet.getRow(i).getCell(i1).setCellStyle(getStyle(wb,0));
            }
        }
        //4.返回
        response.setContentType("application/vnd.ms-excel;charset=utf-8");
        ServletOutputStream outputStream = response.getOutputStream();
        wb.write(outputStream);
        outputStream.flush();
        outputStream.close();

    }
    /**
     * 获取样式
     *
     * @param styleNum
     * @return
     */
    public HSSFCellStyle getStyle(HSSFWorkbook wb, Integer styleNum) {
        List<HSSFCellStyle> styles=new ArrayList<>();
        List<HSSFFont> fonts=new ArrayList<>();
        if(styles==null){
            for (int i = 0; i < 11; i++) {
                HSSFCellStyle cellStyle = wb.createCellStyle();
                HSSFFont cellFont = wb.createFont();
                styles.add(cellStyle);
                fonts.add(cellFont);

            }
        }
        HSSFCellStyle style=styles.get(styleNum);
        HSSFFont font=fonts.get(styleNum);
        switch (styleNum) {
            case (0): {//设置边框
                style.setBorderRight(BorderStyle.THIN);//右边框
                style.setBorderBottom(BorderStyle.THIN);//下边框
                style.setBorderLeft(BorderStyle.THIN);//左边框
                style.setBorderTop(BorderStyle.THIN);//上边框
            }
            break;
            case (1): {//方正小标宋_GBK 16 垂直居中,左对齐
                font.setFontName("方正小标宋_GBK");
                font.setFontHeightInPoints((short) 16);//字体大小
                style.setFont(font);
                style.setVerticalAlignment(VerticalAlignment.CENTER);//垂直居中
                style.setAlignment(HorizontalAlignment.LEFT);//跨列居中
                //font.setBold(true);//粗体
                style.setBorderTop(BorderStyle.THIN);
                style.setBorderRight(BorderStyle.THIN);
                style.setBorderLeft(BorderStyle.THIN);
                style.setBorderBottom(BorderStyle.THIN);

            }
            break;
            case (2): {//样式2：宋体16 垂直，水平居中 加粗
                 style = cellstyles.get(2);
                font.setFontName("宋体");
                style.setVerticalAlignment(VerticalAlignment.CENTER);//垂直居中
                style.setAlignment(HorizontalAlignment.CENTER_SELECTION);//跨列居中
                font.setBold(true);//粗体
                font.setFontHeightInPoints((short) 16);//字体大小
                style.setFont(font);

            }
            break;
            case (3): {//宋体12 加粗水平居左，垂直居中
                font.setFontName("宋体");
                font.setBold(true);//粗体
                font.setFontHeightInPoints((short) 12);//字体大小
                style.setFont(font);
                style.setVerticalAlignment(VerticalAlignment.CENTER);//垂直居中
                style.setAlignment(HorizontalAlignment.LEFT);


            }
            break;
            case (4): {//宋体10 加粗 垂直水平居中
                font.setFontName("宋体");
                style.setVerticalAlignment(VerticalAlignment.CENTER);//垂直居中
                style.setAlignment(HorizontalAlignment.CENTER_SELECTION);
                font.setBold(true);//粗体
                font.setFontHeightInPoints((short) 10);//字体大小
                style.setFont(font);
                style.setBorderTop(BorderStyle.THIN);
                style.setBorderRight(BorderStyle.THIN);
                style.setBorderLeft(BorderStyle.THIN);
                style.setBorderBottom(BorderStyle.THIN);
            }
            break;
            case(5):{//宋体11 加粗 垂直水平居中
                font.setFontName("宋体");
                style.setVerticalAlignment(VerticalAlignment.CENTER);//垂直居中
                style.setAlignment(HorizontalAlignment.CENTER_SELECTION);
                font.setBold(true);//粗体
                font.setFontHeightInPoints((short) 11);//字体大小
                style.setFont(font);
                style.setBorderTop(BorderStyle.THIN);
                style.setBorderRight(BorderStyle.THIN);
                style.setBorderLeft(BorderStyle.THIN);
                style.setBorderBottom(BorderStyle.THIN);
            }
            break;
            case(6):{//宋体11 垂直居中，水平居右
                font.setFontName("宋体");
                style.setVerticalAlignment(VerticalAlignment.CENTER);//垂直居中
                style.setAlignment(HorizontalAlignment.CENTER);//水平
                //font.setBold(true);//粗体
                font.setFontHeightInPoints((short) 11);//字体大小
                style.setFont(font);
                style.setBorderTop(BorderStyle.THIN);
                style.setBorderRight(BorderStyle.THIN);
                style.setBorderLeft(BorderStyle.THIN);
                style.setBorderBottom(BorderStyle.THIN);
            }
            break;
            case(7):{//宋体10 垂直水平居中，自动换行
                font.setFontName("宋体");
                style.setVerticalAlignment(VerticalAlignment.CENTER);//垂直居中
                style.setAlignment(HorizontalAlignment.CENTER_SELECTION);
                //font.setBold(true);//粗体
                font.setFontHeightInPoints((short) 10);//字体大小
                style.setFont(font);
                style.setWrapText(true);//自动换行
                style.setBorderTop(BorderStyle.THIN);
                style.setBorderRight(BorderStyle.THIN);
                style.setBorderLeft(BorderStyle.THIN);
                style.setBorderBottom(BorderStyle.THIN);
            }
            break;
            case(8):{//宋体12 加粗 垂直水平合并居中，自动换行
                font.setFontName("宋体");
                style.setVerticalAlignment(VerticalAlignment.CENTER);//垂直居中
                style.setAlignment(HorizontalAlignment.CENTER_SELECTION);
                font.setBold(true);//粗体
                font.setFontHeightInPoints((short) 12);//字体大小
                style.setFont(font);
                style.setWrapText(true);//自动换行
                style.setBorderTop(BorderStyle.THIN);
                style.setBorderRight(BorderStyle.THIN);
                style.setBorderLeft(BorderStyle.THIN);
                style.setBorderBottom(BorderStyle.THIN);
            }
            break;
            case(9):{//宋体12 加粗 垂直水平合并居左，自动换行
                font.setFontName("宋体");
                style.setVerticalAlignment(VerticalAlignment.CENTER);//垂直居中
                style.setAlignment(HorizontalAlignment.LEFT);
                font.setBold(true);//粗体
                font.setFontHeightInPoints((short) 12);//字体大小
                style.setFont(font);
                style.setWrapText(true);//自动换行
                style.setBorderTop(BorderStyle.THIN);
                style.setBorderRight(BorderStyle.THIN);
                style.setBorderLeft(BorderStyle.THIN);
                style.setBorderBottom(BorderStyle.THIN);
            }
            break;
            case(10):{//样式10，宋12 垂直中 合并左 自动换行
                font.setFontName("宋体");
                style.setVerticalAlignment(VerticalAlignment.CENTER);//垂直居中
                style.setAlignment(HorizontalAlignment.LEFT);
                //font.setBold(true);//粗体
                font.setFontHeightInPoints((short) 12);//字体大小
                style.setFont(font);
                style.setWrapText(true);//自动换行4
                style.setBorderTop(BorderStyle.THIN);
                style.setBorderRight(BorderStyle.THIN);
                style.setBorderLeft(BorderStyle.THIN);
                style.setBorderBottom(BorderStyle.THIN);
            }
            break;
        }

        return style;
    }

}
