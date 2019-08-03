package com.vivo.sprinbootpoi.controller;


import com.vivo.sprinbootpoi.entity.User;
import com.vivo.sprinbootpoi.mapper.UserMapper;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;

import javax.servlet.http.HttpServletResponse;
import java.io.OutputStream;
import java.util.List;
//接口api:     http://localhost:8080/swagger-ui.html#/
@Controller
@RequestMapping("/user")
public class ExcelController {

    @Autowired
    private UserMapper userMapper;

    @RequestMapping("/test")
    @ResponseBody
    public String voTest(){
        return "进入系统";
    }


    @RequestMapping("/getExcel")
    @ResponseBody
    public void getExcel (HttpServletResponse response) throws Exception {
        //得到实体类集合
        List<User> userList = userMapper.getAll();
        //创建一个excle文件
        HSSFWorkbook wb = new HSSFWorkbook();
        //创建一个工作表
        HSSFSheet sheet =wb.createSheet("附件 1   融资担保机构主要数据月报表");
        //
        HSSFRow row = null;

        /**
         * 创建了一个第一行，由10+1列组成
         */
        row = sheet.createRow(0);
        row.setHeight((short)(26.25*20));
        row.createCell(0).setCellValue("     附件  1                   融资担保机构主要数据月报表");
        row.getCell(0).setCellStyle(getStyle(wb,1));//设置样式
        for(int i = 1;i <= 9;i++){
            row.createCell(i).setCellStyle(getStyle(wb,1));
        }
        CellRangeAddress rowRegion = new CellRangeAddress(0,0,0,6);
        sheet.addMergedRegion(rowRegion);

        /**
         * 第二行年月
         */
        row = sheet.createRow(1);
        row.setHeight((short)(26.25*20));
        row.createCell(0).setCellValue("  年   月");

        row.getCell(0).setCellStyle(getStyle(wb,3));//设置样式
        for(int i = 1;i <= 9;i++){
            row.createCell(i).setCellStyle(getStyle(wb,3));
        }

        CellRangeAddress rowRegion2 = new CellRangeAddress(1,1,0,6);

        sheet.addMergedRegion(rowRegion2);

        /*CellRangeAddress columnRegion = new CellRangeAddress(1,4,0,0);
        sheet.addMergedRegion(columnRegion);*/
        /**
         * 第三行填报单位
         */
        row = sheet.createRow(2);
        row.setHeight((short)(26.25*20));
        row.createCell(0).setCellValue("填报单位 ");
        row.createCell(5).setCellValue("单位：万元、户、笔");
        row.getCell(0).setCellStyle(getStyle(wb,3));//设置样式
        row.getCell(5).setCellStyle(getStyle(wb,3));//设置样式

        for(int i = 2;i <= 4;i++){
            row.createCell(i).setCellStyle(getStyle(wb,3));
        }
        for(int i = 6;i <= 9;i++){
            row.createCell(i).setCellStyle(getStyle(wb,3));
        }
        CellRangeAddress rowRegion3 = new CellRangeAddress(2,2,0,3);
        CellRangeAddress rowRegion3_2 = new CellRangeAddress(2,2,5,5);
        sheet.addMergedRegion(rowRegion3);
        sheet.addMergedRegion(rowRegion3_2);
        /**
         * 4-5行数据标题
         */
         row = sheet.createRow(3);
         row.createCell(0).setCellValue("序号");
         row.getCell(0).setCellStyle(getStyle(wb,3));
        row.createCell(0).setCellValue("序号");
        row.getCell(0).setCellStyle(getStyle(wb,3));
        row.createCell(0).setCellValue("序号");
        row.getCell(0).setCellStyle(getStyle(wb,3));
        row.createCell(0).setCellValue("序号");
        row.getCell(0).setCellStyle(getStyle(wb,3));
        row.createCell(0).setCellValue("序号");
        row.getCell(0).setCellStyle(getStyle(wb,3));


        /**
         * 创建第二行
         */
       /* row = sheet.createRow(1);
        row.createCell(0).setCellStyle(getStyle(wb,3));
        row.setHeight((short)(22.50*20));
        row.createCell(1).setCellValue("用户Id");
        row.createCell(2).setCellValue("用户名");
        row.createCell(3).setCellValue("用户密码");
        for(int i = 1;i <= 3;i++){
            row.getCell(i).setCellStyle(getStyle(wb,1));
        }

        for(int i = 0;i<userList.size();i++){
            row = sheet.createRow(i+2);
            User user = userList.get(i);
            row.createCell(1).setCellValue(user.getId());
            row.createCell(2).setCellValue(user.getName());
            row.createCell(3).setCellValue(user.getAge());
            for(int j = 1;j <= 3;j++){
                row.getCell(j).setCellStyle(getStyle(wb,2));
            }
        }*/

        //默认行高
        sheet.setDefaultRowHeight((short)(16.5*20));
        //列宽自适应
        for(int i=0;i<=13;i++){
            sheet.autoSizeColumn(i);
        }

        response.setContentType("application/vnd.ms-excel;charset=utf-8");
        OutputStream os = response.getOutputStream();
        wb.write(os);
        os.flush();
        os.close();
    }

    /**
     * 获取样式
     * @param hssfWorkbook
     * @param styleNum
     * @return
     */
    public HSSFCellStyle getStyle(HSSFWorkbook hssfWorkbook, Integer styleNum){
        HSSFCellStyle style = hssfWorkbook.createCellStyle();
        style.setBorderRight(BorderStyle.THIN);//右边框
        style.setBorderBottom(BorderStyle.THIN);//下边框

        HSSFFont font = hssfWorkbook.createFont();
        font.setFontName("微软雅黑");//设置字体为微软雅黑

        HSSFPalette palette = hssfWorkbook.getCustomPalette();//拿到palette颜色板,可以根据需要设置颜色
        switch (styleNum){

            case(0):{//16号加粗
                style.setAlignment(HorizontalAlignment.CENTER_SELECTION);//跨列居中
                font.setBold(true);//粗体
                font.setFontHeightInPoints((short) 16);//字体大小
                style.setFont(font);
                //palette.setColorAtIndex(HSSFColor.BLUE.index,(byte)184,(byte)204,(byte)228);//替换颜色板中的颜色
                //style.setFillForegroundColor(HSSFColor.BLUE.index);
                //style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            }
            break;
            case(1):{//16号不加粗
                //font.setBold(true);//粗体
                font.setFontHeightInPoints((short) 16);//字体大小
                style.setFont(font);
            }
            break;
            case(2):{//11号加粗
                font.setBold(true);//粗体
                font.setFontHeightInPoints((short) 11);//字体大小
                style.setFont(font);
            }
            break;
            case(3):{//11号不加粗
                //font.setBold(true);//粗体
                font.setFontHeightInPoints((short) 11);//字体大小
                style.setFont(font);
            }
            break;
        }

        return style;
    }
}
