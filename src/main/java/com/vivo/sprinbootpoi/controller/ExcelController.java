package com.vivo.sprinbootpoi.controller;


import com.sun.org.apache.bcel.internal.generic.NEW;
import com.vivo.sprinbootpoi.entity.User;
import com.vivo.sprinbootpoi.mapper.UserMapper;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;

import javax.servlet.http.HttpServletResponse;
import javax.swing.*;
import javax.swing.border.Border;
import javax.swing.text.Style;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
//接口api:     http://localhost:8080/swagger-ui.html#/
@Controller
@RequestMapping("/user")
public class ExcelController {
    //创建样式集合
    private List<HSSFCellStyle> styles=new ArrayList<>();
    private List<HSSFFont> fonts=new ArrayList<>();
    //@Autowired
    //private UserMapper userMapper;
    @RequestMapping("/getExcel")
    @ResponseBody
    public void getExcel(HttpServletResponse response) throws Exception {
        //得到实体类集合
        //List<User> userList = userMapper.getAll();
        //创建一个excle文件
        HSSFWorkbook wb = new HSSFWorkbook();

        //创建一个工作表
        HSSFSheet sheet = wb.createSheet("附件 1   融资担保机构主要数据月报表");
        HSSFSheet sheet2=wb.createSheet(" 附件 2   融资担保机构主要数据年报表");

        //
        HSSFRow row = null;

        /**
         * 生成sheet1:附件 1   融资担保机构主要数据月报表
         */
        getSheet1(wb, sheet);

        /**
         * 生成sheet2:附件 2   融资担保机构主要数据年报表
         */
       getSheet2(wb, sheet2);

        response.setContentType("application/vnd.ms-excel;charset=utf-8");
        OutputStream os = response.getOutputStream();
        wb.write(os);
        os.flush();
        os.close();
    }

    /**
     * 生成sheet1
     * @param wb
     * @param sheet
     */
    private void getSheet1(HSSFWorkbook wb, HSSFSheet sheet) {
        HSSFRow row;
        row = sheet.createRow(0);
        row.setHeight((short) (26.25 * 20));
        row.createCell(0).setCellValue("     附件  1                   融资担保机构主要数据月报表");
        //row.getCell(0).setCellStyle(getStyle(wb, 1));//设置样式
        for (int i = 1; i <= 9; i++) {
            row.createCell(i);
        }
        CellRangeAddress rowRegion = new CellRangeAddress(0, 0, 0, 6);
        sheet.addMergedRegion(rowRegion);
        /**
         * 第二行年月
         */
        row = sheet.createRow(1);
        row.setHeight((short) (26.25 * 20));
        row.createCell(0).setCellValue("  年   月");
        for (int i = 1; i <= 9; i++) {
            row.createCell(i);
        }

        //合并单元格
        CellRangeAddress rowRegion2 = new CellRangeAddress(1, 1, 0, 6);
        sheet.addMergedRegion(rowRegion2);
        /*CellRangeAddress columnRegion = new CellRangeAddress(1,4,0,0);
        sheet.addMergedRegion(columnRegion);*/
        /**
         * 第三行填报单位
         */
        row = sheet.createRow(2);
        //row.setHeight((short)(26.25*20));
        //创建所有单元格
        for (int i = 0; i <= 9; i++) {
            row.createCell(i);
        }
        //设置特殊单元格的样式
        row.getCell(0).setCellValue("填报单位 ");
        row.getCell(5).setCellValue("单位：万元、户、笔");//F3
        //合并单元格
        CellRangeAddress rowRegion3 = new CellRangeAddress(2, 2, 0, 3);
        sheet.addMergedRegion(rowRegion3);


        /**
         * 4-5行数据标题
         */
        row = sheet.createRow(3);
        //创建所有单元格
        for (int i = 0; i <= 9; i++) {
            row.createCell(i);
        }
        row.getCell(0).setCellValue("序号");
        row.getCell(1).setCellValue("项目");
        row.getCell(4).setCellValue("数额");
        row.getCell(5).setCellValue("年初数");
        row.getCell(6).setCellValue("上年同期");
        row.getCell(7).setCellValue("月报表取数来源");
        // row.getCell(7).setCellStyle(getStyle(wb, 0));
        row.getCell(8).setCellValue("年度报表取数来源");
        // row.getCell(8).setCellStyle(getStyle(wb, 0));
        row.getCell(9).setCellValue("指标");
        // row.getCell(9).setCellStyle(getStyle(wb, 0));

        row = sheet.createRow(4);
        for (int i = 0; i <= 9; i++) {
            row.createCell(i);
        }
        //合并单元格
        CellRangeAddress rowRegion4 = new CellRangeAddress(3, 4, 0, 0);
        CellRangeAddress rowRegion4_1 = new CellRangeAddress(3, 4, 1, 3);
        CellRangeAddress rowRegion4_2 = new CellRangeAddress(3, 4, 4, 4);
        CellRangeAddress rowRegion4_3 = new CellRangeAddress(3, 4, 5, 5);
        CellRangeAddress rowRegion4_4 = new CellRangeAddress(3, 4, 6, 6);
        CellRangeAddress rowRegion4_5 = new CellRangeAddress(3, 4, 7, 7);
        CellRangeAddress rowRegion4_6 = new CellRangeAddress(3, 4, 8, 8);
        CellRangeAddress rowRegion4_7 = new CellRangeAddress(3, 4, 9, 9);
        sheet.addMergedRegion(rowRegion4);
        sheet.addMergedRegion(rowRegion4_1);
        sheet.addMergedRegion(rowRegion4_2);
        sheet.addMergedRegion(rowRegion4_3);
        sheet.addMergedRegion(rowRegion4_4);
        sheet.addMergedRegion(rowRegion4_5);
        sheet.addMergedRegion(rowRegion4_6);
        sheet.addMergedRegion(rowRegion4_7);
        /**
         * 6-98行内容的填充
         */
        for (int i = 5; i < 103; i++) {
            row = sheet.createRow(i);
           // row.setRowStyle(getStyle(wb, 3));
            //创建所有单元格
            for (int j = 0; j <= 6; j++) {
                row.createCell(j);
            }
            for (int k = 7; k <= 9; k++) {
                row.createCell(k);
            }
        }
        //填充数据,并合并
        //6行注册资本金
        sheet.getRow(5).getCell(0).setCellValue("1");
        sheet.getRow(5).getCell(1).setCellValue("1.  注册资本金");
        //sheet.getRow(5).getCell(1).setCellStyle(getStyle(wb, 0));
        CellRangeAddress rowRegion5 = new CellRangeAddress(5, 5, 1, 3);
        sheet.addMergedRegion(rowRegion5);
        //7-10行
        sheet.getRow(6).getCell(0).setCellValue("2");//A7
        sheet.getRow(6).getCell(1).setCellValue("2.   实收资本");//B7
        //sheet.getRow(6).getCell(1).setCellStyle(getStyle(wb, 2));
        sheet.getRow(7).getCell(1).setCellValue("其中");//B8
        // sheet.getRow(7).getCell(1).setCellStyle(getStyle(wb, 2));
        sheet.getRow(7).getCell(2).setCellValue("2.1   国有资本");//C8
        sheet.getRow(8).getCell(2).setCellValue("2.2   民营资本");//C9
        sheet.getRow(9).getCell(2).setCellValue("2.3   外资");//C10
        CellRangeAddress B7_D7 = new CellRangeAddress(6, 6, 1, 3);//B7-D7
        CellRangeAddress A7_A10 = new CellRangeAddress(6, 9, 0, 0);//A7-A10
        CellRangeAddress B8_B10 = new CellRangeAddress(7, 9, 1, 1);//B8-B10
        CellRangeAddress C8_D8 = new CellRangeAddress(7, 7, 2, 3);//C8-D8
        CellRangeAddress C9_D9 = new CellRangeAddress(8, 8, 2, 3);//C9-D9
        CellRangeAddress C10_D10 = new CellRangeAddress(9, 9, 2, 3);//C10-D10
        sheet.addMergedRegion(B7_D7);
        sheet.addMergedRegion(A7_A10);
        sheet.addMergedRegion(B8_B10);
        sheet.addMergedRegion(C8_D8);
        sheet.addMergedRegion(C9_D9);
        sheet.addMergedRegion(C10_D10);
        //11-12行
        sheet.getRow(10).getCell(0).setCellValue("3");//A11
        sheet.getRow(11).getCell(0).setCellValue("4");//A12
        sheet.getRow(10).getCell(1).setCellValue("3.资产总额");//b11
        //sheet.getRow(10).getCell(1).setCellStyle(getStyle(wb, 2));
        sheet.getRow(11).getCell(1).setCellValue("4.净资产");//b12
        //sheet.getRow(11).getCell(1).setCellStyle(getStyle(wb, 2));
        CellRangeAddress B11_D11 = new CellRangeAddress(10, 10, 1, 3);//B11_D11
        CellRangeAddress B12_D12 = new CellRangeAddress(11, 11, 1, 3);//B12_D12
        sheet.addMergedRegion(B11_D11);
        sheet.addMergedRegion(B12_D12);
        //13-26行:5. 本年累计发生担保金额
        sheet.getRow(12).getCell(1).setCellValue("5. 本年累计发生担保金额");//B13 5. 本年累计发生担保金额
        sheet.getRow(12).getCell(0).setCellValue("5");//A13 5
        sheet.getRow(13).getCell(1).setCellValue("按业务类型划分");//B14 按业务类型划分
        sheet.getRow(13).getCell(2).setCellValue("5.1融资担保业务");//C14 5.1融资担保业务
        sheet.getRow(18).getCell(2).setCellValue("5.2非融担保业务");//C19 5.2非融担保业务
        sheet.getRow(13).getCell(3).setCellValue("5.1.1 借款担保业务");//D14 5.1.1 借款担保业务
        sheet.getRow(14).getCell(3).setCellValue("5.1.1.2其中：互联网担保业务");//D15 5.1.1.2其中：互联网担保业务
        sheet.getRow(15).getCell(3).setCellValue("5.1.2 发行债券担保业务");//D16 5.1.2 发行债券担保业务
        sheet.getRow(16).getCell(3).setCellValue("5.1.3 其他融资担保业务");//D17 5.1.3 其他融资担保业务
        sheet.getRow(17).getCell(3).setCellValue("5.1.4 融资担保业务小计");//D18 5.1.4 融资担保业务小计
        sheet.getRow(18).getCell(3).setCellValue("5.2.1 工程履约担保");//D19 5.2.1 工程履约担保
        sheet.getRow(19).getCell(3).setCellValue("5.2.2 诉讼保全担保");//D20 5.2.2 诉讼保全担保
        sheet.getRow(20).getCell(3).setCellValue("5.2.3 其他非融担保业务");//d21 5.2.3 其他非融担保业务
        sheet.getRow(21).getCell(3).setCellValue("5.2.4 非融担保业务小计");//d22 5.2.4 非融担保业务小计
        sheet.getRow(22).getCell(1).setCellValue("按客户类型");//B23 按客户类型
        sheet.getRow(22).getCell(2).setCellValue("5.3 涉农担保");//C23 5.3 涉农担保
        sheet.getRow(23).getCell(2).setCellValue("5.4 小型企业担保");//C24 5.4 小型企业担保
        sheet.getRow(24).getCell(2).setCellValue("5.5  微型企业担保");//C25 5.5  微型企业担保
        sheet.getRow(25).getCell(2).setCellValue("5.6 个人担保");//C26 5.6 个人担保
        sheet.getRow(22).getCell(7).setCellValue("对应指标月末数-月初数");//H23 对应指标月末数-月初数
        sheet.getRow(22).getCell(8).setCellValue("对应指标数");//I23 对应指标数
        sheet.getRow(22).getCell(9).setCellValue("本年累计涉农担保发生担保金额");//J23 本年累计涉农担保发生担保金额
        for (int i = 12; i < 26; i++) {
            sheet.getRow(i).getCell(5).setCellValue("——");//f13-f22 ——
        }
        sheet.getRow(12).getCell(7).setCellValue("各业务类型月度金额之和");//h13 各业务类型月度金额之和
        sheet.getRow(12).getCell(8).setCellValue("各业务类型年度金额之和"); //i13 各业务类型年度金额之和
        sheet.getRow(13).getCell(7).setCellValue("对应指标月末数-月初数");//h14 对应指标月末数-月初数
        sheet.getRow(13).getCell(8).setCellValue("对应指标数");//i14 对应指标数
        sheet.getRow(13).getCell(9).setCellValue("本年累计借款担保业务发生担保金额");//j14 本年累计借款担保业务发生担保金额
        CellRangeAddress A13_A22 = new CellRangeAddress(12, 25, 0, 0);//A13_A26
        CellRangeAddress B13_D13 = new CellRangeAddress(12, 12, 1, 3);//B13_D13
        CellRangeAddress B14_B22 = new CellRangeAddress(13, 21, 1, 1);//B14_B22
        CellRangeAddress C14_C18 = new CellRangeAddress(13, 17, 2, 2); //C14_C18
        CellRangeAddress C19_C22 = new CellRangeAddress(18, 21, 2, 2);//C19_C22
        CellRangeAddress B23_B26 = new CellRangeAddress(22, 25, 1, 1);//B23_B26
        CellRangeAddress C23_D23 = new CellRangeAddress(22, 22, 2, 3);//C23_D23
        CellRangeAddress C24_D24 = new CellRangeAddress(23, 23, 2, 3);//C24_D24
        CellRangeAddress C25_D25 = new CellRangeAddress(24, 24, 2, 3);//C25_D25
        CellRangeAddress C26_D26 = new CellRangeAddress(25, 25, 2, 3);//C26_D26
        sheet.addMergedRegion(A13_A22);
        sheet.addMergedRegion(B13_D13);
        sheet.addMergedRegion(B14_B22);
        sheet.addMergedRegion(C14_C18);
        sheet.addMergedRegion(C19_C22);
        sheet.addMergedRegion(B23_B26);
        sheet.addMergedRegion(C23_D23);
        sheet.addMergedRegion(C24_D24);
        sheet.addMergedRegion(C25_D25);
        sheet.addMergedRegion(C26_D26);
        //27-40行:6.本年累计发生担保户数
        sheet.getRow(26).getCell(1).setCellValue("6.本年累计发生担保户数");//B13+14 6.本年累计发生担保户数
        sheet.getRow(26).getCell(0).setCellValue("6");//A13+14 5
        sheet.getRow(27).getCell(1).setCellValue("按业务类型划分");//B14+14 按业务类型划分
        sheet.getRow(27).getCell(2).setCellValue("6.1融资担保业务");//C14+14 6.1融资担保业务
        sheet.getRow(32).getCell(2).setCellValue("6.2非融担保业务");//C19+14 6.2非融担保业务
        sheet.getRow(27).getCell(3).setCellValue("6.1.1 借款担保业务");//D14+14 6.1.1 借款担保业务
        sheet.getRow(28).getCell(3).setCellValue("6.1.1.2其中：互联网担保业务");//D15+14 6.1.1.2其中：互联网担保业务
        sheet.getRow(29).getCell(3).setCellValue("6.1.2 发行债券担保业务");//D16+14 6.1.2 发行债券担保业务
        sheet.getRow(30).getCell(3).setCellValue("6.1.3 其他融资担保业务");//D17+14 6.1.3 其他融资担保业务
        sheet.getRow(31).getCell(3).setCellValue("6.1.4 融资担保业务小计");//D18+14 6.1.4 融资担保业务小计
        sheet.getRow(32).getCell(3).setCellValue("6.2.1 工程履约担保");//D19+14 6.2.1 工程履约担保
        sheet.getRow(33).getCell(3).setCellValue("6.2.2 诉讼保全担保");//D20+14 6.2.2 诉讼保全担保
        sheet.getRow(34).getCell(3).setCellValue("6.2.3 其他非融担保业务");//d21+14 6.2.3 其他非融担保业务
        sheet.getRow(36).getCell(3).setCellValue("6.2.4 非融担保业务小计");//d22+14 6.2.4 非融担保业务小计
        sheet.getRow(36).getCell(1).setCellValue("按客户类型");//B23+14 按客户类型
        sheet.getRow(36).getCell(2).setCellValue("6.3 涉农担保");//C23+14 6.3 涉农担保
        sheet.getRow(37).getCell(2).setCellValue("6.4 小型企业担保");//C24+14 6.4 小型企业担保
        sheet.getRow(38).getCell(2).setCellValue("6.5  微型企业担保");//C25+14 6.5  微型企业担保
        sheet.getRow(39).getCell(2).setCellValue("6.6 个人担保");//C26+14 6.6 个人担保
        sheet.getRow(36).getCell(7).setCellValue("对应指标月末数-月初数");//H23+14 对应指标月末数-月初数
        sheet.getRow(36).getCell(8).setCellValue("对应指标数");//I23+14 对应指标数
        sheet.getRow(36).getCell(9).setCellValue("本年累计涉农担保发生担保金额");//J23+14 本年累计涉农担保发生担保金额
        for (int i = 26; i < 40; i++) {
            sheet.getRow(i).getCell(5).setCellValue("——");//f13+14-f22+14 ——
        }
        sheet.getRow(26).getCell(7).setCellValue("各业务类型月度金额之和");//h13+14 各业务类型月度金额之和
        sheet.getRow(26).getCell(8).setCellValue("各业务类型年度金额之和"); //i13+14 各业务类型年度金额之和
        sheet.getRow(27).getCell(7).setCellValue("对应指标月末数-月初数");//h14+14 对应指标月末数-月初数
        sheet.getRow(27).getCell(8).setCellValue("对应指标数");//i14+14 对应指标数
        sheet.getRow(27).getCell(9).setCellValue("本年累计借款担保业务发生担保金额");//j14+14 本年累计借款担保业务发生担保金额
        CellRangeAddress A27_A40 = new CellRangeAddress(26, 39, 0, 0);//A27_A40
        CellRangeAddress B27_D27 = new CellRangeAddress(26, 26, 1, 3);//B27_D27
        CellRangeAddress B28_B36 = new CellRangeAddress(27, 35, 1, 1);//B28_B36
        CellRangeAddress C28_C32 = new CellRangeAddress(27, 31, 2, 2); //C28_C32
        CellRangeAddress C33_C36 = new CellRangeAddress(32, 35, 2, 2);//C33_C36
        CellRangeAddress B37_B40 = new CellRangeAddress(36, 39, 1, 1);//B37_B40
        CellRangeAddress C37_D37 = new CellRangeAddress(36, 36, 2, 3);//C37_D37
        CellRangeAddress C38_D38 = new CellRangeAddress(37, 37, 2, 3);//C38_D38
        CellRangeAddress C39D39 = new CellRangeAddress(38, 38, 2, 3);//C39D39
        CellRangeAddress C40_D40 = new CellRangeAddress(39, 39, 2, 3);//C40_D40
        sheet.addMergedRegion(A27_A40);
        sheet.addMergedRegion(B27_D27);
        sheet.addMergedRegion(B28_B36);
        sheet.addMergedRegion(C28_C32);
        sheet.addMergedRegion(C33_C36);
        sheet.addMergedRegion(B37_B40);
        sheet.addMergedRegion(C37_D37);
        sheet.addMergedRegion(C38_D38);
        sheet.addMergedRegion(C39D39);
        sheet.addMergedRegion(C40_D40);

        //27-40行:7.在保余额
        sheet.getRow(40).getCell(1).setCellValue("7.在保余额");//B13+14+14 7.在保余额
        sheet.getRow(40).getCell(0).setCellValue("7");//A13+14+14 7
        sheet.getRow(41).getCell(1).setCellValue("按业务类型划分");//B14+14+14 按业务类型划分
        sheet.getRow(41).getCell(2).setCellValue("7.1融资担保业务");//C14+14+14 7.1融资担保业务
        sheet.getRow(46).getCell(2).setCellValue("7.2非融担保业务");//C19+14+14 7.2非融担保业务
        sheet.getRow(41).getCell(3).setCellValue("7.1.1 借款担保业务");//D14+14+14 7.1.1 借款担保业务
        sheet.getRow(42).getCell(3).setCellValue("7.1.1.2其中：互联网担保业务");//D15+14+14 7.1.1.2其中：互联网担保业务
        sheet.getRow(43).getCell(3).setCellValue("7.1.2 发行债券担保业务");//D16+14+14 7.1.2 发行债券担保业务
        sheet.getRow(44).getCell(3).setCellValue("7.1.3 其他融资担保业务");//D17+14+14 7.1.3 其他融资担保业务
        sheet.getRow(45).getCell(3).setCellValue("7.1.4 融资担保业务小计");//D18+14+14 7.1.4 融资担保业务小计
        sheet.getRow(46).getCell(3).setCellValue("7.2.1 工程履约担保");//D19+14+14 7.2.1 工程履约担保
        sheet.getRow(47).getCell(3).setCellValue("7.2.2 诉讼保全担保");//D20+14+14 7.2.2 诉讼保全担保
        sheet.getRow(48).getCell(3).setCellValue("7.2.3 其他非融担保业务");//d21+14+14 7.2.3 其他非融担保业务
        sheet.getRow(50).getCell(3).setCellValue("7.2.4 非融担保业务小计");//d22+14+14 7.2.4 非融担保业务小计
        sheet.getRow(50).getCell(1).setCellValue("按客户类型");//B23+14+14 按客户类型
        sheet.getRow(50).getCell(2).setCellValue("7.3 涉农担保");//C23+14+14 7.3 涉农担保
        sheet.getRow(51).getCell(2).setCellValue("7.4 小型企业担保");//C24+14+14 7.4 小型企业担保
        sheet.getRow(52).getCell(2).setCellValue("7.5  微型企业担保");//C25+14+14 7.5  微型企业担保
        sheet.getRow(53).getCell(2).setCellValue("7.6 个人担保");//C26+14+14 7.6 个人担保
        sheet.getRow(50).getCell(7).setCellValue("对应指标月末数-月初数");//H23+14+14 对应指标月末数-月初数
        sheet.getRow(50).getCell(8).setCellValue("对应指标数");//I23+14+14 对应指标数
        sheet.getRow(50).getCell(9).setCellValue("本年累计涉农担保发生担保金额");//J23+14+14 本年累计涉农担保发生担保金额
        for (int i = 40; i < 54; i++) {
            sheet.getRow(i).getCell(5).setCellValue("——");//f13+14+14-f22+14+14 ——
        }
        sheet.getRow(40).getCell(7).setCellValue("各业务类型月度金额之和");//h13+14+14 各业务类型月度金额之和
        sheet.getRow(40).getCell(8).setCellValue("各业务类型年度金额之和"); //i13+14+14 各业务类型年度金额之和
        sheet.getRow(41).getCell(7).setCellValue("对应指标月末数-月初数");//h14+14+14 对应指标月末数-月初数
        sheet.getRow(41).getCell(8).setCellValue("对应指标数");//i14+14+14 对应指标数
        sheet.getRow(41).getCell(9).setCellValue("本年累计借款担保业务发生担保金额");//j14+14+14 本年累计借款担保业务发生担保金额
        CellRangeAddress A41_A54 = new CellRangeAddress(40, 53, 0, 0);//A41_A54
        CellRangeAddress B41_D41 = new CellRangeAddress(40, 40, 1, 3);//B41_D41
        CellRangeAddress B42_B50 = new CellRangeAddress(41, 49, 1, 1);//B42_B50
        CellRangeAddress C42_C46 = new CellRangeAddress(41, 45, 2, 2); //C42_C46
        CellRangeAddress C47_C50 = new CellRangeAddress(46, 49, 2, 2);//C47_C50
        CellRangeAddress B51_B54 = new CellRangeAddress(50, 53, 1, 1);//B51_B54
        CellRangeAddress C51_D51 = new CellRangeAddress(50, 50, 2, 3);//C51_D51
        CellRangeAddress C52_D52 = new CellRangeAddress(51, 51, 2, 3);//C52_D52
        CellRangeAddress C53_D53 = new CellRangeAddress(52, 52, 2, 3);//C53_D53
        CellRangeAddress C54_D54 = new CellRangeAddress(53, 53, 2, 3);//C54_D54
        sheet.addMergedRegion(A41_A54);
        sheet.addMergedRegion(B41_D41);
        sheet.addMergedRegion(B42_B50);
        sheet.addMergedRegion(C42_C46);
        sheet.addMergedRegion(C47_C50);
        sheet.addMergedRegion(B51_B54);
        sheet.addMergedRegion(C51_D51);
        sheet.addMergedRegion(C52_D52);
        sheet.addMergedRegion(C53_D53);
        sheet.addMergedRegion(C54_D54);


        //56行8.融资担保责任余额（按条例及配套制度计量）
        sheet.getRow(55).getCell(0).setCellValue("8");//A56 8
        sheet.getRow(55).getCell(1).setCellValue("8.融资担保责任余额（按条例及配套制度计量）");//B56 8.融资担保责任余额（按条例及配套制度计量）
        sheet.getRow(55).getCell(7).setCellValue("各业务类型月度融资担保责任余额之和");//H56 各业务类型月度融资担保责任余额之和
        sheet.getRow(55).getCell(8).setCellValue("各业务类型年度融资担保责任余额之和");//I56 各业务类型年度融资担保责任余额之和
        sheet.getRow(55).getCell(9).setCellValue("\uF06E 借款类担保责任余额 = 单户在保余额500万元人民币以下（含500万元人民币）的小微企业借款类担保在保余额 * 75% + 单户在保余额");//J56  借款类担保责任余额 = 单户在保余额500万元人民币以下（含500万元人民币）的小微企业借款类担保在保余额 * 75% + 单户在保余额
        CellRangeAddress B56_D56 = new CellRangeAddress(55, 55, 1, 3);//B56_D56
        sheet.addMergedRegion(B56_D56);

        //57-71（14） 9.在保户数
        int startLine=56;//57行开始
        sheet.getRow(startLine).getCell(1).setCellValue("9.在保户数");//B13+14+14+16 7.在保余额
        sheet.getRow(startLine).getCell(0).setCellValue("7");//A13+14+14+16 7
        sheet.getRow(startLine+1).getCell(1).setCellValue("按业务类型划分");//B14+14+14+16 按业务类型划分
        sheet.getRow(startLine+1).getCell(2).setCellValue("9.1融资担保业务");//C14+14+14+16 7.1融资担保业务
        sheet.getRow(startLine+7).getCell(2).setCellValue("9.2非融担保业务");//C19+14+14+16 7.2非融担保业务
        sheet.getRow(startLine+1).getCell(3).setCellValue("7.1.1 借款担保业务");//D14+14+14+16 7.1.1 借款担保业务
        sheet.getRow(startLine+2).getCell(3).setCellValue("7.1.1.2其中：互联网担保业务");//D15+14+14+16 7.1.1.2其中：互联网担保业务
        sheet.getRow(startLine+3).getCell(3).setCellValue("7.1.2 发行债券担保业务");//D16+14+14+16 7.1.2 发行债券担保业务
        sheet.getRow(startLine+4).getCell(3).setCellValue("7.1.3 其他融资担保业务");//D17+14+14+16 7.1.3 其他融资担保业务
        sheet.getRow(startLine+5).getCell(3).setCellValue("9.1.3.1 其中：保本基金担保业务");
        sheet.getRow(startLine+6).getCell(3).setCellValue("7.1.4 融资担保业务小计");//D18+14+14+16 7.1.4 融资担保业务小计
        sheet.getRow(startLine+7).getCell(3).setCellValue("7.2.1 工程履约担保");//D19+14+14+16 7.2.1 工程履约担保
        sheet.getRow(startLine+8).getCell(3).setCellValue("7.2.2 诉讼保全担保");//D20+14+14+16 7.2.2 诉讼保全担保
        sheet.getRow(startLine+9).getCell(3).setCellValue("7.2.3 其他非融担保业务");//d21+14+14+16 7.2.3 其他非融担保业务
        sheet.getRow(startLine+10).getCell(3).setCellValue("7.2.4 非融担保业务小计");//d22+14+14+16 7.2.4 非融担保业务小计
        sheet.getRow(startLine+11).getCell(1).setCellValue("按客户类型");//B23+14+14+16 按客户类型
        sheet.getRow(startLine+11).getCell(2).setCellValue("7.3 涉农担保");//C23+14+14+16 7.3 涉农担保
        sheet.getRow(startLine+12).getCell(2).setCellValue("7.4 小型企业担保");//C24+14+14+16 7.4 小型企业担保
        sheet.getRow(startLine+13).getCell(2).setCellValue("7.5  微型企业担保");//C25+14+14+16 7.5  微型企业担保
        sheet.getRow(startLine+14).getCell(2).setCellValue("7.6 个人担保");//C26+14+14+16 7.6 个人担保
        sheet.getRow(startLine+11).getCell(7).setCellValue("对应指标月末数-月初数");//H23+14+14+16 对应指标月末数-月初数
        sheet.getRow(startLine+11).getCell(8).setCellValue("对应指标数");//I23+14+14+16 对应指标数
        sheet.getRow(startLine+11).getCell(8).setCellValue("对应指标数");//I23+14+14+16 对应指标数
        sheet.getRow(startLine+11).getCell(9).setCellValue("涉农担保业务在保户数");
        /*for (int i = 40; i < 54; i++) {
            sheet.getRow(i).getCell(5).setCellValue("——");
        }*/
        sheet.getRow(startLine).getCell(7).setCellValue("各业务类型月度金额之和");//h13+14+14+16 各业务类型月度金额之和
        sheet.getRow(startLine).getCell(8).setCellValue("各业务类型年度金额之和"); //i13+14+14+16 各业务类型年度金额之和
        sheet.getRow(startLine+1).getCell(7).setCellValue("对应指标月末数-月初数");//h14+14+14+16 对应指标月末数-月初数
        sheet.getRow(startLine+1).getCell(8).setCellValue("对应指标数");//i14+14+14+16 对应指标数
        sheet.getRow(startLine+1).getCell(9).setCellValue("借款担保业务在保户数");//j14+14+14+16 本年累计借款担保业务发生担保金额
        CellRangeAddress A57_A71 = new CellRangeAddress(startLine, startLine+14, 0, 0);//A57_A71
        CellRangeAddress B58_B67 = new CellRangeAddress(57, 66, 1, 1);//B58_B67
        CellRangeAddress B68_B71 = new CellRangeAddress(67, 70, 1, 1);//B68_B71
        CellRangeAddress B57_D57 = new CellRangeAddress(56, 56, 1, 3);//B57_D57
        CellRangeAddress C58_C63 = new CellRangeAddress(57, 62, 2, 2); //C58_C63
        CellRangeAddress C64_C67 = new CellRangeAddress(63, 66, 2, 2);//C64_C67
        CellRangeAddress C68_D68 = new CellRangeAddress(67, 67, 2, 3);//C68_D68
        CellRangeAddress C69_D69 = new CellRangeAddress(68, 68, 2, 3);//C69_D69
        CellRangeAddress C70_D70 = new CellRangeAddress(69, 69, 2, 3);//C70_D70
        CellRangeAddress C71_D71 = new CellRangeAddress(70, 70, 2, 3);//C71_D71
        sheet.addMergedRegion(A57_A71);
        sheet.addMergedRegion(B58_B67);
        sheet.addMergedRegion(B68_B71);
        sheet.addMergedRegion(B57_D57);
        sheet.addMergedRegion(C58_C63);
        sheet.addMergedRegion(C64_C67);
        sheet.addMergedRegion(C68_D68);
        sheet.addMergedRegion(C69_D69);
        sheet.addMergedRegion(C70_D70);
        sheet.addMergedRegion(C71_D71);
        //72-75行 损益类指标
        sheet.getRow(71).getCell(0).setCellValue("10");
        sheet.getRow(73).getCell(0).setCellValue("11");
        sheet.getRow(74).getCell(0).setCellValue("12");
        sheet.getRow(72).getCell(1).setCellValue("损益类指标");
        sheet.getRow(71).getCell(2).setCellValue("10.本年累计营业收入");
        sheet.getRow(72).getCell(2).setCellValue("其中  10.1本年累计担保费收入");
        sheet.getRow(73).getCell(2).setCellValue("11.本年累计净利润");
        sheet.getRow(74).getCell(2).setCellValue("12.本年累计缴纳税金");
        for (int i = 71; i < 75; i++) {
            sheet.getRow(i).getCell(5).setCellValue("——");
        }
        sheet.getRow(72).getCell(7).setCellValue("指标月末数-月初数");//h73  指标月末数-月初数
        sheet.getRow(72).getCell(8).setCellValue("指标月末数-月初数");//I73 指标年末数-年初数
        sheet.getRow(72).getCell(9).setCellValue("指标月末数-月初数");//J73  累计担保费收入
        CellRangeAddress a72_a73 = new CellRangeAddress(71, 72, 0, 0);//a72_a73
        CellRangeAddress B72_B75 = new CellRangeAddress(71, 74, 1, 1);//B72_B75
        CellRangeAddress C72_D72 = new CellRangeAddress(71, 71, 2, 3); //C72_D72
        CellRangeAddress C73_D73 = new CellRangeAddress(72, 72, 2, 3);//C73_D73
        CellRangeAddress C74_D74 = new CellRangeAddress(73, 73, 2, 3);//C74_D74
        CellRangeAddress C75_D75 = new CellRangeAddress(74, 74, 2, 3); //C75_D75
        sheet.addMergedRegion(a72_a73);
        sheet.addMergedRegion(B72_B75);
        sheet.addMergedRegion(C72_D72);
        sheet.addMergedRegion(C73_D73);
        sheet.addMergedRegion(C74_D74);
        sheet.addMergedRegion(C75_D75);
        //76-98行 风险类指标
        sheet.getRow(75).getCell(0).setCellValue("13");//A76 13
        sheet.getRow(77).getCell(0).setCellValue("14");//A78 14
        sheet.getRow(78).getCell(0).setCellValue("15");//A79 15
        sheet.getRow(80).getCell(0).setCellValue("16");//A81 16
        sheet.getRow(81).getCell(0).setCellValue("17");//A82 17
        sheet.getRow(83).getCell(0).setCellValue("18");//A84 18
        sheet.getRow(84).getCell(0).setCellValue("19");//A85 19
        sheet.getRow(86).getCell(0).setCellValue("20");//A87 20
        sheet.getRow(87).getCell(0).setCellValue("21");//A88 21
        sheet.getRow(91).getCell(0).setCellValue("22");//A92 22
        sheet.getRow(94).getCell(0).setCellValue("23");//A95 23
        sheet.getRow(75).getCell(1).setCellValue("风险类指标");//B76 风险类指标
        sheet.getRow(92).getCell(1).setCellValue("其中");//B93 其中
        sheet.getRow(95).getCell(1).setCellValue("其中");//B96 其中
        sheet.getRow(75).getCell(2).setCellValue("13.本年累计解除担保业务金额");//C76 13.本年累计解除担保业务金额
        sheet.getRow(76).getCell(2).setCellValue(" 其中  13.1本年累计解除融资担保业务金额");//C77  其中  13.1本年累计解除融资担保业务金额
        sheet.getRow(77).getCell(2).setCellValue("14.累计解除担保业务金额");//C78 14.累计解除担保业务金额
        sheet.getRow(78).getCell(2).setCellValue("15.本年累计担保业务代偿金额");//C79 15.本年累计担保业务代偿金额
        sheet.getRow(79).getCell(2).setCellValue("       其中    15.1本年累计融资担保代偿金额");//C80        其中    15.1本年累计融资担保代偿金额
        sheet.getRow(80).getCell(2).setCellValue("16.累计担保业务代偿金额");//C81 16.累计担保业务代偿金额
        sheet.getRow(81).getCell(2).setCellValue("17.本年累计担保业务损失金额");//C82 17.本年累计担保业务损失金额
        sheet.getRow(82).getCell(2).setCellValue(" 其中   17.1 本年累计融资担保损失金额");//C83  其中   17.1 本年累计融资担保损失金额
        sheet.getRow(83).getCell(2).setCellValue("18.累计担保损失金额");//C84 18.累计担保损失金额
        sheet.getRow(84).getCell(2).setCellValue("19.代偿余额");//C85 19.代偿余额
        sheet.getRow(85).getCell(2).setCellValue("其中   19.1融资担保代偿余额");//C86 其中   19.1融资担保代偿余额
        sheet.getRow(86).getCell(2).setCellValue("20.逾期未代偿金额");//C87 20.逾期未代偿金额
        sheet.getRow(87).getCell(2).setCellValue("21.担保准备金");//C88 21.担保准备金
        sheet.getRow(88).getCell(2).setCellValue("其中");//C89 其中
        sheet.getRow(88).getCell(3).setCellValue("21.1担保赔偿准备金");//D89 21.1担保赔偿准备金
        sheet.getRow(89).getCell(3).setCellValue("21.2未到期责任准备金");//D90 21.2未到期责任准备金
        sheet.getRow(90).getCell(3).setCellValue("21.3一般风险准备金");//D91 21.3一般风险准备金
        sheet.getRow(91).getCell(1).setCellValue("22.对外投资");//B92 22.对外投资
        sheet.getRow(92).getCell(2).setCellValue("22.1股权投资");//C93 22.1股权投资
        sheet.getRow(93).getCell(2).setCellValue("22.2 委托贷款");//C94 22.2 委托贷款
        sheet.getRow(94).getCell(1).setCellValue("23.在职人数");//B95 23.在职人数
        sheet.getRow(95).getCell(2).setCellValue("3.1研究生及以上");//C96 3.1研究生及以上
        sheet.getRow(96).getCell(2).setCellValue("23.2本科");//C97 23.2本科
        sheet.getRow(97).getCell(2).setCellValue("23.3大专及以下");//C98 23.3大专及以下

        sheet.getRow(75).getCell(5).setCellValue("——");//F76-F77 ——
        sheet.getRow(78).getCell(5).setCellValue("——");//F79-F80 ——
        sheet.getRow(81).getCell(5).setCellValue("——"); //F82-F83 ——
        sheet.getRow(76).getCell(5).setCellValue("——");//F76-F77 ——
        sheet.getRow(79).getCell(5).setCellValue("——");//F79-F80 ——
        sheet.getRow(82).getCell(5).setCellValue("——"); //F82-F83 ——
        sheet.getRow(72).getCell(7).setCellValue("指标月末数-月初数");//H73 指标月末数-月初数
        sheet.getRow(75).getCell(7).setCellValue("指标月末数-月初数");//H76 指标月末数-月初数
        sheet.getRow(76).getCell(7).setCellValue("指标月末数-月初数");//H77 指标月末数-月初数
        sheet.getRow(77).getCell(7).setCellValue("指标数");//H78 指标数
        sheet.getRow(78).getCell(7).setCellValue("指标月末数-月初数");//H79 指标月末数-月初数
        sheet.getRow(79).getCell(7).setCellValue("指标月末数-月初数");//H80 指标月末数-月初数
        sheet.getRow(80).getCell(7).setCellValue("A:指标数");//H81 A:指标数
        sheet.getRow(81).getCell(7).setCellValue("指标月末数-月初数");//H82 指标月末数-月初数
        sheet.getRow(82).getCell(7).setCellValue("指标月末数-月初数");//H83 指标月末数-月初数
        sheet.getRow(83).getCell(7).setCellValue("B：指标数");//H84 B：指标数
        sheet.getRow(84).getCell(7).setCellValue("A-B-C月末数-C月初数");//H85 A-B-C月末数-C月初数
        sheet.getRow(85).getCell(7).setCellValue("A-B-C月末数-C月初数 （融资担保类）");//H86 A-B-C月末数-C月初数 （融资担保类）
        sheet.getRow(86).getCell(7).setCellValue("指标月末数-月初数");//H87 指标月末数-月初数
        sheet.getRow(72).getCell(8).setCellValue("指标年末数-年初数");//I73 指标年末数-年初数
        sheet.getRow(75).getCell(8).setCellValue("指标年末数-年初数");//I76 指标年末数-年初数
        sheet.getRow(76).getCell(8).setCellValue("指标年末数-年初数");//I77 指标年末数-年初数
        sheet.getRow(77).getCell(8).setCellValue("指标数");//I78 指标数
        sheet.getRow(78).getCell(8).setCellValue("指标年末数-年初数");//I79 指标年末数-年初数
        sheet.getRow(79).getCell(8).setCellValue("指标年末数-年初数");//I80 指标年末数-年初数
        sheet.getRow(80).getCell(8).setCellValue("A1:指标数");//I81 A1:指标数
        sheet.getRow(81).getCell(8).setCellValue("指标年末数-年初数");//I82 指标年末数-年初数
        sheet.getRow(82).getCell(8).setCellValue("指标年末数-年初数");//I83 指标年末数-年初数
        sheet.getRow(83).getCell(8).setCellValue("B1:指标数");//I84 B1:指标数
        sheet.getRow(84).getCell(8).setCellValue("C年末数-C年初数");//I85 C年末数-C年初数
        sheet.getRow(85).getCell(8).setCellValue("C年末数-C年初数（融资担保类）");//I86 C年末数-C年初数（融资担保类）
        sheet.getRow(86).getCell(8).setCellValue("指标年末数-年初数");//I87 指标年末数-年初数
        sheet.getRow(72).getCell(9).setCellValue("累计担保费收入");//J73 累计担保费收入
        sheet.getRow(75).getCell(9).setCellValue("累计解除担保业务金额");//J76 累计解除担保业务金额
        sheet.getRow(76).getCell(9).setCellValue("累计解除融资担保业务金额");//J77 累计解除融资担保业务金额
        sheet.getRow(77).getCell(9).setCellValue("累计解除担保业务金额");//J78 累计解除担保业务金额
        sheet.getRow(78).getCell(9).setCellValue("累计担保代偿金额");//J79 累计担保代偿金额
        sheet.getRow(79).getCell(9).setCellValue("累计解除融资担保代偿金额");//J80 累计解除融资担保代偿金额
        sheet.getRow(80).getCell(9).setCellValue("累计担保代偿金额");//J81 累计担保代偿金额
        sheet.getRow(81).getCell(9).setCellValue("累计担保损失金额");//J82 累计担保损失金额
        sheet.getRow(82).getCell(9).setCellValue("累计解除融资担保损失金额");//J83 累计解除融资担保损失金额
        sheet.getRow(83).getCell(9).setCellValue("累计担保损失金额");//J84 累计担保损失金额
        sheet.getRow(84).getCell(9).setCellValue("C:累计收回担保损失金额");//J85 C:累计收回担保损失金额
        sheet.getRow(85).getCell(9).setCellValue("累计收回融资担保损失金额");//J86 累计收回融资担保损失金额
        sheet.getRow(86).getCell(9).setCellValue("逾期未代偿金额");//J87 逾期未代偿金额

        CellRangeAddress A76_A77=new CellRangeAddress(75,76,0,0);//A76_A77
        CellRangeAddress A79_A80=new CellRangeAddress(78,79,0,0);//A79_A80
        CellRangeAddress A82_A83=new CellRangeAddress(81,82,0,0);//A82_A83
        CellRangeAddress A85_A86=new CellRangeAddress(84,85,0,0);//A85_A86
        CellRangeAddress A88_A91=new CellRangeAddress(87,90,0,0);//A88_A91
        CellRangeAddress A92_A94=new CellRangeAddress(91,93,0,0);//A92_A94
        CellRangeAddress A95_A98=new CellRangeAddress(94,97,0,0);//A95_A98
        CellRangeAddress B76_B91=new CellRangeAddress(75,90,1,1);//B76_B91
        CellRangeAddress B92_D92=new CellRangeAddress(91,91,1,3);//B92_D92
        CellRangeAddress B93_B94=new CellRangeAddress(92,93,1,1);//B93_B94
        CellRangeAddress B95_D95=new CellRangeAddress(94,94,1,3);//B95_D95
        CellRangeAddress B96_B98=new CellRangeAddress(95,97,1,1);//B96_B98
        CellRangeAddress C89_C91=new CellRangeAddress(88,90,2,2);//C89_C91
        CellRangeAddress C93_D93=new CellRangeAddress(92,92,2,3);//C93_D93
        CellRangeAddress C94_D94=new CellRangeAddress(93,93,2,3);//C94_D94
        CellRangeAddress C96_D96=new CellRangeAddress(95,95,2,3);//C96_D96
        CellRangeAddress C97_D97=new CellRangeAddress(96,96,2,3);//C97_D97
        CellRangeAddress C98_D98=new CellRangeAddress(97,97,2,3);//C98_D98
        for (int i = 75; i < 88; i++) {
            CellRangeAddress cdrange=new CellRangeAddress(i,i,2,3);//C76_D76~C88_D88
            sheet.addMergedRegion(cdrange);
        }
        sheet.addMergedRegion(A76_A77);
        sheet.addMergedRegion(A79_A80);
        sheet.addMergedRegion(A82_A83);
        sheet.addMergedRegion(A85_A86);
        sheet.addMergedRegion(A88_A91);
        sheet.addMergedRegion(A92_A94);
        sheet.addMergedRegion(A95_A98);
        sheet.addMergedRegion(B76_B91);
        sheet.addMergedRegion(B92_D92);
        sheet.addMergedRegion(B93_B94);
        sheet.addMergedRegion(B95_D95);
        sheet.addMergedRegion(B96_B98);
        sheet.addMergedRegion(C89_C91);
        sheet.addMergedRegion(C93_D93);
        sheet.addMergedRegion(C94_D94);
        sheet.addMergedRegion(C96_D96);
        sheet.addMergedRegion(C97_D97);
        sheet.addMergedRegion(C98_D98);

        //99-103末尾信息

        sheet.getRow(98).getCell(2).setCellValue("审核人：");//C99  审核人：
        sheet.getRow(98).getCell(5).setCellValue("联系电话：");//F99 联系电话：
        sheet.getRow(99).getCell(0).setCellValue("填报说明：1.表格中，除“——”框中不填数据，其余均要求填写数据，若没有发生业务，填“0”；2.“涉农担保”是指融资担保法人机构为“三农”融资提供担保的业务，包含新型农业经营主体；小型企业、微型企业划分标准及统计口径按照《关于印发中小企业划型标准规定的通知》（工信部联企业[2011]300号）的有关规定执行，其中，微型企业包括微型企业、个体工商户及小微企业主；3.“融资担保责任余额”的计量见《融资担保责任余额计量办法》；4.“损失金额”是指有诉讼判决书或仲裁书和强制执行书，或者其他足以证明损失已形成的证据，证明代偿已无法收回的业务余额；5.涉及“当年累计”指标应填写本公司在本年度累计发生数据，涉及“累计”指标应填写公司成立以来发生数据；6.“逾期未代偿金额”是指担保业务已逾期公司应进行代偿而目前未代偿业务；7.涉及“担保业务收入”、“净利润”、“税金”应填列当年累计数额及去年同期累计数额。");//A100 填报说明：1.表格中，除“——”框中不填数据，其余均要求填写数据，若没有发生业务，填“0”；2.“涉农担保”是指融资担保法人机构为“三农”融资提供担保的业务，包含新型农业经营主体；小型企业、微型企业划分标准及统计口径按照《关于印发中小企业划型标准规定的通知》（工信部联企业[2011]300号）的有关规定执行，其中，微型企业包括微型企业、个体工商户及小微企业主；3.“融资担保责任余额”的计量见《融资担保责任余额计量办法》；4.“损失金额”是指有诉讼判决书或仲裁书和强制执行书，或者其他足以证明损失已形成的证据，证明代偿已无法收回的业务余额；5.涉及“当年累计”指标应填写本公司在本年度累计发生数据，涉及“累计”指标应填写公司成立以来发生数据；6.“逾期未代偿金额”是指担保业务已逾期公司应进行代偿而目前未代偿业务；7.涉及“担保业务收入”、“净利润”、“税金”应填列当年累计数额及去年同期累计数额。
        CellRangeAddress A100_G103=new CellRangeAddress(99,102,0,6);//A100-G103
        sheet.addMergedRegion(A100_G103);
        /**
         *统一样式设置
         */
        //默认行高
        sheet.setDefaultRowHeight((short) (33));
        //列宽
        sheet.setColumnWidth(0,6*256);
        sheet.setColumnWidth(1,12*256);
        sheet.setColumnWidth(2,22*256);
        sheet.setColumnWidth(3,33*256);
        sheet.setColumnWidth(4,9*256);
        sheet.setColumnWidth(5,9*256);
        sheet.setColumnWidth(6,15*256);
        sheet.setColumnWidth(7,18*256);
        sheet.setColumnWidth(8,19*256);
        sheet.setColumnWidth(9,40*256);

        //1.样式0：全局设置边框
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

        //2.样式1：方正小标宋_GBK 16 垂直居中
        sheet.getRow(0).getCell(0).setCellStyle(getStyle(wb,1));//A1
        //3.样式2：宋体16 垂直，水平居中 加粗
        sheet.getRow(1).getCell(0).setCellStyle(getStyle(wb,2));//A2
        //4.样式3：宋体12 加粗水平居左，垂直居中
        sheet.getRow(2).getCell(0).setCellStyle(getStyle(wb,3));//A3
        sheet.getRow(2).getCell(5).setCellStyle(getStyle(wb,3));//F3
        sheet.getRow(98).getCell(2).setCellStyle(getStyle(wb,3));//C99
        sheet.getRow(98).getCell(5).setCellStyle(getStyle(wb,3));//F99

        //5.样式4：宋体10 加粗 垂直水平居中
        sheet.getRow(3).getCell(7).setCellStyle(getStyle(wb,4));//H4
        sheet.getRow(3).getCell(8).setCellStyle(getStyle(wb,4));//I4
        sheet.getRow(3).getCell(9).setCellStyle(getStyle(wb,4));//J4
        //6.样式5：宋体11 加粗 垂直水平居中
        sheet.getRow(5).getCell(1).setCellStyle(getStyle(wb,5)); //B6
        sheet.getRow(5).getCell(2).setCellStyle(getStyle(wb,5)); //C6
        sheet.getRow(5).getCell(3).setCellStyle(getStyle(wb,5)); //D6

        sheet.getRow(6).getCell(1).setCellStyle(getStyle(wb,5)); //B7
        sheet.getRow(10).getCell(1).setCellStyle(getStyle(wb,5)); //B11
        sheet.getRow(11).getCell(1).setCellStyle(getStyle(wb,5)); //B12
        sheet.getRow(12).getCell(1).setCellStyle(getStyle(wb,5)); //B13
        sheet.getRow(26).getCell(1).setCellStyle(getStyle(wb,5)); //B27
        sheet.getRow(40).getCell(1).setCellStyle(getStyle(wb,5)); //B41
        sheet.getRow(91).getCell(1).setCellStyle(getStyle(wb,5)); //B92
        sheet.getRow(94).getCell(1).setCellStyle(getStyle(wb,5)); //B95
        sheet.getRow(55).getCell(1).setCellStyle(getStyle(wb,5)); //B56
        sheet.getRow(56).getCell(1).setCellStyle(getStyle(wb,5)); //B57
        sheet.getRow(71).getCell(2).setCellStyle(getStyle(wb,5)); //C72
        sheet.getRow(73).getCell(2).setCellStyle(getStyle(wb,5)); //C74
        sheet.getRow(74).getCell(2).setCellStyle(getStyle(wb,5)); //C75
        sheet.getRow(77).getCell(2).setCellStyle(getStyle(wb,5)); //C78
        sheet.getRow(78).getCell(2).setCellStyle(getStyle(wb,5)); //C79
        sheet.getRow(80).getCell(2).setCellStyle(getStyle(wb,5)); //C81
        sheet.getRow(81).getCell(2).setCellStyle(getStyle(wb,5)); //C82
        sheet.getRow(83).getCell(2).setCellStyle(getStyle(wb,5)); //C84
        sheet.getRow(84).getCell(2).setCellStyle(getStyle(wb,5)); //C85
        sheet.getRow(86).getCell(2).setCellStyle(getStyle(wb,5)); //C87
        sheet.getRow(87).getCell(2).setCellStyle(getStyle(wb,5)); //C88
        //7.样式6：宋体11 垂直居中，水平居中
        sheet.getRow(3).getCell(1).setCellStyle(getStyle(wb,6));//B4
        //E4_G98
        for (int i = 4; i < 7; i++) {
            for (int i1 = 3; i1 < 98; i1++) {
                sheet.getRow(i1).getCell(i).setCellStyle(getStyle(wb,6));
            }
        }
        //8.样式7.宋体10 垂直水平居中，自动换行
        //A4_A95
        for (int i = 3; i < 95; i++) {
            sheet.getRow(i).getCell(0).setCellStyle(getStyle(wb,7));
        }
        //H6_I103
        for (int i = 7; i < 9; i++) {
            for (int i1 =5; i1 < 103; i1++) {
                sheet.getRow(i1).getCell(i).setCellStyle(getStyle(wb,7));
            }
        }
        //9.样式8.宋体12 加粗 垂直水平合并居中，自动换行
        sheet.getRow(7).getCell(1).setCellStyle(getStyle(wb,8)); //B8
        sheet.getRow(71).getCell(1).setCellStyle(getStyle(wb,8)); //B72
        sheet.getRow(75).getCell(1).setCellStyle(getStyle(wb,8)); //B76
        sheet.getRow(92).getCell(1).setCellStyle(getStyle(wb,8)); //B93
        sheet.getRow(95).getCell(1).setCellStyle(getStyle(wb,8)); //B96
        //10.样式9.宋体12 加粗 垂直水平合并居左，自动换行
        sheet.getRow(13).getCell(1).setCellStyle(getStyle(wb,9)); //B14
        sheet.getRow(22).getCell(1).setCellStyle(getStyle(wb,9)); //B23
        sheet.getRow(27).getCell(1).setCellStyle(getStyle(wb,9)); //B28
        sheet.getRow(36).getCell(1).setCellStyle(getStyle(wb,9)); //B37
        sheet.getRow(41).getCell(1).setCellStyle(getStyle(wb,9)); //B42
        sheet.getRow(51).getCell(1).setCellStyle(getStyle(wb,9)); //B52
        sheet.getRow(57).getCell(1).setCellStyle(getStyle(wb,9)); //B58
        sheet.getRow(67).getCell(1).setCellStyle(getStyle(wb,9)); //B68

        //11.样式10，宋12 垂直中 合并左 自动换行
        sheet.getRow(99).getCell(0).setCellStyle(getStyle(wb,10));//A100
        //12.样式11，宋11 垂直中 水平中 自动换行
        sheet.getRow(13).getCell(2).setCellStyle(getStyle(wb,11));//C14
        sheet.getRow(18).getCell(2).setCellStyle(getStyle(wb,11));//C19
        sheet.getRow(27).getCell(2).setCellStyle(getStyle(wb,11));//C28
        sheet.getRow(32).getCell(2).setCellStyle(getStyle(wb,11));//C33
        sheet.getRow(41).getCell(2).setCellStyle(getStyle(wb,11));//C42
        sheet.getRow(47).getCell(2).setCellStyle(getStyle(wb,11));//C47
        sheet.getRow(57).getCell(2).setCellStyle(getStyle(wb,11));//C58
        sheet.getRow(63).getCell(2).setCellStyle(getStyle(wb,11));//C64
        sheet.getRow(88).getCell(2).setCellStyle(getStyle(wb,11));//C89
        //13.样式12，宋10 垂直居中 自动换行
        //I6_I103
        for (int i = 5; i < 103; i++) {
            sheet.getRow(i).getCell(9).setCellStyle(getStyle(wb,12));
        }
    }

    /**
     * 生成sheet2
     * @param wb
     * @param sheet
     */
    private void getSheet2(HSSFWorkbook wb, HSSFSheet sheet) {
        HSSFRow row;
        row = sheet.createRow(0);
        row.setHeight((short) (26.25 * 20));
        row.createCell(0).setCellValue("附件  2                  融资担保机构主要数据年报表");
        //row.getCell(0).setCellStyle(getStyle(wb, 1));//设置样式
        for (int i = 1; i <= 9; i++) {
            row.createCell(i);
        }
        CellRangeAddress rowRegion = new CellRangeAddress(0, 0, 0, 6);
        sheet.addMergedRegion(rowRegion);
        /**
         * 第二行年月
         */
        row = sheet.createRow(1);
        row.setHeight((short) (26.25 * 20));
        row.createCell(0).setCellValue("  年   月");
        for (int i = 1; i <= 9; i++) {
            row.createCell(i);
        }

        //合并单元格
        CellRangeAddress rowRegion2 = new CellRangeAddress(1, 1, 0, 6);
        sheet.addMergedRegion(rowRegion2);
        /*CellRangeAddress columnRegion = new CellRangeAddress(1,4,0,0);
        sheet.addMergedRegion(columnRegion);*/
        /**
         * 第三行填报单位
         */
        row = sheet.createRow(2);
        //row.setHeight((short)(26.25*20));
        //创建所有单元格
        for (int i = 0; i <= 9; i++) {
            row.createCell(i);
        }
        //设置特殊单元格的样式
        row.getCell(0).setCellValue("填报单位 ");
        row.getCell(5).setCellValue("单位：万元、户、笔");//F3
        //合并单元格
        CellRangeAddress rowRegion3 = new CellRangeAddress(2, 2, 0, 3);
        sheet.addMergedRegion(rowRegion3);


        /**
         * 4-5行数据标题
         */
        row = sheet.createRow(3);
        //创建所有单元格
        for (int i = 0; i <= 9; i++) {
            row.createCell(i);
        }
        row.getCell(0).setCellValue("序号");
        row.getCell(1).setCellValue("项目");
        row.getCell(4).setCellValue("数额");
        row.getCell(5).setCellValue("年初数");
        row.getCell(6).setCellValue("上年同期");
        row.getCell(7).setCellValue("月报表取数来源");
        // row.getCell(7).setCellStyle(getStyle(wb, 0));
        row.getCell(8).setCellValue("年度报表取数来源");
        // row.getCell(8).setCellStyle(getStyle(wb, 0));
        row.getCell(9).setCellValue("指标");
        // row.getCell(9).setCellStyle(getStyle(wb, 0));

        row = sheet.createRow(4);
        for (int i = 0; i <= 9; i++) {
            row.createCell(i);
        }
        //合并单元格
        CellRangeAddress rowRegion4 = new CellRangeAddress(3, 4, 0, 0);
        CellRangeAddress rowRegion4_1 = new CellRangeAddress(3, 4, 1, 3);
        CellRangeAddress rowRegion4_2 = new CellRangeAddress(3, 4, 4, 4);
        CellRangeAddress rowRegion4_3 = new CellRangeAddress(3, 4, 5, 5);
        CellRangeAddress rowRegion4_4 = new CellRangeAddress(3, 4, 6, 6);
        CellRangeAddress rowRegion4_5 = new CellRangeAddress(3, 4, 7, 7);
        CellRangeAddress rowRegion4_6 = new CellRangeAddress(3, 4, 8, 8);
        CellRangeAddress rowRegion4_7 = new CellRangeAddress(3, 4, 9, 9);
        sheet.addMergedRegion(rowRegion4);
        sheet.addMergedRegion(rowRegion4_1);
        sheet.addMergedRegion(rowRegion4_2);
        sheet.addMergedRegion(rowRegion4_3);
        sheet.addMergedRegion(rowRegion4_4);
        sheet.addMergedRegion(rowRegion4_5);
        sheet.addMergedRegion(rowRegion4_6);
        sheet.addMergedRegion(rowRegion4_7);
        /**
         * 6-98行内容的填充
         */
        for (int i = 5; i < 103; i++) {
            row = sheet.createRow(i);
            // row.setRowStyle(getStyle(wb, 3));
            //创建所有单元格
            for (int j = 0; j <= 6; j++) {
                row.createCell(j);
            }
            for (int k = 7; k <= 9; k++) {
                row.createCell(k);
            }
        }
        //填充数据,并合并
        //6行注册资本金
        sheet.getRow(5).getCell(0).setCellValue("1");
        sheet.getRow(5).getCell(1).setCellValue("1.  注册资本金");
        //sheet.getRow(5).getCell(1).setCellStyle(getStyle(wb, 0));
        CellRangeAddress rowRegion5 = new CellRangeAddress(5, 5, 1, 3);
        sheet.addMergedRegion(rowRegion5);
        //7-10行
        sheet.getRow(6).getCell(0).setCellValue("2");//A7
        sheet.getRow(6).getCell(1).setCellValue("2.   实收资本");//B7
        //sheet.getRow(6).getCell(1).setCellStyle(getStyle(wb, 2));
        sheet.getRow(7).getCell(1).setCellValue("其中");//B8
        // sheet.getRow(7).getCell(1).setCellStyle(getStyle(wb, 2));
        sheet.getRow(7).getCell(2).setCellValue("2.1   国有资本");//C8
        sheet.getRow(8).getCell(2).setCellValue("2.2   民营资本");//C9
        sheet.getRow(9).getCell(2).setCellValue("2.3   外资");//C10
        CellRangeAddress B7_D7 = new CellRangeAddress(6, 6, 1, 3);//B7-D7
        CellRangeAddress A7_A10 = new CellRangeAddress(6, 9, 0, 0);//A7-A10
        CellRangeAddress B8_B10 = new CellRangeAddress(7, 9, 1, 1);//B8-B10
        CellRangeAddress C8_D8 = new CellRangeAddress(7, 7, 2, 3);//C8-D8
        CellRangeAddress C9_D9 = new CellRangeAddress(8, 8, 2, 3);//C9-D9
        CellRangeAddress C10_D10 = new CellRangeAddress(9, 9, 2, 3);//C10-D10
        sheet.addMergedRegion(B7_D7);
        sheet.addMergedRegion(A7_A10);
        sheet.addMergedRegion(B8_B10);
        sheet.addMergedRegion(C8_D8);
        sheet.addMergedRegion(C9_D9);
        sheet.addMergedRegion(C10_D10);
        //11-12行
        sheet.getRow(10).getCell(0).setCellValue("3");//A11
        sheet.getRow(11).getCell(0).setCellValue("4");//A12
        sheet.getRow(10).getCell(1).setCellValue("3.资产总额");//b11
        //sheet.getRow(10).getCell(1).setCellStyle(getStyle(wb, 2));
        sheet.getRow(11).getCell(1).setCellValue("4.净资产");//b12
        //sheet.getRow(11).getCell(1).setCellStyle(getStyle(wb, 2));
        CellRangeAddress B11_D11 = new CellRangeAddress(10, 10, 1, 3);//B11_D11
        CellRangeAddress B12_D12 = new CellRangeAddress(11, 11, 1, 3);//B12_D12
        sheet.addMergedRegion(B11_D11);
        sheet.addMergedRegion(B12_D12);
        //13-26行:5. 本年累计发生担保金额
        sheet.getRow(12).getCell(1).setCellValue("5. 本年累计发生担保金额");//B13 5. 本年累计发生担保金额
        sheet.getRow(12).getCell(0).setCellValue("5");//A13 5
        sheet.getRow(13).getCell(1).setCellValue("按业务类型划分");//B14 按业务类型划分
        sheet.getRow(13).getCell(2).setCellValue("5.1融资担保业务");//C14 5.1融资担保业务
        sheet.getRow(18).getCell(2).setCellValue("5.2非融担保业务");//C19 5.2非融担保业务
        sheet.getRow(13).getCell(3).setCellValue("5.1.1 借款担保业务");//D14 5.1.1 借款担保业务
        sheet.getRow(14).getCell(3).setCellValue("5.1.1.2其中：互联网担保业务");//D15 5.1.1.2其中：互联网担保业务
        sheet.getRow(15).getCell(3).setCellValue("5.1.2 发行债券担保业务");//D16 5.1.2 发行债券担保业务
        sheet.getRow(16).getCell(3).setCellValue("5.1.3 其他融资担保业务");//D17 5.1.3 其他融资担保业务
        sheet.getRow(17).getCell(3).setCellValue("5.1.4 融资担保业务小计");//D18 5.1.4 融资担保业务小计
        sheet.getRow(18).getCell(3).setCellValue("5.2.1 工程履约担保");//D19 5.2.1 工程履约担保
        sheet.getRow(19).getCell(3).setCellValue("5.2.2 诉讼保全担保");//D20 5.2.2 诉讼保全担保
        sheet.getRow(20).getCell(3).setCellValue("5.2.3 其他非融担保业务");//d21 5.2.3 其他非融担保业务
        sheet.getRow(21).getCell(3).setCellValue("5.2.4 非融担保业务小计");//d22 5.2.4 非融担保业务小计
        sheet.getRow(22).getCell(1).setCellValue("按客户类型");//B23 按客户类型
        sheet.getRow(22).getCell(2).setCellValue("5.3 涉农担保");//C23 5.3 涉农担保
        sheet.getRow(23).getCell(2).setCellValue("5.4 小型企业担保");//C24 5.4 小型企业担保
        sheet.getRow(24).getCell(2).setCellValue("5.5  微型企业担保");//C25 5.5  微型企业担保
        sheet.getRow(25).getCell(2).setCellValue("5.6 个人担保");//C26 5.6 个人担保
        sheet.getRow(22).getCell(7).setCellValue("对应指标月末数-月初数");//H23 对应指标月末数-月初数
        sheet.getRow(22).getCell(8).setCellValue("对应指标数");//I23 对应指标数
        sheet.getRow(22).getCell(9).setCellValue("本年累计涉农担保发生担保金额");//J23 本年累计涉农担保发生担保金额
        for (int i = 12; i < 26; i++) {
            sheet.getRow(i).getCell(5).setCellValue("——");//f13-f22 ——
        }
        sheet.getRow(12).getCell(7).setCellValue("各业务类型月度金额之和");//h13 各业务类型月度金额之和
        sheet.getRow(12).getCell(8).setCellValue("各业务类型年度金额之和"); //i13 各业务类型年度金额之和
        sheet.getRow(13).getCell(7).setCellValue("对应指标月末数-月初数");//h14 对应指标月末数-月初数
        sheet.getRow(13).getCell(8).setCellValue("对应指标数");//i14 对应指标数
        sheet.getRow(13).getCell(9).setCellValue("本年累计借款担保业务发生担保金额");//j14 本年累计借款担保业务发生担保金额
        CellRangeAddress A13_A22 = new CellRangeAddress(12, 25, 0, 0);//A13_A26
        CellRangeAddress B13_D13 = new CellRangeAddress(12, 12, 1, 3);//B13_D13
        CellRangeAddress B14_B22 = new CellRangeAddress(13, 21, 1, 1);//B14_B22
        CellRangeAddress C14_C18 = new CellRangeAddress(13, 17, 2, 2); //C14_C18
        CellRangeAddress C19_C22 = new CellRangeAddress(18, 21, 2, 2);//C19_C22
        CellRangeAddress B23_B26 = new CellRangeAddress(22, 25, 1, 1);//B23_B26
        CellRangeAddress C23_D23 = new CellRangeAddress(22, 22, 2, 3);//C23_D23
        CellRangeAddress C24_D24 = new CellRangeAddress(23, 23, 2, 3);//C24_D24
        CellRangeAddress C25_D25 = new CellRangeAddress(24, 24, 2, 3);//C25_D25
        CellRangeAddress C26_D26 = new CellRangeAddress(25, 25, 2, 3);//C26_D26
        sheet.addMergedRegion(A13_A22);
        sheet.addMergedRegion(B13_D13);
        sheet.addMergedRegion(B14_B22);
        sheet.addMergedRegion(C14_C18);
        sheet.addMergedRegion(C19_C22);
        sheet.addMergedRegion(B23_B26);
        sheet.addMergedRegion(C23_D23);
        sheet.addMergedRegion(C24_D24);
        sheet.addMergedRegion(C25_D25);
        sheet.addMergedRegion(C26_D26);
        //27-40行:6.本年累计发生担保户数
        sheet.getRow(26).getCell(1).setCellValue("6.本年累计发生担保户数");//B13+14 6.本年累计发生担保户数
        sheet.getRow(26).getCell(0).setCellValue("6");//A13+14 5
        sheet.getRow(27).getCell(1).setCellValue("按业务类型划分");//B14+14 按业务类型划分
        sheet.getRow(27).getCell(2).setCellValue("6.1融资担保业务");//C14+14 6.1融资担保业务
        sheet.getRow(32).getCell(2).setCellValue("6.2非融担保业务");//C19+14 6.2非融担保业务
        sheet.getRow(27).getCell(3).setCellValue("6.1.1 借款担保业务");//D14+14 6.1.1 借款担保业务
        sheet.getRow(28).getCell(3).setCellValue("6.1.1.2其中：互联网担保业务");//D15+14 6.1.1.2其中：互联网担保业务
        sheet.getRow(29).getCell(3).setCellValue("6.1.2 发行债券担保业务");//D16+14 6.1.2 发行债券担保业务
        sheet.getRow(30).getCell(3).setCellValue("6.1.3 其他融资担保业务");//D17+14 6.1.3 其他融资担保业务
        sheet.getRow(31).getCell(3).setCellValue("6.1.4 融资担保业务小计");//D18+14 6.1.4 融资担保业务小计
        sheet.getRow(32).getCell(3).setCellValue("6.2.1 工程履约担保");//D19+14 6.2.1 工程履约担保
        sheet.getRow(33).getCell(3).setCellValue("6.2.2 诉讼保全担保");//D20+14 6.2.2 诉讼保全担保
        sheet.getRow(34).getCell(3).setCellValue("6.2.3 其他非融担保业务");//d21+14 6.2.3 其他非融担保业务
        sheet.getRow(36).getCell(3).setCellValue("6.2.4 非融担保业务小计");//d22+14 6.2.4 非融担保业务小计
        sheet.getRow(36).getCell(1).setCellValue("按客户类型");//B23+14 按客户类型
        sheet.getRow(36).getCell(2).setCellValue("6.3 涉农担保");//C23+14 6.3 涉农担保
        sheet.getRow(37).getCell(2).setCellValue("6.4 小型企业担保");//C24+14 6.4 小型企业担保
        sheet.getRow(38).getCell(2).setCellValue("6.5  微型企业担保");//C25+14 6.5  微型企业担保
        sheet.getRow(39).getCell(2).setCellValue("6.6 个人担保");//C26+14 6.6 个人担保
        sheet.getRow(36).getCell(7).setCellValue("对应指标月末数-月初数");//H23+14 对应指标月末数-月初数
        sheet.getRow(36).getCell(8).setCellValue("对应指标数");//I23+14 对应指标数
        sheet.getRow(36).getCell(9).setCellValue("本年累计涉农担保发生担保金额");//J23+14 本年累计涉农担保发生担保金额
        for (int i = 26; i < 40; i++) {
            sheet.getRow(i).getCell(5).setCellValue("——");//f13+14-f22+14 ——
        }
        sheet.getRow(26).getCell(7).setCellValue("各业务类型月度金额之和");//h13+14 各业务类型月度金额之和
        sheet.getRow(26).getCell(8).setCellValue("各业务类型年度金额之和"); //i13+14 各业务类型年度金额之和
        sheet.getRow(27).getCell(7).setCellValue("对应指标月末数-月初数");//h14+14 对应指标月末数-月初数
        sheet.getRow(27).getCell(8).setCellValue("对应指标数");//i14+14 对应指标数
        sheet.getRow(27).getCell(9).setCellValue("本年累计借款担保业务发生担保金额");//j14+14 本年累计借款担保业务发生担保金额
        CellRangeAddress A27_A40 = new CellRangeAddress(26, 39, 0, 0);//A27_A40
        CellRangeAddress B27_D27 = new CellRangeAddress(26, 26, 1, 3);//B27_D27
        CellRangeAddress B28_B36 = new CellRangeAddress(27, 35, 1, 1);//B28_B36
        CellRangeAddress C28_C32 = new CellRangeAddress(27, 31, 2, 2); //C28_C32
        CellRangeAddress C33_C36 = new CellRangeAddress(32, 35, 2, 2);//C33_C36
        CellRangeAddress B37_B40 = new CellRangeAddress(36, 39, 1, 1);//B37_B40
        CellRangeAddress C37_D37 = new CellRangeAddress(36, 36, 2, 3);//C37_D37
        CellRangeAddress C38_D38 = new CellRangeAddress(37, 37, 2, 3);//C38_D38
        CellRangeAddress C39D39 = new CellRangeAddress(38, 38, 2, 3);//C39D39
        CellRangeAddress C40_D40 = new CellRangeAddress(39, 39, 2, 3);//C40_D40
        sheet.addMergedRegion(A27_A40);
        sheet.addMergedRegion(B27_D27);
        sheet.addMergedRegion(B28_B36);
        sheet.addMergedRegion(C28_C32);
        sheet.addMergedRegion(C33_C36);
        sheet.addMergedRegion(B37_B40);
        sheet.addMergedRegion(C37_D37);
        sheet.addMergedRegion(C38_D38);
        sheet.addMergedRegion(C39D39);
        sheet.addMergedRegion(C40_D40);

        //27-40行:7.在保余额
        sheet.getRow(40).getCell(1).setCellValue("7.在保余额");//B13+14+14 7.在保余额
        sheet.getRow(40).getCell(0).setCellValue("7");//A13+14+14 7
        sheet.getRow(41).getCell(1).setCellValue("按业务类型划分");//B14+14+14 按业务类型划分
        sheet.getRow(41).getCell(2).setCellValue("7.1融资担保业务");//C14+14+14 7.1融资担保业务
        sheet.getRow(46).getCell(2).setCellValue("7.2非融担保业务");//C19+14+14 7.2非融担保业务
        sheet.getRow(41).getCell(3).setCellValue("7.1.1 借款担保业务");//D14+14+14 7.1.1 借款担保业务
        sheet.getRow(42).getCell(3).setCellValue("7.1.1.2其中：互联网担保业务");//D15+14+14 7.1.1.2其中：互联网担保业务
        sheet.getRow(43).getCell(3).setCellValue("7.1.2 发行债券担保业务");//D16+14+14 7.1.2 发行债券担保业务
        sheet.getRow(44).getCell(3).setCellValue("7.1.3 其他融资担保业务");//D17+14+14 7.1.3 其他融资担保业务
        sheet.getRow(45).getCell(3).setCellValue("7.1.4 融资担保业务小计");//D18+14+14 7.1.4 融资担保业务小计
        sheet.getRow(46).getCell(3).setCellValue("7.2.1 工程履约担保");//D19+14+14 7.2.1 工程履约担保
        sheet.getRow(47).getCell(3).setCellValue("7.2.2 诉讼保全担保");//D20+14+14 7.2.2 诉讼保全担保
        sheet.getRow(48).getCell(3).setCellValue("7.2.3 其他非融担保业务");//d21+14+14 7.2.3 其他非融担保业务
        sheet.getRow(50).getCell(3).setCellValue("7.2.4 非融担保业务小计");//d22+14+14 7.2.4 非融担保业务小计
        sheet.getRow(50).getCell(1).setCellValue("按客户类型");//B23+14+14 按客户类型
        sheet.getRow(50).getCell(2).setCellValue("7.3 涉农担保");//C23+14+14 7.3 涉农担保
        sheet.getRow(51).getCell(2).setCellValue("7.4 小型企业担保");//C24+14+14 7.4 小型企业担保
        sheet.getRow(52).getCell(2).setCellValue("7.5  微型企业担保");//C25+14+14 7.5  微型企业担保
        sheet.getRow(53).getCell(2).setCellValue("7.6 个人担保");//C26+14+14 7.6 个人担保
        sheet.getRow(50).getCell(7).setCellValue("对应指标月末数-月初数");//H23+14+14 对应指标月末数-月初数
        sheet.getRow(50).getCell(8).setCellValue("对应指标数");//I23+14+14 对应指标数
        sheet.getRow(50).getCell(9).setCellValue("本年累计涉农担保发生担保金额");//J23+14+14 本年累计涉农担保发生担保金额
        for (int i = 40; i < 54; i++) {
            sheet.getRow(i).getCell(5).setCellValue("——");//f13+14+14-f22+14+14 ——
        }
        sheet.getRow(40).getCell(7).setCellValue("各业务类型月度金额之和");//h13+14+14 各业务类型月度金额之和
        sheet.getRow(40).getCell(8).setCellValue("各业务类型年度金额之和"); //i13+14+14 各业务类型年度金额之和
        sheet.getRow(41).getCell(7).setCellValue("对应指标月末数-月初数");//h14+14+14 对应指标月末数-月初数
        sheet.getRow(41).getCell(8).setCellValue("对应指标数");//i14+14+14 对应指标数
        sheet.getRow(41).getCell(9).setCellValue("本年累计借款担保业务发生担保金额");//j14+14+14 本年累计借款担保业务发生担保金额
        CellRangeAddress A41_A54 = new CellRangeAddress(40, 53, 0, 0);//A41_A54
        CellRangeAddress B41_D41 = new CellRangeAddress(40, 40, 1, 3);//B41_D41
        CellRangeAddress B42_B50 = new CellRangeAddress(41, 49, 1, 1);//B42_B50
        CellRangeAddress C42_C46 = new CellRangeAddress(41, 45, 2, 2); //C42_C46
        CellRangeAddress C47_C50 = new CellRangeAddress(46, 49, 2, 2);//C47_C50
        CellRangeAddress B51_B54 = new CellRangeAddress(50, 53, 1, 1);//B51_B54
        CellRangeAddress C51_D51 = new CellRangeAddress(50, 50, 2, 3);//C51_D51
        CellRangeAddress C52_D52 = new CellRangeAddress(51, 51, 2, 3);//C52_D52
        CellRangeAddress C53_D53 = new CellRangeAddress(52, 52, 2, 3);//C53_D53
        CellRangeAddress C54_D54 = new CellRangeAddress(53, 53, 2, 3);//C54_D54
        sheet.addMergedRegion(A41_A54);
        sheet.addMergedRegion(B41_D41);
        sheet.addMergedRegion(B42_B50);
        sheet.addMergedRegion(C42_C46);
        sheet.addMergedRegion(C47_C50);
        sheet.addMergedRegion(B51_B54);
        sheet.addMergedRegion(C51_D51);
        sheet.addMergedRegion(C52_D52);
        sheet.addMergedRegion(C53_D53);
        sheet.addMergedRegion(C54_D54);


        //56行8.融资担保责任余额（按条例及配套制度计量）
        sheet.getRow(55).getCell(0).setCellValue("8");//A56 8
        sheet.getRow(55).getCell(1).setCellValue("8.融资担保责任余额（按条例及配套制度计量）");//B56 8.融资担保责任余额（按条例及配套制度计量）
        sheet.getRow(55).getCell(7).setCellValue("各业务类型月度融资担保责任余额之和");//H56 各业务类型月度融资担保责任余额之和
        sheet.getRow(55).getCell(8).setCellValue("各业务类型年度融资担保责任余额之和");//I56 各业务类型年度融资担保责任余额之和
        sheet.getRow(55).getCell(9).setCellValue("\uF06E 借款类担保责任余额 = 单户在保余额500万元人民币以下（含500万元人民币）的小微企业借款类担保在保余额 * 75% + 单户在保余额");//J56  借款类担保责任余额 = 单户在保余额500万元人民币以下（含500万元人民币）的小微企业借款类担保在保余额 * 75% + 单户在保余额
        CellRangeAddress B56_D56 = new CellRangeAddress(55, 55, 1, 3);//B56_D56
        sheet.addMergedRegion(B56_D56);

        //57-71（14） 9.在保户数
        int startLine=56;//57行开始
        sheet.getRow(startLine).getCell(1).setCellValue("9.在保户数");//B13+14+14+16 7.在保余额
        sheet.getRow(startLine).getCell(0).setCellValue("7");//A13+14+14+16 7
        sheet.getRow(startLine+1).getCell(1).setCellValue("按业务类型划分");//B14+14+14+16 按业务类型划分
        sheet.getRow(startLine+1).getCell(2).setCellValue("9.1融资担保业务");//C14+14+14+16 7.1融资担保业务
        sheet.getRow(startLine+7).getCell(2).setCellValue("9.2非融担保业务");//C19+14+14+16 7.2非融担保业务
        sheet.getRow(startLine+1).getCell(3).setCellValue("7.1.1 借款担保业务");//D14+14+14+16 7.1.1 借款担保业务
        sheet.getRow(startLine+2).getCell(3).setCellValue("7.1.1.2其中：互联网担保业务");//D15+14+14+16 7.1.1.2其中：互联网担保业务
        sheet.getRow(startLine+3).getCell(3).setCellValue("7.1.2 发行债券担保业务");//D16+14+14+16 7.1.2 发行债券担保业务
        sheet.getRow(startLine+4).getCell(3).setCellValue("7.1.3 其他融资担保业务");//D17+14+14+16 7.1.3 其他融资担保业务
        sheet.getRow(startLine+5).getCell(3).setCellValue("9.1.3.1 其中：保本基金担保业务");
        sheet.getRow(startLine+6).getCell(3).setCellValue("7.1.4 融资担保业务小计");//D18+14+14+16 7.1.4 融资担保业务小计
        sheet.getRow(startLine+7).getCell(3).setCellValue("7.2.1 工程履约担保");//D19+14+14+16 7.2.1 工程履约担保
        sheet.getRow(startLine+8).getCell(3).setCellValue("7.2.2 诉讼保全担保");//D20+14+14+16 7.2.2 诉讼保全担保
        sheet.getRow(startLine+9).getCell(3).setCellValue("7.2.3 其他非融担保业务");//d21+14+14+16 7.2.3 其他非融担保业务
        sheet.getRow(startLine+10).getCell(3).setCellValue("7.2.4 非融担保业务小计");//d22+14+14+16 7.2.4 非融担保业务小计
        sheet.getRow(startLine+11).getCell(1).setCellValue("按客户类型");//B23+14+14+16 按客户类型
        sheet.getRow(startLine+11).getCell(2).setCellValue("7.3 涉农担保");//C23+14+14+16 7.3 涉农担保
        sheet.getRow(startLine+12).getCell(2).setCellValue("7.4 小型企业担保");//C24+14+14+16 7.4 小型企业担保
        sheet.getRow(startLine+13).getCell(2).setCellValue("7.5  微型企业担保");//C25+14+14+16 7.5  微型企业担保
        sheet.getRow(startLine+14).getCell(2).setCellValue("7.6 个人担保");//C26+14+14+16 7.6 个人担保
        sheet.getRow(startLine+11).getCell(7).setCellValue("对应指标月末数-月初数");//H23+14+14+16 对应指标月末数-月初数
        sheet.getRow(startLine+11).getCell(8).setCellValue("对应指标数");//I23+14+14+16 对应指标数
        sheet.getRow(startLine+11).getCell(8).setCellValue("对应指标数");//I23+14+14+16 对应指标数
        sheet.getRow(startLine+11).getCell(9).setCellValue("涉农担保业务在保户数");
        /*for (int i = 40; i < 54; i++) {
            sheet.getRow(i).getCell(5).setCellValue("——");
        }*/
        sheet.getRow(startLine).getCell(7).setCellValue("各业务类型月度金额之和");//h13+14+14+16 各业务类型月度金额之和
        sheet.getRow(startLine).getCell(8).setCellValue("各业务类型年度金额之和"); //i13+14+14+16 各业务类型年度金额之和
        sheet.getRow(startLine+1).getCell(7).setCellValue("对应指标月末数-月初数");//h14+14+14+16 对应指标月末数-月初数
        sheet.getRow(startLine+1).getCell(8).setCellValue("对应指标数");//i14+14+14+16 对应指标数
        sheet.getRow(startLine+1).getCell(9).setCellValue("借款担保业务在保户数");//j14+14+14+16 本年累计借款担保业务发生担保金额
        CellRangeAddress A57_A71 = new CellRangeAddress(startLine, startLine+14, 0, 0);//A57_A71
        CellRangeAddress B58_B67 = new CellRangeAddress(57, 66, 1, 1);//B58_B67
        CellRangeAddress B68_B71 = new CellRangeAddress(67, 70, 1, 1);//B68_B71
        CellRangeAddress B57_D57 = new CellRangeAddress(56, 56, 1, 3);//B57_D57
        CellRangeAddress C58_C63 = new CellRangeAddress(57, 62, 2, 2); //C58_C63
        CellRangeAddress C64_C67 = new CellRangeAddress(63, 66, 2, 2);//C64_C67
        CellRangeAddress C68_D68 = new CellRangeAddress(67, 67, 2, 3);//C68_D68
        CellRangeAddress C69_D69 = new CellRangeAddress(68, 68, 2, 3);//C69_D69
        CellRangeAddress C70_D70 = new CellRangeAddress(69, 69, 2, 3);//C70_D70
        CellRangeAddress C71_D71 = new CellRangeAddress(70, 70, 2, 3);//C71_D71
        sheet.addMergedRegion(A57_A71);
        sheet.addMergedRegion(B58_B67);
        sheet.addMergedRegion(B68_B71);
        sheet.addMergedRegion(B57_D57);
        sheet.addMergedRegion(C58_C63);
        sheet.addMergedRegion(C64_C67);
        sheet.addMergedRegion(C68_D68);
        sheet.addMergedRegion(C69_D69);
        sheet.addMergedRegion(C70_D70);
        sheet.addMergedRegion(C71_D71);
        //72-75行 损益类指标
        sheet.getRow(71).getCell(0).setCellValue("10");
        sheet.getRow(73).getCell(0).setCellValue("11");
        sheet.getRow(74).getCell(0).setCellValue("12");
        sheet.getRow(72).getCell(1).setCellValue("损益类指标");
        sheet.getRow(71).getCell(2).setCellValue("10.本年累计营业收入");
        sheet.getRow(72).getCell(2).setCellValue("其中  10.1本年累计担保费收入");
        sheet.getRow(73).getCell(2).setCellValue("11.本年累计净利润");
        sheet.getRow(74).getCell(2).setCellValue("12.本年累计缴纳税金");
        for (int i = 71; i < 75; i++) {
            sheet.getRow(i).getCell(5).setCellValue("——");
        }
        sheet.getRow(72).getCell(7).setCellValue("指标月末数-月初数");//h73  指标月末数-月初数
        sheet.getRow(72).getCell(8).setCellValue("指标月末数-月初数");//I73 指标年末数-年初数
        sheet.getRow(72).getCell(9).setCellValue("指标月末数-月初数");//J73  累计担保费收入
        CellRangeAddress a72_a73 = new CellRangeAddress(71, 72, 0, 0);//a72_a73
        CellRangeAddress B72_B75 = new CellRangeAddress(71, 74, 1, 1);//B72_B75
        CellRangeAddress C72_D72 = new CellRangeAddress(71, 71, 2, 3); //C72_D72
        CellRangeAddress C73_D73 = new CellRangeAddress(72, 72, 2, 3);//C73_D73
        CellRangeAddress C74_D74 = new CellRangeAddress(73, 73, 2, 3);//C74_D74
        CellRangeAddress C75_D75 = new CellRangeAddress(74, 74, 2, 3); //C75_D75
        sheet.addMergedRegion(a72_a73);
        sheet.addMergedRegion(B72_B75);
        sheet.addMergedRegion(C72_D72);
        sheet.addMergedRegion(C73_D73);
        sheet.addMergedRegion(C74_D74);
        sheet.addMergedRegion(C75_D75);
        //76-98行 风险类指标
        sheet.getRow(75).getCell(0).setCellValue("13");//A76 13
        sheet.getRow(77).getCell(0).setCellValue("14");//A78 14
        sheet.getRow(78).getCell(0).setCellValue("15");//A79 15
        sheet.getRow(80).getCell(0).setCellValue("16");//A81 16
        sheet.getRow(81).getCell(0).setCellValue("17");//A82 17
        sheet.getRow(83).getCell(0).setCellValue("18");//A84 18
        sheet.getRow(84).getCell(0).setCellValue("19");//A85 19
        sheet.getRow(86).getCell(0).setCellValue("20");//A87 20
        sheet.getRow(87).getCell(0).setCellValue("21");//A88 21
        sheet.getRow(91).getCell(0).setCellValue("22");//A92 22
        sheet.getRow(94).getCell(0).setCellValue("23");//A95 23
        sheet.getRow(75).getCell(1).setCellValue("风险类指标");//B76 风险类指标
        sheet.getRow(92).getCell(1).setCellValue("其中");//B93 其中
        sheet.getRow(95).getCell(1).setCellValue("其中");//B96 其中
        sheet.getRow(75).getCell(2).setCellValue("13.本年累计解除担保业务金额");//C76 13.本年累计解除担保业务金额
        sheet.getRow(76).getCell(2).setCellValue(" 其中  13.1本年累计解除融资担保业务金额");//C77  其中  13.1本年累计解除融资担保业务金额
        sheet.getRow(77).getCell(2).setCellValue("14.累计解除担保业务金额");//C78 14.累计解除担保业务金额
        sheet.getRow(78).getCell(2).setCellValue("15.本年累计担保业务代偿金额");//C79 15.本年累计担保业务代偿金额
        sheet.getRow(79).getCell(2).setCellValue("       其中    15.1本年累计融资担保代偿金额");//C80        其中    15.1本年累计融资担保代偿金额
        sheet.getRow(80).getCell(2).setCellValue("16.累计担保业务代偿金额");//C81 16.累计担保业务代偿金额
        sheet.getRow(81).getCell(2).setCellValue("17.本年累计担保业务损失金额");//C82 17.本年累计担保业务损失金额
        sheet.getRow(82).getCell(2).setCellValue(" 其中   17.1 本年累计融资担保损失金额");//C83  其中   17.1 本年累计融资担保损失金额
        sheet.getRow(83).getCell(2).setCellValue("18.累计担保损失金额");//C84 18.累计担保损失金额
        sheet.getRow(84).getCell(2).setCellValue("19.代偿余额");//C85 19.代偿余额
        sheet.getRow(85).getCell(2).setCellValue("其中   19.1融资担保代偿余额");//C86 其中   19.1融资担保代偿余额
        sheet.getRow(86).getCell(2).setCellValue("20.逾期未代偿金额");//C87 20.逾期未代偿金额
        sheet.getRow(87).getCell(2).setCellValue("21.担保准备金");//C88 21.担保准备金
        sheet.getRow(88).getCell(2).setCellValue("其中");//C89 其中
        sheet.getRow(88).getCell(3).setCellValue("21.1担保赔偿准备金");//D89 21.1担保赔偿准备金
        sheet.getRow(89).getCell(3).setCellValue("21.2未到期责任准备金");//D90 21.2未到期责任准备金
        sheet.getRow(90).getCell(3).setCellValue("21.3一般风险准备金");//D91 21.3一般风险准备金
        sheet.getRow(91).getCell(1).setCellValue("22.对外投资");//B92 22.对外投资
        sheet.getRow(92).getCell(2).setCellValue("22.1股权投资");//C93 22.1股权投资
        sheet.getRow(93).getCell(2).setCellValue("22.2 委托贷款");//C94 22.2 委托贷款
        sheet.getRow(94).getCell(1).setCellValue("23.在职人数");//B95 23.在职人数
        sheet.getRow(95).getCell(2).setCellValue("3.1研究生及以上");//C96 3.1研究生及以上
        sheet.getRow(96).getCell(2).setCellValue("23.2本科");//C97 23.2本科
        sheet.getRow(97).getCell(2).setCellValue("23.3大专及以下");//C98 23.3大专及以下

        sheet.getRow(75).getCell(5).setCellValue("——");//F76-F77 ——
        sheet.getRow(78).getCell(5).setCellValue("——");//F79-F80 ——
        sheet.getRow(81).getCell(5).setCellValue("——"); //F82-F83 ——
        sheet.getRow(76).getCell(5).setCellValue("——");//F76-F77 ——
        sheet.getRow(79).getCell(5).setCellValue("——");//F79-F80 ——
        sheet.getRow(82).getCell(5).setCellValue("——"); //F82-F83 ——
        sheet.getRow(72).getCell(7).setCellValue("指标月末数-月初数");//H73 指标月末数-月初数
        sheet.getRow(75).getCell(7).setCellValue("指标月末数-月初数");//H76 指标月末数-月初数
        sheet.getRow(76).getCell(7).setCellValue("指标月末数-月初数");//H77 指标月末数-月初数
        sheet.getRow(77).getCell(7).setCellValue("指标数");//H78 指标数
        sheet.getRow(78).getCell(7).setCellValue("指标月末数-月初数");//H79 指标月末数-月初数
        sheet.getRow(79).getCell(7).setCellValue("指标月末数-月初数");//H80 指标月末数-月初数
        sheet.getRow(80).getCell(7).setCellValue("A:指标数");//H81 A:指标数
        sheet.getRow(81).getCell(7).setCellValue("指标月末数-月初数");//H82 指标月末数-月初数
        sheet.getRow(82).getCell(7).setCellValue("指标月末数-月初数");//H83 指标月末数-月初数
        sheet.getRow(83).getCell(7).setCellValue("B：指标数");//H84 B：指标数
        sheet.getRow(84).getCell(7).setCellValue("A-B-C月末数-C月初数");//H85 A-B-C月末数-C月初数
        sheet.getRow(85).getCell(7).setCellValue("A-B-C月末数-C月初数 （融资担保类）");//H86 A-B-C月末数-C月初数 （融资担保类）
        sheet.getRow(86).getCell(7).setCellValue("指标月末数-月初数");//H87 指标月末数-月初数
        sheet.getRow(72).getCell(8).setCellValue("指标年末数-年初数");//I73 指标年末数-年初数
        sheet.getRow(75).getCell(8).setCellValue("指标年末数-年初数");//I76 指标年末数-年初数
        sheet.getRow(76).getCell(8).setCellValue("指标年末数-年初数");//I77 指标年末数-年初数
        sheet.getRow(77).getCell(8).setCellValue("指标数");//I78 指标数
        sheet.getRow(78).getCell(8).setCellValue("指标年末数-年初数");//I79 指标年末数-年初数
        sheet.getRow(79).getCell(8).setCellValue("指标年末数-年初数");//I80 指标年末数-年初数
        sheet.getRow(80).getCell(8).setCellValue("A1:指标数");//I81 A1:指标数
        sheet.getRow(81).getCell(8).setCellValue("指标年末数-年初数");//I82 指标年末数-年初数
        sheet.getRow(82).getCell(8).setCellValue("指标年末数-年初数");//I83 指标年末数-年初数
        sheet.getRow(83).getCell(8).setCellValue("B1:指标数");//I84 B1:指标数
        sheet.getRow(84).getCell(8).setCellValue("C年末数-C年初数");//I85 C年末数-C年初数
        sheet.getRow(85).getCell(8).setCellValue("C年末数-C年初数（融资担保类）");//I86 C年末数-C年初数（融资担保类）
        sheet.getRow(86).getCell(8).setCellValue("指标年末数-年初数");//I87 指标年末数-年初数
        sheet.getRow(72).getCell(9).setCellValue("累计担保费收入");//J73 累计担保费收入
        sheet.getRow(75).getCell(9).setCellValue("累计解除担保业务金额");//J76 累计解除担保业务金额
        sheet.getRow(76).getCell(9).setCellValue("累计解除融资担保业务金额");//J77 累计解除融资担保业务金额
        sheet.getRow(77).getCell(9).setCellValue("累计解除担保业务金额");//J78 累计解除担保业务金额
        sheet.getRow(78).getCell(9).setCellValue("累计担保代偿金额");//J79 累计担保代偿金额
        sheet.getRow(79).getCell(9).setCellValue("累计解除融资担保代偿金额");//J80 累计解除融资担保代偿金额
        sheet.getRow(80).getCell(9).setCellValue("累计担保代偿金额");//J81 累计担保代偿金额
        sheet.getRow(81).getCell(9).setCellValue("累计担保损失金额");//J82 累计担保损失金额
        sheet.getRow(82).getCell(9).setCellValue("累计解除融资担保损失金额");//J83 累计解除融资担保损失金额
        sheet.getRow(83).getCell(9).setCellValue("累计担保损失金额");//J84 累计担保损失金额
        sheet.getRow(84).getCell(9).setCellValue("C:累计收回担保损失金额");//J85 C:累计收回担保损失金额
        sheet.getRow(85).getCell(9).setCellValue("累计收回融资担保损失金额");//J86 累计收回融资担保损失金额
        sheet.getRow(86).getCell(9).setCellValue("逾期未代偿金额");//J87 逾期未代偿金额

        CellRangeAddress A76_A77=new CellRangeAddress(75,76,0,0);//A76_A77
        CellRangeAddress A79_A80=new CellRangeAddress(78,79,0,0);//A79_A80
        CellRangeAddress A82_A83=new CellRangeAddress(81,82,0,0);//A82_A83
        CellRangeAddress A85_A86=new CellRangeAddress(84,85,0,0);//A85_A86
        CellRangeAddress A88_A91=new CellRangeAddress(87,90,0,0);//A88_A91
        CellRangeAddress A92_A94=new CellRangeAddress(91,93,0,0);//A92_A94
        CellRangeAddress A95_A98=new CellRangeAddress(94,97,0,0);//A95_A98
        CellRangeAddress B76_B91=new CellRangeAddress(75,90,1,1);//B76_B91
        CellRangeAddress B92_D92=new CellRangeAddress(91,91,1,3);//B92_D92
        CellRangeAddress B93_B94=new CellRangeAddress(92,93,1,1);//B93_B94
        CellRangeAddress B95_D95=new CellRangeAddress(94,94,1,3);//B95_D95
        CellRangeAddress B96_B98=new CellRangeAddress(95,97,1,1);//B96_B98
        CellRangeAddress C89_C91=new CellRangeAddress(88,90,2,2);//C89_C91
        CellRangeAddress C93_D93=new CellRangeAddress(92,92,2,3);//C93_D93
        CellRangeAddress C94_D94=new CellRangeAddress(93,93,2,3);//C94_D94
        CellRangeAddress C96_D96=new CellRangeAddress(95,95,2,3);//C96_D96
        CellRangeAddress C97_D97=new CellRangeAddress(96,96,2,3);//C97_D97
        CellRangeAddress C98_D98=new CellRangeAddress(97,97,2,3);//C98_D98
        for (int i = 75; i < 88; i++) {
            CellRangeAddress cdrange=new CellRangeAddress(i,i,2,3);//C76_D76~C88_D88
            sheet.addMergedRegion(cdrange);
        }
        sheet.addMergedRegion(A76_A77);
        sheet.addMergedRegion(A79_A80);
        sheet.addMergedRegion(A82_A83);
        sheet.addMergedRegion(A85_A86);
        sheet.addMergedRegion(A88_A91);
        sheet.addMergedRegion(A92_A94);
        sheet.addMergedRegion(A95_A98);
        sheet.addMergedRegion(B76_B91);
        sheet.addMergedRegion(B92_D92);
        sheet.addMergedRegion(B93_B94);
        sheet.addMergedRegion(B95_D95);
        sheet.addMergedRegion(B96_B98);
        sheet.addMergedRegion(C89_C91);
        sheet.addMergedRegion(C93_D93);
        sheet.addMergedRegion(C94_D94);
        sheet.addMergedRegion(C96_D96);
        sheet.addMergedRegion(C97_D97);
        sheet.addMergedRegion(C98_D98);

        //99-103末尾信息

        sheet.getRow(98).getCell(2).setCellValue("审核人：");//C99  审核人：
        sheet.getRow(98).getCell(5).setCellValue("联系电话：");//F99 联系电话：
        sheet.getRow(99).getCell(0).setCellValue("填报说明：1.表格中，除“——”框中不填数据，其余均要求填写数据，若没有发生业务，填“0”；2.“涉农担保”是指融资担保法人机构为“三农”融资提供担保的业务，包含新型农业经营主体；小型企业、微型企业划分标准及统计口径按照《关于印发中小企业划型标准规定的通知》（工信部联企业[2011]300号）的有关规定执行，其中，微型企业包括微型企业、个体工商户及小微企业主；3.“融资担保责任余额”的计量见《融资担保责任余额计量办法》；4.“损失金额”是指有诉讼判决书或仲裁书和强制执行书，或者其他足以证明损失已形成的证据，证明代偿已无法收回的业务余额；5.涉及“当年累计”指标应填写本公司在本年度累计发生数据，涉及“累计”指标应填写公司成立以来发生数据；6.“逾期未代偿金额”是指担保业务已逾期公司应进行代偿而目前未代偿业务；7.涉及“担保业务收入”、“净利润”、“税金”应填列当年累计数额及去年同期累计数额。");//A100 填报说明：1.表格中，除“——”框中不填数据，其余均要求填写数据，若没有发生业务，填“0”；2.“涉农担保”是指融资担保法人机构为“三农”融资提供担保的业务，包含新型农业经营主体；小型企业、微型企业划分标准及统计口径按照《关于印发中小企业划型标准规定的通知》（工信部联企业[2011]300号）的有关规定执行，其中，微型企业包括微型企业、个体工商户及小微企业主；3.“融资担保责任余额”的计量见《融资担保责任余额计量办法》；4.“损失金额”是指有诉讼判决书或仲裁书和强制执行书，或者其他足以证明损失已形成的证据，证明代偿已无法收回的业务余额；5.涉及“当年累计”指标应填写本公司在本年度累计发生数据，涉及“累计”指标应填写公司成立以来发生数据；6.“逾期未代偿金额”是指担保业务已逾期公司应进行代偿而目前未代偿业务；7.涉及“担保业务收入”、“净利润”、“税金”应填列当年累计数额及去年同期累计数额。
        CellRangeAddress A100_G103=new CellRangeAddress(99,102,0,6);//A100-G103
        sheet.addMergedRegion(A100_G103);
        /**
         *统一样式设置
         */
        //默认行高
        sheet.setDefaultRowHeight((short) (33));
        //列宽
        sheet.setColumnWidth(0,6*256);
        sheet.setColumnWidth(1,12*256);
        sheet.setColumnWidth(2,22*256);
        sheet.setColumnWidth(3,33*256);
        sheet.setColumnWidth(4,9*256);
        sheet.setColumnWidth(5,9*256);
        sheet.setColumnWidth(6,15*256);
        sheet.setColumnWidth(7,18*256);
        sheet.setColumnWidth(8,19*256);
        sheet.setColumnWidth(9,40*256);

        //1.样式0：全局设置边框
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

        //2.样式1：方正小标宋_GBK 16 垂直居中
        sheet.getRow(0).getCell(0).setCellStyle(getStyle(wb,1));//A1
        //3.样式2：宋体16 垂直，水平居中 加粗
        sheet.getRow(1).getCell(0).setCellStyle(getStyle(wb,2));//A2
        //4.样式3：宋体12 加粗水平居左，垂直居中
        sheet.getRow(2).getCell(0).setCellStyle(getStyle(wb,3));//A3
        sheet.getRow(2).getCell(5).setCellStyle(getStyle(wb,3));//F3
        sheet.getRow(98).getCell(2).setCellStyle(getStyle(wb,3));//C99
        sheet.getRow(98).getCell(5).setCellStyle(getStyle(wb,3));//F99

        //5.样式4：宋体10 加粗 垂直水平居中
        sheet.getRow(3).getCell(7).setCellStyle(getStyle(wb,4));//H4
        sheet.getRow(3).getCell(8).setCellStyle(getStyle(wb,4));//I4
        sheet.getRow(3).getCell(9).setCellStyle(getStyle(wb,4));//J4
        //6.样式5：宋体11 加粗 垂直水平居中
        sheet.getRow(5).getCell(1).setCellStyle(getStyle(wb,5)); //B6
        sheet.getRow(5).getCell(2).setCellStyle(getStyle(wb,5)); //C6
        sheet.getRow(5).getCell(3).setCellStyle(getStyle(wb,5)); //D6

        sheet.getRow(6).getCell(1).setCellStyle(getStyle(wb,5)); //B7
        sheet.getRow(10).getCell(1).setCellStyle(getStyle(wb,5)); //B11
        sheet.getRow(11).getCell(1).setCellStyle(getStyle(wb,5)); //B12
        sheet.getRow(12).getCell(1).setCellStyle(getStyle(wb,5)); //B13
        sheet.getRow(26).getCell(1).setCellStyle(getStyle(wb,5)); //B27
        sheet.getRow(40).getCell(1).setCellStyle(getStyle(wb,5)); //B41
        sheet.getRow(91).getCell(1).setCellStyle(getStyle(wb,5)); //B92
        sheet.getRow(94).getCell(1).setCellStyle(getStyle(wb,5)); //B95
        sheet.getRow(55).getCell(1).setCellStyle(getStyle(wb,5)); //B56
        sheet.getRow(56).getCell(1).setCellStyle(getStyle(wb,5)); //B57
        sheet.getRow(71).getCell(2).setCellStyle(getStyle(wb,5)); //C72
        sheet.getRow(73).getCell(2).setCellStyle(getStyle(wb,5)); //C74
        sheet.getRow(74).getCell(2).setCellStyle(getStyle(wb,5)); //C75
        sheet.getRow(77).getCell(2).setCellStyle(getStyle(wb,5)); //C78
        sheet.getRow(78).getCell(2).setCellStyle(getStyle(wb,5)); //C79
        sheet.getRow(80).getCell(2).setCellStyle(getStyle(wb,5)); //C81
        sheet.getRow(81).getCell(2).setCellStyle(getStyle(wb,5)); //C82
        sheet.getRow(83).getCell(2).setCellStyle(getStyle(wb,5)); //C84
        sheet.getRow(84).getCell(2).setCellStyle(getStyle(wb,5)); //C85
        sheet.getRow(86).getCell(2).setCellStyle(getStyle(wb,5)); //C87
        sheet.getRow(87).getCell(2).setCellStyle(getStyle(wb,5)); //C88
        //7.样式6：宋体11 垂直居中，水平居中
        sheet.getRow(3).getCell(1).setCellStyle(getStyle(wb,6));//B4
        //E4_G98
        for (int i = 4; i < 7; i++) {
            for (int i1 = 3; i1 < 98; i1++) {
                sheet.getRow(i1).getCell(i).setCellStyle(getStyle(wb,6));
            }
        }
        //8.样式7.宋体10 垂直水平居中，自动换行
        //A4_A95
        for (int i = 3; i < 95; i++) {
            sheet.getRow(i).getCell(0).setCellStyle(getStyle(wb,7));
        }
        //H6_I103
        for (int i = 7; i < 9; i++) {
            for (int i1 =5; i1 < 103; i1++) {
                sheet.getRow(i1).getCell(i).setCellStyle(getStyle(wb,7));
            }
        }
        //9.样式8.宋体12 加粗 垂直水平合并居中，自动换行
        sheet.getRow(7).getCell(1).setCellStyle(getStyle(wb,8)); //B8
        sheet.getRow(71).getCell(1).setCellStyle(getStyle(wb,8)); //B72
        sheet.getRow(75).getCell(1).setCellStyle(getStyle(wb,8)); //B76
        sheet.getRow(92).getCell(1).setCellStyle(getStyle(wb,8)); //B93
        sheet.getRow(95).getCell(1).setCellStyle(getStyle(wb,8)); //B96
        //10.样式9.宋体12 加粗 垂直水平合并居左，自动换行
        sheet.getRow(13).getCell(1).setCellStyle(getStyle(wb,9)); //B14
        sheet.getRow(22).getCell(1).setCellStyle(getStyle(wb,9)); //B23
        sheet.getRow(27).getCell(1).setCellStyle(getStyle(wb,9)); //B28
        sheet.getRow(36).getCell(1).setCellStyle(getStyle(wb,9)); //B37
        sheet.getRow(41).getCell(1).setCellStyle(getStyle(wb,9)); //B42
        sheet.getRow(51).getCell(1).setCellStyle(getStyle(wb,9)); //B52
        sheet.getRow(57).getCell(1).setCellStyle(getStyle(wb,9)); //B58
        sheet.getRow(67).getCell(1).setCellStyle(getStyle(wb,9)); //B68

        //11.样式10，宋12 垂直中 合并左 自动换行
        sheet.getRow(99).getCell(0).setCellStyle(getStyle(wb,10));//A100
        //12.样式11，宋11 垂直中 水平中 自动换行
        sheet.getRow(13).getCell(2).setCellStyle(getStyle(wb,11));//C14
        sheet.getRow(18).getCell(2).setCellStyle(getStyle(wb,11));//C19
        sheet.getRow(27).getCell(2).setCellStyle(getStyle(wb,11));//C28
        sheet.getRow(32).getCell(2).setCellStyle(getStyle(wb,11));//C33
        sheet.getRow(41).getCell(2).setCellStyle(getStyle(wb,11));//C42
        sheet.getRow(47).getCell(2).setCellStyle(getStyle(wb,11));//C47
        sheet.getRow(57).getCell(2).setCellStyle(getStyle(wb,11));//C58
        sheet.getRow(63).getCell(2).setCellStyle(getStyle(wb,11));//C64
        sheet.getRow(88).getCell(2).setCellStyle(getStyle(wb,11));//C89
        //13.样式12，宋10 垂直居中 自动换行
        //I6_I103
        for (int i = 5; i < 103; i++) {
            sheet.getRow(i).getCell(9).setCellStyle(getStyle(wb,12));
        }
    }
    /**
     * 获取样式
     *
     * @param
     * @param styleNum
     * @return
     */
    public HSSFCellStyle getStyle(HSSFWorkbook wb, Integer styleNum) {
        if(styles.size()==0){
            for (int i = 0; i < 13; i++) {
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
            case(11):{//宋11 垂直中 水平中 自动换行
                font.setFontName("宋体");
                style.setVerticalAlignment(VerticalAlignment.CENTER);//垂直居中
                style.setAlignment(HorizontalAlignment.CENTER);
                //font.setBold(true);//粗体
                font.setFontHeightInPoints((short) 11);//字体大小
                style.setFont(font);
                style.setWrapText(true);//自动换行4
                style.setBorderTop(BorderStyle.THIN);
                style.setBorderRight(BorderStyle.THIN);
                style.setBorderLeft(BorderStyle.THIN);
                style.setBorderBottom(BorderStyle.THIN);
            }
            break;
            case(12):{//样式12，宋10 垂直居中 自动换行
                font.setFontName("宋体");
                style.setVerticalAlignment(VerticalAlignment.CENTER);//垂直居中
                //style.setAlignment(HorizontalAlignment.CENTER);
                //font.setBold(true);//粗体
                font.setFontHeightInPoints((short) 10);//字体大小
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
