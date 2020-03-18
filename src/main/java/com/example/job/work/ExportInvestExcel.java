package com.example.job.work;

import com.example.job.util.ExcelPoiUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.util.StringUtils;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;

public class ExportInvestExcel {
    public static final String FILE_PATH = "src/main/resources/excel/";

    public static void main(String[] args) throws Exception{
        //导出三年滚动投资计划信息
//        exportInvestExcel("/Users/yangmengfan/Downloads/三年滚动投资计划信息表.xlsx");

//        exportInvestImplementExcel("/Users/yangmengfan/Downloads/三年滚动投资计划实施计划信息表.xlsx");
//        exportContractExcel("/Users/yangmengfan/Downloads/合同会签信息表.xlsx");
//        exportContractPayExcel("/Users/yangmengfan/Downloads/合同支付信息表.xlsx");
//        exportWorkApprovalExcel("/Users/yangmengfan/Downloads/开工审批信息表.xlsx");

        //这条数据不导入，因为被投资计划信息表触发生成
//        exportInvestRestructureExcel("/Users/yangmengfan/Downloads/投资计划调整信息表.xlsx");
//        exportInvestTenderResultExcel("/Users/yangmengfan/Downloads/招标结果信息表.xlsx");
        exportInvestTenderCaseExcel("/Users/yangmengfan/Downloads/招标采购方案信息表.xlsx");
//        exportCompleteReceiveExcel("/Users/yangmengfan/Downloads/竣工验收信息表.xlsx");
//        exportProjectManageExcel("/Users/yangmengfan/Downloads/项目建设管理进度情况信息表.xlsx");
    }

    private static void exportProjectManageExcel(String distination) throws Exception{
        Map<String,Integer> manulCorrect= new HashMap<String,Integer>(){{
            put("项目分标",0);
        }};
        Map<Integer, Integer> numberMapping = ExcelPoiUtil.getNumberMapping(manulCorrect, "investDataExcel/项目建设管理进度情况信息表.xls", 0, "investExcelTemplate/项目建设管理进度情况信息表.xlsx", 1);
        Workbook wbTem = ExcelPoiUtil.exportExcel(numberMapping,"investDataExcel/项目建设管理进度情况信息表.xls","investExcelTemplate/项目建设管理进度情况信息表.xlsx",2);
        dealProjectManageData(wbTem,distination);
    }

    private static void dealProjectManageData(Workbook wbTem, String distination) throws Exception{
        Sheet sheet = wbTem.getSheetAt(0);
        for (int r = 2; r<=sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);


        }
        OutputStream fileOut = new FileOutputStream(distination);
        wbTem.write(fileOut);
    }

    private static void exportCompleteReceiveExcel(String distination) throws Exception{
        Map<String,Integer> manulCorrect= new HashMap<String,Integer>(){{

        }};
        Map<Integer, Integer> numberMapping = ExcelPoiUtil.getNumberMapping(manulCorrect, "investDataExcel/竣工验收信息表.xls", 0, "investExcelTemplate/竣工验收信息表.xlsx", 1);
        Workbook wbTem = ExcelPoiUtil.exportExcel(numberMapping,"investDataExcel/竣工验收信息表.xls","investExcelTemplate/竣工验收信息表.xlsx",1);
        dealCompleteReceiveData(wbTem,distination);
    }

    private static void dealCompleteReceiveData(Workbook wbTem, String distination) throws Exception{
        Sheet sheet = wbTem.getSheetAt(0);
        for (int r = 2; r<=sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);
            Cell cell = row.getCell(10);
            if ("黄河水利水电开发总公司".equals(cell.getStringCellValue())){
                cell.setCellValue("开发公司(C02)");
            }
            if ("水利部小浪底水利枢纽管理中心".equals(cell.getStringCellValue())){
                cell.setCellValue("小浪底管理中心(C01)");
            }
            if ("云南合宇投资有限公司".equals(cell.getStringCellValue())){
                cell.setCellValue("云南合宇公司(C0305)");
            }

        }
        OutputStream fileOut = new FileOutputStream(distination);
        wbTem.write(fileOut);
    }

    private static void exportInvestTenderCaseExcel(String distination) throws Exception {
        Map<String,Integer> manulCorrect= new HashMap<String,Integer>(){{
            put("招标采购类型",7);
            put("招标采购方式",8);
            put("项目总投资（如有分标段情况则录入每个标段的投资）",4);
            put("项目本年投资（如有分标段情况则录入每个标段的投资）",6);
            put("录入人名称",19);
        }};
        Map<Integer, Integer> numberMapping = ExcelPoiUtil.getNumberMapping(manulCorrect, "investDataExcel/招标采购方案信息表.xls", 0, "investExcelTemplate/招标采购方案信息表.xlsx", 1);
        Workbook wbTem = ExcelPoiUtil.exportExcel(numberMapping,"investDataExcel/招标采购方案信息表.xls","investExcelTemplate/招标采购方案信息表.xlsx",1);
        dealInvestTenderCaseData(wbTem,distination);
    }

    private static void dealInvestTenderCaseData(Workbook wbTem, String distination) throws Exception{
        Sheet sheet = wbTem.getSheetAt(0);
        Set<String> set = new HashSet<String>(){{
            add("单一来源");
            add("协议采购");
            add("电子商务采购");
            add("零星采购");
            add("直接采购");
            add("续签");
            add("补充协议");
        }};

        for (int r = 2; r<=sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);
            Cell cell7 = row.getCell(7);
            Cell cell8 = row.getCell(8);
            if (!StringUtils.hasText(cell7.getStringCellValue()) && StringUtils.hasText(cell8.getStringCellValue())){
                cell8.setCellValue("");
            }
            if ("/".equals(cell7.getStringCellValue())){
                cell7.setCellValue("");
            }
            if ("/".equals(cell8.getStringCellValue())){
                cell8.setCellValue("");
            }
        }
        OutputStream fileOut = new FileOutputStream(distination);
        wbTem.write(fileOut);
    }

    private static void exportInvestTenderResultExcel(String distination) throws Exception{
        Map<String,Integer> manulCorrect= new HashMap<String,Integer>(){{
            put("异议或质疑事项及处理情况",22);
            put("中标金额（元)",20);
            put("招标采购类型",5);
            put("发中标通知书时间",13);

        }};
        Map<Integer, Integer> numberMapping = ExcelPoiUtil.getNumberMapping(manulCorrect, "investDataExcel/招标结果信息表.xls", 0, "investExcelTemplate/招标采购结果信息表.xlsx", 1);
        Workbook wbTem = ExcelPoiUtil.exportExcel(numberMapping,"investDataExcel/招标结果信息表.xls","investExcelTemplate/招标采购结果信息表.xlsx",1);
        dealInvestTenderResultData(wbTem,distination);
    }

    private static void dealInvestTenderResultData(Workbook wbTem, String distination) throws Exception{
        Sheet sheet = wbTem.getSheetAt(0);
        Set<String> set = new HashSet<String>(){{

        }};

        for (int r = 2; r<=sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);
            String method = row.getCell(4).getStringCellValue();
            String compete = row.getCell(5).getStringCellValue();

            if ("竞争性方式".equals(compete)){
                row.getCell(4).setCellValue("竞争性招标");
                row.getCell(5).setCellValue(method);
            }else if ("".equals(compete)){
                row.getCell(4).setCellValue("");
            }else if ("其他".equals(compete)){
                row.getCell(4).setCellValue("其他");
                row.getCell(5).setCellValue("");
                Cell cell6 = row.getCell(6);
                if (cell6 == null){
                    cell6 = row.createCell(6);
                }
                row.getCell(6).setCellValue(method);;
            }else{
                System.out.println(compete);
            }
        }
        OutputStream fileOut = new FileOutputStream(distination);
        wbTem.write(fileOut);
    }

    private static void exportInvestRestructureExcel(String distination) throws Exception{
        Map<String,Integer> manulCorrect= new HashMap<String,Integer>(){{
            put("公司",3);
            put("必要性和可行性",4);
            put("计划总投资（元）",10);
            put("起始年份",6);
            put("年度",7);
            put("计划类型",9);

        }};
        Map<Integer, Integer> numberMapping = ExcelPoiUtil.getNumberMapping(manulCorrect, "investDataExcel/投资计划调整信息表.xls", 0, "investExcelTemplate/投资计划调整信息表.xlsx", 1);
        Workbook wbTem = ExcelPoiUtil.exportExcel(numberMapping,"investDataExcel/投资计划调整信息表.xls","investExcelTemplate/投资计划调整信息表.xlsx",1);
        dealInvestRestructureData(wbTem,distination);
    }

    private static void dealInvestRestructureData(Workbook wbTem, String distination) throws Exception{
        Sheet sheet = wbTem.getSheetAt(0);
        for (int r = 2; r<=sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);

        }
        OutputStream fileOut = new FileOutputStream(distination);
        wbTem.write(fileOut);
    }

    private static void exportWorkApprovalExcel(String distination) throws Exception {
        Map<String,Integer> manulCorrect= new HashMap<String,Integer>(){{

        }};
        Map<Integer, Integer> numberMapping = ExcelPoiUtil.getNumberMapping(manulCorrect, "investDataExcel/开工审批信息表.xls", 0, "investExcelTemplate/开工审批信息表.xlsx", 1);
        Workbook wbTem = ExcelPoiUtil.exportExcel(numberMapping,"investDataExcel/开工审批信息表.xls","investExcelTemplate/开工审批信息表.xlsx",1);
        dealWorkApprovalExcel(wbTem,distination);
    }

    private static void dealWorkApprovalExcel(Workbook wbTem, String distination) throws Exception{
        Sheet sheet = wbTem.getSheetAt(0);

        OutputStream fileOut = new FileOutputStream(distination);
        wbTem.write(fileOut);
    }

    private static void exportContractPayExcel(String distination) throws Exception {
        Map<String,Integer> manulCorrect= new HashMap<String,Integer>(){{
            put("累计材料调差款（万元）",8);
            put("累计实际支付金额（元）",4);
            put("累计应付金额（元）",5);
            put("累计支付预付款（万元）",9);
            put("累计合同工程款（万元）",6);
            put("累计其他扣款（万元）",13);
            put("累计扣还预付款（万元）",10);
            put("累计退还质保金（万元）",12);
            put("累计扣留质保金（万元）",11);
            put("累计变更工程款（万元）",7);
        }};
        Map<Integer, Integer> numberMapping = ExcelPoiUtil.getNumberMapping(manulCorrect, "investDataExcel/合同支付信息表.xls", 0, "investExcelTemplate/合同支付信息表.xlsx", 1);
        Workbook wbTem = ExcelPoiUtil.exportExcel(numberMapping,"investDataExcel/合同支付信息表.xls","investExcelTemplate/合同支付信息表.xlsx",2);
        dealContractPayData(wbTem,distination);

    }

    private static void dealContractPayData(Workbook wbTem, String distination) throws Exception {
        Sheet sheet = wbTem.getSheetAt(0);
        for (int r = 3; r<=sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);

        }
        OutputStream fileOut = new FileOutputStream(distination);
        wbTem.write(fileOut);
    }

    private static void exportContractExcel(String distination) throws Exception {
        Map<String,Integer> manulCorrect= new HashMap<String,Integer>(){{
            put("合同开工时间",13);
            put("合同完工时间",14);
            put("合同金额(元)",9);
        }};
        Map<Integer, Integer> numberMapping = ExcelPoiUtil.getNumberMapping(manulCorrect, "investDataExcel/合同会签信息表.xls", 0, "investExcelTemplate/合同会签信息表.xlsx", 1);
        Workbook wbTem = ExcelPoiUtil.exportExcel(numberMapping,"investDataExcel/合同会签信息表.xls","investExcelTemplate/合同会签信息表.xlsx",1);
        dealContractData(wbTem,distination);
    }

    private static void dealContractData(Workbook wbTem, String distination) throws Exception{
        Sheet sheet = wbTem.getSheetAt(0);
        for (int r = 2; r<=sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);

        }
        OutputStream fileOut = new FileOutputStream(distination);
        wbTem.write(fileOut);
    }

    private static void exportInvestImplementExcel(String distination) throws Exception {
        Map<String,Integer> manulCorrect= new HashMap<String,Integer>(){{
            put("竣工（合同）验收时间",13);
            put("竣工（合同）结算时间",14);
        }};

        Map<Integer, Integer> numberMapping = ExcelPoiUtil.getNumberMapping(manulCorrect, "investDataExcel/三年滚动投资计划实施计划信息表.xls", 0, "investExcelTemplate/三年滚动投资计划实施计划信息表.xlsx", 1);
        Workbook wbTem = ExcelPoiUtil.exportExcel(numberMapping,"investDataExcel/三年滚动投资计划实施计划信息表.xls","investExcelTemplate/三年滚动投资计划实施计划信息表.xlsx",1);
        dealInvestImpData(wbTem,distination);
    }

    private static void dealInvestImpData(Workbook wbTem, String distination) throws IOException {
        Sheet sheet = wbTem.getSheetAt(0);
        for (int r = 2; r<=sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);
        }
        OutputStream fileOut = new FileOutputStream(distination);
        wbTem.write(fileOut);
    }

    private static void exportInvestExcel(String distination) throws Exception {
        Map<String,Integer> manulCorrect= new HashMap<String,Integer>(){{
            put("公司",3);
            put("必要性和可行性",4);
            put("计划总投资（元）",10);
            put("起始年份",6);
            put("年度",7);
            put("计划类型",9);
        }};
        Map<Integer, Integer> numberMapping = ExcelPoiUtil.getNumberMapping(manulCorrect,"investDataExcel/三年滚动投资计划信息表.xls", 0, "investExcelTemplate/三年滚动投资计划信息表.xlsx", 1);
        Workbook wbTem = ExcelPoiUtil.exportExcel(numberMapping,"investExcelTemplate/三年滚动投资计划信息表.xlsx","investDataExcel/三年滚动投资计划信息表.xls",1);
        //保存修改后的文件
        dealInvestData(wbTem,distination);
    }

    private static void dealInvestData(Workbook wbTem,String distination) throws IOException {
        Sheet sheet = wbTem.getSheetAt(0);
        for (int r = 2; r<=sheet.getLastRowNum(); r++){
            Row row = sheet.getRow(r);
            for (int c = 0; c <= row.getLastCellNum(); c++){

                if (c == 3){
                    if ("黄河小浪底水资源投资有限公司".equals(row.getCell(3).getStringCellValue())){
                        row.getCell(3).setCellValue("投资公司(C03)");
                    }
                    if ("水利部小浪底水利枢纽管理中心".equals(row.getCell(3).getStringCellValue())){
                        row.getCell(3).setCellValue("小浪底管理中心(C01)");
                    }
                    if ("黄河水利水电开发总公司".equals(row.getCell(3).getStringCellValue())){
                        row.getCell(3).setCellValue("开发公司(C02)");
                    }
                }
                if (c == 7){
                    String tem = row.getCell(7).getStringCellValue().replace("年度", "");
                    row.getCell(7).setCellValue(tem);
                }
                if (c == 9){
                    String tem = row.getCell(9).getStringCellValue().replace("建议计划", "建设计划");
                    row.getCell(9).setCellValue(tem);
                }
            }
        }

        OutputStream fileOut = new FileOutputStream(distination);
        wbTem.write(fileOut);
    }
}
