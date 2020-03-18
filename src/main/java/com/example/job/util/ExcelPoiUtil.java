package com.example.job.util;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

public class ExcelPoiUtil {
    public static final String FILE_PATH = "src/main/resources/excel/";


    public static void main(String[] args) throws Exception{



        return;
    }



    /**
     * 获取模版表和目标表字段对应
     * @param dataExcel
     * @param dataNum
     * @param temExcel
     * @param temNum
     * @return
     * @throws Exception
     */
    public static Map<Integer,Integer> getNumberMapping(Map<String,Integer> manulCorrect,String dataExcel,int dataNum,String temExcel,int temNum) throws Exception {

        Map<String, Integer> dataMap = getHeaderColMap(dataExcel, dataNum);
        Map<String, Integer> temMap = getHeaderColMap(temExcel, temNum);


        Map<Integer,Integer> positionMap = new HashMap<>();
        StringBuffer sb= new StringBuffer();
        StringBuffer sb2= new StringBuffer();
        dataMap.forEach((key,value) -> {
            if (temMap.containsKey(key)){
                positionMap.put(value,temMap.get(key));
            }else{
                sb.append(key+" ");
                if (manulCorrect.containsKey(key)){
                    positionMap.put(value,manulCorrect.get(key));
                }else{
                    sb2.append(key+" ");
                }
            }
        });
        System.out.println("进行校准的字段为: "+sb.toString());
        System.out.println("其中丢弃字段为: "+sb2.toString());

        return positionMap;
    }

    private static Map<String,Integer> getHeaderColMap(String fileName,Integer num)throws Exception{
        //读取模版xlsx文件
        InputStream inpData = new FileInputStream(FILE_PATH+fileName);
        Workbook dataBook = WorkbookFactory.create(inpData);
        Sheet sheetData = dataBook.getSheetAt(0);
        Map<String,Integer> excelMap = new HashMap<>();
        Row row = sheetData.getRow(num);
        for (int i=0; i<row.getLastCellNum(); i++){
            excelMap.put(row.getCell(i).getStringCellValue(),i);
        }

        return excelMap;
    }

    /**
     * 返回句柄以对数据进行其他替换操作
     * @param numMapping
     * @throws IOException
     */
    public static Workbook exportExcel(Map<Integer, Integer> numMapping,String dataName,String temName,Integer startNum) throws IOException {
        //读取模版xlsx文件
        InputStream inp = new FileInputStream(FILE_PATH+temName);
        Workbook wbTem = WorkbookFactory.create(inp);
        Sheet sheetTem = wbTem.getSheetAt(0);

        //读取数据xlsx文件
        InputStream inpData = new FileInputStream(FILE_PATH+dataName);
        Workbook wbData = WorkbookFactory.create(inpData);


        //获取第一个工作表信息
        Sheet sheetData = wbData.getSheetAt(0);
        Row row = sheetData.getRow(2);
        //遍历数据表，创建新数据
        for (int rowNum = 1; rowNum <= sheetData.getLastRowNum(); rowNum++){
            Row rowTem = sheetTem.createRow(rowNum + startNum);
            for (int col = 0; col <= sheetData.getRow(rowNum).getLastCellNum(); col ++){
                //设置数据应该放在模版excel第几列
                if (numMapping.containsKey(col)){
                    Cell cellTem = rowTem.createCell(numMapping.get(col));
                    DataFormatter formatter = new DataFormatter();
                    cellTem.setCellValue(formatter.formatCellValue(sheetData.getRow(rowNum).getCell(col)));
                }
            }
        }
        return wbTem;
    }


}
