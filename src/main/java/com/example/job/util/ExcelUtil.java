package com.example.job.util;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Number;
import jxl.write.*;
import org.junit.Test;

import java.io.File;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;

/**
 * @Auther: ymfa
 * @Date: 2020/2/14 13:07
 * @Description:
 */
public class ExcelUtil {
    public static final String FILE_PATH = "src/main/resources/excel/";
    public static void main(String[] args) throws Exception{
//        readXLS("test.xls");
        BigDecimal bd = new BigDecimal("1.200853632E7");
        String str = bd.toPlainString();
        System.out.println(str);
    }

    // 创建表格
    public static void createXLS() {
        try {
            // 打开文件
            WritableWorkbook book = Workbook.createWorkbook(new File("test.xls"));
            // 生成名为“第一页”的工作表，参数0表示这是第一页
            WritableSheet sheet = book.createSheet("first page", 0);
            // 在Label对象的构造子中指名单元格位置是第一列第一行(0,0)单元格内容为test
            Label label = new Label(0, 0, "test");
            // 将定义好的单元格添加到工作表中
            sheet.addCell(label);
            // 生成一个保存数字的单元格,
            Number number = new Number(1, 0, 789);
            sheet.addCell(number);
            // 写入数据并关闭文件
            book.write();
            book.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // 读取Excel
    public static List<List<String>> readXLS(String pathName,Integer hangshu) {
        List<List<String>> lists = new ArrayList<>();

        try {
            Workbook book = Workbook.getWorkbook(new File(FILE_PATH+pathName+".xls"));
            Sheet sheet = book.getSheet(0);
            Cell cell = sheet.getCell(0, 0);
            for (int row = hangshu; row<sheet.getRows(); row++ ){
                List rowLine = new ArrayList();
                for (int col = 0; col<sheet.getColumns(); col++){
                    String contents = sheet.getCell(col, row).getContents();
                    System.out.print(contents+" ");
                    rowLine.add(contents);
                }
                System.out.println();
                lists.add(rowLine);
            }

            book.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return lists;
    }

    // 修改Excel的类，添加一个工作表
    public static void insertLine(List<List<String>> lists) {
        try {
            File file = new File(FILE_PATH+"newRI.xls");
            Workbook wb = Workbook.getWorkbook(file);
            // 打开一个文件的副本，并且指定数据写回到原文件
            WritableWorkbook book = Workbook.createWorkbook(new File("newRI.xls"), wb);
            WritableSheet sheet = book.getSheet(0);

            int rows = sheet.getRows();
            int column = 0;
            for (List<String> list: lists){
                for (String str : list){
                    Label label = new Label(column,rows,str);
                    sheet.addCell(label);
                    column++;
                }
                rows++;
                column=0;
            }

            book.write();
            book.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // 操作图片
    public static void addImage() throws Exception {
        WritableWorkbook wwb = Workbook.createWorkbook(new File("test.xls"));
        WritableSheet ws = wwb.createSheet("Test Sheet 1", 0);
        File file = new File("test.png");
        WritableImage image = new WritableImage(1, 4, 6, 18, file);
        ws.addImage(image);
        wwb.write();
        wwb.close();
    }

    // 单元格合并
    public static void mergeSheet(WritableWorkbook book, int sheetIndex, int x, int y, int x1, int y1) {
        WritableSheet sheet = book.getSheet(sheetIndex);
        try {
            sheet.mergeCells(x, y, x1, y1);
            // 合并第x列第x行到第x1列第y1行的所有单元格
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // 列宽行高设置
    public static void setColView(WritableSheet sheet, int colIndex, int rowIndex, int colW, int rowH) {
        try {
            // 将第colIndex列的宽度设为colW
            sheet.setColumnView(colIndex, colW);
            // 将第rowIndex行的高度设为rowH
            sheet.setRowView(rowIndex, rowH);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void setFont() {
        WritableWorkbook book;
        try {
            book = Workbook.createWorkbook(new File("test.xls"));
            WritableFont font1 = new WritableFont(WritableFont.TIMES, 16, WritableFont.BOLD);// 字体为TIMES，字号16，加粗显示
            WritableCellFormat format1 = new WritableCellFormat(font1);
            // 使用了WritableCellFormat类，这个类非常重要，通过它可以指定单元格的各种属性，后面的单元格格式化中会有更多描述
            Label label = new Label(0, 0, "data 4 test", format1);
            // 使用了Label类的构造子，指定了字串被赋予那种格式
            // 把水平对齐方式指定为居中
            format1.setAlignment(jxl.format.Alignment.CENTRE);
            // 把垂直对齐方式指定为居中
            format1.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
            // 设置自动换行
            format1.setWrap(true);
        } catch (Exception e) {
            e.printStackTrace();
        }

    }
}
