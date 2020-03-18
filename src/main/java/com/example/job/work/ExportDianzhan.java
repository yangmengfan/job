package com.example.job.work;

import com.example.job.util.DateUtil;
import com.example.job.util.ExcelUtil;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

public class ExportDianzhan {

    public static void main(String[] args) {
        importRiData();
    }

    public static void importRiData(){
        List<String> duringDate = DateUtil.getDuringDate("2020-02-24", "2020-12-31");
        for (String str:
             duringDate) {
            System.out.println(str);
        }
    }

    public static void exportDianzhan(){
        Map<String,String> dianzhanMap = new HashMap<>();

        //读取电站主字段
        List<List<String>> dangan = ExcelUtil.readXLS("参控股电站档案", 4);
        for (List<String> list: dangan){
            if (list.size() != 10){
                throw new RuntimeException("档案"+list.get(0));
            }
            String key =list.get(3)+list.get(7)+list.get(8);
            if (key.equals("古里卡河20201")){
                System.out.println(list.get(0));
            }
            dianzhanMap.put(key,list.get(0));
        }

        List<List<String>> ribao = ExcelUtil.readXLS("参控股电站日报", 3);
        List<List<String>> newRibao = new ArrayList<>();
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm");
        for (List<String> list: ribao){
            Date date = null;

            try {
                date = sdf.parse(list.get(3));
            } catch (ParseException e) {
                e.printStackTrace();
            }
            Integer year = date.getYear()+1900;
            Integer month = date.getMonth()+1;

            String key = list.get(5)+year+month;
            if (key.equals("古里卡河20201")){
                System.out.println(list.get(0));
            }
            List<String> listri = new ArrayList<>();
            listri.add(dianzhanMap.get(key));
            listri.add(list.get(5));
            listri.add(list.get(3));
            listri.add(list.get(6));
            newRibao.add(listri);
        }

        System.out.println(newRibao);
        ExcelUtil.insertLine(newRibao);

    }
}
