package com.example.job.util;

import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

public class DateUtil {
    public static void main(String[] args) {

    }
    public static List<String> getDuringDate(String startStr,String endStr){
        List<String> list = new ArrayList<>();
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
        try {
            //起始日期
            Date start = sdf.parse(startStr);
            //结束日期
            Date end = sdf.parse(endStr);
            Date temp = start;
            Calendar calendar = Calendar.getInstance();
            calendar.setTime(start);

            while (temp.getTime() < end.getTime()) {
                temp = calendar.getTime();
                list.add(sdf.format(temp));
                //天数+1
                calendar.add(Calendar.DAY_OF_MONTH, 1);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }

        return list;
    }
}
