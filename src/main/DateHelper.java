package main;

import lombok.Data;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

@Data
public class DateHelper {

    private List<String> dateList;
    private List<String> dateListWithMonth;

    DateHelper(String startTime, String endTime) throws ParseException {
        initDateList(startTime, endTime);
    }

    private void initDateList(String startTime, String endTime) throws ParseException {
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
        // 声明保存日期集合
        dateList = new ArrayList<>();
        dateListWithMonth = new ArrayList<>();
        // 转化成日期类型
        Date startDate = sdf.parse(startTime);
        Date endDate = sdf.parse(endTime);
        //用Calendar 进行日期比较判断
        Calendar calendar = Calendar.getInstance();
        while (startDate.getTime() <= endDate.getTime()) {
            // 把日期添加到集合
            dateList.add(String.valueOf(startDate.getDate()));
            dateListWithMonth.add(startDate.getMonth() + 1 + "." + startDate.getDate());
            // 设置日期
            calendar.setTime(startDate);
            //把日期增加一天
            calendar.add(Calendar.DATE, 1);
            // 获取增加后的日期
            startDate = calendar.getTime();
        }
    }
}
