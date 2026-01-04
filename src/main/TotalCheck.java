package main;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.ParseException;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import static java.util.Arrays.asList;
import static main.Constant.*;
import static main.FileHelper.copyFile;
import static main.FileHelper.folderReader;

public class TotalCheck {

    static final ArrayList<String> VACATIONS = new ArrayList<>(asList("月", "年", "其", "公"));
    static ArrayList<Attendance> verifiedList = new ArrayList<>();
    static ArrayList<Attendance> unverifiedList = new ArrayList<>();
    static List<String> HEADER;
    static Integer DAYS;


    public void startCheck() throws IOException, ParseException {
        copyFile(COMMON_PATH);
        DateHelper dateHelper = new DateHelper(COMMON_START, COMMON_END);
        HEADER = dateHelper.getDateList();
        DAYS = HEADER.size();
        TotalCheck check = new TotalCheck();
        HashMap<String, Attendance> dailyMap = check.dailyAttendance();
        HashMap<String, Attendance> totalMap = check.totalAttendance();
        check.mapCheck(dailyMap, totalMap);
        check.write();
    }

    public void write() {
        Collections.sort(unverifiedList);
        new ExcelWriter().write(HEADER, verifiedList, unverifiedList);
    }

    private void mapCheck(HashMap<String, Attendance> dailyMap, HashMap<String, Attendance> totalMap) {
        dailyMap.forEach((k, v) -> {
            if (totalMap.containsKey(k)) {
                dateSetCheck(v, totalMap.get(k));
                totalMap.remove(k);
            } else {
                unverifiedList.add(v);
            }
        });
        totalMap.forEach((k, v) -> {
            unverifiedList.add(v);
        });
    }

    private void dateSetCheck(Attendance daily, Attendance total) {
        HashSet<String> dailySet = daily.getDateSet();
        HashSet<String> totalSet = total.getDateSet();
        if (dailySet == null || totalSet == null) {
            unverifiedList.add(daily);
            unverifiedList.add(total);
            return;
        }
        if (dailySet.size() != totalSet.size()) {
            unverifiedList.add(daily);
            unverifiedList.add(total);
            return;
        }
        if (dailySet.containsAll(totalSet)) {
            verifiedList.add(daily);
            verifiedList.add(total);
        } else {
            unverifiedList.add(daily);
            unverifiedList.add(total);
        }
    }

    private HashMap<String, Attendance> totalAttendance() throws ParseException, IOException {
        String filePath = TOTAL_PATH;
        Workbook workbook = excelReader(filePath);
        HashMap<String, Attendance> map = new HashMap<>();
        Row row;
        Sheet sheet;
        List<String> dateListWithMonthList = new DateHelper(COMMON_START, COMMON_END).getDateListWithMonth();
        for (int sheetNum = 0; sheetNum < 2; sheetNum++) {
            sheet = workbook.getSheetAt(sheetNum);
            for (int rowNum = 4; rowNum <= sheet.getLastRowNum(); rowNum++) {
                row = sheet.getRow(rowNum);
                if (row == null) {
                    break;
                }
                String name = row.getCell(1).getStringCellValue().strip();
                String dept;
                String membership;
                int startNum = 5;
                if (sheetNum == 0) {
                    dept = getDept((row.getCell(3).getStringCellValue())).strip();
                    membership = "核心";
                } else {
                    dept = supportDept(row);
                    membership = "支持";
                    startNum++;
                }
                Attendance attendance = new Attendance(name, dept, membership, null, DAYS);
                for (int i = startNum; i < startNum + dateListWithMonthList.size(); i++) {
                    String att = row.getCell(i).getStringCellValue().strip();
                    if (!VACATIONS.contains(att) && !att.isEmpty()) {
                        attendance.getDateSet().add(dateListWithMonthList.get(i - startNum));
                    }
                    attendance.getDateList()[i - startNum] = att;
                }
                map.put(name + "-" + dept, attendance);
            }
        }
        return map;
    }

    private String supportDept(Row row) {
        HashMap<String, String> map = new HashMap<>();
        map.put("上飞公司", "制造中队");
        map.put("上航公司", "综合办");
        map.put("新闻中心", "综合办");
        map.put("上飞院", "工程中队");
        map.put("客服公司", "工程中队");
        map.put("北研中心", "工程中队");
        map.put("总部", "领导");
        String dept = getDept((row.getCell(4).getStringCellValue())).strip();
        if (!dept.equals("制造中队")) {
            dept = map.get((row.getCell(5).getStringCellValue()).strip());
        }
        return dept;
    }

    private HashMap<String, Attendance> dailyAttendance() throws IOException {
        File[] files = folderReader(COMMON_PATH);
        Workbook workbook;
        HashMap<String, Attendance> map = new HashMap<>();
        for (File file : files) {
            workbook = excelReader(file.getPath());
            String date = dateExtractor(file.getName());
            Row row;
            Sheet sheet;
            for (int sheetNum = 0; sheetNum < workbook.getNumberOfSheets(); sheetNum++) {
                sheet = workbook.getSheetAt(sheetNum);
                String city = String.valueOf(sheet.getSheetName().charAt(0));
                for (int rowNum = 2; rowNum < sheet.getLastRowNum(); rowNum++) {
                    row = sheet.getRow(rowNum);
                    if (row == null) {
                        break;
                    }
                    if (row.getCell(1).getStringCellValue().isBlank()) {
                        continue;
                    }
                    String name = row.getCell(1).getStringCellValue().strip();
                    String dept = getDept((row.getCell(2).getStringCellValue())).strip();
                    String hotel = getDept((row.getCell(4).getStringCellValue())).strip();
                    if (name.length() > 10) {
                        continue;
                    }
                    String key = name + "-" + dept;
                    if (map.containsKey(key)) {
                        Attendance attendance = map.get(key);
                        attendance.getDateSet().add(date);
                        int idx = ((Integer.parseInt((date.split("\\.")[1]))) - 26 + DAYS) % DAYS;
                        attendance.getDateList()[idx] = city;
                    } else {
                        map.put(key, new Attendance(name, dept, hotel, city, getMembership(sheetNum), date, DAYS));
                    }
                }
            }
        }
        return map;
    }

    private String getMembership(int sheet) {
        if (sheet % 2 == 0) {
            return "核心";
        }
        return "支持";
    }

    private Workbook excelReader(String filePath) throws IOException {
        InputStream inputStream = new FileInputStream(filePath);
        return WorkbookFactory.create(inputStream);
    }

    public String dateExtractor(String fileName) throws IOException {
        Pattern pattern = Pattern.compile("(\\d{1,2})月(\\d{1,2})日");
        Matcher matcher = pattern.matcher(fileName);

        if (matcher.find()) {
            int month = Integer.parseInt(matcher.group(1));
            int day = Integer.parseInt(matcher.group(2));
            return month + "." + day;
        } else {
            System.out.println("未匹配到日期: " + fileName);
        }
        return fileName;
    }

    private String getDept(String dept) {
        if (dept.contains("安全")) {
            return "安全质量适航";
        }
        if (dept.contains("领导")) {
            return "领导";
        }
        if (dept.contains("综合")) {
            return "综合办";
        }
        return dept;
    }

    private boolean isLeave(String status) {
        status = status.strip();
        return status.equals("离开");
    }
}
