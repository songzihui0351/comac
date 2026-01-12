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
import java.util.HashMap;
import java.util.List;

import static main.Constant.*;
import static main.FileHelper.folderReader;

public class DYCheck {

    List<String> HEADER;
    Integer DAYS;

    public void start() throws ParseException, IOException {
        DateHelper dateHelper = new DateHelper(START, END);
        HEADER = dateHelper.getDateList();
        DAYS = HEADER.size();

        HashMap<String, Attendance> core = check(CORE);
        HashMap<String, Attendance> support = check(SUPPORT);
        ExcelWriter writer = new ExcelWriter();
        writer.write(HEADER, core, support);
    }

    private HashMap<String, Attendance> check(int sheetNum) throws IOException {
        File[] files = folderReader(PATH);
        Workbook workbook;
        HashMap<String, Attendance> map = new HashMap<>();
        DAYS = files.length;
        for (int i = 0; i < DAYS; i++) {
            File file = files[i];
            workbook = excelReader(file.getPath());
            Row row;
            Sheet sheet = workbook.getSheetAt(sheetNum);
            for (int j = 2; j < sheet.getLastRowNum(); j++) {
                row = sheet.getRow(j);
                if (row == null) {
                    break;
                }
                if (row.getCell(1) == null) {
                    System.out.println(file.getName());
                }
                if (row.getCell(1).getStringCellValue().isEmpty()) {
                    continue;
                }
                String name = (row.getCell(1).getStringCellValue()).strip();
                String dept = (row.getCell(2).getStringCellValue()).strip();
                String hotel = (row.getCell(4).getStringCellValue()).strip();
                Attendance attendance;
                if (map.containsKey(name)) {
                    attendance = map.get(name);
                    attendance.getDateList()[i] = "1";
                } else {
                    attendance = new Attendance(name, dept, hotel);
                    String[] dateList = new String[DAYS];
                    dateList[i] = "1";
                    attendance.setDateList(dateList);
                    map.put(name, attendance);
                }
            }
            workbook.close();
        }
        return map;
    }

    private Workbook excelReader(String filePath) throws IOException {
        InputStream inputStream = new FileInputStream(filePath);
        return WorkbookFactory.create(inputStream);
    }

}
