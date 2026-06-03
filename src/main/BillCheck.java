package main;


import com.spire.xls.CellRange;
import com.spire.xls.HorizontalAlignType;
import com.spire.xls.VerticalAlignType;
import com.spire.xls.Worksheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.awt.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import static main.Constant.*;
import static main.FileHelper.folderReader;

public class BillCheck {


    public void start() throws IOException {
        ArrayList<Bill> bills = generateBill();
        for (String name : HOTEL_NAME) {
            writeToExcel(bills, name);
        }
    }

    private ArrayList<Bill> generateBill() throws IOException {
        File[] files = folderReader(BILL_PATH);
        Workbook workbook;
        HashMap<String, Bill> map = new HashMap<>();
        ArrayList<Bill> bills = new ArrayList<>();
        for (File file : files) {
            try (InputStream inputStream = new FileInputStream(file.getPath())) {
                workbook = WorkbookFactory.create(inputStream);
                Sheet sheet = workbook.getSheetAt(0);
                int rowNum = sheet.getLastRowNum();

                for (int i = 2; i < rowNum; i++) {
                    Row row = sheet.getRow(i);
                    if (row == null) continue;
                    String name = row.getCell(1).getStringCellValue();
                    if (name == null || name.isBlank()) continue;
                    String dept = getDept(row.getCell(2).getStringCellValue());
                    String date = row.getCell(3).getStringCellValue();
                    String lastDate = "";
                    Pattern pattern = Pattern.compile("\\d{1,2}月(\\d{1,2})日");
                    Matcher matcher = pattern.matcher(file.getName());
                    if (matcher.find()) {
                        lastDate = String.format("%02d", Integer.parseInt(matcher.group(1)));
                    }
                    String hotel = row.getCell(4).getStringCellValue();
                    String remark = row.getCell(5).getStringCellValue();
                    String key = name + "-" + dept + "-" + hotel;
                    if (remark.equals("入住")) {
                        Bill bill = new Bill(name, dept, date, hotel);
                        if (lastDate.equals("25")) {
                            bill.setLastDay("25日在住");
                            bill.setDays(0);
                        }
                        map.put(key, bill);
                    } else if (remark.equals("离开")) {
                        Bill bill = map.get(key);
                        if (bill == null) {
                            continue;
                        }
                        bill.incrementDays(date, remark);
                        bills.add(bill);
                        map.remove(key);
                    } else if (map.containsKey(key)) {
                        Bill bill = map.get(key);
                        if (lastDate.equals("25")) {
                            bill.setLastDay("25日在住");
                            remark = bill.getLastDay();
                        }
                        bill.incrementDays(date, remark);
                    } else {
                        for (String k : map.keySet()) {
                            if (k.startsWith(name + "-" + dept)) {
                                Bill bill = map.get(k);
                                if (bill == null) {
                                    continue;
                                }
                                bill.incrementDays(date, remark);
                                bills.add(bill);
                                map.remove(k);
                                break;
                            }
                        }
                        map.put(key, new Bill(name, dept, date, hotel));
                    }
                }
            }
        }
        bills.addAll(map.values());
        return bills;
    }

    public void writeToExcel(ArrayList<Bill> bills, String sheetName) throws IOException {
        com.spire.xls.Workbook wb = new com.spire.xls.Workbook();
        Worksheet sheet = wb.getWorksheets().get(0);
        List<Bill> list = bills.stream().filter(bill -> bill.getHotel().equals(sheetName)).toList();
        sheet.insertArray(makeMatrix(list), 1, 1);
        setStyle(sheet);
        //保存文档
        wb.saveToFile("output/" + sheetName + ".xlsx");
    }

    private Object[][] makeMatrix(List<Bill> list) {
        Object[][] matrix = new Object[list.size() + 2][];
        matrix[0] = getHeader();
        int i = 1;
        int total = 0;
        for (Bill bill : list) {
            Object[] row = new Object[11];
            int totalDays = 0;
            row[0] = i;
            row[1] = bill.getDept();
            row[2] = bill.getName();
            row[3] = bill.getArrival();
            row[4] = bill.getDeparture();
            row[5] = bill.getDays();
            row[6] = bill.getLastDay() == null ? 0 : 1;
            totalDays = (int) row[5] + (int) row[6];
            row[7] = totalDays;
            row[8] = PRICE;
            row[9] = totalDays * PRICE;
            row[10] = bill.getLastDay();
            total += totalDays * PRICE;
            matrix[i++] = row;
        }
        Object[] row = new Object[11];
        row[8] = "总计";
        row[9] = total;
        matrix[list.size() + 1] = row;
        return matrix;
    }

    private void setStyle(Worksheet worksheet) {
        //行高列宽
        CellRange global = worksheet.getAllocatedRange();
        global.setRowHeight(20);
        global.setColumnWidth(5);
        for (int i = 2; i < 12; i++) {
            worksheet.setColumnWidth(i, 15);
        }
        //居中
        global.getCellStyle().setHorizontalAlignment(HorizontalAlignType.Center);
        global.getCellStyle().setVerticalAlignment(VerticalAlignType.Center);
        //字体
        global.getCellStyle().getExcelFont().setFontName("宋体");
        global.getCellStyle().getExcelFont().setSize(12);
        CellRange header = worksheet.getCellRange(1, 1, 1, worksheet.getColumns().length);
        header.getCellStyle().getExcelFont().isBold(true);
        header.getCellStyle().getExcelFont().setSize(14);
        //背景颜色
        global.getStyle().setColor(new Color(230, 170, 130));
        global.getStyle().setColor(new Color(255, 255, 255));
        header.getStyle().setColor(new Color(200, 200, 200));
        //边框
        global.borderAround();
        global.borderInside();
    }

    private Object[] getHeader() {
        Object[] header = new Object[11];
        header[0] = "序号";
        header[1] = "所属中队";
        header[2] = "姓名";
        header[3] = "入住时间";
        header[4] = "离店时间";
        header[5] = "住宿天数";
        header[6] = "天数调整";
        header[7] = "调整后天数";
        header[8] = "单价（元）";
        header[9] = "金额（元）";
        header[10] = "备注";
        return header;
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

}
