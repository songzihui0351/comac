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
import java.util.HashMap;

import static main.Constant.*;
import static main.FileHelper.folderReader;

public class BillCheck {


    public void start() throws IOException {
        HashMap<String, Bill> bills = generateBill();
        writeToExcel();
    }


    private HashMap<String, Bill> generateBill() throws IOException {
        File[] files = folderReader(BILL_PATH);
        Workbook workbook;
        HashMap<String, Bill> map = new HashMap<>();
        for (File file : files) {
            try (InputStream inputStream = new FileInputStream(new File(file.getPath()))) {
                workbook = WorkbookFactory.create(inputStream);
                Sheet sheet = workbook.getSheetAt(0);
                int rowNum = sheet.getLastRowNum();
                for (int i = 2; i < rowNum; i++) {
                    Row row = sheet.getRow(i);
                    if (row == null) continue;
                    String name = row.getCell(1).getStringCellValue();
                    if (name == null || name.isBlank()) continue;
                    String dept = row.getCell(2).getStringCellValue();
                    String date = row.getCell(3).getStringCellValue();
                    String remark = row.getCell(5).getStringCellValue();
                    String key = name + "-" + dept;
                    if (map.containsKey(key)) {
                        map.get(key).incrementDays(date, remark);
                    } else {
                        map.put(key, new Bill(name, dept, date));
                    }
                }
            }
        }
        return map;
    }

    public void writeToExcel() throws IOException {
        com.spire.xls.Workbook wb = new com.spire.xls.Workbook();
        Worksheet sheet = wb.getWorksheets().get(0);
        sheet.insertArray(makeMatrix(generateBill()), 1, 1);

        sheet.setName(SHEET_NAME);
        setStyle(sheet);
        //保存文档
        wb.saveToFile("output/" + SHEET_NAME + ".xlsx");
    }

    private Object[][] makeMatrix(HashMap<String, Bill> map) {
        Object[][] matrix = new Object[map.size() + 2][];
        matrix[0] = getHeader();
        int i = 1;
        int total = 0;
        for (Bill bill : map.values()) {
            Object[] row = new Object[11];
            row[0] = i;
            row[1] = bill.getDept();
            row[2] = bill.getName();
            row[3] = bill.getArrival();
            row[4] = bill.getDeparture();
            row[5] = bill.getDays();
            row[6] = 0;
            row[7] = bill.getDays();
            row[8] = PRICE;
            row[9] = bill.getDays() * PRICE;
            row[10] = "";
            matrix[i++] = row;
            total += bill.getDays() * PRICE;
        }
        Object[] row = new Object[11];
        row[8] = "总计";
        row[9] = total;
        matrix[map.size() + 1] = row;
        return matrix;
    }

    private void setStyle(Worksheet sheet) {
        //行高列宽
        CellRange global = sheet.getAllocatedRange();
        CellRange dateRange = sheet.getCellRange(2, 4, sheet.getLastRow(), 5);
        global.setRowHeight(20);
        global.setColumnWidth(5);
        for (int i = 2; i < 12; i++) {
            sheet.setColumnWidth(i, 15);
        }
        //日期
        dateRange.getCellStyle().setNumberFormat("yyyy-M-d");
        //居中
        global.getCellStyle().setHorizontalAlignment(HorizontalAlignType.Center);
        global.getCellStyle().setVerticalAlignment(VerticalAlignType.Center);
        //字体
        global.getCellStyle().getExcelFont().setFontName("宋体");
        global.getCellStyle().getExcelFont().setSize(12);
        CellRange header = sheet.getCellRange(1, 1, 1, sheet.getColumns().length);
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

}
