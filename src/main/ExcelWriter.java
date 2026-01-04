package main;

import com.spire.xls.*;

import java.awt.*;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelWriter {

    public static void main(String[] args) throws IOException {
        //创建Workbook对象
        Workbook workbook = new Workbook();

        //获取第一个工作表
        Worksheet sheet = workbook.getWorksheets().get(0);

        //设置字体名称
        sheet.getCellRange("B1").setText("字体名称：Comic Sans MS");
        sheet.getCellRange("B1").getCellStyle().getExcelFont().setFontName("Comic Sans MS");

        //设置字体大小
        sheet.getCellRange("B2").setText("字体大小：20");
        sheet.getCellRange("B2").getCellStyle().getExcelFont().setSize(20);

        //设置字体颜色
        sheet.getCellRange("B3").setText("字体颜色：蓝色");
        sheet.getCellRange("B3").getCellStyle().getExcelFont().setColor(Color.blue);

        //设置加粗
        sheet.getCellRange("B4").setText("字体样式：加粗");
        sheet.getCellRange("B4").getCellStyle().getExcelFont().isBold(true);

        //设置下划线
        sheet.getCellRange("B5").setText("字体样式：下划线");
        sheet.getCellRange("B5").getCellStyle().getExcelFont().setUnderline(FontUnderlineType.Single);

        //设置斜体
        sheet.getCellRange("B6").setText("字体样式：斜体");
        sheet.getCellRange("B6").getCellStyle().getExcelFont().isItalic(true);

        //保存文档
        workbook.saveToFile("output/FontStyles.xlsx");
    }

    public void writeToExcel(List<String> header, HashMap<String, Attendance> map) {
        HashMap<String, Attendance> core = new HashMap<>();
        HashMap<String, Attendance> support = new HashMap<>();
        for (Map.Entry<String, Attendance> entry : map.entrySet()) {
            if (isCore(entry.getValue().getCity())) {
                core.put(entry.getKey(), entry.getValue());
            } else {
                support.put(entry.getKey(), entry.getValue());
            }
        }
        write(header, core, support);
    }

    private boolean isCore(String sheetName) {
        return sheetName.contains("核心");
    }

    public void write(List<String> header, HashMap<String, Attendance> core, HashMap<String, Attendance> support) {

        com.spire.xls.Workbook wb = new Workbook();

        Worksheet coreSheet = wb.getWorksheets().get(0);
        Worksheet supportSheet = wb.getWorksheets().get(1);
        coreSheet.insertArray(makeMatrix(header, core), 1, 1);
        supportSheet.insertArray(makeMatrix(header, support), 1, 1);
        coreSheet.setName("核心");
        supportSheet.setName("支持");
        setStyle(coreSheet);
        setStyle(supportSheet);
        //保存文档
        wb.saveToFile("output/考勤统计.xlsx");
    }

    public void write(List<String> header, ArrayList<Attendance> verifiedList, ArrayList<Attendance> unverifiedList) {

        com.spire.xls.Workbook wb = new Workbook();

        Worksheet coreSheet = wb.getWorksheets().get(0);
        Worksheet supportSheet = wb.getWorksheets().get(1);
        coreSheet.insertArray(makeMatrix(header, verifiedList), 1, 1);
        supportSheet.insertArray(makeMatrix(header, unverifiedList), 1, 1);
        coreSheet.setName("已核对");
        supportSheet.setName("待核对");
        setStyle(coreSheet);
        setStyle(supportSheet);
        //保存文档
        wb.saveToFile("output/考勤统计.xlsx");
    }

    private void setStyle(Worksheet worksheet) {
        //行高列宽
        CellRange global = worksheet.getAllocatedRange();
        global.setRowHeight(20);
        global.setColumnWidth(5);
        worksheet.setColumnWidth(1, 15);
        worksheet.setColumnWidth(2, 15);
        worksheet.setColumnWidth(3, 15);
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
        //各地染色
        CellRange[] dy = worksheet.findAllString("东", false, false);
        for (CellRange cell : dy) {
            cell.getStyle().setColor(new Color(100, 222, 179));
        }
        CellRange[] nc = worksheet.findAllString("南", false, false);
        for (CellRange cell : nc) {
            cell.getStyle().setColor(new Color(250, 200, 100));
        }
        CellRange[] fp = worksheet.findAllString("富", false, false);
        for (CellRange cell : fp) {
            cell.getStyle().setColor(new Color(250, 120, 100));
        }
        CellRange[] e = worksheet.findAllString("鄂", false, false);
        for (CellRange cell : e) {
            cell.getStyle().setColor(new Color(170, 150, 200));
        }
        CellRange[] dun = worksheet.findAllString("敦", false, false);
        for (CellRange cell : dun) {
            cell.getStyle().setColor(new Color(50, 150, 200));
        }
        CellRange[] ty = worksheet.findAllString("太", false, false);
        for (CellRange cell : ty) {
            cell.getStyle().setColor(new Color(130, 180, 70));
        }
        String range = "A2" + ":C" + worksheet.getRows().length;
        worksheet.getCellRange(range).getStyle().setColor(new Color(150, 180, 250));
        //边框
        global.borderAround();
        global.borderInside();
    }


    private Object[][] makeMatrix(List<String> header, HashMap<String, Attendance> map) {
        Object[][] matrix = new Object[map.size() + 1][];
        makeHeader(header, matrix);
        int i = 1;
        for (Attendance data : map.values()) {
            Object[] row = new Object[data.getDateList().length + 3];
            row[0] = data.getHotel();
            row[1] = data.getDept();
            row[2] = data.getName();
            System.arraycopy(data.getDateList(), 0, row, 3, data.getDateList().length);
            matrix[i++] = row;
        }
        return matrix;
    }

    private Object[][] makeMatrix(List<String> header, ArrayList<Attendance> list) {
        Object[][] matrix = new Object[list.size() + 1][];
        makeHeader(header, matrix);
        int i = 1;
        for (Attendance data : list) {
            Object[] row = new Object[data.getDateList().length + 3];
            row[0] = data.getHotel();
            row[1] = data.getDept();
            row[2] = data.getName();
            System.arraycopy(data.getDateList(), 0, row, 3, data.getDateList().length);
            matrix[i++] = row;
        }
        return matrix;
    }

    private void makeHeader(List<String> header, Object[][] matrix) {
        Object[] row = new Object[header.size() + 3];
        row[0] = "住宿";
        row[1] = "中队";
        row[2] = "姓名";
        System.arraycopy(header.toArray(), 0, row, 3, header.size());
        matrix[0] = row;
    }
}
