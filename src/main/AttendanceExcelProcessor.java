package main;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.function.Function;
import java.util.stream.Collectors;

/**
 * 考勤数据批量处理程序（新增统计列最终版）
 * 严格按照您指定的Excel列读取，新增4列考勤统计，序号自动递增，所有功能完全匹配需求
 */
public class AttendanceExcelProcessor {
    // ==================== 您指定的固定列索引（Excel列从1开始，POI从0开始） ====================
    private static final int COLUMN_INDEX_NAME = 1;         // 姓名列：Excel第2列 → 索引1
    private static final int COLUMN_INDEX_TEAM = 4;         // 考勤组列：Excel第5列 → 索引4
    private static final int COLUMN_INDEX_DATE = 5;         // 考勤日期列：Excel第6列 → 索引5
    private static final int COLUMN_INDEX_ON_DUTY = 13;     // 上班打卡状态列：Excel第14列 → 索引13
    private static final int COLUMN_INDEX_OFF_DUTY = 16;    // 下班打卡状态列：Excel第17列 → 索引16
    private static final int SHEET_INDEX = 0;                // 固定读取第1个工作表（索引0）
    // ====================================================================================================

    // 基础配置常量
    private static final String INPUT_DIRECTORY = "input/attendance"; // 输入文件目录
    private static final String OUTPUT_FILE_PATH = "output/考勤结果.xlsx"; // 输出文件路径
    private static final SimpleDateFormat DATE_FORMAT = new SimpleDateFormat("yyyy-MM-dd");

    // 字体样式常量
    private static final short HEADER_FONT_SIZE = 11;
    private static final short DATA_FONT_SIZE = 11;

    // 颜色常量（符合视觉规范，柔和不刺眼）
    private static final String HEADER_FILL_COLOR = "0070C0"; // 表头蓝色背景
    private static final String NORMAL_FILL_COLOR = "D5E8D4";  // 正常-浅绿色
    private static final String MISS_CARD_FILL_COLOR = "FFF2CC"; // 缺卡-黄色
    private static final String NO_CARD_FILL_COLOR = "F8CECC";   // 无卡-浅红色
    private static final String ZEBRA_FILL_COLOR = "EBF1F8";    // 斑马纹浅蓝
    private static final String WHITE_FILL_COLOR = "FFFFFF";     // 白色背景

    public static void main(String[] args) {
        try {
            // 1. 初始化输入目录（不存在则自动创建）
            initInputDirectory();
            // 2. 批量读取目录下所有Excel文件，合并数据
            List<Employee> mergedEmployees = batchProcessFiles();
            // 3. 生成最终合并Excel文件
            generateMergedExcel(mergedEmployees);
            System.out.println("✅ 批量处理完成！");
            System.out.println("📊 共处理文件数：" + getExcelFiles().size());
            System.out.println("👥 共处理员工数：" + mergedEmployees.size());
            System.out.println("📅 共处理考勤天数：" + mergedEmployees.get(0).getDailyStatus().size());
            System.out.println("📁 结果文件已保存至：" + OUTPUT_FILE_PATH);
        } catch (Exception e) {
            System.err.println("❌ 处理失败：" + e.getMessage());
            e.printStackTrace();
        }
    }

    /**
     * 初始化输入目录，不存在则自动创建
     */
    private static void initInputDirectory() {
        File directory = new File(INPUT_DIRECTORY);
        if (!directory.exists()) {
            boolean created = directory.mkdirs();
            if (created) {
                System.out.println("📂 已自动创建输入目录：" + INPUT_DIRECTORY);
            }
        }
    }

    /**
     * 获取目录下所有Excel文件（.xlsx/.xls）
     */
    private static List<File> getExcelFiles() {
        File directory = new File(INPUT_DIRECTORY);
        File[] files = directory.listFiles((dir, name) ->
                name.toLowerCase().endsWith(".xlsx") || name.toLowerCase().endsWith(".xls")
        );
        if (files == null || files.length == 0) {
            throw new RuntimeException("输入目录 " + INPUT_DIRECTORY + " 中未找到任何Excel文件");
        }
        return Arrays.asList(files);
    }

    /**
     * 批量处理所有文件，合并员工数据
     */
    private static List<Employee> batchProcessFiles() throws IOException {
        Map<String, Employee> employeeMap = new LinkedHashMap<>();
        Set<String> allDates = new TreeSet<>();
        List<File> excelFiles = getExcelFiles();

        for (File file : excelFiles) {
            System.out.println("📖 正在处理文件：" + file.getName());
            FileProcessResult result = processSingleFile(file);
            // 合并日期
            allDates.addAll(result.getAllDates());
            // 合并员工数据
            for (Employee employee : result.getEmployees()) {
                String employeeKey = employee.getName() + "_" + employee.getTeam();
                if (employeeMap.containsKey(employeeKey)) {
                    // 已有该员工，合并考勤数据
                    employeeMap.get(employeeKey).mergeDailyStatus(employee.getDailyStatus());
                } else {
                    // 新增员工
                    employeeMap.put(employeeKey, employee);
                }
            }
        }

        // 补全所有员工的所有日期（无数据设为无卡）
        for (Employee employee : employeeMap.values()) {
            for (String date : allDates) {
                employee.getDailyStatus().putIfAbsent(date, AttendanceStatus.NO_CARD);
            }
        }

        return new ArrayList<>(employeeMap.values());
    }

    /**
     * 处理单个Excel文件，严格按您指定的列索引读取数据
     */
    private static FileProcessResult processSingleFile(File file) throws IOException {
        Map<String, Employee> employeeMap = new LinkedHashMap<>();
        Set<String> allDates = new TreeSet<>();

        try (Workbook workbook = new XSSFWorkbook(new FileInputStream(file))) {
            // 按固定索引读取工作表
            Sheet sheet = workbook.getSheetAt(SHEET_INDEX);
            if (sheet == null) {
                throw new RuntimeException("文件 " + file.getName() + " 中未找到索引为 " + SHEET_INDEX + " 的工作表");
            }

            System.out.println("  ✅ 按指定列索引读取数据，开始处理...");

            // 遍历数据行（跳过表头行，从第2行开始，索引=1）
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                // 严格按您指定的列索引读取数据
                String name = getCellValue(row.getCell(COLUMN_INDEX_NAME));
                String rawTeam = getCellValue(row.getCell(COLUMN_INDEX_TEAM));
                // 自动去除「（固定工时）」，只保留前面的中队名称
                String team = rawTeam.replace("（固定工时）", "").trim();
                String dateStr = getCellValue(row.getCell(COLUMN_INDEX_DATE));
                String onDutyStatus = getCellValue(row.getCell(COLUMN_INDEX_ON_DUTY));
                String offDutyStatus = getCellValue(row.getCell(COLUMN_INDEX_OFF_DUTY));

                // 跳过空行（姓名和日期为空则跳过）
                if (name.isEmpty() || dateStr.isEmpty()) continue;

                // 收集所有考勤日期
                allDates.add(dateStr);
                // 生成员工唯一标识（姓名+中队，避免重名）
                String employeeKey = name + "_" + team;
                Employee employee = employeeMap.computeIfAbsent(employeeKey, k -> new Employee(name, team));

                // 按您的规则判断考勤状态
                AttendanceStatus status = judgeAttendanceStatus(onDutyStatus, offDutyStatus);
                employee.getDailyStatus().put(dateStr, status);
            }
        }

        System.out.println("  ✅ 文件处理完成，共读取 " + employeeMap.size() + " 名员工数据");
        return new FileProcessResult(new ArrayList<>(employeeMap.values()), allDates);
    }

    /**
     * 考勤状态判断规则（完全匹配您的需求）
     * 正常：上下班打卡状态均为正常
     * 缺卡：仅上班或仅下班有正常打卡，另一个缺卡/空
     * 无卡：上下班均无有效打卡
     */
    private static AttendanceStatus judgeAttendanceStatus(String onDutyStatus, String offDutyStatus) {
        if ("正常".equals(onDutyStatus) && "正常".equals(offDutyStatus)) {
            return AttendanceStatus.NORMAL;
        } else if (("正常".equals(onDutyStatus) && ("缺卡".equals(offDutyStatus) || offDutyStatus.isEmpty()))
                || ("正常".equals(offDutyStatus) && ("缺卡".equals(onDutyStatus) || onDutyStatus.isEmpty()))) {
            return AttendanceStatus.MISS_CARD;
        } else {
            return AttendanceStatus.NO_CARD;
        }
    }

    /**
     * 生成最终合并的Excel文件（新增4列统计核心修改）
     */
    private static void generateMergedExcel(List<Employee> employees) throws IOException {
        if (employees.isEmpty()) {
            throw new RuntimeException("无有效考勤数据可生成");
        }

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("合并考勤结果");
            int currentRow = 0;
            final int fixedColumnCount = 3; // 固定列：序号、中队、姓名
            final int dateStartColumnIndex = fixedColumnCount;

            // 1. 生成表头行
            Row headerRow = sheet.createRow(currentRow++);
            // 1.1 写入固定表头
            String[] fixedHeaders = {"序号", "中队", "姓名"};
            for (int i = 0; i < fixedColumnCount; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(fixedHeaders[i]);
                setHeaderCellStyle(workbook, cell);
            }
            // 1.2 写入日期表头（仅保留日，去除年月）
            Set<String> fullDates = employees.get(0).getDailyStatus().keySet();
            // 完整日期转仅日的表头，同时保留完整日期的顺序
            List<String> dayHeaders = fullDates.stream()
                    .map(date -> date.split("-")[2]) // 拆分yyyy-MM-dd，取日部分
                    .map(day -> String.valueOf(Integer.parseInt(day))) // 去除前导零，01→1
                    .collect(Collectors.toList());
            int currentHeaderCol = dateStartColumnIndex;
            for (String day : dayHeaders) {
                Cell cell = headerRow.createCell(currentHeaderCol++);
                cell.setCellValue(day);
                setHeaderCellStyle(workbook, cell);
            }
            // ==================== 新增：写入统计列表头 ====================
            final int dateEndColumnIndex = currentHeaderCol - 1; // 日期列最后一列索引
            String[] statHeaders = {"正常数", "缺卡数", "无卡数", "总天数"};
            for (String statHeader : statHeaders) {
                Cell cell = headerRow.createCell(currentHeaderCol++);
                cell.setCellValue(statHeader);
                setHeaderCellStyle(workbook, cell);
            }
            final int statStartColumnIndex = dateEndColumnIndex + 1; // 统计列起始索引
            // =================================================================

            // 2. 生成数据行（序号自动递增）
            int serialNum = 1;
            for (Employee employee : employees) {
                Row dataRow = sheet.createRow(currentRow++);
                // 2.1 写入固定列数据
                dataRow.createCell(0).setCellValue(serialNum++); // 序号自动递增
                dataRow.createCell(1).setCellValue(employee.getTeam());
                dataRow.createCell(2).setCellValue(employee.getName());

                // 2.2 写入考勤数据（严格按表头顺序，确保对齐）
                int currentDataCol = dateStartColumnIndex;
                for (AttendanceStatus status : employee.getDailyStatus().values()) {
                    dataRow.createCell(currentDataCol++).setCellValue(status.getDesc());
                }

                // ==================== 新增：写入统计列数据 ====================
                // 统计该员工的考勤状态数量
                Map<AttendanceStatus, Long> statusCount = employee.getDailyStatus().values().stream()
                        .collect(Collectors.groupingBy(Function.identity(), Collectors.counting()));
                long normalCount = statusCount.getOrDefault(AttendanceStatus.NORMAL, 0L);
                long missCardCount = statusCount.getOrDefault(AttendanceStatus.MISS_CARD, 0L);
                long noCardCount = statusCount.getOrDefault(AttendanceStatus.NO_CARD, 0L);
                long totalDays = employee.getDailyStatus().size();

                // 写入统计数据
                currentDataCol = statStartColumnIndex;
                dataRow.createCell(currentDataCol++).setCellValue(normalCount);
                dataRow.createCell(currentDataCol++).setCellValue(missCardCount);
                dataRow.createCell(currentDataCol++).setCellValue(noCardCount);
                dataRow.createCell(currentDataCol++).setCellValue(totalDays);
                // =================================================================

                // 2.3 设置数据行基础样式
                setDataCellStyle(workbook, dataRow, serialNum % 2 == 0);
            }

            // 3. 设置条件格式（仅日期列，统计列不设置）
            int totalDataRows = currentRow - 1;
            setConditionalFormatting(sheet, dateStartColumnIndex, dateEndColumnIndex, 1, totalDataRows);

            // 4. 自动调整列宽（包含统计列）
            autoFitColumnWidth(sheet);
            // 5. 冻结窗格（冻结前3列和表头行，滚动查看更方便）
            sheet.createFreezePane(fixedColumnCount, 1);

            // 写入文件
            try (FileOutputStream fos = new FileOutputStream(OUTPUT_FILE_PATH)) {
                workbook.write(fos);
            }
        }
    }

    /**
     * 条件格式设置方法（完全匹配您的颜色要求）
     */
    private static void setConditionalFormatting(Sheet sheet, int startCol, int endCol, int startRow, int endRow) {
        // 边界校验，避免无效范围
        if (startCol > endCol || startRow > endRow) {
            System.out.println("⚠️ 条件格式范围无效，跳过设置");
            return;
        }

        SheetConditionalFormatting conditionalFormatting = sheet.getSheetConditionalFormatting();
        // 创建范围对象
        CellRangeAddress rangeAddress = new CellRangeAddress(startRow, endRow, startCol, endCol);
        CellRangeAddress[] rangeArray = new CellRangeAddress[]{rangeAddress};

        // 1. 正常-浅绿色填充
        ConditionalFormattingRule normalRule = conditionalFormatting.createConditionalFormattingRule(
                ComparisonOperator.EQUAL, "\"正常\""
        );
        PatternFormatting normalPattern = normalRule.createPatternFormatting();
        normalPattern.setFillBackgroundColor(new XSSFColor(java.awt.Color.decode("#" + NORMAL_FILL_COLOR), null));
        normalPattern.setFillPattern(PatternFormatting.SOLID_FOREGROUND);

        // 2. 缺卡-黄色填充
        ConditionalFormattingRule missCardRule = conditionalFormatting.createConditionalFormattingRule(
                ComparisonOperator.EQUAL, "\"缺卡\""
        );
        PatternFormatting missCardPattern = missCardRule.createPatternFormatting();
        missCardPattern.setFillBackgroundColor(new XSSFColor(java.awt.Color.decode("#" + MISS_CARD_FILL_COLOR), null));
        missCardPattern.setFillPattern(PatternFormatting.SOLID_FOREGROUND);

        // 3. 无卡-浅红色填充
        ConditionalFormattingRule noCardRule = conditionalFormatting.createConditionalFormattingRule(
                ComparisonOperator.EQUAL, "\"无卡\""
        );
        PatternFormatting noCardPattern = noCardRule.createPatternFormatting();
        noCardPattern.setFillBackgroundColor(new XSSFColor(java.awt.Color.decode("#" + NO_CARD_FILL_COLOR), null));
        noCardPattern.setFillPattern(PatternFormatting.SOLID_FOREGROUND);

        // 调用标准方法，无任何编译报错
        ConditionalFormattingRule[] ruleArray = new ConditionalFormattingRule[]{normalRule, missCardRule, noCardRule};
        conditionalFormatting.addConditionalFormatting(rangeArray, ruleArray);
    }

    /**
     * 设置表头单元格样式
     */
    private static void setHeaderCellStyle(XSSFWorkbook workbook, Cell cell) {
        XSSFCellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setFontName("微软雅黑");
        font.setFontHeightInPoints(HEADER_FONT_SIZE);
        font.setBold(true);
        font.setColor(IndexedColors.WHITE.getIndex());
        style.setFont(font);

        XSSFColor headerFillColor = new XSSFColor(java.awt.Color.decode("#" + HEADER_FILL_COLOR), null);
        style.setFillForegroundColor(headerFillColor);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);

        cell.setCellStyle(style);
    }

    /**
     * 设置数据行基础样式
     */
    private static void setDataCellStyle(XSSFWorkbook workbook, Row row, boolean isZebra) {
        XSSFCellStyle baseStyle = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setFontName("微软雅黑");
        font.setFontHeightInPoints(DATA_FONT_SIZE);
        baseStyle.setFont(font);

        String fillColorHex = isZebra ? ZEBRA_FILL_COLOR : WHITE_FILL_COLOR;
        XSSFColor fillColor = new XSSFColor(java.awt.Color.decode("#" + fillColorHex), null);
        baseStyle.setFillForegroundColor(fillColor);
        baseStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        baseStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        baseStyle.setBorderBottom(BorderStyle.THIN);
        baseStyle.setBorderTop(BorderStyle.THIN);
        baseStyle.setBorderLeft(BorderStyle.THIN);
        baseStyle.setBorderRight(BorderStyle.THIN);

        // 固定列左对齐
        XSSFCellStyle leftAlignStyle = workbook.createCellStyle();
        leftAlignStyle.cloneStyleFrom(baseStyle);
        leftAlignStyle.setAlignment(HorizontalAlignment.LEFT);
        for (int i = 0; i < 3; i++) {
            Cell cell = row.getCell(i);
            if (cell != null) {
                cell.setCellStyle(leftAlignStyle);
            }
        }

        // 考勤列和统计列居中对齐
        XSSFCellStyle centerAlignStyle = workbook.createCellStyle();
        centerAlignStyle.cloneStyleFrom(baseStyle);
        centerAlignStyle.setAlignment(HorizontalAlignment.CENTER);
        for (int i = 3; i < row.getLastCellNum(); i++) {
            Cell cell = row.getCell(i);
            if (cell != null) {
                cell.setCellStyle(centerAlignStyle);
            }
        }
    }

    /**
     * 自动调整列宽
     */
    private static void autoFitColumnWidth(Sheet sheet) {
        int maxColNum = sheet.getRow(0).getLastCellNum();
        for (int colIndex = 0; colIndex < maxColNum; colIndex++) {
            int maxWidth = 0;
            for (int rowIndex = 0; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                if (row == null) continue;
                Cell cell = row.getCell(colIndex);
                if (cell == null) continue;
                String cellValue = getCellValue(cell);
                int currentWidth = cellValue.chars().map(c -> c > 0xFF ? 2 : 1).sum();
                if (currentWidth > maxWidth) {
                    maxWidth = currentWidth;
                }
            }
            int finalWidth = Math.max(8, Math.min(maxWidth + 3, 50));
            sheet.setColumnWidth(colIndex, finalWidth * 256);
        }
    }

    /**
     * 获取单元格字符串值
     */
    private static String getCellValue(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue().trim();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return DATE_FORMAT.format(cell.getDateCellValue());
                } else {
                    return String.valueOf((int) Math.round(cell.getNumericCellValue()));
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return getCellValue(cell.getCachedFormulaResultType(), cell);
            default:
                return "";
        }
    }

    /**
     * 处理公式单元格值
     */
    private static String getCellValue(CellType cellType, Cell cell) {
        if (cellType == CellType.STRING) {
            return cell.getStringCellValue().trim();
        } else if (cellType == CellType.NUMERIC) {
            if (DateUtil.isCellDateFormatted(cell)) {
                return DATE_FORMAT.format(cell.getDateCellValue());
            } else {
                return String.valueOf((int) Math.round(cell.getNumericCellValue()));
            }
        } else if (cellType == CellType.BOOLEAN) {
            return String.valueOf(cell.getBooleanCellValue());
        }
        return "";
    }

    /**
     * 考勤状态枚举
     */
    private enum AttendanceStatus {
        NORMAL("正常"),
        MISS_CARD("缺卡"),
        NO_CARD("无卡");

        private final String desc;

        AttendanceStatus(String desc) {
            this.desc = desc;
        }

        public String getDesc() {
            return desc;
        }
    }

    /**
     * 员工信息实体（支持多文件数据合并）
     */
    private static class Employee {
        private final String name;
        private final String team;
        // key: 完整日期(yyyy-MM-dd)，value: 考勤状态
        private final Map<String, AttendanceStatus> dailyStatus = new TreeMap<>();

        public Employee(String name, String team) {
            this.name = name;
            this.team = team;
        }

        public String getName() {
            return name;
        }

        public String getTeam() {
            return team;
        }

        public Map<String, AttendanceStatus> getDailyStatus() {
            return dailyStatus;
        }

        // 合并其他文件的同员工考勤数据
        public void mergeDailyStatus(Map<String, AttendanceStatus> otherStatus) {
            otherStatus.forEach((date, status) -> {
                // 已有数据不覆盖，无数据则新增
                this.dailyStatus.putIfAbsent(date, status);
            });
        }

        @Override
        public boolean equals(Object o) {
            if (this == o) return true;
            if (o == null || getClass() != o.getClass()) return false;
            Employee employee = (Employee) o;
            return Objects.equals(name, employee.name) && Objects.equals(team, employee.team);
        }

        @Override
        public int hashCode() {
            return Objects.hash(name, team);
        }
    }

    /**
     * 单文件处理结果封装
     */
    private static class FileProcessResult {
        private final List<Employee> employees;
        private final Set<String> allDates;

        public FileProcessResult(List<Employee> employees, Set<String> allDates) {
            this.employees = employees;
            this.allDates = allDates;
        }

        public List<Employee> getEmployees() {
            return employees;
        }

        public Set<String> getAllDates() {
            return allDates;
        }
    }
}