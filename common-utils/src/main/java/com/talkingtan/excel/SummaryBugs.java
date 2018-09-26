package com.talkingtan.excel;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.sl.usermodel.ColorStyle;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import java.awt.*;
import java.awt.Color;
import java.io.File;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.List;

/**
 * 统计bug修复率
 */
public class SummaryBugs {
    public static String FILE_LOCATION = "F:/temp/";

    public static String TOTAL_XLS = FILE_LOCATION + "summary";
    public static String NEW_XLS = FILE_LOCATION + "new.xls";
    public static String FIX_XLS = FILE_LOCATION + "fixed.xls";

    public static Map<String, String> NAMES_MAP = new HashMap<>();
    public static List<String> NAMES_LIST = new ArrayList<>();

    public static Map<String, StaffBugBean> notFixedStaffBugMap = new HashMap<String, StaffBugBean>();
    public static Map<String, StaffBugBean> fixedStaffBugMap = new HashMap<String, StaffBugBean>();


    static {
        SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd");
        Calendar c = Calendar.getInstance();
        TOTAL_XLS = TOTAL_XLS + "-" + format.format(c.getTime()) + ".xls";

        NAMES_LIST.add("liangzhiwei");
        NAMES_LIST.add("liupan");
        NAMES_LIST.add("liudajun");
        NAMES_LIST.add("huzongyong");
        NAMES_LIST.add("fengtao");

        NAMES_LIST.add("lizihao");
        NAMES_LIST.add("pengyi");
        NAMES_LIST.add("xuezhigang");

        NAMES_LIST.add("duyumin");
        NAMES_LIST.add("linhuaxin");
        NAMES_LIST.add("zhuanghuanbin");
        NAMES_LIST.add("yiyujie");

        NAMES_LIST.add("donghongping");
        NAMES_LIST.add("yezhihao");
        NAMES_LIST.add("liyuwen");
        NAMES_LIST.add("yangmengfei");

        NAMES_LIST.add("keweimeng");
        NAMES_LIST.add("huangkunting");
        NAMES_LIST.add("liangbingkun");
        NAMES_LIST.add("chenkai");
        NAMES_LIST.add("zhanchangru");

        NAMES_MAP.put("liangzhiwei", "梁志伟");
        NAMES_MAP.put("liupan", "刘攀");
        NAMES_MAP.put("liudajun", "刘大军");
        NAMES_MAP.put("huzongyong", "胡宗勇");
        NAMES_MAP.put("fengtao", "冯涛");

        NAMES_MAP.put("lizihao", "李子豪");
        NAMES_MAP.put("pengyi", "彭毅");
        NAMES_MAP.put("xuezhigang", "薛志刚");

        NAMES_MAP.put("linhuaxin", "林华新");
        NAMES_MAP.put("duyumin", "杜玉敏");
        NAMES_MAP.put("zhuanghuanbin", "庄焕滨");
        NAMES_MAP.put("yiyujie", "衣玉杰");

        NAMES_MAP.put("donghongping", "董红苹");
        NAMES_MAP.put("yezhihao", "叶志豪");
        NAMES_MAP.put("liyuwen", "李玉文");
        NAMES_MAP.put("yangmengfei", "杨梦飞");

        NAMES_MAP.put("huangkunting", "黄坤庭");
        NAMES_MAP.put("liangbingkun", "梁炳坤");
        NAMES_MAP.put("keweimeng", "柯伟梦");
        NAMES_MAP.put("chenkai", "陈凯");
        NAMES_MAP.put("zhanchangru", "詹昌如");
    }

    //new or open or reopen
    public static void readBugNumber(String fileName, int type) throws Exception {
        XSSFWorkbook wb = (XSSFWorkbook) WorkbookFactory.create(new File(fileName));

        XSSFSheet sheet = wb.getSheetAt(0);
        int rowNumber = sheet.getLastRowNum() - sheet.getFirstRowNum();
        //从第二行开始
        for (int i = 1; i < rowNumber; i++) {
            XSSFRow row = sheet.getRow(i);

            StaffBugBean sb = new StaffBugBean();
            sb.setName(String.valueOf(row.getCell(0)));
            sb.setLow(StringUtils.isEmpty(String.valueOf(row.getCell(1))) || "null".equals(String.valueOf(row.getCell(1))) ? 0 : row.getCell(1).getNumericCellValue());
            sb.setMedium(StringUtils.isEmpty(String.valueOf(row.getCell(2))) || "null".equals(String.valueOf(row.getCell(2))) ? 0 : row.getCell(2).getNumericCellValue());
            sb.setHigh(StringUtils.isEmpty(String.valueOf(row.getCell(3))) || "null".equals(String.valueOf(row.getCell(3))) ? 0 : row.getCell(3).getNumericCellValue());
            sb.setTotal(StringUtils.isEmpty(String.valueOf(row.getCell(4))) || "null".equals(String.valueOf(row.getCell(4))) ? 0 : row.getCell(4).getNumericCellValue());

            if (0 == type) {
                notFixedStaffBugMap.put(sb.getName(), sb);
            } else {
                fixedStaffBugMap.put(sb.getName(), sb);
            }
        }

    }

    //生成汇总excel
    public static void generalTotalExcel() throws Exception {
        Workbook wb = new XSSFWorkbook();
        XSSFSheet sheet = (XSSFSheet) wb.createSheet("bug汇总统计");


        XSSFRow header = sheet.createRow(0);
        XSSFRow title = sheet.createRow(1);

        createCell(wb, header, 0, "姓名");
        createCell(wb, header, 1, "未修复");
        createCell(wb, title, 1, "low");
        createCell(wb, title, 2, "medium");
        createCell(wb, title, 3, "high");
        createCell(wb, title, 4, "合计");
        createCell(wb, header, 5, "已修复");
        createCell(wb, title, 5, "low");
        createCell(wb, title, 6, "medium");
        createCell(wb, title, 7, "high");
        createCell(wb, title, 8, "合计");
        createCell(wb, header, 9, "bug修复率");
        createCell(wb, header, 10, "排名");

        sheet.addMergedRegion(new CellRangeAddress(0, 1, 0, 0));
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 1, 4));
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 5, 8));
        sheet.addMergedRegion(new CellRangeAddress(0, 1, 9, 9));
        sheet.addMergedRegion(new CellRangeAddress(0, 1, 10, 10));

        int totalRowNum = NAMES_LIST.size();
        for (int i = 0; i < totalRowNum; i++) {
            XSSFRow row = sheet.createRow(i + 2);
            String name = NAMES_LIST.get(i);
            StaffBugBean notFixedStaff = notFixedStaffBugMap.get(name);
            StaffBugBean fixedStaff = fixedStaffBugMap.get(name);

            createCell(wb, row, 0, NAMES_MAP.get(name));
            createCellNumber(wb, row, 1, notFixedStaff == null ? 0 : notFixedStaff.getLow());
            createCellNumber(wb, row, 2, notFixedStaff == null ? 0 : notFixedStaff.getMedium());
            createCellNumber(wb, row, 3, notFixedStaff == null ? 0 : notFixedStaff.getHigh());
            createCellNumber(wb, row, 4, notFixedStaff == null ? 0 : notFixedStaff.getTotal());
            createCellNumber(wb, row, 5, fixedStaff == null ? 0 : fixedStaff.getLow());
            createCellNumber(wb, row, 6, fixedStaff == null ? 0 : fixedStaff.getMedium());
            createCellNumber(wb, row, 7, fixedStaff == null ? 0 : fixedStaff.getHigh());
            createCellNumber(wb, row, 8, fixedStaff == null ? 0 : fixedStaff.getTotal());

            double total = (notFixedStaff == null ? 0 : notFixedStaff.getTotal()) + (fixedStaff == null ? 0 : fixedStaff.getTotal());

            createCellNumber(wb, row, 9, fixedStaff == null ? 0 : (total == 0 ? 0 : (fixedStaff.getTotal() / total)));
            createCellNumber(wb, row, 10, 0);
        }


        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream(TOTAL_XLS);
        wb.write(fileOut);
        fileOut.close();
    }

    /**
     * Creates a cell and aligns it a certain way.
     *
     * @param wb     the workbook
     * @param row    the row to create the cell in
     * @param column the column number to create the cell in
     */
    private static void createCell(Workbook wb, Row row, int column, String value) {
        Cell cell = row.createCell(column);
        CellStyle cellStyle = wb.createCellStyle();

        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        Font font = wb.createFont();
        font.setColor(IndexedColors.WHITE.getIndex());
        font.setBold(true);

        if(row.getRowNum() <= 1) {
            switch (column) {
                case 0:
                    cellStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
                    cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                    break;
                case 1:
                    cellStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
                    cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                    break;
                case 2:
                    cellStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
                    cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                    break;
                case 3:
                    cellStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
                    cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                    break;
                case 4:
                    cellStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
                    cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                    break;
                case 5:
                    cellStyle.setFillForegroundColor(IndexedColors.GREEN.getIndex());
                    cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                    break;
                case 6:
                    cellStyle.setFillForegroundColor(IndexedColors.GREEN.getIndex());
                    cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                    break;
                case 7:
                    cellStyle.setFillForegroundColor(IndexedColors.GREEN.getIndex());
                    cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                    break;
                case 8:
                    cellStyle.setFillForegroundColor(IndexedColors.GREEN.getIndex());
                    cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                    break;
                case 9:
                    cellStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
                    cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                    break;
                case 10:
                    cellStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
                    cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                    break;
            }

            cellStyle.setFont(font);
        }




        cell.setCellStyle(cellStyle);
        cell.setCellValue(value);
    }

    /**
     * Creates a cell and aligns it a certain way.
     *
     * @param wb     the workbook
     * @param row    the row to create the cell in
     * @param column the column number to create the cell in
     */
    private static void createCellNumber(Workbook wb, Row row, int column, double value) {
        Cell cell = row.createCell(column);
        CellStyle cellStyle = wb.createCellStyle();

        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        cell.setCellStyle(cellStyle);
        cell.setCellValue(value);

        //bug修复率
        if (10 == column && row.getRowNum() == 2) {
            cell.setCellType(CellType.FORMULA);
            cell.setCellFormula("RANK(J3,$J$3:$J$19)");
        }
    }

    public static void main(String[] args) throws Exception {
        readBugNumber(NEW_XLS, 0);
        readBugNumber(FIX_XLS, 1);
        generalTotalExcel();

        System.out.println("success");
    }
}
