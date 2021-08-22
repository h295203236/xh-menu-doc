package org.mooon;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.mooon.biz.*;
import org.mooon.util.Triple3;

import java.io.IOException;
import java.security.InvalidParameterException;
import java.text.ParseException;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExportReports {

    public static void main(String[] args) throws ParseException, IOException {

        String path = "/Users/mooon/Documents/私人/临时/火星13";
        String sourceExcelFileName = "省店8.23-8.27早、中餐菜单.xlsx";
        String sheetName = "省店中晚菜单"; // 菜单所在sheet名称
        String dayStartAt = "2021-08-23";   // 起始菜单日期
        String dayEndAt = "2021-08-27";   // 起始菜单日期
        exportAllReports(path, sourceExcelFileName, sheetName, dayStartAt, dayEndAt);
    }

    public static void exportAllReports(String path, String sourceExcelFileName,
                                        String sheetName, String dayStartAt, String dayEndAt) throws ParseException, IOException {

        // 午餐参数
        ImportSetting importSettingOfLunch = new ImportSetting();
        importSettingOfLunch.sheetName = sheetName; // 菜单所在sheet名称
        importSettingOfLunch.dayStartAt = dayStartAt;   // 起始菜单日期
        importSettingOfLunch.dayAtRow = 3;                // 日期所在行号
        importSettingOfLunch.menuStartRow = 5;            // 菜单开始行号
        importSettingOfLunch.menuEndRow = 20;             // 菜单结束行号
        importSettingOfLunch.dayStartCell = 'B';          // 日期起始列名
        importSettingOfLunch.dayEndCell = 'K';            // 日期截止列名

        // 晚餐参数
        Map<String, Triple3<Integer, Integer, Integer>> cateMap = new HashMap<>();
        // tripe3 起始行号、截止行号、优先级
        int lu = 1;
        cateMap.put("档口", new Triple3<>(25, 29, 2));
        cateMap.put("晚餐小炒", new Triple3<>(30, 34, 3));
        cateMap.put("卤味打包", new Triple3<>(35, 37, lu));
        cateMap.put("现包主食", new Triple3<>(38, 38, 4));

        // 导入配置
        ImportSetting importSettingOfDinner = new ImportSetting();
        importSettingOfDinner.sheetName = sheetName;
        importSettingOfDinner.dayAtRow = 23;                // 日期所在行号
        importSettingOfDinner.menuStartRow = 25;            // 菜单开始行号
        importSettingOfDinner.menuEndRow = 38;             // 菜单结束行号
        importSettingOfDinner.dayStartCell = 'B';          // 日期起始列名
        importSettingOfDinner.dayEndCell = 'K';            // 日期截止列名

        Date startDay = ExcelHelper.parseDate(dayStartAt);
        Date endDay = ExcelHelper.parseDate(dayEndAt);
        String dateOfSheet = String.format("%s-%s", ExcelHelper.formatDateOfSheet(startDay), ExcelHelper.formatDateOfSheet(endDay));
        String dateOfTitle = String.format("%s-%s", ExcelHelper.formatDateOfTile(startDay), ExcelHelper.formatDateOfTile(endDay));

        // workbook
        Workbook workbook = new XSSFWorkbook(path + "/" + sourceExcelFileName);

        // 午餐报告导出
        String filenameOfLunch = String.format("4.中餐价目表(%s).xlsx", dateOfSheet);
        exportLunchReport1(workbook, path, filenameOfLunch, importSettingOfLunch);

        // 晚餐报告导出
        String filenameOfDinner = String.format("2.晚餐价目表(%s).xlsx", dateOfSheet);
        exportDinnerReport1(workbook, path, filenameOfDinner, importSettingOfDinner, cateMap, lu);

        // 分开菜单报告导出
        // menuStartRow-3: 晚餐起始行号，用于中晚餐拆分报告
        String filenameOfSplitMenu = String.format("5.中晚餐分开菜单(%s).xlsx", dateOfSheet);
        exportSplitMenuReport1(path, sourceExcelFileName, sheetName, filenameOfSplitMenu,
                importSettingOfDinner.menuStartRow - 3);

        workbook.close();
    }

    protected static void exportLunchReport1(Workbook workbook, String path, String filename, ImportSetting importSetting) throws ParseException, IOException {
        // 导入配置
        //importSetting.sheetName = "省店中晚菜单"; // 菜单所在sheet名称
        //importSetting.dayStartAt = "2021-08-23";   // 起始菜单日期
        //importSetting.dayAtRow = 3;                // 日期所在行号
        //importSetting.dayStartCell = 'B';          // 日期起始列名
        //importSetting.dayEndCell = 'K';            // 日期截止列名
        //importSetting.menuStartRow = 5;            // 菜单开始行号
        //importSetting.menuEndRow = 20;             // 菜单结束行号
        // 导出配置
        ExportSetting exportSetting = new ExportSetting();
        exportSetting.fileName = filename;
        exportSetting.title = "%s 中餐周菜单";    // 标题名称
        exportSetting.tileRowHigh = 50;         // 标题行高
        exportSetting.menuRowHigh = 80;         // 菜单行高
        exportSetting.menuCellWidth = 7000;     // 菜单列宽
        exportSetting.menuRowMax = 7;           //菜单最大行数
        exportSetting.menuDataCellMax = 3;      // 菜单最大列数
        exportSetting.startRow = 2;             // 数据写入起始行号

        Sheet sheet = workbook.getSheet(importSetting.sheetName);
        if (sheet == null) {
            throw new InvalidParameterException(String.format("该sheet:{%s}不存在", importSetting.sheetName));
        }

        LunchService lunchService = new LunchService(importSetting, exportSetting);
        List<MealMenu> menus = lunchService.getMealMenu(sheet);
        lunchService.import2LunchMenuOfReport1(menus, path + "/" + exportSetting.fileName);
    }

    protected static void exportDinnerReport1(Workbook workbook, String path,
                                              String filenameOfDinner,
                                              ImportSetting importSetting,
                                              Map<String, Triple3<Integer, Integer, Integer>> cateMap,
                                              int lu) throws ParseException, IOException {
        //Map<String, Triple3<Integer, Integer, Integer>> cateMap = new HashMap<>();
        //int ruWeiPro = 1;
        //cateMap.put("档口", new Triple3<>(25, 29, 2));
        //cateMap.put("晚餐小炒", new Triple3<>(30, 34, 3));
        //cateMap.put("卤味打包", new Triple3<>(35, 37, ruWeiPro));
        //cateMap.put("现包主食", new Triple3<>(38, 38, 4));

        // 导入配置
        //importSetting.sheetName = "省店中晚菜单"; // 菜单所在sheet名称
        //importSetting.dayStartAt = "2021-08-23";   // 起始菜单日期
        //importSetting.dayAtRow = 23;                // 日期所在行号
        //importSetting.dayStartCell = 'B';          // 日期起始列名
        //importSetting.dayEndCell = 'K';            // 日期截止列名
        //importSetting.menuStartRow = 25;            // 菜单开始行号
        //importSetting.menuEndRow = 38;             // 菜单结束行号
        // 导出配置
        ExportSetting exportSetting = new ExportSetting();
        exportSetting.fileName = filenameOfDinner;
        exportSetting.title = "%s 晚餐周菜单";    // 标题名称
        exportSetting.tileRowHigh = 50;         // 标题行高
        exportSetting.menuRowHigh = 80;         // 菜单行高
        exportSetting.menuCellWidth = 7000;     // 菜单列宽
        exportSetting.firstCellWidth = 2000;    // 第一列列宽（中餐列别名称使用列）
        exportSetting.menuDataCellMax = 3;      // 菜单最大列数
        exportSetting.startRow = 2;             // 数据写入起始行号

        Sheet sheet = workbook.getSheet(importSetting.sheetName);
        if (sheet == null) {
            throw new InvalidParameterException(String.format("该sheet:{%s}不存在", importSetting.sheetName));
        }

        // 晚餐菜单
        DinnerService dinnerService = new DinnerService(importSetting, exportSetting);
        Map<Date, Map<Integer, MealMenu>> menus = dinnerService.getMealMenu(sheet, cateMap);
        dinnerService.import2MealMenuOfReport1(menus, path + "/" + exportSetting.fileName);

        // 卤味预订单
        Date startDay = ExcelHelper.parseDate(importSetting.dayStartAt);
        Date endDay = ExcelHelper.addDay(startDay, menus.size() - 1);
        String dateOfSheet = String.format("%s-%s", ExcelHelper.formatDateOfSheet(startDay), ExcelHelper.formatDateOfSheet(endDay));
        String dateOfTitle = String.format("%s-%s", ExcelHelper.formatDateOfTile(startDay), ExcelHelper.formatDateOfTile(endDay));

        String filenameOfLuWei = String.format("1.卤味预订单(%s).docx", dateOfSheet);
        DinnerRuWeiService ruWeiService = new DinnerRuWeiService();
        ruWeiService.export(path + "/" + filenameOfLuWei, dateOfTitle, menus, lu);

        // 晚餐预订单
        String fileNameOf3 = String.format("3.晚餐预订单(%s).xlsx", dateOfSheet);
        DinnerOrdersService ordersService = new DinnerOrdersService(path, "/" + fileNameOf3);
        ordersService.exportReport1(menus);
    }

    protected static void exportSplitMenuReport1(String path,
                                                 String sourceFileName,
                                                 String sourceSheet,
                                                 String fileName,
                                                 int dinnerStartAt) {
        SplitMenuService splitMenuService = new SplitMenuService(path, sourceFileName, sourceSheet, fileName);
        splitMenuService.exportReport1(dinnerStartAt);
    }
}
