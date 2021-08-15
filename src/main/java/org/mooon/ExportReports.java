package org.mooon;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.mooon.biz.*;
import org.mooon.util.Triple3;

import java.io.File;
import java.io.IOException;
import java.text.ParseException;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExportReports {

    public static void main(String[] args) throws ParseException, IOException {

        String path = "/Users/mooon/Documents/私人/临时/火星11";
        String filePath = path + "/省新华书店8.16-8.20早、中餐菜单.xlsx";

        Workbook workbook = WorkbookFactory.create(new File(filePath));
        exportLunchReport1(workbook, path);
        exportDinnerReport1(workbook, path);
    }

    public static void exportLunchReport1(Workbook workbook, String path) throws ParseException, IOException {
        // 导入配置
        ImportSetting importSetting = new ImportSetting();
        importSetting.sheetName = "省店中、晚餐菜单"; // 菜单所在sheet名称
        importSetting.dayStartAt = "2021-08-16";   // 起始菜单日期
        importSetting.dayAtRow = 3;                // 日期所在行号
        importSetting.dayStartCell = 'B';          // 日期起始列名
        importSetting.dayEndCell = 'K';            // 日期截止列名
        importSetting.menuStartRow = 5;            // 菜单开始行号
        importSetting.menuEndRow = 21;             // 菜单结束行号
        // 导出配置
        ExportSetting exportSetting = new ExportSetting();
        exportSetting.fileName = "4.中餐价目表(8.16-8.20).xlsx";
        exportSetting.title = "%s 中餐周菜单";    // 标题名称
        exportSetting.tileRowHigh = 50;         // 标题行高
        exportSetting.menuRowHigh = 80;         // 菜单行高
        exportSetting.menuCellWidth = 7000;     // 菜单列宽
        exportSetting.menuRowMax = 7;           //菜单最大行数
        exportSetting.menuDataCellMax = 3;      // 菜单最大列数
        exportSetting.startRow = 2;             // 数据写入起始行号

        Sheet sheet = workbook.getSheet(importSetting.sheetName);
        if (sheet == null) {
            return;
        }

        LunchService lunchService = new LunchService(importSetting, exportSetting);
        List<MealMenu> menus = lunchService.getMealMenu(sheet);
        lunchService.import2LunchMenuOfReport1(menus, path + "/" + exportSetting.fileName);
    }

    public static void exportDinnerReport1(Workbook workbook, String path) throws ParseException, IOException {
        Map<String, Triple3<Integer, Integer, Integer>> cateMap = new HashMap<>();
        int ruWeiPro = 1;
        cateMap.put("档口", new Triple3<>(27, 31, 2));
        cateMap.put("晚餐小炒", new Triple3<>(32, 36, 3));
        cateMap.put("卤味打包", new Triple3<>(37, 39, ruWeiPro));
        cateMap.put("现包主食", new Triple3<>(40, 40, 4));

        // 导入配置
        ImportSetting importSetting = new ImportSetting();
        importSetting.sheetName = "省店中、晚餐菜单"; // 菜单所在sheet名称
        importSetting.dayStartAt = "2021-08-16";   // 起始菜单日期
        importSetting.dayAtRow = 25;                // 日期所在行号
        importSetting.dayStartCell = 'B';          // 日期起始列名
        importSetting.dayEndCell = 'K';            // 日期截止列名
        importSetting.menuStartRow = 27;            // 菜单开始行号
        importSetting.menuEndRow = 40;             // 菜单结束行号
        // 导出配置
        ExportSetting exportSetting = new ExportSetting();
        exportSetting.fileName = "2.晚餐价目表(8.16-8.20).xlsx";
        exportSetting.title = "%s 晚餐周菜单";    // 标题名称
        exportSetting.tileRowHigh = 50;         // 标题行高
        exportSetting.menuRowHigh = 80;         // 菜单行高
        exportSetting.menuCellWidth = 7000;     // 菜单列宽
        exportSetting.firstCellWidth = 2000;    // 第一列列宽（中餐列别名称使用列）
        exportSetting.menuDataCellMax = 3;      // 菜单最大列数
        exportSetting.startRow = 2;             // 数据写入起始行号

        Sheet sheet = workbook.getSheet(importSetting.sheetName);
        if (sheet == null) {
            return;
        }

        // 晚餐预订单
        DinnerService dinnerService = new DinnerService(importSetting, exportSetting);
        Map<String, Map<Integer, MealMenu>> menus = dinnerService.getMealMenu(sheet, cateMap);
        dinnerService.import2MealMenuOfReport1(menus, path + "/" + exportSetting.fileName);

        // 卤味预订单
        DinnerRuWeiService ruWeiService = new DinnerRuWeiService();
        Date startDay = ExcelHelper.parseDate(importSetting.dayStartAt);
        Date endDay = ExcelHelper.addDay(startDay, menus.size());
        String fileName = String.format("1.卤味预订单(%s-%s).docx", ExcelHelper.formatDateOfTile(startDay), ExcelHelper.formatDateOfTile(endDay));
        String title = String.format("%s-%s", ExcelHelper.formatDateOfTile(startDay), ExcelHelper.formatDateOfTile(endDay));
        ruWeiService.export(path + "/" + fileName, title, menus, ruWeiPro);
    }

}
