package org.mooon.biz;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.mooon.ExportSetting;
import org.mooon.ImportSetting;

import java.io.*;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

import static org.mooon.biz.ExcelHelper.DATE_FORMAT;
import static org.mooon.biz.ExcelHelper.DATE_FORMAT_OF_TITLE;

public class LunchService {

    private final ImportSetting importSetting;
    private final ExportSetting exportSetting;
    public LunchService(ImportSetting importSetting, ExportSetting exportSetting) {
        this.importSetting = importSetting;
        this.exportSetting = exportSetting;
    }

    public List<MealMenu> getMealMenu(Sheet sheet) throws ParseException {
        // 解析每日菜单日期
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat(DATE_FORMAT);
        Date startDay = simpleDateFormat.parse(importSetting.dayStartAt);
        List<DayMenuContext> contexts = new LinkedList<>();
        Row row = sheet.getRow(importSetting.dayAtRow - 1);
        int start = 0;
        for (int col = importSetting.dayStartCell - 'A'; col < importSetting.dayEndCell - 'A' + 1; col++) {
            String value = row.getCell(col).getStringCellValue();
            if (value == null || value.trim().isEmpty()) {
                value = row.getCell(col + 1).getStringCellValue();
            }
            Calendar calendar = Calendar.getInstance();
            calendar.setTime(startDay);
            calendar.add(Calendar.DAY_OF_MONTH, start);
            DayMenuContext context = new DayMenuContext();
            context.day = value;
            context.dayOfDate = calendar.getTime();
            context.nameColNum = col;
            context.priceColNum = col + 1;
            contexts.add(context);
            col++;
            start++;
        }

        // 获取每日详细菜单
        List<MealMenu> menus = new LinkedList<>();
        for (DayMenuContext context : contexts) {
            MealMenu menu = new MealMenu(context.day, context.dayOfDate);
            for (int i = importSetting.menuStartRow - 1; i < importSetting.menuEndRow; i++) {
                Row rowI = sheet.getRow(i);
                String name = rowI.getCell(context.nameColNum).getStringCellValue().trim();
                String price = null;
                try {
                    double v = rowI.getCell(context.priceColNum).getNumericCellValue();
                    price = ExcelHelper.formatPrice(v);
                } catch (Exception e) {
                    //
                    price = rowI.getCell(context.priceColNum).getStringCellValue();
                }
                menu.addMenu(name, price);
            }
            menus.add(menu);
        }
        menus.forEach(System.out::println);

        return menus;
    }

    public void import2LunchMenuOfReport1(List<MealMenu> menus, String filePath) throws IOException {
        SimpleDateFormat dfOfSheet = new SimpleDateFormat(ExcelHelper.DATE_FORMAT_OF_SHEET);
        Workbook workbook = new XSSFWorkbook();
        for (MealMenu menu : menus) {
            Sheet sheet = workbook.createSheet(dfOfSheet.format(menu.getForDayOfDate()));
            // 创建标题
            createTitle(workbook, sheet, menu);
            // 创建菜单
            createData(workbook, sheet, menu);
        }
        try(FileOutputStream fos = new FileOutputStream(filePath)) {
            workbook.write(fos);
            fos.flush();
        } catch (Exception e) {
            throw new RuntimeException(e);
        } finally {
            workbook.close();
        }
    }

    public void createTitle(Workbook workbook, Sheet sheet, MealMenu menu) {
        SimpleDateFormat df = new SimpleDateFormat(DATE_FORMAT_OF_TITLE);
        String title = String.format(exportSetting.title, df.format(menu.getForDayOfDate()));
        Row rowTitle = sheet.createRow(0);
        rowTitle.setHeightInPoints(exportSetting.tileRowHigh);
        Cell cellTitle = rowTitle.createCell(0, CellType.STRING);
        // 设置样式
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setWrapText(true);
        // 设置字体
        Font font = workbook.createFont();
        font.setBold(true);
        font.setFontName(exportSetting.font);
        font.setFontHeightInPoints(exportSetting.tileFontSize);
        style.setFont(font);

        // 合并单元格
        CellRangeAddress region = new CellRangeAddress(0, 0, 0, exportSetting.menuDataCellMax - 1);
        sheet.addMergedRegion(region);
        cellTitle.setCellStyle(style);
        cellTitle.setCellValue(title);
    }

    public void createData(Workbook workbook, Sheet sheet, MealMenu menu) {
        int startRow = exportSetting.startRow - 1;
        int startCell = exportSetting.startCell - 'A';
        int totalRow = exportSetting.menuRowMax + 1;
        int totalCell = exportSetting.menuDataCellMax;
        for (int i = startRow; i < totalRow; i++) {
            Row rowMenu = sheet.createRow(i);
            rowMenu.setHeightInPoints(exportSetting.menuRowHigh);
        }
        int index = 0;
        for (int j = startCell; j < totalCell; j++) {
            // 创建菜单列
            // 设置样式
            sheet.setColumnWidth(j, exportSetting.menuCellWidth);
            CellStyle style = workbook.createCellStyle();
            style.setAlignment(HorizontalAlignment.CENTER);
            style.setVerticalAlignment(VerticalAlignment.CENTER);
            style.setWrapText(true);
            style.setBorderTop(BorderStyle.THIN);
            style.setBorderBottom(BorderStyle.THIN);
            style.setBorderLeft(BorderStyle.THIN);
            style.setBorderRight(BorderStyle.THIN);
            // 设置字体
            Font font = workbook.createFont();
            font.setBold(true);
            font.setFontName(exportSetting.font);
            font.setFontHeightInPoints(exportSetting.menuFontSize);
            style.setFont(font);
            for (int i = startRow; i < totalRow; i++) {
                System.out.printf("create row:%d,cell:%d%n", i,  j);
                Row row = sheet.getRow( i);
                Cell cell = row.createCell(startCell + j, CellType.STRING);
                cell.setCellStyle(style);
                if (index == menu.menuList().size()) {
                    // 追加米饭列
                    cell.setCellValue("米饭\n1元");
                } else if (index < menu.menuList().size()) {
                    // 创建每日详细菜单列
                    MealMenu.Node m = menu.menuList().get(index);
                    cell.setCellValue(String.format("%s\n%s", m.getName(), m.getPrice()));
                } else {
                    cell.setCellValue("");
                }
                index++;
            }
        }
    }

}
