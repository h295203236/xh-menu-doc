package org.mooon.biz;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.mooon.ExportSetting;
import org.mooon.ImportSetting;
import org.mooon.util.Triple3;
import org.mooon.util.Tuple2;

import java.io.FileOutputStream;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

import static org.mooon.biz.ExcelHelper.*;

public class DinnerService {

    private final ImportSetting importSetting;
    private final ExportSetting exportSetting;
    public DinnerService(ImportSetting importSetting, ExportSetting exportSetting) {
        this.importSetting = importSetting;
        this.exportSetting = exportSetting;
    }

    public Map<String, Map<Integer, MealMenu>> getMealMenu(Sheet sheet, Map<String, Triple3<Integer, Integer, Integer>> cateMap) throws ParseException {
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
        Map<String, Map<Integer, MealMenu>> menusMap = new LinkedHashMap<>();
        for (DayMenuContext context : contexts) {
            Map<Integer, MealMenu> menuMap = menusMap.computeIfAbsent(context.day, s -> new LinkedHashMap<>());
            for (int i = importSetting.menuStartRow - 1; i < importSetting.menuEndRow; i++) {
                Tuple2<String, Integer> tuple2 = getCate(i, cateMap);
                Row rowI = sheet.getRow(i);
                String name = rowI.getCell(context.nameColNum).getStringCellValue().trim();
                String price = null;
                try {
                    double v = rowI.getCell(context.priceColNum).getNumericCellValue();
                    price = formatPrice(v);
                } catch (Exception e) {
                    //
                    price = rowI.getCell(context.priceColNum).getStringCellValue();
                }
                MealMenu menu = menuMap.computeIfAbsent(tuple2.getSecond(), s -> new MealMenu(context.day, context.dayOfDate, tuple2.getFirst(), tuple2.getSecond()));
                menu.addMenu(name, price);
            }
        }

        return menusMap;
    }

    public void import2MealMenuOfReport1(Map<String, Map<Integer, MealMenu>> menus, String filePath) throws IOException {
        SimpleDateFormat dfOfSheet = new SimpleDateFormat(DATE_FORMAT_OF_SHEET);
        Workbook workbook = new XSSFWorkbook();
        for (Map<Integer, MealMenu> menuMap : menus.values()) {
            MealMenu menu = null;
            for (MealMenu value : menuMap.values()) {
                menu = value;
                break;
            }
            Sheet sheet = workbook.createSheet(dfOfSheet.format(menu.getForDayOfDate()));
            // 创建标题
            createTitle(workbook, sheet, menu);
            // 创建菜单
            createData(workbook, sheet, menuMap);
        }
        try (FileOutputStream fos = new FileOutputStream(filePath)) {
            workbook.write(fos);
            fos.flush();
        } catch (Exception e) {
            throw new RuntimeException(e);
        } finally {
            workbook.close();
        }
    }

    public Tuple2<String, Integer> getCate(int row, Map<String, Triple3<Integer, Integer, Integer>> cateMap) {
        for (String s : cateMap.keySet()) {
            Triple3<Integer, Integer, Integer> tuple2 = cateMap.get(s);
            if (row + 1 >= tuple2.getFirst() && row + 1 <= tuple2.getSecond()) {
                return new Tuple2<>(s, tuple2.getThird());
            }
        }
        return null;
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
        CellRangeAddress region = new CellRangeAddress(0, 0, 0, exportSetting.menuDataCellMax);
        sheet.addMergedRegion(region);
        cellTitle.setCellStyle(style);
        cellTitle.setCellValue(title);
    }

    public void createData(Workbook workbook, Sheet sheet, Map<Integer, MealMenu> menuMap) {
        int startRow = exportSetting.startRow - 1;
        int startCell = exportSetting.startCell - 'A';
        // 设置字体
        Font font = workbook.createFont();
        font.setBold(true);
        font.setFontName(exportSetting.font);
        font.setFontHeightInPoints(exportSetting.menuFontSize);
        Map<Integer, MealMenu> sortedMap = new TreeMap<>(menuMap);
        for (MealMenu menu : sortedMap.values()) {
            int totalCell = exportSetting.menuDataCellMax;
            int totalRow = (menu.menuList().size() / totalCell) + (menu.menuList().size() % totalCell > 0 ? 1 : 0);
            int index = 0;
            for (int i = 0; i < totalRow; i++) {
                Row row = sheet.createRow(startRow);
                row.setHeightInPoints(exportSetting.menuRowHigh);
                // 设置样式
                CellStyle style = workbook.createCellStyle();
                style.setAlignment(HorizontalAlignment.CENTER);
                style.setVerticalAlignment(VerticalAlignment.CENTER);
                style.setWrapText(true);
                style.setBorderTop(BorderStyle.THIN);
                style.setBorderBottom(BorderStyle.THIN);
                style.setBorderLeft(BorderStyle.THIN);
                style.setBorderRight(BorderStyle.THIN);
                style.setFont(font);
                for (int j = startCell; j < totalCell + 1; j++) {
                    System.out.printf("create row:%d,cell:%d%n", i,  j);
                    Cell cell = row.createCell(startCell + j, CellType.STRING);
                    cell.setCellStyle(style);
                    if (j == startCell) {
                        cell.setCellValue(menu.category);
                        sheet.setColumnWidth(startCell, exportSetting.firstCellWidth);
                        continue;
                    }
                    sheet.setColumnWidth(j, exportSetting.menuCellWidth);
                    if (index < menu.menuList().size()) {
                        // 创建每日详细菜单列
                        MealMenu.Node m = menu.menuList().get(index);
                        if (m.getName() == null || m.getName().trim().isEmpty()) {
                            cell.setCellValue("");
                        } else {
                            cell.setCellValue(String.format("%s\n%s", m.getName(), m.getPrice()));
                        }
                    } else {
                        cell.setCellValue("");
                    }
                    index++;
                }
                startRow++;
            }
            // 合并类别单元格
            if (totalRow < 2) {
                continue;
            }
            CellRangeAddress region = new CellRangeAddress(startRow - totalRow, startRow - 1, startCell, startCell);
            sheet.addMergedRegionUnsafe(region);
        }
    }
}
