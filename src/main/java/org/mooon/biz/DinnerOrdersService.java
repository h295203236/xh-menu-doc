package org.mooon.biz;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

/**
 * 午餐预定单导出
 */
@Slf4j
public class DinnerOrdersService {

    private final String path;
    private final String fileName;
    public DinnerOrdersService(String path, String fileName) {
        this.path = path;
        this.fileName = fileName;
    }

    public void exportReport1(Map<Date, Map<Integer, MealMenu>> menuMap) {
        System.out.println("starting export " + fileName);

        try (InputStream is = this.getClass().getClassLoader().getResourceAsStream("晚餐预订单模版.xlsx");
             Workbook workbook = new XSSFWorkbook(is)) {
            int sheetIndex = 0;
            for (Map.Entry<Date, Map<Integer, MealMenu>> dateMapEntry : menuMap.entrySet()) {
                sheetIndex++;

                // get all menu node list
                List<MealMenu.Node> nodeList = dateMapEntry.getValue().values().stream().flatMap(s -> s.menuList().stream()).collect(Collectors.toList());

                Sheet sheet = workbook.cloneSheet(0);
                workbook.setSheetName(sheetIndex, ExcelHelper.formatDateOfSheet(dateMapEntry.getKey()));
                sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
                sheet.getPrintSetup().setFitHeight((short) 1);
                sheet.getPrintSetup().setFitWidth((short) 1);
                // 纵向打印
                sheet.getPrintSetup().setLandscape(true);

                int firstRowNum = sheet.getFirstRowNum();
                int lastRowNum = sheet.getLastRowNum();

                // set first row title
                Row rowTitle = sheet.getRow(firstRowNum);
                Cell cellTitle = rowTitle.getCell(0);
                String title = cellTitle.getStringCellValue();
                cellTitle.setCellValue(title.replace("{date}", ExcelHelper.formatDateOfTile(dateMapEntry.getKey())));

                boolean isMenuRow = false;
                int nodeIndex = 0;
                int totalNodes = nodeList.size();
                for (int i = firstRowNum + 1; i <= lastRowNum; i++) {
                    MealMenu.Node node = null;
                    if (nodeIndex < totalNodes) {
                        node = nodeList.get(nodeIndex);
                    }

                    Row row = sheet.getRow(i);
                    int firstCellNum = row.getFirstCellNum();
                    int lastCellNum = row.getLastCellNum();
                    for (int j = firstCellNum; j <= lastCellNum; j++) {
                        Cell cell = row.getCell(j);
                        if (cell == null) {
                            continue;
                        }
                        String value = getCellValue(cell);
                        if ("{name}".equalsIgnoreCase(value)) {
                            isMenuRow = true;
                            if (nodeIndex >= totalNodes) {
                                cell.setCellValue("");
                                continue;
                            }
                            formatName(row, cell, node.getName(), firstCellNum == j);
                        } else if ("{price}".equalsIgnoreCase(value)) {
                            isMenuRow = true;
                            if (nodeIndex >= totalNodes) {
                                cell.setCellValue("");
                                continue;
                            }
                            formatPrice(cell, node.getPrice());
                        }
                    }
                    if (isMenuRow) {
                        nodeIndex++;
                        isMenuRow = false;
                    }
                }
            }

            try(FileOutputStream fos = new FileOutputStream(path + fileName)) {
                workbook.removeSheetAt(0);
                workbook.write(fos);
                fos.flush();
            } catch (IOException e) {
                log.error("can not export, ", e);
            }
        } catch (Exception e) {
            log.error("can not export, ", e);
        }
    }

    private void formatName(Row row, Cell cell, String name, boolean isNeedResizeH) {
        int len = name.length();
        if (len <= 6) {
            cell.setCellValue(name);
            return;
        }
        String left = name.substring(0, 4);
        String right = name.substring(4);
        cell.setCellValue(left + "\n" + right);
        if (isNeedResizeH) {
            float currH = row.getHeightInPoints();
            row.setHeightInPoints(currH * 1.8f);
        }
    }

    private void formatPrice(Cell cell, String price) {
        if (!price.contains("/")) {
            if ("0元".equalsIgnoreCase(price)) {
                cell.setCellValue("");
                return;
            }
            cell.setCellValue(price.replace("元", ""));
            return;
        }
        String[] values = price.split("/");
        cell.setCellValue(values[0] + "/份");
    }

    private String getCellValue(Cell cell) {
        try {
            return cell.getStringCellValue();
        } catch (IllegalStateException e) {
            double v = cell.getNumericCellValue();
            BigDecimal bg = BigDecimal.valueOf(v).setScale(2, RoundingMode.UP);
            double num = bg.doubleValue();
            if (Math.round(num) - num == 0) {
                return (long) num + "";
            }
            return num + "";
        }
    }
}
