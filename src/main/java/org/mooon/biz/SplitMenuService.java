package org.mooon.biz;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.LinkedList;
import java.util.List;
import java.util.stream.Collectors;

/**
 * 拆分菜单导出
 */
@Slf4j
public final class SplitMenuService {
    private final String path;
    private final String sourceFileName;
    private final String sourceSheet;
    private final String fileName;

    public SplitMenuService(String path, String sourceFileName, String sourceSheet, String fileName) {
        this.path = path;
        this.sourceFileName = sourceFileName;
        this.sourceSheet = sourceSheet;
        this.fileName = fileName;
    }

    public void exportReport1(int dinnerStart) {
        System.out.println("staring export " + fileName);
        try (InputStream is = new FileInputStream(path + "/" + sourceFileName);
             Workbook workbook = new XSSFWorkbook(is)) {
            int at = -1;
            List<Integer> notNeededList = new LinkedList<>();
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                Sheet existSheet = workbook.getSheetAt(i);
                if (!sourceSheet.equalsIgnoreCase(existSheet.getSheetName())) {
                    notNeededList.add(i);
                    continue;
                }
                at = i;
            }
            if (at == -1) {
                log.error("no sheet {} found.", sourceSheet);
                return;
            }

            Sheet lunchSheet = workbook.cloneSheet(at);
            workbook.setSheetName(workbook.getSheetIndex(lunchSheet.getSheetName()), "午餐");
            //removeRegions(lunchSheet, dinnerStart-1, lunchSheet.getLastRowNum());
            for (int i = lunchSheet.getLastRowNum(); i > dinnerStart - 1; i--) {
                removeRow(lunchSheet, dinnerStart - 1);
            }

            Sheet dinnerSheet = workbook.cloneSheet(at);
            workbook.setSheetName(workbook.getSheetIndex(dinnerSheet.getSheetName()), "晚餐");
            removeRegions(dinnerSheet, 0, dinnerStart - 1);
            for (int i = 0; i < dinnerStart - 1; i++) {
                removeRow(dinnerSheet, 0);
            }
            int lastCellNum = dinnerSheet.getRow(1).getLastCellNum();
            dinnerSheet.addMergedRegion(new CellRangeAddress(0, 0, 0, lastCellNum - 1));

            // 打印设置
            printSetUp(lunchSheet);
            printSetUp(dinnerSheet);

            // 移除不需要的sheet
            workbook.removeSheetAt(at);
            for (Integer integer : notNeededList) {
                workbook.removeSheetAt(integer);
            }

            try (FileOutputStream fos = new FileOutputStream(path + "/" + fileName)) {
                workbook.write(fos);
                fos.flush();
            } catch (IOException e) {
                log.error("export error.", e);
            }
        } catch (Exception e) {
            log.error("can not export, ", e);
        }
    }

    private static void removeRow(Sheet sheet, int rowIndex) {
        int lastRow = sheet.getLastRowNum();
        if (rowIndex >= 0 && rowIndex < lastRow) {
            sheet.shiftRows(rowIndex + 1, lastRow, -1);
        }
        if (rowIndex == lastRow) {
            sheet.removeRow(sheet.getRow(rowIndex));
        }
    }

    private void removeRegions(Sheet sheet, int startRow, int endRow) {
        List<CellRangeAddress> rangeAddressList = sheet.getMergedRegions();
        for (int i = rangeAddressList.size() - 1; i >= 0; i--) {
            CellRangeAddress addresses = rangeAddressList.get(i);
            if (addresses.getFirstRow() >= startRow && addresses.getLastRow() <= endRow) {
                sheet.removeMergedRegion(i);
            }
        }
    }

    private void printSetUp(Sheet sheet) {
        sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
        sheet.setMargin(Sheet.TopMargin, 0.75);
        sheet.setMargin(Sheet.BottomMargin, 0.75);
        sheet.setMargin(Sheet.LeftMargin, 0.25);
        sheet.setMargin(Sheet.RightMargin, 0.25);
        // 纵向打印
        sheet.getPrintSetup().setLandscape(true);
    }
}
