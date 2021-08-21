package org.mooon.biz;

import org.apache.poi.sl.usermodel.VerticalAlignment;
import org.apache.poi.xwpf.usermodel.*;
import org.mooon.util.StringUtil;
import org.mooon.util.WeekDayUtil;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;

public class DinnerRuWeiService {

    public DinnerRuWeiService() {}

    public void export(String filePath, String title, Map<String, Map<Integer, MealMenu>> menuMap, int ruWeiPro) {
        XWPFDocument doc = new XWPFDocument();
        // 写入文件
        try (doc; FileOutputStream fos = new FileOutputStream(filePath)) {
            createDocReport1(doc, title, menuMap, ruWeiPro);
            doc.write(fos);
            fos.flush();
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    public void createDocReport1(XWPFDocument doc, String title, Map<String, Map<Integer, MealMenu>> menuMap, int ruWeiPro) {
        // 创建标题
        XWPFParagraph titleParagraph = doc.createParagraph();
        titleParagraph.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun titleParagraphRun = titleParagraph.createRun();
        titleParagraphRun.setText(title);
        titleParagraphRun.addCarriageReturn();
        titleParagraphRun.setText("卤味预订单");
        titleParagraphRun.addCarriageReturn();
        titleParagraphRun.addCarriageReturn();
        titleParagraphRun.addCarriageReturn();
        titleParagraphRun.addCarriageReturn();
        titleParagraphRun.addCarriageReturn();
        titleParagraphRun.setColor("000000");
        titleParagraphRun.setBold(true);
        titleParagraphRun.setFontSize(20);
        titleParagraphRun.setFontFamily("宋体");

        // 创建一个多行3列的表格
        XWPFTable table = doc.createTable(menuMap.size(), 3);
        table.setTableAlignment(TableRowAlign.LEFT);
        table.setWidthType(TableWidthType.DXA);
        table.setWidth(8000);
        int startRow = 0;
        for (String s : menuMap.keySet()) {
            // 创建日期
            XWPFTableRow row = table.getRow(startRow);
            row.setHeight(1500);
            XWPFTableCell cell = row.getCell(0);
            cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
            cell.setWidthType(TableWidthType.PCT);
            cell.setWidth("20.0%");
            XWPFParagraph paragraph = cell.addParagraph();
            paragraph.setAlignment(ParagraphAlignment.CENTER);
            paragraph.setBorderLeft(Borders.THICK);
            paragraph.setBorderRight(Borders.THICK);
            XWPFRun run = paragraph.createRun();
            run.setBold(true);
            run.setFontSize(16);
            run.setFontFamily("宋体");
            run.setText(WeekDayUtil.getNickyWeekDay(s));
            // 设置菜单
            MealMenu menu = menuMap.get(s).get(ruWeiPro);
            int size = menu.menuList().size();
            int totalCells = size / 2 + (size % 2 > 0 ? 1 : 0);
            for (int i = 0; i < totalCells; i++) {
                XWPFTableCell cell2 = row.getCell(i + 1);
                cell2.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
                cell2.setWidthType(TableWidthType.PCT);
                cell2.setWidth("40.0%");
                XWPFParagraph paragraph2 = cell2.addParagraph();
                paragraph2.setAlignment(ParagraphAlignment.LEFT);
                paragraph2.setVerticalAlignment(TextAlignment.CENTER);
                XWPFRun run2 = paragraph2.createRun();
                run2.setBold(true);
                run2.setFontSize(16);
                run2.setFontFamily("宋体");

                String v1;
                MealMenu.Node node1 = menu.menuList().get(i * 2);
                if (node1.getName() == null || node1.getName().trim().isEmpty()) {
                    v1 = formatText("", "");
                } else {
                    v1 = formatText(node1.getName(), node1.getPrice());
                }
                run2.setText(v1);
                run2.addBreak(BreakClear.RIGHT);

                if (i * 2 + 1 < size) {
                    String v2;
                    MealMenu.Node node2 = menu.menuList().get(i * 2+ 1 );
                    if (node2.getName() == null || node2.getName().trim().isEmpty()) {
                        v2 = formatText("", "");
                    } else {
                        v2 = formatText(node2.getName(), node2.getPrice());
                    }
                    run2.setText(v2);
                    run2.addCarriageReturn();
                }
            }
            startRow++;
        }
    }

    /**
     * 格式化内容
     * @param left  左字符串
     * @param right 右字符串
     */
    private String formatText(String left, String right) {
        return " " + formatText(16, left, right);
    }

    /**
     * 格式化内容
     * @param totalLen 总长度
     * @param left  左字符串
     * @param right 右字符串
     */
    private String formatText(int totalLen, String left, String right) {
        int leftLen = StringUtil.withLength(left);
        int rightLen = StringUtil.withLength(right);
        return left + " ".repeat(Math.max(0, totalLen - leftLen - rightLen)) + right;
    }
}
