package org.mooon.biz;

import org.apache.poi.xwpf.usermodel.*;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;

public class DinnerRuWeiService {

    public DinnerRuWeiService() {}

    public void export(String filePath, String title, Map<String, Map<Integer, MealMenu>> menuMap, int ruWeiPro) throws IOException {
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
        int startRow = 0;
        for (String s : menuMap.keySet()) {
            // 创建日期
            XWPFTableRow row = table.getRow(startRow);
            row.setHeight(250);
            XWPFParagraph paragraph = row.getCell(0).addParagraph();
            paragraph.setAlignment(ParagraphAlignment.CENTER);
            paragraph.setBorderLeft(Borders.THICK);
            paragraph.setBorderRight(Borders.THICK);
            XWPFRun run = paragraph.createRun();
            run.setBold(true);
            run.setFontSize(16);
            run.setFontFamily("宋体");
            run.setText(s);
            // 设置菜单
            MealMenu menu = menuMap.get(s).get(ruWeiPro);
            int size = menu.menuList().size();
            int totalCells = size / 2 + (size % 2 > 0 ? 1 : 0);
            for (int i = 0; i < totalCells; i++) {
                XWPFParagraph paragraph2 = row.getCell(i + 1).addParagraph();
                paragraph2.setAlignment(ParagraphAlignment.LEFT);
                XWPFRun run2 = paragraph2.createRun();
                run2.setBold(true);
                run2.setFontSize(16);
                run2.setFontFamily("宋体");

                String v1;
                MealMenu.Node node1 = menu.menuList().get(i * 2);
                if (node1.getName() == null || node1.getName().trim().isEmpty()) {
                    v1 = "";
                } else {
                    v1 = String.format("%s   %s", node1.getName(), node1.getPrice());
                }
                if (!v1.isEmpty()) {
                    run2.setText(v1);
                }

                if (i * 2 + 1 < size) {
                    String v2;
                    MealMenu.Node node2 = menu.menuList().get(i * 2+ 1 );
                    if (node2.getName() == null || node2.getName().trim().isEmpty()) {
                        v2 = "";
                    } else {
                        v2 = String.format("%s   %s", node2.getName(), node2.getPrice());
                    }

                    if (!v1.isEmpty()) {
                        run2.addBreak();
                        run2.setText(v2, 1);
                    } else {
                        run2.setText(v2);
                    }
                }
            }
            startRow++;
        }
    }

}
