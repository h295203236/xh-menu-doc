package org.mooon;

public class ExportSetting {

    /**
     * 导出文件名称
     */
    public String fileName = "";

    /**
     * 第一行标题名称
     */
    public String title = "%s 中餐周菜单";

    /**
     * 标题字体大小
     */
    public short tileFontSize = 20;

    /**
     * 标题行高
     */
    public short tileRowHigh = 50;

    /**
     * 菜单行高
     */
    public short menuRowHigh = 80;

    /**
     * 菜单字体大小
     */
    public short menuFontSize = 20;

    /**
     * 菜单列宽
     */
    public int menuCellWidth = 7000;

    /**
     * 第一列列宽（中餐列别名称使用列）
     */
    public int firstCellWidth = 2000;

    /**
     * 菜单最大行数
     */
    public int menuRowMax = 7;

    /**
     * 菜单最大列宽
     */
    public int menuDataCellMax = 3;

    /**
     * 数据写入起始行号
     */
    public int startRow = 2;
    /**
     * 数据写入起始列号
     */
    public char startCell = 'A';

    /**
     * 字体
     */
    public String font = "宋体";

    /**
     * 字体大小
     */
    public int fontSize = 20;

    /**
     * 内容居中
     */
    public boolean contentMiddle = true;
}
