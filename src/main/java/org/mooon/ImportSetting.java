package org.mooon;

public class ImportSetting {
    /**
     * 菜单所在sheet名称
     */
    public String sheetName = null;
    /**
     * 起始日期
     */
    public String dayStartAt = "2021-08-16";

    /**
     * 日期所属行
     */
    public int dayAtRow = 3;

    /**
     * 日期开始列
     */
    public char dayStartCell = 'B';

    /**
     * 日期结束列
     */
    public char dayEndCell = 'K';

    /**
     * 菜单开始行
     */
    public int menuStartRow = 5;

    /**
     * 菜单结束行
     */
    public int menuEndRow = 21;
}
