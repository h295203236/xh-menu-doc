package org.mooon.util;

public final class StringUtil {
    private StringUtil() {}

    /**
     * 计算字符长度，中文字符算2，其他英文字符和数字算1
     * @param s 字符串
     */
    public static int withLength(String s) {
        int sum = 0;
        final String string = s.trim();
        for (char c : string.toCharArray()) {
            if (isChinese(c)) {
                sum += 2;
            } else {
                sum += 1;
            }
        }
        return sum;
    }

    /**
     * 判断是否是汉字
     * @param c 字符
     */
    public static boolean isChinese(char c) {
        Character.UnicodeScript sc = Character.UnicodeScript.of(c);
        return sc == Character.UnicodeScript.HAN;
    }
}
