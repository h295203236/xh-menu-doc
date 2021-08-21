package org.mooon.util;

import java.util.LinkedHashMap;
import java.util.Map;

/**
 * 解析每周日期信息
 */
public final class WeekDayUtil {

    private WeekDayUtil() {}

    static final String[] WEEK_DAYS = { "星期一","星期二","星期三","星期四","星期五","星期六","星期日" };
    static final Map<String, String> WEEK_DAYS_MAP;
    static final Map<String, String> WEEK_DAYS_NICKY_MAP;
    static {
        WEEK_DAYS_MAP = new LinkedHashMap<>();
        WEEK_DAYS_MAP.put("星期一", "周一");
        WEEK_DAYS_MAP.put("星期二", "周二");
        WEEK_DAYS_MAP.put("星期三", "周三");
        WEEK_DAYS_MAP.put("星期四", "周四");
        WEEK_DAYS_MAP.put("星期五", "周五");
        WEEK_DAYS_MAP.put("星期六", "周六");
        WEEK_DAYS_MAP.put("星期日", "周日");
        WEEK_DAYS_NICKY_MAP = new LinkedHashMap<>();
        WEEK_DAYS_NICKY_MAP.put("周一", "星期一");
        WEEK_DAYS_NICKY_MAP.put("周二", "星期二");
        WEEK_DAYS_NICKY_MAP.put("周三", "星期三");
        WEEK_DAYS_NICKY_MAP.put("周四", "星期四");
        WEEK_DAYS_NICKY_MAP.put("周五", "星期五");
        WEEK_DAYS_NICKY_MAP.put("周六", "星期六");
        WEEK_DAYS_NICKY_MAP.put("周日", "星期日");
    }

    public static String getNickyWeekDay(String weekDay) {
        return WEEK_DAYS_MAP.get(weekDay);
    }
}
