package org.mooon.biz;

import java.math.BigDecimal;
import java.math.RoundingMode;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

public final class ExcelHelper {

    public static final String DATE_FORMAT = "yyyy-MM-dd";
    public static final String DATE_FORMAT_OF_TITLE = "yyyy年MM月dd日";
    public static final String DATE_FORMAT_OF_SHEET = "MM.dd";

    private ExcelHelper() {}

    public static String formatDateOfTile(Date date) throws ParseException {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat(DATE_FORMAT_OF_TITLE);
        return simpleDateFormat.format(date);
    }

    public static String formatDateOfSheet(Date date) throws ParseException {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat(DATE_FORMAT_OF_SHEET);
        return simpleDateFormat.format(date);
    }

    public static String formatDate(Date date, String format) throws ParseException {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat(format);
        return simpleDateFormat.format(date);
    }

    public static Date parseDate(String day) throws ParseException {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat(DATE_FORMAT);
        return simpleDateFormat.parse(day);
    }

    public static Date addDay(Date date, int day) throws ParseException {
        Calendar calendar = Calendar.getInstance();
        calendar.setTime(date);
        calendar.add(Calendar.DAY_OF_MONTH, day);
        return calendar.getTime();
    }

    public static Date addDay(String date, int day) throws ParseException {
        Date dt = parseDate(date);
        Calendar calendar = Calendar.getInstance();
        calendar.setTime(dt);
        calendar.add(Calendar.DAY_OF_MONTH, day);
        return calendar.getTime();
    }

    public static String formatPrice(Double price) {
        if (price == null) {
            return "免费";
        }
        BigDecimal bg = BigDecimal.valueOf(price).setScale(2, RoundingMode.UP);
        double num = bg.doubleValue();
        if (Math.round(num) - num == 0) {
            return (long) num + "元";
        }
        return num + "元";
    }
}
