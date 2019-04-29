package com.tellyes.platform.toolkit.utils;

import java.time.LocalDate;
import java.time.ZoneId;
import java.util.Date;
import java.util.Optional;

/**
 * LocalDate工具类
 * @author xiehai
 * @date 2018/07/25 17:33
 * @Copyright(c) tellyes tech. inc. co.,ltd
 * @see java.time.LocalDate
 */
public interface LocalDateUtil {
    /**
     * 时区id
     */
    ZoneId ZONE_ID = ZoneId.systemDefault();
    /**
     * {@link Date} to {@link LocalDate}
     * @param date 日期
     * @return {@link LocalDate}
     */
    static LocalDate toLocalDate(Date date) {
        return
            Optional.ofNullable(date)
                .map(d -> d.toInstant().atZone(ZONE_ID).toLocalDate())
                .orElse(null);
    }

    /**
     * {@link LocalDate} to {@link Date}
     * @param localDate 日期
     * @return {@link Date}
     */
    static Date toDate(LocalDate localDate) {
        return
            Optional.ofNullable(localDate)
                .map(l -> Date.from(l.atStartOfDay(ZONE_ID).toInstant()))
                .orElse(null);
    }
}
