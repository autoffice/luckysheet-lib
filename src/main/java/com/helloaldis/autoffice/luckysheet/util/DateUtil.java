package com.helloaldis.autoffice.luckysheet.util;

import java.time.Instant;
import java.time.format.DateTimeFormatter;

public class DateUtil {
    public static String toJsonTimeString(long millis) {
        return DateTimeFormatter.ISO_INSTANT.format(Instant.ofEpochMilli(millis));
    }
}
