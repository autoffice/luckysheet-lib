package io.github.autoffice.luckysheet.util;

public class Util {
    public static <T> T requireNonNullElse(T obj, T defaultObj) {
        return obj != null ? obj : defaultObj;
    }
}
