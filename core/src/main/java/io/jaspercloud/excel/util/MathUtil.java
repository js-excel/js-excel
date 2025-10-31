package io.jaspercloud.excel.util;

public class MathUtil {

    public static Integer parseInt(String text) {
        try {
            int i = Integer.parseInt(text);
            return i;
        } catch (NumberFormatException e) {
            return null;
        }
    }
}
