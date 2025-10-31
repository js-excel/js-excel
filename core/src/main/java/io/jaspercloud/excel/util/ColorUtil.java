package io.jaspercloud.excel.util;

import java.util.Locale;

public class ColorUtil {

    public static String getRGBHex(byte[] rgb) {
        if (rgb == null) {
            return null;
        }

        StringBuilder sb = new StringBuilder();
        for (byte c : rgb) {
            int i = c & 0xff;
            String cs = Integer.toHexString(i);
            if (cs.length() == 1) {
                sb.append('0');
            }
            sb.append(cs);
        }
        return sb.toString().toUpperCase(Locale.ROOT);
    }
}
