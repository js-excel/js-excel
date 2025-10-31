package io.jaspercloud.excel.util;

import org.apache.poi.ss.usermodel.*;

import java.util.HashMap;
import java.util.Map;

public class UniverEnumUtil {

    private static Map<BorderStyle, Integer> borderMap = new HashMap<>();
    private static Map<FontUnderline, Integer> underlineMap = new HashMap<>();
    private static Map<HorizontalAlignment, Integer> htMap = new HashMap<>();
    private static Map<VerticalAlignment, Integer> vtMap = new HashMap<>();
    private static Map<CellType, Integer> cellTypeMap = new HashMap<>();

    static {
        borderMap.put(BorderStyle.NONE, 0);
        borderMap.put(BorderStyle.THIN, 1);
        borderMap.put(BorderStyle.MEDIUM, 8);
        borderMap.put(BorderStyle.DASHED, 4);
        borderMap.put(BorderStyle.DOTTED, 3);
        borderMap.put(BorderStyle.THICK, 13);
        borderMap.put(BorderStyle.DOUBLE, 7);
        borderMap.put(BorderStyle.HAIR, 2);
        borderMap.put(BorderStyle.MEDIUM_DASHED, 9);
        borderMap.put(BorderStyle.DASH_DOT, 5);
        borderMap.put(BorderStyle.MEDIUM_DASH_DOT, 10);
        borderMap.put(BorderStyle.DASH_DOT_DOT, 6);
        borderMap.put(BorderStyle.MEDIUM_DASH_DOT_DOT, 11);
        borderMap.put(BorderStyle.SLANTED_DASH_DOT, 12);
        underlineMap.put(FontUnderline.SINGLE, 12);
        underlineMap.put(FontUnderline.DOUBLE, 10);
        underlineMap.put(FontUnderline.NONE, 11);
        htMap.put(HorizontalAlignment.GENERAL, 1);
        htMap.put(HorizontalAlignment.LEFT, 1);
        htMap.put(HorizontalAlignment.CENTER, 2);
        htMap.put(HorizontalAlignment.RIGHT, 3);
        htMap.put(HorizontalAlignment.JUSTIFY, 4);
        htMap.put(HorizontalAlignment.DISTRIBUTED, 6);
        vtMap.put(VerticalAlignment.TOP, 1);
        vtMap.put(VerticalAlignment.CENTER, 2);
        vtMap.put(VerticalAlignment.BOTTOM, 3);
        vtMap.put(VerticalAlignment.JUSTIFY, 2);
        cellTypeMap.put(CellType.NUMERIC, 2);
        cellTypeMap.put(CellType.STRING, 1);
        cellTypeMap.put(CellType.BOOLEAN, 3);
    }

    public static int getUniverBorderCode(BorderStyle borderStyle) {
        return borderMap.getOrDefault(borderStyle, BorderStyle.NONE.ordinal());
    }

    public static int getUniverUnderlineCode(FontUnderline underline) {
        return underlineMap.getOrDefault(underline, FontUnderline.NONE.ordinal());
    }

    public static int getUniverHorizontalAlignmentCode(HorizontalAlignment alignment) {
        return htMap.getOrDefault(alignment, HorizontalAlignment.LEFT.ordinal());
    }

    public static int getUniverVerticalAlignmentCode(VerticalAlignment alignment) {
        return vtMap.getOrDefault(alignment, VerticalAlignment.CENTER.ordinal());
    }

    public static int getUniverCellTypeCode(CellType cellType) {
        return cellTypeMap.getOrDefault(cellType, CellType._NONE.ordinal());
    }
}
