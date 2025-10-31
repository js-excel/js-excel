package io.jaspercloud.excel.util;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;

public class SheetUtil {

    public static void addEmptyRow(XSSFRow row) {
        addEmptyCell(row, 26);
    }

    public static void addEmptyCell(XSSFRow row, int count) {
        for (int i = 0; i < count; i++) {
            XSSFCell cell = row.createCell(i);
            cell.setCellStyle(row.getRowStyle());
        }
    }
}
