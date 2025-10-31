package io.jaspercloud.excel.util;

import org.apache.poi.ss.usermodel.CellType;

import java.util.HashMap;
import java.util.Map;

public class ExcelEnumUtil {

    private static Map<Integer, CellType> cellTypeMap = new HashMap<>();

    static {
        cellTypeMap.put(2, CellType.NUMERIC);
        cellTypeMap.put(1, CellType.STRING);
        cellTypeMap.put(3, CellType.BOOLEAN);
    }

    public static CellType getCellType(Integer cellType) {
        return cellTypeMap.getOrDefault(cellType, CellType.STRING);
    }
}
