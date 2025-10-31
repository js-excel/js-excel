package io.jaspercloud.excel;

import io.jaspercloud.excel.domain.*;
import io.jaspercloud.excel.util.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellFill;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.*;

import java.util.Map;
import java.util.Optional;
import java.util.TreeMap;

public class JsExcel {

    private XSSFWorkbook workbook;

    public JsExcel(XSSFWorkbook workbook) {
        this.workbook = workbook;
    }

    public UniverWorkbook render() {
        UniverWorkbook univerWorkbook = new UniverWorkbook();
        for (int s = 0; s < workbook.getNumberOfSheets(); s++) {
            XSSFSheet sheet = workbook.getSheetAt(s);
            int columnCount = calcColumnCount(sheet);
            UniverSheet univerSheet = new UniverSheet();
            univerSheet.setId(String.valueOf(s));
            univerSheet.setName(sheet.getSheetName());
            univerSheet.setHidden(SheetVisibility.HIDDEN.equals(workbook.getSheetVisibility(s)) ? 1 : 0);
            processColumns(sheet, univerWorkbook, univerSheet, columnCount);
            processRows(workbook, univerWorkbook, sheet, columnCount, univerSheet);
            processCellMergeRange(sheet, univerSheet);
            processFreezeRow(sheet, univerSheet);
            univerSheet.setRowCount(sheet.getPhysicalNumberOfRows() + 30);
            if (columnCount > 0) {
                univerSheet.setColumnCount(columnCount);
            }
            reCalcCellStyle(univerWorkbook, univerSheet);
            univerWorkbook.addSheets(univerSheet);
        }
        return univerWorkbook;
    }

    private void reCalcCellStyle(UniverWorkbook univerWorkbook, UniverSheet univerSheet) {
        Map<Integer, UniverColumnData> columnData = univerSheet.getColumnData();
        Map<Integer, UniverRowData> rowData = univerSheet.getRowData();
        Map<Integer, Map<Integer, UniverCell>> cellData = univerSheet.getCellData();
        for (Map.Entry<Integer, UniverRowData> row : rowData.entrySet()) {
            Integer r = row.getKey();
            Integer rowStyleId = MathUtil.parseInt(row.getValue().getS());
            for (Map.Entry<Integer, UniverColumnData> column : columnData.entrySet()) {
                Integer c = column.getKey();
                Integer colStyleId = MathUtil.parseInt(column.getValue().getS());
                Map<Integer, UniverCell> cellMap = cellData.computeIfAbsent(r, k -> new TreeMap<>());
                UniverCell univerCell = cellMap.computeIfAbsent(c, k -> new UniverCell());
                String s = univerCell.getS();
                if (null != s) {
                    continue;
                }
                if (null != rowStyleId) {
                    s = String.valueOf(rowStyleId);
                } else if (null != colStyleId) {
                    s = String.valueOf(colStyleId);
                }
                univerCell.setS(s);
            }
        }
    }

    private void processFreezeRow(XSSFSheet sheet, UniverSheet univerSheet) {
        if (null == sheet.getPaneInformation()) {
            return;
        }
        int freezeRow = sheet.getPaneInformation().getHorizontalSplitTopRow();
        UniverFreeze freeze = new UniverFreeze();
        freeze.setStartRow(freezeRow);
        freeze.setYSplit(freezeRow);
        univerSheet.setFreeze(freeze);
    }

    private void processRows(XSSFWorkbook workbook, UniverWorkbook univerWorkbook, XSSFSheet sheet, int columnCount, UniverSheet univerSheet) {
        for (CTRow ctRow : sheet.getCTWorksheet().getSheetData().getRowList()) {
            int r = getRowNum(ctRow);
            XSSFRow row = sheet.getRow(r);
            processRow(workbook, univerWorkbook, univerSheet, row);
            if (row.getPhysicalNumberOfCells() <= 0) {
                SheetUtil.addEmptyCell(row, columnCount);
            }
            for (CTCell ctCell : ctRow.getCList()) {
                int c = getColNum(ctCell);
                XSSFCell cell = row.getCell(c);
                String loc = cell.getCTCell().getR();
                processCell(workbook, univerWorkbook, univerSheet, cell);
            }
        }
    }

    private int getRowNum(CTRow ctRow) {
        int r = (int) ctRow.getR() - 1;
        return r;
    }

    private int getColNum(CTCell ctCell) {
        int c = new CellReference(ctCell.getR()).getCol();
        return c;
    }

    private void processCellMergeRange(XSSFSheet sheet, UniverSheet univerSheet) {
        for (int i = 0; i < sheet.getMergedRegions().size(); i++) {
            CellRangeAddress rangeAddress = sheet.getMergedRegions().get(i);
            UniverCellRange range = new UniverCellRange();
            range.setRangeType(0);
            range.setStartRow(rangeAddress.getFirstRow());
            range.setEndRow(rangeAddress.getLastRow());
            range.setStartColumn(rangeAddress.getFirstColumn());
            range.setEndColumn(rangeAddress.getLastColumn());
            univerSheet.addCellMergeRange(range);
        }
    }

    private void processColumns(XSSFSheet sheet, UniverWorkbook univerWorkbook, UniverSheet univerSheet, int columnCount) {
        for (int i = 0; i < columnCount; i++) {
            float columnWidth = sheet.getColumnWidthInPixels(i);
            float defWidth = sheet.getDefaultColumnWidth() * Units.DEFAULT_CHARACTER_WIDTH;
            if (columnWidth == defWidth) {
                continue;
            }
            UniverColumnData columnData = new UniverColumnData();
            columnData.setW((int) columnWidth);
            univerSheet.addColumnData(i, columnData);
        }
        for (CTCols cols : sheet.getCTWorksheet().getColsList()) {
            for (CTCol col : cols.getColList()) {
                int min = (int) (col.getMin() - 1);
                int max = (int) (col.getMax() - 1);
                for (int i = min; i <= max; i++) {
                    XSSFCellStyle cellStyle = (XSSFCellStyle) sheet.getWorkbook().getSheet(sheet.getSheetName()).getColumnStyle(i);
                    processStyle(workbook, univerWorkbook, cellStyle);
                    int styleId = cellStyle.getIndex();
                    univerSheet.getColumnData().get(i).setS(String.valueOf(styleId));
                }
            }
        }
    }

    private int calcColumnCount(XSSFSheet sheet) {
        int maxCellColumnCount = 0;
        CTWorksheet ctWorksheet = sheet.getCTWorksheet();
        CTSheetData sheetData = ctWorksheet.getSheetData();
        for (CTRow ctRow : sheetData.getRowList()) {
            for (CTCell ctCell : ctRow.getCList()) {
                int c = new CellReference(ctCell.getR()).getCol() + 1;
                if (c > maxCellColumnCount) {
                    maxCellColumnCount = c;
                }
            }
        }
        int columnDefineCount = ctWorksheet.getColsList().stream().flatMap(e -> e.getColList().stream())
                .map(e -> (int) e.getMax()).max(Integer::compareTo).orElse(0);
        int maxColumnCount = Math.max(columnDefineCount, maxCellColumnCount);
        return maxColumnCount;
    }

    private void processRow(XSSFWorkbook workbook, UniverWorkbook univerWorkbook, UniverSheet univerSheet, XSSFRow row) {
        UniverRowData univerRowData = new UniverRowData();
        if (row.getHeightInPoints() != row.getSheet().getDefaultRowHeightInPoints()) {
            int height = Math.round(row.getHeightInPoints());
            height = height < univerSheet.getDefaultRowHeight() ? univerSheet.getDefaultRowHeight() : height;
            univerRowData.setH(height);
        } else {
            univerRowData.setH(univerSheet.getDefaultRowHeight());
        }
        univerSheet.addRowData(row.getRowNum(), univerRowData);
        XSSFCellStyle rowStyle = row.getRowStyle();
        if (null == rowStyle) {
            return;
        }
        processStyle(workbook, univerWorkbook, rowStyle);
        int styleId = rowStyle.getIndex();
        univerRowData.setS(String.valueOf(styleId));
    }

    private void processCell(XSSFWorkbook workbook, UniverWorkbook univerWorkbook, UniverSheet univerSheet, XSSFCell cell) {
        //univerCell
        UniverCell univerCell = new UniverCell();
        univerCell.setT(UniverEnumUtil.getUniverCellTypeCode(cell.getCellType()));
        univerCell.setV(new DataFormatter().formatCellValue(cell));
        univerSheet.addCell(cell.getRowIndex(), cell.getColumnIndex(), univerCell);
        //univerStyle
        XSSFCellStyle cellStyle = cell.getCellStyle();
        if (null == cellStyle) {
            return;
        }
        processStyle(workbook, univerWorkbook, cellStyle);
        int styleId = cellStyle.getIndex();
        univerCell.setS(String.valueOf(styleId));
    }

    private void processStyle(XSSFWorkbook workbook, UniverWorkbook univerWorkbook, XSSFCellStyle cellStyle) {
        if (null == cellStyle) {
            return;
        }
        int styleId = cellStyle.getIndex();
        CTXf ctXf = workbook.getStylesSource().getCellXfAt(styleId);
        UniverStyle univerStyle = univerWorkbook.getOrCreateStyle(String.valueOf(styleId));
        int fontId = (int) ctXf.getFontId();
        XSSFFont font = workbook.getStylesSource().getFontAt(fontId);
        if (null != font.getXSSFColor()) {
            String fc = ColorUtil.getRGBHex(font.getXSSFColor().getRGB());
            univerStyle.setCl(new UniverColorStyle(String.format("#%s", fc)));
        }
        int fillId = (int) ctXf.getFillId();
        XSSFCellFill fill = workbook.getStylesSource().getFillAt(fillId);
        if (null != fill.getFillForegroundColor()) {
            String bc;
            if (fill.getFillForegroundColor().isThemed()) {
                XSSFColor themeColor = workbook.getStylesSource().getTheme().getThemeColor(fill.getFillForegroundColor().getTheme());
                bc = ColorUtil.getRGBHex(themeColor.getRGB());
            } else {
                bc = Optional.ofNullable(fill.getFillForegroundColor())
                        .map(e -> ColorUtil.getRGBHex(e.getRGB())).orElse(null);
            }
            if (null != bc) {
                univerStyle.setBg(new UniverColorStyle(String.format("#%s", bc)));
            }
        }
        if (null != cellStyle.getAlignment()) {
            univerStyle.setHt(UniverEnumUtil.getUniverHorizontalAlignmentCode(cellStyle.getAlignment()));
        }
        if (null != cellStyle.getVerticalAlignment()) {
            univerStyle.setVt(UniverEnumUtil.getUniverVerticalAlignmentCode(cellStyle.getVerticalAlignment()));
        }
        univerStyle.setTb(cellStyle.getWrapText() ? 3 : 0);
        univerStyle.setN(new UniverNumfmt(cellStyle.getDataFormatString()));
        //font
        processFont(cellStyle, univerStyle);
        //border
        processBorder(cellStyle, univerStyle);
    }

    private void processFont(XSSFCellStyle cellStyle, UniverStyle univerStyle) {
        if (null == cellStyle.getFont()) {
            return;
        }
        boolean bold = cellStyle.getFont().getBold();
        univerStyle.setBl(bold ? 1 : 0);
        boolean italic = cellStyle.getFont().getItalic();
        univerStyle.setIt(italic ? 1 : 0);
        FontUnderline fontUnderline = FontUnderline.valueOf(cellStyle.getFont().getUnderline());
        if (!FontUnderline.NONE.equals(fontUnderline)) {
            UniverTextDecoration decoration = new UniverTextDecoration();
            decoration.setS(1);
            decoration.setT(UniverEnumUtil.getUniverUnderlineCode(fontUnderline));
            univerStyle.setUl(decoration);
        }
        if (cellStyle.getFont().getStrikeout()) {
            UniverTextDecoration decoration = new UniverTextDecoration();
            decoration.setS(1);
            univerStyle.setSt(decoration);
        }
    }

    private void processBorder(XSSFCellStyle cellStyle, UniverStyle univerStyle) {
        BorderStyle borderLeft = cellStyle.getBorderLeft();
        BorderStyle borderRight = cellStyle.getBorderRight();
        BorderStyle borderTop = cellStyle.getBorderTop();
        BorderStyle borderBottom = cellStyle.getBorderBottom();
        if (BorderStyle.NONE.equals(borderLeft) && BorderStyle.NONE.equals(borderRight)
                && BorderStyle.NONE.equals(borderTop) && BorderStyle.NONE.equals(borderBottom)) {
            return;
        }
        UniverBorderData borderData = new UniverBorderData();
        {
            UniverBorderStyleData borderStyleData = new UniverBorderStyleData();
            borderStyleData.setS(Integer.valueOf(UniverEnumUtil.getUniverBorderCode(cellStyle.getBorderLeft())));
            XSSFColor color = cellStyle.getLeftBorderXSSFColor();
            if (null != color && null != color.getRGB()) {
                borderStyleData.setCl(new UniverColorStyle(String.format("#%s", ColorUtil.getRGBHex(color.getRGB()))));
            }
            borderData.setL(borderStyleData);
        }
        {
            UniverBorderStyleData borderStyleData = new UniverBorderStyleData();
            borderStyleData.setS(Integer.valueOf(UniverEnumUtil.getUniverBorderCode(cellStyle.getBorderRight())));
            XSSFColor color = cellStyle.getRightBorderXSSFColor();
            if (null != color && null != color.getRGB()) {
                borderStyleData.setCl(new UniverColorStyle(String.format("#%s", ColorUtil.getRGBHex(color.getRGB()))));
            }
            borderData.setR(borderStyleData);
        }
        {
            UniverBorderStyleData borderStyleData = new UniverBorderStyleData();
            borderStyleData.setS(Integer.valueOf(UniverEnumUtil.getUniverBorderCode(cellStyle.getBorderTop())));
            XSSFColor color = cellStyle.getTopBorderXSSFColor();
            if (null != color && null != color.getRGB()) {
                borderStyleData.setCl(new UniverColorStyle(String.format("#%s", ColorUtil.getRGBHex(color.getRGB()))));
            }
            borderData.setT(borderStyleData);
        }
        {
            UniverBorderStyleData borderStyleData = new UniverBorderStyleData();
            borderStyleData.setS(Integer.valueOf(UniverEnumUtil.getUniverBorderCode(cellStyle.getBorderBottom())));
            XSSFColor color = cellStyle.getBottomBorderXSSFColor();
            if (null != color && null != color.getRGB()) {
                borderStyleData.setCl(new UniverColorStyle(String.format("#%s", ColorUtil.getRGBHex(color.getRGB()))));
            }
            borderData.setB(borderStyleData);
        }
        univerStyle.setBd(borderData);
    }

    public void parse(UniverWorkbook univerWorkbook) {
        for (UniverSheet univerSheet : univerWorkbook.getSheets().values()) {
            XSSFSheet sheet = workbook.createSheet(univerSheet.getName());
            parseSheet(univerSheet, sheet);
        }
    }

    private void parseSheet(UniverSheet univerSheet, XSSFSheet sheet) {
        Map<Integer, Map<Integer, UniverCell>> cellData = univerSheet.getCellData();
        for (Map.Entry<Integer, Map<Integer, UniverCell>> rowEntry : cellData.entrySet()) {
            Integer r = rowEntry.getKey();
            XSSFRow row = getOrCreateRow(sheet, r);
            for (Map.Entry<Integer, UniverCell> colEntry : rowEntry.getValue().entrySet()) {
                Integer c = colEntry.getKey();
                UniverCell univerCell = colEntry.getValue();
                XSSFCell cell = getOrCreateCell(row, c);
                CellType cellType = ExcelEnumUtil.getCellType(univerCell.getT());
                cell.setCellType(cellType);
                Object v = univerCell.getV();
                if (null != v) {
                    if (CellType.BOOLEAN.equals(cellType)) {
                        if (v instanceof Integer) {
                            cell.setCellValue(((Integer) v) == 1);
                        } else if (v instanceof Boolean) {
                            cell.setCellValue((Boolean) v);
                        } else {
                            throw new UnsupportedOperationException();
                        }
                    } else if (CellType.NUMERIC.equals(cellType)) {
                        if (v instanceof Integer) {
                            cell.setCellValue((int) v);
                        } else if (v instanceof Long) {
                            cell.setCellValue((long) v);
                        } else if (v instanceof Float) {
                            cell.setCellValue((float) v);
                        } else if (v instanceof Double) {
                            cell.setCellValue((double) v);
                        } else {
                            cell.setCellValue(Double.parseDouble((String) v));
                        }
                    } else {
                        cell.setCellValue((String) v);
                    }
                }
            }
        }
    }

    private XSSFRow getOrCreateRow(XSSFSheet sheet, Integer r) {
        XSSFRow row = sheet.getRow(r);
        if (null == row) {
            row = sheet.createRow(r);
        }
        return row;
    }

    private XSSFCell getOrCreateCell(XSSFRow row, Integer c) {
        XSSFCell cell = row.getCell(c);
        if (null == cell) {
            cell = row.createCell(c);
        }
        return cell;
    }
}
