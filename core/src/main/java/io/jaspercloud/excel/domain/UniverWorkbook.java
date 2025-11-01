package io.jaspercloud.excel.domain;

import io.jaspercloud.excel.JsExcel;
import io.jaspercloud.excel.exception.ExcelException;
import io.jaspercloud.excel.util.UniverEnumUtil;
import lombok.Getter;
import lombok.Setter;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayOutputStream;
import java.util.*;

@Setter
@Getter
public class UniverWorkbook {

    private String id = "workbook";
    private String appVersion = "0.10.12";
    private String name = "UniverSheet";
    private String locale = "zhCN";
    private Map<String, UniverSheet> sheets = new HashMap<>();
    private Map<String, UniverStyle> styles = new HashMap<>();
    private Map<String, List<UniverProtectionRange>> protectionRanges = new HashMap<>();
    private Map<String, List<UniverComponent>> components = new HashMap<>();
    private Map<String, UniverSheetFilter> filters = new HashMap<>();

    public XSSFWorkbook export() {
        XSSFWorkbook workbook = new XSSFWorkbook();
        new JsExcel(workbook).parse(this);
        return workbook;
    }

    public byte[] exportBytes() throws Exception {
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        export().write(outputStream);
        return outputStream.toByteArray();
    }

    public void addSheets(UniverSheet sheet) {
        sheets.put(sheet.getId(), sheet);
    }

    public UniverStyle getOrCreateStyle(String id) {
        UniverStyle univerStyle = styles.computeIfAbsent(id, k -> new UniverStyle());
        return univerStyle;
    }

    public void addProtectionRange(String sheetId, UniverProtectionRange range) {
        List<UniverProtectionRange> rangeList = protectionRanges.computeIfAbsent(sheetId, k -> new ArrayList<>());
        rangeList.add(range);
    }

    public void addSheetFilter(String sheetId, UniverSheetFilter filter) {
        filters.put(sheetId, filter);
    }

    public void addProtectionRange(String sheetName, String range) {
        UniverSheet univerSheet = sheets.values().stream().filter(e -> StringUtils.equals(e.getName(), sheetName)).findAny()
                .orElseThrow(() -> new ExcelException("not found sheet: " + sheetName));
        List<UniverProtectionRange> rangeList = protectionRanges.computeIfAbsent(univerSheet.getId(), k -> new ArrayList<>());
        rangeList.add(new UniverProtectionRange(range, true));
    }

    public void addComponent(String sheetName, ComponentEnum type, String componentRange, String dataRange) {
        UniverSheet univerSheet = sheets.values().stream().filter(e -> StringUtils.equals(e.getName(), sheetName)).findAny()
                .orElseThrow(() -> new ExcelException("not found sheet: " + sheetName));
        List<UniverComponent> rangeList = components.computeIfAbsent(univerSheet.getId(), k -> new ArrayList<>());
        CellRangeAddress rangeAddress = CellRangeAddress.valueOf(componentRange);
        UniverCellRange range = new UniverCellRange();
        range.setRangeType(0);
        range.setStartRow(rangeAddress.getFirstRow());
        range.setEndRow(rangeAddress.getLastRow());
        range.setStartColumn(rangeAddress.getFirstColumn());
        range.setEndColumn(rangeAddress.getLastColumn());
        UniverComponent component = new UniverComponent();
        component.setUid(String.valueOf(rangeList.size()));
        component.setType(type.name());
        component.setFormula1(dataRange);
        component.setRanges(Arrays.asList(range));
        rangeList.add(component);
    }

    public void addComponent(String sheetName, ComponentEnum type, String componentRange) {
        UniverSheet univerSheet = sheets.values().stream().filter(e -> StringUtils.equals(e.getName(), sheetName)).findAny()
                .orElseThrow(() -> new ExcelException("not found sheet: " + sheetName));
        List<UniverComponent> rangeList = components.computeIfAbsent(univerSheet.getId(), k -> new ArrayList<>());
        CellRangeAddress rangeAddress = CellRangeAddress.valueOf(componentRange);
        UniverCellRange range = new UniverCellRange();
        range.setRangeType(0);
        range.setStartRow(rangeAddress.getFirstRow());
        range.setEndRow(rangeAddress.getLastRow());
        range.setStartColumn(rangeAddress.getFirstColumn());
        range.setEndColumn(rangeAddress.getLastColumn());
        UniverComponent component = new UniverComponent();
        component.setUid(String.valueOf(rangeList.size()));
        component.setType(type.name());
        component.setRanges(Arrays.asList(range));
        if (ComponentEnum.checkbox.equals(type)) {
            component.setFormula1("1");
            component.setFormula2("0");
        }
        rangeList.add(component);
        getSheets().forEach((n, s) -> {
            for (int r : s.getCellData().keySet()) {
                Map<Integer, UniverCell> row = s.getCellData().get(r);
                for (int c : row.keySet()) {
                    UniverCell cell = row.get(c);
                    if (inRange(r, c, range) && null != cell.getS()) {
                        UniverStyle univerStyle = getStyles().get(cell.getS());
                        univerStyle.setCl(new UniverColorStyle("#143EE3"));
                        univerStyle.setHt(UniverEnumUtil.getUniverHorizontalAlignmentCode(HorizontalAlignment.CENTER));
                        univerStyle.setVt(UniverEnumUtil.getUniverVerticalAlignmentCode(VerticalAlignment.CENTER));
                    }
                }
            }
        });
    }

    private boolean inRange(int r, int c, UniverCellRange range) {
        return 1 == 1
                && r >= range.getStartRow()
                && r <= range.getEndRow()
                && c >= range.getStartColumn()
                && c <= range.getEndColumn();
    }
}
