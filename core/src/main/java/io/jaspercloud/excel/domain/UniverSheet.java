package io.jaspercloud.excel.domain;

import lombok.Getter;
import lombok.Setter;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;

@Setter
@Getter
public class UniverSheet {

    private String id;
    private String name;
    private Integer rowCount;
    private Integer columnCount;
    private Integer hidden;
    private Map<Integer, Map<Integer, UniverCell>> cellData = new TreeMap<>();
    private Map<Integer, UniverRowData> rowData = new TreeMap<>();
    private Map<Integer, UniverColumnData> columnData = new TreeMap<>();
    private List<UniverCellRange> mergeData = new ArrayList<>();
    private UniverColumnHeader columnHeader = new UniverColumnHeader();
    private UniverRowHeader rowHeader = new UniverRowHeader();
    private UniverFreeze freeze = new UniverFreeze(-1, -1, 0, 0);
    private Integer defaultColumnWidth = 88;
    private Integer defaultRowHeight = 24;
    private Integer rightToLeft = 0;
    private Integer scrollLeft = 0;
    private Integer scrollTop = 0;
    private Integer showGridlines = 1;
    private String tabColor = "";
    private Integer zoomRatio = 1;

    public void addCell(int r, int c, UniverCell cell) {
        cellData.computeIfAbsent(r, k -> new TreeMap<>()).put(c, cell);
    }

    public void addRowData(int r, UniverRowData data) {
        rowData.put(r, data);
    }

    public void addColumnData(int c, UniverColumnData data) {
        columnData.put(c, data);
    }

    public void addCellMergeRange(UniverCellRange range) {
        mergeData.add(range);
    }
}
