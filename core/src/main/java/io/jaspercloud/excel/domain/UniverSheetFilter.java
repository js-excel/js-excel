package io.jaspercloud.excel.domain;

import lombok.Getter;
import lombok.Setter;

@Setter
@Getter
public class UniverSheetFilter {

    private Integer startRow;
    private Integer endRow;
    private Integer startColumn;
    private Integer endColumn;
}
