package io.jaspercloud.excel.domain;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@AllArgsConstructor
@NoArgsConstructor
@Setter
@Getter
public class UniverFreeze {

    private Integer startColumn;
    private Integer startRow;
    private Integer xSplit;
    private Integer ySplit;
}
