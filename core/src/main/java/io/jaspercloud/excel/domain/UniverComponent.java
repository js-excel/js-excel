package io.jaspercloud.excel.domain;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

import java.util.ArrayList;
import java.util.List;

@AllArgsConstructor
@NoArgsConstructor
@Setter
@Getter
public class UniverComponent {

    private String uid;
    private String type;
    private String formula1;
    private String formula2;
    private List<UniverCellRange> ranges = new ArrayList<>();
}
