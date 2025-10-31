package io.jaspercloud.excel.domain;

import lombok.Getter;
import lombok.Setter;

@Setter
@Getter
public class UniverBorderData {

    private UniverBorderStyleData l;
    private UniverBorderStyleData r;
    private UniverBorderStyleData t;
    private UniverBorderStyleData b;
}
