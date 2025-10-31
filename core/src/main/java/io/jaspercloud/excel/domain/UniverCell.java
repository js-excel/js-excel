package io.jaspercloud.excel.domain;

import lombok.Getter;
import lombok.Setter;

@Setter
@Getter
public class UniverCell {

    //styleId
    private String s;
    private String f;
    private Object v;
    //cellType
    private Integer t;

    @Override
    public String toString() {
        return "UniverCell{" +
                "v=" + v +
                '}';
    }
}
