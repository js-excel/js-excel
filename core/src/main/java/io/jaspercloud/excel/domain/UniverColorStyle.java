package io.jaspercloud.excel.domain;

import lombok.Getter;
import lombok.Setter;

@Setter
@Getter
public class UniverColorStyle {

    private String rgb;

    public UniverColorStyle() {
    }

    public UniverColorStyle(String rgb) {
        this.rgb = rgb;
    }
}
