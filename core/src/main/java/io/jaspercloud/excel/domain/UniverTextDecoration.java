package io.jaspercloud.excel.domain;

import lombok.Getter;
import lombok.Setter;

@Setter
@Getter
public class UniverTextDecoration {

    //show
    private Integer s;
    //color
    private UniverColorStyle cl;
    //TextDecoration @see https://reference.univer.ai/zh-CN/globals#textdecoration
    private Integer t;
}
