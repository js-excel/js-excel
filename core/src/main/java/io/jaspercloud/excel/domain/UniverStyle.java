package io.jaspercloud.excel.domain;

import lombok.Getter;
import lombok.Setter;

@Setter
@Getter
public class UniverStyle {

    //bold 0: false 1: true
    private Integer bl;
    //italic 0: false 1: true
    private Integer it;
    //strikethrough
    private UniverTextDecoration st;
    //underline
    private UniverTextDecoration ul;
    //foreground
    private UniverColorStyle cl;
    //background
    private UniverColorStyle bg;
    //border
    private UniverBorderData bd;
    //horizontalAlignment
    private Integer ht;
    //verticalAlignment
    private Integer vt;
    //wrapStrategy
    private Integer tb;
    //Numfmt
    private UniverNumfmt n;
}
