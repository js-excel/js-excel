package io.jaspercloud.excel.support;

import java.lang.annotation.*;

@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Inherited
public @interface ExcelPropertyName {

    String value();
}
