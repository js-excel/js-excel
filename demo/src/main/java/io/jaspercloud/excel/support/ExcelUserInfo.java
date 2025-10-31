package io.jaspercloud.excel.support;

import com.alibaba.excel.annotation.ExcelProperty;
import lombok.Getter;
import lombok.Setter;

import javax.validation.constraints.NotEmpty;
import javax.validation.constraints.Pattern;

@Setter
@Getter
public class ExcelUserInfo {

    @NotEmpty
    @ExcelPropertyName("姓名")
    @ExcelProperty(index = 0)
    private String name;

    @NotEmpty
    @ExcelPropertyName("性别")
    @ExcelProperty(index = 1)
    private String gender;

    @Pattern(regexp = "^[0-9]{11}$", message = "格式错误")
    @NotEmpty
    @ExcelPropertyName("手机号")
    @ExcelProperty(index = 2)
    private String phone;

    @NotEmpty
    @ExcelPropertyName("岗位")
    @ExcelProperty(index = 3)
    private String job;

    @ExcelPropertyName("是否启用")
    @ExcelProperty(index = 4)
    private Integer enable;

    @ExcelPropertyName("结果")
    @ExcelProperty(index = 5)
    private String msg;
}
