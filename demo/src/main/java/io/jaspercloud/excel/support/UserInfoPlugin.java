package io.jaspercloud.excel.support;

import io.jaspercloud.excel.domain.ComponentEnum;
import io.jaspercloud.excel.domain.Plugin;
import io.jaspercloud.excel.domain.UniverWorkbook;

public class UserInfoPlugin implements Plugin {

    @Override
    public void run(UniverWorkbook univerWorkbook) {
        univerWorkbook.addProtectionRange("用户信息", "1:1");
        univerWorkbook.addProtectionRange("用户信息", "2:2");
        univerWorkbook.addComponent("用户信息", ComponentEnum.list, "B3:B65535", "=性别!$A:$A");
        univerWorkbook.addComponent("用户信息", ComponentEnum.list, "D3:D65535", "=岗位!$A:$A");
        univerWorkbook.addComponent("用户信息", ComponentEnum.listMultiple, "E3:E65535", "=系统!$A:$A");
        univerWorkbook.addComponent("用户信息", ComponentEnum.checkbox, "F3:F65535");
    }
}
