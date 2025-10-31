package io.jaspercloud.excel;

import cn.hutool.core.io.IoUtil;
import cn.hutool.core.io.copy.StreamCopier;
import cn.hutool.json.JSONObject;
import io.jaspercloud.excel.domain.ComponentEnum;
import io.jaspercloud.excel.domain.UniverWorkbook;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.*;

public class JsExcelTest {

    @Test
    public void test() throws Exception {
        File file = new File(new File(System.getProperty("user.dir")).getParent(), "template/demo.xlsx");
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        try (InputStream inputStream = new FileInputStream(file)) {
            new StreamCopier(IoUtil.DEFAULT_BUFFER_SIZE, file.length()).copy(inputStream, outputStream);
        }
        XSSFWorkbook workbook = new XSSFWorkbook(OPCPackage.open(new ByteArrayInputStream(outputStream.toByteArray())));
        UniverWorkbook univerWorkbook = new JsExcel(workbook).render();
        univerWorkbook.addProtectionRange("用户信息", "1:1");
        univerWorkbook.addProtectionRange("用户信息", "2:2");
        univerWorkbook.addComponent("用户信息", ComponentEnum.list, "B3:B65535", "=性别!$A:$A");
        univerWorkbook.addComponent("用户信息", ComponentEnum.list, "D3:D65535", "=岗位!$A:$A");
        univerWorkbook.addComponent("用户信息", ComponentEnum.checkbox, "E3:E65535");
        JSONObject jsonObject = new JSONObject(univerWorkbook);
        String json = jsonObject.toString().replaceAll("(^\\{)|(\\}$)", "");
        String jsonPretty = jsonObject.toStringPretty();
        System.out.println();
    }
}
