package io.jaspercloud.excel.controller;

import cn.hutool.json.JSONUtil;
import com.alibaba.excel.EasyExcel;
import io.jaspercloud.excel.JsExcel;
import io.jaspercloud.excel.domain.UniverWorkbook;
import io.jaspercloud.excel.service.UserInfoService;
import io.jaspercloud.excel.support.DataLoader;
import io.jaspercloud.excel.support.ExcelUserInfo;
import io.jaspercloud.excel.support.UserInfoPlugin;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RestController;

import javax.annotation.Resource;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.List;
import java.util.Map;

@RestController
public class ExcelController {

    @Resource
    private UserInfoService userInfoService;

    @GetMapping("/template")
    public UniverWorkbook data() throws Exception {
        XSSFWorkbook workbook = new XSSFWorkbook(OPCPackage.open(new ByteArrayInputStream(DataLoader.excelTemplate())));
        UniverWorkbook univerWorkbook = new JsExcel(workbook).render(new UserInfoPlugin());
        return univerWorkbook;
    }

    @GetMapping("/sample")
    public UniverWorkbook sample() throws Exception {
        List<ExcelUserInfo> list = DataLoader.csvData();
        ByteArrayOutputStream excelOutputStream = new ByteArrayOutputStream();
        EasyExcel.write(excelOutputStream)
                .withTemplate(new ByteArrayInputStream(DataLoader.excelTemplate()))
                .sheet("用户信息")
                .doWrite(list);
        XSSFWorkbook workbook = new XSSFWorkbook(OPCPackage.open(new ByteArrayInputStream(excelOutputStream.toByteArray())));
        UniverWorkbook univerWorkbook = new JsExcel(workbook).render(new UserInfoPlugin());
        return univerWorkbook;
    }

    @PostMapping("/importData")
    public Map<String, Object> importData(@RequestBody String json) throws Exception {
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook();
        UniverWorkbook univerWorkbook = JSONUtil.parseObj(json).toBean(UniverWorkbook.class);
        new JsExcel(xssfWorkbook).parse(univerWorkbook);
        Map<String, Object> result = userInfoService.importData(univerWorkbook);
        return result;
    }
}
