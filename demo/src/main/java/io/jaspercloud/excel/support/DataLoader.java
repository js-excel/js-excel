package io.jaspercloud.excel.support;

import cn.hutool.core.text.csv.CsvParser;
import cn.hutool.core.text.csv.CsvReadConfig;
import cn.hutool.core.text.csv.CsvRow;
import org.springframework.util.FileCopyUtils;

import java.io.File;
import java.io.StringReader;
import java.util.ArrayList;
import java.util.List;

public class DataLoader {

    public static byte[] excelTemplate() throws Exception {
        byte[] excelBytes = FileCopyUtils.copyToByteArray(new File(System.getProperty("user.dir"), "template/demo.xlsx"));
        return excelBytes;
    }

    public static List<ExcelUserInfo> csvData() throws Exception {
        byte[] csvBytes = FileCopyUtils.copyToByteArray(new File(System.getProperty("user.dir"), "template/demo.csv"));
        CsvParser csvParser = new CsvParser(new StringReader(new String(csvBytes)),
                CsvReadConfig.defaultConfig());
        List<ExcelUserInfo> list = new ArrayList<>();
        while (csvParser.hasNext()) {
            CsvRow next = csvParser.next();
            ExcelUserInfo userInfo = new ExcelUserInfo();
            userInfo.setName(next.get(0));
            userInfo.setGender(next.get(1));
            userInfo.setPhone(next.get(2));
            userInfo.setJob(next.get(3));
            userInfo.setSystems(next.get(4));
            userInfo.setEnable(Integer.parseInt(next.get(5)));
            list.add(userInfo);
        }
        return list;
    }
}
