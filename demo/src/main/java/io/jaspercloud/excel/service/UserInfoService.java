package io.jaspercloud.excel.service;

import cn.hutool.json.JSONArray;
import cn.hutool.json.JSONObject;
import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.support.ExcelTypeEnum;
import io.jaspercloud.excel.JsExcel;
import io.jaspercloud.excel.domain.UniverWorkbook;
import io.jaspercloud.excel.support.DataLoader;
import io.jaspercloud.excel.support.ExcelPropertyName;
import io.jaspercloud.excel.support.ExcelUserInfo;
import io.jaspercloud.excel.support.UserInfoPlugin;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.util.ReflectionUtils;

import javax.validation.ConstraintViolation;
import javax.validation.Validation;
import javax.validation.Validator;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.lang.reflect.Field;
import java.util.*;

@Service
public class UserInfoService {

    private Validator validator = Validation.buildDefaultValidatorFactory().getValidator();

    public Map<String, Object> importData(UniverWorkbook univerWorkbook) throws Exception {
        List<ExcelUserInfo> list = new ArrayList<>();
        JSONArray jsonArray = new JSONArray();
        Map<String, String> headers = parseHeaders(ExcelUserInfo.class);
        EasyExcel.read(new ByteArrayInputStream(univerWorkbook.exportBytes()))
                .excelType(ExcelTypeEnum.XLSX)
                .head(ExcelUserInfo.class)
                .registerReadListener(new AnalysisEventListener<ExcelUserInfo>() {
                    @Override
                    public void invoke(ExcelUserInfo data, AnalysisContext context) {
                        if (null == data.getEnable()) {
                            data.setEnable(0);
                        }
                        list.add(data);
                        Set<ConstraintViolation<ExcelUserInfo>> violationSet = validator.validate(data);
                        Iterator<ConstraintViolation<ExcelUserInfo>> iterator = violationSet.iterator();
                        if (iterator.hasNext()) {
                            ConstraintViolation<ExcelUserInfo> violation = iterator.next();
                            String columnName = violation.getPropertyPath().toString();
                            String cnName = headers.get(columnName);
                            String message = String.format("%s: %s", cnName, violation.getMessage());
                            data.setMsg(message);
                            return;
                        }
                        data.setMsg("导入成功");
                        jsonArray.add(new JSONObject(data));
                    }

                    @Override
                    public void doAfterAllAnalysed(AnalysisContext context) {

                    }
                })
                .sheet("用户信息")
                .headRowNumber(2)
                .doRead();
        UniverWorkbook result = writeExcel(list);
        Map<String, Object> map = new HashMap<>();
        map.put("workbook", result);
        map.put("table", jsonArray);
        return map;
    }

    private UniverWorkbook writeExcel(List<ExcelUserInfo> list) throws Exception {
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        EasyExcel.write(outputStream)
                .withTemplate(new ByteArrayInputStream(DataLoader.excelTemplate()))
                .sheet("用户信息")
                .doWrite(list);
        XSSFWorkbook workbook = new XSSFWorkbook(OPCPackage.open(new ByteArrayInputStream(outputStream.toByteArray())));
        UniverWorkbook univerWorkbook = new JsExcel(workbook).render(new UserInfoPlugin());
        return univerWorkbook;
    }

    private Map<String, String> parseHeaders(Class<ExcelUserInfo> clazz) {
        Map<String, String> map = new TreeMap<>();
        ReflectionUtils.doWithFields(clazz, new ReflectionUtils.FieldCallback() {
            @Override
            public void doWith(Field field) throws IllegalArgumentException, IllegalAccessException {
                ReflectionUtils.makeAccessible(field);
                ExcelPropertyName excelPropertyName = field.getAnnotation(ExcelPropertyName.class);
                String name = field.getName();
                String cnName = excelPropertyName.value();
                map.put(name, cnName);
            }
        });
        return map;
    }
}
