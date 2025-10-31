# 复选框

设置F列复选框

```java
UniverWorkbook univerWorkbook = new JsExcel(workbook).render();
univerWorkbook.addComponent("用户信息", ComponentEnum.checkbox, "F3:F65535");
```

# 下拉单选

设置第B、D列下拉单选

```java
UniverWorkbook univerWorkbook = new JsExcel(workbook).render();
univerWorkbook.addComponent("用户信息", ComponentEnum.list, "B3:B65535", "=性别!$A:$A");
univerWorkbook.addComponent("用户信息", ComponentEnum.list, "D3:D65535", "=岗位!$A:$A");
```

设置第E列下拉单选

# 下拉多选
```java
UniverWorkbook univerWorkbook = new JsExcel(workbook).render();
univerWorkbook.addComponent("用户信息", ComponentEnum.listMultiple, "E3:E65535", "=系统!$A:$A");
```