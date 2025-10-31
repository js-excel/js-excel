设置第1、2行不可编辑

```java
UniverWorkbook univerWorkbook = new JsExcel(workbook).render();
univerWorkbook.addProtectionRange("用户信息", "1:1");
univerWorkbook.addProtectionRange("用户信息", "2:2");
```