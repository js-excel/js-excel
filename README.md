# js-excel

# 简介

js-excel是一个在线excel数据导入、导出的框架

# 优点

- 支持在web页面上直接编辑数据上传后端
- 简化了office文件编辑数据再上传和下载报错信息的繁琐流程
- js-excel所有流程在web界面完成，web上直接反馈数据导入结果
- 可直接复制粘贴office中的数据到web中

# 依赖引用
```xml
<dependency>
  <groupId>io.jaspercloud.excel</groupId>
  <artifactId>js-excel</artifactId>
  <version>todo</version>
</dependency>
```

## 纯前端演示demo

<p>不含后端代码，前端mock数据<p>

[https://js-excel.github.io](https://js-excel.github.io)

## 前后端完整版本

### 前端

- univer 源码重新构建
- vue3
- element plus

### 后端

- spring boot 2.7.6
- easyExcel

### 目录结构

```
js-excel
 ├─demo
 │  └─src 
 │     └─java 后端excel处理
 │       └─resources
 │          └─static web代码
 └─template 模板
    ├─demo.xlsx 导入导出模板
    └─demo.csv  mock数据
```

## excel支持

- 单元格
    - 前景色
    - 背景色
    - 字号
    - 加粗
    - 倾斜
    - 下划线
    - 删除线
    - 边框
    - 对齐方式
    - 合并单元格
    - 冻结（禁止编辑）
- 行/列
    - 指定宽/高
- 组件
    - 复选框
    - 下拉单选
    - 下拉多选
- 操作
    - 增/删多行
    - 复制
    - 粘贴
    - 过滤 todo
- 表格
    - 隐藏表格