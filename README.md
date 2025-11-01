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
    - [x] 前景色
    - [x] 背景色
    - [x] 字号
    - [x] 加粗
    - [x] 倾斜
    - [x] 下划线
    - [x] 删除线
    - [x] 边框
    - [x] 对齐方式
    - [x] 合并单元格
    - [x] [冻结](doc/protection.md)（单元格禁止编辑）
- 行/列
    - [x] 指定宽/高
- 组件
    - [x] [复选框](doc/component.md#复选框)
    - [x] [下拉单选](doc/component.md#下拉单选)
    - [x] [下拉多选](doc/component.md#下拉多选)
- 操作
    - [x] 增/删多行
    - [x] 复制
    - [x] 粘贴
    - [ ] 过滤
- 表格
    - [x] 隐藏sheet表格