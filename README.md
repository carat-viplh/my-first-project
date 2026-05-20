# 批量文件修改对比 & 自动修改工具

基于模板 diff，批量扫描子文件并按模板差异自动替换内容。

## 快速启动

```bash
cd file_compare_tool
npm install
npm start
```

浏览器打开：**http://localhost:3000**

## 使用步骤

1. 上传 **原始文件** 与 **修改后模板文件**（txt / docx / pdf；doc 请另存为 docx）
2. 上传 **批量文件** 或选择 **整个文件夹**
3. 点击「开始对比与分析」查看差异行号与批量扫描结果
4. 点击「一键批量自动修改」写回 txt / docx，并下载修改后的文件

## 功能说明

| 功能 | 说明 |
|------|------|
| 文件解析 | mammoth(docx)、pdf-parse(pdf)、原生读取 txt |
| 对比 | `diff` 行级对比，识别新增/删除/修改并标注行号 |
| 批量遍历 | 在子文件中查找模板「旧内容」，输出行号与新旧文本 |
| 自动修改 | txt 按行替换；docx 在 document.xml 中替换；pdf 仅对比 |

## 技术栈

- 前端：HTML + Tailwind CSS + 原生 JavaScript
- 后端：Node.js + Express + multer + diff + mammoth + pdf-parse + jszip
