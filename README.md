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
2. 添加 **批量文件**（可多次选择、多选累加，列表中可单独移除）
3. 点击「开始对比与分析」查看差异行号与批量扫描结果
4. 点击「一键批量自动修改」写回 txt / docx，并下载修改后的文件

## 功能说明

| 功能 | 说明 |
|------|------|
| 文件解析 | mammoth(docx)、pdf-parse(pdf)、原生读取 txt |
| 对比 | `diff` 行级对比，识别新增/删除/修改并标注行号 |
| 批量遍历 | 在子文件中查找模板「旧内容」，输出行号与新旧文本 |
| 自动修改 | txt 按行替换；docx 在 document.xml 中替换；pdf 仅对比 |

## 部署到 Vercel（Serverless）

本项目为 Express 单应用，已通过 `api/index.js` 作为无服务器入口，并由 `vercel.json` 将全部请求重写至该函数。

1. 连接 GitHub 仓库并导入项目，**根目录**指向本仓库。
2. 使用默认 **Node 18+**；框架选 Other / 留空即可。
3. **说明**：Vercel 上上传目录使用 `/tmp`，实例间不共享、冷启动后临时文件会清空；适合对比与短时会话。长时间或大文件生产环境建议使用 VPS / Docker 跑 `npm start`。

本地与常见问题：若在 Serverless 中出现 500，请先在 Vercel 项目 **Functions → Logs** 查看具体报错（多为旧部署调用了 `listen`、或写入了非 `/tmp` 目录）。

## 技术栈

- 前端：HTML + Tailwind CSS + 原生 JavaScript
- 后端：Node.js + Express + multer + diff + mammoth + pdf-parse + jszip
