# Word 格式调整器

![Release](https://img.shields.io/github/v/release/lhzyyds666/word-format-adjuster)
![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20macOS%20%7C%20Linux-2f6fed)
![License](https://img.shields.io/badge/license-MIT-2ea44f)

一个本地运行的 DOCX 桌面工具，用来整理论文、报告和模板文档的格式。它可以读取目标文档、提取模板的结构化格式，并把页面、正文、标题、页眉页脚、表格、题注和编号规则套用到目标 `.docx` 中。

[下载最新版](https://github.com/lhzyyds666/word-format-adjuster/releases/latest)

## 功能

- 模板提取：从模板 `.docx` 中提取页面设置、正文样式、标题样式、表格样式、题注规则、页眉页脚规则和多级编号配置。
- 一键套用：把提取出的结构化格式应用到目标文档，并生成新的 `.docx`。
- 手动调整：支持页面、字体、字号、颜色、段前段后、行距、缩进、标题层级、页眉页脚、题注和表格等格式项。
- 交叉引用：扫描标题、题注和书签，辅助插入 `REF` / `PAGEREF` 字段。
- 本地处理：文档在本机读取和转换，不需要上传到在线服务。

## 安全边界

- 只支持 `.docx`，暂不直接处理旧版 `.doc`。
- 模板模式默认只提取格式，不复制模板正文内容、图片、批注、修订痕迹或作者隐私信息。
- 插入交叉引用后，需要在 Word/WPS 中更新域，才能刷新最终页码和引用文本。

## 开发

```powershell
npm install
npm run dev
```

## 验证

```powershell
npm test
npm run check
```
