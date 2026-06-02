# Word 格式调整器

本项目是一个本地 Windows 桌面应用，用于读取 `.docx`、提取模板格式、套用结构化格式，并提供半自动交叉引用入口。

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

## v1 范围

- 支持 `.docx`，不直接支持旧 `.doc`。
- 模板提取只提取结构化格式，不复制模板正文、图片、批注、修订痕迹。
- 交叉引用采用半自动扫描和插入字段，插入后需要在 Word/WPS 中更新域。
