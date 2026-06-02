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

## 远程自动打包

仓库包含 GitHub Actions workflow：`.github/workflows/build.yml`。

- push 到 `main` 或手动运行 workflow 时，会在 GitHub 的 Windows、macOS、Linux runner 上分别执行 `npm ci`、`npm test` 和桌面打包。
- 打包产物不会提交进 git，会作为 Actions artifact 上传，保留 14 天。
- 推送 `v*` 标签时，会额外创建/更新 GitHub Release，并上传三平台产物。
- 优先上传 Electron Forge maker 产物；如果 runner 没有生成 maker 文件，会自动用 `electron-packager` 生成便携版包作为兜底。

发布新版本示例：

```powershell
git tag v0.1.0
git push origin v0.1.0
```

## v1 范围

- 支持 `.docx`，不直接支持旧 `.doc`。
- 模板提取只提取结构化格式，不复制模板正文、图片、批注、修订痕迹。
- 交叉引用采用半自动扫描和插入字段，插入后需要在 Word/WPS 中更新域。
