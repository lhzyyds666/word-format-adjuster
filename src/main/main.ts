import { app, BrowserWindow, dialog, ipcMain } from "electron";
import { existsSync } from "node:fs";
import fs from "node:fs/promises";
import path from "node:path";
import started from "electron-squirrel-startup";
import { applyTemplateProfile, extractTemplateProfile } from "../docx/profile";
import { insertCrossReference, scanReferenceTargets } from "../docx/references";
import { parseTemplateProfilePreset } from "../shared/profilePreset";
import type { ApplyTemplateOptions, FilePayload, ReferenceInsertOptions, TemplateProfile } from "../shared/types";

if (started) app.quit();

declare const MAIN_WINDOW_VITE_DEV_SERVER_URL: string | undefined;
declare const MAIN_WINDOW_VITE_NAME: string;

const createWindow = (): void => {
  const mainWindow = new BrowserWindow({
    width: 1360,
    height: 860,
    minWidth: 1120,
    minHeight: 720,
    title: "Word 格式调整器",
    backgroundColor: "#0b0f19",
    autoHideMenuBar: true,
    titleBarStyle: "hidden",
    titleBarOverlay: process.platform === "win32" ? {
      color: "#0d1320",
      symbolColor: "#f8fafc",
      height: 52
    } : true,
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false
    }
  });

  if (MAIN_WINDOW_VITE_DEV_SERVER_URL) {
    mainWindow.loadURL(MAIN_WINDOW_VITE_DEV_SERVER_URL);
  } else {
    const namedRenderer = path.join(__dirname, `../renderer/${MAIN_WINDOW_VITE_NAME}/index.html`);
    const renderer = existsSync(namedRenderer) ? namedRenderer : path.join(__dirname, "../renderer/index.html");
    mainWindow.loadFile(renderer);
  }
};

app.whenReady().then(() => {
  registerIpc();
  createWindow();

  app.on("activate", () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") app.quit();
});

function registerIpc(): void {
  ipcMain.handle("docx:open", async () => openDocxDialog("选择目标 Word 文档"));
  ipcMain.handle("docx:open-template", async () => openDocxDialog("选择模板 Word 文档"));
  ipcMain.handle("docx:save", async (_event, name: string, data: ArrayBuffer) => saveDocxDialog(name, data));
  ipcMain.handle("preset:open", async () => openPresetDialog());
  ipcMain.handle("preset:save", async (_event, name: string, profile: TemplateProfile) => savePresetDialog(name, profile));
  ipcMain.handle("docx:extract-template", async (_event, file: FilePayload) =>
    extractTemplateProfile(file.data, file.name)
  );
  ipcMain.handle(
    "docx:apply-template",
    async (_event, file: FilePayload, profile: TemplateProfile, options: ApplyTemplateOptions) =>
      applyTemplateProfile(file.data, profile, options)
  );
  ipcMain.handle("docx:scan-references", async (_event, file: FilePayload) => scanReferenceTargets(file.data));
  ipcMain.handle("docx:insert-cross-reference", async (_event, file: FilePayload, options: ReferenceInsertOptions) =>
    insertCrossReference(file.data, options)
  );
}

async function openDocxDialog(title: string): Promise<FilePayload | null> {
  const result = await dialog.showOpenDialog({
    title,
    properties: ["openFile"],
    filters: [{ name: "Word 文档", extensions: ["docx"] }]
  });
  if (result.canceled || !result.filePaths[0]) return null;
  const filePath = result.filePaths[0];
  const buffer = await fs.readFile(filePath);
  return {
    name: path.basename(filePath),
    path: filePath,
    data: buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength)
  };
}

async function saveDocxDialog(name: string, data: ArrayBuffer): Promise<string | null> {
  const defaultName = name.replace(/\.docx$/i, "") + "-formatted.docx";
  const result = await dialog.showSaveDialog({
    title: "保存调整后的 Word 文档",
    defaultPath: defaultName,
    filters: [{ name: "Word 文档", extensions: ["docx"] }]
  });
  if (result.canceled || !result.filePath) return null;
  await fs.writeFile(result.filePath, Buffer.from(data));
  return result.filePath;
}

async function openPresetDialog(): Promise<TemplateProfile | null> {
  const result = await dialog.showOpenDialog({
    title: "导入模板预设",
    properties: ["openFile"],
    filters: [{ name: "模板预设", extensions: ["json"] }]
  });
  if (result.canceled || !result.filePaths[0]) return null;
  const content = await fs.readFile(result.filePaths[0], "utf8");
  return parseTemplateProfilePreset(content);
}

async function savePresetDialog(name: string, profile: TemplateProfile): Promise<string | null> {
  const result = await dialog.showSaveDialog({
    title: "保存模板预设",
    defaultPath: `${name || profile.sourceName}.json`,
    filters: [{ name: "模板预设", extensions: ["json"] }]
  });
  if (result.canceled || !result.filePath) return null;
  await fs.writeFile(result.filePath, JSON.stringify(profile, null, 2), "utf8");
  return result.filePath;
}
