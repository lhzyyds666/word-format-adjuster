import { contextBridge, ipcRenderer } from "electron";
import type {
  ApplyTemplateOptions,
  DesktopApi,
  FilePayload,
  ReferenceInsertOptions,
  TemplateProfile
} from "../shared/types";

const api: DesktopApi = {
  openDocx: () => ipcRenderer.invoke("docx:open"),
  openTemplate: () => ipcRenderer.invoke("docx:open-template"),
  openPreset: () => ipcRenderer.invoke("preset:open"),
  saveDocx: (name: string, data: ArrayBuffer) => ipcRenderer.invoke("docx:save", name, data),
  savePreset: (name: string, profile: TemplateProfile) => ipcRenderer.invoke("preset:save", name, profile),
  extractTemplate: (file: FilePayload) => ipcRenderer.invoke("docx:extract-template", file),
  applyTemplate: (file: FilePayload, profile: TemplateProfile, options: ApplyTemplateOptions) =>
    ipcRenderer.invoke("docx:apply-template", file, profile, options),
  scanReferences: (file: FilePayload) => ipcRenderer.invoke("docx:scan-references", file),
  insertCrossReference: (file: FilePayload, options: ReferenceInsertOptions) =>
    ipcRenderer.invoke("docx:insert-cross-reference", file, options)
};

contextBridge.exposeInMainWorld("wordAdjuster", api);
