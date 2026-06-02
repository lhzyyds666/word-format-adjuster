import JSZip from "jszip";
import { DOMParser, XMLSerializer } from "@xmldom/xmldom";

export const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
export const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

export type DocxArchive = {
  zip: JSZip;
  readText: (path: string) => Promise<string | null>;
  writeText: (path: string, content: string) => void;
  list: (pattern: RegExp) => string[];
  toArrayBuffer: () => Promise<ArrayBuffer>;
};

export async function openDocxArchive(data: ArrayBuffer): Promise<DocxArchive> {
  const zip = await JSZip.loadAsync(data);
  return {
    zip,
    readText: async (path) => {
      const file = zip.file(path);
      return file ? file.async("string") : null;
    },
    writeText: (path, content) => {
      zip.file(path, content);
    },
    list: (pattern) =>
      Object.keys(zip.files)
        .filter((path) => !zip.files[path].dir)
        .filter((path) => pattern.test(path)),
    toArrayBuffer: async () => zip.generateAsync({ type: "arraybuffer" })
  };
}

export function parseXml(xml: string): Document {
  const parser = new DOMParser();
  const doc = parser.parseFromString(xml, "application/xml") as unknown as Document;
  const error = doc.getElementsByTagName("parsererror")[0];
  if (error) {
    throw new Error(`XML 解析失败: ${error.textContent ?? "unknown parser error"}`);
  }
  return doc;
}

export function serializeXml(doc: Document): string {
  return new XMLSerializer().serializeToString(doc as never);
}

export function serializeElement(element: Element): string {
  return new XMLSerializer().serializeToString(element as never);
}

export function qn(doc: Document, localName: string): Element {
  return doc.createElementNS(W_NS, `w:${localName}`);
}

export function getAttr(element: Element | null | undefined, localName: string): string | null {
  if (!element) return null;
  return element.getAttributeNS(W_NS, localName) ?? element.getAttribute(`w:${localName}`);
}

export function setAttr(element: Element, localName: string, value: string | number | boolean): void {
  element.setAttributeNS(W_NS, `w:${localName}`, String(value));
}

export function first(parent: Document | Element | null | undefined, tagName: string): Element | null {
  if (!parent) return null;
  return parent.getElementsByTagName(`w:${tagName}`)[0] ?? null;
}

export function children(parent: Document | Element, tagName: string): Element[] {
  return Array.from(parent.childNodes).filter(
    (node): node is Element => node.nodeType === 1 && (node as Element).localName === tagName
  );
}

export function directChild(parent: Document | Element | null | undefined, tagName: string): Element | null {
  if (!parent) return null;
  return children(parent, tagName)[0] ?? null;
}

export function descendants(parent: Document | Element, tagName: string): Element[] {
  return Array.from(parent.getElementsByTagName(`w:${tagName}`));
}

export function ensureChild(parent: Element, tagName: string): Element {
  const found = children(parent, tagName)[0];
  if (found) return found;
  const child = qn(parent.ownerDocument, tagName);
  parent.appendChild(child);
  return child;
}

export function ensureFirstChild(parent: Element, tagName: string): Element {
  const found = children(parent, tagName)[0];
  if (found) return found;
  const child = qn(parent.ownerDocument, tagName);
  parent.insertBefore(child, parent.firstChild);
  return child;
}

export function textContent(element: Element): string {
  return descendants(element, "t")
    .map((node) => node.textContent ?? "")
    .join("");
}

export function normalizeStyleName(value: string): string {
  return value.toLowerCase().replace(/\s+/g, " ").trim();
}

export function findStyleByName(stylesDoc: Document, styleName: string): Element | null {
  const wanted = normalizeStyleName(styleName);
  return descendants(stylesDoc, "style").find((style) => {
    const name = children(style, "name")[0];
    return normalizeStyleName(getAttr(name, "val") ?? "") === wanted;
  }) ?? null;
}

export function findParagraphStyleId(stylesDoc: Document, styleName: string): string | null {
  const style = findStyleByName(stylesDoc, styleName);
  return style ? getAttr(style, "styleId") : null;
}

export function ensureRunProperties(parent: Element): Element {
  return ensureFirstChild(parent, "rPr");
}

export function ensureParagraphProperties(parent: Element): Element {
  return ensureFirstChild(parent, "pPr");
}

export function upsertLeaf(parent: Element, tagName: string, attributes: Record<string, string | number | boolean>): Element {
  const node = ensureChild(parent, tagName);
  for (const [key, value] of Object.entries(attributes)) {
    setAttr(node, key, value);
  }
  return node;
}

export function cloneElementInto(targetDoc: Document, source: Element): Element {
  return targetDoc.importNode(source, true) as Element;
}
