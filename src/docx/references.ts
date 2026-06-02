import { ReferenceInsertOptions, ReferenceTarget, TransformReport } from "../shared/types";
import {
  children,
  directChild,
  descendants,
  ensureParagraphProperties,
  first,
  getAttr,
  openDocxArchive,
  parseXml,
  qn,
  serializeElement,
  serializeXml,
  setAttr,
  textContent
} from "./ooxml";

export async function scanReferenceTargets(data: ArrayBuffer): Promise<ReferenceTarget[]> {
  const archive = await openDocxArchive(data);
  const documentXml = await archive.readText("word/document.xml");
  if (!documentXml) throw new Error("这不是有效的 .docx：缺少 word/document.xml");
  const doc = parseXml(documentXml);
  const stylesXml = await archive.readText("word/styles.xml");
  const stylesDoc = stylesXml ? parseXml(stylesXml) : null;
  const headingStyleIds = stylesDoc ? collectReferenceHeadingStyleIds(stylesDoc) : new Set<string>();
  const targets: ReferenceTarget[] = [];
  const paragraphs = descendants(doc, "p");

  paragraphs.forEach((p, index) => {
    const label = textContent(p).trim().replace(/\s+/g, " ");
    if (!label) return;
    const style = getAttr(first(first(p, "pPr") ?? p, "pStyle"), "val") ?? "";
    const bookmarkStart = first(p, "bookmarkStart");
    const bookmark = getAttr(bookmarkStart, "name") ?? makeBookmark(label, index);

    if (isReferenceHeadingStyle(style, headingStyleIds)) {
      targets.push({ id: `heading-${index}`, type: "heading", label, bookmark, paragraphIndex: index });
      return;
    }

    if (/^(图|表|公式)\s*\d+/.test(label)) {
      targets.push({ id: `caption-${index}`, type: "caption", label, bookmark, paragraphIndex: index });
      return;
    }

    if (bookmarkStart) {
      targets.push({ id: `bookmark-${index}`, type: "bookmark", label, bookmark, paragraphIndex: index });
    }
  });

  return targets;
}

function collectReferenceHeadingStyleIds(stylesDoc: Document): Set<string> {
  const ids = new Set<string>();
  for (const style of descendants(stylesDoc, "style")) {
    const styleId = getAttr(style, "styleId");
    if (!styleId) continue;

    const name = getAttr(directChild(style, "name"), "val") ?? "";
    const outlineLevel = Number.parseInt(
      getAttr(directChild(directChild(style, "pPr"), "outlineLvl"), "val") ?? "",
      10
    );
    if (/^(heading|标题)\s*[1-9]/i.test(name) || (Number.isFinite(outlineLevel) && outlineLevel >= 0 && outlineLevel <= 8)) {
      ids.add(styleId);
    }
  }
  return ids;
}

function isReferenceHeadingStyle(styleId: string, headingStyleIds: Set<string>): boolean {
  return headingStyleIds.has(styleId) || /heading|标题/i.test(styleId);
}

export async function insertCrossReference(
  data: ArrayBuffer,
  options: ReferenceInsertOptions
): Promise<{ data: ArrayBuffer; report: TransformReport }> {
  const archive = await openDocxArchive(data);
  const documentXml = await archive.readText("word/document.xml");
  if (!documentXml) throw new Error("这不是有效的 .docx：缺少 word/document.xml");
  const doc = parseXml(documentXml);
  const body = first(doc, "body");
  if (!body) throw new Error("文档缺少 body");

  const bookmarkAdded = ensureBookmarkExists(doc, options.targetBookmark);
  const field = createReferenceField(doc, options);
  const replacedMarker = replaceReferenceMarker(doc, options.targetBookmark, field);

  if (!replacedMarker) {
    const paragraph = qn(doc, "p");
    ensureParagraphProperties(paragraph);
    paragraph.appendChild(field);
    insertParagraphBeforeFinalSection(body, paragraph);
  }

  const instruction = referenceInstruction(options);
  archive.writeText("word/document.xml", serializeXml(doc));
  return {
    data: await archive.toArrayBuffer(),
    report: {
      changedParts: ["word/document.xml"],
      changedItems: [
        ...(bookmarkAdded ? [`补充引用目标书签: ${options.targetBookmark}`] : []),
        replacedMarker ? `替换交叉引用占位符: [[REF:${options.targetBookmark}]]` : `插入交叉引用字段: ${instruction}`
      ],
      skippedItems: [],
      warnings: ["插入后请在 Word/WPS 中更新域，以显示真实引用文字、编号或页码。"],
      updateFieldsRequired: ["交叉引用字段需要更新域。"]
    }
  };
}

function createReferenceField(doc: Document, options: ReferenceInsertOptions): Element {
  const field = qn(doc, "fldSimple");
  setAttr(field, "instr", referenceInstruction(options));
  const fieldRun = qn(doc, "r");
  const text = qn(doc, "t");
  text.textContent = options.display === "page" ? "页码引用" : options.display === "number" ? "编号引用" : "交叉引用";
  fieldRun.appendChild(text);
  field.appendChild(fieldRun);
  return field;
}

function referenceInstruction(options: ReferenceInsertOptions): string {
  if (options.display === "page") {
    return `PAGEREF ${options.targetBookmark}${options.withHyperlink ? " \\h" : ""}`;
  }
  const switches = [options.display === "number" ? "\\r" : "", options.withHyperlink ? "\\h" : ""].filter(Boolean);
  return `REF ${options.targetBookmark}${switches.length ? ` ${switches.join(" ")}` : ""}`;
}

function insertParagraphBeforeFinalSection(body: Element, paragraph: Element): void {
  const bodyChildren = Array.from(body.childNodes);
  let lastSectPr: ChildNode | null = null;
  for (let index = bodyChildren.length - 1; index >= 0; index -= 1) {
    const node = bodyChildren[index];
    if (node.nodeType === 1 && (node as Element).localName === "sectPr") {
      lastSectPr = node;
      break;
    }
  }
  if (lastSectPr) body.insertBefore(paragraph, lastSectPr);
  else body.appendChild(paragraph);
}

function ensureBookmarkExists(doc: Document, bookmark: string): boolean {
  if (descendants(doc, "bookmarkStart").some((item) => getAttr(item, "name") === bookmark)) return false;

  const paragraphs = descendants(doc, "p");
  const target = paragraphs.find((paragraph, index) => {
    const label = textContent(paragraph).trim().replace(/\s+/g, " ");
    return label ? makeBookmark(label, index) === bookmark : false;
  });
  if (!target) throw new Error(`找不到引用目标书签：${bookmark}`);

  addBookmark(target, bookmark, nextBookmarkId(doc));
  return true;
}

function addBookmark(paragraph: Element, name: string, id: number): void {
  const start = qn(paragraph.ownerDocument, "bookmarkStart");
  setAttr(start, "id", id);
  setAttr(start, "name", name);
  const end = qn(paragraph.ownerDocument, "bookmarkEnd");
  setAttr(end, "id", id);

  const paragraphProperties = children(paragraph, "pPr")[0];
  paragraph.insertBefore(start, paragraphProperties?.nextSibling ?? paragraph.firstChild);
  paragraph.appendChild(end);
}

function nextBookmarkId(doc: Document): number {
  const ids = [...descendants(doc, "bookmarkStart"), ...descendants(doc, "bookmarkEnd")]
    .map((item) => Number.parseInt(getAttr(item, "id") ?? "", 10))
    .filter(Number.isFinite);
  return ids.length ? Math.max(...ids) + 1 : 1;
}

function replaceReferenceMarker(doc: Document, bookmark: string, field: Element): boolean {
  const marker = `[[REF:${bookmark}]]`;
  
  const targetParagraph = descendants(doc, "p").find((p) => textContent(p).includes(marker));
  if (!targetParagraph) return false;

  normalizeParagraphRuns(targetParagraph);

  const runs = descendants(targetParagraph, "r");
  const mappings: Array<{
    run: Element;
    tElement: Element | null;
    text: string;
    startIdx: number;
    endIdx: number;
  }> = [];

  let logicalText = "";
  for (const run of runs) {
    const tElement = children(run, "t")[0] ?? null;
    const text = tElement ? (tElement.textContent ?? "") : "";
    const startIdx = logicalText.length;
    logicalText += text;
    const endIdx = logicalText.length;
    mappings.push({ run, tElement, text, startIdx, endIdx });
  }

  const markerIndex = logicalText.indexOf(marker);
  if (markerIndex === -1) return false;
  const markerEndIndex = markerIndex + marker.length;

  const intersecting = mappings.filter(
    (m) => !(m.endIdx <= markerIndex || m.startIdx >= markerEndIndex)
  );
  if (intersecting.length === 0) return false;

  const firstMapping = intersecting[0];
  const lastMapping = intersecting[intersecting.length - 1];
  const parent = firstMapping.run.parentNode;
  if (!parent) return false;

  const beforeText = firstMapping.text.slice(0, markerIndex - firstMapping.startIdx);
  const afterText = lastMapping.text.slice(markerEndIndex - lastMapping.startIdx);

  if (firstMapping.run === lastMapping.run) {
    if (beforeText) {
      if (firstMapping.tElement) {
        firstMapping.tElement.textContent = beforeText;
        if (/^\s|\s$/.test(beforeText)) firstMapping.tElement.setAttribute("xml:space", "preserve");
      }
      parent.insertBefore(field, firstMapping.run.nextSibling);
      if (afterText) {
        const afterRun = createTextRun(doc, firstMapping.run, afterText);
        parent.insertBefore(afterRun, field.nextSibling);
      }
    } else {
      parent.insertBefore(field, firstMapping.run);
      if (afterText) {
        if (firstMapping.tElement) {
          firstMapping.tElement.textContent = afterText;
          if (/^\s|\s$/.test(afterText)) firstMapping.tElement.setAttribute("xml:space", "preserve");
        }
      } else {
        parent.removeChild(firstMapping.run);
      }
    }
  } else {
    if (beforeText) {
      if (firstMapping.tElement) {
        firstMapping.tElement.textContent = beforeText;
        if (/^\s|\s$/.test(beforeText)) firstMapping.tElement.setAttribute("xml:space", "preserve");
      }
      parent.insertBefore(field, firstMapping.run.nextSibling);
    } else {
      parent.insertBefore(field, firstMapping.run);
      parent.removeChild(firstMapping.run);
    }

    for (let i = 1; i < intersecting.length - 1; i++) {
      const run = intersecting[i].run;
      if (run.parentNode) run.parentNode.removeChild(run);
    }

    if (afterText) {
      if (lastMapping.tElement) {
        lastMapping.tElement.textContent = afterText;
        if (/^\s|\s$/.test(afterText)) lastMapping.tElement.setAttribute("xml:space", "preserve");
      }
    } else {
      if (lastMapping.run.parentNode) lastMapping.run.parentNode.removeChild(lastMapping.run);
    }
  }

  return true;
}

function normalizeParagraphRuns(paragraph: Element): void {
  const childNodes = Array.from(paragraph.childNodes);
  let i = 0;
  while (i < childNodes.length - 1) {
    const current = childNodes[i];
    const next = childNodes[i + 1];

    if (
      current.nodeType === 1 &&
      next.nodeType === 1 &&
      (current as Element).localName === "r" &&
      (next as Element).localName === "r"
    ) {
      const currentEl = current as Element;
      const nextEl = next as Element;

      if (isSafeRun(currentEl) && isSafeRun(nextEl)) {
        const currentRPr = children(currentEl, "rPr")[0];
        const nextRPr = children(nextEl, "rPr")[0];

        const currentRPrXml = currentRPr ? serializeElement(currentRPr) : "";
        const nextRPrXml = nextRPr ? serializeElement(nextRPr) : "";

        const currentT = children(currentEl, "t")[0];
        const nextT = children(nextEl, "t")[0];

        if (currentRPrXml === nextRPrXml && currentT && nextT) {
          currentT.textContent = (currentT.textContent ?? "") + (nextT.textContent ?? "");
          if (/^\s|\s$/.test(currentT.textContent)) {
            currentT.setAttribute("xml:space", "preserve");
          }
          paragraph.removeChild(nextEl);
          childNodes.splice(i + 1, 1);
          continue;
        }
      }
    }
    i++;
  }
}

function isSafeRun(run: Element): boolean {
  const unsafeTags = [
    "drawing",
    "fldChar",
    "instrText",
    "bookmarkStart",
    "bookmarkEnd",
    "hyperlink",
    "tab",
    "br",
    "lastRenderedPageBreak"
  ];
  for (const tag of unsafeTags) {
    if (descendants(run, tag).length > 0) return false;
  }
  return true;
}

function createTextRun(doc: Document, sourceRun: Element, text: string): Element {
  const run = qn(doc, "r");
  const runProperties = children(sourceRun, "rPr")[0];
  if (runProperties) run.appendChild(doc.importNode(runProperties, true));
  const textNode = qn(doc, "t");
  if (/^\s|\s$/.test(text)) textNode.setAttribute("xml:space", "preserve");
  textNode.textContent = text;
  run.appendChild(textNode);
  return run;
}

function makeBookmark(label: string, index: number): string {
  const compact = label.replace(/[^\p{L}\p{N}]+/gu, "_").replace(/^_+|_+$/g, "");
  return `wfa_${index}_${compact.slice(0, 24) || "target"}`;
}
