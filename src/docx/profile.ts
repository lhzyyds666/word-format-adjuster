import {
  ApplyTemplateOptions,
  CaptionProfile,
  HeaderFooterProfile,
  HeadingStyle,
  PageSettings,
  ParagraphStyle,
  TableProfile,
  TemplateProfile,
  TextStyle,
  TransformReport
} from "../shared/types";
import {
  children,
  directChild,
  descendants,
  ensureChild,
  ensureParagraphProperties,
  ensureRunProperties,
  findParagraphStyleId,
  first,
  getAttr,
  openDocxArchive,
  parseXml,
  qn,
  R_NS,
  serializeElement,
  serializeXml,
  setAttr,
  textContent,
  upsertLeaf,
  W_NS
} from "./ooxml";

const DEFAULT_TEXT: TextStyle = {
  eastAsiaFont: "宋体",
  asciiFont: "Times New Roman",
  sizeHalfPoints: 24,
  color: "000000"
};

const DEFAULT_PARAGRAPH: ParagraphStyle = {
  alignment: "both",
  firstLineTwips: 480,
  hangingTwips: 0,
  leftTwips: 0,
  rightTwips: 0,
  line: 360,
  before: 0,
  after: 0,
  keepLines: false,
  keepNext: false
};

export function defaultTemplateProfile(sourceName = "默认中文论文格式"): TemplateProfile {
  const bodyText = { ...DEFAULT_TEXT };
  const bodyParagraph = { ...DEFAULT_PARAGRAPH };
  return {
    sourceName,
    extractedAt: new Date().toISOString(),
    page: {
      widthTwips: 11906,
      heightTwips: 16838,
      orientation: "portrait",
      margins: {
        top: 1440,
        right: 1440,
        bottom: 1440,
        left: 1440,
        header: 720,
        footer: 720,
        gutter: 0
      }
    },
    body: {
      text: bodyText,
      paragraph: bodyParagraph
    },
    headings: [1, 2, 3, 4].map((level) => ({
      ...bodyText,
      ...bodyParagraph,
      level: level as 1 | 2 | 3 | 4,
      outlineLevel: level - 1,
      bold: true,
      sizeHalfPoints: level <= 2 ? 30 : level === 3 ? 28 : 24,
      alignment: level === 1 ? "center" : "left",
      firstLineTwips: 0,
      before: level === 1 ? 360 : 240,
      after: 120,
      keepNext: true
    })),
    headersFooters: {
      differentFirstPage: false,
      oddEvenPages: false,
      font: { ...bodyText, sizeHalfPoints: 18 },
      paragraph: { alignment: "center", before: 0, after: 0, line: 240 },
      headerContent: {
        mode: "preserve",
        text: "中国石油大学（华东）本科毕业设计（论文）",
        styleRef: "1"
      },
      pageNumberFormat: "decimal",
      hasTopBorder: false,
      hasBottomBorder: true
    },
    tables: {
      font: { ...bodyText, sizeHalfPoints: 21 },
      headerBold: true,
      alignment: "center",
      cellMarginTwips: 120,
      repeatHeaderRow: true,
      borderColor: "808080"
    },
    captions: {
      font: { ...bodyText, sizeHalfPoints: 21 },
      paragraph: { alignment: "center", before: 120, after: 120, line: 300 },
      labels: ["图", "表", "公式"]
    },
    numbering: {
      hasNumberingXml: false,
      abstractNumCount: 0,
      numCount: 0
    },
    warnings: []
  };
}

export async function extractTemplateProfile(data: ArrayBuffer, sourceName: string): Promise<TemplateProfile> {
  const archive = await openDocxArchive(data);
  const documentXml = await archive.readText("word/document.xml");
  const stylesXml = await archive.readText("word/styles.xml");
  if (!documentXml || !stylesXml) {
    throw new Error("这不是有效的 .docx：缺少 word/document.xml 或 word/styles.xml");
  }

  const profile = defaultTemplateProfile(sourceName);
  const documentDoc = parseXml(documentXml);
  const stylesDoc = parseXml(stylesXml);
  const settingsXml = await archive.readText("word/settings.xml");
  const relsXml = await archive.readText("word/_rels/document.xml.rels");
  const contentTypesXml = await archive.readText("[Content_Types].xml");
  profile.rawParts = {
    stylesXml,
    settingsXml: settingsXml ?? undefined,
    headerFooterParts: await readHeaderFooterParts(archive),
    headerFooterRelationships: relsXml ? extractHeaderFooterRelationships(relsXml) : [],
    headerFooterContentTypeOverrides: contentTypesXml ? extractHeaderFooterContentTypes(contentTypesXml) : [],
    sectionHeaderFooterReferences: extractSectionHeaderFooterReferences(documentDoc),
    frontMatterParagraphs: extractFrontMatterParagraphBlueprints(documentDoc, stylesDoc),
    bodyHeadingParagraphs: extractBodyHeadingParagraphBlueprints(documentDoc, stylesDoc)
  };

  const sectPr = descendants(documentDoc, "sectPr").at(-1);
  if (sectPr) profile.page = extractPageSettings(sectPr, profile.page);

  const normal = findParagraphStyle(stylesDoc, "Normal") ?? findParagraphStyle(stylesDoc, "正文");
  if (normal) {
    profile.body.text = extractTextStyle(normal, profile.body.text);
    profile.body.paragraph = extractParagraphStyle(normal, profile.body.paragraph);
  }

  profile.headings = profile.headings.map((fallback) => {
    const style =
      findParagraphStyle(stylesDoc, `heading ${fallback.level}`) ??
      findParagraphStyle(stylesDoc, `标题 ${fallback.level}`);
    return style
      ? {
          ...fallback,
          ...extractTextStyle(style, fallback),
          ...extractParagraphStyle(style, fallback)
        }
      : fallback;
  });

  const headerFooterPart = await findRepresentativeHeaderFooterPart(archive);
  if (headerFooterPart) {
    const xml = await archive.readText(headerFooterPart);
    if (xml) {
      profile.headersFooters = extractHeaderFooterProfile(parseXml(xml), profile.headersFooters, stylesDoc);
    }
  }

  profile.tables = extractTableProfile(documentDoc, stylesDoc, profile.tables);
  profile.captions = extractCaptionProfile(documentDoc, stylesDoc, profile.captions);

  const numberingXml = await archive.readText("word/numbering.xml");
  if (numberingXml) {
    const numberingDoc = parseXml(numberingXml);
    profile.numbering = {
      hasNumberingXml: true,
      abstractNumCount: descendants(numberingDoc, "abstractNum").length,
      numCount: descendants(numberingDoc, "num").length
    };
  }

  const warnings = [];
  if (descendants(documentDoc, "ins").length || descendants(documentDoc, "del").length) {
    warnings.push("模板包含修订痕迹，提取时不会迁移修订内容。");
  }
  if (archive.list(/^word\/comments.*\.xml$/).length) {
    warnings.push("模板包含批注，提取时不会迁移批注。");
  }
  profile.warnings = warnings;
  return profile;
}

export async function applyTemplateProfile(
  data: ArrayBuffer,
  profile: TemplateProfile,
  options: ApplyTemplateOptions
): Promise<{ data: ArrayBuffer; report: TransformReport }> {
  const archive = await openDocxArchive(data);
  const documentXml = await archive.readText("word/document.xml");
  const stylesXml = await archive.readText("word/styles.xml");
  if (!documentXml || !stylesXml) {
    throw new Error("这不是有效的 .docx：缺少 word/document.xml 或 word/styles.xml");
  }

  const report: TransformReport = {
    changedParts: [],
    changedItems: [],
    skippedItems: [],
    warnings: [],
    updateFieldsRequired: []
  };

  const documentDoc = parseXml(documentXml);
  const stylesDoc = parseXml(stylesXml);
  let stylesAlreadyReplaced = false;
  const shouldReplaceStylesXml = options.body || options.headings;

  if (shouldReplaceStylesXml && profile.rawParts?.stylesXml && canSafelyReplaceStyles(documentDoc, profile.rawParts.stylesXml)) {
    archive.writeText("word/styles.xml", profile.rawParts.stylesXml);
    stylesAlreadyReplaced = true;
    report.changedItems.push("模板 styles.xml 原样样式表");
  } else if (shouldReplaceStylesXml && profile.rawParts?.stylesXml) {
    report.warnings.push("模板样式表未原样替换：目标文档存在模板中没有的 styleId，已回退到按样式名局部同步。");
  } else if (profile.rawParts?.stylesXml) {
    report.skippedItems.push("未选择正文或标题范围，跳过模板 styles.xml 原样替换。");
  }

  if (profile.rawParts?.settingsXml && (options.page || options.headersFooters)) {
    archive.writeText("word/settings.xml", profile.rawParts.settingsXml);
    report.changedParts.push("word/settings.xml");
    report.changedItems.push("模板 settings.xml 文档设置");
  } else if (profile.rawParts?.settingsXml) {
    report.skippedItems.push("未选择页面或页眉页脚范围，跳过模板 settings.xml 原样替换。");
  }

  if (options.page) {
    for (const sectPr of descendants(documentDoc, "sectPr")) {
      applyPageSettings(sectPr, profile.page);
    }
    report.changedItems.push("页面设置");
  }

  if (options.body) {
    const headingStyleIds = collectHeadingStyleIds(stylesDoc);
    if (!stylesAlreadyReplaced) {
      applyStyleByNames(stylesDoc, ["Normal", "正文"], profile.body.text, profile.body.paragraph);
      applyParagraphFormatting(documentDoc, profile.body.text, profile.body.paragraph, headingStyleIds);
    } else {
      report.skippedItems.push("已原样套用模板 styles.xml，跳过全篇正文直格式覆盖，避免破坏封面和题名页直接格式。");
    }
    report.changedItems.push("正文样式和段落格式");
  }

  if (options.headings) {
    if (!stylesAlreadyReplaced) {
      for (const heading of profile.headings) {
        applyStyleByNames(stylesDoc, [`heading ${heading.level}`, `标题 ${heading.level}`], heading, heading);
      }
    }
    report.changedItems.push("标题 1-4 样式");
  }

  if ((options.body || options.headings) && profile.rawParts?.frontMatterParagraphs?.length) {
    const applied = applyFrontMatterBlueprint(documentDoc, profile.rawParts.frontMatterParagraphs);
    if (applied > 0) {
      report.changedItems.push(`模板前置页格式蓝图（${applied} 段）`);
    }
  }

  if (options.headings && profile.rawParts?.bodyHeadingParagraphs?.length) {
    const activeStylesDoc =
      stylesAlreadyReplaced && profile.rawParts?.stylesXml ? parseXml(profile.rawParts.stylesXml) : stylesDoc;
    const applied = applyBodyHeadingBlueprints(documentDoc, activeStylesDoc, profile.rawParts.bodyHeadingParagraphs);
    if (applied > 0) {
      report.changedItems.push(`模板正文标题段落格式蓝图（${applied} 段）`);
    }
  }

  if (options.headersFooters) {
    await applyHeaderFooterDocumentSettings(archive, profile.headersFooters, report);
    const blueprintApplied = await applyHeaderFooterBlueprint(archive, documentDoc, profile, report);
    if (!blueprintApplied) {
      if (profile.headersFooters.headerContent.mode !== "preserve") {
        await applyManualHeaderContent(archive, documentDoc, profile.headersFooters, report);
      }
      for (const part of archive.list(/^word\/(header|footer)\d+\.xml$/)) {
        const xml = await archive.readText(part);
        if (!xml) continue;
        const partDoc = parseXml(xml);
        applyHeaderFooterFormatting(partDoc, profile.headersFooters);
        archive.writeText(part, serializeXml(partDoc));
        report.changedParts.push(part);
      }
    }
    for (const sectPr of descendants(documentDoc, "sectPr")) {
      applyHeaderFooterSectionSettings(sectPr, profile.headersFooters);
    }
    report.changedItems.push("页眉页脚样式");
  }

  if (options.tables) {
    const activeStylesDoc =
      stylesAlreadyReplaced && profile.rawParts?.stylesXml ? parseXml(profile.rawParts.stylesXml) : stylesDoc;
    const skippedFrontMatterTables = collectFrontMatterTables(documentDoc, activeStylesDoc);
    applyTableFormatting(documentDoc, profile.tables, skippedFrontMatterTables);
    if (skippedFrontMatterTables.size > 0) {
      report.skippedItems.push(`跳过前置页版式表格（${skippedFrontMatterTables.size} 个），避免破坏封面信息栏。`);
    }
    report.changedItems.push("表格基础样式");
  }

  if (options.captions) {
    applyCaptionFormatting(documentDoc, profile.captions);
    report.changedItems.push("题注样式");
    report.updateFieldsRequired.push("题注编号和交叉引用建议在 Word/WPS 中全选后更新域。");
  }

  if (options.numbering) {
    const templateNumbering = profile.numbering.hasNumberingXml;
    if (templateNumbering) {
      report.warnings.push("已识别模板编号体系；v1 只同步标题样式，不直接复制 numbering.xml，避免破坏目标文档编号关系。");
    } else {
      report.skippedItems.push("模板没有 numbering.xml");
    }
  }

  archive.writeText("word/document.xml", serializeXml(documentDoc));
  if (!stylesAlreadyReplaced) {
    archive.writeText("word/styles.xml", serializeXml(stylesDoc));
  }
  report.changedParts.push("word/document.xml", "word/styles.xml");

  return { data: await archive.toArrayBuffer(), report };
}

function findParagraphStyle(stylesDoc: Document, name: string): Element | null {
  const id = findParagraphStyleId(stylesDoc, name);
  if (!id) return null;
  return descendants(stylesDoc, "style").find((style) => getAttr(style, "styleId") === id) ?? null;
}

function extractPageSettings(sectPr: Element, fallback: PageSettings): PageSettings {
  const pgSz = first(sectPr, "pgSz");
  const pgMar = first(sectPr, "pgMar");
  return {
    widthTwips: numberAttr(pgSz, "w", fallback.widthTwips),
    heightTwips: numberAttr(pgSz, "h", fallback.heightTwips),
    orientation: (getAttr(pgSz, "orient") as PageSettings["orientation"]) ?? fallback.orientation,
    margins: {
      top: numberAttr(pgMar, "top", fallback.margins.top),
      right: numberAttr(pgMar, "right", fallback.margins.right),
      bottom: numberAttr(pgMar, "bottom", fallback.margins.bottom),
      left: numberAttr(pgMar, "left", fallback.margins.left),
      header: numberAttr(pgMar, "header", fallback.margins.header),
      footer: numberAttr(pgMar, "footer", fallback.margins.footer),
      gutter: numberAttr(pgMar, "gutter", fallback.margins.gutter)
    }
  };
}

function extractTextStyle(style: Element, fallback: TextStyle): TextStyle {
  const rPr = first(style, "rPr");
  const fonts = first(rPr ?? style, "rFonts");
  const sz = first(rPr ?? style, "sz");
  const color = first(rPr ?? style, "color");
  const bold = directChild(rPr, "b") ?? first(rPr ?? style, "b");
  const italic = directChild(rPr, "i") ?? first(rPr ?? style, "i");
  return {
    eastAsiaFont: getAttr(fonts, "eastAsia") ?? fallback.eastAsiaFont,
    asciiFont: getAttr(fonts, "ascii") ?? getAttr(fonts, "hAnsi") ?? fallback.asciiFont,
    sizeHalfPoints: numberAttr(sz, "val", fallback.sizeHalfPoints),
    color: getAttr(color, "val") ?? fallback.color,
    bold: readOnOff(bold, fallback.bold),
    italic: readOnOff(italic, fallback.italic)
  };
}

function extractParagraphStyle(style: Element, fallback: ParagraphStyle): ParagraphStyle {
  const pPr = first(style, "pPr");
  const jc = first(pPr ?? style, "jc");
  const ind = first(pPr ?? style, "ind");
  const spacing = first(pPr ?? style, "spacing");
  return {
    alignment: (getAttr(jc, "val") as ParagraphStyle["alignment"]) ?? fallback.alignment,
    firstLineTwips: numberAttr(ind, "firstLine", fallback.firstLineTwips),
    hangingTwips: numberAttr(ind, "hanging", fallback.hangingTwips),
    leftTwips: numberAttr(ind, "left", fallback.leftTwips),
    rightTwips: numberAttr(ind, "right", fallback.rightTwips),
    line: numberAttr(spacing, "line", fallback.line),
    before: numberAttr(spacing, "before", fallback.before),
    after: numberAttr(spacing, "after", fallback.after),
    keepLines: readOnOff(first(pPr ?? style, "keepLines"), fallback.keepLines),
    keepNext: readOnOff(first(pPr ?? style, "keepNext"), fallback.keepNext)
  };
}

async function findRepresentativeHeaderFooterPart(
  archive: Awaited<ReturnType<typeof openDocxArchive>>
): Promise<string | undefined> {
  const headerParts = archive.list(/^word\/header\d+\.xml$/);
  let firstPart = headerParts[0];
  let firstTextPart: string | undefined;
  for (const part of headerParts) {
    const xml = await archive.readText(part);
    if (!xml) continue;
    if (/\bSTYLEREF\b/i.test(xml)) return part;
    if (!firstTextPart && />[^<\s][^<]*</.test(xml)) firstTextPart = part;
  }
  return firstTextPart ?? firstPart;
}

function extractHeaderFooterProfile(
  doc: Document,
  fallback: HeaderFooterProfile,
  stylesDoc?: Document
): HeaderFooterProfile {
  const firstRun = descendants(doc, "r")[0];
  const firstParagraph = descendants(doc, "p")[0];
  const pStyleId = getAttr(first(firstParagraph, "pStyle"), "val");
  const paragraphStyle = pStyleId
    ? descendants(stylesDoc ?? doc, "style").find((style) => getAttr(style, "styleId") === pStyleId) ?? null
    : null;
  const instructionText = descendants(doc, "instrText")
    .map((node) => node.textContent ?? "")
    .join(" ");
  const styleRef = /\bSTYLEREF\s+([^\\\s]+)/i.exec(instructionText)?.[1] ?? fallback.headerContent.styleRef;
  const headerText = descendants(doc, "t")
    .map((node) => node.textContent ?? "")
    .join("")
    .trim();
  const styleParagraph = paragraphStyle ? extractParagraphStyle(paragraphStyle, fallback.paragraph as ParagraphStyle) : fallback.paragraph;
  const directParagraph = firstParagraph
    ? extractParagraphStyle(firstParagraph, styleParagraph as ParagraphStyle)
    : styleParagraph;
  return {
    ...fallback,
    font: firstRun ? extractTextStyle(firstRun, fallback.font) : fallback.font,
    paragraph: directParagraph,
    headerContent: {
      mode: /\bSTYLEREF\b/i.test(instructionText) ? "styleRef" : headerText ? "staticText" : "empty",
      text: /\bSTYLEREF\b/i.test(instructionText) ? fallback.headerContent.text : headerText,
      styleRef
    },
    hasTopBorder: hasParagraphBorder(firstParagraph, paragraphStyle, "top", fallback.hasTopBorder),
    hasBottomBorder: hasParagraphBorder(firstParagraph, paragraphStyle, "bottom", fallback.hasBottomBorder)
  };
}

function hasParagraphBorder(
  paragraph: Element | null,
  paragraphStyle: Element | null,
  side: "top" | "bottom",
  fallback: boolean
): boolean {
  const directBorder = first(first(first(paragraph, "pPr"), "pBdr"), side);
  if (directBorder) return getAttr(directBorder, "val") !== "none";
  const styleBorder = first(first(first(paragraphStyle, "pPr"), "pBdr"), side);
  if (styleBorder) return getAttr(styleBorder, "val") !== "none";
  return fallback;
}

function extractTableProfile(doc: Document, stylesDoc: Document, fallback: TableProfile): TableProfile {
  const frontMatterTables = collectFrontMatterTables(doc, stylesDoc);
  const table = descendants(doc, "tbl").find((candidate) => !frontMatterTables.has(candidate));
  if (!table) return fallback;
  const firstRun = descendants(table, "r")[0];
  const tblPr = directChild(table, "tblPr") ?? first(table, "tblPr");
  const jc = directChild(tblPr, "jc") ?? first(tblPr, "jc");
  const firstRow = children(table, "tr")[0] ?? null;
  const firstCell = firstRow ? descendants(firstRow, "tc")[0] ?? null : null;
  const tableMargin = extractCellMarginTwips(directChild(tblPr, "tblCellMar") ?? first(tblPr, "tblCellMar"));
  const cellMargin = extractCellMarginTwips(first(directChild(firstCell, "tcPr"), "tcMar"));
  const firstRowRuns = firstRow ? descendants(firstRow, "r") : [];
  const trPr = directChild(firstRow, "trPr") ?? first(firstRow, "trPr");
  const tblHeader = directChild(trPr, "tblHeader") ?? first(trPr, "tblHeader");

  return {
    ...fallback,
    font: firstRun ? extractTextStyle(firstRun, fallback.font) : fallback.font,
    alignment: extractTableAlignment(jc, fallback.alignment),
    cellMarginTwips: tableMargin ?? cellMargin ?? fallback.cellMarginTwips,
    borderColor: extractBorderColor(directChild(tblPr, "tblBorders") ?? first(tblPr, "tblBorders"), fallback.borderColor),
    headerBold: firstRowRuns.length
      ? firstRowRuns.some((run) => readOnOff(directChild(first(run, "rPr"), "b") ?? first(first(run, "rPr"), "b"), false))
      : fallback.headerBold,
    repeatHeaderRow: firstRow ? readOnOff(tblHeader, false) : fallback.repeatHeaderRow
  };
}

function extractCaptionProfile(
  doc: Document,
  stylesDoc: Document,
  fallback: CaptionProfile
): CaptionProfile {
  const headingStyleIds = collectHeadingStyleIds(stylesDoc);
  const captionItems = descendants(doc, "p")
    .map((p) => {
      const style = getParagraphStyleId(p);
      if (headingStyleIds.has(style) || /heading|标题/i.test(style)) return null;

      const label = extractCaptionLabel(textContent(p).trim());
      return label ? { paragraph: p, label } : null;
    })
    .filter((item): item is { paragraph: Element; label: string } => Boolean(item));

  const foundLabels = new Set<string>(captionItems.map((item) => item.label));

  const labels = foundLabels.size > 0 ? Array.from(foundLabels) : fallback.labels;

  const captionParagraph = captionItems[0]?.paragraph;
  if (!captionParagraph) {
    return { ...fallback, labels };
  }

  const firstRun = descendants(captionParagraph, "r")[0];
  return {
    ...fallback,
    font: firstRun ? extractTextStyle(firstRun, fallback.font) : fallback.font,
    paragraph: extractParagraphStyle(captionParagraph, fallback.paragraph as ParagraphStyle),
    labels
  };
}

function applyPageSettings(sectPr: Element, page: PageSettings): void {
  const pgSz = ensureChild(sectPr, "pgSz");
  setAttr(pgSz, "w", page.widthTwips);
  setAttr(pgSz, "h", page.heightTwips);
  setAttr(pgSz, "orient", page.orientation);

  const pgMar = ensureChild(sectPr, "pgMar");
  for (const [key, value] of Object.entries(page.margins)) {
    setAttr(pgMar, key, value);
  }
}

function applyStyleByNames(stylesDoc: Document, names: string[], text: TextStyle, paragraph: ParagraphStyle): void {
  for (const name of names) {
    const style = findParagraphStyle(stylesDoc, name);
    if (style) {
      applyTextStyle(ensureRunProperties(style), text);
      applyParagraphStyle(ensureParagraphProperties(style), paragraph);
      return;
    }
  }
}

function applyParagraphFormatting(
  doc: Document,
  text: TextStyle,
  paragraph: ParagraphStyle,
  skipStyleIds: Set<string>
): void {
  for (const p of descendants(doc, "p")) {
    const style = getAttr(first(first(p, "pPr") ?? p, "pStyle"), "val") ?? "";
    if (skipStyleIds.has(style) || /heading|标题/i.test(style)) continue;
    applyParagraphStyle(ensureParagraphProperties(p), paragraph);
    for (const r of descendants(p, "r")) {
      applyTextStyle(ensureRunProperties(r), text);
    }
  }
}

function collectHeadingStyleIds(stylesDoc: Document): Set<string> {
  const ids = new Set<string>();
  for (const style of descendants(stylesDoc, "style")) {
    const name = getAttr(first(style, "name"), "val") ?? "";
    const styleId = getAttr(style, "styleId");
    if (styleId && /^(heading|标题)\s*[1-9]/i.test(name)) {
      ids.add(styleId);
    }
  }
  return ids;
}

function canSafelyReplaceStyles(documentDoc: Document, templateStylesXml: string): boolean {
  const templateStylesDoc = parseXml(templateStylesXml);
  const templateStyleIds = new Set(
    descendants(templateStylesDoc, "style")
      .map((style) => getAttr(style, "styleId"))
      .filter((id): id is string => Boolean(id))
  );
  const usedStyleIds = new Set<string>();
  for (const tagName of ["pStyle", "rStyle", "tblStyle"]) {
    for (const styleRef of descendants(documentDoc, tagName)) {
      const styleId = getAttr(styleRef, "val");
      if (styleId) usedStyleIds.add(styleId);
    }
  }
  return Array.from(usedStyleIds).every((styleId) => templateStyleIds.has(styleId));
}

type FrontMatterParagraphBlueprint = NonNullable<
  NonNullable<TemplateProfile["rawParts"]>["frontMatterParagraphs"]
>[number];

type BodyHeadingParagraphBlueprint = NonNullable<
  NonNullable<TemplateProfile["rawParts"]>["bodyHeadingParagraphs"]
>[number];

function extractFrontMatterParagraphBlueprints(
  documentDoc: Document,
  stylesDoc: Document
): FrontMatterParagraphBlueprint[] {
  const paragraphs = bodyParagraphs(documentDoc);
  const firstSectionCount = countFirstSectionParagraphs(documentDoc);
  const headingStyleIds = collectHeadingStyleIds(stylesDoc);
  const firstBodyHeadingIndex = paragraphs.findIndex((paragraph, index) => {
    if (index < firstSectionCount) return false;
    return isBodyHeadingParagraph(paragraph, headingStyleIds);
  });
  const frontMatterLimit =
    firstBodyHeadingIndex >= 0
      ? firstBodyHeadingIndex
      : Math.min(paragraphs.length, Math.max(firstSectionCount, 80));
  const blueprints: FrontMatterParagraphBlueprint[] = [];
  const seen = new Set<string>();

  const addBlueprint = (
    paragraph: Element,
    match: FrontMatterParagraphBlueprint["match"],
    index: number,
    textKey?: string
  ) => {
    const identity = `${match}:${match === "index" ? index : textKey ?? ""}`;
    if (seen.has(identity)) return;
    seen.add(identity);
    blueprints.push(makeParagraphBlueprint(paragraph, match, index, textKey));
  };

  for (let index = 0; index < Math.min(firstSectionCount, paragraphs.length); index += 1) {
    addBlueprint(paragraphs[index], "index", index);
  }

  for (let index = firstSectionCount; index < frontMatterLimit; index += 1) {
    const textKey = normalizeFrontMatterText(textContent(paragraphs[index]));
    if (isKnownFrontMatterHeading(textKey)) {
      addBlueprint(paragraphs[index], "text", index, textKey);
    }

    const previous = paragraphs[index - 1];
    const previousText = previous ? textContent(previous).trim() : "";
    if ((textKey === "摘要" || textKey === "abstract") && previous && previousText.length > 0) {
      addBlueprint(previous, "index", index - 1);
    }
  }

  return blueprints;
}

function extractBodyHeadingParagraphBlueprints(
  documentDoc: Document,
  stylesDoc: Document
): BodyHeadingParagraphBlueprint[] {
  const paragraphs = bodyParagraphs(documentDoc);
  const bodyStart = findBodyStartParagraphIndex(documentDoc, stylesDoc);
  if (bodyStart < 0) return [];

  const headingStyleIds = collectHeadingStyleIds(stylesDoc);
  const byStyleId = new Map<string, BodyHeadingParagraphBlueprint>();
  for (let index = bodyStart; index < paragraphs.length; index += 1) {
    const paragraph = paragraphs[index];
    const styleId = getParagraphStyleId(paragraph);
    if (!styleId || !headingStyleIds.has(styleId) || textContent(paragraph).trim().length === 0) continue;
    if (!byStyleId.has(styleId)) {
      byStyleId.set(styleId, {
        styleId,
        pPrXml: makeParagraphBlueprint(paragraph, "index", index).pPrXml,
        runRPrXmls: makeParagraphBlueprint(paragraph, "index", index).runRPrXmls
      });
    }
  }
  return Array.from(byStyleId.values());
}

function bodyParagraphs(doc: Document): Element[] {
  const body = first(doc, "body");
  return body ? children(body, "p") : descendants(doc, "p");
}

function findBodyStartParagraphIndex(doc: Document, stylesDoc: Document): number {
  const paragraphs = bodyParagraphs(doc);
  const firstSectionCount = countFirstSectionParagraphs(doc);
  const headingStyleIds = collectHeadingStyleIds(stylesDoc);
  return paragraphs.findIndex((paragraph, index) => {
    if (index < firstSectionCount) return false;
    return isBodyHeadingParagraph(paragraph, headingStyleIds);
  });
}

function countFirstSectionParagraphs(doc: Document): number {
  const body = first(doc, "body");
  if (!body) return 0;

  let count = 0;
  for (const child of Array.from(body.childNodes)) {
    if (child.nodeType !== 1) continue;
    const element = child as Element;
    if (element.localName === "sectPr") break;
    if (element.localName !== "p") continue;

    count += 1;
    const pPr = children(element, "pPr")[0];
    if (pPr && children(pPr, "sectPr")[0]) break;
  }
  return count;
}

function collectFrontMatterTables(doc: Document, stylesDoc: Document): Set<Element> {
  const body = first(doc, "body");
  const tables = new Set<Element>();
  if (!body) return tables;

  const headingStyleIds = collectHeadingStyleIds(stylesDoc);
  const hasBodyHeading = Array.from(body.childNodes).some(
    (child) => child.nodeType === 1 && (child as Element).localName === "p" && isBodyHeadingParagraph(child as Element, headingStyleIds)
  );
  if (!hasBodyHeading) return tables;

  let beforeBody = true;
  for (const child of Array.from(body.childNodes)) {
    if (child.nodeType !== 1) continue;
    const element = child as Element;

    if (beforeBody && element.localName === "tbl") {
      tables.add(element);
      continue;
    }

    if (element.localName === "p" && isBodyHeadingParagraph(element, headingStyleIds)) {
      beforeBody = false;
    }
  }
  return tables;
}

function isBodyHeadingParagraph(paragraph: Element, headingStyleIds: Set<string>): boolean {
  const style = getParagraphStyleId(paragraph);
  return headingStyleIds.has(style) && textContent(paragraph).trim().length > 0;
}

function getParagraphStyleId(paragraph: Element): string {
  return getAttr(first(first(paragraph, "pPr") ?? paragraph, "pStyle"), "val") ?? "";
}

function makeParagraphBlueprint(
  paragraph: Element,
  match: FrontMatterParagraphBlueprint["match"],
  index: number,
  textKey?: string
): FrontMatterParagraphBlueprint {
  const pPr = children(paragraph, "pPr")[0];
  return {
    match,
    index,
    textKey,
    pPrXml: pPr ? serializeElement(pPr) : undefined,
    runRPrXmls: descendants(paragraph, "r").map((run) => {
      const rPr = children(run, "rPr")[0];
      return rPr ? serializeElement(rPr) : null;
    })
  };
}

function normalizeFrontMatterText(text: string): string {
  return text.replace(/\s+/g, "").replace(/：/g, ":").trim().toLowerCase();
}

function isKnownFrontMatterHeading(textKey: string): boolean {
  return (
    textKey === "学位论文原创性声明" ||
    textKey === "学位论文版权使用授权书" ||
    textKey === "摘要" ||
    textKey === "abstract" ||
    textKey === "目录" ||
    textKey.includes("毕业设计") ||
    textKey.startsWith("题目:")
  );
}

function applyFrontMatterBlueprint(doc: Document, blueprints: FrontMatterParagraphBlueprint[]): number {
  const paragraphs = bodyParagraphs(doc);
  const consumed = new Set<number>();
  let applied = 0;

  for (const blueprint of blueprints) {
    const targetIndex =
      blueprint.match === "index"
        ? blueprint.index
        : paragraphs.findIndex(
            (paragraph, index) =>
              !consumed.has(index) && normalizeFrontMatterText(textContent(paragraph)) === blueprint.textKey
          );
    if (targetIndex === undefined || targetIndex < 0 || targetIndex >= paragraphs.length) continue;

    applyParagraphBlueprint(paragraphs[targetIndex], blueprint);
    consumed.add(targetIndex);
    applied += 1;
  }

  return applied;
}

function applyBodyHeadingBlueprints(
  doc: Document,
  stylesDoc: Document,
  blueprints: BodyHeadingParagraphBlueprint[]
): number {
  const paragraphs = bodyParagraphs(doc);
  const bodyStart = findBodyStartParagraphIndex(doc, stylesDoc);
  if (bodyStart < 0) return 0;

  const byStyleId = new Map(blueprints.map((blueprint) => [blueprint.styleId, blueprint]));
  let applied = 0;
  for (let index = bodyStart; index < paragraphs.length; index += 1) {
    const blueprint = byStyleId.get(getParagraphStyleId(paragraphs[index]));
    if (!blueprint) continue;
    applyParagraphBlueprint(paragraphs[index], blueprint);
    applied += 1;
  }
  return applied;
}

function applyParagraphBlueprint(
  paragraph: Element,
  blueprint: FrontMatterParagraphBlueprint | BodyHeadingParagraphBlueprint
): void {
  if (blueprint.pPrXml !== undefined) {
    replacePropertyChild(paragraph, "pPr", blueprint.pPrXml);
  }

  const runs = descendants(paragraph, "r");
  runs.forEach((run, index) => {
    const rPrXml = index < blueprint.runRPrXmls.length ? blueprint.runRPrXmls[index] : blueprint.runRPrXmls[0];
    if (rPrXml !== undefined) {
      replacePropertyChild(run, "rPr", rPrXml);
    }
  });
}

function replacePropertyChild(parent: Element, tagName: "pPr" | "rPr", xml: string | null): void {
  for (const child of children(parent, tagName)) {
    parent.removeChild(child);
  }
  if (!xml) return;

  const replacement = parent.ownerDocument.importNode(parseXml(xml).documentElement, true);
  parent.insertBefore(replacement, parent.firstChild);
}

async function readHeaderFooterParts(
  archive: Awaited<ReturnType<typeof openDocxArchive>>
): Promise<Record<string, string>> {
  const parts: Record<string, string> = {};
  for (const part of archive.list(/^word\/(header|footer)\d+\.xml$/)) {
    const xml = await archive.readText(part);
    if (xml) parts[part] = xml;
  }
  return parts;
}

function extractHeaderFooterRelationships(
  relsXml: string
): NonNullable<NonNullable<TemplateProfile["rawParts"]>["headerFooterRelationships"]> {
  const relsDoc = parseXml(relsXml);
  return Array.from(relsDoc.getElementsByTagName("Relationship"))
    .map((relationship) => {
      const type = relationship.getAttribute("Type") ?? "";
      if (!isHeaderFooterRelationshipType(type)) return null;
      return {
        id: relationship.getAttribute("Id") ?? "",
        type: type.endsWith("/header") ? ("header" as const) : ("footer" as const),
        target: relationship.getAttribute("Target") ?? ""
      };
    })
    .filter((item): item is NonNullable<typeof item> => Boolean(item?.id && item.target));
}

function extractHeaderFooterContentTypes(contentTypesXml: string): string[] {
  return (
    contentTypesXml.match(
      /<Override\b(?=[^>]*PartName="\/word\/(?:header|footer)\d+\.xml")(?=[^>]*ContentType="application\/vnd\.openxmlformats-officedocument\.wordprocessingml\.(?:header|footer)\+xml")[^>]*\/>/g
    ) ?? []
  );
}

function extractSectionHeaderFooterReferences(
  documentDoc: Document
): NonNullable<NonNullable<TemplateProfile["rawParts"]>["sectionHeaderFooterReferences"]> {
  return Array.from(documentDoc.getElementsByTagName("w:sectPr")).map((section) =>
    Array.from(section.childNodes)
      .filter(
        (child): child is Element =>
          child.nodeType === 1 &&
          ((child as Element).localName === "headerReference" || (child as Element).localName === "footerReference")
      )
      .map((child) => ({
        kind: child.localName as "headerReference" | "footerReference",
        type: getAttr(child, "type") || "default",
        relationshipId: child.getAttributeNS(R_NS, "id") ?? child.getAttribute("r:id") ?? ""
      }))
      .filter((ref) => Boolean(ref.relationshipId))
  );
}

function isHeaderFooterRelationshipType(type: string): boolean {
  return (
    type === "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" ||
    type === "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
  );
}

function mergeHeaderFooterContentTypes(contentTypesXml: string, overrides: string[]): string {
  const withoutOld = contentTypesXml.replace(
    /\s*<Override\b(?=[^>]*PartName="\/word\/(?:header|footer)\d+\.xml")(?=[^>]*ContentType="application\/vnd\.openxmlformats-officedocument\.wordprocessingml\.(?:header|footer)\+xml")[^>]*\/>/g,
    ""
  );
  return withoutOld.replace("</Types>", `${overrides.join("")}</Types>`);
}

function applyHeaderFooterFormatting(doc: Document, profile: HeaderFooterProfile): void {
  for (const p of descendants(doc, "p")) {
    const pPr = ensureParagraphProperties(p);
    applyParagraphStyle(pPr, {
      ...DEFAULT_PARAGRAPH,
      ...profile.paragraph
    });
    applyParagraphBorders(pPr, profile.hasTopBorder, profile.hasBottomBorder);
  }
  for (const r of descendants(doc, "r")) {
    applyTextStyle(ensureRunProperties(r), profile.font);
  }
}

function applyHeaderFooterSectionSettings(sectPr: Element, profile: HeaderFooterProfile): void {
  toggleLeaf(sectPr, "titlePg", profile.differentFirstPage);
  upsertLeaf(sectPr, "pgNumType", { fmt: profile.pageNumberFormat });
}

async function applyHeaderFooterDocumentSettings(
  archive: Awaited<ReturnType<typeof openDocxArchive>>,
  profile: HeaderFooterProfile,
  report: TransformReport
): Promise<void> {
  const settingsPath = "word/settings.xml";
  const settingsXml = await archive.readText(settingsPath);
  const settingsDoc = settingsXml
    ? parseXml(settingsXml)
    : parseXml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="${W_NS}"/>`);
  toggleLeaf(settingsDoc.documentElement, "evenAndOddHeaders", profile.oddEvenPages);
  archive.writeText(settingsPath, serializeXml(settingsDoc));
  report.changedParts.push(settingsPath);
}

async function applyManualHeaderContent(
  archive: Awaited<ReturnType<typeof openDocxArchive>>,
  documentDoc: Document,
  profile: HeaderFooterProfile,
  report: TransformReport
): Promise<void> {
  const headerPath = nextHeaderPath(archive);
  archive.writeText(headerPath, createManualHeaderXml(profile));
  await addHeaderContentTypeOverride(archive, headerPath);
  const relationshipId = await addDocumentHeaderRelationship(archive, headerPath.replace(/^word\//, ""));

  for (const section of descendants(documentDoc, "sectPr")) {
    for (const child of Array.from(section.childNodes)) {
      if (child.nodeType === 1 && (child as Element).localName === "headerReference") {
        section.removeChild(child);
      }
    }
    const headerReference = qn(documentDoc, "headerReference");
    setAttr(headerReference, "type", "default");
    headerReference.setAttributeNS(R_NS, "r:id", relationshipId);
    section.insertBefore(headerReference, section.firstChild);
  }

  report.changedItems.push("手动页眉内容");
  report.changedParts.push(headerPath, "word/_rels/document.xml.rels", "[Content_Types].xml");
  if (profile.headerContent.mode === "styleRef") {
    report.updateFieldsRequired.push("页眉章节名使用 STYLEREF 字段，建议在 Word/WPS 中全选后更新域。");
  }
}

function nextHeaderPath(archive: Awaited<ReturnType<typeof openDocxArchive>>): string {
  const max = archive.list(/^word\/header\d+\.xml$/).reduce((current, part) => {
    const match = /header(\d+)\.xml$/.exec(part);
    return match ? Math.max(current, Number.parseInt(match[1], 10)) : current;
  }, 0);
  return `word/header${max + 1}.xml`;
}

async function addDocumentHeaderRelationship(
  archive: Awaited<ReturnType<typeof openDocxArchive>>,
  target: string
): Promise<string> {
  const relsPath = "word/_rels/document.xml.rels";
  const relsXml = await archive.readText(relsPath);
  const relsDoc = relsXml
    ? parseXml(relsXml)
    : parseXml(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'
      );
  const relationships = Array.from(relsDoc.getElementsByTagName("Relationship"));
  const maxRid = relationships.reduce((max, relationship) => {
    const match = /^rId(\d+)$/.exec(relationship.getAttribute("Id") ?? "");
    return match ? Math.max(max, Number.parseInt(match[1], 10)) : max;
  }, 0);
  const id = `rId${maxRid + 1}`;
  const relationship = relsDoc.createElement("Relationship");
  relationship.setAttribute("Id", id);
  relationship.setAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header");
  relationship.setAttribute("Target", target);
  relsDoc.documentElement.appendChild(relationship);
  archive.writeText(relsPath, serializeXml(relsDoc));
  return id;
}

async function addHeaderContentTypeOverride(
  archive: Awaited<ReturnType<typeof openDocxArchive>>,
  headerPath: string
): Promise<void> {
  const contentTypesPath = "[Content_Types].xml";
  const partName = `/${headerPath}`;
  const override =
    `<Override PartName="${partName}" ` +
    'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>';
  const contentTypesXml =
    (await archive.readText(contentTypesPath)) ??
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>';
  if (contentTypesXml.includes(`PartName="${partName}"`)) return;
  archive.writeText(contentTypesPath, contentTypesXml.replace("</Types>", `${override}</Types>`));
}

function createManualHeaderXml(profile: HeaderFooterProfile): string {
  const rPr = runPropertiesXml(profile.font);
  const runs =
    profile.headerContent.mode === "empty"
      ? ""
      : profile.headerContent.mode === "styleRef"
        ? styleRefRunsXml(rPr, profile.headerContent.styleRef)
        : `<w:r>${rPr}<w:t>${escapeXml(profile.headerContent.text)}</w:t></w:r>`;
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<w:hdr xmlns:w="${W_NS}">` +
    `<w:p>${paragraphPropertiesXml(profile)}${runs}</w:p>` +
    "</w:hdr>"
  );
}

function paragraphPropertiesXml(profile: HeaderFooterProfile): string {
  const borderXml =
    profile.hasTopBorder || profile.hasBottomBorder
      ? `<w:pBdr>${profile.hasTopBorder ? '<w:top w:val="single" w:color="auto" w:sz="4" w:space="1"/>' : ""}${
          profile.hasBottomBorder ? '<w:bottom w:val="single" w:color="auto" w:sz="4" w:space="1"/>' : ""
        }</w:pBdr>`
      : "";
  return (
    "<w:pPr>" +
    `<w:jc w:val="${profile.paragraph.alignment}"/>` +
    `<w:spacing w:line="${profile.paragraph.line}" w:lineRule="auto" w:before="${profile.paragraph.before}" w:after="${profile.paragraph.after}"/>` +
    borderXml +
    "</w:pPr>"
  );
}

function runPropertiesXml(style: TextStyle): string {
  return (
    "<w:rPr>" +
    `<w:rFonts w:ascii="${escapeXml(style.asciiFont)}" w:hAnsi="${escapeXml(style.asciiFont)}" w:eastAsia="${escapeXml(style.eastAsiaFont)}"/>` +
    `<w:sz w:val="${style.sizeHalfPoints}"/><w:szCs w:val="${style.sizeHalfPoints}"/>` +
    `<w:color w:val="${escapeXml(style.color)}"/>` +
    (style.bold ? "<w:b/>" : "") +
    (style.italic ? "<w:i/>" : "") +
    "</w:rPr>"
  );
}

function styleRefRunsXml(rPr: string, styleRef: string): string {
  const escapedStyleRef = escapeXml(styleRef || "1");
  return (
    `<w:r>${rPr}<w:fldChar w:fldCharType="begin"/></w:r>` +
    `<w:r>${rPr}<w:instrText xml:space="preserve"> STYLEREF ${escapedStyleRef} \\* MERGEFORMAT </w:instrText></w:r>` +
    `<w:r>${rPr}<w:fldChar w:fldCharType="separate"/></w:r>` +
    `<w:r>${rPr}<w:t>当前章节标题</w:t></w:r>` +
    `<w:r>${rPr}<w:fldChar w:fldCharType="end"/></w:r>`
  );
}

function escapeXml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

async function applyHeaderFooterBlueprint(
  archive: Awaited<ReturnType<typeof openDocxArchive>>,
  documentDoc: Document,
  profile: TemplateProfile,
  report: TransformReport
): Promise<boolean> {
  const raw = profile.rawParts;
  if (
    !raw?.headerFooterParts ||
    !raw.headerFooterRelationships?.length ||
    !raw.sectionHeaderFooterReferences?.length
  ) {
    return false;
  }

  let strippedRelationshipContent = false;
  for (const [part, xml] of Object.entries(raw.headerFooterParts)) {
    const sanitized = sanitizeHeaderFooterPartXml(xml);
    strippedRelationshipContent ||= sanitized.stripped;
    archive.writeText(part, sanitized.xml);
  }
  if (strippedRelationshipContent) {
    report.warnings.push("模板页眉页脚包含图片、形状或外部链接，已安全剥离这些关系型内容，避免导出文件出现断开的资源引用。");
  }

  const relsPath = "word/_rels/document.xml.rels";
  const relsXml = await archive.readText(relsPath);
  if (!relsXml) {
    report.warnings.push("目标文档缺少 document.xml.rels，无法迁移页眉页脚关系。");
    return false;
  }
  const relsDoc = parseXml(relsXml);
  const relsRoot = relsDoc.documentElement;
  const existingRelationships = Array.from(relsDoc.getElementsByTagName("Relationship"));
  for (const relationship of existingRelationships) {
    const type = relationship.getAttribute("Type") ?? "";
    if (isHeaderFooterRelationshipType(type) && relationship.parentNode) {
      relationship.parentNode.removeChild(relationship);
    }
  }

  const maxRid = existingRelationships.reduce((max, relationship) => {
    const id = relationship.getAttribute("Id") ?? "";
    const match = /^rId(\d+)$/.exec(id);
    return match ? Math.max(max, Number.parseInt(match[1], 10)) : max;
  }, 0);
  const relationshipIdMap = new Map<string, string>();
  raw.headerFooterRelationships.forEach((relationship, index) => {
    const newId = `rId${maxRid + index + 1}`;
    relationshipIdMap.set(relationship.id, newId);
    const node = relsDoc.createElement("Relationship");
    node.setAttribute("Id", newId);
    node.setAttribute(
      "Type",
      `http://schemas.openxmlformats.org/officeDocument/2006/relationships/${relationship.type}`
    );
    node.setAttribute("Target", relationship.target);
    relsRoot.appendChild(node);
  });
  archive.writeText(relsPath, serializeXml(relsDoc));

  const contentTypesXml = await archive.readText("[Content_Types].xml");
  if (contentTypesXml && raw.headerFooterContentTypeOverrides?.length) {
    archive.writeText("[Content_Types].xml", mergeHeaderFooterContentTypes(contentTypesXml, raw.headerFooterContentTypeOverrides));
  }

  const sections = Array.from(documentDoc.getElementsByTagName("w:sectPr"));
  sections.forEach((section, index) => {
    for (const child of Array.from(section.childNodes)) {
      if (
        child.nodeType === 1 &&
        ((child as Element).localName === "headerReference" || (child as Element).localName === "footerReference")
      ) {
        section.removeChild(child);
      }
    }
    const refs =
      raw.sectionHeaderFooterReferences?.[index] ??
      raw.sectionHeaderFooterReferences?.[raw.sectionHeaderFooterReferences.length - 1] ??
      [];
    const insertBefore = section.firstChild;
    for (const ref of refs) {
      const mappedId = relationshipIdMap.get(ref.relationshipId);
      if (!mappedId) continue;
      const node = qn(documentDoc, ref.kind);
      setAttr(node, "type", ref.type);
      node.setAttributeNS(R_NS, "r:id", mappedId);
      section.insertBefore(node, insertBefore);
    }
  });

  report.changedItems.push("模板页眉页脚 XML parts 和分节引用");
  report.changedParts.push("word/_rels/document.xml.rels", "[Content_Types].xml");
  return true;
}

function sanitizeHeaderFooterPartXml(xml: string): { xml: string; stripped: boolean } {
  const doc = parseXml(xml);
  let stripped = false;

  for (const localName of ["drawing", "pict", "object", "AlternateContent"]) {
    for (const element of Array.from(doc.getElementsByTagName("*")).filter((item) => item.localName === localName)) {
      if (element.parentNode) {
        element.parentNode.removeChild(element);
        stripped = true;
      }
    }
  }

  for (const hyperlink of Array.from(doc.getElementsByTagName("w:hyperlink"))) {
    const relationshipId = hyperlink.getAttributeNS(R_NS, "id") ?? hyperlink.getAttribute("r:id");
    if (!relationshipId || !hyperlink.parentNode) continue;
    while (hyperlink.firstChild) {
      hyperlink.parentNode.insertBefore(hyperlink.firstChild, hyperlink);
    }
    hyperlink.parentNode.removeChild(hyperlink);
    stripped = true;
  }

  return { xml: serializeXml(doc), stripped };
}

function applyTableFormatting(doc: Document, profile: TableProfile, skipTables = new Set<Element>()): void {
  for (const table of descendants(doc, "tbl")) {
    if (skipTables.has(table)) continue;
    const tblPr = ensureChild(table, "tblPr");
    upsertLeaf(tblPr, "jc", { val: profile.alignment });
    applyTableBorders(tblPr, profile.borderColor);
    const rows = children(table, "tr");
    for (const row of rows) {
      const isFirstRow = row === rows[0];
      if (isFirstRow) {
        toggleLeaf(ensureChild(row, "trPr"), "tblHeader", profile.repeatHeaderRow);
      }
      for (const cell of descendants(row, "tc")) {
        const tcPr = ensureChild(cell, "tcPr");
        const tcMar = ensureChild(tcPr, "tcMar");
        for (const side of ["top", "left", "bottom", "right"]) {
          upsertLeaf(tcMar, side, { w: profile.cellMarginTwips, type: "dxa" });
        }
        for (const r of descendants(cell, "r")) {
          applyTextStyle(ensureRunProperties(r), {
            ...profile.font,
            bold: profile.headerBold && isFirstRow ? true : profile.font.bold
          });
        }
      }
    }
  }
}

function applyTableBorders(tblPr: Element, color: string): void {
  const borders = ensureChild(tblPr, "tblBorders");
  for (const side of ["top", "left", "bottom", "right", "insideH", "insideV"]) {
    upsertLeaf(borders, side, { val: "single", sz: 4, space: 0, color });
  }
}

function applyCaptionFormatting(doc: Document, profile: CaptionProfile): void {
  for (const p of descendants(doc, "p")) {
    const text = textContent(p).trim();
    if (!isCaptionText(text, profile.labels)) continue;
    applyParagraphStyle(ensureParagraphProperties(p), {
      ...DEFAULT_PARAGRAPH,
      ...profile.paragraph
    });
    for (const r of descendants(p, "r")) {
      applyTextStyle(ensureRunProperties(r), profile.font);
    }
  }
}

function applyTextStyle(rPr: Element, style: TextStyle): void {
  upsertLeaf(rPr, "rFonts", {
    ascii: style.asciiFont,
    hAnsi: style.asciiFont,
    eastAsia: style.eastAsiaFont
  });
  upsertLeaf(rPr, "sz", { val: style.sizeHalfPoints });
  upsertLeaf(rPr, "szCs", { val: style.sizeHalfPoints });
  upsertLeaf(rPr, "color", { val: style.color });
  toggleLeaf(rPr, "b", Boolean(style.bold));
  toggleLeaf(rPr, "i", Boolean(style.italic));
}

function applyParagraphStyle(pPr: Element, style: ParagraphStyle): void {
  upsertLeaf(pPr, "jc", { val: style.alignment });
  upsertLeaf(pPr, "ind", {
    firstLine: style.firstLineTwips,
    hanging: style.hangingTwips,
    left: style.leftTwips,
    right: style.rightTwips
  });
  upsertLeaf(pPr, "spacing", {
    line: style.line,
    lineRule: "auto",
    before: style.before,
    after: style.after
  });
  toggleLeaf(pPr, "keepLines", style.keepLines);
  toggleLeaf(pPr, "keepNext", style.keepNext);
}

function applyParagraphBorders(pPr: Element, hasTopBorder: boolean, hasBottomBorder: boolean): void {
  const pBdr = ensureChild(pPr, "pBdr");
  toggleBorder(pBdr, "top", hasTopBorder);
  toggleBorder(pBdr, "bottom", hasBottomBorder);
  if (children(pBdr, "top").length === 0 && children(pBdr, "bottom").length === 0 && pBdr.parentNode) {
    pBdr.parentNode.removeChild(pBdr);
  }
}

function toggleBorder(parent: Element, tagName: string, enabled: boolean): void {
  const existing = children(parent, tagName)[0];
  if (!enabled) {
    if (existing?.parentNode) existing.parentNode.removeChild(existing);
    return;
  }
  upsertLeaf(parent, tagName, { val: "single", sz: 4, space: 1, color: "auto" });
}

function toggleLeaf(parent: Element, tagName: string, enabled: boolean): void {
  const existing = children(parent, tagName)[0];
  if (enabled && !existing) parent.appendChild(qn(parent.ownerDocument, tagName));
  if (!enabled && existing && existing.parentNode) existing.parentNode.removeChild(existing);
}

function readOnOff(element: Element | null | undefined, fallback: boolean): boolean;
function readOnOff(element: Element | null | undefined, fallback?: boolean): boolean | undefined;
function readOnOff(element: Element | null | undefined, fallback?: boolean): boolean | undefined {
  if (!element) return fallback;
  const value = (getAttr(element, "val") ?? "").trim().toLowerCase();
  if (!value) return true;
  return !["0", "false", "off", "no"].includes(value);
}

function extractCellMarginTwips(margins: Element | null | undefined): number | null {
  for (const side of ["top", "left", "bottom", "right"]) {
    const value = numberAttr(directChild(margins, side) ?? first(margins, side), "w", Number.NaN);
    if (Number.isFinite(value)) return value;
  }
  return null;
}

function extractBorderColor(borders: Element | null | undefined, fallback: string): string {
  for (const side of ["top", "left", "bottom", "right", "insideH", "insideV"]) {
    const color = getAttr(directChild(borders, side) ?? first(borders, side), "color");
    if (color && color.toLowerCase() !== "auto") return color;
  }
  return fallback;
}

function extractTableAlignment(
  jc: Element | null | undefined,
  fallback: TableProfile["alignment"]
): TableProfile["alignment"] {
  const value = getAttr(jc, "val");
  if (value === "center") return "center";
  if (value === "left" || value === "start") return "left";
  return fallback;
}

const CAPTION_NUMERAL_SOURCE = String.raw`(?:\d+|[一二三四五六七八九十百千万零〇两]+)`;
const CAPTION_NUMBER_SOURCE = String.raw`(?:第\s*)?(?:[（(]?${CAPTION_NUMERAL_SOURCE}(?:[.．\-－—、]\s*${CAPTION_NUMERAL_SOURCE})*[）)]?)`;

function extractCaptionLabel(text: string): string | null {
  const match = new RegExp(String.raw`^([^\d\s]{1,8})\s*${CAPTION_NUMBER_SOURCE}`, "i").exec(text.trim());
  const label = match?.[1]?.trim();
  return label && isLikelyCaptionLabel(label) ? label : null;
}

function isCaptionText(text: string, labels: string[]): boolean {
  const normalized = text.trim();
  return labels.some((label) => {
    const pattern = new RegExp(String.raw`^${escapeRegExp(label)}\s*${CAPTION_NUMBER_SOURCE}`, "i");
    return pattern.test(normalized);
  });
}

function isLikelyCaptionLabel(label: string): boolean {
  return /图|表|式|fig\.?|figure|table|eq\.?|equation/i.test(label);
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function numberAttr(element: Element | null | undefined, name: string, fallback: number): number {
  const value = getAttr(element, name);
  if (!value) return fallback;
  const parsed = Number.parseInt(value, 10);
  return Number.isFinite(parsed) ? parsed : fallback;
}
