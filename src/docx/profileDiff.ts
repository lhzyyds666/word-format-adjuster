import type {
  CaptionProfile,
  FormatDiffItem,
  FormatDiffReport,
  FormatDiffSection,
  HeaderFooterProfile,
  HeadingStyle,
  PageSettings,
  ParagraphStyle,
  TableProfile,
  TemplateProfile,
  TextStyle
} from "../shared/types";

const TWIPS_PER_MM = 56.7;

export function compareTemplateProfiles(target: TemplateProfile, expected: TemplateProfile): FormatDiffReport {
  const items: FormatDiffItem[] = [];
  const add = (id: string, section: FormatDiffSection, label: string, current: string, wanted: string) => {
    if (current !== wanted) items.push({ id, section, label, current, expected: wanted });
  };

  comparePage(target.page, expected.page, add);
  compareTextStyle("body.text", "正文", "正文文字", target.body.text, expected.body.text, add);
  compareParagraphStyle("body.paragraph", "正文", "正文段落", target.body.paragraph, expected.body.paragraph, add);

  for (const expectedHeading of expected.headings) {
    const targetHeading = target.headings.find((heading) => heading.level === expectedHeading.level);
    if (!targetHeading) {
      add(`heading.${expectedHeading.level}`, "标题", `${expectedHeading.level} 级标题`, "缺失", "存在");
      continue;
    }
    compareHeading(targetHeading, expectedHeading, add);
  }

  compareHeaderFooter(target.headersFooters, expected.headersFooters, add);
  compareTable(target.tables, expected.tables, add);
  compareCaption(target.captions, expected.captions, add);
  add(
    "numbering.available",
    "编号",
    "编号体系",
    target.numbering.hasNumberingXml ? "已识别" : "无编号文件",
    expected.numbering.hasNumberingXml ? "已识别" : "无编号文件"
  );

  return {
    checkedAt: new Date().toISOString(),
    targetName: target.sourceName,
    profileName: expected.sourceName,
    items
  };
}

function comparePage(
  current: PageSettings,
  expected: PageSettings,
  add: (id: string, section: FormatDiffSection, label: string, current: string, expected: string) => void
): void {
  add("page.size", "页面", "纸张尺寸", formatPageSize(current), formatPageSize(expected));
  add("page.orientation", "页面", "页面方向", orientationLabel(current.orientation), orientationLabel(expected.orientation));
  add("page.margin.top", "页面", "上边距", formatMm(current.margins.top), formatMm(expected.margins.top));
  add("page.margin.right", "页面", "右边距", formatMm(current.margins.right), formatMm(expected.margins.right));
  add("page.margin.bottom", "页面", "下边距", formatMm(current.margins.bottom), formatMm(expected.margins.bottom));
  add("page.margin.left", "页面", "左边距", formatMm(current.margins.left), formatMm(expected.margins.left));
  add("page.margin.header", "页面", "页眉距", formatMm(current.margins.header), formatMm(expected.margins.header));
  add("page.margin.footer", "页面", "页脚距", formatMm(current.margins.footer), formatMm(expected.margins.footer));
}

function compareHeading(
  current: HeadingStyle,
  expected: HeadingStyle,
  add: (id: string, section: FormatDiffSection, label: string, current: string, expected: string) => void
): void {
  const prefix = `heading.${expected.level}`;
  const labelPrefix = `${expected.level} 级标题`;
  compareTextStyle(prefix, "标题", labelPrefix, current, expected, add);
  compareParagraphStyle(prefix, "标题", labelPrefix, current, expected, add);
  add(`${prefix}.outlineLevel`, "标题", `${labelPrefix} 大纲级别`, String(current.outlineLevel), String(expected.outlineLevel));
}

function compareHeaderFooter(
  current: HeaderFooterProfile,
  expected: HeaderFooterProfile,
  add: (id: string, section: FormatDiffSection, label: string, current: string, expected: string) => void
): void {
  add(
    "headersFooters.mode",
    "页眉页脚",
    "页眉内容",
    headerModeLabel(current.headerContent.mode),
    headerModeLabel(expected.headerContent.mode)
  );
  add("headersFooters.text", "页眉页脚", "固定页眉文字", current.headerContent.text, expected.headerContent.text);
  add("headersFooters.styleRef", "页眉页脚", "章节样式编号", current.headerContent.styleRef, expected.headerContent.styleRef);
  compareTextStyle("headersFooters.font", "页眉页脚", "页眉页脚文字", current.font, expected.font, add);
  add("headersFooters.alignment", "页眉页脚", "页眉页脚对齐", alignmentLabel(current.paragraph.alignment), alignmentLabel(expected.paragraph.alignment));
  add("headersFooters.line", "页眉页脚", "页眉页脚行距", formatLine(current.paragraph.line), formatLine(expected.paragraph.line));
  add("headersFooters.pageNumber", "页眉页脚", "页码格式", current.pageNumberFormat, expected.pageNumberFormat);
  add("headersFooters.firstPage", "页眉页脚", "首页不同", yesNo(current.differentFirstPage), yesNo(expected.differentFirstPage));
  add("headersFooters.oddEven", "页眉页脚", "奇偶页不同", yesNo(current.oddEvenPages), yesNo(expected.oddEvenPages));
  add("headersFooters.topBorder", "页眉页脚", "页眉上边线", yesNo(current.hasTopBorder), yesNo(expected.hasTopBorder));
  add("headersFooters.bottomBorder", "页眉页脚", "页脚下边线", yesNo(current.hasBottomBorder), yesNo(expected.hasBottomBorder));
}

function compareTable(
  current: TableProfile,
  expected: TableProfile,
  add: (id: string, section: FormatDiffSection, label: string, current: string, expected: string) => void
): void {
  compareTextStyle("tables.font", "表格", "表格文字", current.font, expected.font, add);
  add("tables.alignment", "表格", "表格对齐", tableAlignmentLabel(current.alignment), tableAlignmentLabel(expected.alignment));
  add("tables.cellMargin", "表格", "单元格边距", formatMm(current.cellMarginTwips), formatMm(expected.cellMarginTwips));
  add("tables.headerBold", "表格", "表头加粗", yesNo(current.headerBold), yesNo(expected.headerBold));
  add("tables.repeatHeader", "表格", "重复标题行", yesNo(current.repeatHeaderRow), yesNo(expected.repeatHeaderRow));
  add("tables.borderColor", "表格", "边框颜色", formatColor(current.borderColor), formatColor(expected.borderColor));
}

function compareCaption(
  current: CaptionProfile,
  expected: CaptionProfile,
  add: (id: string, section: FormatDiffSection, label: string, current: string, expected: string) => void
): void {
  compareTextStyle("captions.font", "题注", "题注文字", current.font, expected.font, add);
  add("captions.alignment", "题注", "题注对齐", alignmentLabel(current.paragraph.alignment), alignmentLabel(expected.paragraph.alignment));
  add("captions.before", "题注", "题注段前", formatPtFromTwips(current.paragraph.before), formatPtFromTwips(expected.paragraph.before));
  add("captions.after", "题注", "题注段后", formatPtFromTwips(current.paragraph.after), formatPtFromTwips(expected.paragraph.after));
  add("captions.line", "题注", "题注行距", formatLine(current.paragraph.line), formatLine(expected.paragraph.line));
  add("captions.labels", "题注", "题注标签", current.labels.join("、"), expected.labels.join("、"));
}

function compareTextStyle(
  idPrefix: string,
  section: FormatDiffSection,
  labelPrefix: string,
  current: TextStyle,
  expected: TextStyle,
  add: (id: string, section: FormatDiffSection, label: string, current: string, expected: string) => void
): void {
  add(`${idPrefix}.eastAsiaFont`, section, `${labelPrefix}中文字体`, current.eastAsiaFont, expected.eastAsiaFont);
  add(`${idPrefix}.asciiFont`, section, `${labelPrefix}英文字体`, current.asciiFont, expected.asciiFont);
  add(`${idPrefix}.size`, section, `${labelPrefix}字号`, formatHalfPoints(current.sizeHalfPoints), formatHalfPoints(expected.sizeHalfPoints));
  add(`${idPrefix}.color`, section, `${labelPrefix}颜色`, formatColor(current.color), formatColor(expected.color));
  add(`${idPrefix}.bold`, section, `${labelPrefix}加粗`, yesNo(Boolean(current.bold)), yesNo(Boolean(expected.bold)));
  add(`${idPrefix}.italic`, section, `${labelPrefix}斜体`, yesNo(Boolean(current.italic)), yesNo(Boolean(expected.italic)));
}

function compareParagraphStyle(
  idPrefix: string,
  section: FormatDiffSection,
  labelPrefix: string,
  current: ParagraphStyle,
  expected: ParagraphStyle,
  add: (id: string, section: FormatDiffSection, label: string, current: string, expected: string) => void
): void {
  add(`${idPrefix}.alignment`, section, `${labelPrefix}对齐`, alignmentLabel(current.alignment), alignmentLabel(expected.alignment));
  add(`${idPrefix}.firstLine`, section, `${labelPrefix}首行缩进`, formatMm(current.firstLineTwips), formatMm(expected.firstLineTwips));
  add(`${idPrefix}.hanging`, section, `${labelPrefix}悬挂缩进`, formatMm(current.hangingTwips), formatMm(expected.hangingTwips));
  add(`${idPrefix}.left`, section, `${labelPrefix}左缩进`, formatMm(current.leftTwips), formatMm(expected.leftTwips));
  add(`${idPrefix}.right`, section, `${labelPrefix}右缩进`, formatMm(current.rightTwips), formatMm(expected.rightTwips));
  add(`${idPrefix}.before`, section, `${labelPrefix}段前`, formatPtFromTwips(current.before), formatPtFromTwips(expected.before));
  add(`${idPrefix}.after`, section, `${labelPrefix}段后`, formatPtFromTwips(current.after), formatPtFromTwips(expected.after));
  add(`${idPrefix}.line`, section, `${labelPrefix}行距`, formatLine(current.line), formatLine(expected.line));
  add(`${idPrefix}.keepLines`, section, `${labelPrefix}孤行控制`, yesNo(current.keepLines), yesNo(expected.keepLines));
  add(`${idPrefix}.keepNext`, section, `${labelPrefix}与下段同页`, yesNo(current.keepNext), yesNo(expected.keepNext));
}

function formatPageSize(page: PageSettings): string {
  return `${formatMm(page.widthTwips)} x ${formatMm(page.heightTwips)}`;
}

function formatMm(value: number): string {
  return `${formatNumber(value / TWIPS_PER_MM)} mm`;
}

function formatPtFromTwips(value: number): string {
  return `${formatNumber(value / 20)} pt`;
}

function formatHalfPoints(value: number): string {
  return `${formatNumber(value / 2)} pt`;
}

function formatLine(value: number): string {
  return `${formatNumber(value / 240)} 倍`;
}

function formatColor(value: string): string {
  const cleaned = value.replace("#", "").toUpperCase();
  return `#${cleaned || "000000"}`;
}

function formatNumber(value: number): string {
  const rounded = Number(value.toFixed(1));
  return Number.isInteger(rounded) ? String(rounded) : String(rounded);
}

function orientationLabel(value: PageSettings["orientation"]): string {
  return value === "landscape" ? "横向" : "纵向";
}

function alignmentLabel(value: ParagraphStyle["alignment"]): string {
  if (value === "center") return "居中";
  if (value === "right") return "右对齐";
  if (value === "both") return "两端对齐";
  return "左对齐";
}

function tableAlignmentLabel(value: TableProfile["alignment"]): string {
  return value === "center" ? "居中" : "左对齐";
}

function headerModeLabel(value: HeaderFooterProfile["headerContent"]["mode"]): string {
  if (value === "empty") return "空页眉";
  if (value === "staticText") return "固定文字";
  if (value === "styleRef") return "当前一级标题";
  return "保留原页眉";
}

function yesNo(value: boolean): string {
  return value ? "是" : "否";
}
