import type { TemplateProfile } from "./types";

export function parseTemplateProfilePreset(content: string): TemplateProfile {
  let parsed: unknown;
  try {
    parsed = JSON.parse(content);
  } catch {
    throw new Error("预设 JSON 解析失败，请选择由本工具保存的模板预设。");
  }

  if (!isTemplateProfile(parsed)) {
    throw new Error("不是有效的模板预设：缺少必要的格式字段。");
  }

  return parsed;
}

function isTemplateProfile(value: unknown): value is TemplateProfile {
  if (!isRecord(value)) return false;
  return (
    typeof value.sourceName === "string" &&
    typeof value.extractedAt === "string" &&
    hasPage(value.page) &&
    hasBody(value.body) &&
    Array.isArray(value.headings) &&
    value.headings.length > 0 &&
    value.headings.every(hasHeading) &&
    hasHeaderFooter(value.headersFooters) &&
    hasTable(value.tables) &&
    hasCaption(value.captions) &&
    hasNumbering(value.numbering) &&
    Array.isArray(value.warnings) &&
    value.warnings.every((item) => typeof item === "string") &&
    hasRawParts(value.rawParts)
  );
}

function hasPage(value: unknown): boolean {
  return (
    hasRecord(value) &&
    hasNumber(value, "widthTwips") &&
    hasNumber(value, "heightTwips") &&
    isOneOf(value.orientation, ["portrait", "landscape"]) &&
    hasMargins(value.margins)
  );
}

function hasBody(value: unknown): boolean {
  return hasRecord(value) && hasTextStyle(value.text) && hasParagraphStyle(value.paragraph);
}

function hasHeading(value: unknown): boolean {
  return (
    hasTextStyle(value) &&
    hasParagraphStyle(value) &&
    hasRecord(value) &&
    hasNumber(value, "outlineLevel") &&
    isOneOf(value.level, [1, 2, 3, 4])
  );
}

function hasTextStyle(value: unknown): boolean {
  return (
    hasRecord(value) &&
    typeof value.eastAsiaFont === "string" &&
    typeof value.asciiFont === "string" &&
    hasNumber(value, "sizeHalfPoints") &&
    typeof value.color === "string" &&
    (value.bold === undefined || typeof value.bold === "boolean") &&
    (value.italic === undefined || typeof value.italic === "boolean")
  );
}

function hasParagraphStyle(value: unknown): boolean {
  return (
    hasRecord(value) &&
    hasParagraphBase(value) &&
    hasNumber(value, "firstLineTwips") &&
    hasNumber(value, "hangingTwips") &&
    hasNumber(value, "leftTwips") &&
    hasNumber(value, "rightTwips") &&
    typeof value.keepLines === "boolean" &&
    typeof value.keepNext === "boolean"
  );
}

function hasParagraphBase(value: unknown): value is Record<string, unknown> {
  return (
    hasRecord(value) &&
    isOneOf(value.alignment, ["left", "center", "right", "both"]) &&
    hasNumber(value, "line") &&
    hasNumber(value, "before") &&
    hasNumber(value, "after")
  );
}

function hasMargins(value: unknown): boolean {
  return (
    hasRecord(value) &&
    hasNumber(value, "top") &&
    hasNumber(value, "right") &&
    hasNumber(value, "bottom") &&
    hasNumber(value, "left") &&
    hasNumber(value, "header") &&
    hasNumber(value, "footer") &&
    hasNumber(value, "gutter")
  );
}

function hasHeaderFooter(value: unknown): boolean {
  return (
    hasRecord(value) &&
    typeof value.differentFirstPage === "boolean" &&
    typeof value.oddEvenPages === "boolean" &&
    hasTextStyle(value.font) &&
    hasParagraphBase(value.paragraph) &&
    hasRecord(value.headerContent) &&
    isOneOf(value.headerContent.mode, ["preserve", "empty", "staticText", "styleRef"]) &&
    typeof value.headerContent.text === "string" &&
    typeof value.headerContent.styleRef === "string" &&
    isOneOf(value.pageNumberFormat, ["decimal", "lowerRoman", "upperRoman"]) &&
    typeof value.hasTopBorder === "boolean" &&
    typeof value.hasBottomBorder === "boolean"
  );
}

function hasTable(value: unknown): boolean {
  return (
    hasRecord(value) &&
    hasTextStyle(value.font) &&
    typeof value.headerBold === "boolean" &&
    isOneOf(value.alignment, ["left", "center"]) &&
    hasNumber(value, "cellMarginTwips") &&
    typeof value.repeatHeaderRow === "boolean" &&
    typeof value.borderColor === "string"
  );
}

function hasCaption(value: unknown): boolean {
  return (
    hasRecord(value) &&
    hasTextStyle(value.font) &&
    hasParagraphBase(value.paragraph) &&
    Array.isArray(value.labels) &&
    value.labels.every((item) => typeof item === "string")
  );
}

function hasNumbering(value: unknown): boolean {
  return (
    hasRecord(value) &&
    typeof value.hasNumberingXml === "boolean" &&
    hasNumber(value, "abstractNumCount") &&
    hasNumber(value, "numCount")
  );
}

function hasRawParts(value: unknown): boolean {
  if (value === undefined) return true;
  if (!hasRecord(value)) return false;
  return (
    optionalString(value.stylesXml) &&
    optionalString(value.settingsXml) &&
    optionalStringRecord(value.headerFooterParts) &&
    optionalHeaderFooterRelationships(value.headerFooterRelationships) &&
    optionalStringArray(value.headerFooterContentTypeOverrides) &&
    optionalSectionHeaderFooterReferences(value.sectionHeaderFooterReferences) &&
    optionalFrontMatterParagraphs(value.frontMatterParagraphs) &&
    optionalBodyHeadingParagraphs(value.bodyHeadingParagraphs)
  );
}

function optionalString(value: unknown): boolean {
  return value === undefined || typeof value === "string";
}

function optionalStringArray(value: unknown): boolean {
  return value === undefined || (Array.isArray(value) && value.every((item) => typeof item === "string"));
}

function optionalStringRecord(value: unknown): boolean {
  return value === undefined || (hasRecord(value) && Object.values(value).every((item) => typeof item === "string"));
}

function optionalHeaderFooterRelationships(value: unknown): boolean {
  return (
    value === undefined ||
    (Array.isArray(value) &&
      value.every(
        (item) =>
          hasRecord(item) &&
          typeof item.id === "string" &&
          isOneOf(item.type, ["header", "footer"]) &&
          typeof item.target === "string"
      ))
  );
}

function optionalSectionHeaderFooterReferences(value: unknown): boolean {
  return (
    value === undefined ||
    (Array.isArray(value) &&
      value.every(
        (section) =>
          Array.isArray(section) &&
          section.every(
            (item) =>
              hasRecord(item) &&
              isOneOf(item.kind, ["headerReference", "footerReference"]) &&
              typeof item.type === "string" &&
              typeof item.relationshipId === "string"
          )
      ))
  );
}

function optionalFrontMatterParagraphs(value: unknown): boolean {
  return (
    value === undefined ||
    (Array.isArray(value) &&
      value.every(
        (item) =>
          hasRecord(item) &&
          isOneOf(item.match, ["index", "text"]) &&
          (item.index === undefined || typeof item.index === "number") &&
          (item.textKey === undefined || typeof item.textKey === "string") &&
          optionalString(item.pPrXml) &&
          Array.isArray(item.runRPrXmls) &&
          item.runRPrXmls.every((rPr) => rPr === null || typeof rPr === "string")
      ))
  );
}

function optionalBodyHeadingParagraphs(value: unknown): boolean {
  return (
    value === undefined ||
    (Array.isArray(value) &&
      value.every(
        (item) =>
          hasRecord(item) &&
          typeof item.styleId === "string" &&
          optionalString(item.pPrXml) &&
          Array.isArray(item.runRPrXmls) &&
          item.runRPrXmls.every((rPr) => rPr === null || typeof rPr === "string")
      ))
  );
}

function hasNumber(value: Record<string, unknown>, key: string): boolean {
  return typeof value[key] === "number" && Number.isFinite(value[key]);
}

function isOneOf<T extends string | number>(value: unknown, allowed: readonly T[]): value is T {
  return (allowed as readonly unknown[]).includes(value);
}

function hasRecord(value: unknown): value is Record<string, unknown> {
  return isRecord(value);
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}
