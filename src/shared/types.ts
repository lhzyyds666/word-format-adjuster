export type FilePayload = {
  name: string;
  path?: string;
  data: ArrayBuffer;
};

export type PageSettings = {
  widthTwips: number;
  heightTwips: number;
  orientation: "portrait" | "landscape";
  margins: {
    top: number;
    right: number;
    bottom: number;
    left: number;
    header: number;
    footer: number;
    gutter: number;
  };
};

export type TextStyle = {
  eastAsiaFont: string;
  asciiFont: string;
  sizeHalfPoints: number;
  color: string;
  bold?: boolean;
  italic?: boolean;
};

export type ParagraphStyle = {
  alignment: "left" | "center" | "right" | "both";
  firstLineTwips: number;
  hangingTwips: number;
  leftTwips: number;
  rightTwips: number;
  line: number;
  before: number;
  after: number;
  keepLines: boolean;
  keepNext: boolean;
};

export type HeadingStyle = TextStyle &
  ParagraphStyle & {
    level: 1 | 2 | 3 | 4;
    outlineLevel: number;
  };

export type HeaderFooterProfile = {
  differentFirstPage: boolean;
  oddEvenPages: boolean;
  font: TextStyle;
  paragraph: Pick<ParagraphStyle, "alignment" | "before" | "after" | "line">;
  headerContent: {
    mode: "preserve" | "empty" | "staticText" | "styleRef";
    text: string;
    styleRef: string;
  };
  pageNumberFormat: "decimal" | "lowerRoman" | "upperRoman";
  hasTopBorder: boolean;
  hasBottomBorder: boolean;
};

export type TableProfile = {
  font: TextStyle;
  headerBold: boolean;
  alignment: "left" | "center";
  cellMarginTwips: number;
  repeatHeaderRow: boolean;
  borderColor: string;
};

export type CaptionProfile = {
  font: TextStyle;
  paragraph: Pick<ParagraphStyle, "alignment" | "before" | "after" | "line">;
  labels: string[];
};

export type NumberingProfile = {
  hasNumberingXml: boolean;
  abstractNumCount: number;
  numCount: number;
};

export type TemplateProfile = {
  sourceName: string;
  extractedAt: string;
  page: PageSettings;
  body: {
    text: TextStyle;
    paragraph: ParagraphStyle;
  };
  headings: HeadingStyle[];
  headersFooters: HeaderFooterProfile;
  tables: TableProfile;
  captions: CaptionProfile;
  numbering: NumberingProfile;
  rawParts?: {
    stylesXml?: string;
    settingsXml?: string;
    headerFooterParts?: Record<string, string>;
    headerFooterRelationships?: Array<{
      id: string;
      type: "header" | "footer";
      target: string;
    }>;
    headerFooterContentTypeOverrides?: string[];
    sectionHeaderFooterReferences?: Array<
      Array<{
        kind: "headerReference" | "footerReference";
        type: string;
        relationshipId: string;
      }>
    >;
    frontMatterParagraphs?: Array<{
      match: "index" | "text";
      index?: number;
      textKey?: string;
      pPrXml?: string;
      runRPrXmls: Array<string | null>;
    }>;
    bodyHeadingParagraphs?: Array<{
      styleId: string;
      pPrXml?: string;
      runRPrXmls: Array<string | null>;
    }>;
  };
  warnings: string[];
};

export type ApplyTemplateOptions = {
  page: boolean;
  body: boolean;
  headings: boolean;
  headersFooters: boolean;
  tables: boolean;
  captions: boolean;
  numbering: boolean;
};

export type TransformReport = {
  changedParts: string[];
  changedItems: string[];
  skippedItems: string[];
  warnings: string[];
  updateFieldsRequired: string[];
};

export type FormatDiffSection = "页面" | "正文" | "标题" | "页眉页脚" | "表格" | "题注" | "编号";

export type FormatDiffItem = {
  id: string;
  section: FormatDiffSection;
  label: string;
  current: string;
  expected: string;
};

export type FormatDiffReport = {
  checkedAt: string;
  targetName: string;
  profileName: string;
  items: FormatDiffItem[];
};

export type ReferenceTarget = {
  id: string;
  type: "heading" | "caption" | "bookmark";
  label: string;
  bookmark: string;
  paragraphIndex: number;
};

export type ReferenceInsertOptions = {
  targetBookmark: string;
  display: "text" | "number" | "page";
  withHyperlink: boolean;
};

export type DesktopApi = {
  openDocx: () => Promise<FilePayload | null>;
  openTemplate: () => Promise<FilePayload | null>;
  openPreset: () => Promise<TemplateProfile | null>;
  saveDocx: (name: string, data: ArrayBuffer) => Promise<string | null>;
  savePreset: (name: string, profile: TemplateProfile) => Promise<string | null>;
  extractTemplate: (file: FilePayload) => Promise<TemplateProfile>;
  applyTemplate: (
    file: FilePayload,
    profile: TemplateProfile,
    options: ApplyTemplateOptions
  ) => Promise<{ data: ArrayBuffer; report: TransformReport }>;
  scanReferences: (file: FilePayload) => Promise<ReferenceTarget[]>;
  insertCrossReference: (
    file: FilePayload,
    options: ReferenceInsertOptions
  ) => Promise<{ data: ArrayBuffer; report: TransformReport }>;
};

export const DEFAULT_APPLY_OPTIONS: ApplyTemplateOptions = {
  page: true,
  body: true,
  headings: true,
  headersFooters: true,
  tables: true,
  captions: true,
  numbering: true
};
