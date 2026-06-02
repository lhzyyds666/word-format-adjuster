import { describe, expect, test } from "vitest";
import JSZip from "jszip";
import { DEFAULT_APPLY_OPTIONS } from "../src/shared/types";
import { applyTemplateProfile, defaultTemplateProfile, extractTemplateProfile } from "../src/docx/profile";

describe("template profile", () => {
  test("extracts structured page and body style without copying template body text", async () => {
    const template = await makeDocx({
      bodyText: "模板里的秘密正文不应该被复制",
      marginTop: "1200",
      bodySize: "26",
      withComments: true
    });

    const profile = await extractTemplateProfile(template, "template.docx");

    expect(profile.page.margins.top).toBe(1200);
    expect(profile.body.text.sizeHalfPoints).toBe(26);
    expect(profile.warnings).toContain("模板包含批注，提取时不会迁移批注。");
    expect(JSON.stringify(profile)).not.toContain("模板里的秘密正文");
  });

  test("extracts style-reference chapter header from a synthetic template", async () => {
    const sample = await makeDocxWithStyleRefHeader();
    const profile = await extractTemplateProfile(sample, "style-ref-template.docx");

    expect(profile.headersFooters.headerContent.mode).toBe("styleRef");
    expect(profile.headersFooters.headerContent.styleRef).toBe("1");
    expect(profile.headersFooters.font.eastAsiaFont).toBe("楷体_GB2312");
    expect(profile.headersFooters.font.sizeHalfPoints).toBe(21);
    expect(profile.headersFooters.paragraph.alignment).toBe("center");
    expect(profile.headersFooters.hasBottomBorder).toBe(true);
  });

  test("applies template styles to target document and preserves target body text", async () => {
    const template = await makeDocx({ bodyText: "模板正文", marginTop: "1200", bodySize: "26" });
    const target = await makeDocx({ bodyText: "目标正文", marginTop: "1440", bodySize: "24" });
    const profile = await extractTemplateProfile(template, "template.docx");

    const result = await applyTemplateProfile(target, profile, DEFAULT_APPLY_OPTIONS);
    const zip = await JSZip.loadAsync(result.data);
    const documentXml = await zip.file("word/document.xml")!.async("string");
    const stylesXml = await zip.file("word/styles.xml")!.async("string");

    expect(documentXml).toContain("目标正文");
    expect(documentXml).not.toContain("模板正文");
    expect(documentXml).toContain('w:top="1200"');
    expect(stylesXml).toContain('w:val="26"');
    expect(result.report.changedItems).toContain("页面设置");
  });

  test("does not apply body direct paragraph formatting to numeric heading style ids", async () => {
    const template = await makeDocx({ bodyText: "模板正文", marginTop: "1200", bodySize: "26", headingStyleId: "2" });
    const target = await makeDocx({ bodyText: "目标正文", marginTop: "1440", bodySize: "24", headingStyleId: "2" });
    const profile = await extractTemplateProfile(template, "template.docx");

    const result = await applyTemplateProfile(target, profile, DEFAULT_APPLY_OPTIONS);
    const zip = await JSZip.loadAsync(result.data);
    const documentXml = await zip.file("word/document.xml")!.async("string");
    const headingParagraph = documentXml.match(/<w:p>\s*<w:pPr><w:pStyle w:val="2"\/>[\s\S]*?<\/w:p>/)?.[0] ?? "";

    expect(headingParagraph).not.toContain('w:firstLine="480"');
  });

  test("applies template direct formatting to body heading paragraphs", async () => {
    const template = await makeDocxWithBodyHeading({
      headingText: "TEMPLATE HEADING",
      headingAlignment: "center",
      headingBefore: "156",
      headingAfter: "156",
      headingLine: "360"
    });
    const target = await makeDocxWithBodyHeading({
      headingText: "TARGET HEADING",
      headingAlignment: "both",
      headingBefore: "160",
      headingAfter: "160",
      headingLine: "240"
    });
    const profile = await extractTemplateProfile(template, "template.docx");

    const result = await applyTemplateProfile(target, profile, DEFAULT_APPLY_OPTIONS);
    const zip = await JSZip.loadAsync(result.data);
    const documentXml = await zip.file("word/document.xml")!.async("string");
    const headingParagraph = documentXml.match(/<w:p>[\s\S]*?TARGET HEADING[\s\S]*?<\/w:p>/)?.[0] ?? "";

    expect(documentXml).toContain("TARGET HEADING");
    expect(documentXml).not.toContain("TEMPLATE HEADING");
    expect(headingParagraph).toContain('w:jc w:val="center"');
    expect(headingParagraph).toContain('w:before="156"');
    expect(headingParagraph).toContain('w:after="156"');
    expect(headingParagraph).toContain('w:line="360"');
  });

  test("copies template styles and settings when all used style ids exist in template", async () => {
    const template = await makeDocx({ bodyText: "模板正文", marginTop: "1200", bodySize: "26", settingsMarker: "template-settings" });
    const target = await makeDocx({ bodyText: "目标正文", marginTop: "1440", bodySize: "24", settingsMarker: "target-settings" });
    const profile = await extractTemplateProfile(template, "template.docx");

    const result = await applyTemplateProfile(target, profile, DEFAULT_APPLY_OPTIONS);
    const templateZip = await JSZip.loadAsync(template);
    const outputZip = await JSZip.loadAsync(result.data);
    const templateStyles = await templateZip.file("word/styles.xml")!.async("string");
    const outputStyles = await outputZip.file("word/styles.xml")!.async("string");
    const outputSettings = await outputZip.file("word/settings.xml")!.async("string");

    expect(outputStyles).toBe(templateStyles);
    expect(outputSettings).toContain("template-settings");
    expect(outputSettings).not.toContain("target-settings");
  });

  test("does not copy raw styles or settings when matching apply ranges are disabled", async () => {
    const template = await makeDocx({ bodyText: "模板正文", marginTop: "1200", bodySize: "26", settingsMarker: "template-settings" });
    const target = await makeDocx({ bodyText: "目标正文", marginTop: "1440", bodySize: "24", settingsMarker: "target-settings" });
    const profile = await extractTemplateProfile(template, "template.docx");

    const result = await applyTemplateProfile(target, profile, {
      page: false,
      body: false,
      headings: false,
      headersFooters: false,
      tables: false,
      captions: false,
      numbering: false
    });
    const outputZip = await JSZip.loadAsync(result.data);
    const outputDocument = await outputZip.file("word/document.xml")!.async("string");
    const outputStyles = await outputZip.file("word/styles.xml")!.async("string");
    const outputSettings = await outputZip.file("word/settings.xml")!.async("string");

    expect(outputDocument).toContain('w:top="1440"');
    expect(outputStyles).toContain('w:val="24"');
    expect(outputStyles).not.toContain('w:val="26"');
    expect(outputSettings).toContain("target-settings");
    expect(outputSettings).not.toContain("template-settings");
  });

  test("copies template header and footer blueprint into the target document", async () => {
    const template = await makeDocx({
      bodyText: "模板正文",
      marginTop: "1200",
      bodySize: "26",
      headerFooterBlueprint: "template"
    });
    const target = await makeDocx({
      bodyText: "目标正文",
      marginTop: "1440",
      bodySize: "24",
      headerFooterBlueprint: "target"
    });
    const profile = await extractTemplateProfile(template, "template.docx");

    const result = await applyTemplateProfile(target, profile, DEFAULT_APPLY_OPTIONS);
    const outputZip = await JSZip.loadAsync(result.data);
    const documentXml = await outputZip.file("word/document.xml")!.async("string");
    const relsXml = await outputZip.file("word/_rels/document.xml.rels")!.async("string");

    expect(outputZip.file("word/header2.xml")).toBeTruthy();
    expect(outputZip.file("word/footer1.xml")).toBeTruthy();
    expect(documentXml).toContain('w:type="even"');
    expect(relsXml).toContain('Target="header2.xml"');
    expect(relsXml).toContain('Target="footer1.xml"');
    expect(await outputZip.file("word/header2.xml")!.async("string")).toContain("模板偶数页眉");
  });

  test("strips relationship based drawing content from copied header and footer blueprints", async () => {
    const template = await makeDocx({
      bodyText: "模板正文",
      marginTop: "1200",
      bodySize: "26",
      headerFooterBlueprint: "template",
      headerWithDrawing: true
    });
    const target = await makeDocx({
      bodyText: "目标正文",
      marginTop: "1440",
      bodySize: "24",
      headerFooterBlueprint: "target"
    });
    const profile = await extractTemplateProfile(template, "template.docx");

    const result = await applyTemplateProfile(target, profile, DEFAULT_APPLY_OPTIONS);
    const outputZip = await JSZip.loadAsync(result.data);
    const headerXml = await outputZip.file("word/header1.xml")!.async("string");

    expect(headerXml).toContain("模板页眉带图片");
    expect(headerXml).not.toContain("<w:drawing");
    expect(headerXml).not.toContain("r:embed");
    expect(result.report.warnings.join("\n")).toContain("已安全剥离");
  });

  test("applies front matter paragraph formatting without copying template text", async () => {
    const template = await makeDocxWithCover({
      coverText: "TEMPLATE COVER TEXT",
      bodyText: "TEMPLATE BODY",
      coverAlignment: "center",
      coverFont: "SimHei",
      coverSize: "72",
      coverBold: true,
      bodySize: "26"
    });
    const target = await makeDocxWithCover({
      coverText: "TARGET COVER TEXT",
      bodyText: "TARGET BODY",
      coverAlignment: "left",
      coverFont: "SimSun",
      coverSize: "24",
      coverBold: false,
      bodySize: "24"
    });
    const profile = await extractTemplateProfile(template, "template.docx");

    const result = await applyTemplateProfile(target, profile, DEFAULT_APPLY_OPTIONS);
    const zip = await JSZip.loadAsync(result.data);
    const documentXml = await zip.file("word/document.xml")!.async("string");
    const coverParagraph = documentXml.match(/<w:p>[\s\S]*?TARGET COVER TEXT[\s\S]*?<\/w:p>/)?.[0] ?? "";

    expect(documentXml).toContain("TARGET COVER TEXT");
    expect(documentXml).toContain("TARGET BODY");
    expect(documentXml).not.toContain("TEMPLATE COVER TEXT");
    expect(documentXml).not.toContain("TEMPLATE BODY");
    expect(coverParagraph).toContain('w:jc w:val="center"');
    expect(coverParagraph).toContain('w:sz w:val="72"');
    expect(coverParagraph).toContain("SimHei");
    expect(coverParagraph).toContain("<w:b");
    expect(result.report.skippedItems).toContain(
      "已原样套用模板 styles.xml，跳过全篇正文直格式覆盖，避免破坏封面和题名页直接格式。"
    );
  });

  test("skips front matter layout tables when applying table formatting", async () => {
    const template = await makeDocxWithLayoutAndDataTables({
      frontTableFont: "TemplateCoverFont",
      frontTableSize: "32",
      dataTableFont: "TemplateDataFont",
      dataTableSize: "20"
    });
    const target = await makeDocxWithLayoutAndDataTables({
      frontTableFont: "TargetCoverFont",
      frontTableSize: "32",
      dataTableFont: "TargetDataFont",
      dataTableSize: "24"
    });
    const profile = await extractTemplateProfile(template, "template.docx");

    const result = await applyTemplateProfile(target, profile, DEFAULT_APPLY_OPTIONS);
    const zip = await JSZip.loadAsync(result.data);
    const documentXml = await zip.file("word/document.xml")!.async("string");
    const frontTableParagraph = documentXml.match(/<w:p>[\s\S]*?FRONT TABLE[\s\S]*?<\/w:p>/)?.[0] ?? "";
    const dataTableParagraph = documentXml.match(/<w:p>[\s\S]*?DATA TABLE[\s\S]*?<\/w:p>/)?.[0] ?? "";

    expect(frontTableParagraph).toContain("TargetCoverFont");
    expect(frontTableParagraph).not.toContain("TemplateDataFont");
    expect(frontTableParagraph).not.toContain("<w:b");
    expect(dataTableParagraph).toContain("TemplateDataFont");
    expect(dataTableParagraph).toContain('w:sz w:val="20"');
    expect(dataTableParagraph).toContain("<w:b");
    expect(result.report.skippedItems).toContain("跳过前置页版式表格（1 个），避免破坏封面信息栏。");
  });

  test("applies structured manual header, footer, page and table controls", async () => {
    const target = await makeDocx({
      bodyText: "TARGET BODY",
      marginTop: "1440",
      bodySize: "24",
      headerFooterBlueprint: "target"
    });
    const profile = defaultTemplateProfile("manual");
    profile.page.orientation = "landscape";
    profile.headersFooters = {
      ...profile.headersFooters,
      differentFirstPage: true,
      oddEvenPages: true,
      pageNumberFormat: "upperRoman",
      hasTopBorder: true,
      hasBottomBorder: false,
      font: { ...profile.headersFooters.font, sizeHalfPoints: 22 }
    };
    profile.tables = {
      ...profile.tables,
      borderColor: "FF0000",
      cellMarginTwips: 240,
      repeatHeaderRow: true
    };

    const result = await applyTemplateProfile(target, profile, DEFAULT_APPLY_OPTIONS);
    const zip = await JSZip.loadAsync(result.data);
    const documentXml = await zip.file("word/document.xml")!.async("string");
    const settingsXml = await zip.file("word/settings.xml")!.async("string");
    const headerXml = await zip.file("word/header1.xml")!.async("string");

    expect(documentXml).toContain('w:orient="landscape"');
    expect(documentXml).toContain("<w:titlePg");
    expect(documentXml).toContain('w:fmt="upperRoman"');
    expect(documentXml).toContain("<w:tblBorders");
    expect(documentXml).toContain('w:color="FF0000"');
    expect(documentXml).toContain('w:w="240"');
    expect(documentXml).toContain("<w:tblHeader");
    expect(settingsXml).toContain("<w:evenAndOddHeaders");
    expect(headerXml).toContain('w:sz w:val="22"');
    expect(headerXml).toContain("<w:pBdr");
    expect(headerXml).toContain("<w:top");
    expect(headerXml).not.toContain("<w:bottom");
  });

  test("creates a manual STYLEREF chapter header when target has no header", async () => {
    const target = await makeDocx({ bodyText: "TARGET BODY", marginTop: "1440", bodySize: "24" });
    const profile = defaultTemplateProfile("manual-style-ref");
    profile.headersFooters = {
      ...profile.headersFooters,
      font: {
        ...profile.headersFooters.font,
        eastAsiaFont: "楷体_GB2312",
        asciiFont: "楷体_GB2312",
        sizeHalfPoints: 21
      },
      paragraph: { ...profile.headersFooters.paragraph, alignment: "center" },
      headerContent: { mode: "styleRef", text: "", styleRef: "1" },
      hasTopBorder: false,
      hasBottomBorder: true
    };

    const result = await applyTemplateProfile(target, profile, DEFAULT_APPLY_OPTIONS);
    const zip = await JSZip.loadAsync(result.data);
    const documentXml = await zip.file("word/document.xml")!.async("string");
    const relsXml = await zip.file("word/_rels/document.xml.rels")!.async("string");
    const contentTypesXml = await zip.file("[Content_Types].xml")!.async("string");
    const headerXml = await zip.file("word/header1.xml")!.async("string");

    expect(documentXml).toContain("<w:headerReference");
    expect(documentXml).toContain('w:type="default"');
    expect(relsXml).toContain('Target="header1.xml"');
    expect(contentTypesXml).toContain('/word/header1.xml');
    expect(headerXml).toContain("STYLEREF 1");
    expect(headerXml).toContain("楷体_GB2312");
    expect(headerXml).toContain('w:sz w:val="21"');
    expect(headerXml).toContain("<w:bottom");
    expect(result.report.updateFieldsRequired.join("\n")).toContain("STYLEREF");
  });

  test("extracts custom table cell margin, border color, header bold, repeat header row and caption labels", async () => {
    const template = await makeCustomTableAndCaptionDocx();
    const profile = await extractTemplateProfile(template, "custom_template.docx");

    expect(profile.tables.cellMarginTwips).toBe(240);
    expect(profile.tables.borderColor).toBe("FF0000");
    expect(profile.tables.headerBold).toBe(true);
    expect(profile.tables.repeatHeaderRow).toBe(true);
    expect(profile.tables.alignment).toBe("left");
    expect(profile.captions.labels).toContain("图表");
    expect(profile.captions.labels).toContain("附图");
  });

  test("removes an existing repeat header row marker when table profile disables it", async () => {
    const target = await makeDocxWithExistingRepeatHeader();
    const profile = defaultTemplateProfile("manual-table");
    profile.tables = { ...profile.tables, repeatHeaderRow: false };

    const result = await applyTemplateProfile(target, profile, {
      ...DEFAULT_APPLY_OPTIONS,
      page: false,
      body: false,
      headings: false,
      headersFooters: false,
      captions: false,
      numbering: false
    });
    const zip = await JSZip.loadAsync(result.data);
    const documentXml = await zip.file("word/document.xml")!.async("string");

    expect(documentXml).not.toContain("<w:tblHeader");
  });

  test("applies caption formatting only to label plus number paragraphs", async () => {
    const target = await makeCaptionPrecisionDocx();
    const profile = defaultTemplateProfile("manual-caption");
    profile.captions = {
      ...profile.captions,
      labels: ["图"],
      font: { ...profile.captions.font, color: "FF0000", sizeHalfPoints: 22 }
    };

    const result = await applyTemplateProfile(target, profile, {
      ...DEFAULT_APPLY_OPTIONS,
      page: false,
      body: false,
      headings: false,
      headersFooters: false,
      tables: false,
      numbering: false
    });
    const zip = await JSZip.loadAsync(result.data);
    const documentXml = await zip.file("word/document.xml")!.async("string");
    const normalParagraph = documentXml.match(/<w:p>[\s\S]*?图书馆介绍[\s\S]*?<\/w:p>/)?.[0] ?? "";
    const captionParagraph = documentXml.match(/<w:p>[\s\S]*?图 2 系统结构[\s\S]*?<\/w:p>/)?.[0] ?? "";

    expect(normalParagraph).not.toContain('w:color w:val="FF0000"');
    expect(captionParagraph).toContain('w:color w:val="FF0000"');
  });
});

async function makeDocx(input: {
  bodyText: string;
  marginTop: string;
  bodySize: string;
  withComments?: boolean;
  headingStyleId?: string;
  settingsMarker?: string;
  headerFooterBlueprint?: "target" | "template";
  headerWithDrawing?: boolean;
}): Promise<ArrayBuffer> {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", contentTypes(input.withComments, input.headerFooterBlueprint));
  zip.folder("_rels")!.file(".rels", rootRels());
  zip.folder("word")!.file("document.xml", documentXml(input.bodyText, input.marginTop, input.headingStyleId ?? "Heading1", input.headerFooterBlueprint));
  zip.folder("word")!.file("styles.xml", stylesXml(input.bodySize, input.headingStyleId ?? "Heading1"));
  zip.folder("word")!.file("settings.xml", settingsXml(input.settingsMarker ?? "default-settings"));
  zip.folder("word")!.file("numbering.xml", numberingXml());
  zip.folder("word")!.folder("_rels")!.file("document.xml.rels", documentRels(input.headerFooterBlueprint));
  if (input.withComments) zip.folder("word")!.file("comments.xml", "<w:comments xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"/>");
  if (input.headerFooterBlueprint === "target") {
    zip.folder("word")!.file("header1.xml", headerFooterXml("目标旧页眉"));
  }
  if (input.headerFooterBlueprint === "template") {
    zip.folder("word")!.file("header1.xml", input.headerWithDrawing ? headerFooterDrawingXml("模板页眉带图片") : headerFooterXml("模板默认页眉"));
    zip.folder("word")!.file("header2.xml", headerFooterXml("模板偶数页眉"));
    zip.folder("word")!.file("footer1.xml", headerFooterXml("模板页脚"));
    if (input.headerWithDrawing) {
      zip.folder("word")!.folder("_rels")!.file(
        "header1.xml.rels",
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImg1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/></Relationships>'
      );
      zip.folder("word")!.folder("media")!.file("image1.png", "fake image bytes");
    }
  }
  return zip.generateAsync({ type: "arraybuffer" });
}

async function makeDocxWithStyleRefHeader(): Promise<ArrayBuffer> {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", contentTypes(undefined, "target"));
  zip.folder("_rels")!.file(".rels", rootRels());
  zip.folder("word")!.file("document.xml", documentXml("正文", "1440", "Heading1", "target"));
  zip.folder("word")!.file("styles.xml", stylesXml("24", "Heading1"));
  zip.folder("word")!.file("settings.xml", settingsXml("style-ref-header-test"));
  zip.folder("word")!.file("numbering.xml", numberingXml());
  zip.folder("word")!.folder("_rels")!.file("document.xml.rels", documentRels("target"));
  zip.folder("word")!.file("header1.xml", styleRefHeaderXml());
  return zip.generateAsync({ type: "arraybuffer" });
}

function contentTypes(withComments?: boolean, headerFooterBlueprint?: "target" | "template"): string {
  const headerFooterOverrides =
    headerFooterBlueprint === "target"
      ? '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
      : headerFooterBlueprint === "template"
        ? '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/header2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
        : "";
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
  ${headerFooterOverrides}
  ${withComments ? '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>' : ""}
</Types>`;
}

function rootRels(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;
}

function documentRels(headerFooterBlueprint?: "target" | "template"): string {
  const rels =
    headerFooterBlueprint === "target"
      ? '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>'
      : headerFooterBlueprint === "template"
        ? '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header2.xml"/><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>'
        : "";
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`;
}

function documentXml(
  text: string,
  marginTop: string,
  headingStyleId: string,
  headerFooterBlueprint?: "target" | "template"
): string {
  const sectRefs =
    headerFooterBlueprint === "target"
      ? '<w:headerReference w:type="default" r:id="rId3"/>'
      : headerFooterBlueprint === "template"
        ? '<w:headerReference w:type="default" r:id="rId3"/><w:headerReference w:type="even" r:id="rId4"/><w:footerReference w:type="default" r:id="rId5"/>'
        : "";
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:pPr><w:pStyle w:val="Normal"/></w:pPr>
      <w:r><w:t>${text}</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:pStyle w:val="${headingStyleId}"/></w:pPr>
      <w:r><w:t>第一章</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>图 1 测试图</w:t></w:r>
    </w:p>
    <w:tbl><w:tr><w:tc><w:p><w:r><w:t>表格</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
    <w:sectPr>
      ${sectRefs}
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="${marginTop}" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`;
}

function stylesXml(bodySize: string, headingStyleId: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:pPr><w:jc w:val="both"/><w:spacing w:line="360" w:before="0" w:after="0"/><w:ind w:firstLine="480"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="宋体"/><w:sz w:val="${bodySize}"/><w:color w:val="000000"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="${headingStyleId}">
    <w:name w:val="heading 1"/>
    <w:pPr><w:jc w:val="center"/><w:spacing w:line="360" w:before="240" w:after="120"/><w:ind w:firstLine="0"/></w:pPr>
    <w:rPr><w:b/><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="宋体"/><w:sz w:val="30"/><w:color w:val="000000"/></w:rPr>
  </w:style>
</w:styles>`;
}

function numberingXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="0"/><w:num w:numId="1"/></w:numbering>`;
}

function settingsXml(marker: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docVar w:name="${marker}" w:val="1"/></w:settings>`;
}

function headerFooterXml(text: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:hdr>`;
}

function styleRefHeaderXml(): string {
  const rPr =
    '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="楷体_GB2312"/><w:sz w:val="21"/><w:szCs w:val="21"/><w:color w:val="000000"/></w:rPr>';
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:pPr>
      <w:jc w:val="center"/>
      <w:pBdr><w:bottom w:val="single" w:color="auto" w:sz="4" w:space="1"/></w:pBdr>
    </w:pPr>
    <w:r>${rPr}<w:fldChar w:fldCharType="begin"/></w:r>
    <w:r>${rPr}<w:instrText xml:space="preserve"> STYLEREF 1 \\* MERGEFORMAT </w:instrText></w:r>
    <w:r>${rPr}<w:fldChar w:fldCharType="separate"/></w:r>
    <w:r>${rPr}<w:t>当前章节标题</w:t></w:r>
    <w:r>${rPr}<w:fldChar w:fldCharType="end"/></w:r>
  </w:p>
</w:hdr>`;
}

function headerFooterDrawingXml(text: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
       xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <w:p>
    <w:r><w:t>${text}</w:t></w:r>
    <w:r>
      <w:drawing>
        <wp:inline>
          <a:graphic>
            <a:graphicData>
              <a:blip r:embed="rIdImg1"/>
            </a:graphicData>
          </a:graphic>
        </wp:inline>
      </w:drawing>
    </w:r>
  </w:p>
</w:hdr>`;
}

async function makeDocxWithCover(input: {
  coverText: string;
  bodyText: string;
  coverAlignment: "left" | "center" | "right";
  coverFont: string;
  coverSize: string;
  coverBold: boolean;
  bodySize: string;
}): Promise<ArrayBuffer> {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", contentTypes());
  zip.folder("_rels")!.file(".rels", rootRels());
  zip.folder("word")!.file("document.xml", documentXmlWithCover(input));
  zip.folder("word")!.file("styles.xml", stylesXml(input.bodySize, "Heading1"));
  zip.folder("word")!.file("settings.xml", settingsXml("cover-test"));
  zip.folder("word")!.folder("_rels")!.file("document.xml.rels", documentRels());
  return zip.generateAsync({ type: "arraybuffer" });
}

function documentXmlWithCover(input: {
  coverText: string;
  bodyText: string;
  coverAlignment: "left" | "center" | "right";
  coverFont: string;
  coverSize: string;
  coverBold: boolean;
}): string {
  const coverRPr = coverRunPropertiesXml(input.coverFont, input.coverSize, input.coverBold);
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:pPr>
        <w:jc w:val="${input.coverAlignment}"/>
        <w:rPr>${coverRPr}</w:rPr>
        <w:sectPr>
          <w:pgSz w:w="11906" w:h="16838"/>
          <w:pgMar w:top="1200" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
        </w:sectPr>
      </w:pPr>
      <w:r><w:rPr>${coverRPr}</w:rPr><w:t>${input.coverText}</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:pStyle w:val="Normal"/></w:pPr>
      <w:r><w:t>${input.bodyText}</w:t></w:r>
    </w:p>
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`;
}

function coverRunPropertiesXml(font: string, size: string, bold: boolean): string {
  return `<w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:eastAsia="${font}"/><w:sz w:val="${size}"/><w:szCs w:val="${size}"/><w:color w:val="000000"/>${bold ? "<w:b/>" : ""}`;
}

async function makeDocxWithLayoutAndDataTables(input: {
  frontTableFont: string;
  frontTableSize: string;
  dataTableFont: string;
  dataTableSize: string;
}): Promise<ArrayBuffer> {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", contentTypes());
  zip.folder("_rels")!.file(".rels", rootRels());
  zip.folder("word")!.file("document.xml", documentXmlWithLayoutAndDataTables(input));
  zip.folder("word")!.file("styles.xml", stylesXml("24", "Heading1"));
  zip.folder("word")!.file("settings.xml", settingsXml("layout-table-test"));
  zip.folder("word")!.folder("_rels")!.file("document.xml.rels", documentRels());
  return zip.generateAsync({ type: "arraybuffer" });
}

function documentXmlWithLayoutAndDataTables(input: {
  frontTableFont: string;
  frontTableSize: string;
  dataTableFont: string;
  dataTableSize: string;
}): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p><w:r><w:t>Front matter</w:t></w:r></w:p>
    ${tableXml("FRONT TABLE", input.frontTableFont, input.frontTableSize)}
    <w:p>
      <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
      <w:r><w:t>Chapter 1</w:t></w:r>
    </w:p>
    ${tableXml("DATA TABLE", input.dataTableFont, input.dataTableSize, true)}
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`;
}

function tableXml(text: string, font: string, size: string, bold = false): string {
  const rPr = coverRunPropertiesXml(font, size, bold);
  return `<w:tbl><w:tr><w:tc><w:p><w:r><w:rPr>${rPr}</w:rPr><w:t>${text}</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`;
}

async function makeDocxWithBodyHeading(input: {
  headingText: string;
  headingAlignment: "both" | "center";
  headingBefore: string;
  headingAfter: string;
  headingLine: string;
}): Promise<ArrayBuffer> {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", contentTypes());
  zip.folder("_rels")!.file(".rels", rootRels());
  zip.folder("word")!.file("document.xml", documentXmlWithBodyHeading(input));
  zip.folder("word")!.file("styles.xml", stylesXml("24", "2"));
  zip.folder("word")!.file("settings.xml", settingsXml("heading-test"));
  zip.folder("word")!.folder("_rels")!.file("document.xml.rels", documentRels());
  return zip.generateAsync({ type: "arraybuffer" });
}

function documentXmlWithBodyHeading(input: {
  headingText: string;
  headingAlignment: "both" | "center";
  headingBefore: string;
  headingAfter: string;
  headingLine: string;
}): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:pPr>
        <w:sectPr>
          <w:pgSz w:w="11906" w:h="16838"/>
          <w:pgMar w:top="1200" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
        </w:sectPr>
      </w:pPr>
    </w:p>
    <w:p>
      <w:pPr><w:pStyle w:val="2"/><w:pageBreakBefore/><w:spacing w:before="${input.headingBefore}" w:after="${input.headingAfter}" w:line="${input.headingLine}" w:lineRule="auto"/><w:jc w:val="${input.headingAlignment}"/></w:pPr>
      <w:r><w:rPr><w:b/><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="黑体"/><w:sz w:val="32"/></w:rPr><w:t>${input.headingText}</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:pStyle w:val="Normal"/></w:pPr>
      <w:r><w:t>Body paragraph</w:t></w:r>
    </w:p>
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`;
}

async function makeCustomTableAndCaptionDocx(): Promise<ArrayBuffer> {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", contentTypes());
  zip.folder("_rels")!.file(".rels", rootRels());
  
  const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
      <w:r><w:t>第一章 正文开始</w:t></w:r>
    </w:p>
    <w:tbl>
      <w:tblPr>
        <w:jc w:val="left"/>
        <w:tblCellMar>
          <w:top w:w="240" w:type="dxa"/>
        </w:tblCellMar>
        <w:tblBorders>
          <w:top w:val="single" w:sz="4" w:space="0" w:color="FF0000"/>
        </w:tblBorders>
      </w:tblPr>
      <w:tr>
        <w:trPr>
          <w:tblHeader/>
        </w:trPr>
        <w:tc>
          <w:p>
            <w:r>
              <w:rPr>
                <w:b/>
              </w:rPr>
              <w:t>Header</w:t>
            </w:r>
          </w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
    <w:p>
      <w:r><w:t>图表 1-1 系统测试图</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>附图 2 架构图</w:t></w:r>
    </w:p>
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>
  </w:body>
</w:document>`;

  zip.folder("word")!.file("document.xml", docXml);
  zip.folder("word")!.file("styles.xml", stylesXml("24", "Heading1"));
  return zip.generateAsync({ type: "arraybuffer" });
}

async function makeDocxWithExistingRepeatHeader(): Promise<ArrayBuffer> {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", contentTypes());
  zip.folder("_rels")!.file(".rels", rootRels());
  zip.folder("word")!.file(
    "document.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tr>
        <w:trPr><w:tblHeader/></w:trPr>
        <w:tc><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>Body</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>
  </w:body>
</w:document>`
  );
  zip.folder("word")!.file("styles.xml", stylesXml("24", "Heading1"));
  return zip.generateAsync({ type: "arraybuffer" });
}

async function makeCaptionPrecisionDocx(): Promise<ArrayBuffer> {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", contentTypes());
  zip.folder("_rels")!.file(".rels", rootRels());
  zip.folder("word")!.file(
    "document.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>图书馆介绍</w:t></w:r></w:p>
    <w:p><w:r><w:t>图 2 系统结构</w:t></w:r></w:p>
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>
  </w:body>
</w:document>`
  );
  zip.folder("word")!.file("styles.xml", stylesXml("24", "Heading1"));
  return zip.generateAsync({ type: "arraybuffer" });
}
