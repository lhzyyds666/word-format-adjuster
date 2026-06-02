import { describe, expect, test } from "vitest";
import JSZip from "jszip";
import { insertCrossReference, scanReferenceTargets } from "../src/docx/references";

describe("cross references", () => {
  test("scans headings, captions and bookmarks as selectable reference targets", async () => {
    const docx = await makeReferenceDocx();

    const targets = await scanReferenceTargets(docx);

    expect(targets.map((item) => item.type)).toEqual(["heading", "caption", "bookmark"]);
    expect(targets[0].label).toBe("第一章 绪论");
    expect(targets[1].label).toBe("图 1 系统结构");
  });

  test("scans numeric heading style ids by resolving styles.xml names", async () => {
    const docx = await makeNumericHeadingStyleReferenceDocx();

    const targets = await scanReferenceTargets(docx);

    expect(targets[0]).toMatchObject({
      type: "heading",
      label: "第一章 绪论",
      paragraphIndex: 0
    });
  });

  test("inserts REF field before final section properties", async () => {
    const docx = await makeReferenceDocx();

    const result = await insertCrossReference(docx, {
      targetBookmark: "bm_existing",
      display: "text",
      withHyperlink: true
    });
    const zip = await JSZip.loadAsync(result.data);
    const documentXml = await zip.file("word/document.xml")!.async("string");

    expect(documentXml).toContain("REF bm_existing \\h");
    expect(documentXml.indexOf("REF bm_existing")).toBeLessThan(documentXml.indexOf("<w:sectPr"));
    expect(result.report.updateFieldsRequired).toContain("交叉引用字段需要更新域。");
  });

  test("adds a real bookmark when inserting a generated heading reference", async () => {
    const docx = await makeReferenceDocx();
    const [heading] = await scanReferenceTargets(docx);

    const result = await insertCrossReference(docx, {
      targetBookmark: heading.bookmark,
      display: "number",
      withHyperlink: true
    });
    const zip = await JSZip.loadAsync(result.data);
    const documentXml = await zip.file("word/document.xml")!.async("string");

    expect(documentXml).toContain(`w:name="${heading.bookmark}"`);
    expect(documentXml).toContain(`REF ${heading.bookmark} \\r \\h`);
  });

  test("replaces a REF placeholder instead of appending at the document end", async () => {
    const docx = await makeReferenceDocx({ placeholder: true });

    const result = await insertCrossReference(docx, {
      targetBookmark: "bm_existing",
      display: "page",
      withHyperlink: true
    });
    const zip = await JSZip.loadAsync(result.data);
    const documentXml = await zip.file("word/document.xml")!.async("string");

    expect(documentXml).not.toContain("[[REF:bm_existing]]");
    expect(documentXml).toContain("PAGEREF bm_existing \\h");
    expect(documentXml.indexOf("请见")).toBeLessThan(documentXml.indexOf("PAGEREF bm_existing"));
    expect(documentXml.indexOf("PAGEREF bm_existing")).toBeLessThan(documentXml.indexOf("继续"));
  });

  test("replaces a REF placeholder split across runs with different formatting properties", async () => {
    const docx = await makeSplitReferenceDocx();

    const result = await insertCrossReference(docx, {
      targetBookmark: "bm_existing",
      display: "text",
      withHyperlink: true
    });
    const zip = await JSZip.loadAsync(result.data);
    const documentXml = await zip.file("word/document.xml")!.async("string");

    expect(documentXml).not.toContain("[[");
    expect(documentXml).not.toContain("REF:");
    expect(documentXml).not.toContain("bm_existing]]");
    expect(documentXml).toContain("REF bm_existing \\h");
    expect(documentXml).toContain("<w:b/>"); // Bold formatting in first run preserved
    expect(documentXml).toContain("<w:i/>"); // Italic formatting in last run preserved
  });

  test("replaces a REF placeholder fragmented across adjacent same-style runs", async () => {
    const docx = await makeSameStyleFragmentedReferenceDocx();

    const result = await insertCrossReference(docx, {
      targetBookmark: "bm_existing",
      display: "text",
      withHyperlink: true
    });
    const zip = await JSZip.loadAsync(result.data);
    const documentXml = await zip.file("word/document.xml")!.async("string");

    expect(documentXml).not.toContain("[[");
    expect(documentXml).not.toContain("REF:");
    expect(documentXml).not.toContain("bm_existing]]");
    expect(documentXml).toContain("Before ");
    expect(documentXml).toContain("REF bm_existing \\h");
    expect(documentXml).toContain(" after");
    expect(documentXml.indexOf("Before ")).toBeLessThan(documentXml.indexOf("REF bm_existing"));
    expect(documentXml.indexOf("REF bm_existing")).toBeLessThan(documentXml.indexOf(" after"));
  });

  test("replaces a REF placeholder with surrounding text in the same run and preserves order and format", async () => {
    const docx = await makeSurroundingReferenceDocx();

    const result = await insertCrossReference(docx, {
      targetBookmark: "bm_existing",
      display: "text",
      withHyperlink: true
    });
    const zip = await JSZip.loadAsync(result.data);
    const documentXml = await zip.file("word/document.xml")!.async("string");

    expect(documentXml).not.toContain("[[REF:bm_existing]]");
    expect(documentXml).toContain("参见");
    expect(documentXml).toContain("的细节");
    expect(documentXml).toContain("REF bm_existing \\h");
    expect(documentXml.indexOf("参见")).toBeLessThan(documentXml.indexOf("REF bm_existing"));
    expect(documentXml.indexOf("REF bm_existing")).toBeLessThan(documentXml.indexOf("的细节"));
  });
});

async function makeReferenceDocx(options: { placeholder?: boolean } = {}): Promise<ArrayBuffer> {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", contentTypes());
  zip.folder("_rels")!.file(".rels", rootRels());
  zip.folder("word")!.file("document.xml", documentXml(options));
  zip.folder("word")!.file("styles.xml", stylesXml());
  return zip.generateAsync({ type: "arraybuffer" });
}

async function makeNumericHeadingStyleReferenceDocx(): Promise<ArrayBuffer> {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", contentTypes());
  zip.folder("_rels")!.file(".rels", rootRels());
  zip.folder("word")!.file(
    "document.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr><w:pStyle w:val="2"/></w:pPr>
      <w:r><w:t>第一章 绪论</w:t></w:r>
    </w:p>
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>
  </w:body>
</w:document>`
  );
  zip.folder("word")!.file(
    "styles.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="2"><w:name w:val="heading 1"/></w:style>
</w:styles>`
  );
  return zip.generateAsync({ type: "arraybuffer" });
}

function contentTypes(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`;
}

function rootRels(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;
}

function documentXml(options: { placeholder?: boolean } = {}): string {
  const placeholder = options.placeholder
    ? "<w:p><w:r><w:t>请见 [[REF:bm_existing]] 继续</w:t></w:r></w:p>"
    : "";
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
      <w:r><w:t>第一章 绪论</w:t></w:r>
    </w:p>
    <w:p><w:r><w:t>图 1 系统结构</w:t></w:r></w:p>
    <w:p>
      <w:bookmarkStart w:id="1" w:name="bm_existing"/>
      <w:r><w:t>已有书签</w:t></w:r>
      <w:bookmarkEnd w:id="1"/>
    </w:p>
    ${placeholder}
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>
  </w:body>
</w:document>`;
}

function stylesXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/></w:style>
</w:styles>`;
}

async function makeSplitReferenceDocx(): Promise<ArrayBuffer> {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", contentTypes());
  zip.folder("_rels")!.file(".rels", rootRels());
  
  const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:bookmarkStart w:id="1" w:name="bm_existing"/>
      <w:r><w:t>已有书签</w:t></w:r>
      <w:bookmarkEnd w:id="1"/>
    </w:p>
    <w:p>
      <w:r>
        <w:rPr>
          <w:b/>
        </w:rPr>
        <w:t>请见 [[</w:t>
      </w:r>
      <w:r>
        <w:rPr>
          <w:b/>
          <w:i/>
        </w:rPr>
        <w:t>REF:bm_existing</w:t>
      </w:r>
      <w:r>
        <w:rPr>
          <w:i/>
        </w:rPr>
        <w:t>]] 继续</w:t>
      </w:r>
    </w:p>
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>
  </w:body>
</w:document>`;
  zip.folder("word")!.file("document.xml", docXml);
  zip.folder("word")!.file("styles.xml", stylesXml());
  return zip.generateAsync({ type: "arraybuffer" });
}

async function makeSameStyleFragmentedReferenceDocx(): Promise<ArrayBuffer> {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", contentTypes());
  zip.folder("_rels")!.file(".rels", rootRels());

  const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:bookmarkStart w:id="1" w:name="bm_existing"/>
      <w:r><w:t>Existing bookmark</w:t></w:r>
      <w:bookmarkEnd w:id="1"/>
    </w:p>
    <w:p>
      <w:r><w:rPr><w:b/></w:rPr><w:t>Before [[</w:t></w:r>
      <w:r><w:rPr><w:b/></w:rPr><w:t>REF:</w:t></w:r>
      <w:r><w:rPr><w:b/></w:rPr><w:t>bm_existing]] after</w:t></w:r>
    </w:p>
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>
  </w:body>
</w:document>`;
  zip.folder("word")!.file("document.xml", docXml);
  zip.folder("word")!.file("styles.xml", stylesXml());
  return zip.generateAsync({ type: "arraybuffer" });
}

async function makeSurroundingReferenceDocx(): Promise<ArrayBuffer> {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", contentTypes());
  zip.folder("_rels")!.file(".rels", rootRels());
  
  const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:bookmarkStart w:id="1" w:name="bm_existing"/>
      <w:r><w:t>已有书签</w:t></w:r>
      <w:bookmarkEnd w:id="1"/>
    </w:p>
    <w:p>
      <w:r>
        <w:rPr>
          <w:b/>
        </w:rPr>
        <w:t>参见[[REF:bm_existing]]的细节</w:t>
      </w:r>
    </w:p>
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>
  </w:body>
</w:document>`;
  zip.folder("word")!.file("document.xml", docXml);
  zip.folder("word")!.file("styles.xml", stylesXml());
  return zip.generateAsync({ type: "arraybuffer" });
}
