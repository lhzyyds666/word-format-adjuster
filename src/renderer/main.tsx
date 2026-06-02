import React from "react";
import ReactDOM from "react-dom/client";
import {
  CheckCircle2,
  FileText,
  Layers,
  Link2,
  RotateCcw,
  Save,
  ScanLine,
  Settings2,
  SlidersHorizontal,
  Upload,
  Wand2
} from "lucide-react";
import {
  DEFAULT_APPLY_OPTIONS,
  ApplyTemplateOptions,
  FilePayload,
  FormatDiffReport,
  HeaderFooterProfile,
  HeadingStyle,
  PageSettings,
  ParagraphStyle,
  ReferenceInsertOptions,
  ReferenceTarget,
  TableProfile,
  TemplateProfile,
  TextStyle,
  TransformReport
} from "../shared/types";
import { defaultTemplateProfile } from "../docx/profile";
import { compareTemplateProfiles } from "../docx/profileDiff";
import "./styles.css";

const DEFAULT_PROFILE = defaultTemplateProfile("中文论文通用");
const TWIPS_PER_MM = 56.7;

type ManualSection = "page" | "body" | "headings" | "headersFooters" | "tables";

const manualSections: Array<{ id: ManualSection; label: string; hint: string }> = [
  { id: "page", label: "页面", hint: "纸张、方向、页边距" },
  { id: "body", label: "正文", hint: "字体、段落、行距" },
  { id: "headings", label: "标题", hint: "1-4 级标题样式" },
  { id: "headersFooters", label: "页眉页脚", hint: "页码、边线、首页" },
  { id: "tables", label: "表格题注", hint: "表格、题注格式" }
];

function App() {
  const [target, setTarget] = React.useState<FilePayload | null>(null);
  const [template, setTemplate] = React.useState<FilePayload | null>(null);
  const [profile, setProfile] = React.useState<TemplateProfile>(DEFAULT_PROFILE);
  const [options, setOptions] = React.useState<ApplyTemplateOptions>(DEFAULT_APPLY_OPTIONS);
  const [report, setReport] = React.useState<TransformReport | null>(null);
  const [output, setOutput] = React.useState<ArrayBuffer | null>(null);
  const [references, setReferences] = React.useState<ReferenceTarget[]>([]);
  const [selectedReference, setSelectedReference] = React.useState<string>("");
  const [referenceDisplay, setReferenceDisplay] = React.useState<ReferenceInsertOptions["display"]>("text");
  const [referenceHyperlink, setReferenceHyperlink] = React.useState(true);
  const [diffReport, setDiffReport] = React.useState<FormatDiffReport | null>(null);
  const [status, setStatus] = React.useState("准备就绪");
  const [busy, setBusy] = React.useState(false);

  const run = async (label: string, task: () => Promise<void>) => {
    if (busy) return;
    setBusy(true);
    setStatus(label);
    try {
      await task();
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
    } finally {
      setBusy(false);
    }
  };

  const openTarget = () =>
    run("正在打开目标文档...", async () => {
      const file = await window.wordAdjuster.openDocx();
      if (!file) return;
      setTarget(file);
      setOutput(null);
      setReport(null);
      setDiffReport(null);
      setReferences([]);
      setStatus(`已载入目标文档：${file.name}`);
    });

  const openTemplate = () =>
    run("正在提取模板格式...", async () => {
      const file = await window.wordAdjuster.openTemplate();
      if (!file) return;
      const extracted = await window.wordAdjuster.extractTemplate(file);
      setTemplate(file);
      setProfile(extracted);
      setOutput(null);
      setReport(null);
      setDiffReport(null);
      setStatus(`已提取模板：${file.name}`);
    });

  const openPreset = () =>
    run("正在导入模板预设...", async () => {
      const loaded = await window.wordAdjuster.openPreset();
      if (!loaded) return;
      setTemplate(null);
      setProfile(loaded);
      setOutput(null);
      setReport(null);
      setDiffReport(null);
      setStatus(`已导入预设：${loaded.sourceName}`);
    });

  const applyTemplate = () =>
    run("正在套用当前格式...", async () => {
      if (!target) throw new Error("请先选择目标文档。");
      const result = await window.wordAdjuster.applyTemplate(target, profile, options);
      setOutput(result.data);
      setReport(result.report);
      setDiffReport(null);
      setStatus("格式已套用，可以导出新文档。");
    });

  const saveOutput = () =>
    run("正在保存文档...", async () => {
      if (!target || !output) throw new Error("还没有可保存的输出文档。");
      const saved = await window.wordAdjuster.saveDocx(target.name, output);
      if (saved) setStatus(`已保存：${saved}`);
    });

  const savePreset = () =>
    run("正在保存预设...", async () => {
      const saved = await window.wordAdjuster.savePreset(profile.sourceName, profile);
      if (saved) setStatus(`预设已保存：${saved}`);
    });

  const scanReferences = () =>
    run("正在扫描可引用目标...", async () => {
      if (!target) throw new Error("请先选择目标文档。");
      const source = output ? { ...target, data: output } : target;
      const found = await window.wordAdjuster.scanReferences(source);
      setReferences(found);
      setSelectedReference(found[0]?.bookmark ?? "");
      setStatus(`找到 ${found.length} 个可引用目标。`);
    });

  const checkFormat = () =>
    run("正在体检格式差异...", async () => {
      if (!target) throw new Error("请先选择目标文档。");
      const source = output ? { ...target, data: output } : target;
      const targetProfile = await window.wordAdjuster.extractTemplate(source);
      const nextReport = compareTemplateProfiles(targetProfile, profile);
      setDiffReport(nextReport);
      setStatus(`体检完成：发现 ${nextReport.items.length} 处格式差异。`);
    });

  const insertReference = () =>
    run("正在插入交叉引用...", async () => {
      if (!target) throw new Error("请先选择目标文档。");
      if (!selectedReference) throw new Error("请先选择引用目标。");
      const sourceData = output ?? target.data;
      const result = await window.wordAdjuster.insertCrossReference(
        { ...target, data: sourceData },
        { targetBookmark: selectedReference, display: referenceDisplay, withHyperlink: referenceHyperlink }
      );
      setOutput(result.data);
      setReport(result.report);
      const changedItems = result.report.changedItems.join("\n");
      setStatus(changedItems.includes("[[REF:") ? "已替换交叉引用占位符。" : "已在文档末尾插入交叉引用字段。");
    });

  const updateApplyOption = (key: keyof ApplyTemplateOptions, checked: boolean) => {
    setOptions((current) => ({ ...current, [key]: checked }));
    setOutput(null);
    setReport(null);
    setDiffReport(null);
    setStatus("套用范围已更新，点击“套用当前格式”重新生成文档。");
  };

  const resetManualProfile = () => {
    setProfile(defaultTemplateProfile("中文论文通用"));
    setOptions(DEFAULT_APPLY_OPTIONS);
    setOutput(null);
    setReport(null);
    setDiffReport(null);
    setStatus("已恢复默认中文论文格式。");
  };

  const updateManualProfile = (
    applyKey: keyof ApplyTemplateOptions,
    updater: (current: TemplateProfile) => TemplateProfile
  ) => {
    setProfile((current) => {
      const next = updater(current);
      return {
        ...next,
        sourceName: next.sourceName.includes("手动调整") ? next.sourceName : `${next.sourceName}（手动调整）`,
        extractedAt: new Date().toISOString(),
        rawParts: invalidateRawPartsForManualChange(next, applyKey)
      };
    });
    setOptions((current) => ({ ...current, [applyKey]: true }));
    setOutput(null);
    setReport(null);
    setDiffReport(null);
    setStatus("手动格式已更新，点击“套用当前格式”生成新文档。");
  };

  return (
    <main className="desktop-app">
      <header className="app-titlebar">
        <div className="app-mark">W</div>
        <div>
          <h1>Word 格式调整器</h1>
          <span>本地 DOCX 桌面工具</span>
        </div>
      </header>

      <nav className="command-bar" aria-label="主操作">
        <button onClick={openTarget} className="primary" disabled={busy}>
          <FileText size={16} />
          打开目标
        </button>
        <button onClick={openTemplate} disabled={busy}>
          <Layers size={16} />
          导入模板
        </button>
        <button onClick={openPreset} disabled={busy}>
          <Upload size={16} />
          导入预设
        </button>
        <button onClick={checkFormat} disabled={busy || !target}>
          <ScanLine size={16} />
          格式体检
        </button>
        <button onClick={applyTemplate} disabled={busy || !target}>
          <Wand2 size={16} />
          套用当前格式
        </button>
        <button onClick={saveOutput} disabled={busy || !output}>
          <Save size={16} />
          导出
        </button>
        <button onClick={savePreset} disabled={busy}>
          <Settings2 size={16} />
          保存预设
        </button>
        <button onClick={resetManualProfile} disabled={busy}>
          <RotateCcw size={16} />
          恢复默认
        </button>
      </nav>

      <section className="app-grid">
        <aside className="sidebar">
          <SectionTitle title="文件" />
          <FileRow icon={<FileText size={16} />} label="目标" value={target?.name ?? "未打开"} active={Boolean(target)} />
          <FileRow icon={<Layers size={16} />} label="模板" value={template?.name ?? profile.sourceName} active={Boolean(template)} />
          <FileRow icon={<CheckCircle2 size={16} />} label="输出" value={output ? "已生成" : "未生成"} active={Boolean(output)} />

          <SectionTitle title="套用范围" />
          <OptionGrid options={options} onChange={updateApplyOption} disabled={busy} />
        </aside>

        <section className="workbench">
          <ManualFormatPanel profile={profile} updateProfile={updateManualProfile} />

          <section className="tool-panel template-panel">
            <div className="panel-title">
              <h2>
                <Settings2 size={14} />
                当前格式摘要
              </h2>
              <span>{profile.sourceName}</span>
            </div>
            <TemplateSummary profile={profile} />
          </section>

          <section className="tool-panel report-panel">
            <div className="panel-title">
              <h2>
                <CheckCircle2 size={14} />
                变更报告
              </h2>
              <span>{report ? `${report.changedItems.length} 项` : "空"}</span>
            </div>
            {report ? (
              <ul>
                {report.changedItems.map((item) => (
                  <li key={item}>{item}</li>
                ))}
                {report.skippedItems.map((item) => (
                  <li key={item} className="muted-item">
                    {item}
                  </li>
                ))}
                {report.warnings.map((item) => (
                  <li key={item} className="warn">
                    {item}
                  </li>
                ))}
                {report.updateFieldsRequired.map((item) => (
                  <li key={item} className="field">
                    {item}
                  </li>
                ))}
              </ul>
            ) : (
              <div className="empty-state">尚未执行操作</div>
            )}
          </section>
        </section>

        <aside className="task-pane">
          <section className="tool-panel diff-panel">
            <div className="panel-title">
              <h2>
                <ScanLine size={14} />
                格式体检
              </h2>
              <span>{diffReport ? `${diffReport.items.length} 项` : "未体检"}</span>
            </div>
            {diffReport ? (
              diffReport.items.length > 0 ? (
                <div className="diff-list">
                  {diffReport.items.map((item) => (
                    <div className="diff-row" key={item.id}>
                      <div>
                        <span>{item.section}</span>
                        <strong>{item.label}</strong>
                      </div>
                      <p>
                        <em>{item.current}</em>
                        <b>→</b>
                        <em>{item.expected}</em>
                      </p>
                    </div>
                  ))}
                </div>
              ) : (
                <div className="empty-state">当前文档与所选格式一致</div>
              )
            ) : (
              <div className="empty-state">打开目标后可检查与当前格式的差异</div>
            )}
          </section>

          <section className="tool-panel">
            <div className="panel-title">
              <h2>
                <Link2 size={14} />
                交叉引用
              </h2>
              <button className="icon-button" onClick={scanReferences} disabled={busy || !target} title="扫描引用目标">
                <ScanLine size={15} />
              </button>
            </div>
            <div className="reference-stack">
              <select value={selectedReference} disabled={busy} onChange={(event) => setSelectedReference(event.target.value)}>
                {references.length === 0 && <option value="">暂无目标</option>}
                {references.map((item) => (
                  <option value={item.bookmark} key={item.id}>
                    {item.type} / {item.label}
                  </option>
                ))}
              </select>
              <div className="reference-options">
                <SelectField
                  label="显示内容"
                  value={referenceDisplay}
                  onChange={(value) => setReferenceDisplay(value as ReferenceInsertOptions["display"])}
                  disabled={busy}
                  options={[
                    ["text", "引用文字"],
                    ["number", "编号"],
                    ["page", "页码"]
                  ]}
                />
                <CheckField label="带超链接" checked={referenceHyperlink} disabled={busy} onChange={setReferenceHyperlink} />
              </div>
              <button onClick={insertReference} disabled={busy || !selectedReference}>
                <Link2 size={16} />
                插入
              </button>
            </div>
          </section>
        </aside>
      </section>

      <footer className="statusbar" data-busy={busy}>
        <ScanLine size={14} />
        <span>{status}</span>
      </footer>
    </main>
  );
}

function ManualFormatPanel({
  profile,
  updateProfile
}: {
  profile: TemplateProfile;
  updateProfile: (applyKey: keyof ApplyTemplateOptions, updater: (current: TemplateProfile) => TemplateProfile) => void;
}) {
  const [activeSection, setActiveSection] = React.useState<ManualSection>("page");
  const [headingLevel, setHeadingLevel] = React.useState<HeadingStyle["level"]>(1);
  const activeHeading =
    profile.headings.find((heading) => heading.level === headingLevel) ??
    DEFAULT_PROFILE.headings[headingLevel - 1] ??
    DEFAULT_PROFILE.headings[0];

  const updatePage = (patch: (page: PageSettings) => PageSettings) =>
    updateProfile("page", (current) => ({ ...current, page: patch(current.page) }));

  const updateBodyText = (patch: (text: TextStyle) => TextStyle) =>
    updateProfile("body", (current) => ({ ...current, body: { ...current.body, text: patch(current.body.text) } }));

  const updateBodyParagraph = (patch: (paragraph: ParagraphStyle) => ParagraphStyle) =>
    updateProfile("body", (current) => ({
      ...current,
      body: { ...current.body, paragraph: patch(current.body.paragraph) }
    }));

  const updateHeading = (patch: (heading: HeadingStyle) => HeadingStyle) =>
    updateProfile("headings", (current) => {
      const patched = patch(
        current.headings.find((heading) => heading.level === headingLevel) ??
          DEFAULT_PROFILE.headings[headingLevel - 1] ??
          DEFAULT_PROFILE.headings[0]
      );
      const hasLevel = current.headings.some((heading) => heading.level === headingLevel);
      return {
        ...current,
        headings: (hasLevel
          ? current.headings.map((heading) => (heading.level === headingLevel ? patched : heading))
          : [...current.headings, patched]
        ).sort((a, b) => a.level - b.level)
      };
    });

  const updateHeaderFooter = (patch: (headersFooters: HeaderFooterProfile) => HeaderFooterProfile) =>
    updateProfile("headersFooters", (current) => ({ ...current, headersFooters: patch(current.headersFooters) }));

  const updateTable = (patch: (tables: TableProfile) => TableProfile) =>
    updateProfile("tables", (current) => ({ ...current, tables: patch(current.tables) }));

  const updateCaption = (patch: (captions: TemplateProfile["captions"]) => TemplateProfile["captions"]) =>
    updateProfile("captions", (current) => ({ ...current, captions: patch(current.captions) }));

  return (
    <section className="tool-panel manual-panel">
      <div className="panel-title">
        <h2>
          <SlidersHorizontal size={14} />
          手动调整格式
        </h2>
        <span>按分类调整，套用范围会自动勾选对应项</span>
      </div>

      <div className="manual-layout">
        <div className="format-nav" aria-label="格式分类">
          {manualSections.map((section) => (
            <button
              type="button"
              key={section.id}
              className="format-tab"
              data-active={section.id === activeSection}
              onClick={() => setActiveSection(section.id)}
            >
              <strong>{section.label}</strong>
              <span>{section.hint}</span>
            </button>
          ))}
        </div>

        <div className="format-editor">
          {activeSection === "page" && (
            <ControlGroup title="页面设置">
              <SelectField
                label="纸张"
                value={paperPreset(profile.page)}
                onChange={(value) => updatePage((page) => applyPaperPreset(page, value))}
                options={[
                  ["a4", "A4"],
                  ["letter", "Letter"],
                  ["custom", "自定义"]
                ]}
              />
              <SelectField
                label="方向"
                value={profile.page.orientation}
                onChange={(value) => updatePage((page) => ({ ...page, orientation: value as PageSettings["orientation"] }))}
                options={[
                  ["portrait", "纵向"],
                  ["landscape", "横向"]
                ]}
              />
              <NumberField
                label="上边距 mm"
                value={twipsToMm(profile.page.margins.top)}
                min={0}
                step={0.5}
                onChange={(value) => updatePage((page) => ({ ...page, margins: { ...page.margins, top: mmToTwips(value) } }))}
              />
              <NumberField
                label="下边距 mm"
                value={twipsToMm(profile.page.margins.bottom)}
                min={0}
                step={0.5}
                onChange={(value) =>
                  updatePage((page) => ({ ...page, margins: { ...page.margins, bottom: mmToTwips(value) } }))
                }
              />
              <NumberField
                label="左边距 mm"
                value={twipsToMm(profile.page.margins.left)}
                min={0}
                step={0.5}
                onChange={(value) => updatePage((page) => ({ ...page, margins: { ...page.margins, left: mmToTwips(value) } }))}
              />
              <NumberField
                label="右边距 mm"
                value={twipsToMm(profile.page.margins.right)}
                min={0}
                step={0.5}
                onChange={(value) =>
                  updatePage((page) => ({ ...page, margins: { ...page.margins, right: mmToTwips(value) } }))
                }
              />
              <NumberField
                label="页眉距 mm"
                value={twipsToMm(profile.page.margins.header)}
                min={0}
                step={0.5}
                onChange={(value) =>
                  updatePage((page) => ({ ...page, margins: { ...page.margins, header: mmToTwips(value) } }))
                }
              />
              <NumberField
                label="页脚距 mm"
                value={twipsToMm(profile.page.margins.footer)}
                min={0}
                step={0.5}
                onChange={(value) =>
                  updatePage((page) => ({ ...page, margins: { ...page.margins, footer: mmToTwips(value) } }))
                }
              />
            </ControlGroup>
          )}

          {activeSection === "body" && (
            <ControlGroup title="正文格式">
              <TextField
                label="中文字体"
                value={profile.body.text.eastAsiaFont}
                onChange={(value) => updateBodyText((text) => ({ ...text, eastAsiaFont: value }))}
              />
              <TextField
                label="英文字体"
                value={profile.body.text.asciiFont}
                onChange={(value) => updateBodyText((text) => ({ ...text, asciiFont: value }))}
              />
              <NumberField
                label="字号 pt"
                value={halfPointsToPt(profile.body.text.sizeHalfPoints)}
                min={6}
                step={0.5}
                onChange={(value) => updateBodyText((text) => ({ ...text, sizeHalfPoints: ptToHalfPoints(value) }))}
              />
              <ColorField
                label="颜色"
                value={profile.body.text.color}
                onChange={(value) => updateBodyText((text) => ({ ...text, color: value }))}
              />
              <SelectField
                label="对齐"
                value={profile.body.paragraph.alignment}
                onChange={(value) =>
                  updateBodyParagraph((paragraph) => ({ ...paragraph, alignment: value as ParagraphStyle["alignment"] }))
                }
                options={alignmentOptions}
              />
              <NumberField
                label="首行缩进 mm"
                value={twipsToMm(profile.body.paragraph.firstLineTwips)}
                step={0.5}
                onChange={(value) => updateBodyParagraph((paragraph) => ({ ...paragraph, firstLineTwips: mmToTwips(value) }))}
              />
              <NumberField
                label="段前 pt"
                value={twipsToPt(profile.body.paragraph.before)}
                min={0}
                step={0.5}
                onChange={(value) => updateBodyParagraph((paragraph) => ({ ...paragraph, before: ptToTwips(value) }))}
              />
              <NumberField
                label="段后 pt"
                value={twipsToPt(profile.body.paragraph.after)}
                min={0}
                step={0.5}
                onChange={(value) => updateBodyParagraph((paragraph) => ({ ...paragraph, after: ptToTwips(value) }))}
              />
              <NumberField
                label="行距倍数"
                value={lineToMultiple(profile.body.paragraph.line)}
                min={1}
                step={0.05}
                onChange={(value) => updateBodyParagraph((paragraph) => ({ ...paragraph, line: multipleToLine(value) }))}
              />
              <CheckField
                label="孤行控制"
                checked={profile.body.paragraph.keepLines}
                onChange={(checked) => updateBodyParagraph((paragraph) => ({ ...paragraph, keepLines: checked }))}
              />
            </ControlGroup>
          )}

          {activeSection === "headings" && (
            <ControlGroup title="标题格式">
              <SelectField
                label="标题级别"
                value={String(headingLevel)}
                onChange={(value) => setHeadingLevel(Number(value) as HeadingStyle["level"])}
                options={[
                  ["1", "一级标题"],
                  ["2", "二级标题"],
                  ["3", "三级标题"],
                  ["4", "四级标题"]
                ]}
              />
              <TextField
                label="中文字体"
                value={activeHeading.eastAsiaFont}
                onChange={(value) => updateHeading((heading) => ({ ...heading, eastAsiaFont: value }))}
              />
              <NumberField
                label="字号 pt"
                value={halfPointsToPt(activeHeading.sizeHalfPoints)}
                min={6}
                step={0.5}
                onChange={(value) => updateHeading((heading) => ({ ...heading, sizeHalfPoints: ptToHalfPoints(value) }))}
              />
              <SelectField
                label="对齐"
                value={activeHeading.alignment}
                onChange={(value) => updateHeading((heading) => ({ ...heading, alignment: value as ParagraphStyle["alignment"] }))}
                options={alignmentOptions}
              />
              <NumberField
                label="段前 pt"
                value={twipsToPt(activeHeading.before)}
                min={0}
                step={0.5}
                onChange={(value) => updateHeading((heading) => ({ ...heading, before: ptToTwips(value) }))}
              />
              <NumberField
                label="段后 pt"
                value={twipsToPt(activeHeading.after)}
                min={0}
                step={0.5}
                onChange={(value) => updateHeading((heading) => ({ ...heading, after: ptToTwips(value) }))}
              />
              <CheckField
                label="加粗"
                checked={Boolean(activeHeading.bold)}
                onChange={(checked) => updateHeading((heading) => ({ ...heading, bold: checked }))}
              />
              <CheckField
                label="与下段同页"
                checked={activeHeading.keepNext}
                onChange={(checked) => updateHeading((heading) => ({ ...heading, keepNext: checked }))}
              />
            </ControlGroup>
          )}

          {activeSection === "headersFooters" && (
            <ControlGroup title="页眉页脚">
              <SelectField
                label="页眉内容"
                value={profile.headersFooters.headerContent.mode}
                onChange={(value) =>
                  updateHeaderFooter((headersFooters) => ({
                    ...headersFooters,
                    headerContent: {
                      ...headersFooters.headerContent,
                      mode: value as HeaderFooterProfile["headerContent"]["mode"]
                    }
                  }))
                }
                options={[
                  ["preserve", "保留原页眉"],
                  ["empty", "空页眉"],
                  ["staticText", "固定文字"],
                  ["styleRef", "当前一级标题"]
                ]}
              />
              <TextField
                label="固定页眉文字"
                value={profile.headersFooters.headerContent.text}
                onChange={(value) =>
                  updateHeaderFooter((headersFooters) => ({
                    ...headersFooters,
                    headerContent: { ...headersFooters.headerContent, text: value }
                  }))
                }
              />
              <TextField
                label="章节样式编号"
                value={profile.headersFooters.headerContent.styleRef}
                onChange={(value) =>
                  updateHeaderFooter((headersFooters) => ({
                    ...headersFooters,
                    headerContent: { ...headersFooters.headerContent, styleRef: value }
                  }))
                }
              />
              <TextField
                label="中文字体"
                value={profile.headersFooters.font.eastAsiaFont}
                onChange={(value) =>
                  updateHeaderFooter((headersFooters) => ({
                    ...headersFooters,
                    font: { ...headersFooters.font, eastAsiaFont: value }
                  }))
                }
              />
              <TextField
                label="英文字体"
                value={profile.headersFooters.font.asciiFont}
                onChange={(value) =>
                  updateHeaderFooter((headersFooters) => ({
                    ...headersFooters,
                    font: { ...headersFooters.font, asciiFont: value }
                  }))
                }
              />
              <NumberField
                label="字号 pt"
                value={halfPointsToPt(profile.headersFooters.font.sizeHalfPoints)}
                min={6}
                step={0.5}
                onChange={(value) =>
                  updateHeaderFooter((headersFooters) => ({
                    ...headersFooters,
                    font: { ...headersFooters.font, sizeHalfPoints: ptToHalfPoints(value) }
                  }))
                }
              />
              <SelectField
                label="对齐"
                value={profile.headersFooters.paragraph.alignment}
                onChange={(value) =>
                  updateHeaderFooter((headersFooters) => ({
                    ...headersFooters,
                    paragraph: { ...headersFooters.paragraph, alignment: value as ParagraphStyle["alignment"] }
                  }))
                }
                options={alignmentOptions}
              />
              <SelectField
                label="页码格式"
                value={profile.headersFooters.pageNumberFormat}
                onChange={(value) =>
                  updateHeaderFooter((headersFooters) => ({
                    ...headersFooters,
                    pageNumberFormat: value as HeaderFooterProfile["pageNumberFormat"]
                  }))
                }
                options={[
                  ["decimal", "1, 2, 3"],
                  ["lowerRoman", "i, ii, iii"],
                  ["upperRoman", "I, II, III"]
                ]}
              />
              <CheckField
                label="首页不同"
                checked={profile.headersFooters.differentFirstPage}
                onChange={(checked) => updateHeaderFooter((headersFooters) => ({ ...headersFooters, differentFirstPage: checked }))}
              />
              <CheckField
                label="奇偶页不同"
                checked={profile.headersFooters.oddEvenPages}
                onChange={(checked) => updateHeaderFooter((headersFooters) => ({ ...headersFooters, oddEvenPages: checked }))}
              />
              <CheckField
                label="页眉上边线"
                checked={profile.headersFooters.hasTopBorder}
                onChange={(checked) => updateHeaderFooter((headersFooters) => ({ ...headersFooters, hasTopBorder: checked }))}
              />
              <CheckField
                label="页脚下边线"
                checked={profile.headersFooters.hasBottomBorder}
                onChange={(checked) => updateHeaderFooter((headersFooters) => ({ ...headersFooters, hasBottomBorder: checked }))}
              />
            </ControlGroup>
          )}

          {activeSection === "tables" && (
            <ControlGroup title="表格与题注">
              <NumberField
                label="表格字号 pt"
                value={halfPointsToPt(profile.tables.font.sizeHalfPoints)}
                min={6}
                step={0.5}
                onChange={(value) =>
                  updateTable((tables) => ({ ...tables, font: { ...tables.font, sizeHalfPoints: ptToHalfPoints(value) } }))
                }
              />
              <SelectField
                label="表格对齐"
                value={profile.tables.alignment}
                onChange={(value) => updateTable((tables) => ({ ...tables, alignment: value as TableProfile["alignment"] }))}
                options={[
                  ["left", "左对齐"],
                  ["center", "居中"]
                ]}
              />
              <NumberField
                label="单元格边距 mm"
                value={twipsToMm(profile.tables.cellMarginTwips)}
                min={0}
                step={0.5}
                onChange={(value) => updateTable((tables) => ({ ...tables, cellMarginTwips: mmToTwips(value) }))}
              />
              <ColorField
                label="边框颜色"
                value={profile.tables.borderColor}
                onChange={(value) => updateTable((tables) => ({ ...tables, borderColor: value }))}
              />
              <CheckField
                label="表头加粗"
                checked={profile.tables.headerBold}
                onChange={(checked) => updateTable((tables) => ({ ...tables, headerBold: checked }))}
              />
              <CheckField
                label="重复标题行"
                checked={profile.tables.repeatHeaderRow}
                onChange={(checked) => updateTable((tables) => ({ ...tables, repeatHeaderRow: checked }))}
              />
              <NumberField
                label="题注字号 pt"
                value={halfPointsToPt(profile.captions.font.sizeHalfPoints)}
                min={6}
                step={0.5}
                onChange={(value) =>
                  updateCaption((captions) => ({
                    ...captions,
                    font: { ...captions.font, sizeHalfPoints: ptToHalfPoints(value) }
                  }))
                }
              />
              <NumberField
                label="题注段后 pt"
                value={twipsToPt(profile.captions.paragraph.after)}
                min={0}
                step={0.5}
                onChange={(value) =>
                  updateCaption((captions) => ({
                    ...captions,
                    paragraph: { ...captions.paragraph, after: ptToTwips(value) }
                  }))
                }
              />
            </ControlGroup>
          )}
        </div>
      </div>
    </section>
  );
}

function ControlGroup({ title, children }: { title: string; children: React.ReactNode }) {
  return (
    <section className="control-group">
      <div className="control-group-title">{title}</div>
      <div className="control-grid">{children}</div>
    </section>
  );
}

function NumberField({
  label,
  value,
  min,
  step,
  onChange
}: {
  label: string;
  value: number;
  min?: number;
  step?: number;
  onChange: (value: number) => void;
}) {
  return (
    <label className="control-field">
      <span>{label}</span>
      <input
        type="number"
        value={formatNumber(value)}
        min={min}
        step={step}
        onChange={(event) => {
          const next = Number(event.target.value);
          if (Number.isFinite(next)) onChange(next);
        }}
      />
    </label>
  );
}

function TextField({ label, value, onChange }: { label: string; value: string; onChange: (value: string) => void }) {
  return (
    <label className="control-field">
      <span>{label}</span>
      <input type="text" value={value} onChange={(event) => onChange(event.target.value)} />
    </label>
  );
}

function ColorField({ label, value, onChange }: { label: string; value: string; onChange: (value: string) => void }) {
  return (
    <label className="control-field color-field">
      <span>{label}</span>
      <input type="color" value={`#${normalizeColor(value)}`} onChange={(event) => onChange(cleanColor(event.target.value))} />
    </label>
  );
}

function SelectField({
  label,
  value,
  onChange,
  options,
  disabled
}: {
  label: string;
  value: string;
  onChange: (value: string) => void;
  options: Array<[string, string]>;
  disabled?: boolean;
}) {
  return (
    <label className="control-field">
      <span>{label}</span>
      <select value={value} disabled={disabled} onChange={(event) => onChange(event.target.value)}>
        {options.map(([optionValue, labelText]) => (
          <option value={optionValue} key={optionValue}>
            {labelText}
          </option>
        ))}
      </select>
    </label>
  );
}

function CheckField({
  label,
  checked,
  onChange,
  disabled
}: {
  label: string;
  checked: boolean;
  onChange: (checked: boolean) => void;
  disabled?: boolean;
}) {
  return (
    <label className="check-field">
      <input type="checkbox" checked={checked} disabled={disabled} onChange={(event) => onChange(event.target.checked)} />
      <span>{label}</span>
    </label>
  );
}

function SectionTitle({ title }: { title: string }) {
  return <div className="section-title">{title}</div>;
}

function FileRow({
  icon,
  label,
  value,
  active
}: {
  icon: React.ReactNode;
  label: string;
  value: string;
  active: boolean;
}) {
  return (
    <div className="file-row" data-active={active}>
      <div className="file-icon">{icon}</div>
      <div>
        <span>{label}</span>
        <strong title={value}>{value}</strong>
      </div>
    </div>
  );
}

function TemplateSummary({ profile }: { profile: TemplateProfile }) {
  return (
    <div className="summary">
      <Property
        label="纸张"
        value={`${profile.page.orientation === "portrait" ? "纵向" : "横向"} ${profile.page.widthTwips}x${profile.page.heightTwips}`}
      />
      <Property
        label="页边距"
        value={`${profile.page.margins.top}/${profile.page.margins.right}/${profile.page.margins.bottom}/${profile.page.margins.left}`}
      />
      <Property
        label="正文"
        value={`${profile.body.text.eastAsiaFont} / ${profile.body.text.asciiFont} / ${profile.body.text.sizeHalfPoints / 2}pt`}
      />
      <Property
        label="段落"
        value={`前${profile.body.paragraph.before} 后${profile.body.paragraph.after} 行${profile.body.paragraph.line}`}
      />
      <Property label="标题" value={`${profile.headings.length} 级，一级${headingAlignmentLabel(profile.headings[0]?.alignment)}`} />
      <Property
        label="页眉页脚"
        value={`${headerModeLabel(profile.headersFooters.headerContent.mode)} / ${
          profile.headersFooters.differentFirstPage ? "首页不同" : "首页相同"
        } / ${profile.headersFooters.pageNumberFormat}`}
      />
      <Property label="编号" value={profile.numbering.hasNumberingXml ? "已识别" : "无编号文件"} />
      {profile.warnings.map((warning) => (
        <p className="warning" key={warning}>
          {warning}
        </p>
      ))}
    </div>
  );
}

function Property({ label, value }: { label: string; value: string }) {
  return (
    <div className="property-row">
      <span>{label}</span>
      <strong>{value}</strong>
    </div>
  );
}

function OptionGrid({
  options,
  onChange,
  disabled
}: {
  options: ApplyTemplateOptions;
  onChange: (key: keyof ApplyTemplateOptions, checked: boolean) => void;
  disabled?: boolean;
}) {
  const labels: Record<keyof ApplyTemplateOptions, string> = {
    page: "页面",
    body: "正文",
    headings: "标题",
    headersFooters: "页眉页脚",
    tables: "表格",
    captions: "题注",
    numbering: "编号"
  };
  return (
    <div className="option-grid">
      {(Object.keys(labels) as Array<keyof ApplyTemplateOptions>).map((key) => (
        <label key={key}>
          <input
            type="checkbox"
            checked={options[key]}
            disabled={disabled}
            onChange={(event) => onChange(key, event.target.checked)}
          />
          {labels[key]}
        </label>
      ))}
    </div>
  );
}

const alignmentOptions: Array<[string, string]> = [
  ["left", "左对齐"],
  ["center", "居中"],
  ["right", "右对齐"],
  ["both", "两端对齐"]
];

function invalidateRawPartsForManualChange(profile: TemplateProfile, applyKey: keyof ApplyTemplateOptions) {
  if (!profile.rawParts) return undefined;
  const rawParts = { ...profile.rawParts };
  if (applyKey === "body" || applyKey === "headings") {
    delete rawParts.stylesXml;
    delete rawParts.frontMatterParagraphs;
    delete rawParts.bodyHeadingParagraphs;
  }
  if (applyKey === "headersFooters") {
    delete rawParts.headerFooterParts;
    delete rawParts.headerFooterRelationships;
    delete rawParts.headerFooterContentTypeOverrides;
    delete rawParts.sectionHeaderFooterReferences;
  }
  if (applyKey === "page") {
    delete rawParts.settingsXml;
  }
  return rawParts;
}

function paperPreset(page: PageSettings) {
  if (page.widthTwips === 11906 && page.heightTwips === 16838) return "a4";
  if (page.widthTwips === 12240 && page.heightTwips === 15840) return "letter";
  return "custom";
}

function applyPaperPreset(page: PageSettings, preset: string): PageSettings {
  if (preset === "a4") return { ...page, widthTwips: 11906, heightTwips: 16838 };
  if (preset === "letter") return { ...page, widthTwips: 12240, heightTwips: 15840 };
  return page;
}

function twipsToMm(value: number) {
  return Number((value / TWIPS_PER_MM).toFixed(1));
}

function mmToTwips(value: number) {
  return Math.max(0, Math.round(value * TWIPS_PER_MM));
}

function twipsToPt(value: number) {
  return Number((value / 20).toFixed(1));
}

function ptToTwips(value: number) {
  return Math.max(0, Math.round(value * 20));
}

function halfPointsToPt(value: number) {
  return Number((value / 2).toFixed(1));
}

function ptToHalfPoints(value: number) {
  return Math.max(1, Math.round(value * 2));
}

function lineToMultiple(value: number) {
  return Number((value / 240).toFixed(2));
}

function multipleToLine(value: number) {
  return Math.max(120, Math.round(value * 240));
}

function normalizeColor(value: string) {
  const color = cleanColor(value);
  return color.length === 6 ? color : "000000";
}

function cleanColor(value: string) {
  return value.replace("#", "").replace(/[^0-9a-f]/gi, "").slice(0, 6).toUpperCase();
}

function formatNumber(value: number) {
  if (!Number.isFinite(value)) return "";
  return Number.isInteger(value) ? String(value) : String(value);
}

function headingAlignmentLabel(value?: ParagraphStyle["alignment"]) {
  if (value === "center") return "居中";
  if (value === "right") return "右对齐";
  if (value === "both") return "两端对齐";
  return "左对齐";
}

function headerModeLabel(value: HeaderFooterProfile["headerContent"]["mode"]) {
  if (value === "empty") return "空页眉";
  if (value === "staticText") return "固定文字";
  if (value === "styleRef") return "当前一级标题";
  return "保留页眉";
}

ReactDOM.createRoot(document.getElementById("root")!).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);
