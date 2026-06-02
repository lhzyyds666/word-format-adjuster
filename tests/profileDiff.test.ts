import { describe, expect, test } from "vitest";
import { defaultTemplateProfile } from "../src/docx/profile";
import { compareTemplateProfiles } from "../src/docx/profileDiff";

describe("template profile diff", () => {
  test("reports differences between a target profile and the selected profile", () => {
    const target = defaultTemplateProfile("target");
    const expected = defaultTemplateProfile("expected");
    target.page.margins.top = 1200;
    target.body.text.sizeHalfPoints = 22;
    target.headings[0].alignment = "left";
    target.tables.borderColor = "808080";

    expected.page.margins.top = 1440;
    expected.body.text.sizeHalfPoints = 24;
    expected.headings[0].alignment = "center";
    expected.tables.borderColor = "FF0000";

    const report = compareTemplateProfiles(target, expected);

    expect(report.items).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ id: "page.margin.top", current: "21.2 mm", expected: "25.4 mm" }),
        expect.objectContaining({ id: "body.text.size", current: "11 pt", expected: "12 pt" }),
        expect.objectContaining({ id: "heading.1.alignment", current: "左对齐", expected: "居中" }),
        expect.objectContaining({ id: "tables.borderColor", current: "#808080", expected: "#FF0000" })
      ])
    );
  });

  test("returns an empty item list when profiles match", () => {
    const profile = defaultTemplateProfile("same");

    const report = compareTemplateProfiles(profile, structuredClone(profile));

    expect(report.items).toEqual([]);
  });
});
