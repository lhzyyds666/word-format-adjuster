import { describe, expect, test } from "vitest";
import { defaultTemplateProfile } from "../src/docx/profile";
import { parseTemplateProfilePreset } from "../src/shared/profilePreset";

describe("template profile presets", () => {
  test("parses a saved template profile preset", () => {
    const profile = defaultTemplateProfile("saved preset");

    const parsed = parseTemplateProfilePreset(JSON.stringify(profile));

    expect(parsed.sourceName).toBe("saved preset");
    expect(parsed.body.text.eastAsiaFont).toBe(profile.body.text.eastAsiaFont);
    expect(parsed.headersFooters.headerContent.mode).toBe("preserve");
  });

  test("rejects json that is not a template profile", () => {
    expect(() => parseTemplateProfilePreset('{"sourceName":"bad"}')).toThrow("不是有效的模板预设");
  });

  test("rejects incomplete nested preset structures", () => {
    const profile = defaultTemplateProfile("broken preset") as unknown as Record<string, unknown>;
    profile.headersFooters = {
      differentFirstPage: false,
      oddEvenPages: false
    };

    expect(() => parseTemplateProfilePreset(JSON.stringify(profile))).toThrow("不是有效的模板预设");
  });

  test("rejects invalid enum values and malformed raw parts", () => {
    const profile = defaultTemplateProfile("broken raw") as unknown as Record<string, unknown>;
    profile.page = { ...(profile.page as Record<string, unknown>), orientation: "diagonal" };
    expect(() => parseTemplateProfilePreset(JSON.stringify(profile))).toThrow("不是有效的模板预设");

    const rawProfile = defaultTemplateProfile("broken raw") as unknown as Record<string, unknown>;
    rawProfile.rawParts = {
      headerFooterRelationships: [{ id: "rId1", type: "image", target: "header1.xml" }]
    };
    expect(() => parseTemplateProfilePreset(JSON.stringify(rawProfile))).toThrow("不是有效的模板预设");
  });
});
