import { describe, expect, it } from "vitest";
import {
  NIH_DOC_LANES,
  buildEvidencePaths,
  parseValidationLane,
} from "../nih-validation.mjs";

describe("NIH validation helpers", () => {
  it("maps lane aliases to the correct fixed document copy", () => {
    expect(parseValidationLane("a")).toMatchObject({
      lane: "copy-a",
      fileName: "NIH_Biocoating_CraftV6.docx",
    });
    expect(parseValidationLane("copy-b")).toMatchObject({
      lane: "copy-b",
      fileName: "NIH_Biocoating_CraftV6 copy.docx",
    });
    expect(parseValidationLane("source")).toMatchObject({
      lane: "source-of-truth",
      readOnly: true,
      fileName: "NIH_Biocoating_CraftV6 source-of-truth.docx",
    });
  });

  it("rejects unknown lanes with the supported lane list", () => {
    expect(() => parseValidationLane("side-quest")).toThrowError(
      /copy-a, copy-b, source-of-truth/,
    );
  });

  it("builds deterministic evidence paths under scripts/bridge output", () => {
    const result = buildEvidencePaths({
      rootDir: "/repo",
      lane: NIH_DOC_LANES[1],
      scenario: "track-changes-smoke",
      timestamp: "2026-04-11T01-35-00Z",
    });

    expect(result.baseDir).toBe(
      "/repo/scripts/bridge/artifacts/copy-b/track-changes-smoke/2026-04-11T01-35-00Z",
    );
    expect(result.preReadbackPath).toMatch(/pre-readback\.json$/);
    expect(result.postReadbackPath).toMatch(/post-readback\.json$/);
    expect(result.summaryPath).toMatch(/summary\.json$/);
  });
});
