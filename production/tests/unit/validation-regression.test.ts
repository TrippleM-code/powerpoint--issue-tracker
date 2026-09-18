import { describe, expect, it } from "vitest";
import { normalizeIssueId, validateIssueDraft } from "../../src/domain/validation";

describe("Issue ID normalization regression", () => {
  it("trims and uppercases", () => {
    expect(normalizeIssueId("  mep-001  ")).toBe("MEP-001");
  });

  it("preserves internal characters", () => {
    expect(normalizeIssueId("mep-a/001")).toBe("MEP-A/001");
  });
});

describe("Issue draft validation regression", () => {
  it("accepts a complete draft", () => {
    expect(validateIssueDraft({
      id: "MEP-001",
      areaCode: "B1",
      roomSpace: "Pump Room",
      description: "Pipe conflicts with cable tray."
    })).toEqual([]);
  });

  it("rejects whitespace-only values", () => {
    expect(validateIssueDraft({
      id: "   ",
      areaCode: " ",
      roomSpace: "\t",
      description: "\n"
    })).toEqual([
      "Issue ID is required.",
      "Area Code is required.",
      "Room / Space is required.",
      "Description is required."
    ]);
  });

  it("reports only the missing field", () => {
    expect(validateIssueDraft({
      id: "MEP-002",
      areaCode: "L1",
      roomSpace: "",
      description: "Access clearance required."
    })).toEqual(["Room / Space is required."]);
  });
});
