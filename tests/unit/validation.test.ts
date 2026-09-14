import { describe, expect, it } from "vitest";
import { normalizeIssueId, validateIssueDraft } from "../../src/domain/validation";

describe("normalizeIssueId", () => {
  it("trims and uppercases IDs", () => {
    expect(normalizeIssueId(" mep-001 ")).toBe("MEP-001");
  });
});

describe("validateIssueDraft", () => {
  it("accepts a complete draft", () => {
    expect(validateIssueDraft({
      id: "MEP-001",
      areaCode: "BASEMENT",
      roomSpace: "TX Room Corridor",
      description: "Check route above water tank."
    })).toEqual([]);
  });

  it("rejects missing required fields", () => {
    const errors = validateIssueDraft({
      id: "",
      areaCode: "",
      roomSpace: "",
      description: ""
    });
    expect(errors).toHaveLength(4);
  });
});
