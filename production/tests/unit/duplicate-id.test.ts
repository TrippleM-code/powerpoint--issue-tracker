import { describe, expect, it } from "vitest";
import { isDuplicateIssueId } from "../../src/domain/validation";

describe("isDuplicateIssueId", () => {
  it("detects an exact duplicate", () => {
    expect(isDuplicateIssueId("MEP-001", ["MEP-001", "MEP-002"])).toBe(true);
  });

  it("detects duplicates case-insensitively and ignores surrounding spaces", () => {
    expect(isDuplicateIssueId(" mep-001 ", ["MEP-001"])).toBe(true);
  });

  it("allows a genuinely new Issue ID", () => {
    expect(isDuplicateIssueId("MEP-003", ["MEP-001", "MEP-002"])).toBe(false);
  });

  it("does not treat an empty candidate as a duplicate", () => {
    expect(isDuplicateIssueId("   ", ["MEP-001"])).toBe(false);
  });
});
