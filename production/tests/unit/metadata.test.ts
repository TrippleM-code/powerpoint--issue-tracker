import { describe, expect, it } from "vitest";
import { parseIssueMetadata } from "../../src/domain/metadata";

const issue = { id: "A", areaCode: "L1", roomSpace: "Room", description: "Issue", createdAt: "2026-09-20T00:00:00Z", actions: [] };

describe("stored issue validation", () => {
  it("reads valid v1 and legacy unversioned metadata", () => {
    expect(parseIssueMetadata(JSON.stringify(issue), "1", "Slide 2")).toEqual(issue);
    expect(parseIssueMetadata(JSON.stringify(issue), undefined, "Slide 2")).toEqual(issue);
  });
  it.each(["null", "[]", '{"id":"BAD"}', "broken", JSON.stringify({ ...issue, createdAt: "invalid" })])(
    "rejects malformed metadata with the slide location: %s", (json) => {
      expect(() => parseIssueMetadata(json, "1", "Slide 2")).toThrow(/Slide 2: invalid IssueFlow metadata/);
    },
  );
  it("does not reinterpret a future schema", () => {
    expect(() => parseIssueMetadata(JSON.stringify(issue), "2", "Slide 2")).toThrow(/Unsupported schema/);
  });
  it("rejects duplicate action IDs", () => {
    const action = { id: "a", party: "MEP", required: "Check", status: "Open", createdAt: issue.createdAt };
    expect(() => parseIssueMetadata(JSON.stringify({ ...issue, actions: [action, action] }), "1", "Slide 2")).toThrow(/Duplicate action/);
  });
});
