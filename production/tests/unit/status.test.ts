import { describe, expect, it } from "vitest";
import { computeOverallStatus } from "../../src/domain/status";

describe("computeOverallStatus", () => {
  it("returns No actions for an empty list", () => {
    expect(computeOverallStatus([])).toBe("No actions");
  });

  it("returns Closed when all actions are closed", () => {
    expect(computeOverallStatus(["Closed", "Closed"])).toBe("Closed");
  });

  it("keeps issue open when any action is open", () => {
    expect(computeOverallStatus(["Closed", "Open", "Pending"])).toBe("Open");
  });
});
