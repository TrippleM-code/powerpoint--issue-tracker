import { describe, expect, it } from "vitest";
import { computeOverallStatus, statusCssClass } from "../../src/domain/status";

describe("computeOverallStatus", () => {
  it("returns No actions for an empty list", () => {
    expect(computeOverallStatus([])).toBe("No actions");
  });

  it("returns Closed when all actions are closed", () => {
    expect(computeOverallStatus(["Closed", "Closed"])).toBe("Closed");
  });

  it("returns Open when any action is open", () => {
    expect(computeOverallStatus(["Closed", "Pending", "Open"])).toBe("Open");
  });

  it("returns In Progress when there is no open action but one is in progress", () => {
    expect(computeOverallStatus(["Closed", "In Progress", "Pending"])).toBe("In Progress");
  });

  it("uses Open as conservative fallback for unknown custom statuses", () => {
    expect(computeOverallStatus(["WIP"])).toBe("Open");
  });
});

describe("statusCssClass", () => {
  it("returns a neutral fallback for custom statuses", () => {
    expect(statusCssClass("WIP")).toBe("status-custom");
  });
});
