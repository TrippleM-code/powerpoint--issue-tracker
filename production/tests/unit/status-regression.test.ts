import { describe, expect, it } from "vitest";
import { computeOverallStatus, statusCssClass } from "../../src/domain/status";

describe("computeOverallStatus regression", () => {
  it("returns No actions for an empty list", () => {
    expect(computeOverallStatus([])).toBe("No actions");
  });

  it("returns Closed only when all actions are closed", () => {
    expect(computeOverallStatus(["Closed", " closed ", "CLOSED"])).toBe("Closed");
  });

  it("gives Open priority", () => {
    expect(computeOverallStatus(["Pending", "Closed", "In Progress", "Open"])).toBe("Open");
  });

  it("gives In Progress priority over Pending", () => {
    expect(computeOverallStatus(["Pending", "Closed", "In Progress"])).toBe("In Progress");
  });

  it("returns Pending when appropriate", () => {
    expect(computeOverallStatus(["Closed", "Pending"])).toBe("Pending");
  });

  it("uses Open for custom status fallback", () => {
    expect(computeOverallStatus(["WIP"])).toBe("Open");
  });
});

describe("statusCssClass regression", () => {
  it.each([
    ["Open", "status-open"],
    ["In Progress", "status-in-progress"],
    ["Pending", "status-pending"],
    ["Closed", "status-closed"],
    ["WIP", "status-custom"],
  ])("maps %s to %s", (status, expected) => {
    expect(statusCssClass(status)).toBe(expected);
  });
});
