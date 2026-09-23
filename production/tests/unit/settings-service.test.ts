import { afterEach, describe, expect, it, vi } from "vitest";
import { getDefaultSettings, loadSettings, saveSettings } from "../../src/services/settings-service";

afterEach(() => vi.unstubAllGlobals());

describe("IssueFlow settings regression", () => {
  it("restores the in-memory settings after a failed document save", async () => {
    const previous = { parties: ["MEP"], statuses: ["Open"] };
    let stored: unknown = previous;
    vi.stubGlobal("Office", {
      AsyncResultStatus: { Succeeded: "succeeded" },
      context: { document: { settings: {
        get: () => stored,
        set: (_key: string, value: unknown) => { stored = value; },
        saveAsync: (callback: any) => callback({ status: "failed", error: { message: "Save failed" } }),
      } } },
    });
    await expect(saveSettings(getDefaultSettings())).rejects.toThrow("Save failed");
    expect(stored).toBe(previous);
  });
  it("returns stable default parties and statuses", () => {
    const settings = getDefaultSettings();
    expect(settings.parties.map((party) => party.name)).toEqual([
      "Architect", "C&S", "MEP", "Main Contractor", "ESCS"
    ]);
    expect(settings.statuses).toEqual(["Open", "In Progress", "Pending", "Closed"]);
  });

  it("falls back safely when Office document settings are unavailable", () => {
    vi.stubGlobal("Office", undefined);
    const settings = loadSettings();
    expect(settings.parties.map((party) => party.name)).toEqual([
      "Architect", "C&S", "MEP", "Main Contractor", "ESCS"
    ]);
    expect(settings.statuses).toEqual(["Open", "In Progress", "Pending", "Closed"]);
  });

  it("migrates legacy string parties to Party objects", () => {
    vi.stubGlobal("Office", {
      context: { document: { settings: { get: () => ({
        parties: ["Architect", "MEP", "Main Contractor"],
        statuses: ["Open", "Closed"]
      }) } } }
    });

    const settings = loadSettings();
    expect(settings.parties.map((party) => party.name)).toEqual([
      "Architect", "MEP", "Main Contractor"
    ]);
    expect(settings.parties.every((party) => Boolean(party.id))).toBe(true);
  });

  it("preserves an existing logo during normalization", () => {
    const logo = "data:image/png;base64,AAA";
    vi.stubGlobal("Office", {
      context: { document: { settings: { get: () => ({
        parties: [{ id: "party-mep", name: "MEP", logoDataUrl: logo }],
        statuses: ["Open"]
      }) } } }
    });

    const settings = loadSettings();
    expect(settings.parties).toEqual([
      { id: "party-mep", name: "MEP", logoDataUrl: logo }
    ]);
  });

  it("removes duplicate parties and statuses case-insensitively", () => {
    vi.stubGlobal("Office", {
      context: { document: { settings: { get: () => ({
        parties: ["MEP", "mep", "Architect"],
        statuses: ["Open", "open", "Closed", "closed"]
      }) } } }
    });

    const settings = loadSettings();
    expect(settings.parties.map((party) => party.name)).toEqual(["MEP", "Architect"]);
    expect(settings.statuses).toEqual(["Open", "Closed"]);
  });
});
