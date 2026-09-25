import { afterEach, describe, expect, it, vi } from "vitest";
import { readFileSync } from "node:fs";
import { runInNewContext } from "node:vm";
import { normalizeLogoDataUrl } from "../../src/domain/logo-data";
import { loadSettings, saveSettings } from "../../src/services/settings-service";

const legacy = runInNewContext(
  readFileSync(new URL("../../public/v2/taskpane.js", import.meta.url), "utf8") +
  "\n({ normalizeLogoDataUrl, normalizePartyLogos, getPartyLogos, savePartyLogos })",
  { Office: { onReady: () => undefined } }
);
afterEach(() => vi.unstubAllGlobals());

describe.each([
  ["current", normalizeLogoDataUrl],
  ["legacy", legacy.normalizeLogoDataUrl],
] as const)("%s embedded logo policy", (_name, normalize) => {
  it.each([
    "https://example.invalid/beacon.png", "//example.invalid/beacon",
    "/beacon", "javascript:alert(1)", "data:image/svg+xml;base64,PHN2Zz4=",
    "data:text/html;base64,AAA=", "data:image/png,not-base64",
    "data:image/png;base64,AA A=", "data:image/png;base64,",
    {}, null, "data:image/png;base64," + "A".repeat(1024 * 1024),
  ])("rejects external, unsupported or oversized sources: %#", (source) => {
    expect(normalize(source)).toBeUndefined();
  });
  it.each(["png", "jpeg", "webp"])("preserves embedded %s", (type) => {
    const logo = "data:image/" + type + ";base64,AAA=";
    expect(normalize(" " + logo + " ")).toBe(logo);
  });
});

it("filters persisted current settings on both load and save", async () => {
  let stored: any = { parties: [
    { id: "bad", name: "Remote", logoDataUrl: "https://example.invalid/beacon" },
    { id: "good", name: "Local", logoDataUrl: "data:image/png;base64,AAA=" },
  ], statuses: ["Open"] };
  vi.stubGlobal("Office", { AsyncResultStatus: { Succeeded: "ok" },
    context: { document: { settings: {
      get: () => stored,
      set: (_key: string, value: unknown) => { stored = value; },
      saveAsync: (callback: any) => callback({ status: "ok" }),
    } } } });
  const loaded = loadSettings();
  expect(loaded.parties[0]!.logoDataUrl).toBeUndefined();
  expect(loaded.parties[1]!.logoDataUrl).toBe("data:image/png;base64,AAA=");
  await saveSettings(stored);
  expect(stored.parties[0].logoDataUrl).toBeUndefined();
  expect(stored.parties[1].logoDataUrl).toBe("data:image/png;base64,AAA=");
});

it("legacy logo maps filter values and never inherit sources", () => {
  const input = Object.assign(Object.create({ inherited: "https://example.invalid/beacon" }), {
    Remote: "//example.invalid/beacon", Local: "data:image/png;base64,AAA=",
  });
  const result = legacy.normalizePartyLogos(input);
  expect(Object.getPrototypeOf(result)).toBeNull();
  expect(Object.keys(result)).toEqual(["Local"]);
  expect(result.Local).toBe("data:image/png;base64,AAA=");
});
