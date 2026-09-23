// @vitest-environment jsdom
import { beforeEach, afterEach, describe, expect, it, vi } from "vitest";
import panelHtml from "../../index.html?raw";
import type { Issue } from "../../src/domain/models";

const mock = vi.hoisted(() => ({
  readSelectedIssueState: vi.fn(), readAllIssues: vi.fn(), saveSelectedIssue: vi.fn(),
  renderSelectedIssue: vi.fn(), getSummaryStats: vi.fn(), saveSettings: vi.fn(), loadSettings: vi.fn(),
}));
vi.mock("../../src/services/powerpoint-service", () => ({ PowerPointService: class {
  readSelectedIssueState = mock.readSelectedIssueState;
  readAllIssues = mock.readAllIssues;
  saveSelectedIssue = mock.saveSelectedIssue;
  renderSelectedIssue = mock.renderSelectedIssue;
  getSummaryStats = mock.getSummaryStats;
} }));
vi.mock("../../src/services/settings-service", () => ({
  getDefaultSettings: () => ({ parties: [{ id: "m", name: "MEP" }, { id: "a", name: "Architect" }], statuses: ["Open", "Closed"] }),
  loadSettings: mock.loadSettings, saveSettings: mock.saveSettings,
}));

const base: Issue = { id: "A", areaCode: "L1", roomSpace: "Room", description: "Original", createdAt: "2026-09-20T00:00:00Z",
  actions: [{ id: "a1", party: "MEP", required: "Original action", status: "Open", createdAt: "2026-09-20T00:00:00Z" }] };
const control = (id: string) => document.getElementById(id) as HTMLInputElement;
const click = (id: string) => (document.getElementById(id) as HTMLButtonElement).click();
const button = (container: string, text: string) => {
  const found = [...document.querySelectorAll<HTMLButtonElement>(`${container} button`)].find(b => b.textContent === text);
  if (!found) throw new Error(`Button missing: ${text}`);
  return found;
};
async function start(issue = structuredClone(base)) {
  mock.readSelectedIssueState.mockResolvedValue({ slideId: "slide-A", issue });
  mock.readAllIssues.mockResolvedValue([{ slideId: "slide-A", slideNumber: 1, issue }]);
  const { initializeIssuePanel } = await import("../../src/ui/issue-panel");
  await initializeIssuePanel();
}

beforeEach(() => {
  vi.resetModules(); vi.clearAllMocks();
  document.body.innerHTML = panelHtml;
  mock.loadSettings.mockReturnValue({ parties: [{ id: "m", name: "MEP" }, { id: "a", name: "Architect" }], statuses: ["Open", "Closed"] });
  mock.getSummaryStats.mockResolvedValue({ issueCount: 1, actionCount: 1, openCount: 1, closedCount: 0 });
  mock.saveSettings.mockResolvedValue(undefined);
  mock.renderSelectedIssue.mockResolvedValue(undefined);
  vi.stubGlobal("Office", { context: { document: { addHandlerAsync: vi.fn() } }, EventType: { DocumentSelectionChanged: "selection" } });
});
afterEach(() => vi.unstubAllGlobals());

describe("panel regressions", () => {
  it("rejects Confirm Remove after a slide switch", async () => {
    await start();
    button("#actionList", "Remove").click();
    mock.readSelectedIssueState.mockResolvedValue({ slideId: "slide-B", issue: { ...base, id: "B" } });
    button("#actionList", "Confirm Remove").click();
    await vi.waitFor(() => expect(control("statusBanner").textContent).toMatch(/Selected slide changed/));
    expect(mock.saveSelectedIssue).not.toHaveBeenCalled();
    expect(mock.renderSelectedIssue).not.toHaveBeenCalled();
  });
  it("adds a new action after a blocked edit without overwriting the old one", async () => {
    mock.loadSettings.mockReturnValue({ parties: [{ id: "a", name: "Architect" }], statuses: ["Open"] });
    await start();
    button("#actionList", "Edit").click();
    expect(control("editingActionId").value).toBe("");
    expect(control("saveActionBtn").textContent).toBe("Add Action");
    control("actionRequired").value = "New action";
    click("saveActionBtn");
    await vi.waitFor(() => expect(mock.renderSelectedIssue).toHaveBeenCalledOnce());
    const [saved, slide] = mock.renderSelectedIssue.mock.calls[0]!;
    expect(slide).toBe("slide-A");
    expect(saved.actions).toHaveLength(2);
    expect(saved.actions[0].required).toBe("Original action");
  });
  it("preserves form edits when removing an action", async () => {
    await start(); control("description").value = "Updated draft";
    button("#actionList", "Remove").click(); button("#actionList", "Confirm Remove").click();
    await vi.waitFor(() => expect(mock.renderSelectedIssue).toHaveBeenCalledOnce());
    expect(mock.renderSelectedIssue.mock.calls[0]![0]).toMatchObject({ description: "Updated draft", actions: [] });
    expect(mock.saveSelectedIssue).not.toHaveBeenCalled();
  });
  it("blocks removing a party used on another issue", async () => {
    await start({ ...base, actions: [] });
    mock.readAllIssues.mockResolvedValue([{ slideId: "other", slideNumber: 2, issue: base }]);
    button("#partyLibraryList", "Remove Party").click();
    await vi.waitFor(() => expect(control("statusBanner").textContent).toMatch(/used by an existing action/));
    expect(mock.saveSettings).not.toHaveBeenCalled();
  });
  it("blocks removing a status used on another issue", async () => {
    await start({ ...base, actions: [] });
    mock.readAllIssues.mockResolvedValue([{ slideId: "other", slideNumber: 2, issue: base }]);
    button("#statusLibraryList", "Remove").click();
    await vi.waitFor(() => expect(control("statusBanner").textContent).toMatch(/used by an existing action/));
    expect(mock.saveSettings).not.toHaveBeenCalled();
  });
  it("disables dynamic remove/edit controls during a pending save", async () => {
    await start();
    let finish!: () => void;
    mock.renderSelectedIssue.mockImplementation(() => new Promise<void>(resolve => { finish = resolve; }));
    click("refreshSheetBtn");
    await vi.waitFor(() => expect(mock.renderSelectedIssue).toHaveBeenCalledOnce());
    expect(button("#actionList", "Remove").disabled).toBe(true);
    expect(button("#actionList", "Edit").disabled).toBe(true);
    expect(control("description").disabled).toBe(true);
    finish();
    await vi.waitFor(() => expect(control("description").disabled).toBe(false));
  });
  it("rejects duplicate IDs on copied existing issue slides", async () => {
    await start();
    mock.readAllIssues.mockResolvedValue([{ slideId: "slide-A", slideNumber: 1, issue: base }, { slideId: "copy", slideNumber: 2, issue: base }]);
    click("saveIssueBtn");
    await vi.waitFor(() => expect(control("statusBanner").textContent).toMatch(/already exists on Slide 2/));
    expect(mock.saveSelectedIssue).not.toHaveBeenCalled();
  });
});
