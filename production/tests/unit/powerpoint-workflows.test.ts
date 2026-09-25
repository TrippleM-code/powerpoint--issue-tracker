import { afterEach, describe, expect, it, vi } from "vitest";
import { PowerPointService } from "../../src/services/powerpoint-service";
import { TAGS } from "../../src/storage/tag-names";
import type { Issue } from "../../src/domain/models";

const issue: Issue = { id: "A", areaCode: "L1", roomSpace: "Room", description: "Description", createdAt: "2026-09-20T00:00:00Z", actions: [] };

function host() {
  let sequence = 0;
  const state = { selected: "A", failShapes: false, failMove: false, afterSync: () => {} };
  const tags = (initial: Record<string, string> = {}) => ({
    items: Object.entries(initial).map(([key, value]) => ({ key, value })), load() {},
    add(key: string, value: string) { this.items = this.items.filter(t => t.key !== key); this.items.push({ key, value }); },
  });
  const slides: any[] = [];
  function addSlide(id: string, initial: Record<string, string> = {}) {
    const shapes: any[] = [];
    const addShape = (text = "", options = {}) => {
      if (state.failShapes) throw new Error("Injected shape failure");
      const shape: any = { text, ...options, id: `shape-${++sequence}`, name: "", tags: tags(),
        textFrame: { textRange: { font: {} } }, fill: { setSolidColor() {}, setImage() {} }, lineFormat: {},
        delete: vi.fn(() => { const i = shapes.indexOf(shape); if (i >= 0) shapes.splice(i, 1); }),
      };
      shapes.push(shape); return shape;
    };
    const slide: any = { id, tags: tags(initial), shapes: { items: shapes, load() {}, addTextBox: addShape,
      addGeometricShape: (_type: string, options: Record<string, unknown>) => addShape("", options),
    }, delete: vi.fn(() => { const i = slides.indexOf(slide); if (i >= 0) slides.splice(i, 1); }),
      moveTo(index: number) {
        if (state.failMove) throw new Error("Injected move failure");
        slides.splice(slides.indexOf(slide), 1); slides.splice(index, 0, slide);
      },
    };
    slides.push(slide); return slide;
  }
  const context = { sync: vi.fn(async () => { state.afterSync(); }), presentation: {
    getSelectedSlides: () => ({ items: slides.filter(s => s.id === state.selected), load() {} }),
    setSelectedSlides: (ids: string[]) => { state.selected = ids[0]!; },
    slides: { items: slides, load() {}, getItem: (id: string) => {
      const slide = slides.find(s => s.id === id); if (!slide) throw new Error("Missing slide"); return slide;
    }, getItemAt: (index: number) => slides[index], getCount: () => ({ value: slides.length }),
      add: () => { addSlide(`new-${++sequence}`); },
    },
  } };
  vi.stubGlobal("Office", { context: { requirements: { isSetSupported: () => true }, document: { settings: { get: () => undefined } } } });
  vi.stubGlobal("PowerPoint", { run: async (fn: any) => fn(context) });
  return { state, slides, addSlide, context };
}
afterEach(() => vi.unstubAllGlobals());

describe("PowerPoint service regression boundaries", () => {
  it("keeps typed reference boxes through refresh and bulk settings, with no covering shapes", async () => {
    const h = host(); const a = h.addSlide("A");
    const service = new PowerPointService();
    await service.renderSelectedIssue(issue, "A");
    const fields = () => a.shapes.items.filter((shape: any) => shape.tags.items.some((tag: any) => tag.key === TAGS.referenceRole));
    const originals = [...fields()];
    expect(originals).toHaveLength(3);
    expect(originals.map((shape: any) => shape.text)).toEqual(["", "", ""]);
    originals.forEach((shape: any, index: number) => { shape.text = `Ref-${index}/2026`; });
    const updated = { ...issue, updatedAt: "2026-09-26T12:30:00Z" };
    await service.renderSelectedIssue(updated, "A");
    expect((await service.applySettingsToAllIssueSlides()).failed).toBe(0);
    expect(fields()).toEqual(originals);
    originals.forEach((field: any, index: number) => {
      expect(field.text).toBe(`Ref-${index}/2026`);
      expect(field.delete).not.toHaveBeenCalled();
      const covering = a.shapes.items.filter((shape: any) => shape !== field &&
        shape.left < field.left + field.width && shape.left + shape.width > field.left &&
        shape.top < field.top + field.height && shape.top + shape.height > field.top);
      expect(covering).toHaveLength(0);
    });
    for (const label of ["SOI", "NOV", "COI/CVI"]) {
      expect(a.shapes.items.filter((shape: any) => shape.text === label)).toHaveLength(1);
    }
    for (const prefix of ["Created ", "Updated "]) {
      const date = a.shapes.items.find((shape: any) => shape.text.startsWith(prefix));
      expect(date).toBeDefined(); expect(date.top + date.height).toBeLessThan(76);
    }
    originals[1].delete();
    await service.renderSelectedIssue(updated, "A");
    expect(fields()).toHaveLength(3);
    expect(fields().filter((shape: any) => shape.text === "")).toHaveLength(1);
    expect(fields()).toContain(originals[0]); expect(fields()).toContain(originals[2]);
  });
  it("retains typed references when a refresh fails", async () => {
    const h = host(); const a = h.addSlide("A");
    await new PowerPointService().renderSelectedIssue(issue, "A");
    const reference = a.shapes.items.find((shape: any) => shape.tags.items.some((tag: any) => tag.key === TAGS.referenceRole));
    reference.text = "SOI-keep"; h.state.failShapes = true;
    await expect(new PowerPointService().renderSelectedIssue(issue, "A")).rejects.toThrow(/Injected/);
    expect(reference.text).toBe("SOI-keep"); expect(reference.delete).not.toHaveBeenCalled();
  });
  it("rejects a change during web staging and removes only the staged shapes", async () => {
    const h = host(); const a = h.addSlide("A", { [TAGS.issueJson]: JSON.stringify(issue) });
    const old = a.shapes.addTextBox("previous layout"); old.tags.add(TAGS.managed, "TRUE");
    const newer = { ...issue, description: "Other editor's change" };
    h.state.afterSync = () => {
      if (a.shapes.items.length > 1) a.tags.add(TAGS.issueJson, JSON.stringify(newer));
    };
    await expect(new PowerPointService().renderSelectedIssue(issue, "A", issue)).rejects.toThrow(/changed since it was loaded/);
    expect(old.delete).not.toHaveBeenCalled();
    expect(a.shapes.items).toEqual([old]);
    expect(a.tags.items.find((tag: any) => tag.key === TAGS.issueJson).value).toBe(JSON.stringify(newer));
  });
  it("rejects a changed issue even when its timestamp was not updated", async () => {
    const h = host(); const newer = { ...issue, description: "Newer saved value" };
    const a = h.addSlide("A", { [TAGS.issueJson]: JSON.stringify(newer) });
    await expect(new PowerPointService().saveSelectedIssue({ ...issue, description: "Stale draft" }, "A", issue)).rejects.toThrow(/changed since it was loaded/);
    expect(a.tags.items.find((t: any) => t.key === TAGS.issueJson).value).toBe(JSON.stringify(newer));
  });
  it("rejects a blank-slide draft if another editor created an issue there", async () => {
    const h = host(); h.addSlide("A", { [TAGS.issueJson]: JSON.stringify(issue) });
    await expect(new PowerPointService().renderSelectedIssue(issue, "A", null)).rejects.toThrow(/changed since it was loaded/);
  });
  it("rejects a stale selected slide before writing", async () => {
    const h = host(); h.addSlide("A"); const b = h.addSlide("B"); h.state.selected = "B";
    await expect(new PowerPointService().saveSelectedIssue(issue, "A")).rejects.toThrow(/Selected slide changed/);
    expect(b.tags.items).toEqual([]);
  });
  it("keeps writing to the captured slide when selection changes during awaits", async () => {
    const h = host(); const a = h.addSlide("A"); const b = h.addSlide("B");
    h.state.afterSync = () => { h.state.selected = "B"; };
    await new PowerPointService().saveSelectedIssue(issue, "A");
    expect(a.tags.items.find((t: any) => t.key === TAGS.issueJson)?.value).toBe(JSON.stringify(issue));
    expect(b.tags.items).toEqual([]);
  });
  it("retains the previous layout and metadata when replacement construction fails", async () => {
    const h = host(); const a = h.addSlide("A", { [TAGS.issueJson]: JSON.stringify(issue) });
    const old = a.shapes.addTextBox("old"); old.tags.add(TAGS.managed, "TRUE");
    h.state.failShapes = true;
    await expect(new PowerPointService().renderSelectedIssue({ ...issue, description: "New" }, "A")).rejects.toThrow(/Injected/);
    expect(old.delete).not.toHaveBeenCalled();
    expect(a.tags.items.find((t: any) => t.key === TAGS.issueJson).value).toBe(JSON.stringify(issue));
  });
  it("retains old shapes if the host rejects the staged sync", async () => {
    const h = host(); const a = h.addSlide("A", { [TAGS.issueJson]: JSON.stringify(issue) });
    const old = a.shapes.addTextBox("old"); old.tags.add(TAGS.managed, "TRUE");
    let failed = false;
    h.context.sync.mockImplementation(async () => {
      if (!failed && a.shapes.items.length > 1) { failed = true; throw new Error("Host rejected staged batch"); }
    });
    await expect(new PowerPointService().renderSelectedIssue(issue, "A")).rejects.toThrow(/Host rejected/);
    expect(old.delete).not.toHaveBeenCalled();
    expect(a.shapes.items).toEqual([old]);
  });
  it("preserves manual content and retains all actions when rendering overflow", async () => {
    const h = host(); const a = h.addSlide("A");
    const manual = a.shapes.addTextBox("manual image annotation");
    const old = a.shapes.addTextBox("old"); old.tags.add(TAGS.managed, "TRUE");
    const many = { ...issue, actions: Array.from({ length: 18 }, (_, i) => ({ id: `a${i}`, party: "MEP", required: `Task ${i}`, status: "Open", createdAt: issue.createdAt })) };
    await new PowerPointService().renderSelectedIssue(many, "A");
    expect(manual.delete).not.toHaveBeenCalled(); expect(old.delete).toHaveBeenCalledOnce();
    expect(a.shapes.items.some((s: any) => s.text.includes("8 more action"))).toBe(true);
    expect(a.shapes.items.filter((s: any) => s.text.startsWith("Task "))).toHaveLength(10);
    const stored = JSON.parse(a.tags.items.find((t: any) => t.key === TAGS.issueJson).value);
    expect(stored.actions).toHaveLength(18);
  });
  it.each(["shape", "move"])("retains old summary pages when staging fails at %s", async (step) => {
    const h = host(); h.addSlide("A", { [TAGS.issueJson]: JSON.stringify(issue) });
    const old = h.addSlide("summary", { [TAGS.summary]: "TRUE" });
    h.state.failShapes = step === "shape"; h.state.failMove = step === "move";
    await expect(new PowerPointService().generateSummary()).rejects.toThrow(/Injected/);
    expect(old.delete).not.toHaveBeenCalled(); expect(h.slides).toHaveLength(2);
  });
  it("paginates current actions and replaces previous summaries without modifying issue slides", async () => {
    const h = host();
    const many = { ...issue, actions: Array.from({ length: 18 }, (_, i) => ({ id: `a${i}`, party: "MEP", required: `Task ${i}`, status: "Open", createdAt: issue.createdAt })) };
    const a = h.addSlide("A", { [TAGS.issueJson]: JSON.stringify(many) });
    const old = h.addSlide("summary", { [TAGS.summary]: "TRUE" });
    expect(await new PowerPointService().generateSummary()).toEqual({ slidesCreated: 3, issueCount: 1, replacedCount: 1 });
    expect(old.delete).toHaveBeenCalledOnce(); expect(a.delete).not.toHaveBeenCalled();
    expect(h.slides).not.toContain(old);
    expect(h.slides.flatMap(s => s.shapes.items).filter(s => s.text.startsWith("Task "))).toHaveLength(18);
    await new PowerPointService().generateSummary();
    expect(h.slides).toHaveLength(4);
    expect(h.slides.filter(s => s.tags.items.some((t: any) => t.key === TAGS.summaryArchived && t.value === "TRUE"))).toHaveLength(0);
  });
  it("reports the affected slide instead of crashing or skipping malformed data", async () => {
    const h = host(); h.addSlide("A", { [TAGS.issueJson]: '{"id":"BAD"}' });
    await expect(new PowerPointService().getSummaryStats()).rejects.toThrow(/Slide 1: invalid IssueFlow metadata/);
  });
  it("leaves existing archives intact while replacing current summaries", async () => {
    const h = host(); h.addSlide("A", { [TAGS.issueJson]: JSON.stringify(issue) });
    const archive = h.addSlide("archive", { [TAGS.summary]: "TRUE", [TAGS.summaryArchived]: "TRUE" });
    const note = archive.shapes.addTextBox("Meeting note");
    await new PowerPointService().generateSummary();
    await new PowerPointService().generateSummary();
    expect(archive.delete).not.toHaveBeenCalled();
    expect(archive.shapes.items).toEqual([note]);
    expect((await new PowerPointService().getSummaryStats()).issueCount).toBe(1);
  });
  it("compares persisted card counts across service instances and resets changes on the next refresh", async () => {
    const h = host();
    const openIssue = { ...issue, actions: [{ id: "action", party: "MEP", required: "Fix", status: "Open", createdAt: issue.createdAt }] };
    const a = h.addSlide("A", { [TAGS.issueJson]: JSON.stringify(openIssue) });
    await new PowerPointService().generateSummary();
    expect(h.slides[0].shapes.items.some((s: any) => s.text.includes("First comparison baseline"))).toBe(true);
    expect(JSON.parse(h.slides[0].tags.items.find((t: any) => t.key === TAGS.summaryCounts).value)).toEqual([1,1,0,0,0,0,1,1,1]);
    a.tags.add(TAGS.issueJson, JSON.stringify({ ...openIssue, actions: [{ ...openIssue.actions[0], status: "Closed" }] }));
    await new PowerPointService().generateSummary();
    const changes = h.slides[0].shapes.items.map((s: any) => s.text);
    expect(changes.filter((text: string) => text === "−1")).toHaveLength(3);
    expect(changes.filter((text: string) => text === "+1")).toHaveLength(1);
    expect(changes.filter((text: string) => text === "—")).toHaveLength(5);
    await new PowerPointService().generateSummary();
    expect(h.slides[0].shapes.items.filter((s: any) => s.text === "—")).toHaveLength(9);
  });
  it.each(["broken", "[1]", "[1,1,0,0,0,0,1,-1,1]"])("treats invalid comparison metadata as a new baseline: %s", async (snapshot) => {
    const h = host(); h.addSlide("A", { [TAGS.issueJson]: JSON.stringify(issue) });
    h.addSlide("old", { [TAGS.summary]: "TRUE", [TAGS.summaryType]: "DASHBOARD", [TAGS.summaryCounts]: snapshot });
    await new PowerPointService().generateSummary();
    expect(h.slides[0].shapes.items.some((s: any) => s.text.includes("First comparison baseline"))).toBe(true);
  });
});
