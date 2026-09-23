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
  it("paginates all actions and replaces summaries without modifying issue slides", async () => {
    const h = host();
    const many = { ...issue, actions: Array.from({ length: 18 }, (_, i) => ({ id: `a${i}`, party: "MEP", required: `Task ${i}`, status: "Open", createdAt: issue.createdAt })) };
    const a = h.addSlide("A", { [TAGS.issueJson]: JSON.stringify(many) });
    const old = h.addSlide("summary", { [TAGS.summary]: "TRUE" });
    expect(await new PowerPointService().generateSummary()).toEqual({ slidesCreated: 3, issueCount: 1 });
    expect(old.delete).toHaveBeenCalledOnce(); expect(a.delete).not.toHaveBeenCalled();
    expect(h.slides.at(-1)).toBe(a);
    expect(h.slides.flatMap(s => s.shapes.items).filter(s => s.text.startsWith("Task "))).toHaveLength(18);
    await new PowerPointService().generateSummary();
    expect(h.slides).toHaveLength(4);
  });
  it("reports the affected slide instead of crashing or skipping malformed data", async () => {
    const h = host(); h.addSlide("A", { [TAGS.issueJson]: '{"id":"BAD"}' });
    await expect(new PowerPointService().getSummaryStats()).rejects.toThrow(/Slide 1: invalid IssueFlow metadata/);
  });
});
