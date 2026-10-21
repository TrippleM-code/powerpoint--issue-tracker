import { afterEach, describe, expect, it, vi } from "vitest";
import { PowerPointService } from "../../src/services/powerpoint-service";
import { TAGS } from "../../src/storage/tag-names";
import type { Issue } from "../../src/domain/models";

const issue: Issue = { id: "A", areaCode: "L1", roomSpace: "Room", description: "Description", createdAt: "2026-09-20T00:00:00Z", actions: [] };

function host() {
  let sequence = 0;
  const state = { template: undefined as unknown, selected: "A", failShapes: false, failMove: false, afterSync: () => {} };
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
        setZOrder: vi.fn((position: string) => {
          expect(position).toBe("SendToBack");
          shapes.splice(shapes.indexOf(shape), 1); shapes.unshift(shape);
        }),
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
    pageSetup: {slideWidth:720,slideHeight:540,load() {}},
    getSelectedSlides: () => ({ items: slides.filter(s => s.id === state.selected), load() {} }),
    setSelectedSlides: (ids: string[]) => { state.selected = ids[0]!; },
    slides: { items: slides, load() {}, getItem: (id: string) => {
      const slide = slides.find(s => s.id === id); if (!slide) throw new Error("Missing slide"); return slide;
    }, getItemAt: (index: number) => slides[index], getCount: () => ({ value: slides.length }),
      add: () => { addSlide(`new-${++sequence}`); },
    },
  } };
  vi.stubGlobal("Office", { context: { requirements: { isSetSupported: () => true }, document: { settings: { get: (key: string) => key === "issueflow.template.v1" ? state.template : undefined } } } });
  vi.stubGlobal("PowerPoint", { run: async (fn: any) => fn(context) });
  return { state, slides, addSlide, context };
}
afterEach(() => vi.unstubAllGlobals());

describe("PowerPoint service regression boundaries", () => {
  it.each(["standard", "ice"])("creates one editable notes box and preserves it through %s refreshes", async (format) => {
    const h = host(); h.state.template = { format };
    const slide = h.addSlide("A"), service = new PowerPointService();
    const notes = () => slide.shapes.items.filter((s: any) => s.tags.items.some((t: any) => t.key === TAGS.notesRole));
    await service.renderSelectedIssue(issue, "A");
    expect(notes()).toHaveLength(1);
    const note = notes()[0];
    expect(note.text).toBe("To type Notes");
    expect(note.lineFormat.weight).toBe(1);
    expect(note.tags.items.some((t: any) => t.key === TAGS.managed)).toBe(false);
    note.text = "Private slide-only notes";
    note.left = 120; note.top = 350; note.width = 210; note.height = 90;
    note.textFrame.textRange.font.size = 15;
    await service.renderSelectedIssue(issue, "A");
    expect((await service.applySettingsToAllIssueSlides()).failed).toBe(0);
    h.state.template = { format: format === "ice" ? "standard" : "ice" };
    await service.renderSelectedIssue(issue, "A");
    expect(notes()).toEqual([note]);
    expect(note).toMatchObject({text: "Private slide-only notes", left: 120, top: 350, width: 210, height: 90});
    expect(note.textFrame.textRange.font.size).toBe(15);
    expect(note.delete).not.toHaveBeenCalled();
    await service.generateSummary();
    expect(h.slides.filter(s => s !== slide).flatMap(s => s.shapes.items).some(s => s.text.includes("slide-only notes"))).toBe(false);
  });
  it("uses larger register text and continues every action onto taller paginated rows", async () => {
    const h = host();
    const many = {...issue, actions: Array.from({length: 21}, (_, i) => ({id: String(i), party: "MEP", required: `Unique action ${i}`, status: "Open", createdAt: issue.createdAt}))};
    h.addSlide("A", {[TAGS.issueJson]: JSON.stringify(many)});
    expect((await new PowerPointService().generateSummary()).slidesCreated).toBe(4);
    const registers = h.slides.filter(s => s.shapes.items.some((v: any) => v.text === "ACTION REGISTER"));
    expect(registers.map(s => s.shapes.items.filter((v: any) => v.text.startsWith("Unique action")).length)).toEqual([10,10,1]);
    for (const slide of registers) {
      for (const shape of slide.shapes.items.filter((s: any) => s.text === "Description" || s.text.startsWith("Unique action"))) {
        expect(shape.textFrame.textRange.font.size).toBe(11);
        expect(shape.textFrame.topMargin).toBe(0);
        expect(shape.textFrame.bottomMargin).toBe(0);
        expect(shape.top + shape.height).toBeLessThanOrEqual(540);
      }
    }
  });
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

describe("ICE integration",()=>{
 const template={format:"ice",title:"Shared ICE title",meetingNumber:"003",meetingDate:"2026-10-14"};
 it("converts both ways without losing references or manual content",async()=>{
  const h=host(),a=h.addSlide("A"),service=new PowerPointService();
  await service.renderSelectedIssue(issue,"A");
  const refs=a.shapes.items.filter((s:any)=>s.tags.items.some((t:any)=>t.key===TAGS.referenceRole));
  refs.forEach((s:any,i:number)=>s.text=`REF-${i}`);
  const manual=a.shapes.addTextBox("Manual",{left:100,top:300,width:100,height:100});
  h.state.template=template; await service.renderSelectedIssue(issue,"A");
  expect(a.shapes.items.some((s:any)=>s.text===template.title)).toBe(true);
  expect(a.shapes.items.some((s:any)=>s.text.includes("L1   |   Room"))).toBe(true);
  for(const shape of a.shapes.items){expect(shape.left+shape.width).toBeLessThanOrEqual(720.01);expect(shape.top+shape.height).toBeLessThanOrEqual(540.01);}
  expect(manual.delete).not.toHaveBeenCalled();
  await service.renderSelectedIssue(issue,"A");
  expect(a.shapes.items.filter((s:any)=>s.tags.items.some((t:any)=>t.key===TAGS.referenceRole))).toHaveLength(3);
  for(const shape of a.shapes.items.filter((s:any)=>s.setZOrder.mock.calls.length))expect(a.shapes.items.indexOf(shape)).toBeLessThan(a.shapes.items.indexOf(manual));
  refs.forEach((s:any,i:number)=>{expect(s.text).toBe(`REF-${i}`);expect(s.delete).not.toHaveBeenCalled();});
  h.state.template={format:"standard"};await service.renderSelectedIssue(issue,"A");
  expect(refs[0].left).toBe(776);expect(refs[0].text).toBe("REF-0");
  expect(a.shapes.items.some((s:any)=>s.text==="ISSUE DESCRIPTION")).toBe(true);
 });
 it("shares titles across slides and retains overflow in the register",async()=>{
  const h=host();h.state.template=template;
  const many={...issue,actions:Array.from({length:18},(_,i)=>({id:`a${i}`,party:"MEP",required:`Task ${i}`,status:"Open",createdAt:issue.createdAt}))};
  const a=h.addSlide("A",{[TAGS.issueJson]:JSON.stringify(many)}),b=h.addSlide("B",{[TAGS.issueJson]:JSON.stringify({...issue,id:"B"})});
  const service=new PowerPointService();expect(await service.applySettingsToAllIssueSlides()).toMatchObject({updated:2,failed:0});
  for(const s of [a,b])expect(s.shapes.items.some((v:any)=>v.text===template.title)).toBe(true);
  expect(a.shapes.items.some((s:any)=>s.text.includes("8 more action"))).toBe(true);
  expect((await service.readAllIssues())[0]!.issue.actions).toHaveLength(18);
  await service.generateSummary();expect(h.slides.flatMap(s=>s.shapes.items).some(s=>s.text==="Task 17")).toBe(true);
  const dashboard=h.slides[0]; const label=dashboard.shapes.items.find((s:any)=>s.text==="Total Issues");
  expect(label.textFrame.topMargin).toBe(0); expect(label.textFrame.bottomMargin).toBe(0);
  expect(label.textFrame.textRange.font.size).toBeCloseTo(11*720/960);
  for(const slide of h.slides)for(const shape of slide.shapes.items){expect(shape.left+shape.width).toBeLessThanOrEqual(720.01);expect(shape.top+shape.height).toBeLessThanOrEqual(540.01);}
 });
 it("keeps the old layout on failed ICE staging",async()=>{
  const h=host(),a=h.addSlide("A"),service=new PowerPointService();await service.renderSelectedIssue(issue,"A");
  const before=[...a.shapes.items];h.state.template=template;h.state.failShapes=true;
  await expect(service.renderSelectedIssue(issue,"A")).rejects.toThrow(/Injected/);expect(a.shapes.items).toEqual(before);
 });
});
describe("Standard transparent, contiguous layout",()=>{
 it("preserves manual images across repeated refreshes without opaque body panels or column gutters",async()=>{
  const h=host(),slide=h.addSlide("A"),service=new PowerPointService();
  const picture=slide.shapes.addGeometricShape("Rectangle",{left:200,top:180,width:450,height:200});
  const withAction={...issue,actions:[{id:"a",party:"MEP",required:"Review",status:"Open",createdAt:issue.createdAt}]};
  await service.renderSelectedIssue(withAction,"A");await service.renderSelectedIssue(withAction,"A");
  expect(picture.delete).not.toHaveBeenCalled();expect(picture.left).toBe(200);expect(picture.width).toBe(450);
  expect(picture.setZOrder).not.toHaveBeenCalled();
  for(const shape of slide.shapes.items.filter((s:any)=>s.setZOrder.mock.calls.length))expect(slide.shapes.items.indexOf(shape)).toBeLessThan(slide.shapes.items.indexOf(picture));
  const role=(value:string)=>slide.shapes.items.find((s:any)=>s.tags.items.some((t:any)=>t.key===TAGS.managedRole&&t.value===value));
  for(const value of ["DESC_BODY","LOCATION_BODY","ACTION_0_PARTY_CELL","ACTION_0_REQUIRED_CELL"])expect(role(value).fill.transparency).toBe(1);
  expect(role("DESC_BODY").left+role("DESC_BODY").width).toBe(role("LOCATION_BODY").left);
  expect(role("LOCATION_BODY").left+role("LOCATION_BODY").width).toBe(role("REF_HEADER").left);
  expect(role("REF_HEADER").left+role("REF_HEADER").width).toBe(role("ACTIONS_HEADER").left);
  expect(role("ACTIONS_HEADER").left+role("ACTIONS_HEADER").width).toBe(952);
  await service.renderSelectedIssue(issue,"A");expect(role("ACTIONS_EMPTY_BODY").fill.transparency).toBe(1);
 });
});
