import { afterEach, expect, it, vi } from "vitest";
import { stageShapes } from "../../src/services/staged-render";

afterEach(() => vi.unstubAllGlobals());

it("reports an Office property location after rollback without dumping debug bodies", async () => {
  const h = host(); const original = [...h.items];
  const error = Object.assign(new Error("InvalidArgument"), {
    debugInfo: {errorLocation: "ShapeFill.transparency", fullStatements: "private document text"},
  });
  h.context.sync.mockRejectedValueOnce(error);
  await expect(stageShapes(h.context, h.slide, target => target.shapes.addTextBox("notes")))
    .rejects.toThrow(/^InvalidArgument \(ShapeFill.transparency\)$/);
  expect(h.items).toEqual(original);
});

function host(supported = true) {
  vi.stubGlobal("Office", { context: { requirements: { isSetSupported: () => supported } } });
  const items: any[] = [{ name: "picture" }, { name: "annotation" }];
  const add = (name: string) => {
    const shape = { name, delete() { items.splice(items.indexOf(shape), 1); },
      setZOrder: vi.fn(() => { items.splice(items.indexOf(shape), 1); items.unshift(shape); }) };
    items.push(shape); return shape;
  };
  return { items, context: { sync: vi.fn(async () => {}) },
    slide: { shapes: { addTextBox: add, addGeometricShape: add } } };
}

it("keeps backgrounds ordered below existing pictures and generated text", async () => {
  const h = host();
  await stageShapes(h.context, h.slide, target => {
    target.shapes.addGeometricShape("panel");
    target.shapes.addTextBox("heading");
    target.shapes.addGeometricShape("border");
  });
  expect(h.items.map(s => s.name)).toEqual(["panel", "border", "picture", "annotation", "heading"]);
  expect(h.items[4].setZOrder).not.toHaveBeenCalled();
});

it("rolls back new shapes if applying layer order fails", async () => {
  const h = host(); const original = [...h.items];
  h.context.sync.mockRejectedValueOnce(new Error("creation failed"));
  await expect(stageShapes(h.context, h.slide, target => target.shapes.addGeometricShape("panel"))).rejects.toThrow("creation failed");
  expect(h.items).toEqual(original);
  h.context.sync.mockResolvedValueOnce(undefined).mockRejectedValueOnce(new Error("layer failure"));
  await expect(stageShapes(h.context, h.slide, target => target.shapes.addGeometricShape("panel"))).rejects.toThrow("layer failure");
  expect(h.items).toEqual(original);
});

it("rejects unsupported hosts before creating shapes", async () => {
  const h = host(false); const build = vi.fn();
  await expect(stageShapes(h.context, h.slide, build)).rejects.toThrow("API 1.8");
  expect(build).not.toHaveBeenCalled(); expect(h.items).toHaveLength(2);
});
