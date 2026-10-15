/** Office batches are not transactions. Keep existing content until staging sync succeeds. */
export async function stageShapes(
  context: any,
  slide: any,
  build: (target: any) => void,
): Promise<() => Promise<void>> {
  if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.8")) {
    throw new Error("Keeping the layout behind images requires PowerPoint API 1.8. Update PowerPoint and try again. The existing layout was retained.");
  }
  const created: any[] = [];
  const backgrounds: any[] = [];
  const track = (shape: any) => { created.push(shape); return shape; };
  const target = {
    shapes: {
      addTextBox: (...args: any[]) => track(slide.shapes.addTextBox(...args)),
      addGeometricShape: (...args: any[]) => {
        const shape = track(slide.shapes.addGeometricShape(...args));
        backgrounds.push(shape);
        return shape;
      },
    },
  };
  const rollback = async () => {
    for (const shape of created) shape.delete();
    await context.sync();
  };
  try {
    build(target);
    await context.sync();
    // Reverse traversal preserves the generated backgrounds' relative order.
    // Existing pictures/objects stay untouched; generated text stays in front.
    for (const shape of [...backgrounds].reverse()) shape.setZOrder("SendToBack");
    await context.sync();
    return rollback;
  } catch (error) {
    try {
      await rollback();
    } catch {
      throw new Error(`Refresh failed; the previous layout was retained, but temporary shapes may remain. ${String(error)}`);
    }
    throw error;
  }
}
