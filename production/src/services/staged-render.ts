/** Office batches are not transactions. Keep existing content until staging sync succeeds. */
export async function stageShapes(
  context: any,
  slide: any,
  build: (target: any) => void,
): Promise<() => Promise<void>> {
  const created: any[] = [];
  const track = (shape: any) => { created.push(shape); return shape; };
  const target = {
    shapes: {
      addTextBox: (...args: any[]) => track(slide.shapes.addTextBox(...args)),
      addGeometricShape: (...args: any[]) => track(slide.shapes.addGeometricShape(...args)),
    },
  };
  const rollback = async () => {
    for (const shape of created) shape.delete();
    await context.sync();
  };
  try {
    build(target);
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
