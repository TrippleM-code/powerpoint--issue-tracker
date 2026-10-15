import type { PageSize } from "../layout/ice-meeting-layout";
export async function readIcePageSize(context: any): Promise<PageSize> {
  if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.10")) {
    throw new Error("ICE format needs PowerPoint API 1.10 to read your slide size safely. Update PowerPoint or use Standard format.");
  }
  const page = context.presentation.pageSetup;
  page.load("slideWidth,slideHeight");
  await context.sync();
  const size = {width:page.slideWidth,height:page.slideHeight};
  if (!Number.isFinite(size.width) || !Number.isFinite(size.height) || size.width <= 0 || size.height <= 0) throw new Error("Could not read the presentation slide size.");
  return size;
}
/** Keep existing summary content and pagination; fit coordinates to the ICE deck. */
export function scaledSummaryTarget(slide: any, size: PageSize): any {
  const convert = (o: any) => ({...o,left:o.left*size.width/960,top:o.top*size.height/540,width:o.width*size.width/960,height:o.height*size.height/540});
  return {fontScale:size.width/960,tags:slide.tags,shapes:{
    addTextBox:(text: string,options: any)=>slide.shapes.addTextBox(text,convert(options)),
    addGeometricShape:(type: string,options: any)=>slide.shapes.addGeometricShape(type,convert(options)),
  }};
}
