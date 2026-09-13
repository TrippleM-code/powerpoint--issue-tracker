import type { Issue } from "../domain/models";
import { SCHEMA_VERSION, TAGS } from "../storage/tag-names";

declare const PowerPoint: any;

const TRUE = "TRUE";

function readTag(tags: any[], key: string): string | undefined {
  return tags.find((tag) => tag.key === key)?.value;
}

async function getSelectedSlide(context: any): Promise<any> {
  const slides = context.presentation.getSelectedSlides();
  slides.load("items/id");
  await context.sync();

  if (slides.items.length !== 1) {
    throw new Error("Select exactly one issue slide.");
  }

  return slides.items[0];
}

async function loadShapeTags(context: any, slide: any): Promise<any[]> {
  slide.shapes.load("items/name");
  await context.sync();

  for (const shape of slide.shapes.items) {
    shape.tags.load("items/key,value");
  }
  await context.sync();
  return slide.shapes.items;
}

function addManagedTag(shape: any, role: string): void {
  shape.tags.add(TAGS.managed, TRUE);
  shape.tags.add(TAGS.managedRole, role);
}

function addText(
  slide: any,
  text: string,
  left: number,
  top: number,
  width: number,
  height: number,
  options: { size?: number; bold?: boolean; color?: string } = {}
): any {
  const shape = slide.shapes.addTextBox(text, { left, top, width, height });
  shape.textFrame.textRange.font.size = options.size ?? 10;
  shape.textFrame.textRange.font.bold = options.bold ?? false;
  shape.textFrame.textRange.font.color = options.color ?? "#152336";
  shape.textFrame.verticalAlignment = "Middle";
  return shape;
}

function addFilledRect(
  slide: any,
  left: number,
  top: number,
  width: number,
  height: number,
  fill: string,
  line: string
): any {
  const shape = slide.shapes.addGeometricShape("Rectangle", { left, top, width, height });
  shape.fill.setSolidColor(fill);
  shape.lineFormat.color = line;
  shape.lineFormat.weight = 0.7;
  return shape;
}

function addThinRect(
  slide: any,
  left: number,
  top: number,
  width: number,
  height: number,
  fill: string
): any {
  const safeW = Math.max(width, 0.6);
  const safeH = Math.max(height, 0.6);
  const shape = slide.shapes.addGeometricShape("Rectangle", { left, top, width: safeW, height: safeH });
  shape.fill.setSolidColor(fill);
  shape.lineFormat.transparency = 100;
  return shape;
}

export class PowerPointService {
  async readSelectedIssue(): Promise<Issue | null> {
    return PowerPoint.run(async (context: any) => {
      const slide = await getSelectedSlide(context);
      slide.tags.load("items/key,value");
      await context.sync();

      const json = readTag(slide.tags.items, TAGS.issueJson);
      if (!json) return null;

      try {
        return JSON.parse(json) as Issue;
      } catch {
        throw new Error("This slide contains invalid IssueFlow metadata.");
      }
    });
  }

  async saveSelectedIssue(issue: Issue): Promise<void> {
    await PowerPoint.run(async (context: any) => {
      const slide = await getSelectedSlide(context);
      slide.tags.add(TAGS.app, TRUE);
      slide.tags.add(TAGS.schemaVersion, SCHEMA_VERSION);
      slide.tags.add(TAGS.issueJson, JSON.stringify(issue));
      await context.sync();
    });
  }

  async renderSelectedIssue(issue: Issue): Promise<void> {
    await PowerPoint.run(async (context: any) => {
      const slide = await getSelectedSlide(context);
      const shapes = await loadShapeTags(context, slide);

      // Delete ONLY IssueFlow-managed shapes. Manual PowerPoint content survives.
      for (const shape of shapes) {
        const managed = shape.tags.items.some(
          (tag: any) => tag.key === TAGS.managed && tag.value === TRUE
        );
        if (managed) shape.delete();
      }
      await context.sync();

      slide.tags.add(TAGS.app, TRUE);
      slide.tags.add(TAGS.schemaVersion, SCHEMA_VERSION);
      slide.tags.add(TAGS.issueJson, JSON.stringify(issue));

      const NAVY = "#17344D";
      const GRID = "#C7D6E5";
      const TEXT = "#172538";
      const WHITE = "#FFFFFF";

      // 16:9 widescreen coordinates (points)
      const pageLeft = 8;
      const pageTop = 76;
      const pageBottom = 532;
      const bodyTop = 108;
      const bodyH = pageBottom - bodyTop;

      const descX = pageLeft;
      const descW = 250;
      const stripX = descX + descW + 10;
      const stripW = 38;
      const refX = stripX + stripW + 10;
      const refW = 352;
      const actionsX = refX + refW + 10;
      const actionsW = 320;

      // Issue description
      const descHeader = addFilledRect(slide, descX, pageTop, descW, 32, NAVY, NAVY);
      addManagedTag(descHeader, "DESC_HEADER");
      const descHeaderText = addText(slide, "ISSUE DESCRIPTION", descX + 18, pageTop + 7, descW - 36, 18, {
        size: 9.5, bold: true, color: WHITE
      });
      addManagedTag(descHeaderText, "DESC_HEADER_TEXT");

      const descBody = addFilledRect(slide, descX, bodyTop, descW, bodyH, WHITE, GRID);
      addManagedTag(descBody, "DESC_BODY");
      const descText = addText(slide, issue.description, descX + 20, bodyTop + 16, descW - 40, 60, {
        size: 10.5, color: TEXT
      });
      addManagedTag(descText, "DESC_TEXT");

      // 50/50 vertical Area / Room strip
      const stripBody = addFilledRect(slide, stripX, bodyTop, stripW, bodyH, WHITE, GRID);
      addManagedTag(stripBody, "LOCATION_BODY");
      const half = bodyH / 2;
      const divider = addThinRect(slide, stripX, bodyTop + half, stripW, 0.8, GRID);
      addManagedTag(divider, "LOCATION_DIVIDER");

      const rotatedW = Math.max(60, half - 34);
      const rotatedH = 20;
      const centerX = stripX + stripW / 2;

      const area = addText(
        slide, issue.areaCode,
        centerX - rotatedW / 2, bodyTop + half / 2 - rotatedH / 2,
        rotatedW, rotatedH, { size: 9.5, bold: true, color: TEXT }
      );
      area.rotation = 270;
      addManagedTag(area, "AREA_TEXT");

      const room = addText(
        slide, issue.roomSpace,
        centerX - rotatedW / 2, bodyTop + half + half / 2 - rotatedH / 2,
        rotatedW, rotatedH, { size: 9.5, bold: true, color: TEXT }
      );
      room.rotation = 270;
      addManagedTag(room, "ROOM_TEXT");

      // Reference Images: HEADER + BORDER ONLY.
      const refHeader = addFilledRect(slide, refX, pageTop, refW, 32, NAVY, NAVY);
      addManagedTag(refHeader, "REF_HEADER");
      const refHeaderText = addText(slide, "REFERENCE IMAGES", refX + 18, pageTop + 7, refW - 36, 18, {
        size: 9.5, bold: true, color: WHITE
      });
      addManagedTag(refHeaderText, "REF_HEADER_TEXT");

      for (const [role, l, t, w, h] of [
        ["REF_TOP", refX, bodyTop, refW, 0.8],
        ["REF_BOTTOM", refX, pageBottom, refW, 0.8],
        ["REF_LEFT", refX, bodyTop, 0.8, bodyH],
        ["REF_RIGHT", refX + refW, bodyTop, 0.8, bodyH],
      ] as const) {
        const border = addThinRect(slide, l, t, w, h, GRID);
        addManagedTag(border, role);
      }

      // Actions placeholder for P1 so layout is recognizable without implementing action workflow yet.
      const actHeader = addFilledRect(slide, actionsX, pageTop, actionsW, 32, NAVY, NAVY);
      addManagedTag(actHeader, "ACTIONS_HEADER");
      const actHeaderText = addText(slide, "ACTIONS", actionsX + 18, pageTop + 7, actionsW - 36, 18, {
        size: 9.5, bold: true, color: WHITE
      });
      addManagedTag(actHeaderText, "ACTIONS_HEADER_TEXT");
      const actBody = addFilledRect(slide, actionsX, bodyTop, actionsW, bodyH, WHITE, GRID);
      addManagedTag(actBody, "ACTIONS_BODY");
      const pending = addText(slide, "Actions will be added in Production P2.", actionsX + 18, bodyTop + 18, actionsW - 36, 26, {
        size: 10, color: "#6A7C90"
      });
      addManagedTag(pending, "ACTIONS_PLACEHOLDER");

      await context.sync();
    });
  }
}
