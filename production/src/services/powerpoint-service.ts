import type { Issue } from "../domain/models";
import { computeOverallStatus } from "../domain/status";
import { SCHEMA_VERSION, TAGS } from "../storage/tag-names";
import { loadSettings } from "./settings-service";

declare const PowerPoint: any;
declare const Office: any;

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

function imageDataUrlToBase64(value: string): string {
  const text = String(value || "").trim();
  const comma = text.indexOf(",");

  if (text.startsWith("data:image/") && comma >= 0) {
    return text.slice(comma + 1);
  }

  // Backward compatibility if a future settings version already stores raw Base64.
  return text;
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
  // PowerPoint API expects transparency in the range 0.0–1.0.
  shape.lineFormat.transparency = 1.0;
  return shape;
}


export interface IssueSlideRecord {
  slideId: string;
  slideNumber: number;
  issue: Issue;
}

export interface SelectedIssueState {
  slideId: string;
  issue: Issue | null;
}

export interface ApplyAllResult {
  updated: number;
  failed: number;
  errors: string[];
}

export interface SummaryStats {
  issueCount: number;
  actionCount: number;
  openCount: number;
  closedCount: number;
}

function statusColors(status: string): { fill: string; text: string } {
  switch (status.trim().toLowerCase()) {
    case "open": return { fill: "#FDE8E8", text: "#9F1D1D" };
    case "in progress": return { fill: "#E5F0FB", text: "#185A8B" };
    case "pending": return { fill: "#FFF2D8", text: "#8A5A00" };
    case "closed": return { fill: "#E6F5EA", text: "#27663A" };
    default: return { fill: "#EEF1F5", text: "#46586B" };
  }
}

async function addCleanSummarySlide(context: any, type: string): Promise<any> {
  const slides = context.presentation.slides;
  const count = slides.getCount();
  await context.sync();

  slides.add();
  await context.sync();

  const slide = slides.getItemAt(count.value);
  slide.shapes.load("items/id");
  await context.sync();

  for (const shape of slide.shapes.items) {
    shape.delete();
  }
  await context.sync();

  slide.tags.add(TAGS.summary, TRUE);
  slide.tags.add(TAGS.summaryType, type);
  slide.tags.add(TAGS.app, TRUE);
  return slide;
}

function buildDashboardSlide(slide: any, issues: Issue[]): void {
  const NAVY = "#17344D";
  const GRID = "#D7E1EB";
  const TEXT = "#172538";
  const TRACK = "#E9EEF4";

  // Brighter chart colors for stronger on-screen / PDF visibility.
  const STATUS_OPEN = "#E53935";
  const STATUS_IN_PROGRESS = "#1E88E5";
  const STATUS_PENDING = "#FFB300";
  const STATUS_CLOSED = "#43A047";
  const STATUS_NO_ACTIONS = "#8E24AA";

  const STATUS_COLORS: Record<string, string> = {
    "Open": STATUS_OPEN,
    "In Progress": STATUS_IN_PROGRESS,
    "Pending": STATUS_PENDING,
    "Closed": STATUS_CLOSED,
    "No actions": STATUS_NO_ACTIONS,
  };

  const PARTY_COLORS = [
    "#1E88E5",
    "#E53935",
    "#43A047",
    "#FB8C00",
    "#8E24AA",
    "#00ACC1",
    "#D81B60",
    "#7CB342",
  ];

  const allActions = issues.flatMap((issue) => issue.actions);

  // ---------------------------------------------------------------------------
  // OVERALL TRACKING STATUS = one derived status for each issue.
  // ---------------------------------------------------------------------------
  const issueStatusCounts = new Map<string, number>([
    ["Open", 0],
    ["In Progress", 0],
    ["Pending", 0],
    ["Closed", 0],
    ["No actions", 0],
  ]);

  issues.forEach((issue) => {
    const overall = computeOverallStatus(issue.actions.map((action) => action.status));
    issueStatusCounts.set(overall, (issueStatusCounts.get(overall) || 0) + 1);
  });

  // ---------------------------------------------------------------------------
  // ACTIONS TRACKING STATUS = current (not Closed) actions grouped by party.
  // ---------------------------------------------------------------------------
  const currentActions = allActions.filter(
    (action) => action.status.trim().toLowerCase() !== "closed"
  );

  const currentActionsByParty = new Map<string, number>();
  currentActions.forEach((action) => {
    currentActionsByParty.set(
      action.party,
      (currentActionsByParty.get(action.party) || 0) + 1
    );
  });

  const partyRows = Array.from(currentActionsByParty.entries())
    .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]))
    .slice(0, 7);

  // Header
  addFilledRect(slide, 0, 0, 960, 58, NAVY, NAVY);
  addText(slide, "ISSUEFLOW DASHBOARD", 34, 15, 420, 24, {
    size: 21,
    bold: true,
    color: "#FFFFFF",
  });
  addText(slide, `Refreshed ${new Date().toLocaleString()}`, 690, 19, 230, 16, {
    size: 8,
    color: "#DCE7EF",
  });

  // ===========================================================================
  // 1) OVERALL TRACKING STATUS
  // ===========================================================================
  addText(slide, "OVERALL TRACKING STATUS", 34, 72, 300, 20, {
    size: 13.5,
    bold: true,
    color: TEXT,
  });

  const overallCards: Array<[string, number, string]> = [
    ["Total Issues", issues.length, NAVY],
    ["Open", issueStatusCounts.get("Open") || 0, STATUS_OPEN],
    ["In Progress", issueStatusCounts.get("In Progress") || 0, STATUS_IN_PROGRESS],
    ["Pending", issueStatusCounts.get("Pending") || 0, STATUS_PENDING],
    ["Closed", issueStatusCounts.get("Closed") || 0, STATUS_CLOSED],
    ["No Actions", issueStatusCounts.get("No actions") || 0, STATUS_NO_ACTIONS],
  ];

  overallCards.forEach(([label, value, accent], index) => {
    const left = 28 + index * 151;

    const card = addFilledRect(slide, left, 98, 140, 48, "#FFFFFF", GRID);
    card.lineFormat.weight = 0.8;

    const accentBar = addFilledRect(slide, left, 98, 5, 48, accent, accent);
    accentBar.lineFormat.transparency = 1.0;

    addText(slide, String(value), left + 14, 105, 40, 20, {
      size: 16.5,
      bold: true,
      color: TEXT,
    });

    addText(slide, label, left + 14, 128, 116, 11, {
      size: 7.7,
      color: "#63768B",
    });
  });

  addText(slide, "Issues by overall status", 40, 160, 220, 18, {
    size: 11.5,
    bold: true,
    color: TEXT,
  });

  const issueStatusRows = ["Open", "In Progress", "Pending", "Closed", "No actions"];
  const maxIssueStatus = Math.max(
    1,
    ...issueStatusRows.map((status) => issueStatusCounts.get(status) || 0)
  );

  issueStatusRows.forEach((status, index) => {
    const count = issueStatusCounts.get(status) || 0;
    const top = 187 + index * 21;
    const barColor = STATUS_COLORS[status] || "#607D8B";

    addText(slide, status, 45, top - 1, 108, 14, {
      size: 8.4,
      color: TEXT,
    });

    addFilledRect(slide, 160, top + 1, 555, 12, TRACK, TRACK);

    if (count > 0) {
      addFilledRect(
        slide,
        160,
        top + 1,
        Math.max(10, 555 * count / maxIssueStatus),
        12,
        barColor,
        barColor
      );
    }

    addText(slide, String(count), 730, top - 1, 34, 14, {
      size: 8.5,
      bold: true,
      color: barColor,
    });
  });

  // Divider
  addThinRect(slide, 34, 305, 892, 1, GRID);

  // ===========================================================================
  // 2) ACTIONS TRACKING STATUS
  // ===========================================================================
  addText(slide, "ACTIONS TRACKING STATUS", 34, 318, 300, 20, {
    size: 13.5,
    bold: true,
    color: TEXT,
  });

  const actionCards: Array<[string, number]> = [
    ["Total Actions", allActions.length],
    ["Current Actions", currentActions.length],
    ["Parties with Current Actions", currentActionsByParty.size],
  ];

  actionCards.forEach(([label, value], index) => {
    const left = 40 + index * 298;

    addFilledRect(slide, left, 344, 270, 44, "#FFFFFF", GRID);

    addText(slide, String(value), left + 14, 351, 46, 18, {
      size: 15.5,
      bold: true,
      color: TEXT,
    });

    addText(slide, label, left + 66, 353, 185, 16, {
      size: 8.2,
      color: "#63768B",
    });
  });

  addText(slide, "Current actions by party", 40, 400, 250, 18, {
    size: 11.5,
    bold: true,
    color: TEXT,
  });

  if (partyRows.length === 0) {
    addText(slide, "No current actions.", 44, 430, 300, 20, {
      size: 9.5,
      color: "#63768B",
    });
  } else {
    const maxParty = Math.max(1, ...partyRows.map(([, count]) => count));

    partyRows.forEach(([party, count], index) => {
      const top = 427 + index * 15;
      const barColor = PARTY_COLORS[index % PARTY_COLORS.length] ?? "#1E88E5";

      addText(slide, party, 44, top - 1, 165, 12, {
        size: 7.8,
        color: TEXT,
      });

      addFilledRect(slide, 220, top + 1, 620, 9, TRACK, TRACK);

      addFilledRect(
        slide,
        220,
        top + 1,
        Math.max(10, 620 * count / maxParty),
        9,
        barColor,
        barColor
      );

      addText(slide, String(count), 855, top - 1, 32, 12, {
        size: 7.8,
        bold: true,
        color: barColor,
      });
    });

    if (currentActionsByParty.size > partyRows.length) {
      addText(
        slide,
        `+ ${currentActionsByParty.size - partyRows.length} more party / parties`,
        44,
        427 + partyRows.length * 15 + 3,
        300,
        12,
        { size: 7.2, color: "#63768B" }
      );
    }
  }
}

type RegisterGroup = { issue: Issue; start: number; count: number };

function groupIssuesForRegister(issues: Issue[], maxRows = 8): RegisterGroup[][] {
  const pages: RegisterGroup[][] = [];
  let page: RegisterGroup[] = [];
  let used = 0;

  for (const issue of issues) {
    const rows = Math.max(1, issue.actions.length);
    let start = 0;

    while (start < rows) {
      if (used >= maxRows) {
        pages.push(page);
        page = [];
        used = 0;
      }

      const available = maxRows - used;
      const count = Math.min(available, rows - start);
      page.push({ issue, start, count });
      used += count;
      start += count;

      if (used >= maxRows) {
        pages.push(page);
        page = [];
        used = 0;
      }
    }
  }

  if (page.length) pages.push(page);
  return pages.length ? pages : [[]];
}

function buildRegisterSlide(
  slide: any,
  groups: RegisterGroup[],
  pageNo: number,
  pageCount: number
): void {
  const NAVY = "#17344D";
  const GRID = "#D7E1EB";
  const TEXT = "#172538";

  // Compact register header to maximize usable table area.
  const headerH = 46;
  addFilledRect(slide, 0, 0, 960, headerH, NAVY, NAVY);
  addText(slide, "ACTION REGISTER", 24, 11, 430, 20, {
    size: 17,
    bold: true,
    color: "#FFFFFF"
  });
  addText(slide, `Page ${pageNo} of ${pageCount}`, 812, 14, 116, 14, {
    size: 7.5,
    color: "#DCE7EF"
  });

  // Compact table geometry.
  // 14 action rows fit comfortably on a 16:9 slide while remaining readable.
  const left = 12;
  const top = 58;
  const headH = 24;
  const rowH = 31;

  // Rebalanced widths: keep more space for Description and Action Required.
  const widths = [82, 92, 194, 104, 324, 128] as const;
  const headers = [
    "Issue ID",
    "Area / Room",
    "Description",
    "Action By",
    "Action Required",
    "Status",
  ];

  let x = left;
  headers.forEach((header, index) => {
    const width = widths[index] ?? 0;
    addFilledRect(slide, x, top, width, headH, NAVY, NAVY);
    addText(slide, header, x + 5, top + 5, width - 10, 13, {
      size: 7,
      bold: true,
      color: "#FFFFFF"
    });
    x += width;
  });

  const issueW = widths[0];
  const areaW = widths[1];
  const descW = widths[2];
  const partyW = widths[3];
  const actionW = widths[4];
  const statusW = widths[5];

  let rowIndex = 0;

  groups.forEach((group) => {
    const issue = group.issue;
    const actions = issue.actions.length
      ? issue.actions
      : [{
          id: "none",
          party: "—",
          required: "No actions",
          status: "—",
          createdAt: issue.createdAt,
        }];

    const visible = actions.slice(group.start, group.start + group.count);
    const groupTop = top + headH + rowIndex * rowH;
    const groupHeight = visible.length * rowH;
    const baseFill = rowIndex % 2 === 0 ? "#F8FAFC" : "#FFFFFF";

    let colX = left;

    addFilledRect(slide, colX, groupTop, issueW, groupHeight, baseFill, GRID);
    addText(
      slide,
      issue.id,
      colX + 5,
      groupTop + 4,
      issueW - 10,
      Math.max(16, groupHeight - 8),
      { size: 7.2, bold: true, color: "#1769AA" }
    );
    colX += issueW;

    addFilledRect(slide, colX, groupTop, areaW, groupHeight, baseFill, GRID);
    addText(
      slide,
      [issue.areaCode, issue.roomSpace].filter(Boolean).join("\n") || "—",
      colX + 5,
      groupTop + 4,
      areaW - 10,
      Math.max(16, groupHeight - 8),
      { size: 6.8, color: TEXT }
    );
    colX += areaW;

    addFilledRect(slide, colX, groupTop, descW, groupHeight, baseFill, GRID);
    addText(
      slide,
      issue.description || "—",
      colX + 5,
      groupTop + 4,
      descW - 10,
      Math.max(16, groupHeight - 8),
      { size: 6.8, color: TEXT }
    );
    colX += descW;

    visible.forEach((action, localIndex) => {
      const rowTop = groupTop + localIndex * rowH;

      addFilledRect(slide, colX, rowTop, partyW, rowH, "#FFFFFF", GRID);
      addText(
        slide,
        action.party,
        colX + 5,
        rowTop + 4,
        partyW - 10,
        rowH - 8,
        { size: 6.8, bold: true, color: TEXT }
      );

      const actionX = colX + partyW;
      addFilledRect(slide, actionX, rowTop, actionW, rowH, "#FFFFFF", GRID);
      addText(
        slide,
        action.required,
        actionX + 5,
        rowTop + 4,
        actionW - 10,
        rowH - 8,
        { size: 6.7, color: TEXT }
      );

      const statusX = actionX + actionW;
      const colors = statusColors(action.status);
      addFilledRect(slide, statusX, rowTop, statusW, rowH, colors.fill, GRID);
      addText(
        slide,
        action.status,
        statusX + 5,
        rowTop + 4,
        statusW - 10,
        rowH - 8,
        { size: 6.8, bold: true, color: colors.text }
      );
    });

    rowIndex += visible.length;
  });
}

export class PowerPointService {
  async readSelectedIssueState(): Promise<SelectedIssueState> {
    return PowerPoint.run(async (context: any) => {
      const slide = await getSelectedSlide(context);
      slide.tags.load("items/key,value");
      await context.sync();

      const json = readTag(slide.tags.items, TAGS.issueJson);
      if (!json) {
        return { slideId: slide.id, issue: null };
      }

      try {
        return {
          slideId: slide.id,
          issue: JSON.parse(json) as Issue,
        };
      } catch {
        throw new Error("This slide contains invalid IssueFlow metadata.");
      }
    });
  }

  async readSelectedIssue(): Promise<Issue | null> {
    const state = await this.readSelectedIssueState();
    return state.issue;
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
    const presentationSettings = loadSettings();

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

      // Keep every generated shape inside a 16:9 widescreen slide (960 x 540 pt).
      const pageRight = 952;
      const gap = 8;

      const descX = pageLeft;
      const descW = 230;

      const stripX = descX + descW + gap;
      const stripW = 36;

      const refX = stripX + stripW + gap;
      const refW = 300;

      const actionsX = refX + refW + gap;
      const actionsW = pageRight - actionsX;

      // Production P3 party header: one equal-width cell per configured party.
      const partyHeaderTop = 18;
      const partyHeaderH = 42;
      const identityX = pageRight - 300;
      const partyHeaderX = pageLeft;
      const partyHeaderW = identityX - pageLeft - 8;
      const parties = presentationSettings.parties;
      const partyCellW = partyHeaderW / Math.max(parties.length, 1);

      parties.forEach((party, index) => {
        const cellX = partyHeaderX + index * partyCellW;

        const cell = addFilledRect(
          slide, cellX, partyHeaderTop, partyCellW, partyHeaderH, "#FFFFFF", "#D5DFE9"
        );
        addManagedTag(cell, `PARTY_${index}_CELL`);

        if (party.logoDataUrl) {
          const logoBox = slide.shapes.addGeometricShape("Rectangle", {
            left: cellX + 6,
            top: partyHeaderTop + 5,
            width: Math.max(10, partyCellW - 12),
            height: partyHeaderH - 10,
          });
          logoBox.lineFormat.transparency = 1.0;
          logoBox.fill.setImage(imageDataUrlToBase64(party.logoDataUrl));
          addManagedTag(logoBox, `PARTY_${index}_LOGO`);
        } else {
          const name = addText(
            slide, party.name,
            cellX + 5, partyHeaderTop + 8,
            Math.max(10, partyCellW - 10), partyHeaderH - 16,
            { size: 8, bold: true, color: TEXT }
          );
          addManagedTag(name, `PARTY_${index}_NAME`);
        }
      });

      // Production P2.1 issue identity header.
      const identityTop = 18;
      const identityH = 42;
      const identityW = 292;

      const identityBox = addFilledRect(
        slide, identityX, identityTop, identityW, identityH, "#F8F1F1", "#E4DADA"
      );
      addManagedTag(identityBox, "ISSUE_IDENTITY_BOX");

      const issueLabel = addText(
        slide, "ISSUE ID", identityX + 12, identityTop + 5, 64, 12,
        { size: 7.5, bold: true, color: "#5F5660" }
      );
      addManagedTag(issueLabel, "ISSUE_ID_LABEL");

      const issueValue = addText(
        slide, issue.id, identityX + 12, identityTop + 16, 105, 20,
        { size: 15, bold: true, color: TEXT }
      );
      addManagedTag(issueValue, "ISSUE_ID_VALUE");

      const createdValue = addText(
        slide,
        `Created ${new Date(issue.createdAt).toLocaleString()}`,
        identityX + 124, identityTop + 7, identityW - 136, 12,
        { size: 7.2, color: TEXT }
      );
      addManagedTag(createdValue, "ISSUE_CREATED_VALUE");

      if (issue.updatedAt) {
        const updatedValue = addText(
          slide,
          `Updated ${new Date(issue.updatedAt).toLocaleString()}`,
          identityX + 124, identityTop + 23, identityW - 136, 12,
          { size: 7.2, color: TEXT }
        );
        addManagedTag(updatedValue, "ISSUE_UPDATED_VALUE");
      }

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

      // Actions
      const actHeader = addFilledRect(slide, actionsX, pageTop, actionsW, 32, NAVY, NAVY);
      addManagedTag(actHeader, "ACTIONS_HEADER");
      const actHeaderText = addText(slide, "ACTIONS", actionsX + 18, pageTop + 7, 120, 18, {
        size: 9.5, bold: true, color: WHITE
      });
      addManagedTag(actHeaderText, "ACTIONS_HEADER_TEXT");

      const overallStatus = computeOverallStatus(issue.actions.map((action) => action.status));
      const overallText = addText(
        slide,
        `Overall: ${overallStatus}`,
        actionsX + actionsW - 120,
        pageTop + 7,
        102,
        18,
        { size: 8.5, bold: true, color: overallStatus === "Closed" ? "#DDF4E4" : "#FFD8D8" }
      );
      addManagedTag(overallText, "ACTIONS_OVERALL_STATUS");

      const partyW = 78;
      const statusW = 76;
      const requiredW = actionsW - partyW - statusW;

      if (issue.actions.length === 0) {
        const actBody = addFilledRect(slide, actionsX, bodyTop, actionsW, bodyH, WHITE, GRID);
        addManagedTag(actBody, "ACTIONS_EMPTY_BODY");
        const noActions = addText(
          slide,
          "No actions added.",
          actionsX + 18,
          bodyTop + 18,
          actionsW - 36,
          24,
          { size: 10, color: "#6A7C90" }
        );
        addManagedTag(noActions, "ACTIONS_EMPTY_TEXT");
      } else {
        const rowH = bodyH / issue.actions.length;

        issue.actions.forEach((action, index) => {
          const rowTop = bodyTop + index * rowH;

          const partyCell = addFilledRect(slide, actionsX, rowTop, partyW, rowH, WHITE, GRID);
          addManagedTag(partyCell, `ACTION_${index}_PARTY_CELL`);
          const partyText = addText(
            slide, action.party,
            actionsX + 10, rowTop + 10, partyW - 20, Math.max(24, rowH - 20),
            { size: 9.5, bold: true, color: TEXT }
          );
          addManagedTag(partyText, `ACTION_${index}_PARTY_TEXT`);

          const requiredCell = addFilledRect(
            slide, actionsX + partyW, rowTop, requiredW, rowH, WHITE, GRID
          );
          addManagedTag(requiredCell, `ACTION_${index}_REQUIRED_CELL`);
          const requiredText = addText(
            slide, action.required,
            actionsX + partyW + 12, rowTop + 10, requiredW - 24, Math.max(24, rowH - 20),
            { size: 9.5, color: TEXT }
          );
          addManagedTag(requiredText, `ACTION_${index}_REQUIRED_TEXT`);

          let statusFill = "#EEF1F5";
          let statusTextColor = "#46586B";
          switch (action.status.trim().toLowerCase()) {
            case "open":
              statusFill = "#FDE8E8";
              statusTextColor = "#9F1D1D";
              break;
            case "in progress":
              statusFill = "#E5F0FB";
              statusTextColor = "#185A8B";
              break;
            case "pending":
              statusFill = "#FFF2D8";
              statusTextColor = "#8A5A00";
              break;
            case "closed":
              statusFill = "#E6F5EA";
              statusTextColor = "#27663A";
              break;
          }

          const statusCell = addFilledRect(
            slide, actionsX + partyW + requiredW, rowTop, statusW, rowH, statusFill, GRID
          );
          addManagedTag(statusCell, `ACTION_${index}_STATUS_CELL`);
          const statusText = addText(
            slide, action.status,
            actionsX + partyW + requiredW + 8,
            rowTop + 10,
            statusW - 16,
            Math.max(24, rowH - 20),
            { size: 9.5, bold: true, color: statusTextColor }
          );
          addManagedTag(statusText, `ACTION_${index}_STATUS_TEXT`);
        });
      }

      await context.sync();
    });
  }

  async readAllIssues(): Promise<IssueSlideRecord[]> {
    return PowerPoint.run(async (context: any) => {
      const slides = context.presentation.slides;
      slides.load("items/id");
      await context.sync();

      for (const slide of slides.items) {
        slide.tags.load("items/key,value");
      }
      await context.sync();

      const records: IssueSlideRecord[] = [];

      slides.items.forEach((slide: any, index: number) => {
        const isSummary = readTag(slide.tags.items, TAGS.summary) === TRUE;
        if (isSummary) return;

        const json = readTag(slide.tags.items, TAGS.issueJson);
        if (!json) return;

        try {
          const issue = JSON.parse(json) as Issue;
          if (issue && issue.id) {
            records.push({
              slideId: slide.id,
              slideNumber: index + 1,
              issue,
            });
          }
        } catch {
          // Skip invalid issue metadata; the current-slide workflow reports it directly.
        }
      });

      return records;
    });
  }

  async goToIssue(issueId: string): Promise<void> {
    const targetId = issueId.trim().toLowerCase();
    if (!targetId) {
      throw new Error("Enter an Issue ID.");
    }

    const records = await this.readAllIssues();
    const match = records.find(
      (record) => record.issue.id.trim().toLowerCase() === targetId
    );

    if (!match) {
      throw new Error(`Issue ID "${issueId.trim()}" was not found.`);
    }

    await PowerPoint.run(async (context: any) => {
      context.presentation.setSelectedSlides([match.slideId]);
      await context.sync();
    });
  }

  async getSummaryStats(): Promise<SummaryStats> {
    const records = await this.readAllIssues();
    const issues = records.map((record) => record.issue);
    const actionCount = issues.reduce((sum, issue) => sum + issue.actions.length, 0);
    const closedCount = issues.filter(
      (issue) => computeOverallStatus(issue.actions.map((action) => action.status)) === "Closed"
    ).length;

    return {
      issueCount: issues.length,
      actionCount,
      openCount: issues.length - closedCount,
      closedCount,
    };
  }

  async applySettingsToAllIssueSlides(): Promise<ApplyAllResult> {
    const originalSlideId = await PowerPoint.run(async (context: any) => {
      const selected = context.presentation.getSelectedSlides();
      selected.load("items/id");
      await context.sync();
      return selected.items[0]?.id || "";
    });

    const records = await this.readAllIssues();
    const result: ApplyAllResult = { updated: 0, failed: 0, errors: [] };

    for (const record of records) {
      try {
        await PowerPoint.run(async (context: any) => {
          context.presentation.setSelectedSlides([record.slideId]);
          await context.sync();
        });

        await this.renderSelectedIssue(record.issue);
        result.updated += 1;
      } catch (error) {
        result.failed += 1;
        result.errors.push(
          `${record.issue.id}: ${error instanceof Error ? error.message : String(error)}`
        );
      }
    }

    if (originalSlideId) {
      try {
        await PowerPoint.run(async (context: any) => {
          context.presentation.setSelectedSlides([originalSlideId]);
          await context.sync();
        });
      } catch {
        // Non-fatal: applying settings already completed.
      }
    }

    return result;
  }

  async generateSummary(): Promise<{ slidesCreated: number; issueCount: number }> {
    const records = await this.readAllIssues();
    const issues = records.map((record) => record.issue);
    const pages = groupIssuesForRegister(issues, 14);

    if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.8")) {
      throw new Error(
        "This PowerPoint version cannot move generated summary slides to the front. Update Microsoft 365/PowerPoint and try again."
      );
    }

    await PowerPoint.run(async (context: any) => {
      const slides = context.presentation.slides;
      slides.load("items/id");
      await context.sync();

      for (const slide of slides.items) {
        slide.tags.load("items/key,value");
      }
      await context.sync();

      const existingSummaryIds = slides.items
        .filter((slide: any) => readTag(slide.tags.items, TAGS.summary) === TRUE)
        .map((slide: any) => slide.id);

      existingSummaryIds.forEach((id: string) => slides.getItem(id).delete());
      await context.sync();

      const generatedSlides: any[] = [];

      const dashboard = await addCleanSummarySlide(context, "DASHBOARD");
      buildDashboardSlide(dashboard, issues);
      generatedSlides.push(dashboard);

      for (let index = 0; index < pages.length; index += 1) {
        const page = pages[index];
        if (!page) continue;

        const registerSlide = await addCleanSummarySlide(context, "REGISTER");
        buildRegisterSlide(registerSlide, page, index + 1, pages.length);
        generatedSlides.push(registerSlide);
      }

      await context.sync();

      // Summary must always be the first pages in the presentation.
      // Dashboard = slide 1, register pages follow in their generated order.
      generatedSlides.forEach((slide, index) => {
        slide.moveTo(index);
      });

      await context.sync();
    });

    return {
      slidesCreated: 1 + pages.length,
      issueCount: issues.length,
    };
  }

}
