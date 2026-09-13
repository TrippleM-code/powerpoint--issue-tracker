
const TAG_APP = "PIT_APP";
const TAG_ISSUE_ID = "PIT_ISSUE_ID";
const TAG_BOX_TYPE = "PIT_BOX_TYPE";
const TAG_STATUS = "PIT_STATUS";
const TAG_SUMMARY = "PIT_SUMMARY";
const TAG_STATUS_LIBRARY = "PIT_STATUS_LIBRARY";
const TAG_CREATED_AT = "PIT_CREATED_AT";
const TAG_UPDATED_AT = "PIT_UPDATED_AT";
const TAG_SNAPSHOT_HASH = "PIT_SNAPSHOT_HASH";
const TAG_SUMMARY_TYPE = "PIT_SUMMARY_TYPE";
const TAG_FOOTER = "PIT_FOOTER";
const TAG_LINK_ID = "PIT_LINK_ID";

const BOX_DESCRIPTION = "DESCRIPTION";
const BOX_STATUS = "STATUS";
const BOX_REMARK = "REMARK";
const DEFAULT_STATUSES = ["Open", "In Progress", "Pending", "Closed"];
const ui = {};

Office.onReady((info) => {
  if (info.host !== Office.HostType.PowerPoint) return;
  bindUi();
  initialize().catch(showError);
});

function bindUi() {
  ["issueId","setIssueId","issueStatus","addDescription","addStatus","addRemark",
   "statusSelect","applyStatus","statusList","newStatus","addStatusValue",
   "issueNavigator","goToIssue","resetLayout","repairCurrentSlide","validate","generateSummary","result"]
  .forEach(id => ui[id] = document.getElementById(id));

  ui.setIssueId.addEventListener("click", () => setIssueId().catch(showError));
  ui.addDescription.addEventListener("click", () => addBox(BOX_DESCRIPTION).catch(showError));
  ui.addStatus.addEventListener("click", () => addBox(BOX_STATUS).catch(showError));
  ui.addRemark.addEventListener("click", () => addBox(BOX_REMARK).catch(showError));
  ui.applyStatus.addEventListener("click", () => applyStatus().catch(showError));
  ui.addStatusValue.addEventListener("click", () => addStatusValue().catch(showError));
  ui.goToIssue.addEventListener("click", () => goToIssue().catch(showError));
  ui.resetLayout.addEventListener("click", () => resetBoxLayout().catch(showError));
  ui.repairCurrentSlide.addEventListener("click", () => repairCurrentSlide().catch(showError));
  ui.validate.addEventListener("click", () => validatePresentation().catch(showError));
  ui.generateSummary.addEventListener("click", () => generateSummary().catch(showError));
}

async function initialize() {
  await ensureStatusLibrary();
  await loadCurrentSlideIssue();
  await refreshStatusUi();
  await refreshIssueNavigator();
}

function cleanIssueId(v) { return String(v || "").trim().toUpperCase(); }
function cleanStatus(v) { return String(v || "").trim(); }

async function getSelectedSlide(context) {
  const selected = context.presentation.getSelectedSlides();
  selected.load("items/id");
  await context.sync();
  if (!selected.items.length) throw new Error("Select a slide first.");
  return selected.items[0];
}

async function getTagValue(tags, key, context) {
  const tag = tags.getItemOrNullObject(key);
  tag.load("value,isNullObject");
  await context.sync();
  return tag.isNullObject ? null : tag.value;
}

async function ensureStatusLibrary() {
  await PowerPoint.run(async (context) => {
    const existing = await getTagValue(context.presentation.tags, TAG_STATUS_LIBRARY, context);
    if (!existing) {
      context.presentation.tags.add(TAG_STATUS_LIBRARY, JSON.stringify(DEFAULT_STATUSES));
      await context.sync();
    }
  });
}

async function getStatusLibrary() {
  return PowerPoint.run(async (context) => {
    const value = await getTagValue(context.presentation.tags, TAG_STATUS_LIBRARY, context);
    try {
      const parsed = JSON.parse(value || "[]");
      return Array.isArray(parsed) && parsed.length ? parsed : DEFAULT_STATUSES.slice();
    } catch {
      return DEFAULT_STATUSES.slice();
    }
  });
}

async function saveStatusLibrary(statuses) {
  await PowerPoint.run(async (context) => {
    context.presentation.tags.add(TAG_STATUS_LIBRARY, JSON.stringify(statuses));
    await context.sync();
  });
}

async function loadCurrentSlideIssue() {
  await PowerPoint.run(async (context) => {
    const slide = await getSelectedSlide(context);
    const issueId = await getTagValue(slide.tags, TAG_ISSUE_ID, context);
    ui.issueId.value = issueId || "";
    ui.issueStatus.textContent = issueId ? `Current slide: ${issueId}` : "Current slide has no issue ID.";
  });
}

async function setIssueId() {
  const issueId = cleanIssueId(ui.issueId.value);
  if (!issueId) throw new Error("Enter an issue ID, for example MEP-001.");

  await PowerPoint.run(async (context) => {
    const selected = context.presentation.getSelectedSlides();
    selected.load("items/id");
    const slides = context.presentation.slides;
    slides.load("items/id,items/tags/key,items/tags/value");
    await context.sync();

    if (!selected.items.length) throw new Error("Select a slide first.");
    const slide = selected.items[0];

    for (let i = 0; i < slides.items.length; i++) {
      const candidate = slides.items[i];
      if (candidate.id === slide.id) continue;
      const tag = candidate.tags.items.find(t => t.key === TAG_ISSUE_ID);
      if (cleanIssueId(tag?.value) === issueId) {
        throw new Error(`${issueId} already exists on slide ${i + 1}.`);
      }
    }

    slide.tags.add(TAG_APP, "POWERPOINT_ISSUE_TRACKER");
    slide.tags.add(TAG_ISSUE_ID, issueId);

    const createdTag = slide.tags.getItemOrNullObject(TAG_CREATED_AT);
    createdTag.load("value,isNullObject");
    await context.sync();
    if (createdTag.isNullObject || !createdTag.value) {
      slide.tags.add(TAG_CREATED_AT, new Date().toISOString());
    }

    slide.shapes.load("items/id,items/tags/key,items/tags/value");
    await context.sync();

    for (const shape of slide.shapes.items) {
      const typeTag = shape.tags.items.find(t => t.key === TAG_BOX_TYPE);
      if (!typeTag) continue;
      shape.tags.add(TAG_APP, "POWERPOINT_ISSUE_TRACKER");
      shape.tags.add(TAG_ISSUE_ID, issueId);
    }

    await context.sync();
  });

  ui.issueStatus.textContent = `Current slide: ${issueId}`;
  showResult(`Issue ID set to ${issueId}. Existing tracker boxes were synchronized.`);
  await refreshIssueNavigator();
}

async function requireIssueId() {
  return PowerPoint.run(async (context) => {
    const slide = await getSelectedSlide(context);
    const issueId = await getTagValue(slide.tags, TAG_ISSUE_ID, context);
    if (!issueId) throw new Error("Set the issue ID on this slide first.");
    return { issueId };
  });
}

function defaultBoxOptions(type) {
  if (type === BOX_DESCRIPTION) {
    return { left: 36, top: 110, width: 330, height: 58 };
  }
  if (type === BOX_STATUS) {
    return { left: 390, top: 110, width: 180, height: 58 };
  }
  return { left: 594, top: 110, width: 330, height: 58 };
}


function getTrackerBoxStyle(type, status = "") {
  if (type === BOX_DESCRIPTION) {
    return {
      fill: "#EAF3FF",
      border: "#3A78C2",
      text: "#173A63",
      heading: "#174A7E"
    };
  }

  if (type === BOX_REMARK) {
    return {
      fill: "#F3F4F6",
      border: "#7A8590",
      text: "#26323D",
      heading: "#35424E"
    };
  }

  const s = cleanStatus(status).toLowerCase();

  if (s === "open") {
    return {
      fill: "#FDE8E8",
      border: "#D64545",
      text: "#8A1F1F",
      heading: "#8A1F1F"
    };
  }

  if (s === "in progress") {
    return {
      fill: "#E6F0FF",
      border: "#3A78C2",
      text: "#174A7E",
      heading: "#174A7E"
    };
  }

  if (s === "pending") {
    return {
      fill: "#FFF4D6",
      border: "#D89A2B",
      text: "#7A560F",
      heading: "#7A560F"
    };
  }

  if (s === "closed") {
    return {
      fill: "#E7F6EC",
      border: "#2F8F5B",
      text: "#1E603D",
      heading: "#1E603D"
    };
  }

  // Neutral fallback for custom statuses.
  return {
    fill: "#F3F4F6",
    border: "#7A8590",
    text: "#26323D",
    heading: "#35424E"
  };
}

function trackerBoxHeading(type) {
  if (type === BOX_DESCRIPTION) return "Description";
  if (type === BOX_STATUS) return "Status";
  return "Remark";
}

function applyTrackerBoxStyle(shape, type, status = "") {
  const style = getTrackerBoxStyle(type, status);

  shape.fill.setSolidColor(style.fill);
  shape.lineFormat.color = style.border;
  shape.lineFormat.weight = 1.5;

  const range = shape.textFrame.textRange;
  range.font.name = "Aptos";
  range.font.size = 12;
  range.font.color = style.text;
  range.font.bold = false;

  // Make only the heading on the first line bold.
  // getSubstring is used so the user's body text remains normal weight.
  const heading = trackerBoxHeading(type);
  try {
    const headingRange = range.getSubstring(0, heading.length);
    headingRange.font.bold = true;
    headingRange.font.color = style.heading;
  } catch (e) {
    // Fallback: keep the whole box readable even if substring formatting
    // is unavailable on an older host.
    console.warn("Heading-specific formatting is unavailable.", e);
  }
}

function initialText(type, status) {
  if (type === BOX_DESCRIPTION) return "Description\nEnter description here";
  if (type === BOX_STATUS) return `Status\n${status}`;
  return "Remark\nEnter remark here";
}

async function addBox(type) {
  const { issueId } = await requireIssueId();
  const statuses = await getStatusLibrary();
  const status = statuses[0] || "Open";

  await PowerPoint.run(async (context) => {
    const slide = await getSelectedSlide(context);
    slide.shapes.load("items/id,items/tags/key,items/tags/value");
    await context.sync();

    for (const shape of slide.shapes.items) {
      const typeTag = shape.tags.items.find(t => t.key === TAG_BOX_TYPE);
      if (typeTag?.value === type) throw new Error(`${type} box already exists on this slide.`);
    }

    const box = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, defaultBoxOptions(type));
    box.name = `PIT_${type}_${issueId}`;
    box.textFrame.wordWrap = true;
    box.textFrame.autoSizeSetting = PowerPoint.ShapeAutoSize.autoSizeNone;
    box.textFrame.textRange.text = initialText(type, status);
    applyTrackerBoxStyle(box, type, status);
    box.tags.add(TAG_APP, "POWERPOINT_ISSUE_TRACKER");
    box.tags.add(TAG_ISSUE_ID, issueId);
    box.tags.add(TAG_BOX_TYPE, type);
    if (type === BOX_STATUS) box.tags.add(TAG_STATUS, status);

    await context.sync();
  });

  showResult(`${type} box added.`);
}

async function resetBoxLayout() {
  await PowerPoint.run(async (context) => {
    const slide = await getSelectedSlide(context);
    slide.shapes.load("items/id,items/tags/key,items/tags/value");
    await context.sync();

    const positions = {
      [BOX_DESCRIPTION]: { left: 36, top: 110, width: 330, height: 58 },
      [BOX_STATUS]: { left: 390, top: 110, width: 180, height: 58 },
      [BOX_REMARK]: { left: 594, top: 110, width: 330, height: 58 }
    };

    let count = 0;

    for (const shape of slide.shapes.items) {
      const typeTag = shape.tags.items.find(t => t.key === TAG_BOX_TYPE);
      if (!typeTag) continue;

      const position = positions[typeTag.value];
      if (!position) continue;

      shape.left = position.left;
      shape.top = position.top;
      shape.width = position.width;
      shape.height = position.height;
      shape.textFrame.wordWrap = true;
      shape.textFrame.autoSizeSetting = PowerPoint.ShapeAutoSize.autoSizeNone;

      const statusTag = shape.tags.items.find(t => t.key === TAG_STATUS);
      applyTrackerBoxStyle(shape, typeTag.value, statusTag?.value || "");
      count++;
    }

    await context.sync();
    showResult(`Reset ${count} tracker box(es) to the standard one-line layout.`);
  });
}

async function refreshStatusUi() {
  const statuses = await getStatusLibrary();
  ui.statusSelect.innerHTML = "";
  statuses.forEach(s => {
    const o = document.createElement("option");
    o.value = s; o.textContent = s; ui.statusSelect.appendChild(o);
  });

  ui.statusList.innerHTML = "";
  statuses.forEach(status => {
    const row = document.createElement("div");
    row.className = "status-item";
    const span = document.createElement("span");
    span.textContent = status;
    const btn = document.createElement("button");
    btn.textContent = "Remove";
    btn.addEventListener("click", () => removeStatus(status).catch(showError));
    row.append(span, btn);
    ui.statusList.appendChild(row);
  });
}

async function addStatusValue() {
  const value = cleanStatus(ui.newStatus.value);
  if (!value) throw new Error("Enter a status name.");
  const statuses = await getStatusLibrary();
  if (statuses.some(s => s.toLowerCase() === value.toLowerCase())) throw new Error("That status already exists.");
  statuses.push(value);
  await saveStatusLibrary(statuses);
  ui.newStatus.value = "";
  await refreshStatusUi();
  showResult(`Status "${value}" added.`);
}

async function readIssues() {
  return PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/id,items/tags/key,items/tags/value,items/shapes/items/id,items/shapes/items/tags/key,items/shapes/items/tags/value");
    await context.sync();

    const tracked = [];

    slides.items.forEach((slide, index) => {
      const issueTag = slide.tags.items.find(t => t.key === TAG_ISSUE_ID);
      if (!issueTag) return;

      const issue = {
        issueId: issueTag.value,
        slideId: slide.id,
        slideIndex: index + 1,
        description: "",
        status: "",
        remark: "",
        hasDescription: false,
        hasStatus: false,
        hasRemark: false,
        shapeEntries: []
      };

      slide.shapes.items.forEach(shape => {
        const typeTag = shape.tags.items.find(t => t.key === TAG_BOX_TYPE);
        if (!typeTag) return;

        issue.shapeEntries.push({
          proxy: shape,
          type: typeTag.value,
          storedStatus: shape.tags.items.find(t => t.key === TAG_STATUS)?.value || ""
        });

        shape.textFrame.textRange.load("text");
      });

      tracked.push(issue);
    });

    await context.sync();

    tracked.forEach(issue => {
      issue.shapeEntries.forEach(entry => {
        const text = entry.proxy.textFrame.textRange.text || "";
        const body = text.split(/\r?\n/).slice(1).join("\n").trim();

        if (entry.type === BOX_DESCRIPTION) {
          issue.description = body;
          issue.hasDescription = true;
        } else if (entry.type === BOX_STATUS) {
          issue.status = entry.storedStatus || body;
          issue.hasStatus = true;
        } else if (entry.type === BOX_REMARK) {
          issue.remark = body;
          issue.hasRemark = true;
        }
      });

      delete issue.shapeEntries;
    });

    return tracked;
  });
}

async function repairCurrentSlide() {
  await PowerPoint.run(async (context) => {
    const slide = await getSelectedSlide(context);
    const issueId = await getTagValue(slide.tags, TAG_ISSUE_ID, context);
    if (!issueId) throw new Error("Set the issue ID on this slide first.");

    slide.shapes.load("items/id,items/tags/key,items/tags/value");
    await context.sync();

    let repaired = 0;
    for (const shape of slide.shapes.items) {
      const typeTag = shape.tags.items.find(t => t.key === TAG_BOX_TYPE);
      if (!typeTag) continue;
      shape.tags.add(TAG_APP, "POWERPOINT_ISSUE_TRACKER");
      shape.tags.add(TAG_ISSUE_ID, issueId);

      const statusTag = shape.tags.items.find(t => t.key === TAG_STATUS);
      applyTrackerBoxStyle(shape, typeTag.value, statusTag?.value || "");
      repaired++;
    }
    await context.sync();
    showResult(`Repair complete for ${issueId}. ${repaired} tracker box(es) synchronized.`);
  });
}

async function validatePresentation() {
  const issues = await readIssues();
  const statuses = await getStatusLibrary();
  const errors = [];
  const seen = new Map();

  issues.forEach(issue => {
    const id = cleanIssueId(issue.issueId);
    if (seen.has(id)) errors.push(`${id}: duplicate on slides ${seen.get(id)} and ${issue.slideIndex}.`);
    else seen.set(id, issue.slideIndex);

    if (!issue.hasDescription) errors.push(`${id}: missing Description box.`);
    if (!issue.hasStatus) errors.push(`${id}: missing Status box.`);
    if (!issue.hasRemark) errors.push(`${id}: missing Remark box.`);
    if (issue.hasStatus && !statuses.includes(issue.status)) errors.push(`${id}: invalid status "${issue.status}".`);
  });

  if (!issues.length) return showResult("No tracked issues found.");
  if (!errors.length) return showResult(`Validation passed.\n${issues.length} issue(s) checked.`);
  showResult(`Validation found ${errors.length} problem(s):\n- ${errors.join("\n- ")}`);
}

async function applyStatus() {
  const selectedStatus = ui.statusSelect.value;
  const { issueId } = await requireIssueId();

  await PowerPoint.run(async (context) => {
    const slide = await getSelectedSlide(context);
    slide.shapes.load("items/id,items/tags/key,items/tags/value");
    await context.sync();

    const statusShape = slide.shapes.items.find(shape => {
      const typeTag = shape.tags.items.find(t => t.key === TAG_BOX_TYPE);
      return typeTag?.value === BOX_STATUS;
    });
    if (!statusShape) throw new Error("No Status box exists on this slide.");

    statusShape.tags.add(TAG_STATUS, selectedStatus);
    statusShape.textFrame.textRange.text = `Status\n${selectedStatus}`;
    applyTrackerBoxStyle(statusShape, BOX_STATUS, selectedStatus);
    await context.sync();
  });

  showResult(`${issueId} status updated to ${selectedStatus}.`);
}

async function countStatusUsage(status) {
  const issues = await readIssues();
  return issues.filter(i => i.status === status).map(i => i.slideIndex);
}

async function replaceStatusEverywhere(oldStatus, newStatus) {
  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/shapes/items/tags/key,items/shapes/items/tags/value");
    await context.sync();

    for (const slide of slides.items) {
      for (const shape of slide.shapes.items) {
        const typeTag = shape.tags.items.find(t => t.key === TAG_BOX_TYPE);
        const statusTag = shape.tags.items.find(t => t.key === TAG_STATUS);
        if (typeTag?.value === BOX_STATUS && statusTag?.value === oldStatus) {
          shape.tags.add(TAG_STATUS, newStatus);
          shape.textFrame.textRange.text = `Status\n${newStatus}`;
          applyTrackerBoxStyle(shape, BOX_STATUS, newStatus);
        }
      }
    }
    await context.sync();
  });
}

async function removeStatus(status) {
  const statuses = await getStatusLibrary();
  if (statuses.length <= 1) throw new Error("At least one status must remain.");

  const usage = await countStatusUsage(status);
  if (usage.length) {
    const replacements = statuses.filter(s => s !== status);
    const replacement = window.prompt(
      `"${status}" is used on ${usage.length} slide(s): ${usage.join(", ")}.\nEnter replacement status:\n${replacements.join(" | ")}`,
      replacements[0]
    );
    if (!replacement) return;
    const matched = replacements.find(s => s.toLowerCase() === replacement.trim().toLowerCase());
    if (!matched) throw new Error("Replacement must be one of the existing statuses.");
    await replaceStatusEverywhere(status, matched);
  }

  await saveStatusLibrary(statuses.filter(s => s !== status));
  await refreshStatusUi();
  showResult(`Status "${status}" removed.`);
}


function stableHash(text) {
  let h = 2166136261;
  const s = String(text || "");
  for (let i = 0; i < s.length; i++) {
    h ^= s.charCodeAt(i);
    h = Math.imul(h, 16777619);
  }
  return (h >>> 0).toString(16).padStart(8, "0");
}

function issueContentHash(issue) {
  return stableHash([
    String(issue.description || "").trim(),
    String(issue.status || "").trim(),
    String(issue.remark || "").trim()
  ].join("\u001f"));
}

function formatDate(value) {
  if (!value) return "";
  const d = new Date(value);
  if (Number.isNaN(d.getTime())) return "";
  return d.toLocaleDateString("en-GB", {
    day: "2-digit",
    month: "short",
    year: "numeric"
  });
}

function statusColor(status) {
  const s = String(status || "").toLowerCase();
  if (s === "open") return "#D64545";
  if (s === "in progress") return "#3A78C2";
  if (s === "pending") return "#D89A2B";
  if (s === "closed") return "#2F8F5B";

  // Deterministic fallback for user-defined statuses.
  const palette = ["#6C5CE7", "#008C95", "#B56A2D", "#8C4A7E", "#526D82"];
  const idx = parseInt(stableHash(s).slice(-2), 16) % palette.length;
  return palette[idx];
}

async function refreshIssueNavigator() {
  if (!ui.issueNavigator) return;
  const issues = await readIssues();
  issues.sort((a, b) => a.issueId.localeCompare(b.issueId, undefined, { numeric: true }));
  ui.issueNavigator.innerHTML = "";
  issues.forEach(issue => {
    const option = document.createElement("option");
    option.value = issue.issueId;
    option.textContent = issue.issueId;
    ui.issueNavigator.appendChild(option);
  });
}

async function goToIssue() {
  const issueId = cleanIssueId(ui.issueNavigator.value);
  if (!issueId) throw new Error("No Issue ID is available.");

  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/id,items/tags/key,items/tags/value");
    await context.sync();

    const target = slides.items.find(slide => {
      const tag = slide.tags.items.find(t => t.key === TAG_ISSUE_ID);
      return cleanIssueId(tag?.value) === issueId;
    });

    if (!target) throw new Error(`Could not find ${issueId}.`);
    context.presentation.setSelectedSlides([target.id]);
    await context.sync();
  });

  showResult(`Opened ${issueId}.`);
}

async function syncIssueTrackingMetadata(issues) {
  const now = new Date().toISOString();

  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/id,items/tags/key,items/tags/value");
    await context.sync();

    const slideById = new Map(slides.items.map(s => [s.id, s]));

    for (const issue of issues) {
      const slide = slideById.get(issue.slideId);
      if (!slide) continue;

      const createdTag = slide.tags.items.find(t => t.key === TAG_CREATED_AT);
      const updatedTag = slide.tags.items.find(t => t.key === TAG_UPDATED_AT);
      const snapshotTag = slide.tags.items.find(t => t.key === TAG_SNAPSHOT_HASH);
      const currentHash = issueContentHash(issue);

      if (!createdTag?.value) {
        slide.tags.add(TAG_CREATED_AT, now);
        issue.createdAt = now;
      } else {
        issue.createdAt = createdTag.value;
      }

      if (!snapshotTag?.value) {
        // First V1.4 refresh establishes the baseline and does not count as an update.
        slide.tags.add(TAG_SNAPSHOT_HASH, currentHash);
        issue.updatedAt = updatedTag?.value || "";
      } else if (snapshotTag.value !== currentHash) {
        slide.tags.add(TAG_SNAPSHOT_HASH, currentHash);
        slide.tags.add(TAG_UPDATED_AT, now);
        issue.updatedAt = now;
      } else {
        issue.updatedAt = updatedTag?.value || "";
      }
    }

    await context.sync();
  });
}

async function deleteExistingSummarySlides() {
  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/id,items/tags/key,items/tags/value");
    await context.sync();

    const toDelete = slides.items.filter(slide => {
      const tag = slide.tags.items.find(t => t.key === TAG_SUMMARY);
      return tag?.value === "TRUE";
    });

    for (const slide of toDelete) slide.delete();
    await context.sync();
  });
}

async function addBlankTaggedSlide(context, type) {
  const slides = context.presentation.slides;
  const count = slides.getCount();

  slides.add();
  await context.sync();

  const slide = slides.getItemAt(count.value);
  slide.load("id,shapes/items/id");
  await context.sync();

  // Remove all shapes inherited from the current PowerPoint layout.
  // This gives every generated Dashboard/Register page a clean blank canvas.
  for (const shape of slide.shapes.items) {
    shape.delete();
  }
  await context.sync();

  slide.tags.add(TAG_SUMMARY, "TRUE");
  slide.tags.add(TAG_SUMMARY_TYPE, type);
  slide.tags.add(TAG_APP, "POWERPOINT_ISSUE_TRACKER");

  await context.sync();
  return slide;
}

function addText(slide, text, left, top, width, height, size, bold, color) {
  const box = slide.shapes.addTextBox(text, { left, top, width, height });
  box.textFrame.wordWrap = true;
  box.textFrame.textRange.font.name = "Aptos";
  box.textFrame.textRange.font.size = size;
  box.textFrame.textRange.font.bold = !!bold;
  if (color) box.textFrame.textRange.font.color = color;
  return box;
}

function addCard(slide, left, top, width, height, label, value, accent) {
  const card = slide.shapes.addGeometricShape(
    PowerPoint.GeometricShapeType.roundRectangle,
    { left, top, width, height }
  );
  card.fill.setSolidColor("#FFFFFF");
  card.lineFormat.color = "#DDE3EA";
  card.lineFormat.weight = 1;

  const strip = slide.shapes.addGeometricShape(
    PowerPoint.GeometricShapeType.rectangle,
    { left, top, width: 6, height }
  );
  strip.fill.setSolidColor(accent);
  strip.lineFormat.color = accent;

  addText(slide, String(value), left + 18, top + 10, width - 24, 28, 21, true, "#17212B");
  addText(slide, label, left + 18, top + 40, width - 24, 20, 10, false, "#5E6A75");
}

function createPieChartBase64(statusCounts, total) {
  const canvas = document.createElement("canvas");
  canvas.width = 760;
  canvas.height = 430;
  const ctx = canvas.getContext("2d");
  if (!ctx) return null;

  ctx.fillStyle = "#FFFFFF";
  ctx.fillRect(0, 0, canvas.width, canvas.height);

  const cx = 235;
  const cy = 215;
  const radius = 145;
  const innerRadius = 88;
  let angle = -Math.PI / 2;

  const entries = Object.entries(statusCounts);
  entries.forEach(([status, count]) => {
    const portion = total ? count / total : 0;
    const next = angle + portion * Math.PI * 2;

    ctx.beginPath();
    ctx.arc(cx, cy, radius, angle, next);
    ctx.arc(cx, cy, innerRadius, next, angle, true);
    ctx.closePath();
    ctx.fillStyle = statusColor(status);
    ctx.fill();

    angle = next;
  });

  ctx.fillStyle = "#17212B";
  ctx.textAlign = "center";
  ctx.font = "700 44px Segoe UI";
  ctx.fillText(String(total), cx, cy + 4);
  ctx.font = "20px Segoe UI";
  ctx.fillStyle = "#64707C";
  ctx.fillText("Total Issues", cx, cy + 40);

  ctx.textAlign = "left";
  let ly = 95;
  entries.forEach(([status, count]) => {
    ctx.fillStyle = statusColor(status);
    ctx.fillRect(455, ly - 17, 20, 20);
    ctx.fillStyle = "#17212B";
    ctx.font = "600 22px Segoe UI";
    ctx.fillText(status, 490, ly);
    ctx.fillStyle = "#64707C";
    ctx.font = "20px Segoe UI";
    const pct = total ? Math.round((count / total) * 100) : 0;
    ctx.fillText(`${count}  (${pct}%)`, 490, ly + 28);
    ly += 76;
  });

  return canvas.toDataURL("image/png").split(",")[1];
}

function createTableCellProps(rowCount, columnCount, values) {
  const props = Array.from({ length: rowCount }, () =>
    Array.from({ length: columnCount }, () => ({}))
  );

  for (let c = 0; c < columnCount; c++) {
    props[0][c] = {
      fill: { color: "#18324A" },
      font: { bold: true, color: "#FFFFFF", name: "Aptos", size: 10 }
    };
  }

  for (let r = 1; r < rowCount; r++) {
    const baseFill = r % 2 === 0 ? "#F4F7FA" : "#FFFFFF";
    for (let c = 0; c < columnCount; c++) {
      props[r][c] = {
        fill: { color: baseFill },
        font: { color: "#24313C", name: "Aptos", size: 9.5 }
      };
    }

    // Status column.
    props[r][2] = {
      fill: { color: statusColor(values[r][2]) },
      font: { bold: true, color: "#FFFFFF", name: "Aptos", size: 9.5 }
    };
  }

  return props;
}

async function buildDashboard(slide, issues, statusCounts) {
  const total = issues.length;

  const topBand = slide.shapes.addGeometricShape(
    PowerPoint.GeometricShapeType.rectangle,
    { left: 0, top: 0, width: 960, height: 58 }
  );
  topBand.fill.setSolidColor("#13283A");
  topBand.lineFormat.color = "#13283A";

  addText(slide, "ISSUE MANAGEMENT DASHBOARD", 38, 16, 610, 28, 22, true, "#FFFFFF");
  addText(
    slide,
    `Refreshed ${new Date().toLocaleString("en-GB")}`,
    660, 20, 260, 20, 9.5, false, "#D7E1E9"
  );

  const statusOrder = ["Open", "In Progress", "Pending", "Closed"];
  addCard(slide, 40, 82, 165, 74, "Total Issues", total, "#445B70");

  let cardX = 220;
  statusOrder.forEach(status => {
    addCard(
      slide,
      cardX,
      82,
      165,
      74,
      status,
      statusCounts[status] || 0,
      statusColor(status)
    );
    cardX += 180;
  });

  addText(slide, "Status Distribution", 40, 184, 400, 24, 15, true, "#17212B");
  addText(slide, "Status Overview", 535, 184, 360, 24, 15, true, "#17212B");

  // Power BI-like horizontal bars (stable shape fallback).
  const maxCount = Math.max(1, ...Object.values(statusCounts));
  let y = 225;
  Object.entries(statusCounts).forEach(([status, count]) => {
    addText(slide, status, 535, y, 125, 18, 10, false, "#35424E");
    const bg = slide.shapes.addGeometricShape(
      PowerPoint.GeometricShapeType.rectangle,
      { left: 660, top: y + 2, width: 195, height: 14 }
    );
    bg.fill.setSolidColor("#E8EDF2");
    bg.lineFormat.color = "#E8EDF2";

    const width = 195 * (count / maxCount);
    const bar = slide.shapes.addGeometricShape(
      PowerPoint.GeometricShapeType.rectangle,
      { left: 660, top: y + 2, width: Math.max(4, width), height: 14 }
    );
    bar.fill.setSolidColor(statusColor(status));
    bar.lineFormat.color = statusColor(status);
    addText(slide, String(count), 865, y - 1, 45, 18, 10, true, "#35424E");
    y += 38;
  });

  await slide.context.sync();

  // Try a true donut image generated locally in the task pane.
  const imageBase64 = createPieChartBase64(statusCounts, total);
  if (imageBase64 && typeof slide.shapes.addPicture === "function") {
    try {
      const picture = slide.shapes.addPicture(imageBase64, {
        left: 42,
        top: 214,
        width: 440,
        height: 250
      });
      picture.name = "PIT_Status_Donut";
      await slide.context.sync();
      return;
    } catch (e) {
      console.warn("Donut image API unavailable; using shape fallback.", e);
    }
  }

  // Fallback if addPicture is unavailable.
  let fy = 228;
  Object.entries(statusCounts).forEach(([status, count]) => {
    const sq = slide.shapes.addGeometricShape(
      PowerPoint.GeometricShapeType.rectangle,
      { left: 60, top: fy, width: 16, height: 16 }
    );
    sq.fill.setSolidColor(statusColor(status));
    sq.lineFormat.color = statusColor(status);
    const pct = total ? Math.round((count / total) * 100) : 0;
    addText(slide, `${status}: ${count} (${pct}%)`, 90, fy - 2, 300, 20, 11, false, "#35424E");
    fy += 34;
  });
}

async function buildRegisterSlide(slide, pageIssues, pageNumber, pageCount) {
  const header = slide.shapes.addGeometricShape(
    PowerPoint.GeometricShapeType.rectangle,
    { left: 0, top: 0, width: 960, height: 58 }
  );
  header.fill.setSolidColor("#13283A");
  header.lineFormat.color = "#13283A";

  addText(slide, "ISSUE REGISTER", 38, 16, 420, 28, 22, true, "#FFFFFF");
  addText(slide, `Page ${pageNumber} of ${pageCount}`, 765, 20, 150, 20, 10, false, "#D7E1E9");

  const values = [
    ["ID", "Description", "Status", "Remark"],
    ...pageIssues.map(issue => [
      "",
      issue.description || "",
      issue.status || "",
      issue.remark || ""
    ])
  ];

  const rowCount = values.length;
  const colCount = 4;
  const specificCellProperties = createTableCellProps(rowCount, colCount, values);

  const tableTop = 92;
  const tableLeft = 40;
  const tableWidth = 880;
  const rowHeight = 28;

  slide.shapes.addTable(rowCount, colCount, {
    values,
    left: tableLeft,
    top: tableTop,
    width: tableWidth,
    height: rowCount * rowHeight,
    uniformCellProperties: {
      font: { name: "Aptos", size: 9.5, color: "#24313C" }
    },
    specificCellProperties
  });

  // Overlay Issue IDs as the stable user-facing identity.
  // No ppaction:// hyperlink is assigned because PowerPoint displays a security warning.
  // Use Issue Navigator > Go for direct no-warning navigation by Issue ID.
  pageIssues.forEach((issue, idx) => {
    const y = tableTop + rowHeight * (idx + 1) + 4;
    const idShape = addText(
      slide,
      issue.issueId,
      tableLeft + 7,
      y,
      115,
      18,
      9.5,
      true,
      "#1769AA"
    );
    idShape.tags.add(TAG_LINK_ID, issue.issueId);
  });

}

async function updateIssueFooters(issues) {
  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/id,items/shapes/items/id,items/shapes/items/tags/key,items/shapes/items/tags/value");
    await context.sync();

    const slideById = new Map(slides.items.map(s => [s.id, s]));

    for (const issue of issues) {
      const slide = slideById.get(issue.slideId);
      if (!slide) continue;

      const existing = slide.shapes.items.find(shape =>
        shape.tags.items.some(t => t.key === TAG_FOOTER && t.value === "TRUE")
      );

      const created = formatDate(issue.createdAt);
      const updated = formatDate(issue.updatedAt);
      let footerText = `${issue.issueId}    Created: ${created}`;
      if (updated) footerText += `    Updated: ${updated}`;

      if (existing) {
        existing.textFrame.textRange.text = footerText;
      } else {
        const footer = slide.shapes.addTextBox(footerText, {
          left: 40, top: 522, width: 880, height: 14
        });
        footer.name = `PIT_FOOTER_${issue.issueId}`;
        footer.textFrame.textRange.font.name = "Aptos";
        footer.textFrame.textRange.font.size = 8;
        footer.textFrame.textRange.font.color = "#65727E";
        footer.tags.add(TAG_APP, "POWERPOINT_ISSUE_TRACKER");
        footer.tags.add(TAG_FOOTER, "TRUE");
      }
    }

    await context.sync();
  });
}

async function generateSummary() {
  if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.8")) {
    throw new Error(
      "Professional dashboard/register generation requires PowerPointApi 1.8. " +
      "Your current PowerPoint can still use issue boxes, but this summary feature needs a newer PowerPoint build."
    );
  }

  let issues = await readIssues();
  if (!issues.length) throw new Error("No tracked issues found.");

  issues.sort((a, b) => a.issueId.localeCompare(b.issueId, undefined, { numeric: true }));

  await syncIssueTrackingMetadata(issues);
  await deleteExistingSummarySlides();

  const rowsPerRegisterSlide = 12;
  const pageCount = Math.max(1, Math.ceil(issues.length / rowsPerRegisterSlide));

  let dashboardId = "";
  const registerIds = [];

  // Create blank generated slides first.
  await PowerPoint.run(async (context) => {
    const dashboard = await addBlankTaggedSlide(context, "DASHBOARD");
    dashboardId = dashboard.id;

    for (let i = 0; i < pageCount; i++) {
      const register = await addBlankTaggedSlide(context, "REGISTER");
      registerIds.push(register.id);
    }
  });

  // Keep generated slides at the front. Move from last to first.
  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    const generatedIds = [dashboardId, ...registerIds];

    for (let i = generatedIds.length - 1; i >= 0; i--) {
      const slide = slides.getItem(generatedIds[i]);
      slide.moveTo(0);
    }
    await context.sync();
  });


  const statusCounts = {};
  issues.forEach(issue => {
    const status = cleanStatus(issue.status) || "Missing";
    statusCounts[status] = (statusCounts[status] || 0) + 1;
  });

  // Populate dashboard and registers.
  await PowerPoint.run(async (context) => {
    const dashboard = context.presentation.slides.getItem(dashboardId);
    await buildDashboard(dashboard, issues, statusCounts);

    for (let p = 0; p < pageCount; p++) {
      const register = context.presentation.slides.getItem(registerIds[p]);
      const pageIssues = issues.slice(
        p * rowsPerRegisterSlide,
        (p + 1) * rowsPerRegisterSlide
      );
      await buildRegisterSlide(register, pageIssues, p + 1, pageCount);
    }

    await context.sync();
  });

  await updateIssueFooters(issues);
  await refreshIssueNavigator();

  showResult(
    `Summary refreshed: 1 dashboard + ${pageCount} register page(s) for ${issues.length} issue(s).`
  );
}

function showResult(message) { ui.result.textContent = message; }
function showError(error) {
  console.error(error);
  ui.result.textContent = `Error: ${error?.message || error}`;
}
