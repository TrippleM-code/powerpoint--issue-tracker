
const TAG_APP = "PIT_APP";
const TAG_ISSUE_ID = "PIT_ISSUE_ID";
const TAG_BOX_TYPE = "PIT_BOX_TYPE";
const TAG_STATUS = "PIT_STATUS";
const TAG_SUMMARY = "PIT_SUMMARY";
const TAG_STATUS_LIBRARY = "PIT_STATUS_LIBRARY";

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
   "repairCurrentSlide","validate","generateSummary","result"]
  .forEach(id => ui[id] = document.getElementById(id));

  ui.setIssueId.addEventListener("click", () => setIssueId().catch(showError));
  ui.addDescription.addEventListener("click", () => addBox(BOX_DESCRIPTION).catch(showError));
  ui.addStatus.addEventListener("click", () => addBox(BOX_STATUS).catch(showError));
  ui.addRemark.addEventListener("click", () => addBox(BOX_REMARK).catch(showError));
  ui.applyStatus.addEventListener("click", () => applyStatus().catch(showError));
  ui.addStatusValue.addEventListener("click", () => addStatusValue().catch(showError));
  ui.repairCurrentSlide.addEventListener("click", () => repairCurrentSlide().catch(showError));
  ui.validate.addEventListener("click", () => validatePresentation().catch(showError));
  ui.generateSummary.addEventListener("click", () => generateSummary().catch(showError));
}

async function initialize() {
  await ensureStatusLibrary();
  await loadCurrentSlideIssue();
  await refreshStatusUi();
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
  if (type === BOX_DESCRIPTION) return { left:36, top:360, width:360, height:100 };
  if (type === BOX_STATUS) return { left:420, top:360, width:160, height:50 };
  return { left:36, top:470, width:544, height:90 };
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
    box.fill.setSolidColor("#FFFFFF");
    box.lineFormat.color = "#808080";
    box.lineFormat.weight = 1;
    box.textFrame.wordWrap = true;
    box.textFrame.autoSizeSetting = PowerPoint.ShapeAutoSize.autoSizeNone;
    box.textFrame.textRange.text = initialText(type, status);
    box.textFrame.textRange.font.name = "Aptos";
    box.textFrame.textRange.font.size = 12;
    box.tags.add(TAG_APP, "POWERPOINT_ISSUE_TRACKER");
    box.tags.add(TAG_ISSUE_ID, issueId);
    box.tags.add(TAG_BOX_TYPE, type);
    if (type === BOX_STATUS) box.tags.add(TAG_STATUS, status);

    await context.sync();
  });

  showResult(`${type} box added.`);
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

async function deleteExistingSummarySlides() {
  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/id,items/tags/key,items/tags/value");
    await context.sync();

    const toDelete = [];
    for (const slide of slides.items) {
      const tag = slide.tags.items.find(t => t.key === TAG_SUMMARY);
      if (tag?.value === "TRUE") toDelete.push(slide);
    }

    for (const slide of toDelete) slide.delete();
    await context.sync();
  });
}

async function generateSummary() {
  const issues = await readIssues();
  if (!issues.length) throw new Error("No tracked issues found.");

  await deleteExistingSummarySlides();

  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;

    slides.load("items/id");
    await context.sync();
    const newSlideIndex = slides.items.length;

    slides.add();
    await context.sync();

    slides.load("items/id");
    await context.sync();

    const summarySlide = slides.items[newSlideIndex];
    if (!summarySlide) {
      throw new Error("PowerPoint created the summary slide, but the add-in could not retrieve it.");
    }

    summarySlide.tags.add(TAG_SUMMARY, "TRUE");
    summarySlide.tags.add(TAG_APP, "POWERPOINT_ISSUE_TRACKER");

    const title = summarySlide.shapes.addTextBox("Issue Summary", {
      left: 30, top: 20, width: 650, height: 36
    });
    title.textFrame.textRange.font.size = 24;
    title.textFrame.textRange.font.bold = true;

    const counts = {};
    issues.forEach(i => {
      const key = cleanStatus(i.status) || "Missing";
      counts[key] = (counts[key] || 0) + 1;
    });

    const countText = Object.entries(counts)
      .map(([k, v]) => `${k}: ${v}`)
      .join("   |   ");

    const overview = summarySlide.shapes.addTextBox(
      `Total: ${issues.length}   |   ${countText}`,
      { left: 30, top: 60, width: 650, height: 30 }
    );
    overview.textFrame.textRange.font.size = 12;

    const header = "ID\tDescription\tStatus\tRemark\tLocation";
    const rows = issues.map(i => {
      const desc = (i.description || "").replace(/\s+/g, " ").slice(0, 55);
      const remark = (i.remark || "").replace(/\s+/g, " ").slice(0, 45);
      return `${i.issueId}\t${desc}\t${i.status || "Missing"}\t${remark}\tSlide ${i.slideIndex}`;
    });

    const table = summarySlide.shapes.addTextBox(
      [header, ...rows].join("\n"),
      { left: 30, top: 105, width: 660, height: 380 }
    );
    table.textFrame.wordWrap = true;
    table.textFrame.autoSizeSetting = PowerPoint.ShapeAutoSize.autoSizeTextToFitShape;
    table.textFrame.textRange.font.name = "Aptos";
    table.textFrame.textRange.font.size = 10;

    await context.sync();
  });

  showResult(`Summary generated for ${issues.length} issue(s).`);
}

function showResult(message) { ui.result.textContent = message; }
function showError(error) {
  console.error(error);
  ui.result.textContent = `Error: ${error?.message || error}`;
}
