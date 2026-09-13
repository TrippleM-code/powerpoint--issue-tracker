const TAG_APP = "PIT_APP";
const TAG_ISSUE_ID = "PIT_ISSUE_ID";
const TAG_CREATED_AT = "PIT_CREATED_AT";
const TAG_UPDATED_AT = "PIT_UPDATED_AT";
const TAG_DESCRIPTION_V2 = "PIT_DESCRIPTION_V2";
const TAG_ACTIONS_V2 = "PIT_ACTIONS_V2";
const TAG_PARTY_LIBRARY = "PIT_PARTY_LIBRARY";
const TAG_STATUS_LIBRARY = "PIT_STATUS_LIBRARY";
const TAG_MANAGED = "PIT_MANAGED";
const TAG_MANAGED_ROLE = "PIT_MANAGED_ROLE";
const TAG_SUMMARY = "PIT_SUMMARY";
const TAG_SUMMARY_TYPE = "PIT_SUMMARY_TYPE";

// Legacy tags retained for migration/compatibility.
const TAG_BOX_TYPE = "PIT_BOX_TYPE";
const BOX_DESCRIPTION = "DESCRIPTION";

const DEFAULT_STATUSES = ["Open", "In Progress", "Pending", "Closed"];
const DEFAULT_PARTIES = ["Architect", "C&S", "MEP", "Main Contractor", "Subcontractor"];

const ui = {};
let editingActionId = null;
let toastTimer = null;

Office.onReady((info) => {
  if (info.host !== Office.HostType.PowerPoint) return;
  bindUi();
  initialize().catch(showError);
});

function bindUi() {
  [
    "issueId","setIssueId","issueDescription","saveIssue","createIssueSheet",
    "currentIssueHeading","overallStatusChip","issueActionCount","issueCreatedDate","issueUpdatedDate",
    "issueNavigator","goToIssue",
    "actionParty","actionText","actionStatus","actionRemark","saveAction","clearActionForm",
    "actionEditorTitle","actionList","actionsCountBadge","actionsEmpty",
    "generateSummary","previewIssues","previewActions","previewOpen","previewClosed",
    "partyList","newParty","addParty","statusList","newStatus","addStatus","validatePresentation","toast"
  ].forEach(id => ui[id] = document.getElementById(id));

  document.querySelectorAll(".tab").forEach(tab => {
    tab.addEventListener("click", () => activateTab(tab.dataset.tab));
  });

  ui.setIssueId.addEventListener("click", () => setIssueId().catch(showError));
  ui.saveIssue.addEventListener("click", () => saveIssueData().catch(showError));
  ui.createIssueSheet.addEventListener("click", () => createOrRefreshIssueSheet().catch(showError));
  ui.goToIssue.addEventListener("click", () => goToIssue().catch(showError));

  ui.saveAction.addEventListener("click", () => saveAction().catch(showError));
  ui.clearActionForm.addEventListener("click", clearActionForm);

  ui.generateSummary.addEventListener("click", () => generateSummary().catch(showError));
  ui.addParty.addEventListener("click", () => addParty().catch(showError));
  ui.addStatus.addEventListener("click", () => addStatusValue().catch(showError));
  ui.validatePresentation.addEventListener("click", () => validatePresentation().catch(showError));
}

async function initialize() {
  await ensureLibraries();
  await refreshLibrariesUi();
  await refreshIssueNavigator();
  await loadCurrentIssueSafe();
  await refreshSummaryPreview();
}

function activateTab(name) {
  document.querySelectorAll(".tab").forEach(tab => tab.classList.toggle("active", tab.dataset.tab === name));
  document.querySelectorAll(".tab-panel").forEach(panel => panel.classList.toggle("active", panel.id === `panel-${name}`));
}

function cleanIssueId(value) {
  return String(value || "").trim().toUpperCase();
}

function cleanText(value) {
  return String(value || "").trim();
}

function formatDate(value) {
  if (!value) return "—";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "—";
  return date.toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" });
}

function nowIso() {
  return new Date().toISOString();
}

function statusClass(status) {
  const s = cleanText(status).toLowerCase();
  if (s === "open") return "open";
  if (s === "in progress") return "in-progress";
  if (s === "pending") return "pending";
  if (s === "closed") return "closed";
  return "neutral";
}

function statusColor(status) {
  const s = cleanText(status).toLowerCase();
  if (s === "open") return { fill: "#FDE8E8", line: "#D64545", text: "#8A1F1F" };
  if (s === "in progress") return { fill: "#E6F0FF", line: "#3A78C2", text: "#174A7E" };
  if (s === "pending") return { fill: "#FFF4D6", line: "#D89A2B", text: "#7A560F" };
  if (s === "closed") return { fill: "#E7F6EC", line: "#2F8F5B", text: "#1E603D" };
  return { fill: "#EEF2F5", line: "#7A8590", text: "#35424E" };
}

function computeOverallStatus(actions) {
  if (!actions.length) return "No actions";
  const statuses = actions.map(a => cleanText(a.status).toLowerCase());
  if (statuses.every(s => s === "closed")) return "Closed";
  if (statuses.includes("open")) return "Open";
  if (statuses.includes("in progress")) return "In Progress";
  if (statuses.includes("pending")) return "Pending";
  return actions[0].status || "Open";
}

async function getSelectedSlide(context) {
  const selected = context.presentation.getSelectedSlides();
  selected.load("items/id");
  await context.sync();
  if (!selected.items.length) throw new Error("Select a PowerPoint slide first.");
  return selected.items[0];
}

async function getTagValue(tags, key, context) {
  const tag = tags.getItemOrNullObject(key);
  tag.load("value,isNullObject");
  await context.sync();
  return tag.isNullObject ? null : tag.value;
}

async function ensureLibraries() {
  await PowerPoint.run(async context => {
    const statusValue = await getTagValue(context.presentation.tags, TAG_STATUS_LIBRARY, context);
    if (!statusValue) context.presentation.tags.add(TAG_STATUS_LIBRARY, JSON.stringify(DEFAULT_STATUSES));

    const partyValue = await getTagValue(context.presentation.tags, TAG_PARTY_LIBRARY, context);
    if (!partyValue) context.presentation.tags.add(TAG_PARTY_LIBRARY, JSON.stringify(DEFAULT_PARTIES));

    await context.sync();
  });
}

async function readJsonPresentationTag(tagName, fallback) {
  return PowerPoint.run(async context => {
    const value = await getTagValue(context.presentation.tags, tagName, context);
    try {
      const parsed = JSON.parse(value || "[]");
      return Array.isArray(parsed) && parsed.length ? parsed : fallback.slice();
    } catch {
      return fallback.slice();
    }
  });
}

async function writeJsonPresentationTag(tagName, value) {
  await PowerPoint.run(async context => {
    context.presentation.tags.add(tagName, JSON.stringify(value));
    await context.sync();
  });
}

async function getStatuses() {
  return readJsonPresentationTag(TAG_STATUS_LIBRARY, DEFAULT_STATUSES);
}

async function getParties() {
  return readJsonPresentationTag(TAG_PARTY_LIBRARY, DEFAULT_PARTIES);
}

async function refreshLibrariesUi() {
  const [parties, statuses] = await Promise.all([getParties(), getStatuses()]);

  populateSelect(ui.actionParty, parties);
  populateSelect(ui.actionStatus, statuses);

  renderLibraryList(ui.partyList, parties, value => removeParty(value));
  renderLibraryList(ui.statusList, statuses, value => removeStatus(value));
}

function populateSelect(select, values) {
  const previous = select.value;
  select.innerHTML = "";
  values.forEach(value => {
    const option = document.createElement("option");
    option.value = value;
    option.textContent = value;
    select.appendChild(option);
  });
  if (values.includes(previous)) select.value = previous;
}

function renderLibraryList(container, values, removeHandler) {
  container.innerHTML = "";
  values.forEach(value => {
    const row = document.createElement("div");
    row.className = "library-item";
    const span = document.createElement("span");
    span.textContent = value;
    const button = document.createElement("button");
    button.textContent = "Remove";
    button.addEventListener("click", () => removeHandler(value).catch(showError));
    row.append(span, button);
    container.appendChild(row);
  });
}

async function loadCurrentIssueSafe() {
  try {
    await loadCurrentIssue();
  } catch (error) {
    ui.issueId.value = "";
    ui.issueDescription.value = "";
    updateCurrentIssueUi(null);
    renderActions([]);
  }
}

async function loadCurrentIssue() {
  const issue = await readSelectedIssue();
  ui.issueId.value = issue.issueId || "";
  ui.issueDescription.value = issue.description || "";
  updateCurrentIssueUi(issue);
  renderActions(issue.actions);
}

async function readSelectedIssue() {
  return PowerPoint.run(async context => {
    const slide = await getSelectedSlide(context);
    const issueId = await getTagValue(slide.tags, TAG_ISSUE_ID, context);
    const descriptionTag = await getTagValue(slide.tags, TAG_DESCRIPTION_V2, context);
    const actionsTag = await getTagValue(slide.tags, TAG_ACTIONS_V2, context);
    const createdAt = await getTagValue(slide.tags, TAG_CREATED_AT, context);
    const updatedAt = await getTagValue(slide.tags, TAG_UPDATED_AT, context);

    let description = descriptionTag || "";
    let actions = [];
    try { actions = JSON.parse(actionsTag || "[]"); } catch { actions = []; }
    if (!Array.isArray(actions)) actions = [];

    // Migration helper: if V2 description is empty, try the V1 Description box.
    if (!description) {
      slide.shapes.load("items/id,items/tags/key,items/tags/value");
      await context.sync();
      const legacy = slide.shapes.items.find(shape => {
        const typeTag = shape.tags.items.find(t => t.key === TAG_BOX_TYPE);
        return typeTag?.value === BOX_DESCRIPTION;
      });
      if (legacy) {
        legacy.textFrame.textRange.load("text");
        await context.sync();
        const text = legacy.textFrame.textRange.text || "";
        description = text.split(/\r?\n/).slice(1).join("\n").trim();
      }
    }

    return {
      slideId: slide.id,
      issueId: issueId || "",
      description,
      actions,
      createdAt: createdAt || "",
      updatedAt: updatedAt || ""
    };
  });
}

function updateCurrentIssueUi(issue) {
  const issueId = issue?.issueId || "";
  const actions = issue?.actions || [];
  const overall = computeOverallStatus(actions);

  ui.currentIssueHeading.textContent = issueId || "No issue selected";
  ui.issueActionCount.textContent = String(actions.length);
  ui.issueCreatedDate.textContent = formatDate(issue?.createdAt);
  ui.issueUpdatedDate.textContent = formatDate(issue?.updatedAt);

  ui.overallStatusChip.textContent = overall;
  ui.overallStatusChip.className = `chip ${statusClass(overall)}`;
}

async function setIssueId() {
  const issueId = cleanIssueId(ui.issueId.value);
  if (!issueId) throw new Error("Enter an Issue ID, for example MEP-001.");

  await PowerPoint.run(async context => {
    const slide = await getSelectedSlide(context);
    const slides = context.presentation.slides;
    slides.load("items/id,items/tags/key,items/tags/value");
    await context.sync();

    for (const candidate of slides.items) {
      if (candidate.id === slide.id) continue;
      const tag = candidate.tags.items.find(t => t.key === TAG_ISSUE_ID);
      if (cleanIssueId(tag?.value) === issueId) {
        throw new Error(`${issueId} already exists on another slide.`);
      }
    }

    slide.tags.add(TAG_APP, "ISSUEFLOW_V2");
    slide.tags.add(TAG_ISSUE_ID, issueId);

    const created = await getTagValue(slide.tags, TAG_CREATED_AT, context);
    if (!created) slide.tags.add(TAG_CREATED_AT, nowIso());
    await context.sync();
  });

  await refreshIssueNavigator();
  await loadCurrentIssue();
  showToast(`Issue ID set to ${issueId}.`);
}

async function saveIssueData({ quiet = false } = {}) {
  const issueId = cleanIssueId(ui.issueId.value);
  if (!issueId) throw new Error("Set an Issue ID first.");
  const description = cleanText(ui.issueDescription.value);

  await PowerPoint.run(async context => {
    const slide = await getSelectedSlide(context);
    const currentDescription = await getTagValue(slide.tags, TAG_DESCRIPTION_V2, context) || "";
    slide.tags.add(TAG_DESCRIPTION_V2, description);
    if (currentDescription !== description) slide.tags.add(TAG_UPDATED_AT, nowIso());
    await context.sync();
  });

  await loadCurrentIssue();
  if (!quiet) showToast("Issue data saved.");
}

async function readActionsFromSelectedSlide() {
  const issue = await readSelectedIssue();
  if (!issue.issueId) throw new Error("Set an Issue ID first.");
  return issue;
}

function nextActionId(actions) {
  let max = 0;
  actions.forEach(action => {
    const match = /^ACT-(\d+)$/i.exec(action.id || "");
    if (match) max = Math.max(max, Number(match[1]));
  });
  return `ACT-${String(max + 1).padStart(3, "0")}`;
}

async function saveAction() {
  const issue = await readActionsFromSelectedSlide();
  const party = cleanText(ui.actionParty.value);
  const actionText = cleanText(ui.actionText.value);
  const status = cleanText(ui.actionStatus.value);
  const remark = cleanText(ui.actionRemark.value);

  if (!party) throw new Error("Choose a responsible party.");
  if (!actionText) throw new Error("Enter the required action.");
  if (!status) throw new Error("Choose a status.");

  const actions = issue.actions.slice();
  const timestamp = nowIso();
  const wasEditing = !!editingActionId;

  if (editingActionId) {
    const index = actions.findIndex(a => a.id === editingActionId);
    if (index < 0) throw new Error("The action being edited no longer exists.");
    actions[index] = {
      ...actions[index], party, action: actionText, status, remark, updatedAt: timestamp
    };
  } else {
    actions.push({
      id: nextActionId(actions),
      party,
      action: actionText,
      status,
      remark,
      createdAt: timestamp,
      updatedAt: ""
    });
  }

  await PowerPoint.run(async context => {
    const slide = await getSelectedSlide(context);
    slide.tags.add(TAG_ACTIONS_V2, JSON.stringify(actions));
    slide.tags.add(TAG_UPDATED_AT, timestamp);
    await context.sync();
  });

  clearActionForm();
  await loadCurrentIssue();
  await refreshSummaryPreview();
  showToast(wasEditing ? "Action updated." : "Action added.");
}

function clearActionForm() {
  editingActionId = null;
  ui.actionEditorTitle.textContent = "Add action";
  ui.saveAction.textContent = "+ Add Action";
  ui.actionText.value = "";
  ui.actionRemark.value = "";
}

function renderActions(actions) {
  ui.actionList.innerHTML = "";
  ui.actionsCountBadge.textContent = String(actions.length);
  ui.actionsEmpty.style.display = actions.length ? "none" : "block";

  actions.forEach(action => {
    const card = document.createElement("article");
    card.className = "action-card";

    const top = document.createElement("div");
    top.className = "action-top";
    const party = document.createElement("span");
    party.className = "action-party";
    party.textContent = action.party;
    const chip = document.createElement("span");
    chip.className = `chip ${statusClass(action.status)}`;
    chip.textContent = action.status;
    top.append(party, chip);

    const actionText = document.createElement("div");
    actionText.className = "action-text";
    actionText.textContent = action.action;

    const remark = document.createElement("div");
    remark.className = "action-remark";
    remark.textContent = action.remark ? `Remark: ${action.remark}` : "No remark";

    const controls = document.createElement("div");
    controls.className = "action-controls";
    const edit = document.createElement("button");
    edit.textContent = "Edit";
    edit.addEventListener("click", () => editAction(action));
    const remove = document.createElement("button");
    remove.textContent = "Remove";
    remove.className = "delete";
    remove.addEventListener("click", () => deleteAction(action.id).catch(showError));
    controls.append(edit, remove);

    card.append(top, actionText, remark, controls);
    ui.actionList.appendChild(card);
  });
}

function editAction(action) {
  editingActionId = action.id;
  ui.actionEditorTitle.textContent = `Edit ${action.id}`;
  ui.actionParty.value = action.party;
  ui.actionText.value = action.action;
  ui.actionStatus.value = action.status;
  ui.actionRemark.value = action.remark || "";
  ui.saveAction.textContent = "Save Changes";
  activateTab("actions");
}

async function deleteAction(actionId) {
  const issue = await readActionsFromSelectedSlide();
  const action = issue.actions.find(a => a.id === actionId);
  if (!action) return;
  if (!window.confirm(`Remove ${actionId} (${action.party})?`)) return;

  const actions = issue.actions.filter(a => a.id !== actionId);
  await PowerPoint.run(async context => {
    const slide = await getSelectedSlide(context);
    slide.tags.add(TAG_ACTIONS_V2, JSON.stringify(actions));
    slide.tags.add(TAG_UPDATED_AT, nowIso());
    await context.sync();
  });

  await loadCurrentIssue();
  await refreshSummaryPreview();
  showToast("Action removed.");
}

async function addParty() {
  const value = cleanText(ui.newParty.value);
  if (!value) throw new Error("Enter a party name.");
  const parties = await getParties();
  if (parties.some(p => p.toLowerCase() === value.toLowerCase())) throw new Error("That party already exists.");
  parties.push(value);
  await writeJsonPresentationTag(TAG_PARTY_LIBRARY, parties);
  ui.newParty.value = "";
  await refreshLibrariesUi();
  showToast(`Party "${value}" added.`);
}

async function removeParty(value) {
  const parties = await getParties();
  if (parties.length <= 1) throw new Error("At least one party must remain.");
  const allIssues = await readAllIssues();
  const used = allIssues.some(issue => issue.actions.some(action => action.party === value));
  if (used) throw new Error(`"${value}" is already used by existing actions. Reassign those actions before removing it.`);
  await writeJsonPresentationTag(TAG_PARTY_LIBRARY, parties.filter(p => p !== value));
  await refreshLibrariesUi();
  showToast(`Party "${value}" removed.`);
}

async function addStatusValue() {
  const value = cleanText(ui.newStatus.value);
  if (!value) throw new Error("Enter a status name.");
  const statuses = await getStatuses();
  if (statuses.some(s => s.toLowerCase() === value.toLowerCase())) throw new Error("That status already exists.");
  statuses.push(value);
  await writeJsonPresentationTag(TAG_STATUS_LIBRARY, statuses);
  ui.newStatus.value = "";
  await refreshLibrariesUi();
  showToast(`Status "${value}" added.`);
}

async function removeStatus(value) {
  const statuses = await getStatuses();
  if (statuses.length <= 1) throw new Error("At least one status must remain.");
  const allIssues = await readAllIssues();
  const used = allIssues.some(issue => issue.actions.some(action => action.status === value));
  if (used) throw new Error(`"${value}" is already used by existing actions. Change those actions before removing it.`);
  await writeJsonPresentationTag(TAG_STATUS_LIBRARY, statuses.filter(s => s !== value));
  await refreshLibrariesUi();
  showToast(`Status "${value}" removed.`);
}

function markManaged(shape, role) {
  shape.tags.add(TAG_APP, "ISSUEFLOW_V2");
  shape.tags.add(TAG_MANAGED, "TRUE");
  shape.tags.add(TAG_MANAGED_ROLE, role);
  return shape;
}

function addText(slide, text, left, top, width, height, options = {}) {
  const box = slide.shapes.addTextBox(text, { left, top, width, height });
  box.textFrame.wordWrap = true;
  box.textFrame.textRange.font.name = options.font || "Aptos";
  box.textFrame.textRange.font.size = options.size || 10;
  box.textFrame.textRange.font.bold = !!options.bold;
  if (options.color) box.textFrame.textRange.font.color = options.color;
  return box;
}

function addRect(slide, left, top, width, height, fill, line, radius = false) {
  const rect = slide.shapes.addGeometricShape(
    radius ? PowerPoint.GeometricShapeType.roundRectangle : PowerPoint.GeometricShapeType.rectangle,
    { left, top, width, height }
  );
  rect.fill.setSolidColor(fill);
  rect.lineFormat.color = line;
  rect.lineFormat.weight = 1;
  return rect;
}

async function createOrRefreshIssueSheet() {
  await saveIssueData({ quiet: true });
  const issue = await readSelectedIssue();
  if (!issue.issueId) throw new Error("Set an Issue ID first.");

  await PowerPoint.run(async context => {
    const slide = await getSelectedSlide(context);
    slide.shapes.load("items/id,items/tags/key,items/tags/value");
    await context.sync();

    // Delete only IssueFlow-managed shapes. Manual images/annotations remain untouched.
    for (const shape of slide.shapes.items) {
      const managed = shape.tags.items.find(t => t.key === TAG_MANAGED && t.value === "TRUE");
      if (managed) shape.delete();
    }
    await context.sync();

    const navy = "#13283A";
    const blue = "#1F78B4";
    const border = "#C9D5DE";
    const lightBlue = "#EDF6FC";

    // Header
    markManaged(addRect(slide, 0, 0, 960, 62, navy, navy), "HEADER_BG");
    markManaged(addText(slide, "DAILY PAINKILLER", 32, 16, 155, 22, { size: 11, bold: true, color: "#DDEAF2" }), "HEADER_BRAND");
    markManaged(addText(slide, issue.issueId, 195, 13, 300, 28, { size: 20, bold: true, color: "#FFFFFF" }), "HEADER_ID");
    markManaged(addText(slide, `Created ${formatDate(issue.createdAt)}`, 700, 13, 220, 18, { size: 8.5, color: "#DDEAF2" }), "HEADER_CREATED");
    if (issue.updatedAt) {
      markManaged(addText(slide, `Updated ${formatDate(issue.updatedAt)}`, 700, 31, 220, 18, { size: 8.5, color: "#DDEAF2" }), "HEADER_UPDATED");
    }

    // Section headings
    markManaged(addRect(slide, 28, 84, 246, 28, lightBlue, blue), "DESCRIPTION_HEAD_BG");
    markManaged(addText(slide, "ISSUE DESCRIPTION", 40, 90, 220, 18, { size: 9, bold: true, color: "#174A7E" }), "DESCRIPTION_HEAD");
    markManaged(addRect(slide, 286, 84, 322, 28, "#F4F6F8", border), "REFERENCE_HEAD_BG");
    markManaged(addText(slide, "REFERENCE / IMAGES  ·  MANUAL AREA", 298, 90, 285, 18, { size: 9, bold: true, color: "#4E5B66" }), "REFERENCE_HEAD");
    markManaged(addRect(slide, 620, 84, 312, 28, "#F4F6F8", border), "ACTIONS_HEAD_BG");
    markManaged(addText(slide, "ACTIONS", 632, 90, 180, 18, { size: 9, bold: true, color: "#4E5B66" }), "ACTIONS_HEAD");
    markManaged(addText(slide, `Overall: ${computeOverallStatus(issue.actions)}`, 815, 90, 100, 18, { size: 8.5, bold: true, color: statusColor(computeOverallStatus(issue.actions)).text }), "OVERALL_STATUS");

    // Description area
    markManaged(addRect(slide, 28, 112, 246, 370, "#FFFFFF", border, true), "DESCRIPTION_BG");
    markManaged(addText(slide, issue.description || "No description yet.", 42, 130, 218, 330, { size: 11, color: "#26343F" }), "DESCRIPTION_TEXT");

    // Reference area: border only, with no covering fill. Existing manual images remain visible.
    markManaged(addRect(slide, 286, 112, 322, 1, border, border), "REFERENCE_BORDER_TOP");
    markManaged(addRect(slide, 286, 481, 322, 1, border, border), "REFERENCE_BORDER_BOTTOM");
    markManaged(addRect(slide, 286, 112, 1, 370, border, border), "REFERENCE_BORDER_LEFT");
    markManaged(addRect(slide, 607, 112, 1, 370, border, border), "REFERENCE_BORDER_RIGHT");
    markManaged(addText(slide, "Paste / crop / annotate images freely in this area", 320, 442, 255, 20, { size: 8.5, color: "#94A0AA" }), "REFERENCE_HINT");

    // Actions area
    markManaged(addRect(slide, 620, 112, 312, 370, "#FFFFFF", border, true), "ACTIONS_BG");

    const actions = issue.actions;
    if (!actions.length) {
      markManaged(addText(slide, "No actions yet\nUse IssueFlow > Actions > + Add Action", 655, 230, 245, 60, { size: 11, color: "#7B8791" }), "ACTIONS_EMPTY");
    } else {
      const maxRows = 6;
      const visible = actions.slice(0, maxRows);
      const rowHeight = Math.floor(334 / Math.max(visible.length, 1));
      visible.forEach((action, index) => {
        const top = 124 + index * rowHeight;
        const h = Math.max(48, rowHeight - 6);
        const sc = statusColor(action.status);
        markManaged(addRect(slide, 632, top, 288, h, "#FAFCFD", "#DFE6EB", true), `ACTION_${action.id}_BG`);
        markManaged(addText(slide, action.party, 644, top + 8, 84, 17, { size: 9, bold: true, color: "#26343F" }), `ACTION_${action.id}_PARTY`);
        markManaged(addText(slide, action.action, 735, top + 7, 118, Math.max(20, h - 14), { size: 8.7, color: "#26343F" }), `ACTION_${action.id}_TEXT`);
        const chip = markManaged(addRect(slide, 858, top + 8, 52, 18, sc.fill, sc.line, true), `ACTION_${action.id}_STATUS_BG`);
        markManaged(addText(slide, action.status, 861, top + 11, 46, 12, { size: 7.3, bold: true, color: sc.text }), `ACTION_${action.id}_STATUS`);
        if (action.remark) {
          markManaged(addText(slide, action.remark, 644, top + h - 18, 260, 12, { size: 7.5, color: "#75818B" }), `ACTION_${action.id}_REMARK`);
        }
      });
      if (actions.length > maxRows) {
        markManaged(addText(slide, `+ ${actions.length - maxRows} more action(s) — see Action Register`, 650, 463, 252, 15, { size: 8, color: blue }), "ACTIONS_MORE");
      }
    }

    // Small footer credit.
    markManaged(addText(slide, "Brought to you by Daily Painkiller + Codex", 32, 510, 300, 12, { size: 7.2, color: "#89959E" }), "FOOTER_BRAND");
    await context.sync();
  });

  showToast("Issue sheet refreshed. Manual reference images were preserved.");
}

async function readAllIssues() {
  return PowerPoint.run(async context => {
    const slides = context.presentation.slides;
    slides.load("items/id,items/tags/key,items/tags/value");
    await context.sync();

    const issues = [];
    for (const slide of slides.items) {
      const issueId = slide.tags.items.find(t => t.key === TAG_ISSUE_ID)?.value;
      if (!issueId) continue;
      const description = slide.tags.items.find(t => t.key === TAG_DESCRIPTION_V2)?.value || "";
      const actionsRaw = slide.tags.items.find(t => t.key === TAG_ACTIONS_V2)?.value || "[]";
      const createdAt = slide.tags.items.find(t => t.key === TAG_CREATED_AT)?.value || "";
      const updatedAt = slide.tags.items.find(t => t.key === TAG_UPDATED_AT)?.value || "";
      let actions = [];
      try { actions = JSON.parse(actionsRaw); } catch { actions = []; }
      if (!Array.isArray(actions)) actions = [];
      issues.push({ slideId: slide.id, issueId, description, actions, createdAt, updatedAt });
    }
    issues.sort((a, b) => a.issueId.localeCompare(b.issueId, undefined, { numeric: true }));
    return issues;
  });
}

async function refreshIssueNavigator() {
  const issues = await readAllIssues();
  const previous = ui.issueNavigator.value;
  ui.issueNavigator.innerHTML = "";
  issues.forEach(issue => {
    const option = document.createElement("option");
    option.value = issue.issueId;
    option.textContent = issue.issueId;
    ui.issueNavigator.appendChild(option);
  });
  if (issues.some(i => i.issueId === previous)) ui.issueNavigator.value = previous;
}

async function goToIssue() {
  const issueId = cleanIssueId(ui.issueNavigator.value);
  if (!issueId) throw new Error("No issue is available to navigate to.");

  await PowerPoint.run(async context => {
    const slides = context.presentation.slides;
    slides.load("items/id,items/tags/key,items/tags/value");
    await context.sync();
    const target = slides.items.find(slide => cleanIssueId(slide.tags.items.find(t => t.key === TAG_ISSUE_ID)?.value) === issueId);
    if (!target) throw new Error(`Could not find ${issueId}.`);
    context.presentation.setSelectedSlides([target.id]);
    await context.sync();
  });

  await loadCurrentIssue();
  showToast(`Opened ${issueId}.`);
}

async function deleteExistingSummarySlides() {
  await PowerPoint.run(async context => {
    const slides = context.presentation.slides;
    slides.load("items/id,items/tags/key,items/tags/value");
    await context.sync();
    const generated = slides.items.filter(slide => slide.tags.items.find(t => t.key === TAG_SUMMARY)?.value === "TRUE");
    generated.forEach(slide => slide.delete());
    await context.sync();
  });
}

async function addCleanGeneratedSlide(context, type) {
  const slides = context.presentation.slides;
  const count = slides.getCount();
  slides.add();
  await context.sync();
  const slide = slides.getItemAt(count.value);
  slide.load("id,shapes/items/id");
  await context.sync();
  slide.shapes.items.forEach(shape => shape.delete());
  await context.sync();
  slide.tags.add(TAG_SUMMARY, "TRUE");
  slide.tags.add(TAG_SUMMARY_TYPE, type);
  slide.tags.add(TAG_APP, "ISSUEFLOW_V2");
  await context.sync();
  return slide;
}

function buildDashboard(slide, issues) {
  const allActions = issues.flatMap(issue => issue.actions);
  const statusCounts = {};
  allActions.forEach(action => statusCounts[action.status] = (statusCounts[action.status] || 0) + 1);
  const openIssues = issues.filter(issue => computeOverallStatus(issue.actions) !== "Closed").length;
  const closedIssues = issues.length - openIssues;

  addRect(slide, 0, 0, 960, 62, "#13283A", "#13283A");
  addText(slide, "ISSUEFLOW DASHBOARD", 36, 17, 460, 26, { size: 22, bold: true, color: "#FFFFFF" });
  addText(slide, `Refreshed ${new Date().toLocaleString("en-GB")}`, 690, 20, 230, 17, { size: 8.5, color: "#DCE7EF" });

  const cards = [
    ["Issues", issues.length, "#445B70"],
    ["Open issues", openIssues, "#D64545"],
    ["Closed issues", closedIssues, "#2F8F5B"],
    ["Actions", allActions.length, "#3A78C2"]
  ];
  cards.forEach((card, i) => {
    const left = 40 + i * 220;
    addRect(slide, left, 88, 195, 78, "#FFFFFF", "#DCE4EA", true);
    addRect(slide, left, 88, 6, 78, card[2], card[2]);
    addText(slide, String(card[1]), left + 20, 101, 150, 26, { size: 22, bold: true, color: "#17212B" });
    addText(slide, card[0], left + 20, 135, 150, 16, { size: 9, color: "#68747F" });
  });

  addText(slide, "Action status", 40, 205, 300, 20, { size: 14, bold: true, color: "#17212B" });
  const max = Math.max(1, ...Object.values(statusCounts), 1);
  const statuses = Object.entries(statusCounts);
  statuses.forEach(([status, count], i) => {
    const top = 245 + i * 46;
    const sc = statusColor(status);
    addText(slide, status, 48, top, 130, 18, { size: 10, color: "#3B4853" });
    addRect(slide, 180, top + 2, 500, 15, "#EDF1F4", "#EDF1F4", true);
    addRect(slide, 180, top + 2, Math.max(6, 500 * count / max), 15, sc.line, sc.line, true);
    addText(slide, String(count), 695, top - 1, 50, 18, { size: 10, bold: true, color: "#3B4853" });
  });

  // Actions by party: top 6.
  const partyCounts = {};
  allActions.forEach(action => partyCounts[action.party] = (partyCounts[action.party] || 0) + 1);
  const parties = Object.entries(partyCounts).sort((a,b) => b[1]-a[1]).slice(0,6);
  addText(slide, "Actions by party", 760, 205, 160, 20, { size: 14, bold: true, color: "#17212B" });
  parties.forEach(([party, count], i) => {
    addText(slide, `${party}`, 760, 245 + i*34, 125, 16, { size: 9, color: "#4B5964" });
    addText(slide, `${count}`, 890, 245 + i*34, 30, 16, { size: 9, bold: true, color: "#1F78B4" });
  });

  addText(slide, "Brought to you by Daily Painkiller + Codex", 40, 510, 300, 12, { size: 7.2, color: "#89959E" });
}

function groupIssuesForRegister(issues, maxRows = 9) {
  const pages = [];
  let page = [];
  let used = 0;

  for (const issue of issues) {
    const rows = Math.max(1, issue.actions.length);
    if (rows <= maxRows && used > 0 && used + rows > maxRows) {
      pages.push(page);
      page = [];
      used = 0;
    }

    if (rows <= maxRows) {
      page.push({ issue, start: 0, count: rows });
      used += rows;
    } else {
      let start = 0;
      while (start < rows) {
        if (used === maxRows) {
          pages.push(page); page = []; used = 0;
        }
        const available = maxRows - used;
        const count = Math.min(available, rows - start);
        page.push({ issue, start, count });
        used += count;
        start += count;
        if (used === maxRows) {
          pages.push(page); page = []; used = 0;
        }
      }
    }
  }
  if (page.length) pages.push(page);
  return pages.length ? pages : [[]];
}

function buildRegisterPage(slide, groups, pageNo, pageCount) {
  addRect(slide, 0, 0, 960, 62, "#13283A", "#13283A");
  addText(slide, "ACTION REGISTER", 36, 17, 420, 26, { size: 22, bold: true, color: "#FFFFFF" });
  addText(slide, `Page ${pageNo} of ${pageCount}`, 800, 20, 120, 17, { size: 9, color: "#DCE7EF" });

  const left = 28;
  const top = 92;
  const rowH = 42;
  const widths = [100, 235, 120, 250, 110, 85];
  const headers = ["Issue ID", "Description", "Party", "Action", "Remark", "Status"];
  let x = left;
  headers.forEach((h, i) => {
    addRect(slide, x, top, widths[i], 28, "#18324A", "#18324A");
    addText(slide, h, x + 7, top + 8, widths[i]-12, 12, { size: 8.3, bold: true, color: "#FFFFFF" });
    x += widths[i];
  });

  let rowIndex = 0;
  groups.forEach(group => {
    const issue = group.issue;
    const actions = issue.actions.length ? issue.actions : [{ party: "—", action: "No actions", remark: "", status: "—" }];
    const visibleActions = actions.slice(group.start, group.start + group.count);
    const groupTop = top + 28 + rowIndex * rowH;
    const groupHeight = visibleActions.length * rowH;

    // Merged Issue ID and Description cells.
    addRect(slide, left, groupTop, widths[0], groupHeight, "#F5F8FA", "#D8E1E7");
    addText(slide, issue.issueId, left + 8, groupTop + 10, widths[0]-16, Math.max(18, groupHeight-16), { size: 9.5, bold: true, color: "#1769AA" });

    const descLeft = left + widths[0];
    addRect(slide, descLeft, groupTop, widths[1], groupHeight, "#FFFFFF", "#D8E1E7");
    addText(slide, issue.description || "—", descLeft + 8, groupTop + 8, widths[1]-16, Math.max(20, groupHeight-14), { size: 8.7, color: "#26343F" });

    visibleActions.forEach((action, localIndex) => {
      const y = groupTop + localIndex * rowH;
      const baseX = descLeft + widths[1];
      const values = [action.party || "—", action.action || "—", action.remark || "", action.status || "—"];
      const dataWidths = widths.slice(2);
      let dx = baseX;
      values.forEach((value, i) => {
        const isStatus = i === 3;
        const sc = isStatus ? statusColor(value) : null;
        addRect(slide, dx, y, dataWidths[i], rowH, isStatus ? sc.fill : (rowIndex % 2 ? "#F9FBFC" : "#FFFFFF"), isStatus ? sc.line : "#D8E1E7");
        addText(slide, value, dx + 7, y + 9, dataWidths[i]-14, rowH-14, { size: isStatus ? 8.5 : 8.3, bold: isStatus, color: isStatus ? sc.text : "#26343F" });
        dx += dataWidths[i];
      });
      rowIndex++;
    });
  });

  addText(slide, "Issue ID and Description are grouped; each party action keeps its own status.", 32, 505, 580, 12, { size: 7.2, color: "#89959E" });
  addText(slide, "Daily Painkiller + Codex", 770, 505, 150, 12, { size: 7.2, color: "#89959E" });
}

async function generateSummary() {
  if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.8")) {
    throw new Error("Summary generation requires PowerPointApi 1.8 or later.");
  }

  const issues = await readAllIssues();
  if (!issues.length) throw new Error("No tracked issues found.");

  await deleteExistingSummarySlides();
  const pages = groupIssuesForRegister(issues, 9);
  let dashboardId = "";
  const registerIds = [];

  await PowerPoint.run(async context => {
    const dashboard = await addCleanGeneratedSlide(context, "DASHBOARD");
    dashboardId = dashboard.id;
    for (let i = 0; i < pages.length; i++) {
      const register = await addCleanGeneratedSlide(context, "REGISTER");
      registerIds.push(register.id);
    }
  });

  await PowerPoint.run(async context => {
    const slides = context.presentation.slides;
    const generated = [dashboardId, ...registerIds];
    for (let i = generated.length - 1; i >= 0; i--) slides.getItem(generated[i]).moveTo(0);
    await context.sync();
  });

  await PowerPoint.run(async context => {
    const dashboard = context.presentation.slides.getItem(dashboardId);
    buildDashboard(dashboard, issues);
    pages.forEach((groups, index) => {
      const register = context.presentation.slides.getItem(registerIds[index]);
      buildRegisterPage(register, groups, index + 1, pages.length);
    });
    await context.sync();
  });

  await refreshSummaryPreview();
  showToast(`Summary refreshed: 1 dashboard + ${pages.length} register page(s).`);
}

async function refreshSummaryPreview() {
  const issues = await readAllIssues();
  const actions = issues.flatMap(issue => issue.actions);
  ui.previewIssues.textContent = String(issues.length);
  ui.previewActions.textContent = String(actions.length);
  ui.previewOpen.textContent = String(actions.filter(a => cleanText(a.status).toLowerCase() !== "closed").length);
  ui.previewClosed.textContent = String(actions.filter(a => cleanText(a.status).toLowerCase() === "closed").length);
}

async function validatePresentation() {
  const issues = await readAllIssues();
  const parties = await getParties();
  const statuses = await getStatuses();
  const errors = [];
  const seen = new Set();

  issues.forEach(issue => {
    const id = cleanIssueId(issue.issueId);
    if (!id) errors.push("A tracked slide has no Issue ID.");
    if (seen.has(id)) errors.push(`${id}: duplicate Issue ID.`);
    seen.add(id);
    if (!cleanText(issue.description)) errors.push(`${id}: description is empty.`);
    issue.actions.forEach(action => {
      if (!cleanText(action.party)) errors.push(`${id}/${action.id}: party is empty.`);
      else if (!parties.includes(action.party)) errors.push(`${id}/${action.id}: party "${action.party}" is not in the party library.`);
      if (!cleanText(action.action)) errors.push(`${id}/${action.id}: action text is empty.`);
      if (!cleanText(action.status)) errors.push(`${id}/${action.id}: status is empty.`);
      else if (!statuses.includes(action.status)) errors.push(`${id}/${action.id}: status "${action.status}" is not in the status library.`);
    });
  });

  if (!issues.length) return showToast("No tracked issues found.", true);
  if (!errors.length) return showToast(`Validation passed: ${issues.length} issue(s) checked.`);
  showToast(`Validation found ${errors.length} problem(s). Open the console for details.`, true);
  console.warn("IssueFlow validation", errors);
}

function showToast(message, isError = false) {
  clearTimeout(toastTimer);
  ui.toast.textContent = message;
  ui.toast.className = `toast show${isError ? " error" : ""}`;
  toastTimer = setTimeout(() => ui.toast.className = "toast", 4200);
}

function showError(error) {
  console.error(error);
  showToast(error?.message || String(error), true);
}
