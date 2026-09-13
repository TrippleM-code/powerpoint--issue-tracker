const TAG_APP = "PIT_APP";
const TAG_ISSUE_ID = "PIT_ISSUE_ID";
const TAG_CREATED_AT = "PIT_CREATED_AT";
const TAG_UPDATED_AT = "PIT_UPDATED_AT";
const TAG_DESCRIPTION_V2 = "PIT_DESCRIPTION_V2";
const TAG_AREA_CODE_V2 = "PIT_AREA_CODE_V2";
const TAG_ROOM_NAME_V2 = "PIT_ROOM_NAME_V2";
const TAG_ACTIONS_V2 = "PIT_ACTIONS_V2";
const TAG_PARTY_LIBRARY = "PIT_PARTY_LIBRARY";
const TAG_PARTY_LOGOS = "PIT_PARTY_LOGOS";
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
    "issueId","setIssueId","areaCode","roomName","issueDescription","saveIssue","createIssueSheet",
    "currentIssueHeading","overallStatusChip","issueActionCount","issueCreatedDate","issueUpdatedDate",
    "issueNavigator","goToIssue",
    "actionParty","actionText","actionStatus","saveAction","clearActionForm",
    "actionEditorTitle","actionList","actionsCountBadge","actionsEmpty",
    "generateSummary","previewIssues","previewActions","previewOpen","previewClosed",
    "partyList","newParty","addParty","partyLogoFile","statusList","newStatus","addStatus","validatePresentation","toast"
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
  ui.partyLogoFile.addEventListener("change", () => handlePartyLogoSelected().catch(showError));
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
  const datePart = date.toLocaleDateString("en-GB", {
    day: "2-digit",
    month: "short",
    year: "numeric"
  });
  const timePart = date.toLocaleTimeString("en-GB", {
    hour: "2-digit",
    minute: "2-digit",
    hour12: false
  });
  return `${datePart} ${timePart}`;
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


async function getPartyLogos() {
  return PowerPoint.run(async context => {
    const value = await getTagValue(context.presentation.tags, TAG_PARTY_LOGOS, context);
    try {
      const parsed = JSON.parse(value || "{}");
      return parsed && typeof parsed === "object" && !Array.isArray(parsed) ? parsed : {};
    } catch {
      return {};
    }
  });
}

async function savePartyLogos(logos) {
  await PowerPoint.run(async context => {
    context.presentation.tags.add(TAG_PARTY_LOGOS, JSON.stringify(logos || {}));
    await context.sync();
  });
}

function fileToDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error("Could not read the logo file."));
    reader.onload = () => resolve(String(reader.result || ""));
    reader.readAsDataURL(file);
  });
}

function loadBrowserImage(dataUrl) {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.onerror = () => reject(new Error("Could not open the selected logo image."));
    image.onload = () => resolve(image);
    image.src = dataUrl;
  });
}

async function optimizeLogo(file) {
  if (!["image/png", "image/jpeg", "image/webp"].includes(file.type)) {
    throw new Error("Use PNG, JPG or WebP for logos.");
  }

  const sourceUrl = await fileToDataUrl(file);
  const image = await loadBrowserImage(sourceUrl);
  const scale = Math.min(1, 320 / image.width, 120 / image.height);
  const width = Math.max(1, Math.round(image.width * scale));
  const height = Math.max(1, Math.round(image.height * scale));

  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  const ctx = canvas.getContext("2d");
  if (!ctx) throw new Error("Image processing is unavailable.");

  ctx.clearRect(0, 0, width, height);
  ctx.drawImage(image, 0, 0, width, height);

  let dataUrl = canvas.toDataURL("image/png");
  if (dataUrl.length > 180000) dataUrl = canvas.toDataURL("image/jpeg", 0.82);
  if (dataUrl.length > 220000) {
    throw new Error("Logo is too large after optimization. Please use a simpler or smaller image.");
  }
  return dataUrl;
}

function dataUrlToBase64(dataUrl) {
  const index = String(dataUrl || "").indexOf(",");
  return index >= 0 ? dataUrl.slice(index + 1) : "";
}

let pendingLogoParty = null;

async function choosePartyLogo(party) {
  pendingLogoParty = party;
  ui.partyLogoFile.value = "";
  ui.partyLogoFile.click();
}

async function handlePartyLogoSelected() {
  const file = ui.partyLogoFile.files?.[0];
  const party = pendingLogoParty;
  pendingLogoParty = null;
  if (!file || !party) return;

  const optimized = await optimizeLogo(file);
  const logos = await getPartyLogos();
  logos[party] = optimized;
  await savePartyLogos(logos);
  await refreshLibrariesUi();
  showToast(`Logo saved for ${party}.`);
}

async function removePartyLogo(party) {
  const logos = await getPartyLogos();
  if (!logos[party]) return;
  delete logos[party];
  await savePartyLogos(logos);
  await refreshLibrariesUi();
  showToast(`Logo removed for ${party}.`);
}


async function refreshLibrariesUi() {
  const [parties, statuses, logos] = await Promise.all([getParties(), getStatuses(), getPartyLogos()]);

  populateSelect(ui.actionParty, parties);
  populateSelect(ui.actionStatus, statuses);

  renderPartyLibraryList(ui.partyList, parties, logos);
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


function renderPartyLibraryList(container, parties, logos) {
  container.innerHTML = "";

  parties.forEach(party => {
    const row = document.createElement("div");
    row.className = "library-item";

    const main = document.createElement("div");
    main.className = "party-main";

    if (logos[party]) {
      const img = document.createElement("img");
      img.className = "party-logo-preview";
      img.src = logos[party];
      img.alt = `${party} logo`;
      main.appendChild(img);
    } else {
      const placeholder = document.createElement("div");
      placeholder.className = "party-logo-placeholder";
      placeholder.textContent = "LOGO";
      main.appendChild(placeholder);
    }

    const name = document.createElement("span");
    name.className = "party-name";
    name.textContent = party;
    main.appendChild(name);

    const actions = document.createElement("div");
    actions.className = "party-actions";

    const upload = document.createElement("button");
    upload.textContent = logos[party] ? "Replace" : "Upload";
    upload.addEventListener("click", () => choosePartyLogo(party).catch(showError));
    actions.appendChild(upload);

    if (logos[party]) {
      const removeLogo = document.createElement("button");
      removeLogo.textContent = "Logo ×";
      removeLogo.addEventListener("click", () => removePartyLogo(party).catch(showError));
      actions.appendChild(removeLogo);
    }

    const remove = document.createElement("button");
    remove.textContent = "Remove";
    remove.addEventListener("click", () => removeParty(party).catch(showError));
    actions.appendChild(remove);

    row.append(main, actions);
    container.appendChild(row);
  });
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
    ui.areaCode.value = "";
    ui.roomName.value = "";
    ui.issueDescription.value = "";
    updateCurrentIssueUi(null);
    renderActions([]);
  }
}

async function loadCurrentIssue() {
  const issue = await readSelectedIssue();
  ui.issueId.value = issue.issueId || "";
  ui.areaCode.value = issue.areaCode || "";
  ui.roomName.value = issue.roomName || "";
  ui.issueDescription.value = issue.description || "";
  updateCurrentIssueUi(issue);
  renderActions(issue.actions);
}

async function readSelectedIssue() {
  return PowerPoint.run(async context => {
    const slide = await getSelectedSlide(context);
    const issueId = await getTagValue(slide.tags, TAG_ISSUE_ID, context);
    const descriptionTag = await getTagValue(slide.tags, TAG_DESCRIPTION_V2, context);
    const areaCode = await getTagValue(slide.tags, TAG_AREA_CODE_V2, context);
    const roomName = await getTagValue(slide.tags, TAG_ROOM_NAME_V2, context);
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
      areaCode: areaCode || "",
      roomName: roomName || "",
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
  const areaCode = cleanText(ui.areaCode.value).toUpperCase();
  const roomName = cleanText(ui.roomName.value);
  const description = cleanText(ui.issueDescription.value);

  await PowerPoint.run(async context => {
    const slide = await getSelectedSlide(context);
    const currentDescription = await getTagValue(slide.tags, TAG_DESCRIPTION_V2, context) || "";
    const currentAreaCode = await getTagValue(slide.tags, TAG_AREA_CODE_V2, context) || "";
    const currentRoomName = await getTagValue(slide.tags, TAG_ROOM_NAME_V2, context) || "";

    slide.tags.add(TAG_AREA_CODE_V2, areaCode);
    slide.tags.add(TAG_ROOM_NAME_V2, roomName);
    slide.tags.add(TAG_DESCRIPTION_V2, description);

    if (
      currentDescription !== description ||
      currentAreaCode !== areaCode ||
      currentRoomName !== roomName
    ) {
      slide.tags.add(TAG_UPDATED_AT, nowIso());
    }
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
      ...actions[index], party, action: actionText, status, updatedAt: timestamp
    };
  } else {
    actions.push({
      id: nextActionId(actions),
      party,
      action: actionText,
      status,

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

    card.append(top, actionText, controls);
    ui.actionList.appendChild(card);
  });
}

function editAction(action) {
  editingActionId = action.id;
  ui.actionEditorTitle.textContent = `Edit ${action.id}`;
  ui.actionParty.value = action.party;
  ui.actionText.value = action.action;
  ui.actionStatus.value = action.status;
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


function addLine(slide, left, top, width, height, color = "#C9D5DE", weight = 1) {
  // PowerPoint line options use left/top as the START point and
  // width/height as the END-POINT coordinates, not delta dimensions.
  // Callers in this app use delta-style width/height, so convert them here.
  const line = slide.shapes.addLine(PowerPoint.ConnectorType.straight, {
    left,
    top,
    width: left + width,
    height: top + height
  });
  line.lineFormat.color = color;
  line.lineFormat.weight = weight;
  return line;
}

function addSectionHeader(slide, title, left, top, width, options = {}) {
  const color = options.color || "#19364F";
  const lineColor = options.lineColor || "#8FB6CE";
  markManaged(addText(slide, title, left, top, width, 20, {
    size: options.size || 10,
    bold: true,
    color
  }), `${options.role || title}_TITLE`);
  markManaged(addLine(slide, left, top + 24, width, 0, lineColor, 1.4), `${options.role || title}_LINE`);
}

async function createOrRefreshIssueSheet() {
  await saveIssueData({ quiet: true });
  const issue = await readSelectedIssue();
  if (!issue.issueId) throw new Error("Set an Issue ID first.");

  const [parties, partyLogos] = await Promise.all([getParties(), getPartyLogos()]);

  await PowerPoint.run(async context => {
    const slide = await getSelectedSlide(context);
    slide.shapes.load("items/id,items/tags/key,items/tags/value");
    await context.sync();

    // Delete only IssueFlow-managed shapes. Manual reference images/annotations remain.
    for (const shape of slide.shapes.items) {
      const managed = shape.tags.items.find(t => t.key === TAG_MANAGED && t.value === "TRUE");
      if (managed) shape.delete();
    }
    await context.sync();

    const navy = "#17344C";
    const text = "#26343F";
    const muted = "#6F7F8A";
    const line = "#AFC3D1";
    const softBlue = "#EAF5FB";
    const softRose = "#FFF1F1";
    const softGold = "#FFF8E2";

    // --------------------------------------------------------
    // HEADER: party cells use all parties in the party library.
    // Issue ID and Date/Time sit immediately to the right.
    // --------------------------------------------------------
    const headerTop = 4;
    const headerHeight = 72;
    const issueIdWidth = 132;
    const dateWidth = 178;
    const logoAreaWidth = 960 - issueIdWidth - dateWidth;
    const partyCount = Math.max(parties.length, 1);
    const logoCellWidth = logoAreaWidth / partyCount;

    // Party cells.
    parties.forEach((party, index) => {
      const left = index * logoCellWidth;

      // Soft fill, no box outline. Only vertical separators.
      const bg = addRect(slide, left, headerTop, logoCellWidth, headerHeight, "#F5FBFE", "#F5FBFE");
      bg.lineFormat.weight = 0;
      markManaged(bg, `HEADER_PARTY_BG_${index}`);

      if (index > 0) {
        markManaged(addLine(slide, left, headerTop + 8, 0, headerHeight - 16, "#B8D5E6", 1), `HEADER_PARTY_SEP_${index}`);
      }

      const logoData = partyLogos[party];
      let pictureAdded = false;

      if (logoData && typeof slide.shapes.addPicture === "function") {
        try {
          const base64 = dataUrlToBase64(logoData);
          const maxW = Math.max(34, logoCellWidth - 20);
          const picW = Math.min(maxW, 82);
          const picture = slide.shapes.addPicture(base64, {
            left: left + (logoCellWidth - picW) / 2,
            top: headerTop + 8,
            width: picW,
            height: 38
          });
          markManaged(picture, `HEADER_PARTY_LOGO_${index}`);
          pictureAdded = true;
        } catch (error) {
          console.warn(`Could not render logo for ${party}.`, error);
        }
      }

      markManaged(addText(
        slide,
        party,
        left + 6,
        headerTop + (pictureAdded ? 50 : 26),
        logoCellWidth - 12,
        16,
        {
          size: Math.max(6.5, Math.min(9, logoCellWidth / 13)),
          bold: true,
          color: "#174A6A"
        }
      ), `HEADER_PARTY_NAME_${index}`);
    });

    if (!parties.length) {
      markManaged(addText(slide, "Add parties / logos in Settings", 24, headerTop + 25, logoAreaWidth - 48, 18, {
        size: 9,
        color: muted
      }), "HEADER_NO_PARTIES");
    }

    // Issue ID box.
    const issueLeft = logoAreaWidth;
    const issueBg = addRect(slide, issueLeft, headerTop, issueIdWidth, headerHeight, softRose, softRose);
    issueBg.lineFormat.weight = 0;
    markManaged(issueBg, "HEADER_ISSUE_BG");
    markManaged(addLine(slide, issueLeft, headerTop, 0, headerHeight, "#DAB3B3", 1.2), "HEADER_ISSUE_LEFT");
    markManaged(addText(slide, "ISSUE ID", issueLeft + 10, headerTop + 8, issueIdWidth - 20, 14, {
      size: 7.2, bold: true, color: "#694747"
    }), "HEADER_ISSUE_LABEL");
    markManaged(addText(slide, issue.issueId, issueLeft + 10, headerTop + 25, issueIdWidth - 20, 24, {
      size: 17, bold: true, color: "#17212B"
    }), "HEADER_ISSUE_ID");
    if (issue.areaCode) {
      markManaged(addText(slide, issue.areaCode, issueLeft + 10, headerTop + 52, issueIdWidth - 20, 12, {
        size: 7.2, bold: true, color: "#835F5F"
      }), "HEADER_AREA_CODE");
    }

    // Date / time box.
    const dateLeft = issueLeft + issueIdWidth;
    const dateBg = addRect(slide, dateLeft, headerTop, dateWidth, headerHeight, softGold, softGold);
    dateBg.lineFormat.weight = 0;
    markManaged(dateBg, "HEADER_DATE_BG");
    markManaged(addLine(slide, dateLeft, headerTop, 0, headerHeight, "#E1C971", 1.2), "HEADER_DATE_LEFT");
    markManaged(addText(slide, `Created  ${formatDate(issue.createdAt)}`, dateLeft + 12, headerTop + 14, dateWidth - 24, 16, {
      size: 7.5, bold: true, color: "#3D4850"
    }), "HEADER_CREATED");
    if (issue.updatedAt) {
      markManaged(addText(slide, `Updated  ${formatDate(issue.updatedAt)}`, dateLeft + 12, headerTop + 39, dateWidth - 24, 16, {
        size: 7.5, color: "#3D4850"
      }), "HEADER_UPDATED");
    }

    // Thin baseline only.
    markManaged(addLine(slide, 0, headerTop + headerHeight, 960, 0, "#8DB2C8", 1.4), "HEADER_BASELINE");

    // --------------------------------------------------------
    // MAIN LAYOUT
    // Description | Room strip | Reference Images | Actions
    // --------------------------------------------------------
    const mainTop = 92;
    const mainBottom = 520;
    const contentTop = 126;
    const contentHeight = mainBottom - contentTop;

    const descriptionLeft = 16;
    const descriptionWidth = 248;

    const roomLeft = descriptionLeft + descriptionWidth + 8;
    const roomWidth = 42;

    const referenceLeft = roomLeft + roomWidth + 8;
    const referenceWidth = 286;

    const actionsLeft = referenceLeft + referenceWidth + 12;
    const actionsWidth = 960 - actionsLeft - 16;

    // Headers: text + one underline, no unnecessary border boxes.
    addSectionHeader(slide, "ISSUE DESCRIPTION", descriptionLeft, mainTop, descriptionWidth, {
      role: "DESCRIPTION_HEAD"
    });
    addSectionHeader(slide, "REFERENCE IMAGES", referenceLeft, mainTop, referenceWidth, {
      role: "REFERENCE_HEAD"
    });
    addSectionHeader(slide, "ACTIONS", actionsLeft, mainTop, actionsWidth, {
      role: "ACTIONS_HEAD"
    });

    // Description: only the necessary separator/baseline lines.
    markManaged(addLine(slide, descriptionLeft, contentTop, descriptionWidth, 0, line, 1.2), "DESCRIPTION_TOP");
    markManaged(addLine(slide, descriptionLeft + descriptionWidth, contentTop, 0, contentHeight, line, 1.2), "DESCRIPTION_RIGHT");
    markManaged(addLine(slide, descriptionLeft, mainBottom, descriptionWidth, 0, line, 1.2), "DESCRIPTION_BOTTOM");
    markManaged(addText(slide, issue.description || "No description yet.", descriptionLeft + 12, contentTop + 14, descriptionWidth - 26, contentHeight - 26, {
      size: 10.5,
      color: text
    }), "DESCRIPTION_TEXT");

    // Room / Space strip: vertically rotated text like "Basement-1".
    // Minimal separator lines only.
    markManaged(addLine(slide, roomLeft, contentTop, 0, contentHeight, line, 1.2), "ROOM_LEFT");
    markManaged(addLine(slide, roomLeft + roomWidth, contentTop, 0, contentHeight, line, 1.2), "ROOM_RIGHT");

    const roomLabel = markManaged(addText(
      slide,
      issue.roomName || "Room / Space",
      roomLeft + 9,
      contentTop + 68,
      contentHeight - 136,
      24,
      {
        size: 10.5,
        bold: true,
        color: "#17212B"
      }
    ), "ROOM_NAME");
    roomLabel.rotation = 270;

    // Reference area remains intentionally blank/manual.
    // Only subtle top/bottom separators are generated.
    markManaged(addLine(slide, referenceLeft, contentTop, referenceWidth, 0, line, 1.2), "REFERENCE_TOP");
    markManaged(addLine(slide, referenceLeft + referenceWidth, contentTop, 0, contentHeight, line, 1.2), "REFERENCE_RIGHT");
    markManaged(addLine(slide, referenceLeft, mainBottom, referenceWidth, 0, line, 1.2), "REFERENCE_BOTTOM");

    // --------------------------------------------------------
    // ACTIONS: exactly N rows for N actions.
    // Action By | Action Required | Status.
    // No legacy extra field, no rounded cards, no unused rows.
    // --------------------------------------------------------
    const tableTop = contentTop;
    const tableBottom = mainBottom;
    const headerH = 30;

    const partyW = Math.round(actionsWidth * 0.24);
    const statusW = Math.round(actionsWidth * 0.22);
    const actionW = actionsWidth - partyW - statusW;

    const partyX = actionsLeft;
    const actionX = partyX + partyW;
    const statusX = actionX + actionW;

    // Column headers.
    markManaged(addText(slide, "Action By", partyX + 8, tableTop + 7, partyW - 16, 14, {
      size: 8.4, bold: true, color: "#243746"
    }), "ACTIONS_COL_PARTY");
    markManaged(addText(slide, "Action Required", actionX + 8, tableTop + 7, actionW - 16, 14, {
      size: 8.4, bold: true, color: "#243746"
    }), "ACTIONS_COL_ACTION");
    markManaged(addText(slide, "Status", statusX + 8, tableTop + 7, statusW - 16, 14, {
      size: 8.4, bold: true, color: "#243746"
    }), "ACTIONS_COL_STATUS");

    markManaged(addLine(slide, actionsLeft, tableTop + headerH, actionsWidth, 0, line, 1.2), "ACTIONS_HEADER_LINE");
    markManaged(addLine(slide, actionX, tableTop, 0, tableBottom - tableTop, line, 1.0), "ACTIONS_SEP_1");
    markManaged(addLine(slide, statusX, tableTop, 0, tableBottom - tableTop, line, 1.0), "ACTIONS_SEP_2");
    markManaged(addLine(slide, actionsLeft + actionsWidth, tableTop, 0, tableBottom - tableTop, line, 1.2), "ACTIONS_RIGHT");
    markManaged(addLine(slide, actionsLeft, tableBottom, actionsWidth, 0, line, 1.2), "ACTIONS_BOTTOM");

    const actions = issue.actions || [];
    const overall = computeOverallStatus(actions);
    markManaged(addText(slide, `Overall: ${overall}`, actionsLeft + actionsWidth - 104, mainTop + 3, 100, 14, {
      size: 7.3,
      bold: true,
      color: statusColor(overall).text
    }), "OVERALL_STATUS");

    if (!actions.length) {
      markManaged(addText(slide, "No actions yet", actionsLeft + 10, tableTop + 62, actionsWidth - 20, 20, {
        size: 10,
        color: muted
      }), "ACTIONS_EMPTY");
    } else {
      const rowsTop = tableTop + headerH;
      const rowsHeight = tableBottom - rowsTop;
      const rowHeight = rowsHeight / actions.length;

      actions.forEach((action, index) => {
        const rowTop = rowsTop + index * rowHeight;
        const rowBottom = rowsTop + (index + 1) * rowHeight;
        const rowTextSize = rowHeight >= 70 ? 9.4 : rowHeight >= 48 ? 8.5 : 7.2;
        const sc = statusColor(action.status);

        if (index > 0) {
          markManaged(addLine(slide, actionsLeft, rowTop, actionsWidth, 0, line, 1.0), `ACTION_${action.id}_ROW_LINE`);
        }

        markManaged(addText(slide, action.party || "—", partyX + 8, rowTop + 9, partyW - 16, Math.max(16, rowHeight - 18), {
          size: rowTextSize,
          bold: true,
          color: text
        }), `ACTION_${action.id}_PARTY`);

        markManaged(addText(slide, action.action || "—", actionX + 8, rowTop + 9, actionW - 16, Math.max(16, rowHeight - 18), {
          size: rowTextSize,
          color: text
        }), `ACTION_${action.id}_TEXT`);

        // Status uses text with a subtle status-colored underline rather than a rounded chip.
        markManaged(addText(slide, action.status || "—", statusX + 8, rowTop + 9, statusW - 16, 18, {
          size: rowTextSize,
          bold: true,
          color: sc.text
        }), `ACTION_${action.id}_STATUS`);
        markManaged(addLine(slide, statusX + 8, Math.min(rowBottom - 10, rowTop + 30), statusW - 16, 0, sc.line, 2), `ACTION_${action.id}_STATUS_LINE`);
      });
    }

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
      const areaCode = slide.tags.items.find(t => t.key === TAG_AREA_CODE_V2)?.value || "";
      const roomName = slide.tags.items.find(t => t.key === TAG_ROOM_NAME_V2)?.value || "";
      const actionsRaw = slide.tags.items.find(t => t.key === TAG_ACTIONS_V2)?.value || "[]";
      const createdAt = slide.tags.items.find(t => t.key === TAG_CREATED_AT)?.value || "";
      const updatedAt = slide.tags.items.find(t => t.key === TAG_UPDATED_AT)?.value || "";
      let actions = [];
      try { actions = JSON.parse(actionsRaw); } catch { actions = []; }
      if (!Array.isArray(actions)) actions = [];
      issues.push({ slideId: slide.id, issueId, areaCode, roomName, description, actions, createdAt, updatedAt });
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

  const left = 24;
  const top = 92;
  const rowH = 42;
  const widths = [92, 104, 214, 105, 280, 105];
  const headers = ["Issue ID", "Area / Room", "Description", "Action By", "Action Required", "Status"];

  let x = left;
  headers.forEach((header, index) => {
    addRect(slide, x, top, widths[index], 28, "#18324A", "#18324A");
    addText(slide, header, x + 7, top + 8, widths[index] - 12, 12, {
      size: 8.0,
      bold: true,
      color: "#FFFFFF"
    });
    x += widths[index];
  });

  let rowIndex = 0;

  groups.forEach(group => {
    const issue = group.issue;
    const actions = issue.actions.length
      ? issue.actions
      : [{ party: "—", action: "No actions", status: "—" }];

    const visibleActions = actions.slice(group.start, group.start + group.count);
    const groupTop = top + 28 + rowIndex * rowH;
    const groupHeight = visibleActions.length * rowH;

    // Grouped Issue ID
    addRect(slide, left, groupTop, widths[0], groupHeight, "#F5F8FA", "#D8E1E7");
    addText(slide, issue.issueId, left + 8, groupTop + 10, widths[0] - 16, Math.max(18, groupHeight - 16), {
      size: 9.2,
      bold: true,
      color: "#1769AA"
    });

    // Grouped Area / Room
    const areaLeft = left + widths[0];
    addRect(slide, areaLeft, groupTop, widths[1], groupHeight, "#FBFCFD", "#D8E1E7");
    const areaRoom = [issue.areaCode, issue.roomName].filter(Boolean).join("\n");
    addText(slide, areaRoom || "—", areaLeft + 8, groupTop + 8, widths[1] - 16, Math.max(20, groupHeight - 14), {
      size: 8.3,
      color: "#26343F"
    });

    // Grouped Description
    const descLeft = areaLeft + widths[1];
    addRect(slide, descLeft, groupTop, widths[2], groupHeight, "#FFFFFF", "#D8E1E7");
    addText(slide, issue.description || "—", descLeft + 8, groupTop + 8, widths[2] - 16, Math.max(20, groupHeight - 14), {
      size: 8.3,
      color: "#26343F"
    });

    visibleActions.forEach((action, localIndex) => {
      const y = groupTop + localIndex * rowH;
      let dx = descLeft + widths[2];

      const values = [
        action.party || "—",
        action.action || "—",
        action.status || "—"
      ];
      const dataWidths = widths.slice(3);

      values.forEach((value, index) => {
        const isStatus = index === 2;
        const sc = isStatus ? statusColor(value) : null;
        addRect(
          slide,
          dx,
          y,
          dataWidths[index],
          rowH,
          isStatus ? sc.fill : (rowIndex % 2 ? "#F9FBFC" : "#FFFFFF"),
          isStatus ? sc.line : "#D8E1E7"
        );
        addText(slide, value, dx + 7, y + 9, dataWidths[index] - 14, rowH - 14, {
          size: isStatus ? 8.4 : 8.2,
          bold: isStatus,
          color: isStatus ? sc.text : "#26343F"
        });
        dx += dataWidths[index];
      });

      rowIndex++;
    });
  });
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
    if (!cleanText(issue.areaCode)) errors.push(`${id}: area code is empty.`);
    if (!cleanText(issue.roomName)) errors.push(`${id}: room / space is empty.`);
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
