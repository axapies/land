type Point = {
  x: number;
  y: number;
};

type RectElement = {
  id: string;
  kind: "rect";
  category: ElementCategory;
  name: string;
  x: number;
  y: number;
  width: number;
  height: number;
  rotation: number;
  color: string;
  showInLegend: boolean;
};

type PolyElement = {
  id: string;
  kind: "poly";
  category: ElementCategory;
  name: string;
  points: Point[];
  rotation: 0;
  color: string;
  showInLegend: boolean;
};

type SiteElement = RectElement | PolyElement;

type PlanNote = {
  id: string;
  text: string;
  x: number;
  y: number;
};

type BackgroundImageState = {
  dataUrl: string;
  name: string;
  x: number;
  y: number;
  width: number;
  height: number;
  opacity: number;
  locked: boolean;
};

type Selection =
  | { kind: "none" }
  | { kind: "land" }
  | { kind: "element"; id: string }
  | { kind: "note"; id: string };

type Project = {
  id: string;
  name: string;
  createdAt: string;
  updatedAt: string;
  zoom: number;
  pan: Point;
  land: Point[];
  landFrontEdge: number | null;
  landRotation: number;
  elements: SiteElement[];
  notes: PlanNote[];
  background?: BackgroundImageState;
  showDistances: boolean;
  snapping: boolean;
  selected: Selection;
  history: ProjectSnapshot[];
  future: ProjectSnapshot[];
};

type ProjectSnapshot = Omit<Project, "history" | "future">;

type ElementCategory =
  | "house"
  | "garage"
  | "garden"
  | "driveway"
  | "terrace"
  | "shed"
  | "utility"
  | "custom";

type NewElementDraft = {
  category: ElementCategory;
  name: string;
  kind: "rect" | "poly";
  width: number;
  height: number;
  color: string;
};

type Tool = "select" | "pan" | "draw-land" | "place-rect" | "place-poly" | "place-note";

type PrintSize = "A4" | "A3";

type AppState = {
  projects: Project[];
  activeProjectId: string;
  tool: Tool;
  landDraft: Point[];
  elementDraft: Point[];
  newElement: NewElementDraft;
  landRectWidth: number;
  landRectDepth: number;
  noteText: string;
  printSize: PrintSize;
  includeSummary: boolean;
  showHistory: boolean;
};

type DragState =
  | { type: "none" }
  | { type: "pan"; startClient: Point; startPan: Point }
  | { type: "element"; id: string; lastPoint: Point }
  | { type: "note"; id: string; lastPoint: Point }
  | { type: "land-vertex"; index: number }
  | { type: "element-vertex"; id: string; index: number };

type BBox = {
  minX: number;
  minY: number;
  maxX: number;
  maxY: number;
};

const STORAGE_KEY = "land-layout-planner:v1";
const MAX_HISTORY = 80;
const WORLD_WIDTH = 72;
const WORLD_HEIGHT = 46;
const DEFAULT_PAN: Point = { x: -8, y: 0 };
const LEGACY_DEFAULT_PAN: Point = { x: -8, y: -7 };

const categoryDefaults: Record<
  ElementCategory,
  { label: string; color: string; width: number; height: number }
> = {
  house: { label: "House", color: "#2f6d7e", width: 14, height: 9 },
  garage: { label: "Garage", color: "#8a6d3b", width: 6, height: 6 },
  garden: { label: "Garden", color: "#6e9f58", width: 10, height: 8 },
  driveway: { label: "Driveway", color: "#6e7480", width: 4, height: 14 },
  terrace: { label: "Terrace", color: "#b7845a", width: 7, height: 4 },
  shed: { label: "Shed", color: "#8c6f56", width: 3, height: 3 },
  utility: { label: "Utility area", color: "#7c6ab0", width: 5, height: 4 },
  custom: { label: "Custom item", color: "#3f86b6", width: 5, height: 5 },
};

const app = document.querySelector<HTMLDivElement>("#app");

if (!app) {
  throw new Error("Missing #app root");
}

let state = loadState();
let drag: DragState = { type: "none" };

render();

app.addEventListener("click", handleClick);
app.addEventListener("input", handleInput);
app.addEventListener("change", handleChange);
app.addEventListener("pointerdown", handlePointerDown);
app.addEventListener("dblclick", handleDoubleClick);
window.addEventListener("pointermove", handlePointerMove);
window.addEventListener("pointerup", handlePointerUp);
window.addEventListener("keydown", handleKeyDown);

function loadState(): AppState {
  const fallback = createInitialState();
  const raw = localStorage.getItem(STORAGE_KEY);

  if (!raw) {
    return fallback;
  }

  try {
    const parsed = JSON.parse(raw) as Partial<AppState>;
    const projects = Array.isArray(parsed.projects)
      ? parsed.projects.map(normalizeProject).filter(Boolean)
      : [];

    if (!projects.length) {
      return fallback;
    }

    const activeProjectId =
      typeof parsed.activeProjectId === "string" &&
      projects.some((project) => project.id === parsed.activeProjectId)
        ? parsed.activeProjectId
        : projects[0].id;

    return {
      projects,
      activeProjectId,
      tool: "select",
      landDraft: [],
      elementDraft: [],
      newElement: normalizeNewElement(parsed.newElement),
      landRectWidth: normalizePositiveNumber(parsed.landRectWidth, 23),
      landRectDepth: normalizePositiveNumber(parsed.landRectDepth, 45.5),
      noteText: typeof parsed.noteText === "string" ? parsed.noteText : "Check sightline here",
      printSize: parsed.printSize === "A3" ? "A3" : "A4",
      includeSummary: parsed.includeSummary !== false,
      showHistory: false,
    };
  } catch {
    return fallback;
  }
}

function createInitialState(): AppState {
  const project = createProject("Option A");

  return {
    projects: [project],
    activeProjectId: project.id,
    tool: "select",
    landDraft: [],
    elementDraft: [],
    newElement: {
      category: "house",
      name: "House",
      kind: "rect",
      width: 14,
      height: 9,
      color: categoryDefaults.house.color,
    },
    landRectWidth: 23,
    landRectDepth: 45.5,
    noteText: "Check sightline here",
    printSize: "A4",
    includeSummary: true,
    showHistory: false,
  };
}

function createProject(name: string, seeded = true): Project {
  const now = new Date().toISOString();
  const houseId = uid("element");
  const drivewayId = uid("element");
  const land = seeded
    ? [
        { x: 0, y: 0 },
        { x: 42, y: 0 },
        { x: 42, y: 28 },
        { x: 0, y: 28 },
      ]
    : [];
  const elements: SiteElement[] = seeded
    ? [
        {
          id: houseId,
          kind: "rect",
          category: "house",
          name: "House",
          x: 18,
          y: 11,
          width: 14,
          height: 9,
          rotation: 0,
          color: categoryDefaults.house.color,
          showInLegend: true,
        },
        {
          id: drivewayId,
          kind: "rect",
          category: "driveway",
          name: "Driveway",
          x: 34,
          y: 8,
          width: 4,
          height: 14,
          rotation: 0,
          color: categoryDefaults.driveway.color,
          showInLegend: true,
        },
      ]
    : [];
  const notes: PlanNote[] = seeded
    ? [{ id: uid("note"), text: "North garden kept open", x: 7, y: 24 }]
    : [];

  return {
    id: uid("project"),
    name,
    createdAt: now,
    updatedAt: now,
    zoom: 1,
    pan: { ...DEFAULT_PAN },
    land,
    landFrontEdge: seeded ? 2 : null,
    landRotation: 0,
    elements,
    notes,
    showDistances: true,
    snapping: true,
    selected: seeded ? { kind: "element", id: houseId } : { kind: "none" },
    history: [],
    future: [],
  };
}

function normalizeProject(input: unknown): Project {
  const source = input as Partial<Project>;
  const project = createProject(typeof source.name === "string" ? source.name : "Imported project");

  project.id = typeof source.id === "string" ? source.id : project.id;
  project.createdAt = typeof source.createdAt === "string" ? source.createdAt : project.createdAt;
  project.updatedAt = typeof source.updatedAt === "string" ? source.updatedAt : project.updatedAt;
  project.zoom = clamp(toNumber(source.zoom, 1), 0.45, 4);
  project.pan = normalizePan(source.pan, project.pan);
  project.land = normalizePoints(source.land, project.land);
  project.landFrontEdge = normalizeLandFrontEdge(source.landFrontEdge, project.land);
  project.landRotation = toNumber(source.landRotation, 0);
  project.elements = Array.isArray(source.elements)
    ? source.elements.map(normalizeElement).filter(Boolean)
    : project.elements;
  project.notes = Array.isArray(source.notes)
    ? source.notes.map(normalizeNote).filter(Boolean)
    : project.notes;
  project.background = normalizeBackground(source.background);
  project.showDistances = source.showDistances !== false;
  project.snapping = source.snapping !== false;
  project.selected = { kind: "none" };
  project.history = [];
  project.future = [];

  return project;
}

function normalizeNewElement(input: unknown): NewElementDraft {
  const draft = (input ?? {}) as Partial<NewElementDraft>;
  const category = isElementCategory(draft.category) ? draft.category : "house";
  const defaults = categoryDefaults[category];

  return {
    category,
    name: typeof draft.name === "string" && draft.name.trim() ? draft.name : defaults.label,
    kind: draft.kind === "poly" ? "poly" : "rect",
    width: clamp(toNumber(draft.width, defaults.width), 0.2, 200),
    height: clamp(toNumber(draft.height, defaults.height), 0.2, 200),
    color: typeof draft.color === "string" ? draft.color : defaults.color,
  };
}

function normalizePositiveNumber(value: unknown, fallback: number) {
  return clamp(toNumber(value, fallback), 0.1, 10000);
}

function normalizeElement(input: unknown): SiteElement | undefined {
  const source = input as Partial<SiteElement>;
  const category = isElementCategory(source.category) ? source.category : "custom";
  const defaults = categoryDefaults[category];

  if (source.kind === "poly") {
    const points = normalizePoints((source as Partial<PolyElement>).points, []);
    if (points.length < 3) {
      return undefined;
    }

    return {
      id: typeof source.id === "string" ? source.id : uid("element"),
      kind: "poly",
      category,
      name: typeof source.name === "string" ? source.name : defaults.label,
      points,
      rotation: 0,
      color: typeof source.color === "string" ? source.color : defaults.color,
      showInLegend: source.showInLegend !== false,
    };
  }

  return {
    id: typeof source.id === "string" ? source.id : uid("element"),
    kind: "rect",
    category,
    name: typeof source.name === "string" ? source.name : defaults.label,
    x: toNumber((source as Partial<RectElement>).x, 8),
    y: toNumber((source as Partial<RectElement>).y, 8),
    width: clamp(toNumber((source as Partial<RectElement>).width, defaults.width), 0.2, 200),
    height: clamp(toNumber((source as Partial<RectElement>).height, defaults.height), 0.2, 200),
    rotation: toNumber((source as Partial<RectElement>).rotation, 0),
    color: typeof source.color === "string" ? source.color : defaults.color,
    showInLegend: source.showInLegend !== false,
  };
}

function normalizeNote(input: unknown): PlanNote | undefined {
  const source = input as Partial<PlanNote>;

  if (typeof source.text !== "string" || !source.text.trim()) {
    return undefined;
  }

  return {
    id: typeof source.id === "string" ? source.id : uid("note"),
    text: source.text,
    x: toNumber(source.x, 0),
    y: toNumber(source.y, 0),
  };
}

function normalizeBackground(input: unknown): BackgroundImageState | undefined {
  const source = input as Partial<BackgroundImageState>;

  if (typeof source.dataUrl !== "string" || !source.dataUrl.startsWith("data:image/")) {
    return undefined;
  }

  return {
    dataUrl: source.dataUrl,
    name: typeof source.name === "string" ? source.name : "Background image",
    x: toNumber(source.x, 0),
    y: toNumber(source.y, 0),
    width: clamp(toNumber(source.width, 40), 0.2, 500),
    height: clamp(toNumber(source.height, 25), 0.2, 500),
    opacity: clamp(toNumber(source.opacity, 0.45), 0.05, 1),
    locked: source.locked !== false,
  };
}

function render() {
  const project = activeProject();
  const selectedElement =
    project.selected.kind === "element"
      ? project.elements.find((element) => element.id === project.selected.id)
      : undefined;
  const selectedNote =
    project.selected.kind === "note"
      ? project.notes.find((note) => note.id === project.selected.id)
      : undefined;
  const warnings = getWarnings(project);

  document.body.dataset.printSize = state.printSize;
  document.body.classList.toggle("with-summary", state.includeSummary);

  app.innerHTML = `
    <div class="shell">
      <header class="topbar">
        <div class="brand">
          <span class="brand-mark" aria-hidden="true"></span>
          <div>
            <strong>Land Layout Planner</strong>
            <small>Metric SVG drafting workspace</small>
          </div>
        </div>
        <nav class="project-tabs" aria-label="Open projects">
          ${state.projects.map(renderProjectTab).join("")}
          <button class="tab tab-action" data-action="new-project" type="button">+ New blank</button>
          <button class="tab tab-action ${state.showHistory ? "active" : ""}" data-action="toggle-history" type="button">History</button>
        </nav>
        <div class="top-actions">
          <button data-action="undo" type="button" ${project.history.length ? "" : "disabled"}>Undo</button>
          <button data-action="redo" type="button" ${project.future.length ? "" : "disabled"}>Redo</button>
          <button data-action="export-json" type="button">Export JSON</button>
          <label class="file-button">
            Import JSON
            <input id="import-project" type="file" accept="application/json,.json" />
          </label>
          <button class="primary" data-action="print" type="button">Print / PDF</button>
        </div>
      </header>
      ${state.showHistory ? renderProjectHistory(project.id) : ""}

      <main class="planner">
        <aside class="panel panel-left">
          ${renderWorkspacePanel(project)}
          ${renderAddElementPanel()}
          ${renderBackgroundPanel(project)}
        </aside>

        <section class="workspace">
          <div class="canvas-strip">
            <div>
              <strong>${escapeHtml(project.name)}</strong>
              <span>${project.land.length} boundary points - ${project.elements.length} elements - ${project.notes.length} notes</span>
            </div>
            <div class="view-controls">
              <label>
                Zoom
                <input data-field="project.zoom" type="range" min="0.45" max="4" step="0.05" value="${project.zoom}" />
              </label>
              <button data-action="zoom-reset" type="button">Reset view</button>
            </div>
          </div>
          <div class="canvas-frame">
            ${renderSvg(project, warnings)}
          </div>
          <div class="statusbar">
            <span>Scale: 1 canvas unit = 1 m - Measurements round to 0.1 m</span>
            <span>${toolStatus()}</span>
          </div>
        </section>

        <aside class="panel panel-right">
          ${renderInspector(project, selectedElement, selectedNote, warnings)}
          ${renderLegend(project)}
        </aside>
      </main>

      ${renderPrintSummary(project, warnings)}
    </div>
  `;
}

function renderProjectTab(project: Project) {
  const isActive = project.id === state.activeProjectId;

  return `
    <button class="tab ${isActive ? "active" : ""}" data-project-id="${project.id}" type="button">
      ${escapeHtml(project.name)}
    </button>
  `;
}

function renderProjectHistory(activeId: string) {
  const rows = [...state.projects]
    .sort((a, b) => Date.parse(b.updatedAt) - Date.parse(a.updatedAt))
    .map((project) => {
      const isActive = project.id === activeId;
      const landLabel = project.land.length >= 3 ? `${project.land.length} boundary points` : "No land yet";

      return `
        <article class="history-row ${isActive ? "active" : ""}">
          <div>
            <strong>${escapeHtml(project.name)}</strong>
            <small>${escapeHtml(formatDateTime(project.updatedAt))} - ${landLabel} - ${project.elements.length} objects</small>
          </div>
          <div class="history-actions">
            <button data-action="open-history-project" data-project-action-id="${project.id}" type="button" ${
              isActive ? "disabled" : ""
            }>Open</button>
            <button data-action="duplicate-history-project" data-project-action-id="${project.id}" type="button">Duplicate</button>
            <button class="danger-button" data-action="delete-history-project" data-project-action-id="${project.id}" type="button">Delete</button>
          </div>
        </article>
      `;
    })
    .join("");

  return `
    <section class="history-drawer" aria-label="Saved project history">
      <div>
        <strong>Project history</strong>
        <span>Saved locally in this browser. Use export/import to move projects between devices.</span>
      </div>
      <div class="history-list">${rows}</div>
    </section>
  `;
}

function renderWorkspacePanel(project: Project) {
  return `
    <section class="panel-section">
      <div class="section-heading">
        <span>Workspace</span>
        <div class="section-actions">
          <button data-action="duplicate-project" type="button">Duplicate</button>
          <button class="danger-button" data-action="delete-project" type="button">Delete option</button>
        </div>
      </div>
      <label class="field">
        Project name
        <input data-field="project.name" type="text" value="${escapeAttr(project.name)}" />
      </label>
      <div class="toggle-row">
        <label><input data-field="project.snapping" type="checkbox" ${project.snapping ? "checked" : ""} /> Snap</label>
        <label><input data-field="project.showDistances" type="checkbox" ${project.showDistances ? "checked" : ""} /> Distances</label>
      </div>
      <div class="button-grid">
        ${toolButton("select", "Select")}
        ${toolButton("pan", "Pan")}
        ${toolButton("draw-land", "Draw land")}
        ${toolButton("place-note", "Place note")}
      </div>
      <div class="land-setup">
        <div class="field-pair">
          <label class="field">
            Rect width (m)
            <input data-field="landRect.width" type="number" min="0.1" step="0.1" value="${roundM(
              state.landRectWidth,
            )}" />
          </label>
          <label class="field">
            Rect depth (m)
            <input data-field="landRect.depth" type="number" min="0.1" step="0.1" value="${roundM(
              state.landRectDepth,
            )}" />
          </label>
        </div>
        <p class="hint">Rectangle area: ${formatMeters(state.landRectWidth * state.landRectDepth, "m2")}</p>
      </div>
      <div class="button-grid land-actions">
        <button data-action="rect-land" type="button">Create rectangular land</button>
        <button data-action="remove-land" type="button" ${project.land.length ? "" : "disabled"}>Remove land</button>
      </div>
      <p class="hint">Hold Alt while dragging to temporarily bypass snapping. Double-click to finish freeform land or polygon objects.</p>
      ${
        state.landDraft.length
          ? `<div class="draft-actions">
              <strong>${state.landDraft.length} land points drafted</strong>
              <button data-action="finish-land" type="button" ${state.landDraft.length < 3 ? "disabled" : ""}>Finish land</button>
              <button data-action="cancel-land-draft" type="button">Cancel</button>
            </div>`
          : ""
      }
      <label class="field">
        Note text
        <input data-field="noteText" type="text" value="${escapeAttr(state.noteText)}" />
      </label>
    </section>
  `;
}

function renderAddElementPanel() {
  const defaults = categoryDefaults[state.newElement.category];

  return `
    <section class="panel-section">
      <div class="section-heading">
        <span>Add Element</span>
        <button data-action="place-element" type="button">Place</button>
      </div>
      <label class="field">
        Category
        <select data-field="new.category">
          ${Object.entries(categoryDefaults)
            .map(
              ([value, option]) =>
                `<option value="${value}" ${value === state.newElement.category ? "selected" : ""}>${escapeHtml(
                  option.label,
                )}</option>`,
            )
            .join("")}
        </select>
      </label>
      <label class="field">
        Label
        <input data-field="new.name" type="text" value="${escapeAttr(state.newElement.name || defaults.label)}" />
      </label>
      <div class="field-pair">
        <label class="field">
          Shape
          <select data-field="new.kind">
            <option value="rect" ${state.newElement.kind === "rect" ? "selected" : ""}>Rectangle</option>
            <option value="poly" ${state.newElement.kind === "poly" ? "selected" : ""}>Freeform</option>
          </select>
        </label>
        <label class="field">
          Color
          <input data-field="new.color" type="color" value="${escapeAttr(state.newElement.color || defaults.color)}" />
        </label>
      </div>
      <div class="field-pair">
        <label class="field">
          Width (m)
          <input data-field="new.width" type="number" min="0.1" step="0.1" value="${state.newElement.width}" />
        </label>
        <label class="field">
          Depth (m)
          <input data-field="new.height" type="number" min="0.1" step="0.1" value="${state.newElement.height}" />
        </label>
      </div>
      ${
        state.elementDraft.length
          ? `<div class="draft-actions">
              <strong>${state.elementDraft.length} polygon points drafted</strong>
              <button data-action="finish-element-poly" type="button" ${state.elementDraft.length < 3 ? "disabled" : ""}>Finish object</button>
              <button data-action="cancel-element-draft" type="button">Cancel</button>
            </div>`
          : ""
      }
    </section>
  `;
}

function renderBackgroundPanel(project: Project) {
  const background = project.background;

  return `
    <section class="panel-section">
      <div class="section-heading">
        <span>Background</span>
        ${background ? `<button data-action="clear-background" type="button">Clear</button>` : ""}
      </div>
      <label class="file-drop">
        Import survey image
        <input id="background-file" type="file" accept="image/*" />
      </label>
      ${
        background
          ? `<div class="background-fields">
              <small>${escapeHtml(background.name)}</small>
              <div class="field-pair">
                <label class="field">X (m)<input data-field="background.x" type="number" step="0.1" value="${roundM(background.x)}" /></label>
                <label class="field">Y (m)<input data-field="background.y" type="number" step="0.1" value="${roundM(background.y)}" /></label>
              </div>
              <div class="field-pair">
                <label class="field">Width (m)<input data-field="background.width" type="number" min="0.1" step="0.1" value="${roundM(background.width)}" /></label>
                <label class="field">Height (m)<input data-field="background.height" type="number" min="0.1" step="0.1" value="${roundM(background.height)}" /></label>
              </div>
              <label class="field">Opacity<input data-field="background.opacity" type="range" min="0.05" max="1" step="0.05" value="${background.opacity}" /></label>
              <label class="checkline"><input data-field="background.locked" type="checkbox" ${background.locked ? "checked" : ""} /> Lock background</label>
            </div>`
          : `<p class="hint">Import a sketch or survey, then calibrate its real-world width and height in meters.</p>`
      }
    </section>
  `;
}

function renderInspector(
  project: Project,
  selectedElement: SiteElement | undefined,
  selectedNote: PlanNote | undefined,
  warnings: Map<string, string[]>,
) {
  if (project.selected.kind === "land") {
    const area = polygonArea(project.land);
    const bbox = bboxFromPoints(project.land);

    return `
      <section class="panel-section inspector">
        <div class="section-heading">
          <span>Land Boundary</span>
          <button data-action="clear-selection" type="button">Close</button>
        </div>
        <p class="metric-line">Area <strong>${formatMeters(area, "m2")}</strong></p>
        <p class="metric-line">Envelope <strong>${formatMeters(bbox.maxX - bbox.minX)} x ${formatMeters(
          bbox.maxY - bbox.minY,
        )}</strong></p>
        <div class="field-pair">
          <label class="field">Front of land
            <select data-field="land.frontEdge">
              ${renderFrontEdgeOptions(project)}
            </select>
          </label>
          <label class="field">Rotation (deg)
            <input data-field="land.rotation" type="number" step="1" value="${roundM(project.landRotation)}" />
          </label>
        </div>
        <p class="hint">Use rotation to match the land boundary to the north arrow. Positive values rotate clockwise.</p>
        ${renderLandMeasurementControls(project)}
        <p class="hint">Drag vertex handles on the canvas to edit the freeform boundary.</p>
        <button class="danger-button" data-action="remove-land" type="button">Remove land</button>
      </section>
    `;
  }

  if (selectedElement) {
    const elementWarnings = warnings.get(selectedElement.id) ?? [];
    const bbox = bboxOfElement(selectedElement);
    const center = elementCenter(selectedElement);

    return `
      <section class="panel-section inspector">
        <div class="section-heading">
          <span>Selected Object</span>
          <button data-action="delete-selection" type="button">Delete</button>
        </div>
        ${elementWarnings.map((warning) => `<p class="warning">${escapeHtml(warning)}</p>`).join("")}
        <label class="field">Label<input data-field="element.name" type="text" value="${escapeAttr(selectedElement.name)}" /></label>
        <div class="field-pair">
          <label class="field">Category
            <select data-field="element.category">
              ${Object.entries(categoryDefaults)
                .map(
                  ([value, option]) =>
                    `<option value="${value}" ${
                      value === selectedElement.category ? "selected" : ""
                    }>${escapeHtml(option.label)}</option>`,
                )
                .join("")}
            </select>
          </label>
          <label class="field">Color<input data-field="element.color" type="color" value="${escapeAttr(
            selectedElement.color,
          )}" /></label>
        </div>
        <div class="field-pair">
          <label class="field">X (m)<input data-field="element.x" type="number" step="0.1" value="${roundM(center.x)}" /></label>
          <label class="field">Y (m)<input data-field="element.y" type="number" step="0.1" value="${roundM(center.y)}" /></label>
        </div>
        ${
          selectedElement.kind === "rect"
            ? `<div class="field-pair">
                <label class="field">Width (m)<input data-field="element.width" type="number" min="0.1" step="0.1" value="${roundM(
                  selectedElement.width,
                )}" /></label>
                <label class="field">Depth (m)<input data-field="element.height" type="number" min="0.1" step="0.1" value="${roundM(
                  selectedElement.height,
                )}" /></label>
              </div>
              <label class="field">Rotation (deg)<input data-field="element.rotation" type="number" step="1" value="${roundM(
                selectedElement.rotation,
              )}" /></label>`
            : `<p class="metric-line">Envelope <strong>${formatMeters(bbox.maxX - bbox.minX)} x ${formatMeters(
                bbox.maxY - bbox.minY,
              )}</strong></p>
              <p class="hint">Drag polygon vertex handles on the canvas for point editing.</p>`
        }
        <label class="checkline"><input data-field="element.showInLegend" type="checkbox" ${
          selectedElement.showInLegend ? "checked" : ""
        } /> Show in legend</label>
      </section>
    `;
  }

  if (selectedNote) {
    return `
      <section class="panel-section inspector">
        <div class="section-heading">
          <span>Selected Note</span>
          <button data-action="delete-selection" type="button">Delete</button>
        </div>
        <label class="field">Text<input data-field="note.text" type="text" value="${escapeAttr(selectedNote.text)}" /></label>
        <div class="field-pair">
          <label class="field">X (m)<input data-field="note.x" type="number" step="0.1" value="${roundM(selectedNote.x)}" /></label>
          <label class="field">Y (m)<input data-field="note.y" type="number" step="0.1" value="${roundM(selectedNote.y)}" /></label>
        </div>
      </section>
    `;
  }

  return `
    <section class="panel-section inspector">
      <div class="section-heading">
        <span>Inspector</span>
      </div>
      <p class="hint">Select land, an object, or a note to edit exact values. Distance labels appear for the selected object when overlays are enabled.</p>
      <div class="export-options">
        <label class="field">Paper
          <select data-field="printSize">
            <option value="A4" ${state.printSize === "A4" ? "selected" : ""}>A4</option>
            <option value="A3" ${state.printSize === "A3" ? "selected" : ""}>A3</option>
          </select>
        </label>
        <label class="checkline"><input data-field="includeSummary" type="checkbox" ${
          state.includeSummary ? "checked" : ""
        } /> Include summary page</label>
      </div>
    </section>
  `;
}

function renderLegend(project: Project) {
  const legendItems = project.elements.filter((element) => element.showInLegend);

  return `
    <section class="panel-section legend">
      <div class="section-heading">
        <span>Legend</span>
        <small>${legendItems.length} items</small>
      </div>
      ${
        legendItems.length
          ? legendItems
              .map(
                (element) => `
                  <button class="legend-row" data-select-element="${element.id}" type="button">
                    <span class="swatch" style="--swatch:${escapeAttr(element.color)}"></span>
                    <span>${escapeHtml(element.name)}</span>
                    <small>${escapeHtml(categoryDefaults[element.category].label)}</small>
                  </button>
                `,
              )
              .join("")
          : `<p class="hint">No legend-visible objects yet.</p>`
      }
    </section>
  `;
}

function renderFrontEdgeOptions(project: Project) {
  const options = [`<option value="" ${project.landFrontEdge === null ? "selected" : ""}>No front marker</option>`];

  for (let index = 0; index < project.land.length; index += 1) {
    const start = project.land[index];
    const end = project.land[(index + 1) % project.land.length];
    const label = `Edge ${index + 1}: ${roundM(start.x).toFixed(1)}, ${roundM(start.y).toFixed(
      1,
    )} to ${roundM(end.x).toFixed(1)}, ${roundM(end.y).toFixed(1)}`;
    options.push(
      `<option value="${index}" ${project.landFrontEdge === index ? "selected" : ""}>${escapeHtml(label)}</option>`,
    );
  }

  return options.join("");
}

function renderLandMeasurementControls(project: Project) {
  const rows = landSegments(project.land)
    .map(([start, end], index) => {
      const label = `Side ${index + 1}${project.landFrontEdge === index ? " (front)" : ""}`;
      return `
        <label class="field edge-length-field">
          <span>${escapeHtml(label)}</span>
          <input data-field="land.edgeLength.${index}" type="number" min="0.1" step="0.1" value="${roundM(
            distance(start, end),
          ).toFixed(1)}" />
        </label>
      `;
    })
    .join("");

  return `
    <div class="edge-lengths">
      <div class="subheading">Side lengths (m)</div>
      ${rows}
      <p class="hint">Rectangular land keeps opposite sides matched. Freeform side edits move that side's end point.</p>
    </div>
  `;
}

function renderSvg(project: Project, warnings: Map<string, string[]>) {
  const view = currentView(project);
  const selected = project.selected;
  const landPoints = pointsAttr(project.land);
  const hasLand = project.land.length >= 3;
  const viewRight = view.x + view.width;
  const viewBottom = view.y + view.height;
  const selectedElement =
    selected.kind === "element" ? project.elements.find((element) => element.id === selected.id) : undefined;

  return `
    <svg id="plan-svg" viewBox="${view.x} ${view.y} ${view.width} ${view.height}" preserveAspectRatio="xMidYMin meet" role="img" aria-label="Land planning canvas">
      <defs>
        <pattern id="minor-grid" width="1" height="1" patternUnits="userSpaceOnUse">
          <path d="M 1 0 L 0 0 0 1" fill="none" stroke="#b6c3b8" stroke-width="0.045" />
        </pattern>
        <pattern id="major-grid" width="5" height="5" patternUnits="userSpaceOnUse">
          <rect width="5" height="5" fill="url(#minor-grid)" />
          <path d="M 5 0 L 0 0 0 5" fill="none" stroke="#758579" stroke-width="0.11" />
        </pattern>
        <filter id="paper-shadow" x="-20%" y="-20%" width="140%" height="140%">
          <feDropShadow dx="0" dy="0.5" stdDeviation="0.8" flood-color="#1d2b28" flood-opacity="0.12" />
        </filter>
      </defs>
      <rect class="grid-surface" x="${view.x}" y="${view.y}" width="${view.width}" height="${view.height}" fill="url(#major-grid)" />
      <rect class="interaction-catcher" x="${view.x}" y="${view.y}" width="${view.width}" height="${view.height}" />
      ${project.background ? renderBackground(project.background) : ""}
      ${
        hasLand
          ? `<polygon class="land-boundary ${selected.kind === "land" ? "selected" : ""}" data-land="true" points="${landPoints}" />`
          : renderEmptyCanvasHint(view)
      }
      ${hasLand ? renderLandFront(project) : ""}
      ${renderLandVertexHandles(project)}
      ${renderLandDraft()}
      ${project.elements.map((element) => renderElement(project, element, warnings)).join("")}
      ${renderElementDraft()}
      ${project.notes.map((note) => renderNote(note, selected.kind === "note" && selected.id === note.id)).join("")}
      ${selectedElement && project.showDistances ? renderDistanceOverlay(project, selectedElement) : ""}
      ${renderSvgLegend(project, view)}
      ${renderNorthArrow(viewRight - 5, view.y + 4)}
      <g class="scale-chip">
        <rect x="${view.x + 1}" y="${viewBottom - 4.8}" width="18" height="3.2" rx="0.7" />
        <line x1="${view.x + 2.4}" y1="${viewBottom - 2.8}" x2="${view.x + 12.4}" y2="${viewBottom - 2.8}" />
        <text x="${view.x + 2.4}" y="${viewBottom - 1.55}">10 m - zoom ${project.zoom.toFixed(2)}x</text>
      </g>
    </svg>
  `;
}

function renderBackground(background: BackgroundImageState) {
  return `
    <image
      class="background-image ${background.locked ? "locked" : ""}"
      href="${escapeAttr(background.dataUrl)}"
      x="${background.x}"
      y="${background.y}"
      width="${background.width}"
      height="${background.height}"
      opacity="${background.opacity}"
      preserveAspectRatio="none"
    />
  `;
}

function renderEmptyCanvasHint(view: { x: number; y: number; width: number; height: number }) {
  return `
    <g class="empty-canvas-hint">
      <text x="${view.x + view.width / 2}" y="${view.y + view.height / 2 - 1.2}" text-anchor="middle">Blank canvas</text>
      <text x="${view.x + view.width / 2}" y="${view.y + view.height / 2 + 0.4}" text-anchor="middle">Draw land or use the rectangular preset to begin.</text>
    </g>
  `;
}

function renderLandFront(project: Project) {
  const segment = frontEdgeSegment(project);

  if (!segment) {
    return "";
  }

  const [start, end] = segment;
  const midpoint = { x: (start.x + end.x) / 2, y: (start.y + end.y) / 2 };
  const vector = { x: end.x - start.x, y: end.y - start.y };
  const length = Math.max(Math.hypot(vector.x, vector.y), 0.001);
  const normal = { x: -vector.y / length, y: vector.x / length };
  const label = { x: midpoint.x + normal.x * 1.4, y: midpoint.y + normal.y * 1.4 };

  return `
    <g class="land-front-marker">
      <line x1="${start.x}" y1="${start.y}" x2="${end.x}" y2="${end.y}" />
      <text x="${label.x}" y="${label.y}" text-anchor="middle">FRONT</text>
    </g>
  `;
}

function renderLandVertexHandles(project: Project) {
  if (project.selected.kind !== "land" && state.tool !== "draw-land") {
    return "";
  }

  return project.land
    .map(
      (point, index) =>
        `<circle class="vertex-handle land" data-vertex-kind="land" data-index="${index}" cx="${point.x}" cy="${point.y}" r="0.38" />`,
    )
    .join("");
}

function renderLandDraft() {
  if (!state.landDraft.length) {
    return "";
  }

  return `
    <polyline class="draft-line" points="${pointsAttr(state.landDraft)}" />
    ${state.landDraft
      .map(
        (point, index) =>
          `<circle class="draft-point" cx="${point.x}" cy="${point.y}" r="${index === 0 ? "0.48" : "0.34"}" />`,
      )
      .join("")}
  `;
}

function renderElementDraft() {
  if (!state.elementDraft.length) {
    return "";
  }

  return `
    <polyline class="draft-line object" points="${pointsAttr(state.elementDraft)}" />
    ${state.elementDraft.map((point) => `<circle class="draft-point object" cx="${point.x}" cy="${point.y}" r="0.34" />`).join("")}
  `;
}

function renderElement(project: Project, element: SiteElement, warnings: Map<string, string[]>) {
  const isSelected = project.selected.kind === "element" && project.selected.id === element.id;
  const warningClass = warnings.has(element.id) ? "has-warning" : "";
  const points = pointsAttr(pointsOfElement(element));
  const center = elementCenter(element);
  const bbox = bboxOfElement(element);

  return `
    <g class="site-element ${isSelected ? "selected" : ""} ${warningClass}" data-element-id="${element.id}">
      <polygon points="${points}" fill="${escapeAttr(element.color)}" />
      <text x="${center.x}" y="${center.y}" text-anchor="middle">${escapeHtml(element.name)}</text>
      ${element.kind === "rect" ? `<text class="dimension-label" x="${center.x}" y="${bbox.maxY + 1.4}" text-anchor="middle">${formatMeters(element.width)} x ${formatMeters(element.height)}</text>` : ""}
      ${isSelected ? renderElementVertexHandles(element) : ""}
    </g>
  `;
}

function renderSvgLegend(project: Project, view: { x: number; y: number; width: number; height: number }) {
  const items = project.elements.filter((element) => element.showInLegend).slice(0, 8);

  if (!items.length) {
    return "";
  }

  const x = view.x + view.width - 20;
  const y = view.y + view.height - 3.8 - items.length * 1.65;

  return `
    <g class="svg-legend" transform="translate(${x} ${y})">
      <rect width="18" height="${2.1 + items.length * 1.65}" rx="0.75" />
      <text class="svg-legend-title" x="1" y="1.35">Legend</text>
      ${items
        .map(
          (element, index) => `
            <rect class="svg-legend-swatch" x="1" y="${2.05 + index * 1.65}" width="0.9" height="0.9" fill="${escapeAttr(
              element.color,
            )}" />
            <text x="2.35" y="${2.82 + index * 1.65}">${escapeHtml(element.name)}</text>
          `,
        )
        .join("")}
    </g>
  `;
}

function renderElementVertexHandles(element: SiteElement) {
  if (element.kind !== "poly") {
    return "";
  }

  return element.points
    .map(
      (point, index) =>
        `<circle class="vertex-handle object" data-vertex-kind="element" data-element-id="${element.id}" data-index="${index}" cx="${point.x}" cy="${point.y}" r="0.36" />`,
    )
    .join("");
}

function renderNote(note: PlanNote, selected: boolean) {
  return `
    <g class="plan-note ${selected ? "selected" : ""}" data-note-id="${note.id}">
      <circle cx="${note.x}" cy="${note.y}" r="0.42" />
      <text x="${note.x + 0.75}" y="${note.y + 0.3}">${escapeHtml(note.text)}</text>
    </g>
  `;
}

function renderDistanceOverlay(project: Project, element: SiteElement) {
  const box = bboxOfElement(element);
  const center = elementCenter(element);
  const landBox = project.land.length >= 3 ? bboxFromPoints(project.land) : undefined;
  const distances = landBox
    ? [
        { label: `Left ${formatMeters(box.minX - landBox.minX)}`, x: (landBox.minX + box.minX) / 2, y: center.y },
        { label: `Right ${formatMeters(landBox.maxX - box.maxX)}`, x: (landBox.maxX + box.maxX) / 2, y: center.y },
        { label: `Bottom ${formatMeters(box.minY - landBox.minY)}`, x: center.x, y: (landBox.minY + box.minY) / 2 },
        { label: `Top ${formatMeters(landBox.maxY - box.maxY)}`, x: center.x, y: (landBox.maxY + box.maxY) / 2 },
      ]
    : [];
  const nearest = nearestElementDistance(project, element);

  if (!distances.length && !nearest) {
    return "";
  }

  return `
    <g class="distance-overlay">
      ${distances
        .map(
          (item) => `
            <text x="${item.x}" y="${item.y}" text-anchor="middle">${escapeHtml(item.label)}</text>
          `,
        )
        .join("")}
      ${
        nearest
          ? `<line x1="${center.x}" y1="${center.y}" x2="${nearest.center.x}" y2="${nearest.center.y}" />
             <text x="${(center.x + nearest.center.x) / 2}" y="${(center.y + nearest.center.y) / 2 - 0.8}" text-anchor="middle">Nearest ${formatMeters(
               nearest.distance,
             )}</text>`
          : ""
      }
    </g>
  `;
}

function renderNorthArrow(x: number, y: number) {
  return `
    <g class="north-arrow" transform="translate(${x} ${y})">
      <path d="M0,-2.3 L1.05,1.45 L0,0.8 L-1.05,1.45 Z" />
      <text x="0" y="3.1" text-anchor="middle">N</text>
    </g>
  `;
}

function renderPrintSummary(project: Project, warnings: Map<string, string[]>) {
  return `
    <section class="print-summary" aria-label="PDF summary page">
      <h1>${escapeHtml(project.name)}</h1>
      <p>Generated from Land Layout Planner. Units are meters and displayed measurements round to 0.1 m.</p>
      <div class="summary-grid">
        <div><strong>Land area</strong><span>${formatMeters(polygonArea(project.land), "m2")}</span></div>
        <div><strong>Objects</strong><span>${project.elements.length}</span></div>
        <div><strong>Notes</strong><span>${project.notes.length}</span></div>
        <div><strong>Warnings</strong><span>${Array.from(warnings.values()).flat().length}</span></div>
      </div>
      <table>
        <thead><tr><th>Object</th><th>Type</th><th>Size</th><th>Status</th></tr></thead>
        <tbody>
          ${project.elements
            .map((element) => {
              const bbox = bboxOfElement(element);
              const status = (warnings.get(element.id) ?? []).join("; ") || "OK";
              return `<tr><td>${escapeHtml(element.name)}</td><td>${escapeHtml(
                categoryDefaults[element.category].label,
              )}</td><td>${formatMeters(bbox.maxX - bbox.minX)} x ${formatMeters(
                bbox.maxY - bbox.minY,
              )}</td><td>${escapeHtml(status)}</td></tr>`;
            })
            .join("")}
        </tbody>
      </table>
    </section>
  `;
}

function toolButton(tool: Tool, label: string) {
  return `<button class="${state.tool === tool ? "active" : ""}" data-tool="${tool}" type="button">${label}</button>`;
}

function toolStatus() {
  if (state.tool === "draw-land") {
    return "Draw land: click boundary points, double-click to finish.";
  }

  if (state.tool === "place-rect") {
    return "Place rectangle: click the canvas to position the object.";
  }

  if (state.tool === "place-poly") {
    return "Place freeform object: click points, double-click to finish.";
  }

  if (state.tool === "place-note") {
    return "Place note: click the canvas to add the annotation.";
  }

  if (state.tool === "pan") {
    return "Pan: drag the canvas to move the view.";
  }

  return "Select: drag objects or edit exact values in the inspector.";
}

function handleClick(event: MouseEvent) {
  const target = event.target as Element;
  const tab = target.closest<HTMLButtonElement>("[data-project-id]");
  const tool = target.closest<HTMLButtonElement>("[data-tool]");
  const legendSelect = target.closest<HTMLButtonElement>("[data-select-element]");
  const actionButton = target.closest<HTMLButtonElement>("[data-action]");
  const action = actionButton?.dataset.action;
  const actionProjectId = actionButton?.dataset.projectActionId;

  if (tab?.dataset.projectId) {
    state.activeProjectId = tab.dataset.projectId;
    state.tool = "select";
    state.landDraft = [];
    state.elementDraft = [];
    persist();
    render();
    return;
  }

  if (tool?.dataset.tool) {
    state.tool = tool.dataset.tool as Tool;
    if (state.tool !== "draw-land") {
      state.landDraft = [];
    }
    if (state.tool !== "place-poly") {
      state.elementDraft = [];
    }
    render();
    return;
  }

  if (legendSelect?.dataset.selectElement) {
    updateActiveProject((project) => {
      project.selected = { kind: "element", id: legendSelect.dataset.selectElement ?? "" };
    });
    return;
  }

  if (!action) {
    return;
  }

  const project = activeProject();

  if (action === "toggle-history") {
    state.showHistory = !state.showHistory;
    persist();
    render();
    return;
  }

  if (action === "open-history-project" && actionProjectId) {
    state.activeProjectId = actionProjectId;
    state.showHistory = false;
    state.tool = "select";
    state.landDraft = [];
    state.elementDraft = [];
    persist();
    render();
    return;
  }

  if (action === "duplicate-history-project" && actionProjectId) {
    duplicateProject(actionProjectId);
    return;
  }

  if (action === "delete-history-project" && actionProjectId) {
    deleteProjectById(actionProjectId);
    return;
  }

  if (action === "new-project") {
    const newProject = createProject(`Option ${String.fromCharCode(65 + state.projects.length)}`, false);
    state.projects.push(newProject);
    state.activeProjectId = newProject.id;
    state.tool = "select";
    persist();
    render();
    return;
  }

  if (action === "duplicate-project") {
    duplicateProject(project.id);
    return;
  }

  if (action === "delete-project") {
    deleteActiveProject();
    return;
  }

  if (action === "undo") {
    undo();
    return;
  }

  if (action === "redo") {
    redo();
    return;
  }

  if (action === "rect-land") {
    updateActiveProject(
      (draft) => {
        applyRectangularLand(draft, state.landRectWidth, state.landRectDepth);
        draft.selected = { kind: "land" };
      },
      true,
    );
    return;
  }

  if (action === "remove-land") {
    state.landDraft = [];
    updateActiveProject(
      (draft) => {
        draft.land = [];
        draft.landFrontEdge = null;
        draft.landRotation = 0;
        draft.selected = { kind: "none" };
      },
      true,
    );
    return;
  }

  if (action === "finish-land") {
    finishLandDraft();
    return;
  }

  if (action === "cancel-land-draft") {
    state.landDraft = [];
    state.tool = "select";
    render();
    return;
  }

  if (action === "place-element") {
    state.elementDraft = [];
    state.tool = state.newElement.kind === "rect" ? "place-rect" : "place-poly";
    render();
    return;
  }

  if (action === "finish-element-poly") {
    finishElementDraft();
    return;
  }

  if (action === "cancel-element-draft") {
    state.elementDraft = [];
    state.tool = "select";
    render();
    return;
  }

  if (action === "clear-selection") {
    updateActiveProject((draft) => {
      draft.selected = { kind: "none" };
    });
    return;
  }

  if (action === "delete-selection") {
    deleteSelection();
    return;
  }

  if (action === "zoom-reset") {
    updateActiveProject((draft) => {
      draft.zoom = 1;
      draft.pan = { ...DEFAULT_PAN };
    });
    return;
  }

  if (action === "export-json") {
    downloadProject(project);
    return;
  }

  if (action === "print") {
    window.print();
    return;
  }

  if (action === "clear-background") {
    updateActiveProject((draft) => {
      delete draft.background;
    }, true);
  }
}

function handleInput(event: Event) {
  const target = event.target as HTMLInputElement | HTMLSelectElement;
  const field = target.dataset.field;

  if (!field) {
    return;
  }

  if (field.startsWith("new.")) {
    updateNewElement(field, target);
    return;
  }

  if (field.startsWith("landRect.")) {
    updateLandRectangleDraft(field, target);
    persist();
    return;
  }

  if (isDeferredLandGeometryField(field)) {
    return;
  }

  if (field === "noteText") {
    state.noteText = target.value;
    persist();
    return;
  }

  if (field === "printSize") {
    state.printSize = target.value === "A3" ? "A3" : "A4";
    persist();
    render();
    return;
  }

  if (field === "includeSummary") {
    state.includeSummary = (target as HTMLInputElement).checked;
    persist();
    render();
    return;
  }

  updateActiveProject((project) => {
    updateProjectField(project, field, target);
  }, isRecordedField(field));
}

function handleChange(event: Event) {
  const target = event.target as HTMLInputElement;
  const field = target.dataset.field;

  if (field?.startsWith("landRect.")) {
    updateLandRectangleDraft(field, target);
    persist();
    render();
    return;
  }

  if (field && isDeferredLandGeometryField(field)) {
    updateActiveProject((project) => {
      updateProjectField(project, field, target);
    }, true);
    return;
  }

  if (target.id === "background-file" && target.files?.[0]) {
    importBackgroundImage(target.files[0]);
    return;
  }

  if (target.id === "import-project" && target.files?.[0]) {
    importProjectFile(target.files[0]);
  }
}

function handlePointerDown(event: PointerEvent) {
  if ((event.target as Element).closest("[data-action], [data-tool], input, select, button, label")) {
    return;
  }

  const target = event.target as Element;
  const svg = target.closest<SVGSVGElement>("#plan-svg");
  if (!svg) {
    return;
  }

  const point = maybeSnap(clientToWorld(event), event);
  const vertex = target.closest<SVGElement>("[data-vertex-kind]");
  const elementNode = target.closest<SVGElement>("[data-element-id]");
  const noteNode = target.closest<SVGElement>("[data-note-id]");
  const isLand = Boolean(target.closest("[data-land]"));

  event.preventDefault();

  if (vertex?.dataset.vertexKind === "land") {
    pushHistory(activeProject());
    activeProject().selected = { kind: "land" };
    drag = { type: "land-vertex", index: Number(vertex.dataset.index ?? 0) };
    render();
    return;
  }

  if (vertex?.dataset.vertexKind === "element" && vertex.dataset.elementId) {
    pushHistory(activeProject());
    activeProject().selected = { kind: "element", id: vertex.dataset.elementId };
    drag = {
      type: "element-vertex",
      id: vertex.dataset.elementId,
      index: Number(vertex.dataset.index ?? 0),
    };
    render();
    return;
  }

  if (elementNode?.dataset.elementId && state.tool === "select") {
    updateActiveProject((project) => {
      project.selected = { kind: "element", id: elementNode.dataset.elementId ?? "" };
    }, false);
    pushHistory(activeProject());
    drag = { type: "element", id: elementNode.dataset.elementId, lastPoint: point };
    return;
  }

  if (noteNode?.dataset.noteId && state.tool === "select") {
    updateActiveProject((project) => {
      project.selected = { kind: "note", id: noteNode.dataset.noteId ?? "" };
    }, false);
    pushHistory(activeProject());
    drag = { type: "note", id: noteNode.dataset.noteId, lastPoint: point };
    return;
  }

  if (isLand && state.tool === "select") {
    updateActiveProject((project) => {
      project.selected = { kind: "land" };
    });
    return;
  }

  if (state.tool === "pan") {
    drag = {
      type: "pan",
      startClient: { x: event.clientX, y: event.clientY },
      startPan: { ...activeProject().pan },
    };
    return;
  }

  if (state.tool === "draw-land") {
    if (state.landDraft.length >= 3 && distance(point, state.landDraft[0]) < 0.8) {
      finishLandDraft();
      return;
    }

    state.landDraft.push(point);
    render();
    return;
  }

  if (state.tool === "place-rect") {
    const created = createRectElement(point);
    updateActiveProject(
      (project) => {
        project.elements.push(created);
        project.selected = { kind: "element", id: created.id };
      },
      true,
    );
    state.tool = "select";
    return;
  }

  if (state.tool === "place-poly") {
    if (state.elementDraft.length >= 3 && distance(point, state.elementDraft[0]) < 0.8) {
      finishElementDraft();
      return;
    }

    state.elementDraft.push(point);
    render();
    return;
  }

  if (state.tool === "place-note") {
    const note: PlanNote = {
      id: uid("note"),
      text: state.noteText.trim() || "Site note",
      x: point.x,
      y: point.y,
    };
    updateActiveProject(
      (project) => {
        project.notes.push(note);
        project.selected = { kind: "note", id: note.id };
      },
      true,
    );
    state.tool = "select";
    return;
  }

  updateActiveProject((project) => {
    project.selected = { kind: "none" };
  });
}

function handlePointerMove(event: PointerEvent) {
  if (drag.type === "none") {
    return;
  }

  const project = activeProject();

  if (drag.type === "pan") {
    const start = clientPointToWorld(drag.startClient);
    const current = clientPointToWorld({ x: event.clientX, y: event.clientY });
    const dx = current.x - start.x;
    const dy = current.y - start.y;
    project.pan = { x: drag.startPan.x - dx, y: drag.startPan.y - dy };
    project.updatedAt = new Date().toISOString();
    persist();
    render();
    return;
  }

  const nextPoint = maybeSnap(clientToWorld(event), event);

  if (drag.type === "land-vertex") {
    project.land[drag.index] = nextPoint;
    project.updatedAt = new Date().toISOString();
    persist();
    render();
    return;
  }

  if (drag.type === "element-vertex") {
    const element = project.elements.find((candidate) => candidate.id === drag.id);
    if (element?.kind === "poly") {
      element.points[drag.index] = nextPoint;
      project.updatedAt = new Date().toISOString();
      persist();
      render();
    }
    return;
  }

  const delta = { x: nextPoint.x - drag.lastPoint.x, y: nextPoint.y - drag.lastPoint.y };

  if (drag.type === "element") {
    const element = project.elements.find((candidate) => candidate.id === drag.id);
    if (!element) {
      return;
    }

    moveElement(element, delta);
    drag = { ...drag, lastPoint: nextPoint };
    project.updatedAt = new Date().toISOString();
    persist();
    render();
    return;
  }

  if (drag.type === "note") {
    const note = project.notes.find((candidate) => candidate.id === drag.id);
    if (!note) {
      return;
    }

    note.x += delta.x;
    note.y += delta.y;
    drag = { ...drag, lastPoint: nextPoint };
    project.updatedAt = new Date().toISOString();
    persist();
    render();
  }
}

function handlePointerUp() {
  if (drag.type !== "none") {
    drag = { type: "none" };
    render();
  }
}

function handleDoubleClick(event: MouseEvent) {
  const target = event.target as Element;

  if (!target.closest("#plan-svg")) {
    return;
  }

  if (state.tool === "draw-land" && state.landDraft.length >= 3) {
    finishLandDraft();
  }

  if (state.tool === "place-poly" && state.elementDraft.length >= 3) {
    finishElementDraft();
  }
}

function handleKeyDown(event: KeyboardEvent) {
  if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === "z" && event.shiftKey) {
    event.preventDefault();
    redo();
    return;
  }

  if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === "z") {
    event.preventDefault();
    undo();
    return;
  }

  if (event.key === "Escape") {
    state.landDraft = [];
    state.elementDraft = [];
    state.tool = "select";
    drag = { type: "none" };
    render();
  }
}

function updateNewElement(field: string, target: HTMLInputElement | HTMLSelectElement) {
  if (field === "new.category" && isElementCategory(target.value)) {
    const defaults = categoryDefaults[target.value];
    state.newElement.category = target.value;
    state.newElement.name = defaults.label;
    state.newElement.width = defaults.width;
    state.newElement.height = defaults.height;
    state.newElement.color = defaults.color;
  }

  if (field === "new.name") {
    state.newElement.name = target.value;
  }

  if (field === "new.kind") {
    state.newElement.kind = target.value === "poly" ? "poly" : "rect";
  }

  if (field === "new.color") {
    state.newElement.color = target.value;
  }

  if (field === "new.width") {
    state.newElement.width = clamp(toNumber(target.value, state.newElement.width), 0.1, 500);
  }

  if (field === "new.height") {
    state.newElement.height = clamp(toNumber(target.value, state.newElement.height), 0.1, 500);
  }

  persist();
  render();
}

function updateLandRectangleDraft(field: string, target: HTMLInputElement | HTMLSelectElement) {
  const min = 0.1;
  const max = 10000;

  if (field === "landRect.width") {
    state.landRectWidth = clamp(toNumber(target.value, state.landRectWidth), min, max);
  }

  if (field === "landRect.depth") {
    state.landRectDepth = clamp(toNumber(target.value, state.landRectDepth), min, max);
  }
}

function updateProjectField(project: Project, field: string, target: HTMLInputElement | HTMLSelectElement) {
  if (field === "project.name") {
    project.name = target.value || "Untitled project";
  }

  if (field === "project.zoom") {
    project.zoom = clamp(toNumber(target.value, project.zoom), 0.45, 4);
  }

  if (field === "project.snapping") {
    project.snapping = (target as HTMLInputElement).checked;
  }

  if (field === "project.showDistances") {
    project.showDistances = (target as HTMLInputElement).checked;
  }

  if (field.startsWith("element.")) {
    updateSelectedElementField(project, field, target);
  }

  if (field === "land.frontEdge") {
    project.landFrontEdge =
      target.value === "" ? null : normalizeLandFrontEdge(toNumber(target.value, -1), project.land);
  }

  if (field === "land.rotation") {
    rotateLandTo(project, toNumber(target.value, project.landRotation));
  }

  if (field.startsWith("land.edgeLength.")) {
    const index = toNumber(field.replace("land.edgeLength.", ""), -1);
    resizeLandEdge(project, index, toNumber(target.value, landEdgeLength(project, index)));
  }

  if (field.startsWith("note.")) {
    updateSelectedNoteField(project, field, target);
  }

  if (field.startsWith("background.")) {
    updateBackgroundField(project, field, target);
  }
}

function updateSelectedElementField(project: Project, field: string, target: HTMLInputElement | HTMLSelectElement) {
  if (project.selected.kind !== "element") {
    return;
  }

  const element = project.elements.find((candidate) => candidate.id === project.selected.id);
  if (!element) {
    return;
  }

  if (field === "element.name") {
    element.name = target.value || "Untitled object";
  }

  if (field === "element.category" && isElementCategory(target.value)) {
    element.category = target.value;
  }

  if (field === "element.color") {
    element.color = target.value;
  }

  if (field === "element.showInLegend") {
    element.showInLegend = (target as HTMLInputElement).checked;
  }

  if (field === "element.x" || field === "element.y") {
    const center = elementCenter(element);
    const nextCenter = {
      x: field === "element.x" ? toNumber(target.value, center.x) : center.x,
      y: field === "element.y" ? toNumber(target.value, center.y) : center.y,
    };
    const delta = { x: nextCenter.x - center.x, y: nextCenter.y - center.y };
    moveElement(element, delta);
  }

  if (element.kind !== "rect") {
    return;
  }

  if (field === "element.width") {
    element.width = clamp(toNumber(target.value, element.width), 0.1, 500);
  }

  if (field === "element.height") {
    element.height = clamp(toNumber(target.value, element.height), 0.1, 500);
  }

  if (field === "element.rotation") {
    element.rotation = toNumber(target.value, element.rotation);
  }
}

function updateSelectedNoteField(project: Project, field: string, target: HTMLInputElement | HTMLSelectElement) {
  if (project.selected.kind !== "note") {
    return;
  }

  const note = project.notes.find((candidate) => candidate.id === project.selected.id);
  if (!note) {
    return;
  }

  if (field === "note.text") {
    note.text = target.value || "Site note";
  }

  if (field === "note.x") {
    note.x = toNumber(target.value, note.x);
  }

  if (field === "note.y") {
    note.y = toNumber(target.value, note.y);
  }
}

function updateBackgroundField(project: Project, field: string, target: HTMLInputElement | HTMLSelectElement) {
  if (!project.background) {
    return;
  }

  if (field === "background.x") {
    project.background.x = toNumber(target.value, project.background.x);
  }

  if (field === "background.y") {
    project.background.y = toNumber(target.value, project.background.y);
  }

  if (field === "background.width") {
    project.background.width = clamp(toNumber(target.value, project.background.width), 0.1, 500);
  }

  if (field === "background.height") {
    project.background.height = clamp(toNumber(target.value, project.background.height), 0.1, 500);
  }

  if (field === "background.opacity") {
    project.background.opacity = clamp(toNumber(target.value, project.background.opacity), 0.05, 1);
  }

  if (field === "background.locked") {
    project.background.locked = (target as HTMLInputElement).checked;
  }
}

function isRecordedField(field: string) {
  return !["project.zoom", "project.name", "project.snapping", "project.showDistances"].includes(field);
}

function isDeferredLandGeometryField(field: string) {
  return field === "land.rotation" || field.startsWith("land.edgeLength.");
}

function finishLandDraft() {
  if (state.landDraft.length < 3) {
    return;
  }

  const points = [...state.landDraft];
  updateActiveProject(
    (project) => {
      project.land = points;
      project.landFrontEdge = null;
      project.landRotation = 0;
      project.selected = { kind: "land" };
    },
    true,
  );
  state.landDraft = [];
  state.tool = "select";
}

function finishElementDraft() {
  if (state.elementDraft.length < 3) {
    return;
  }

  const element: PolyElement = {
    id: uid("element"),
    kind: "poly",
    category: state.newElement.category,
    name: state.newElement.name || categoryDefaults[state.newElement.category].label,
    points: [...state.elementDraft],
    rotation: 0,
    color: state.newElement.color,
    showInLegend: true,
  };

  updateActiveProject(
    (project) => {
      project.elements.push(element);
      project.selected = { kind: "element", id: element.id };
    },
    true,
  );
  state.elementDraft = [];
  state.tool = "select";
}

function createRectElement(center: Point): RectElement {
  return {
    id: uid("element"),
    kind: "rect",
    category: state.newElement.category,
    name: state.newElement.name || categoryDefaults[state.newElement.category].label,
    x: center.x,
    y: center.y,
    width: state.newElement.width,
    height: state.newElement.height,
    rotation: 0,
    color: state.newElement.color,
    showInLegend: true,
  };
}

function deleteSelection() {
  updateActiveProject(
    (project) => {
      const selected = project.selected;
      if (selected.kind === "element") {
        project.elements = project.elements.filter((element) => element.id !== selected.id);
      }
      if (selected.kind === "note") {
        project.notes = project.notes.filter((note) => note.id !== selected.id);
      }
      project.selected = { kind: "none" };
    },
    true,
  );
}

function duplicateProject(projectId: string) {
  const source = state.projects.find((candidate) => candidate.id === projectId);

  if (!source) {
    return;
  }

  const duplicate = restoreProject(snapshotProject(source));
  duplicate.id = uid("project");
  duplicate.name = `${source.name} copy`;
  duplicate.createdAt = new Date().toISOString();
  duplicate.updatedAt = duplicate.createdAt;
  duplicate.history = [];
  duplicate.future = [];
  duplicate.selected = { kind: "none" };
  state.projects.push(duplicate);
  state.activeProjectId = duplicate.id;
  state.showHistory = false;
  persist();
  render();
}

function deleteActiveProject() {
  deleteProjectById(activeProject().id);
}

function deleteProjectById(projectId: string) {
  const project = state.projects.find((candidate) => candidate.id === projectId);

  if (!project) {
    return;
  }

  const confirmed = window.confirm(`Delete ${project.name}? This removes the option from this browser.`);

  if (!confirmed) {
    return;
  }

  const currentIndex = state.projects.findIndex((candidate) => candidate.id === project.id);
  const deletingActive = state.activeProjectId === project.id;

  if (state.projects.length <= 1) {
    const replacement = createProject("Option A", false);
    state.projects = [replacement];
    state.activeProjectId = replacement.id;
  } else {
    state.projects = state.projects.filter((candidate) => candidate.id !== project.id);
    const nextIndex = Math.min(Math.max(currentIndex, 0), state.projects.length - 1);
    if (deletingActive) {
      state.activeProjectId = state.projects[nextIndex].id;
    }
  }

  state.tool = "select";
  state.landDraft = [];
  state.elementDraft = [];
  drag = { type: "none" };
  persist();
  render();
}

function importBackgroundImage(file: File) {
  const reader = new FileReader();
  reader.addEventListener("load", () => {
    const dataUrl = String(reader.result ?? "");
    const image = new Image();
    image.addEventListener("load", () => {
      const project = activeProject();
      const landBox = bboxFromPoints(project.land);
      const width = Math.max(landBox.maxX - landBox.minX, 30);
      const height = width * (image.naturalHeight / Math.max(image.naturalWidth, 1));

      updateActiveProject(
        (draft) => {
          draft.background = {
            dataUrl,
            name: file.name,
            x: landBox.minX,
            y: landBox.minY,
            width,
            height,
            opacity: 0.45,
            locked: true,
          };
        },
        true,
      );
    });
    image.src = dataUrl;
  });
  reader.readAsDataURL(file);
}

function importProjectFile(file: File) {
  const reader = new FileReader();
  reader.addEventListener("load", () => {
    try {
      const parsed = JSON.parse(String(reader.result ?? "{}"));
      const project = normalizeProject(parsed.project ?? parsed);
      project.id = uid("project");
      project.name = `${project.name} imported`;
      state.projects.push(project);
      state.activeProjectId = project.id;
      state.tool = "select";
      persist();
      render();
    } catch {
      alert("The selected project file could not be imported.");
    }
  });
  reader.readAsText(file);
}

function downloadProject(project: Project) {
  const payload = {
    version: 1,
    exportedAt: new Date().toISOString(),
    project: snapshotProject(project),
  };
  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `${slugify(project.name)}.json`;
  link.click();
  URL.revokeObjectURL(url);
}

function undo() {
  const project = activeProject();
  const previous = project.history.pop();
  if (!previous) {
    return;
  }

  project.future.push(snapshotProject(project));
  const restored = restoreProject(previous);
  restored.history = project.history;
  restored.future = project.future;
  replaceActiveProject(restored);
}

function redo() {
  const project = activeProject();
  const next = project.future.pop();
  if (!next) {
    return;
  }

  project.history.push(snapshotProject(project));
  const restored = restoreProject(next);
  restored.history = project.history;
  restored.future = project.future;
  replaceActiveProject(restored);
}

function updateActiveProject(mutator: (project: Project) => void, record = false) {
  const project = activeProject();

  if (record) {
    pushHistory(project);
  }

  mutator(project);
  project.updatedAt = new Date().toISOString();
  persist();
  render();
}

function pushHistory(project: Project) {
  project.history.push(snapshotProject(project));
  project.history = project.history.slice(-MAX_HISTORY);
  project.future = [];
}

function replaceActiveProject(project: Project) {
  const index = state.projects.findIndex((candidate) => candidate.id === state.activeProjectId);
  if (index >= 0) {
    state.projects[index] = project;
  }
  persist();
  render();
}

function snapshotProject(project: Project): ProjectSnapshot {
  return JSON.parse(
    JSON.stringify({
      id: project.id,
      name: project.name,
      createdAt: project.createdAt,
      updatedAt: project.updatedAt,
      zoom: project.zoom,
      pan: project.pan,
      land: project.land,
      landFrontEdge: project.landFrontEdge,
      landRotation: project.landRotation,
      elements: project.elements,
      notes: project.notes,
      background: project.background,
      showDistances: project.showDistances,
      snapping: project.snapping,
      selected: project.selected,
    }),
  ) as ProjectSnapshot;
}

function restoreProject(snapshot: ProjectSnapshot): Project {
  const restored = JSON.parse(JSON.stringify(snapshot)) as ProjectSnapshot;
  return {
    ...restored,
    landRotation: toNumber(restored.landRotation, 0),
    history: [],
    future: [],
  } as Project;
}

function activeProject() {
  return state.projects.find((project) => project.id === state.activeProjectId) ?? state.projects[0];
}

function persist() {
  localStorage.setItem(
    STORAGE_KEY,
    JSON.stringify({
      ...state,
      projects: state.projects.map((project) => ({
        ...snapshotProject(project),
        history: project.history,
        future: project.future,
      })),
    }),
  );
}

function currentView(project: Project) {
  return {
    x: project.pan.x,
    y: project.pan.y,
    width: WORLD_WIDTH / project.zoom,
    height: WORLD_HEIGHT / project.zoom,
  };
}

function clientToWorld(event: PointerEvent | MouseEvent): Point {
  return clientPointToWorld({ x: event.clientX, y: event.clientY });
}

function clientPointToWorld(clientPoint: Point): Point {
  const svg = document.querySelector<SVGSVGElement>("#plan-svg");
  const project = activeProject();
  const view = currentView(project);

  if (!svg) {
    return { x: 0, y: 0 };
  }

  const matrix = svg.getScreenCTM();
  if (matrix) {
    const point = svg.createSVGPoint();
    point.x = clientPoint.x;
    point.y = clientPoint.y;
    const transformed = point.matrixTransform(matrix.inverse());
    return { x: transformed.x, y: transformed.y };
  }

  const rect = svg.getBoundingClientRect();
  return {
    x: view.x + ((clientPoint.x - rect.left) / rect.width) * view.width,
    y: view.y + ((clientPoint.y - rect.top) / rect.height) * view.height,
  };
}

function maybeSnap(point: Point, event: PointerEvent | MouseEvent): Point {
  const project = activeProject();
  if (!project.snapping || event.altKey) {
    return point;
  }

  return snapPoint(project, point);
}

function snapPoint(project: Project, point: Point) {
  const threshold = 0.55;
  let best = {
    point: {
      x: Math.round(point.x * 2) / 2,
      y: Math.round(point.y * 2) / 2,
    },
    distance: 0.4,
  };

  const candidates = [...project.land, ...project.elements.flatMap(pointsOfElement)];

  for (const candidate of candidates) {
    const candidateDistance = distance(point, candidate);
    if (candidateDistance < threshold && candidateDistance < best.distance) {
      best = { point: candidate, distance: candidateDistance };
    }
  }

  for (const [start, end] of landSegments(project.land)) {
    const projection = projectPointToSegment(point, start, end);
    const projectedDistance = distance(point, projection);
    if (projectedDistance < threshold && projectedDistance < best.distance) {
      best = { point: projection, distance: projectedDistance };
    }
  }

  return best.point;
}

function moveElement(element: SiteElement, delta: Point) {
  if (element.kind === "rect") {
    element.x += delta.x;
    element.y += delta.y;
    return;
  }

  element.points = element.points.map((point) => ({
    x: point.x + delta.x,
    y: point.y + delta.y,
  }));
}

function pointsOfElement(element: SiteElement) {
  if (element.kind === "poly") {
    return element.points;
  }

  const halfWidth = element.width / 2;
  const halfHeight = element.height / 2;
  const corners = [
    { x: element.x - halfWidth, y: element.y - halfHeight },
    { x: element.x + halfWidth, y: element.y - halfHeight },
    { x: element.x + halfWidth, y: element.y + halfHeight },
    { x: element.x - halfWidth, y: element.y + halfHeight },
  ];

  return corners.map((point) => rotatePoint(point, { x: element.x, y: element.y }, element.rotation));
}

function elementCenter(element: SiteElement) {
  if (element.kind === "rect") {
    return { x: element.x, y: element.y };
  }

  return centroid(element.points);
}

function bboxOfElement(element: SiteElement) {
  return bboxFromPoints(pointsOfElement(element));
}

function getWarnings(project: Project) {
  const warnings = new Map<string, string[]>();
  const land = project.land;

  for (const element of project.elements) {
    const elementWarnings: string[] = [];
    const points = pointsOfElement(element);

    if (land.length >= 3 && points.some((point) => !pointInPolygon(point, land))) {
      elementWarnings.push("Object leaves the land boundary.");
    }

    for (const other of project.elements) {
      if (other.id === element.id) {
        continue;
      }

      if (boxesOverlap(bboxOfElement(element), bboxOfElement(other))) {
        elementWarnings.push(`Overlaps ${other.name}.`);
        break;
      }
    }

    if (elementWarnings.length) {
      warnings.set(element.id, elementWarnings);
    }
  }

  return warnings;
}

function nearestElementDistance(project: Project, element: SiteElement) {
  const sourceBox = bboxOfElement(element);
  const sourceCenter = elementCenter(element);
  let nearest: { distance: number; center: Point } | undefined;

  for (const other of project.elements) {
    if (other.id === element.id) {
      continue;
    }

    const otherBox = bboxOfElement(other);
    const dx = Math.max(otherBox.minX - sourceBox.maxX, sourceBox.minX - otherBox.maxX, 0);
    const dy = Math.max(otherBox.minY - sourceBox.maxY, sourceBox.minY - otherBox.maxY, 0);
    const gap = Math.sqrt(dx * dx + dy * dy);

    if (!nearest || gap < nearest.distance) {
      nearest = { distance: gap, center: elementCenter(other) };
    }
  }

  return nearest;
}

function bboxFromPoints(points: Point[]): BBox {
  if (!points.length) {
    return { minX: 0, minY: 0, maxX: 0, maxY: 0 };
  }

  return points.reduce(
    (box, point) => ({
      minX: Math.min(box.minX, point.x),
      minY: Math.min(box.minY, point.y),
      maxX: Math.max(box.maxX, point.x),
      maxY: Math.max(box.maxY, point.y),
    }),
    { minX: points[0].x, minY: points[0].y, maxX: points[0].x, maxY: points[0].y },
  );
}

function boxesOverlap(a: BBox, b: BBox) {
  return a.minX < b.maxX && a.maxX > b.minX && a.minY < b.maxY && a.maxY > b.minY;
}

function pointInPolygon(point: Point, polygon: Point[]) {
  let inside = false;

  for (let i = 0, j = polygon.length - 1; i < polygon.length; j = i++) {
    const current = polygon[i];
    const previous = polygon[j];
    const intersects =
      current.y > point.y !== previous.y > point.y &&
      point.x < ((previous.x - current.x) * (point.y - current.y)) / (previous.y - current.y) + current.x;

    if (intersects) {
      inside = !inside;
    }
  }

  return inside;
}

function polygonArea(points: Point[]) {
  if (points.length < 3) {
    return 0;
  }

  let total = 0;
  for (let index = 0; index < points.length; index += 1) {
    const current = points[index];
    const next = points[(index + 1) % points.length];
    total += current.x * next.y - next.x * current.y;
  }

  return Math.abs(total / 2);
}

function landSegments(points: Point[]) {
  if (points.length < 2) {
    return [];
  }

  return points.map((point, index) => [point, points[(index + 1) % points.length]] as const);
}

function frontEdgeSegment(project: Project) {
  if (project.landFrontEdge === null || project.land.length < 3) {
    return undefined;
  }

  const start = project.land[project.landFrontEdge];
  const end = project.land[(project.landFrontEdge + 1) % project.land.length];

  if (!start || !end) {
    return undefined;
  }

  return [start, end] as const;
}

function landEdgeLength(project: Project, index: number) {
  const segment = landSegments(project.land)[index];
  return segment ? distance(segment[0], segment[1]) : 0.1;
}

function applyRectangularLand(project: Project, width: number, depth: number) {
  const nextWidth = normalizePositiveNumber(width, 23);
  const nextDepth = normalizePositiveNumber(depth, 45.5);
  const rotation = project.land.length >= 3 ? toNumber(project.landRotation, 0) : 0;
  const center = project.land.length >= 3 ? centroid(project.land) : { x: nextWidth / 2, y: nextDepth / 2 };

  project.land = rectanglePoints(center, nextWidth, nextDepth, rotation);
  project.landFrontEdge = normalizeLandFrontEdge(project.landFrontEdge, project.land) ?? 2;
  project.landRotation = rotation;
}

function resizeLandEdge(project: Project, index: number, nextLength: number) {
  const normalizedIndex = normalizeLandFrontEdge(index, project.land);

  if (normalizedIndex === null) {
    return;
  }

  const length = normalizePositiveNumber(nextLength, landEdgeLength(project, normalizedIndex));
  if (resizeRectangularLandEdge(project, normalizedIndex, length)) {
    return;
  }

  const start = project.land[normalizedIndex];
  const endIndex = (normalizedIndex + 1) % project.land.length;
  const end = project.land[endIndex];
  const currentLength = distance(start, end);

  if (currentLength <= 0) {
    return;
  }

  const unit = {
    x: (end.x - start.x) / currentLength,
    y: (end.y - start.y) / currentLength,
  };

  project.land[endIndex] = {
    x: start.x + unit.x * length,
    y: start.y + unit.y * length,
  };
}

function resizeRectangularLandEdge(project: Project, index: number, length: number) {
  const axes = rectangularLandAxes(project.land);

  if (!axes) {
    return false;
  }

  const nextWidth = index % 2 === 0 ? length : axes.width;
  const nextDepth = index % 2 === 1 ? length : axes.depth;
  project.land = rectanglePointsFromAxes(axes.center, axes.unitU, axes.unitV, nextWidth, nextDepth);
  return true;
}

function rectangularLandAxes(points: Point[]) {
  if (points.length !== 4) {
    return undefined;
  }

  const [p0, p1, p2, p3] = points;
  const widthA = distance(p0, p1);
  const widthB = distance(p3, p2);
  const depthA = distance(p1, p2);
  const depthB = distance(p0, p3);
  const width = (widthA + widthB) / 2;
  const depth = (depthA + depthB) / 2;

  if (width <= 0.05 || depth <= 0.05) {
    return undefined;
  }

  const unitU = unitVector(p0, p1);
  const unitV = unitVector(p0, p3);
  const relativeTolerance = 0.03;
  const oppositeSidesMatch =
    Math.abs(widthA - widthB) <= Math.max(width * relativeTolerance, 0.05) &&
    Math.abs(depthA - depthB) <= Math.max(depth * relativeTolerance, 0.05);
  const adjacentSidesSquare = Math.abs(dot(unitU, unitV)) <= 0.03;
  const expectedP2 = { x: p1.x + p3.x - p0.x, y: p1.y + p3.y - p0.y };
  const closesRectangle = distance(expectedP2, p2) <= Math.max(Math.max(width, depth) * relativeTolerance, 0.1);

  if (!oppositeSidesMatch || !adjacentSidesSquare || !closesRectangle) {
    return undefined;
  }

  return {
    center: centroid(points),
    unitU,
    unitV,
    width,
    depth,
  };
}

function rectanglePoints(center: Point, width: number, depth: number, rotation: number) {
  const halfWidth = width / 2;
  const halfDepth = depth / 2;
  const points = [
    { x: center.x - halfWidth, y: center.y - halfDepth },
    { x: center.x + halfWidth, y: center.y - halfDepth },
    { x: center.x + halfWidth, y: center.y + halfDepth },
    { x: center.x - halfWidth, y: center.y + halfDepth },
  ];

  return points.map((point) => roundPoint(rotatePoint(point, center, rotation)));
}

function rectanglePointsFromAxes(center: Point, unitU: Point, unitV: Point, width: number, depth: number) {
  const halfU = { x: unitU.x * width * 0.5, y: unitU.y * width * 0.5 };
  const halfV = { x: unitV.x * depth * 0.5, y: unitV.y * depth * 0.5 };

  return [
    { x: center.x - halfU.x - halfV.x, y: center.y - halfU.y - halfV.y },
    { x: center.x + halfU.x - halfV.x, y: center.y + halfU.y - halfV.y },
    { x: center.x + halfU.x + halfV.x, y: center.y + halfU.y + halfV.y },
    { x: center.x - halfU.x + halfV.x, y: center.y - halfU.y + halfV.y },
  ].map(roundPoint);
}

function rotateLandTo(project: Project, nextRotation: number) {
  const currentRotation = toNumber(project.landRotation, 0);
  const delta = nextRotation - currentRotation;
  project.landRotation = nextRotation;

  if (project.land.length < 3 || Math.abs(delta) < 0.001) {
    return;
  }

  const center = centroid(project.land);
  project.land = project.land.map((point) => {
    const rotated = rotatePoint(point, center, delta);
    return roundPoint(rotated);
  });
}

function projectPointToSegment(point: Point, start: Point, end: Point) {
  const dx = end.x - start.x;
  const dy = end.y - start.y;
  const lengthSquared = dx * dx + dy * dy;

  if (!lengthSquared) {
    return start;
  }

  const t = clamp(((point.x - start.x) * dx + (point.y - start.y) * dy) / lengthSquared, 0, 1);
  return { x: start.x + t * dx, y: start.y + t * dy };
}

function rotatePoint(point: Point, center: Point, degrees: number) {
  const theta = (degrees * Math.PI) / 180;
  const cos = Math.cos(theta);
  const sin = Math.sin(theta);
  const dx = point.x - center.x;
  const dy = point.y - center.y;

  return {
    x: center.x + dx * cos - dy * sin,
    y: center.y + dx * sin + dy * cos,
  };
}

function unitVector(start: Point, end: Point) {
  const length = distance(start, end);

  if (length <= 0) {
    return { x: 1, y: 0 };
  }

  return {
    x: (end.x - start.x) / length,
    y: (end.y - start.y) / length,
  };
}

function dot(a: Point, b: Point) {
  return a.x * b.x + a.y * b.y;
}

function roundPoint(point: Point) {
  return { x: trimNumber(point.x), y: trimNumber(point.y) };
}

function centroid(points: Point[]) {
  if (!points.length) {
    return { x: 0, y: 0 };
  }

  return {
    x: points.reduce((total, point) => total + point.x, 0) / points.length,
    y: points.reduce((total, point) => total + point.y, 0) / points.length,
  };
}

function normalizePoint(input: unknown, fallback: Point) {
  const point = input as Partial<Point>;
  return {
    x: toNumber(point?.x, fallback.x),
    y: toNumber(point?.y, fallback.y),
  };
}

function normalizePan(input: unknown, fallback: Point) {
  const pan = normalizePoint(input, fallback);

  if (pointsMatch(pan, LEGACY_DEFAULT_PAN)) {
    return { ...DEFAULT_PAN };
  }

  return pan;
}

function normalizePoints(input: unknown, fallback: Point[]) {
  if (!Array.isArray(input)) {
    return fallback;
  }

  return input
    .map((point) => normalizePoint(point, { x: Number.NaN, y: Number.NaN }))
    .filter((point) => Number.isFinite(point.x) && Number.isFinite(point.y));
}

function normalizeLandFrontEdge(input: unknown, land: Point[]) {
  const index = Number(input);

  if (!Number.isInteger(index) || index < 0 || index >= land.length || land.length < 3) {
    return null;
  }

  return index;
}

function distance(a: Point, b: Point) {
  return Math.hypot(a.x - b.x, a.y - b.y);
}

function pointsMatch(a: Point, b: Point) {
  return Math.abs(a.x - b.x) < 0.001 && Math.abs(a.y - b.y) < 0.001;
}

function pointsAttr(points: Point[]) {
  return points.map((point) => `${trimNumber(point.x)},${trimNumber(point.y)}`).join(" ");
}

function formatMeters(value: number, unit = "m") {
  return `${roundM(Math.max(value, 0)).toFixed(1)} ${unit}`;
}

function roundM(value: number) {
  return Math.round(value * 10) / 10;
}

function trimNumber(value: number) {
  return Number(value.toFixed(3));
}

function toNumber(value: unknown, fallback: number) {
  if (typeof value === "string") {
    const normalized = value.trim().replace(",", ".");
    if (!normalized) {
      return fallback;
    }
    const next = Number(normalized);
    return Number.isFinite(next) ? next : fallback;
  }

  const next = Number(value);
  return Number.isFinite(next) ? next : fallback;
}

function clamp(value: number, min: number, max: number) {
  return Math.min(Math.max(value, min), max);
}

function uid(prefix: string) {
  return `${prefix}-${Math.random().toString(36).slice(2, 10)}`;
}

function slugify(value: string) {
  return value
    .toLowerCase()
    .trim()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-|-$/g, "");
}

function formatDateTime(value: string) {
  const date = new Date(value);

  if (Number.isNaN(date.getTime())) {
    return "Unknown date";
  }

  return date.toLocaleString(undefined, {
    dateStyle: "medium",
    timeStyle: "short",
  });
}

function isElementCategory(value: unknown): value is ElementCategory {
  return typeof value === "string" && Object.prototype.hasOwnProperty.call(categoryDefaults, value);
}

function escapeHtml(value: string | number) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function escapeAttr(value: string | number) {
  return escapeHtml(value);
}
