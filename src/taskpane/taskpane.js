/* global Office, PowerPoint, FormData, btoa, document, fetch, localStorage, navigator, window */

import "core-js/stable";
import "regenerator-runtime/runtime";

const BASE_URL = "https://gloomeries.github.io/Slide-library";
const STORAGE_KEYS = {
  favorites: "slidebrary:favorites",
  recent: "slidebrary:recent",
};

const materials = [
  {
    id: "title",
    title: "Титульный слайд",
    product: "Общее",
    type: "Презентация",
    format: "16×9",
    style: "Светлая тема",
    tags: ["титул", "обложка", "начало"],
    preview: `${BASE_URL}/assets/previews/title.png`,
    template: `${BASE_URL}/assets/templates/title.pptx`,
  },
  {
    id: "executive_summary",
    title: "Executive Summary",
    product: "Общее",
    type: "Презентация",
    format: "16×9",
    style: "Светлая тема",
    tags: ["итоги", "резюме", "summary"],
    preview: `${BASE_URL}/assets/previews/executive_summary.png`,
    template: `${BASE_URL}/assets/templates/executive_summary.pptx`,
  },
  {
    id: "market_analysis",
    title: "Анализ рынка",
    product: "Аналитика",
    type: "Презентация",
    format: "16×9",
    style: "Светлая тема",
    tags: ["рынок", "аналитика", "исследование"],
    preview: `${BASE_URL}/assets/previews/market_analysis.png`,
    template: `${BASE_URL}/assets/templates/market_analysis.pptx`,
  },
  {
    id: "marketing_plan",
    title: "Маркетинговый план",
    product: "Маркетинг",
    type: "Презентация",
    format: "16×9",
    style: "Светлая тема",
    tags: ["маркетинг", "план", "кампания"],
    preview: `${BASE_URL}/assets/previews/marketing_plan.png`,
    template: `${BASE_URL}/assets/templates/marketing_plan.pptx`,
  },
  {
    id: "prototypes",
    title: "Описание прототипов",
    product: "Продукт",
    type: "Презентация",
    format: "16×9",
    style: "Темная тема",
    tags: ["прототип", "продукт", "интерфейс"],
    preview: `${BASE_URL}/assets/previews/prototypes.png`,
    template: `${BASE_URL}/assets/templates/prototypes.pptx`,
  },
  {
    id: "risk_analysis",
    title: "Матрица рисков",
    product: "Аналитика",
    type: "Презентация",
    format: "16×9",
    style: "Светлая тема",
    tags: ["риски", "матрица", "оценка"],
    preview: `${BASE_URL}/assets/previews/risk_analysis.png`,
    template: `${BASE_URL}/assets/templates/risk_matrix.pptx`,
  },
  {
    id: "roadmap",
    title: "Дорожная карта",
    product: "Продукт",
    type: "Презентация",
    format: "16×9",
    style: "Темная тема",
    tags: ["roadmap", "этапы", "план"],
    preview: `${BASE_URL}/assets/previews/roadmap.png`,
    template: `${BASE_URL}/assets/templates/roadmap.pptx`,
  },
  {
    id: "target_audience",
    title: "Целевая аудитория",
    product: "Маркетинг",
    type: "Презентация",
    format: "16×9",
    style: "Светлая тема",
    tags: ["аудитория", "сегменты", "пользователи"],
    preview: `${BASE_URL}/assets/previews/target_audience.png`,
    template: `${BASE_URL}/assets/templates/target_audience.pptx`,
  },
  {
    id: "team",
    title: "Команда проекта",
    product: "Общее",
    type: "Презентация",
    format: "16×9",
    style: "Темная тема",
    tags: ["команда", "роли", "участники"],
    preview: `${BASE_URL}/assets/previews/team.png`,
    template: `${BASE_URL}/assets/templates/team.pptx`,
  },
  {
    id: "business_process",
    title: "Бизнес-процесс",
    product: "Бизнес",
    type: "Презентация",
    format: "16×9",
    style: "Светлая тема",
    tags: ["процесс", "схема", "бизнес"],
    preview: `${BASE_URL}/assets/previews/business_process.png`,
    template: `${BASE_URL}/assets/templates/business_process.pptx`,
  },
];

const sections = [
  { id: "favorites", label: "Избранное", icon: "heart" },
  { id: "presentations", label: "Презентации", icon: "presentation" },
  { id: "photos", label: "Фотографии", icon: "camera" },
  { id: "illustrations", label: "Иллюстрации", icon: "image" },
  { id: "icons", label: "Иконки", icon: "icons" },
  { id: "logos", label: "Логотипы", icon: "logo" },
  { id: "templates", label: "Шаблоны", icon: "templates" },
  { id: "assistant", label: "ИИ-ассистент", icon: "assistant" },
];

const sectionHints = {
  favorites: "Избранные материалы",
  presentations: "Выберите макет для вашего слайда",
  photos: "Выберите фотографию для вашего слайда",
  illustrations: "Выберите иллюстрацию для вашего слайда",
  icons: "Выберите иконку для вашего слайда",
  logos: "Выберите логотип для вашего слайда",
  templates: "Выберите макет для вашего слайда",
  assistant: "Создайте материал с помощью ИИ-ассистента",
};

const state = {
  section: "presentations",
  tab: "public",
  query: "",
  view: "grid",
  sort: "default",
  filters: { type: "", product: "", format: "", style: "" },
  favorites: new Set(readStoredArray(STORAGE_KEYS.favorites)),
  recent: readStoredArray(STORAGE_KEYS.recent),
  selected: new Set(),
  previewId: null,
  loading: true,
  inserting: false,
  sidebarCollapsed: false,
};

const elements = {};
let toastTimer;

function readStoredArray(key) {
  try {
    const value = JSON.parse(localStorage.getItem(key) || "[]");
    return Array.isArray(value) ? value : [];
  } catch {
    return [];
  }
}

function writeStoredArray(key, value) {
  try {
    localStorage.setItem(key, JSON.stringify(value));
  } catch {
    showToast("Не удалось сохранить данные на этом устройстве");
  }
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function getSectionIcon(name) {
  const icons = {
    heart:
      '<path d="M12 20.5S4 16 4 9.8A4.3 4.3 0 0 1 12 7.6a4.3 4.3 0 0 1 8 2.2c0 6.2-8 10.7-8 10.7Z"/>',
    presentation:
      '<rect x="4" y="4" width="16" height="13" rx="2"/><path d="M8 20l4-3 4 3M8 9h8M8 12h5"/>',
    camera: '<path d="M4 8.5h3l1.4-2h7.2l1.4 2h3v10H4z"/><circle cx="12" cy="13.5" r="3.2"/>',
    image:
      '<rect x="4" y="4" width="16" height="16" rx="2"/><circle cx="9" cy="9" r="1.5"/><path d="m6.5 17 4-4 2.5 2.5 2-2 2.5 3"/>',
    icons:
      '<rect x="4" y="4" width="6" height="6" rx="1.5"/><circle cx="17" cy="7" r="3"/><path d="m7 14-3 6h6z"/><path d="m14 14 6 6m0-6-6 6"/>',
    logo: '<rect x="4" y="4" width="16" height="16" rx="4"/><path d="m8 17 4-10 4 10M9.5 13h5"/>',
    templates: '<rect x="3.5" y="4" width="17" height="16" rx="2"/><path d="M3.5 10h17M10 10v10"/>',
    assistant:
      '<rect x="5" y="5" width="14" height="14" rx="3"/><path d="M9 2v3m6-3v3M9 19v3m6-3v3M2 9h3m-3 6h3m14-6h3m-3 6h3M9 10.5h6M9 14h4"/>',
  };
  return `<svg class="nav-icon" viewBox="0 0 24 24" aria-hidden="true">${icons[name] || icons.templates}</svg>`;
}

function cacheElements() {
  [
    "sectionNav",
    "collapseButton",
    "searchInput",
    "filterButton",
    "filterBadge",
    "sortButton",
    "viewButton",
    "shareButton",
    "moreButton",
    "closeAppButton",
    "sectionHint",
    "statusRegion",
    "library",
    "clearSelectionButton",
    "insertSelectionButton",
    "filterOverlay",
    "closeFiltersButton",
    "resetFiltersButton",
    "filterForm",
    "previewOverlay",
    "closePreviewButton",
    "previewImage",
    "previewTitle",
    "previewFormat",
    "previewProduct",
    "previewInsertButton",
    "toast",
  ].forEach((id) => {
    elements[id] = document.getElementById(id);
  });
}

function renderNavigation() {
  elements.sectionNav.innerHTML = sections
    .map(
      (section) => `
        <button
          class="nav-button${state.section === section.id ? " is-active" : ""}"
          type="button"
          data-section="${section.id}"
          aria-label="${escapeHtml(section.label)}"
          title="${escapeHtml(section.label)}"
        >
          <span class="nav-icon-wrap">${getSectionIcon(section.icon)}</span>
          <span class="nav-label">${escapeHtml(section.label)}</span>
        </button>
      `
    )
    .join("");
}

function populateFilterOptions() {
  ["type", "product", "format", "style"].forEach((field) => {
    const select = elements.filterForm.elements[field];
    const options = [...new Set(materials.map((item) => item[field]))].sort((a, b) =>
      a.localeCompare(b, "ru")
    );
    select.insertAdjacentHTML(
      "beforeend",
      options
        .map((option) => `<option value="${escapeHtml(option)}">${escapeHtml(option)}</option>`)
        .join("")
    );
  });
}

function getVisibleMaterials() {
  if (state.tab === "personal") return [];

  let result = [...materials];
  if (state.section === "favorites") {
    result = result.filter((item) => state.favorites.has(item.id));
  }
  if (!["favorites", "presentations", "templates"].includes(state.section)) result = [];

  const normalizedQuery = state.query.trim().toLocaleLowerCase("ru");
  if (normalizedQuery) {
    result = result.filter((item) =>
      [item.title, item.product, item.type, item.style, ...item.tags]
        .join(" ")
        .toLocaleLowerCase("ru")
        .includes(normalizedQuery)
    );
  }

  Object.entries(state.filters).forEach(([field, value]) => {
    if (value) result = result.filter((item) => item[field] === value);
  });

  if (state.sort === "title-asc") {
    result.sort((a, b) => a.title.localeCompare(b.title, "ru"));
  } else if (state.sort === "title-desc") {
    result.sort((a, b) => b.title.localeCompare(a.title, "ru"));
  }
  return result;
}

function getEmptyMessage() {
  if (state.tab === "personal") {
    return "Личная библиотека появится в следующем этапе. Сейчас все материалы находятся во вкладке «Публичное».";
  }
  if (state.section === "favorites")
    return "В избранном пока ничего нет. Нажмите на сердечко у нужного материала.";
  if (state.section === "photos")
    return "Раздел фотографий готов. Материалы добавим на следующем этапе.";
  if (state.section === "illustrations")
    return "Раздел иллюстраций готов. Материалы добавим на следующем этапе.";
  if (state.section === "icons")
    return "Раздел иконок готов. Материалы добавим на следующем этапе.";
  if (state.section === "logos")
    return "Раздел логотипов готов. Материалы добавим на следующем этапе.";
  if (state.section === "assistant") return "ИИ-ассистент появится на следующем этапе разработки.";
  if (state.query || getActiveFilterCount())
    return "По вашему запросу ничего не найдено. Попробуйте изменить поиск или фильтры.";
  return "В библиотеке пока нет материалов.";
}

function renderStatus() {
  if (state.loading) {
    elements.statusRegion.innerHTML = `
      <div class="loading-grid" aria-label="Загрузка материалов">
        <div class="skeleton"></div><div class="skeleton"></div>
        <div class="skeleton"></div><div class="skeleton"></div>
      </div>`;
    return;
  }
  elements.statusRegion.innerHTML = "";
}

function renderLibrary() {
  renderStatus();
  elements.library.className = `library-grid${state.view === "list" ? " is-list" : ""}`;
  if (state.loading) {
    elements.library.innerHTML = "";
    return;
  }

  const visibleMaterials = getVisibleMaterials();
  if (!visibleMaterials.length) {
    elements.library.innerHTML = `<div class="empty-state">${escapeHtml(getEmptyMessage())}</div>`;
    return;
  }

  elements.library.innerHTML = visibleMaterials
    .map((item) => {
      const isFavorite = state.favorites.has(item.id);
      const isSelected = state.selected.has(item.id);
      return `
        <article
          class="material-card${isSelected ? " is-selected" : ""}"
          data-id="${item.id}"
          tabindex="0"
          aria-label="Открыть предпросмотр: ${escapeHtml(item.title)}"
        >
          <img class="card-preview" src="${item.preview}" alt="Превью: ${escapeHtml(item.title)}" loading="lazy" />
          <p class="card-title">${escapeHtml(item.title)}</p>
          <button
            class="favorite-button${isFavorite ? " is-active" : ""}"
            type="button"
            data-action="favorite"
            aria-label="${isFavorite ? "Удалить из избранного" : "Добавить в избранное"}"
            title="${isFavorite ? "Удалить из избранного" : "Добавить в избранное"}"
          >${isFavorite ? "♥" : "♡"}</button>
          <button
            class="add-button"
            type="button"
            data-action="select"
            aria-label="${isSelected ? "Убрать из выбранного" : "Выбрать для вставки"}"
            title="${isSelected ? "Убрать из выбранного" : "Выбрать для вставки"}"
          >${isSelected ? "✓" : "+"}</button>
        </article>`;
    })
    .join("");
}

function renderControls() {
  elements.sectionHint.textContent = sectionHints[state.section];
  elements.library.dataset.view = state.view;
  document.getElementById("app").classList.toggle("is-sidebar-collapsed", state.sidebarCollapsed);
  elements.collapseButton.setAttribute(
    "aria-label",
    state.sidebarCollapsed ? "Развернуть меню" : "Свернуть меню"
  );
  elements.collapseButton.title = state.sidebarCollapsed ? "Развернуть меню" : "Свернуть меню";
  elements.viewButton.title = state.view === "grid" ? "Показать списком" : "Показать плиткой";
  elements.viewButton.setAttribute("aria-label", elements.viewButton.title);
  elements.sortButton.title =
    state.sort === "default"
      ? "Сортировка: по умолчанию"
      : state.sort === "title-asc"
        ? "Сортировка: А–Я"
        : "Сортировка: Я–А";
  elements.insertSelectionButton.disabled = state.selected.size === 0 || state.inserting;
  elements.clearSelectionButton.disabled = state.selected.size === 0 || state.inserting;
  elements.insertSelectionButton.textContent = state.inserting
    ? "Добавляем…"
    : state.selected.size > 0
      ? `Добавить (${state.selected.size})`
      : "Добавить";

  const activeFilters = getActiveFilterCount();
  elements.filterBadge.hidden = activeFilters === 0;
  elements.filterBadge.textContent = String(activeFilters);

  document.querySelectorAll(".tab").forEach((tab) => {
    const isActive = tab.dataset.tab === state.tab;
    tab.classList.toggle("is-active", isActive);
    tab.setAttribute("aria-selected", String(isActive));
  });
}

function render() {
  renderNavigation();
  renderControls();
  renderLibrary();
}

function getActiveFilterCount() {
  return Object.values(state.filters).filter(Boolean).length;
}

function selectSection(sectionId) {
  state.section = sectionId;
  state.selected.clear();
  render();
}

function toggleFavorite(id) {
  if (state.favorites.has(id)) state.favorites.delete(id);
  else state.favorites.add(id);
  writeStoredArray(STORAGE_KEYS.favorites, [...state.favorites]);
  render();
}

function toggleSelection(id) {
  if (state.selected.has(id)) state.selected.delete(id);
  else state.selected.add(id);
  render();
}

function openPreview(id) {
  const item = materials.find((candidate) => candidate.id === id);
  if (!item) return;
  state.previewId = id;
  elements.previewImage.src = item.preview;
  elements.previewImage.alt = `Превью: ${item.title}`;
  elements.previewTitle.textContent = item.title;
  elements.previewFormat.textContent = item.format;
  elements.previewProduct.textContent = item.product;
  elements.previewOverlay.hidden = false;
  elements.closePreviewButton.focus();
}

function closePreview() {
  elements.previewOverlay.hidden = true;
  state.previewId = null;
}

function openFilters() {
  Object.entries(state.filters).forEach(([field, value]) => {
    elements.filterForm.elements[field].value = value;
  });
  elements.filterOverlay.hidden = false;
  elements.closeFiltersButton.focus();
}

function closeFilters() {
  elements.filterOverlay.hidden = true;
  elements.filterButton.focus();
}

function resetFilters() {
  state.filters = { type: "", product: "", format: "", style: "" };
  elements.filterForm.reset();
  render();
}

function cycleSort() {
  const order = ["default", "title-asc", "title-desc"];
  state.sort = order[(order.indexOf(state.sort) + 1) % order.length];
  render();
  showToast(elements.sortButton.title);
}

function arrayBufferToBase64(buffer) {
  let binary = "";
  const bytes = new Uint8Array(buffer);
  const chunkSize = 8192;
  for (let offset = 0; offset < bytes.length; offset += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + chunkSize));
  }
  return btoa(binary);
}

async function fetchTemplateAsBase64(item) {
  const response = await fetch(item.template);
  if (!response.ok) throw new Error(`Не удалось загрузить «${item.title}» (${response.status})`);
  return arrayBufferToBase64(await response.arrayBuffer());
}

async function insertMaterials(ids) {
  const items = ids.map((id) => materials.find((item) => item.id === id)).filter(Boolean);
  if (!items.length || state.inserting) return;
  if (typeof PowerPoint === "undefined") {
    showToast("Вставка доступна, когда плагин открыт внутри PowerPoint");
    return;
  }

  state.inserting = true;
  renderControls();
  elements.statusRegion.innerHTML = '<div class="status-message">Загружаем выбранные слайды…</div>';

  try {
    const encodedTemplates = [];
    for (const item of items) {
      encodedTemplates.push({ item, base64: await fetchTemplateAsBase64(item) });
    }

    await PowerPoint.run(async (context) => {
      encodedTemplates.forEach(({ base64 }) => context.presentation.insertSlidesFromBase64(base64));
      await context.sync();
    });

    const insertedIds = items.map((item) => item.id);
    state.recent = [
      ...insertedIds,
      ...state.recent.filter((id) => !insertedIds.includes(id)),
    ].slice(0, 20);
    writeStoredArray(STORAGE_KEYS.recent, state.recent);
    state.selected.clear();
    closePreview();
    showToast(items.length === 1 ? "Слайд вставлен" : `Добавлено слайдов: ${items.length}`);
  } catch (error) {
    const message = error instanceof Error ? error.message : "Не удалось вставить выбранные слайды";
    elements.statusRegion.innerHTML = `<div class="error-state">${escapeHtml(message)}. Проверьте интернет и повторите попытку.</div>`;
    showToast("Произошла ошибка при вставке");
  } finally {
    state.inserting = false;
    renderControls();
  }
}

function showToast(message) {
  if (!elements.toast) return;
  window.clearTimeout(toastTimer);
  elements.toast.textContent = message;
  elements.toast.classList.add("is-visible");
  toastTimer = window.setTimeout(() => elements.toast.classList.remove("is-visible"), 2600);
}

function handleLibraryClick(event) {
  const card = event.target.closest(".material-card");
  if (!card) return;
  const id = card.dataset.id;
  const action = event.target.closest("button")?.dataset.action;
  if (action === "favorite") toggleFavorite(id);
  else if (action === "select") toggleSelection(id);
  else openPreview(id);
}

function handleLibraryKeydown(event) {
  if ((event.key === "Enter" || event.key === " ") && event.target.matches(".material-card")) {
    event.preventDefault();
    openPreview(event.target.dataset.id);
  }
}

function bindEvents() {
  elements.collapseButton.addEventListener("click", () => {
    state.sidebarCollapsed = !state.sidebarCollapsed;
    renderControls();
  });

  elements.sectionNav.addEventListener("click", (event) => {
    const button = event.target.closest("[data-section]");
    if (button) selectSection(button.dataset.section);
  });

  document.querySelector(".tabs").addEventListener("click", (event) => {
    const tab = event.target.closest("[data-tab]");
    if (!tab) return;
    state.tab = tab.dataset.tab;
    state.selected.clear();
    render();
  });

  elements.searchInput.addEventListener("input", (event) => {
    state.query = event.target.value;
    renderLibrary();
  });
  elements.filterButton.addEventListener("click", openFilters);
  elements.closeFiltersButton.addEventListener("click", closeFilters);
  elements.resetFiltersButton.addEventListener("click", resetFilters);
  elements.filterForm.addEventListener("submit", (event) => {
    event.preventDefault();
    const data = new FormData(elements.filterForm);
    state.filters = Object.fromEntries(
      ["type", "product", "format", "style"].map((field) => [field, data.get(field) || ""])
    );
    closeFilters();
    render();
  });

  elements.sortButton.addEventListener("click", cycleSort);
  elements.viewButton.addEventListener("click", () => {
    state.view = state.view === "grid" ? "list" : "grid";
    render();
  });
  elements.shareButton.addEventListener("click", async () => {
    const shareData = {
      title: "Slidebrary",
      text: "Библиотека материалов Slidebrary",
      url: BASE_URL,
    };
    try {
      if (navigator.share) await navigator.share(shareData);
      else if (navigator.clipboard) {
        await navigator.clipboard.writeText(BASE_URL);
        showToast("Ссылка скопирована");
      } else showToast("Ссылка: gloomeries.github.io/Slide-library");
    } catch (error) {
      if (error?.name !== "AbortError") showToast("Не удалось поделиться ссылкой");
    }
  });
  elements.moreButton.addEventListener("click", () =>
    showToast("Дополнительные настройки появятся позже")
  );
  elements.closeAppButton.addEventListener("click", () =>
    showToast("Закройте панель крестиком в заголовке PowerPoint")
  );
  elements.library.addEventListener("click", handleLibraryClick);
  elements.library.addEventListener("keydown", handleLibraryKeydown);
  elements.library.addEventListener(
    "error",
    (event) => {
      if (!event.target.matches(".card-preview")) return;
      event.target.alt = "Превью временно недоступно";
      event.target.classList.add("is-broken");
    },
    true
  );

  elements.clearSelectionButton.addEventListener("click", () => {
    state.selected.clear();
    render();
  });
  elements.insertSelectionButton.addEventListener("click", () =>
    insertMaterials([...state.selected])
  );

  elements.closePreviewButton.addEventListener("click", closePreview);
  elements.previewInsertButton.addEventListener("click", () => insertMaterials([state.previewId]));
  elements.previewOverlay.addEventListener("click", (event) => {
    if (event.target === elements.previewOverlay) closePreview();
  });
  elements.filterOverlay.addEventListener("click", (event) => {
    if (event.target === elements.filterOverlay) closeFilters();
  });
  document.addEventListener("keydown", (event) => {
    if (event.key !== "Escape") return;
    if (!elements.previewOverlay.hidden) closePreview();
    else if (!elements.filterOverlay.hidden) closeFilters();
  });
}

let isInitialized = false;

function initialize() {
  if (isInitialized) return;
  isInitialized = true;
  cacheElements();
  populateFilterOptions();
  bindEvents();
  render();
  window.setTimeout(() => {
    state.loading = false;
    render();
  }, 350);
}

if (typeof Office !== "undefined" && Office.onReady) {
  Office.onReady(initialize);
}
document.addEventListener("DOMContentLoaded", () => window.setTimeout(initialize, 250), {
  once: true,
});
