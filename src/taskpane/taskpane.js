/* global Office, PowerPoint, FormData, URLSearchParams, btoa, document, fetch, localStorage, navigator, window */

import "core-js/stable";
import "regenerator-runtime/runtime";

const BASE_URL = "https://gloomeries.github.io/Slide-library";
// Вставьте сюда публичную ссылку на папку Яндекс Диска.
const YANDEX_DISK_PUBLIC_URL = "https://disk.yandex.ru/d/htMsEH_oBBgwEw";
const YANDEX_DISK_PUBLIC_API = "https://cloud-api.yandex.net/v1/disk/public/resources";
const STORAGE_KEYS = {
  favorites: "slidebrary:favorites",
  recent: "slidebrary:recent",
  account: "slidebrary:account",
  proposals: "slidebrary:proposals",
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

const materialFilterMetadata = {
  title: {
    goal: "Информировать",
    structure: "Титульный слайд",
    presentationType: "Информационный",
    visual: "Текст",
    hashtags: "бизнес",
    aiGenerated: false,
  },
  executive_summary: {
    goal: "Отчитаться",
    structure: "Аналитический слайд",
    presentationType: "Отчетный",
    visual: "Инфографика",
    hashtags: "бизнес",
    aiGenerated: false,
  },
  market_analysis: {
    goal: "Отчитаться",
    structure: "Аналитический слайд",
    presentationType: "Отчетный",
    visual: "Диаграмма",
    hashtags: "аналитика",
    aiGenerated: false,
  },
  marketing_plan: {
    goal: "Продать",
    structure: "План",
    presentationType: "Продающий",
    visual: "Инфографика",
    hashtags: "маркетинг",
    aiGenerated: true,
  },
  prototypes: {
    goal: "Обучить",
    structure: "Схема",
    presentationType: "Обучающий",
    visual: "Инфографика",
    hashtags: "продукт",
    aiGenerated: true,
  },
  risk_analysis: {
    goal: "Отчитаться",
    structure: "Аналитический слайд",
    presentationType: "Отчетный",
    visual: "Диаграмма",
    hashtags: "аналитика",
    aiGenerated: false,
  },
  roadmap: {
    goal: "Информировать",
    structure: "План",
    presentationType: "Информационный",
    visual: "Инфографика",
    hashtags: "продукт",
    aiGenerated: false,
  },
  target_audience: {
    goal: "Продать",
    structure: "Аналитический слайд",
    presentationType: "Продающий",
    visual: "Диаграмма",
    hashtags: "маркетинг",
    aiGenerated: false,
  },
  team: {
    goal: "Информировать",
    structure: "Схема",
    presentationType: "Информационный",
    visual: "Фотография",
    hashtags: "бизнес",
    aiGenerated: false,
  },
  business_process: {
    goal: "Обучить",
    structure: "Схема",
    presentationType: "Обучающий",
    visual: "Инфографика",
    hashtags: "бизнес",
    aiGenerated: false,
  },
};

const sections = [
  { id: "favorites", label: "Избранное", icon: "assets/ui/menu/like.svg" },
  { id: "presentations", label: "Презентации", icon: "assets/ui/menu/menu/time/active.svg" },
  { id: "photos", label: "Фотографии", icon: "assets/ui/menu/menu/camera/active.svg" },
  {
    id: "illustrations",
    label: "Иллюстрации",
    icon: "assets/ui/menu/menu/picture-pen/active.svg",
  },
  { id: "icons", label: "Иконки", icon: "assets/ui/menu/menu/bullet/active.svg" },
  { id: "logos", label: "Логотипы", icon: "assets/ui/menu/menu/style/active.svg" },
  {
    id: "templates",
    label: "Шаблоны",
    icon: "assets/ui/menu/menu/grid-rectangle/active.svg",
  },
  { id: "assistant", label: "ИИ-ассистент", icon: "assets/ui/menu/menu/ai/active.svg" },
];

const sectionHints = {
  favorites: "Избранные материалы",
  presentations: "Выберите макет для вашего слайда",
  photos: "Выберите фотографию для вашего слайда",
  illustrations: "Выберите изображение для вашего слайда",
  icons: "Выберите иконку для вашего слайда",
  logos: "Выберите логотип для вашего слайда",
  templates: "Выберите макет для вашего слайда",
  assistant: "Создайте материал с помощью ИИ-ассистента",
};

const sectionFilterConfigs = {
  photos: {
    label: "Продукт группы VK",
    options: ["MAX", "VK", "Сферум", "Одноклассники"],
  },
  illustrations: {
    label: "Тип изображения",
    options: ["3D", "2D", "Фото", "Абстракция"],
  },
  icons: {
    label: "Расширение файла: svg, png, gif",
    options: ["SVG", "PNG", "GIF"],
  },
  logos: {
    label: "Выберите продукт",
    options: ["MAX", "VK", "Сферум", "Одноклассники"],
  },
  templates: {
    label: "Выберите продукт",
    options: ["MAX", "VK", "Сферум", "Одноклассники"],
  },
};

const state = {
  section: "presentations",
  tab: "public",
  query: "",
  view: "grid",
  sort: "default",
  filters: {
    type: "",
    product: "",
    goal: "",
    format: "",
    structure: "",
    presentationType: "",
    visual: "",
    style: "",
    hashtags: "",
    aiGenerated: false,
  },
  favorites: new Set(readStoredArray(STORAGE_KEYS.favorites)),
  recent: readStoredArray(STORAGE_KEYS.recent),
  selected: new Set(),
  previewId: null,
  loading: true,
  inserting: false,
  photosLoading: false,
  photosLoaded: false,
  photosError: "",
  sidebarCollapsed: false,
  account: readStoredObject(STORAGE_KEYS.account),
  templateFile: null,
  templateTags: [],
  sectionFilters: {
    photos: "Все",
    illustrations: "3D",
    icons: "SVG",
    logos: "Все",
    templates: "Все",
  },
};

function hashString(value) {
  let hash = 0;
  for (let index = 0; index < value.length; index += 1) {
    hash = (hash * 31 + value.charCodeAt(index)) >>> 0;
  }
  return hash.toString(36);
}

function getPhotoProduct(path) {
  const normalizedPath = String(path || "").toLocaleLowerCase("ru");
  const products = ["MAX", "VK", "Сферум", "Одноклассники"];
  return (
    products.find((product) => normalizedPath.includes(product.toLocaleLowerCase("ru"))) || "MAX"
  );
}

function isImageResource(resource) {
  return resource.type === "file" && /^image\//i.test(resource.mime_type || "");
}

function makePhotoMaterial(resource, relativePath) {
  const title = resource.name.replace(/\.[^.]+$/, "");
  return {
    id: `yandex-photo-${hashString(resource.path || relativePath || resource.name)}`,
    title,
    product: getPhotoProduct(relativePath),
    type: "Фотография",
    format: "Изображение",
    style: "Фотография",
    tags: ["фотография", getPhotoProduct(relativePath), title],
    preview: resource.preview || resource.file,
    source: resource.file || resource.preview,
    librarySection: "photos",
  };
}

async function fetchYandexFolder(path = "", depth = 0) {
  if (depth > 6) return [];

  const params = new URLSearchParams({
    public_key: YANDEX_DISK_PUBLIC_URL,
    limit: "1000",
    preview_size: "XL",
    preview_crop: "false",
  });
  if (path) params.set("path", path);

  const response = await fetch(`${YANDEX_DISK_PUBLIC_API}?${params.toString()}`);
  if (!response.ok) {
    throw new Error(`Яндекс Диск вернул ошибку ${response.status}`);
  }

  const resource = await response.json();
  const children = resource._embedded?.items || [];
  const photos = children
    .filter(isImageResource)
    .map((item) => makePhotoMaterial(item, path ? `${path}/${item.name}` : item.name));
  const folders = children.filter((item) => item.type === "dir");
  const nestedPhotos = await Promise.all(
    folders.map((folder) =>
      fetchYandexFolder(path ? `${path}/${folder.name}` : folder.name, depth + 1)
    )
  );
  return photos.concat(...nestedPhotos);
}

function updatePhotoFilterOptions(photos) {
  const products = [...new Set(photos.map((photo) => photo.product))].sort((a, b) =>
    a.localeCompare(b, "ru")
  );
  sectionFilterConfigs.photos.options = ["Все", ...products];
}

async function loadYandexPhotos() {
  if (state.photosLoading || state.photosLoaded) return;
  if (!YANDEX_DISK_PUBLIC_URL) {
    state.photosError = "Добавьте публичную ссылку Яндекс Диска в настройку YANDEX_DISK_PUBLIC_URL";
    render();
    return;
  }

  state.photosLoading = true;
  state.photosError = "";
  render();
  try {
    const photos = await fetchYandexFolder();
    materials.push(...photos);
    updatePhotoFilterOptions(photos);
    state.photosLoaded = true;
  } catch (error) {
    state.photosError = error?.message || "Не удалось загрузить фотографии с Яндекс Диска";
  } finally {
    state.photosLoading = false;
    render();
  }
}

const elements = {};
let toastTimer;
let pendingFolderName = "";

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

function readStoredObject(key) {
  try {
    const value = JSON.parse(localStorage.getItem(key) || "{}");
    return value && typeof value === "object" && !Array.isArray(value) ? value : {};
  } catch {
    return {};
  }
}

function writeStoredObject(key, value) {
  try {
    localStorage.setItem(key, JSON.stringify(value));
  } catch {
    showToast("Не удалось сохранить данные на этом устройстве");
  }
}

function formatAccountDate(value) {
  if (!value) return "Ещё не выполнялся";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return value;
  return new Intl.DateTimeFormat("ru-RU", {
    day: "numeric",
    month: "long",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
  }).format(date);
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function getSectionIcon(path) {
  return `<img class="nav-icon" src="${escapeHtml(path)}" alt="" aria-hidden="true" />`;
}

function cacheElements() {
  [
    "sectionNav",
    "collapseButton",
    "collapseIcon",
    "profileButton",
    "searchInput",
    "libraryView",
    "templateView",
    "assistantView",
    "assistantForm",
    "assistantPrompt",
    "assistantPromptCounter",
    "cancelAssistantButton",
    "startAssistantButton",
    "templateForm",
    "templateDropzone",
    "templateFileInput",
    "templateFileLabel",
    "templateNameInput",
    "templateTagInput",
    "templateTags",
    "templateTagCounter",
    "cancelTemplateButton",
    "submitTemplateButton",
    "filterButton",
    "filterBadge",
    "sortButton",
    "viewButton",
    "shareButton",
    "accountOverlay",
    "accountForm",
    "accountEmail",
    "folderPickerButton",
    "folderInput",
    "lastLoginField",
    "indexUpdatedField",
    "cancelAccountButton",
    "sectionHint",
    "sectionFilter",
    "sectionFilterLabel",
    "sectionFilterSelect",
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
        .map(
          (option) =>
            `<option value="${escapeHtml(option)}">${escapeHtml(field === "format" ? option.replace("×", "x") : option)}</option>`
        )
        .join("")
    );
  });
}

function getMaterialFilterValue(item, field) {
  if (field in item) return item[field];
  return materialFilterMetadata[item.id]?.[field];
}

function getVisibleMaterials() {
  if (state.tab === "personal") return [];

  let result = [...materials];
  if (state.section === "favorites") {
    result = result.filter((item) => state.favorites.has(item.id));
  } else if (state.section === "photos") {
    result = result.filter((item) => item.librarySection === "photos");
    if (state.sectionFilters.photos !== "Все") {
      result = result.filter((item) => item.product === state.sectionFilters.photos);
    }
  } else if (["presentations", "templates"].includes(state.section)) {
    result = result.filter((item) => item.librarySection !== "photos");
  } else {
    result = [];
  }
  if (state.section === "templates" && state.sectionFilters.templates !== "MAX") result = [];

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
    if (!value) return;
    result = result.filter((item) => getMaterialFilterValue(item, field) === value);
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
  if (state.section === "photos" && state.photosError) return state.photosError;
  if (state.section === "photos" && state.photosLoaded)
    return "В публичной папке Яндекс Диска пока нет фотографий.";
  if (sectionFilterConfigs[state.section]) {
    return `Для выбранного значения «${state.sectionFilters[state.section]}» пока нет материалов.`;
  }
  if (state.section === "assistant") return "ИИ-ассистент появится на следующем этапе разработки.";
  if (state.query || getActiveFilterCount())
    return "По вашему запросу ничего не найдено. Попробуйте изменить поиск или фильтры.";
  return "В библиотеке пока нет материалов.";
}

function renderStatus() {
  if (state.loading || (state.section === "photos" && state.photosLoading)) {
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
  if (state.loading || (state.section === "photos" && state.photosLoading)) {
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
          >${isSelected ? '<span class="selected-icon" aria-hidden="true">✓</span>' : '<span class="plus-icon" aria-hidden="true"></span>'}</button>
        </article>`;
    })
    .join("");
}

function renderControls() {
  const isTemplateView = state.section === "templates" && state.tab === "personal";
  const isAssistantView = state.section === "assistant";
  elements.libraryView.hidden = isTemplateView || isAssistantView;
  elements.templateView.hidden = !isTemplateView;
  elements.assistantView.hidden = !isAssistantView;
  renderSectionFilter();
  elements.sectionHint.textContent = sectionHints[state.section];
  elements.library.dataset.view = state.view;
  document.getElementById("app").classList.toggle("is-sidebar-collapsed", state.sidebarCollapsed);
  elements.collapseButton.setAttribute(
    "aria-label",
    state.sidebarCollapsed ? "Развернуть меню" : "Свернуть меню"
  );
  elements.collapseButton.title = state.sidebarCollapsed ? "Развернуть меню" : "Свернуть меню";
  elements.collapseIcon.src = state.sidebarCollapsed
    ? "assets/ui/развернуть.svg"
    : "assets/ui/icon.svg";
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

function renderSectionFilter() {
  const config = sectionFilterConfigs[state.section];
  const isVisible = Boolean(config) && state.tab === "public";
  elements.sectionFilter.hidden = !isVisible;
  if (!isVisible) return;

  elements.sectionFilterLabel.textContent = config.label;
  elements.sectionFilterSelect.innerHTML = config.options
    .map(
      (option) =>
        `<option value="${escapeHtml(option)}"${state.sectionFilters[state.section] === option ? " selected" : ""}>${escapeHtml(option)}</option>`
    )
    .join("");
}

function updateTemplateSubmitState() {
  elements.submitTemplateButton.disabled =
    !state.templateFile || !elements.templateNameInput.value.trim();
}

function renderTemplateTags() {
  elements.templateTags.innerHTML = state.templateTags
    .map(
      (tag, index) => `
        <span class="template-tag">
          <span>${escapeHtml(tag)}</span>
          <button type="button" data-tag-index="${index}" aria-label="Удалить хештег ${escapeHtml(tag)}">×</button>
        </span>`
    )
    .join("");
  elements.templateTagCounter.textContent = `${state.templateTags.length} / 25`;
  elements.templateTagInput.disabled = state.templateTags.length >= 25;
}

function addTemplateTag(rawValue) {
  const tag = rawValue.trim().replace(/^#/, "");
  if (!tag || state.templateTags.length >= 25) return;
  if (
    !state.templateTags.some(
      (existing) => existing.toLocaleLowerCase("ru") === tag.toLocaleLowerCase("ru")
    )
  ) {
    state.templateTags.push(tag);
  }
  elements.templateTagInput.value = "";
  renderTemplateTags();
}

function setTemplateFile(file) {
  const allowedExtension = /\.(pptx?|jpe?g|png)$/i.test(file?.name || "");
  const allowedSize = file && file.size <= 20 * 1024 * 1024;
  if (!allowedExtension || !allowedSize) {
    state.templateFile = null;
    elements.templateFileLabel.textContent = "Загрузить файл";
    elements.templateDropzone.classList.remove("has-file");
    showToast(
      !allowedExtension
        ? "Выберите файл pptx, jpg или png"
        : "Размер файла не должен превышать 20 Мб"
    );
    updateTemplateSubmitState();
    return;
  }
  state.templateFile = file;
  elements.templateFileLabel.textContent = file.name;
  elements.templateDropzone.classList.add("has-file");
  updateTemplateSubmitState();
}

function resetTemplateForm() {
  elements.templateForm.reset();
  state.templateFile = null;
  state.templateTags = [];
  elements.templateFileLabel.textContent = "Загрузить файл";
  elements.templateDropzone.classList.remove("has-file", "is-dragover");
  renderTemplateTags();
  updateTemplateSubmitState();
}

function updateAssistantForm() {
  const promptLength = elements.assistantPrompt.value.length;
  elements.assistantPromptCounter.textContent = `${promptLength} / 2000`;
  elements.startAssistantButton.disabled = !elements.assistantPrompt.value.trim();
}

function resetAssistantForm() {
  elements.assistantForm.reset();
  updateAssistantForm();
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
  const compactSections = ["templates", "assistant"];
  const isLeavingCompactView =
    compactSections.includes(state.section) && !compactSections.includes(sectionId);
  state.section = sectionId;
  if (compactSections.includes(sectionId)) state.sidebarCollapsed = true;
  else if (isLeavingCompactView) state.sidebarCollapsed = false;
  state.selected.clear();
  render();
  if (sectionId === "photos") loadYandexPhotos();
  window.scrollTo(0, 0);
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
  elements.previewOverlay.hidden = false;
  elements.closePreviewButton.focus();
}

function closePreview() {
  elements.previewOverlay.hidden = true;
  state.previewId = null;
}

function openFilters() {
  Object.entries(state.filters).forEach(([field, value]) => {
    const control = elements.filterForm.elements[field];
    if (field === "aiGenerated") control.checked = Boolean(value);
    else control.value = value;
  });
  elements.filterOverlay.hidden = false;
}

function closeFilters() {
  elements.filterOverlay.hidden = true;
  elements.filterButton.focus();
}

function openAccount() {
  pendingFolderName = state.account.folderName || "";
  elements.accountEmail.value = state.account.email || "";
  elements.folderPickerButton.querySelector(".folder-path").textContent = pendingFolderName
    ? `//…/${pendingFolderName}`
    : "//… Выбрать…";
  elements.lastLoginField.value = formatAccountDate(state.account.lastLogin);
  elements.indexUpdatedField.value = state.account.indexUpdated
    ? formatAccountDate(state.account.indexUpdated)
    : "Файл не выбран";
  elements.folderInput.value = "";
  elements.accountOverlay.hidden = false;
  elements.accountEmail.focus();
}

function closeAccount() {
  elements.accountOverlay.hidden = true;
  elements.folderInput.value = "";
  pendingFolderName = "";
  elements.profileButton.focus();
}

function getFolderNameFromFiles(files) {
  const firstFile = files?.[0];
  if (!firstFile) return "";
  const relativePath = firstFile.webkitRelativePath || firstFile.name;
  return relativePath.split("/")[0] || firstFile.name;
}

function resetFilters() {
  state.filters = {
    type: "",
    product: "",
    goal: "",
    format: "",
    structure: "",
    presentationType: "",
    visual: "",
    style: "",
    hashtags: "",
    aiGenerated: false,
  };
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
  elements.sectionFilterSelect.addEventListener("change", (event) => {
    state.sectionFilters[state.section] = event.target.value;
    renderLibrary();
  });

  elements.assistantPrompt.addEventListener("input", updateAssistantForm);
  elements.cancelAssistantButton.addEventListener("click", resetAssistantForm);
  elements.assistantForm.addEventListener("submit", (event) => {
    event.preventDefault();
    if (!elements.assistantPrompt.value.trim()) return;
    showToast("Запрос готов. Подключение генерации будет следующим этапом");
  });

  elements.templateDropzone.addEventListener("click", () => elements.templateFileInput.click());
  elements.templateFileInput.addEventListener("change", (event) => {
    const file = event.target.files?.[0];
    if (file) setTemplateFile(file);
  });
  elements.templateDropzone.addEventListener("dragover", (event) => {
    event.preventDefault();
    elements.templateDropzone.classList.add("is-dragover");
  });
  elements.templateDropzone.addEventListener("dragleave", () =>
    elements.templateDropzone.classList.remove("is-dragover")
  );
  elements.templateDropzone.addEventListener("drop", (event) => {
    event.preventDefault();
    elements.templateDropzone.classList.remove("is-dragover");
    const file = event.dataTransfer?.files?.[0];
    if (file) setTemplateFile(file);
  });
  elements.templateNameInput.addEventListener("input", updateTemplateSubmitState);
  elements.templateTagInput.addEventListener("keydown", (event) => {
    if (event.key !== "Enter" && event.key !== ",") return;
    event.preventDefault();
    addTemplateTag(event.target.value);
  });
  elements.templateTagInput.addEventListener("blur", (event) => addTemplateTag(event.target.value));
  elements.templateTags.addEventListener("click", (event) => {
    const removeButton = event.target.closest("[data-tag-index]");
    if (!removeButton) return;
    state.templateTags.splice(Number(removeButton.dataset.tagIndex), 1);
    renderTemplateTags();
  });
  elements.cancelTemplateButton.addEventListener("click", resetTemplateForm);
  elements.templateForm.addEventListener("submit", (event) => {
    event.preventDefault();
    if (!state.templateFile || !elements.templateNameInput.value.trim()) return;
    const data = new FormData(elements.templateForm);
    const proposals = readStoredArray(STORAGE_KEYS.proposals);
    proposals.unshift({
      id: `proposal-${Date.now()}`,
      fileName: state.templateFile.name,
      fileSize: state.templateFile.size,
      title: elements.templateNameInput.value.trim(),
      type: data.get("type") || "",
      product: data.get("product") || "",
      goal: data.get("goal") || "",
      format: data.get("format") || "",
      structure: data.get("structure") || "",
      presentationType: data.get("presentationType") || "",
      visual: data.get("visual") || "",
      style: data.get("style") || "",
      tags: [...state.templateTags],
      aiGenerated: data.get("aiGenerated") === "on",
      createdAt: new Date().toISOString(),
    });
    writeStoredArray(STORAGE_KEYS.proposals, proposals.slice(0, 50));
    resetTemplateForm();
    showToast("Шаблон добавлен в список предложений");
  });

  elements.profileButton.addEventListener("click", openAccount);
  elements.cancelAccountButton.addEventListener("click", closeAccount);
  elements.folderPickerButton.addEventListener("click", () => elements.folderInput.click());
  elements.folderInput.addEventListener("change", (event) => {
    pendingFolderName = getFolderNameFromFiles(event.target.files);
    if (!pendingFolderName) return;
    elements.folderPickerButton.querySelector(".folder-path").textContent =
      `//…/${pendingFolderName}`;
    elements.indexUpdatedField.value = formatAccountDate(new Date().toISOString());
  });
  elements.accountForm.addEventListener("submit", (event) => {
    event.preventDefault();
    const now = new Date().toISOString();
    const folderChanged = Boolean(elements.folderInput.files?.length);
    state.account = {
      email: elements.accountEmail.value.trim(),
      folderName: pendingFolderName,
      lastLogin: now,
      indexUpdated: folderChanged ? now : state.account.indexUpdated || "",
    };
    writeStoredObject(STORAGE_KEYS.account, state.account);
    closeAccount();
    showToast("Данные личного кабинета сохранены");
  });
  elements.accountOverlay.addEventListener("click", (event) => {
    if (event.target === elements.accountOverlay) closeAccount();
  });

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
      [
        "type",
        "product",
        "goal",
        "format",
        "structure",
        "presentationType",
        "visual",
        "style",
        "hashtags",
      ].map((field) => [field, data.get(field) || ""])
    );
    state.filters.aiGenerated = data.get("aiGenerated") === "on";
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
    if (!elements.accountOverlay.hidden) closeAccount();
    else if (!elements.previewOverlay.hidden) closePreview();
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
  renderTemplateTags();
  updateAssistantForm();
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
