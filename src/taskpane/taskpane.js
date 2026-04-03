const BASE_URL = "https://gloomeries.github.io/Slide-library";

const categories = [
  { id: "title", name: "Титул", button: `${BASE_URL}/assets/buttons/title.png`, preview: `${BASE_URL}/assets/previews/title.png`, template: `${BASE_URL}/assets/templates/title.pptx` },
  { id: "executive_summary", name: "Executive Summary", button: `${BASE_URL}/assets/buttons/executive_summary.png`, preview: `${BASE_URL}/assets/previews/executive_summary.png`, template: `${BASE_URL}/assets/templates/executive_summary.pptx` },
  { id: "market_analysis", name: "Анализ рынка", button: `${BASE_URL}/assets/buttons/market_analysis.png`, preview: `${BASE_URL}/assets/previews/market_analysis.png`, template: `${BASE_URL}/assets/templates/market_analysis.pptx` },
  { id: "marketing_plan", name: "Маркетинговый план", button: `${BASE_URL}/assets/buttons/marketing_plan.png`, preview: `${BASE_URL}/assets/previews/marketing_plan.png`, template: `${BASE_URL}/assets/templates/marketing_plan.pptx` },
  { id: "prototypes", name: "Описание прототипов", button: `${BASE_URL}/assets/buttons/prototypes.png`, preview: `${BASE_URL}/assets/previews/prototypes.png`, template: `${BASE_URL}/assets/templates/prototypes.pptx` },
  { id: "risk_analysis", name: "Матрица рисков", button: `${BASE_URL}/assets/buttons/risk_analysis.png`, preview: `${BASE_URL}/assets/previews/risk_analysis.png`, template: `${BASE_URL}/assets/templates/risk_matrix.pptx` },
  { id: "roadmap", name: "Дорожная карта", button: `${BASE_URL}/assets/buttons/roadmap.png`, preview: `${BASE_URL}/assets/previews/roadmap.png`, template: `${BASE_URL}/assets/templates/roadmap.pptx` },
  { id: "target_audience", name: "Целевая аудитория", button: `${BASE_URL}/assets/buttons/target_audience.png`, preview: `${BASE_URL}/assets/previews/target_audience.png`, template: `${BASE_URL}/assets/templates/target_audience.pptx` },
  { id: "team", name: "Команда проекта", button: `${BASE_URL}/assets/buttons/team.png`, preview: `${BASE_URL}/assets/previews/team.png`, template: `${BASE_URL}/assets/templates/team.pptx` },
  { id: "business_process", name: "Бизнес-процесс", button: `${BASE_URL}/assets/buttons/business_process.png`, preview: `${BASE_URL}/assets/previews/business_process.png`, template: `${BASE_URL}/assets/templates/business_process.pptx` },
];

let currentTemplate = null;

Office.onReady(() => {
  renderCategories(categories);
});

function renderCategories(list) {
  const grid = document.getElementById("categoryGrid");
  if (list.length === 0) {
    grid.innerHTML = '<div class="empty" style="grid-column:span 2">Ничего не найдено</div>';
    return;
  }
  grid.innerHTML = list.map(cat => `
    <div class="category-card" onclick="openCatalog('${cat.id}')">
      <img src="${cat.button}" alt="${cat.name}" onerror="this.style.background='#f0f0f0'" />
      <div class="label">${cat.name}</div>
    </div>
  `).join('');
}

function handleSearch() {
  const query = document.getElementById("searchInput").value.toLowerCase();
  const filtered = categories.filter(c => c.name.toLowerCase().includes(query));
  renderCategories(filtered);
}

function openCatalog(categoryId) {
  const cat = categories.find(c => c.id === categoryId);
  if (!cat) return;

  document.getElementById("mainScreen").style.display = "none";
  document.getElementById("catalogScreen").style.display = "block";
  document.getElementById("catalogTitle").textContent = cat.name;

  document.getElementById("slideGrid").innerHTML = `
    <div class="slide-card" onclick="openModal('${cat.preview}', '${cat.name}', '${cat.template}')">
      <img src="${cat.preview}" alt="${cat.name}" />
    </div>
  `;
}

function goBack() {
  document.getElementById("catalogScreen").style.display = "none";
  document.getElementById("mainScreen").style.display = "block";
}

function openModal(previewUrl, title, templateUrl) {
  currentTemplate = templateUrl;
  document.getElementById("modalImg").src = previewUrl;
  document.getElementById("modalTitle").textContent = title;
  document.getElementById("modalOverlay").classList.add("active");
}

function closeModal() {
  document.getElementById("modalOverlay").classList.remove("active");
  currentTemplate = null;
}

function insertSlide() {
  if (!currentTemplate) return;

  fetch(currentTemplate)
    .then(res => res.arrayBuffer())
    .then(buffer => {
      const base64 = arrayBufferToBase64(buffer);
      PowerPoint.run(async (context) => {
        context.presentation.insertSlidesFromBase64(base64);
        await context.sync();
        closeModal();
        showToast();
      });
    })
    .catch(() => {
      showToast("Ошибка при вставке слайда");
    });
}

function arrayBufferToBase64(buffer) {
  let binary = '';
  const bytes = new Uint8Array(buffer);
  for (let i = 0; i < bytes.byteLength; i++) {
    binary += String.fromCharCode(bytes[i]);
  }
  return btoa(binary);
}

function showToast(msg) {
  const toast = document.getElementById("toast");
  toast.textContent = msg || "Слайд вставлен!";
  toast.style.display = "block";
  setTimeout(() => { toast.style.display = "none"; }, 2500);
}

