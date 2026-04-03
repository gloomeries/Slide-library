const slides = [
  { id: 1, title: "Титульный слайд — Компания", category: "Титульные слайды", tags: ["титул", "начало", "компания"], icon: "🏢" },
  { id: 2, title: "Титульный слайд — Проект", category: "Титульные слайды", tags: ["титул", "проект"], icon: "📋" },
  { id: 3, title: "Agenda встречи", category: "Структура и agenda", tags: ["agenda", "план", "структура"], icon: "📌" },
  { id: 4, title: "Структура презентации", category: "Структура и agenda", tags: ["структура", "оглавление"], icon: "🗂️" },
  { id: 5, title: "График — Линейный", category: "Графики и данные", tags: ["график", "данные", "динамика"], icon: "📈" },
  { id: 6, title: "График — Столбчатый", category: "Графики и данные", tags: ["график", "сравнение", "данные"], icon: "📊" },
  { id: 7, title: "Таблица данных", category: "Графики и данные", tags: ["таблица", "данные", "цифры"], icon: "🔢" },
  { id: 8, title: "Команда проекта", category: "Команда и контакты", tags: ["команда", "люди", "роли"], icon: "👥" },
  { id: 9, title: "Контакты", category: "Команда и контакты", tags: ["контакты", "связь"], icon: "📞" },
  { id: 10, title: "Дорожная карта", category: "Дорожная карта", tags: ["roadmap", "план", "этапы"], icon: "🗺️" },
  { id: 11, title: "Этапы проекта", category: "Дорожная карта", tags: ["этапы", "timeline", "план"], icon: "📅" },
];

let activeCategory = "Все";

const categories = ["Все", ...new Set(slides.map(s => s.category))];

Office.onReady(() => {
  renderFilters();
  renderSlides();
});

function renderFilters() {
  const container = document.getElementById("filters");
  container.innerHTML = categories.map(cat => `
    <button class="filter-btn ${cat === activeCategory ? 'active' : ''}"
      onclick="setCategory('${cat}')">${cat}</button>
  `).join('');
}

function setCategory(cat) {
  activeCategory = cat;
  renderFilters();
  renderSlides();
}

function renderSlides() {
  const query = document.getElementById("searchInput").value.toLowerCase();
  const container = document.getElementById("slideList");

  const filtered = slides.filter(s => {
    const matchCat = activeCategory === "Все" || s.category === activeCategory;
    const matchQuery = !query ||
      s.title.toLowerCase().includes(query) ||
      s.tags.some(t => t.includes(query));
    return matchCat && matchQuery;
  });

  if (filtered.length === 0) {
    container.innerHTML = '<div class="empty">Ничего не найдено</div>';
    return;
  }

  container.innerHTML = filtered.map(s => `
    <div class="slide-card">
      <div class="slide-icon">${s.icon}</div>
      <div class="slide-info">
        <div class="slide-title">${s.title}</div>
        <div class="slide-tag">${s.category}</div>
      </div>
      <button class="insert-btn" onclick="insertSlide(${s.id})">Вставить</button>
    </div>
  `).join('');
}

function insertSlide(id) {
  const slide = slides.find(s => s.id === id);
  PowerPoint.run(async (context) => {
    const newSlide = context.presentation.slides.add();
    await context.sync();

    const shapes = newSlide.shapes;
    const textBox = shapes.addTextBox(slide.title, { left: 100, top: 100, width: 600, height: 80 });
    textBox.textFrame.textRange.font.size = 28;
    textBox.textFrame.textRange.font.bold = true;

    const subBox = shapes.addTextBox(slide.category, { left: 100, top: 200, width: 600, height: 40 });
    subBox.textFrame.textRange.font.size = 16;
    subBox.textFrame.textRange.font.color = "#666666";

    await context.sync();

    showToast();
  });
}

function showToast() {
  const toast = document.getElementById("toast");
  toast.style.display = "block";
  setTimeout(() => { toast.style.display = "none"; }, 2000);
}