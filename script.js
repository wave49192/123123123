tailwind.config = {
  darkMode: "class",
  theme: {
    extend: {
      fontFamily: {
        sarabun: ["Sarabun", "sans-serif"],
      },
    },
  },
};

window.addEventListener("DOMContentLoaded", () => {
  // --- MEMORY OBJECTS STATES ---
  let rawQuestionsList = []; // Imported spreadsheet data database
  let filteredQuestions = []; // Current active viewport filtered list
  let currentIndex = 0; // Chosen question index in filteredQuestions array

  // Status trackers synchronized with LocalStorage
  let answersState = {}; // { id: { selected: "A", correct: true, choiceMapping: [...] } }
  let bookmarksState = []; // [ "1", "2" ] - Bookmarked questionNo IDs
  let filterStatus = "all"; // all, unanswered, correct, incorrect, bookmarked
  let filterCaseId = "all"; // all, or specific Case ID string
  let selectedYear = null;
  let searchString = ""; // Search word queries match
  let shuffleChoicesMode = false;

  // UI Target Node Map Bindings
  const activeQuestionText = document.getElementById("activeQuestionText");
  const choicesContainer = document.getElementById("choicesContainer");
  const explanationBox = document.getElementById("explanationBox");

  const displayIndex = document.getElementById("displayIndex");
  const displayTotal = document.getElementById("displayTotal");
  const questionIdBadge = document.getElementById("questionIdBadge");

  const prevQuestionBtn = document.getElementById("prevQuestionBtn");
  const nextQuestionBtn = document.getElementById("nextQuestionBtn");
  const bookmarkButton = document.getElementById("bookmarkButton");
  const gridNavigator = document.getElementById("gridNavigator");

  const statsAnswered = document.getElementById("statsAnswered");
  const statsCorrect = document.getElementById("statsCorrect");
  const statsWrong = document.getElementById("statsWrong");
  const statsPercent = document.getElementById("statsPercent");
  const linearProgressBar = document.getElementById("linearProgressBar");

  const searchTerm = document.getElementById("searchTerm");
  const statusFilter = document.getElementById("statusFilter");
  const caseFilter = document.getElementById("caseFilter");
  const shuffleQsBtn = document.getElementById("shuffleQsBtn");
  const shuffleChoicesBtn = document.getElementById("shuffleChoicesBtn");
  const shuffleChoicesText = document.getElementById("shuffleChoicesText");

  const themeToggler = document.getElementById("themeToggler");
  const themeIcon = document.getElementById("themeIcon");
  const clearStateBtn = document.getElementById("clearStateBtn");
  const navUploadBtn = document.getElementById("navUploadBtn");
  const openAddModalBtn = document.getElementById("openAddModalBtn");

  // Case Cards
  const caseScenarioCard = document.getElementById("caseScenarioCard");
  const caseHeaderLabel = document.getElementById("caseHeaderLabel");
  const caseDescriptionText = document.getElementById(
    "caseDescriptionText",
  );

  // Editor modal elements
  const caseEditorModal = document.getElementById("caseEditorModal");
  const modalContainer = document.getElementById("modalContainer");
  const modalTitle = document.getElementById("modalTitle");
  const caseForm = document.getElementById("caseForm");

  const formYear = document.getElementById("formYear");
  const formCaseId = document.getElementById("formCaseId");
  const formCaseDescription = document.getElementById(
    "formCaseDescription",
  );
  const formQuestionNo = document.getElementById("formQuestionNo");
  const formQuestionText = document.getElementById("formQuestionText");
  const formChoiceA = document.getElementById("formChoiceA");
  const formChoiceB = document.getElementById("formChoiceB");
  const formChoiceC = document.getElementById("formChoiceC");
  const formChoiceD = document.getElementById("formChoiceD");
  const formChoiceE = document.getElementById("formChoiceE");
  const formAnswer = document.getElementById("formAnswer");
  const formExplanation = document.getElementById("formExplanation");
  const editOriginalNo = document.getElementById("editOriginalNo");

  const closeModalBtn = document.getElementById("closeModalBtn");
  const cancelFormBtn = document.getElementById("cancelFormBtn");
  const editQuestionBtn = document.getElementById("editQuestionBtn");
  const landingPage = document.getElementById("landingPage");
  const appShell = document.getElementById("appShell");
  const yearCardsContainer = document.getElementById("yearCardsContainer");
  const backToYearsBtn = document.getElementById("backToYearsBtn");
  // Temporary implicit input for excel trigger
  let excelImporter;

  // --- APPS INITIATION RUN ---
  async function initApp() {
    loadProgressOnlyFromLocalStorage();

    rawQuestionsList = await autoLoadPLEExcelFiles();

    if (
      localStorage.getItem("theme") === "dark" ||
      (!("theme" in localStorage) &&
        window.matchMedia("(prefers-color-scheme: dark)").matches)
    ) {
      document.documentElement.classList.add("dark");
      themeIcon.className = "fa-solid fa-sun text-sm";
    } else {
      document.documentElement.classList.remove("dark");
      themeIcon.className = "fa-solid fa-moon text-sm";
    }

    renderYearLandingPage();
    showLandingPage();

    excelImporter = document.createElement("input");
    excelImporter.type = "file";
    excelImporter.accept = ".xlsx, .xls, .csv";
    excelImporter.addEventListener("change", parseSpreadsheetFile);

    navUploadBtn.addEventListener("click", () => {
      excelImporter.click();
    });
  }



  function renderYearLandingPage() {
    if (!yearCardsContainer) return;

    const years = [...new Set(rawQuestionsList.map((q) => String(q.year || "").trim()))]
      .filter(Boolean)
      .sort((a, b) => Number(b) - Number(a) || b.localeCompare(a));

    if (years.length === 0) {
      yearCardsContainer.innerHTML = `
        <div class="col-span-full text-center py-10 text-rose-400">
          <i class="fa-solid fa-triangle-exclamation text-3xl mb-3"></i>
          <p class="font-extrabold">ไม่พบไฟล์ข้อสอบ Excel</p>
          <p class="text-sm mt-1 text-slate-400">วางไฟล์ PLE_PC1_2563.xlsx ถึง PLE_PC1_2568.xlsx ไว้โฟลเดอร์เดียวกับ index.html</p>
        </div>
      `;
      return;
    }

    yearCardsContainer.innerHTML = years
      .map((year) => {
        const questions = rawQuestionsList.filter((q) => String(q.year) === String(year));
        const total = questions.length;
        const answered = questions.filter((q) => answersState[q.id]?.selected !== undefined).length;
        const correct = questions.filter((q) => answersState[q.id]?.correct).length;
        const percent = answered > 0 ? Math.round((correct / answered) * 100) : 0;

        return `
          <button type="button" data-year="${year}" class="year-card group text-left p-5 rounded-3xl border border-slate-200 dark:border-slate-700 bg-slate-50/70 dark:bg-slate-900/50 hover:border-indigo-400 hover:bg-indigo-50/40 dark:hover:bg-indigo-950/20 transition shadow-sm">
            <div class="flex items-start justify-between gap-3 mb-5">
              <div>
                <p class="text-xs font-bold text-slate-400 uppercase tracking-wider mb-1">ปีข้อสอบ</p>
                <h3 class="text-3xl font-extrabold text-indigo-600 dark:text-indigo-400">${year}</h3>
              </div>
              <span class="w-11 h-11 rounded-2xl bg-indigo-100 dark:bg-indigo-950/50 text-indigo-600 dark:text-indigo-300 flex items-center justify-center group-hover:scale-105 transition">
                <i class="fa-solid fa-arrow-right"></i>
              </span>
            </div>
            <div class="grid grid-cols-3 gap-2 text-center">
              <div class="rounded-2xl bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 p-3">
                <p class="text-lg font-extrabold text-slate-800 dark:text-white">${total}</p>
                <p class="text-[10px] text-slate-400 font-bold">ข้อทั้งหมด</p>
              </div>
              <div class="rounded-2xl bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 p-3">
                <p class="text-lg font-extrabold text-emerald-500">${answered}</p>
                <p class="text-[10px] text-slate-400 font-bold">ทำแล้ว</p>
              </div>
              <div class="rounded-2xl bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 p-3">
                <p class="text-lg font-extrabold text-amber-500">${percent}%</p>
                <p class="text-[10px] text-slate-400 font-bold">ถูก</p>
              </div>
            </div>
          </button>
        `;
      })
      .join("");

    yearCardsContainer.querySelectorAll(".year-card").forEach((button) => {
      button.addEventListener("click", () => startQuizForYear(button.dataset.year));
    });
  }

  function startQuizForYear(year) {
    selectedYear = String(year);
    filterStatus = "all";
    filterCaseId = "all";
    searchString = "";

    if (searchTerm) searchTerm.value = "";
    if (statusFilter) statusFilter.value = "all";
    if (caseFilter) caseFilter.value = "all";

    if (landingPage) landingPage.classList.add("hidden");
    if (appShell) appShell.classList.remove("hidden");
    if (backToYearsBtn) backToYearsBtn.classList.remove("hidden");

    syncAppState();
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  function showLandingPage() {
    selectedYear = null;
    if (appShell) appShell.classList.add("hidden");
    if (landingPage) landingPage.classList.remove("hidden");
    if (backToYearsBtn) backToYearsBtn.classList.add("hidden");
    renderYearLandingPage();
  }

  async function autoLoadPLEExcelFiles() {
    const candidateFiles = [
      "PLE_PC1_2563.xlsx",
      "PLE_PC1_2564.xlsx",
      "PLE_PC1_2565.xlsx",
      "PLE_PC1_2566.xlsx",
      "PLE_PC1_2567.xlsx",
      "PLE_PC1_2568.xlsx",
    ];

    const allQuestions = [];

    for (const fileName of candidateFiles) {
      if (!/^PLE_PC1_\d+\.xlsx$/i.test(fileName)) continue;

      try {
        const response = await fetch(fileName, { cache: "no-store" });

        if (!response.ok) {
          console.warn(`ไม่พบไฟล์ ${fileName}`);
          continue;
        }

        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: "array" });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];

        const rows = XLSX.utils.sheet_to_json(worksheet, {
          defval: "",
        });

        const questions = rows
          .map((row, index) => {
            const year = String(row["Year"] || "").trim();
            const caseId = String(row["Case"] || "").trim();
            const questionNo = Number(row["Question No"] || index + 1);

            return {
              year,
              caseId,
              caseDescription: String(row["Case Description"] || "").trim(),
              id: Number(`${year}${String(questionNo).padStart(3, "0")}`),
              questionNo,
              question: String(row["Question"] || "").trim(),
              choices: {
                A: String(row["Choice A"] || "").trim(),
                B: String(row["Choice B"] || "").trim(),
                C: String(row["Choice C"] || "").trim(),
                D: String(row["Choice D"] || "").trim(),
                E: String(row["Choice E"] || "").trim(),
              },
              answer: normalizeAnswer(row["Answer"]),
              explanation: getExplanationFromRow(row),
            };
          })
          .filter((q) => q.question && q.answer);

        allQuestions.push(...questions);

        console.log(`โหลด ${fileName} สำเร็จ: ${questions.length} ข้อ`);
      } catch (err) {
        console.warn(`โหลดไฟล์ ${fileName} ไม่สำเร็จ`, err);
      }
    }

    if (allQuestions.length === 0) {
      popCustomToast(
        "ไม่พบไฟล์ Excel รูปแบบ PLE_PC1_[ปี].xlsx ในโฟลเดอร์เดียวกัน",
        "error",
      );
    } else {
      popCustomToast(
        `โหลดข้อสอบจาก Excel สำเร็จ ${allQuestions.length} ข้อ`,
        "success",
      );
    }

    return allQuestions;
  }
  function normalizeAnswer(value) {
    const answer = String(value || "")
      .trim()
      .toUpperCase();

    const map = {
      ก: "A",
      ข: "B",
      ค: "C",
      ง: "D",
      จ: "E",
      A: "A",
      B: "B",
      C: "C",
      D: "D",
      E: "E",
    };

    return map[answer] || answer;
  }
  function getExplanationFromRow(row) {
    return String(
      row["explaination"] ||
      row["Explaination"] ||
      row["EXPLAINATION"] ||
      row["explanation"] ||
      row["Explanation"] ||
      row["EXPLANATION"] ||
      row["เฉลยละเอียด"] ||
      row["คำอธิบาย"] ||
      "",
    ).trim();
  }

  // --- COMPILING QUERIES AND FILTERING SYSTEMS ---
  function syncAppState(preserveCurrentIndex = false) {
    updateCaseStudyFilterDropdown();

    let workingSet = [];

    rawQuestionsList.forEach((q) => {
      const history = answersState[q.id];
      const isAnswered = !!history;
      const isCorrect = history && history.correct;
      const isBookmarked = bookmarksState.includes(String(q.id));

      // 1. Check Keywords match
      const searchMatch =
        searchString === "" ||
        String(q.id).toLowerCase().includes(searchString) ||
        q.question.toLowerCase().includes(searchString) ||
        (q.caseDescription &&
          q.caseDescription.toLowerCase().includes(searchString));

      // 2. Status drop list filter matches
      let statusMatch = false;
      if (filterStatus === "all") statusMatch = true;
      else if (filterStatus === "unanswered" && !isAnswered)
        statusMatch = true;
      else if (filterStatus === "correct" && isAnswered && isCorrect)
        statusMatch = true;
      else if (filterStatus === "incorrect" && isAnswered && !isCorrect)
        statusMatch = true;
      else if (filterStatus === "bookmarked" && isBookmarked)
        statusMatch = true;

      // 3. Case studies selection filter matches
      let caseMatch = false;
      if (filterCaseId === "all") {
        caseMatch = true;
      } else {
        const qCase = q.caseId ? String(q.caseId).trim() : "none";
        caseMatch = qCase === String(filterCaseId).trim();
      }

      const yearMatch = selectedYear && String(q.year) === String(selectedYear);

      if (
        searchMatch &&
        statusMatch &&
        caseMatch &&
        yearMatch
      ) {
        workingSet.push(q);
      }
    });

    // Set mapped index values
    filteredQuestions = workingSet;

    if (!preserveCurrentIndex) {
      currentIndex = 0;
    } else if (
      currentIndex >= filteredQuestions.length &&
      filteredQuestions.length > 0
    ) {
      currentIndex = filteredQuestions.length - 1;
    }

    renderDashboardScorecard();
    renderNavigatorGridPanel();
    renderActiveQuestionLayout();
  }

  // --- PRESENTATION COMPONENT RENDER ENGINES ---
  function renderActiveQuestionLayout() {
    if (filteredQuestions.length === 0) {
      // Empty fallback layout
      caseScenarioCard.classList.add("hidden");
      activeQuestionText.innerHTML = `
                        <div class="text-center py-12 text-slate-400">
                            <i class="fa-regular fa-folder-open text-5xl mb-3 block"></i>
                            <p class="font-bold text-base">ไม่พบข้อมูลข้อสอบ</p>
                            <p class="text-xs text-slate-400 mt-1">ลองพิมพ์กล่องค้นหาใหม่ สลับฟิลเตอร์คัดเลือก หรือป้อนชุดข้อสอบของคุณเพิ่มได้</p>
                        </div>
                    `;
      choicesContainer.innerHTML = "";
      explanationBox.classList.add("hidden");
      displayIndex.innerText = "0";
      displayTotal.innerText = "0";
      questionIdBadge.innerText = "";
      prevQuestionBtn.disabled = true;
      nextQuestionBtn.disabled = true;
      return;
    }

    const q = filteredQuestions[currentIndex];
    displayIndex.innerText = currentIndex + 1;
    displayTotal.innerText = filteredQuestions.length;
    questionIdBadge.innerText =
      `(ปี ${q.year || "-"} / Case ${q.caseId || "-"} / ข้อจริง ${q.questionNo || q.id})`;
    activeQuestionText.innerText = q.question;

    // Manage Patient Case display
    if (q.caseDescription && q.caseDescription.trim() !== "") {
      caseScenarioCard.classList.remove("hidden");
      caseHeaderLabel.innerText = `Case ${q.caseId || "ทั่วไป"} ${q.year ? "- ปี " + q.year : ""}`;
      caseDescriptionText.innerHTML = formatCaseDescription(q.caseDescription);
    } else {
      caseScenarioCard.classList.add("hidden");
    }

    // Bookmark visual star toggle
    const isBookmarked = bookmarksState.includes(String(q.id));
    if (isBookmarked) {
      bookmarkButton.className =
        "p-2 rounded-xl bg-amber-500 hover:bg-amber-600 text-white transition shadow-sm";
      bookmarkButton.innerHTML = `<i class="fa-solid fa-bookmark text-sm"></i>`;
    } else {
      bookmarkButton.className =
        "p-2 rounded-xl bg-slate-100 dark:bg-slate-700 hover:bg-slate-200 text-slate-400 dark:text-slate-300 transition";
      bookmarkButton.innerHTML = `<i class="fa-regular fa-bookmark text-sm"></i>`;
    }

    // Create Choices list ordering
    const history = answersState[q.id];
    let activeChoicesOrdering = ["A", "B", "C", "D", "E"];

    if (history && history.choiceMapping) {
      activeChoicesOrdering = history.choiceMapping;
    } else if (shuffleChoicesMode) {
      activeChoicesOrdering.sort(() => Math.random() - 0.5);
    }

    choicesContainer.innerHTML = "";
    const isLocked = history && history.selected !== undefined;

    activeChoicesOrdering.forEach((key) => {
      const text = q.choices[key];
      if (!text) return; // Prevent render null parameters

      const button = document.createElement("button");
      button.type = "button";
      button.className =
        "w-full text-left p-4 rounded-2xl border text-sm sm:text-base font-medium flex items-center justify-between transition-all duration-200 outline-none ";

      if (!isLocked) {
        button.className +=
          "border-slate-200 dark:border-slate-700 hover:border-indigo-400 hover:bg-indigo-50/20 dark:hover:bg-indigo-950/20 cursor-pointer text-slate-700 dark:text-slate-200";
        button.innerHTML = `
                            <div class="flex items-center space-x-3">
                                <span class="w-8 h-8 rounded-lg bg-slate-100 dark:bg-slate-700 dark:text-white flex items-center justify-center font-bold text-xs shrink-0">${key}</span>
                                <span class="leading-relaxed pr-2">${text}</span>
                            </div>
                        `;
        button.addEventListener("click", () =>
          registerAnswerSubmit(q.id, key, activeChoicesOrdering),
        );
      } else {
        button.disabled = true;
        button.className += "cursor-default ";

        const isChosen = history.selected === key;
        const isCorrectAnswer = q.answer === key;

        if (isCorrectAnswer) {
          button.className +=
            "border-emerald-500 bg-emerald-50 dark:bg-emerald-950/30 text-emerald-900 dark:text-emerald-200 font-bold ring-2 ring-emerald-500/20";
          button.innerHTML = `
                                <div class="flex items-center space-x-3">
                                    <span class="w-8 h-8 rounded-lg bg-emerald-500 text-white flex items-center justify-center font-bold text-xs shrink-0">${key}</span>
                                    <span class="leading-relaxed pr-2">${text}</span>
                                </div>
                                <span class="text-emerald-600 dark:text-emerald-400 shrink-0 text-sm font-semibold flex items-center space-x-1">
                                    <i class="fa-solid fa-check-circle text-lg animate-bounce"></i>
                                </span>
                            `;
        } else if (isChosen) {
          button.className +=
            "border-rose-500 bg-rose-50 dark:bg-rose-950/30 text-rose-900 dark:text-rose-200 font-bold ring-2 ring-rose-500/20";
          button.innerHTML = `
                                <div class="flex items-center space-x-3">
                                    <span class="w-8 h-8 rounded-lg bg-rose-500 text-white flex items-center justify-center font-bold text-xs shrink-0">${key}</span>
                                    <span class="leading-relaxed pr-2">${text}</span>
                                </div>
                                <span class="text-rose-600 dark:text-rose-400 shrink-0 text-sm font-semibold flex items-center space-x-1">
                                    <i class="fa-solid fa-circle-xmark text-lg"></i>
                                </span>
                            `;
        } else {
          button.className +=
            "border-slate-200 dark:border-slate-700 opacity-60 text-slate-500 dark:text-slate-400";
          button.innerHTML = `
                                <div class="flex items-center space-x-3">
                                    <span class="w-8 h-8 rounded-lg bg-slate-100 dark:bg-slate-700 dark:text-slate-400 flex items-center justify-center font-bold text-xs shrink-0">${key}</span>
                                    <span class="leading-relaxed pr-2">${text}</span>
                                </div>
                            `;
        }
      }

      choicesContainer.appendChild(button);
    });

    // Display explanations drawer if answered
    if (isLocked) {
      const correct = history.correct;
      explanationBox.className = correct
        ? "rounded-2xl p-5 mb-6 bg-emerald-50 dark:bg-emerald-950/20 border border-emerald-200 dark:border-emerald-800/40 text-slate-700 dark:text-slate-300 transition-all duration-300 animate-fadeIn"
        : "rounded-2xl p-5 mb-6 bg-rose-50 dark:bg-rose-950/20 border border-rose-200 dark:border-rose-800/40 text-slate-700 dark:text-slate-300 transition-all duration-300 animate-fadeIn";

      const explainDesc =
        q.explanation && q.explanation.trim() !== ""
          ? q.explanation
          : "ไม่มีข้อมูลคำอธิบายอ้างอิงเพิ่มเติมสำหรับข้อสอบข้อนี้";

      explanationBox.innerHTML = `
                        <div class="flex items-start space-x-3">
                            <div class="p-2 rounded-xl shrink-0 ${correct ? "bg-emerald-100 dark:bg-emerald-900/50 text-emerald-600 dark:text-emerald-400" : "bg-rose-100 dark:bg-rose-900/50 text-rose-600 dark:text-rose-400"}">
                                <i class="${correct ? "fa-solid fa-circle-check" : "fa-solid fa-circle-xmark"} text-xl"></i>
                            </div>
                            <div class="space-y-1 w-full">
                                <h4 class="font-bold text-sm ${correct ? "text-emerald-800 dark:text-emerald-400" : "text-rose-800 dark:text-rose-400"}">
                                    ${correct ? "ตอบถูกต้อง!" : "ตอบผิดพลาด!"}
                                </h4>
                                <p class="text-xs text-slate-400 dark:text-slate-500 font-sarabun">เฉลยคือกุญแจข้อ: <span class="font-extrabold uppercase">${q.answer}</span></p>
                                <p class="text-sm leading-relaxed text-slate-600 dark:text-slate-300 mt-2 font-medium bg-white/40 dark:bg-slate-900/40 p-3.5 rounded-lg border border-slate-100 dark:border-slate-800/40">${explainDesc}</p>
                            </div>
                        </div>
                    `;
      explanationBox.classList.remove("hidden");
    } else {
      explanationBox.classList.add("hidden");
    }

    prevQuestionBtn.disabled = currentIndex === 0;
    nextQuestionBtn.disabled =
      currentIndex === filteredQuestions.length - 1;

    highlightActiveGridCell(currentIndex);
  }

  function registerAnswerSubmit(qId, selectedKey, choicesOrderingMap) {
    const q = rawQuestionsList.find((item) => item.id === qId);
    if (!q) return;

    const isCorrect = q.answer === selectedKey;

    answersState[qId] = {
      selected: selectedKey,
      correct: isCorrect,
      choiceMapping: choicesOrderingMap,
    };

    saveToLocalStorage();
    renderDashboardScorecard();
    renderNavigatorGridPanel();
    renderActiveQuestionLayout();
  }

  // --- DASHBOARD AND PROGRESS METRICS RENDER ---
  function renderDashboardScorecard() {
    let answered = 0;
    let correct = 0;
    let wrong = 0;

    const scoreQuestions = selectedYear
      ? rawQuestionsList.filter((q) => String(q.year) === String(selectedYear))
      : rawQuestionsList;

    scoreQuestions.forEach((q) => {
      const hist = answersState[q.id];
      if (hist && hist.selected !== undefined) {
        answered++;
        if (hist.correct) correct++;
        else wrong++;
      }
    });

    const total = scoreQuestions.length;
    statsAnswered.innerText = `${answered}/${total}`;
    statsCorrect.innerText = correct;
    statsWrong.innerText = wrong;

    const percentage =
      answered > 0 ? Math.round((correct / answered) * 100) : 0;
    statsPercent.innerText = `${percentage}%`;

    const percentageProgressWidth =
      total > 0 ? (answered / total) * 100 : 0;
    linearProgressBar.style.width = `${percentageProgressWidth}%`;
  }

  // --- RENDER SIDEBAR GRID BUTTONS PANEL ---
  function renderNavigatorGridPanel() {
    gridNavigator.innerHTML = "";
    const yearTotal = selectedYear ? rawQuestionsList.filter((q) => String(q.year) === String(selectedYear)).length : rawQuestionsList.length;
    renderedGridCount.innerText = `${filteredQuestions.length} / ${yearTotal}`;

    filteredQuestions.forEach((q, idx) => {
      const history = answersState[q.id];
      const isAnswered = !!history;
      const isCorrect = history && history.correct;

      const cell = document.createElement("button");
      cell.type = "button";
      cell.innerText = idx + 1;
      cell.title = `ปี ${q.year || "-"} / ข้อจริง ${q.questionNo || q.id}`;
      cell.id = `nav-grid-${idx}`;
      cell.className =
        "p-2.5 text-xs font-bold rounded-xl border transition-all duration-150 text-center cursor-pointer outline-none ";

      if (!isAnswered) {
        cell.className +=
          "bg-slate-100 dark:bg-slate-800 border-slate-200 dark:border-slate-700 text-slate-700 dark:text-slate-300 hover:border-indigo-400";
      } else if (isCorrect) {
        cell.className +=
          "bg-emerald-500 border-emerald-600 text-white hover:bg-emerald-600";
      } else {
        cell.className +=
          "bg-rose-500 border-rose-600 text-white hover:bg-rose-600";
      }

      cell.addEventListener("click", () => {
        currentIndex = idx;
        renderActiveQuestionLayout();
      });

      gridNavigator.appendChild(cell);
    });
  }

  function highlightActiveGridCell(idx) {
    document.querySelectorAll('[id^="nav-grid-"]').forEach((el) => {
      el.classList.remove(
        "ring-4",
        "ring-indigo-500/40",
        "border-indigo-500",
        "dark:border-indigo-400",
      );
    });

    const targetCell = document.getElementById(`nav-grid-${idx}`);
    if (targetCell) {
      targetCell.classList.add(
        "ring-4",
        "ring-indigo-500/40",
        "border-indigo-500",
        "dark:border-indigo-400",
      );
      targetCell.scrollIntoView({ block: "nearest", behavior: "smooth" });
    }
  }

  // --- DYNAMIC SELECTOR UPDATE ENGINES ---
  function updateCaseStudyFilterDropdown() {
    let cases = new Set();
    rawQuestionsList
      .filter((q) => !selectedYear || String(q.year) === String(selectedYear))
      .forEach((q) => {
      if (q.caseId && String(q.caseId).trim() !== "") {
        cases.add(String(q.caseId).trim());
      }
    });

    const sortedCases = Array.from(cases).sort(
      (a, b) => Number(a) - Number(b) || a.localeCompare(b),
    );
    const selectedVal = caseFilter.value;

    caseFilter.innerHTML = '<option value="all">ทุก Case</option>';
    sortedCases.forEach((c) => {
      const opt = document.createElement("option");
      opt.value = c;
      opt.innerText = `Case ${c}`;
      caseFilter.appendChild(opt);
    });

    if (sortedCases.includes(selectedVal)) {
      caseFilter.value = selectedVal;
    } else {
      caseFilter.value = "all";
      filterCaseId = "all";
    }
  }

  // --- ACTIONS CLICK LISTENERS ---

  // Bookmark Trigger
  bookmarkButton.addEventListener("click", () => {
    if (filteredQuestions.length === 0) return;
    const q = filteredQuestions[currentIndex];
    const qIdStr = String(q.id);

    const indexInBookmarks = bookmarksState.indexOf(qIdStr);
    if (indexInBookmarks > -1) {
      bookmarksState.splice(indexInBookmarks, 1);
    } else {
      bookmarksState.push(qIdStr);
    }

    saveToLocalStorage();
    renderActiveQuestionLayout();

    if (filterStatus === "bookmarked") {
      syncAppState(true);
    }
  });

  // Sliding questions pages
  prevQuestionBtn.addEventListener("click", () => {
    if (currentIndex > 0) {
      currentIndex--;
      renderActiveQuestionLayout();
    }
  });

  nextQuestionBtn.addEventListener("click", () => {
    if (currentIndex < filteredQuestions.length - 1) {
      currentIndex++;
      renderActiveQuestionLayout();
    }
  });

  // Keyword searches and dropdowns filters changes listeners
  searchTerm.addEventListener("input", (e) => {
    searchString = e.target.value.toLowerCase().trim();
    syncAppState();
  });

  statusFilter.addEventListener("change", (e) => {
    filterStatus = e.target.value;
    syncAppState();
  });

  caseFilter.addEventListener("change", (e) => {
    filterCaseId = e.target.value;
    syncAppState();
  });

  // Question randomizer shuffling
  shuffleQsBtn.addEventListener("click", () => {
    if (filteredQuestions.length <= 1) return;
    filteredQuestions.sort(() => Math.random() - 0.5);
    currentIndex = 0;
    renderActiveQuestionLayout();
    renderNavigatorGridPanel();
    popCustomToast("สุ่มปรับลำดับรายการข้อสอบเรียบร้อยแล้ว!", "success");
  });

  // Shuffling choice key triggers
  shuffleChoicesBtn.addEventListener("click", () => {
    shuffleChoicesMode = !shuffleChoicesMode;
    if (shuffleChoicesMode) {
      shuffleChoicesBtn.className =
        "flex items-center justify-center space-x-1.5 py-2.5 px-3 bg-indigo-50 dark:bg-indigo-950 text-indigo-600 dark:text-indigo-400 rounded-xl text-xs font-bold transition border border-indigo-200 dark:border-indigo-900/30";
      shuffleChoicesText.innerText = "สลับชอยส์: เปิด";
      popCustomToast("เปิดการทำงานสลับตัวเลือกสำเร็จ!", "info");
    } else {
      shuffleChoicesBtn.className =
        "flex items-center justify-center space-x-1.5 py-2.5 px-3 bg-slate-100 hover:bg-slate-200 dark:bg-slate-700 dark:hover:bg-slate-600 rounded-xl text-xs font-bold transition";
      shuffleChoicesText.innerText = "สลับชอยส์: ปิด";
      popCustomToast("ปิดฟังก์ชันสลับตัวเลือก", "neutral");
    }
    saveToLocalStorage();
    renderActiveQuestionLayout();
  });

  // --- MANUAL CASE EDITOR MODAL WORKFLOW ---
  openAddModalBtn.addEventListener("click", () => {
    modalTitle.innerText = "เพิ่มกรณีศึกษา / โจทย์ชุดใหม่ด้วยตนเอง";
    caseForm.reset();
    editOriginalNo.value = "";

    let nextId = 1;
    if (rawQuestionsList.length > 0) {
      const mapIds = rawQuestionsList
        .map((q) => Number(q.id))
        .filter((id) => !isNaN(id));
      if (mapIds.length > 0) {
        nextId = Math.max(...mapIds) + 1;
      }
    }
    formQuestionNo.value = nextId;
    formYear.value = "2563"; // default baseline

    if (
      filteredQuestions.length > 0 &&
      currentIndex < filteredQuestions.length
    ) {
      const cur = filteredQuestions[currentIndex];
      if (cur.caseId) {
        formCaseId.value = cur.caseId;
        formCaseDescription.value = cur.caseDescription || "";
        formYear.value = cur.year || "2563";
      }
    }

    toggleModalState(true);
  });

  editQuestionBtn.addEventListener("click", () => {
    if (filteredQuestions.length === 0) return;
    const q = filteredQuestions[currentIndex];

    modalTitle.innerText = `แก้ไขรายละเอียดข้อสอบข้อที่ ${q.id}`;
    editOriginalNo.value = q.id;

    // Fill elements
    formYear.value = q.year || "";
    formCaseId.value = q.caseId || "";
    formCaseDescription.value = q.caseDescription || "";
    formQuestionNo.value = q.id;
    formQuestionText.value = q.question || "";
    formChoiceA.value = q.choices.A || "";
    formChoiceB.value = q.choices.B || "";
    formChoiceC.value = q.choices.C || "";
    formChoiceD.value = q.choices.D || "";
    formChoiceE.value = q.choices.E || "";
    formAnswer.value = q.answer || "A";
    formExplanation.value = q.explanation || "";

    toggleModalState(true);
  });

  function toggleModalState(isOpen) {
    if (isOpen) {
      caseEditorModal.classList.remove("hidden");
      setTimeout(() => {
        modalContainer.classList.remove("scale-95", "opacity-0");
        modalContainer.classList.add("scale-100", "opacity-100");
      }, 50);
    } else {
      modalContainer.classList.remove("scale-100", "opacity-100");
      modalContainer.classList.add("scale-95", "opacity-0");
      setTimeout(() => {
        caseEditorModal.classList.add("hidden");
      }, 150);
    }
  }

  closeModalBtn.addEventListener("click", () => toggleModalState(false));
  cancelFormBtn.addEventListener("click", () => toggleModalState(false));

  caseForm.addEventListener("submit", (e) => {
    e.preventDefault();

    const inputNumId = Number(formQuestionNo.value);
    const editMarkerValue = editOriginalNo.value;

    const payload = {
      year: formYear.value.trim(),
      caseId: formCaseId.value.trim(),
      caseDescription: formCaseDescription.value.trim(),
      id: inputNumId,
      question: formQuestionText.value.trim(),
      choices: {
        A: formChoiceA.value.trim(),
        B: formChoiceB.value.trim(),
        C: formChoiceC.value.trim(),
        D: formChoiceD.value.trim(),
        E: formChoiceE.value.trim(),
      },
      answer: formAnswer.value,
      explanation: formExplanation.value.trim(),
    };

    // Propagate case descriptions updates if editing a shared case ID
    if (payload.caseId !== "") {
      rawQuestionsList.forEach((q) => {
        if (String(q.caseId).trim() === String(payload.caseId).trim()) {
          q.caseDescription = payload.caseDescription;
          q.year = payload.year;
        }
      });
    }

    if (editMarkerValue !== "") {
      const origId = Number(editMarkerValue);
      const idx = rawQuestionsList.findIndex((q) => q.id === origId);
      if (idx > -1) {
        rawQuestionsList[idx] = payload;
        popCustomToast(
          `แก้ไขเนื้อหาข้อสอบข้อที่ ${inputNumId} สำเร็จ!`,
          "success",
        );
      }
    } else {
      const duplicateIndex = rawQuestionsList.findIndex(
        (q) => q.id === inputNumId,
      );
      if (duplicateIndex > -1) {
        const confirmOverwrite = confirm(
          `รหัสข้อคำถามเลขที่ ${inputNumId} ซ้ำซ้อน มีอยู่อีกข้อในฐานข้อมูลหลัก ต้องการกดทับเขียนใหม่ใช่หรือไม่?`,
        );
        if (!confirmOverwrite) return;
        rawQuestionsList[duplicateIndex] = payload;
      } else {
        rawQuestionsList.push(payload);
      }
      popCustomToast(
        `เพิ่มโจทย์คำถามข้อที่ ${inputNumId} เข้าระบบคลังเรียบร้อย!`,
        "success",
      );
    }

    saveToLocalStorage();
    toggleModalState(false);

    if (editMarkerValue !== "") {
      syncAppState(true);
    } else {
      syncAppState();
      const matchedIdx = filteredQuestions.findIndex(
        (q) => q.id === inputNumId,
      );
      if (matchedIdx > -1) {
        currentIndex = matchedIdx;
        renderActiveQuestionLayout();
      }
    }
  });
  function formatCaseDescription(text) {
    if (!text) return "";

    const lines = text
      .split(/•|\n/)
      .map(s => s.trim())
      .filter(Boolean);

    if (lines.length <= 1) {
      return text.replace(/\n/g, "<br>");
    }

    const intro = lines.shift();

    return `
            <p class="mb-4 leading-7">${intro}</p>

            <ul class="space-y-2 list-disc pl-6">
              ${lines.map(line => `<li>${line}</li>`).join("")}
            </ul>
          `;
  }
  // --- SHEETJS SPREADSHEET READ AND INTERPRETING PIPELINE ---
  function parseSpreadsheetFile(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function (evt) {
      try {
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const rows = XLSX.utils.sheet_to_json(worksheet);

        if (rows.length === 0) {
          popCustomToast(
            "ไฟล์สเปรดชีตไม่มีข้อมูลโจทย์วิเคราะห์ใด ๆ",
            "error",
          );
          return;
        }

        const alphabetMapTranslate = {
          ก: "A",
          ข: "B",
          ค: "C",
          ง: "D",
          จ: "E",
        };

        // Fuzzy property resolver helper to match case-insensitive headers
        function resolveVal(row, possibleKeys) {
          for (let k of Object.keys(row)) {
            const normalized = k.toLowerCase().replace(/[\s\_\-\.]/g, "");
            for (let target of possibleKeys) {
              if (normalized === target || normalized.includes(target)) {
                return row[k];
              }
            }
          }
          return "";
        }

        const parsed = rows.map((row, idx) => {
          // Extract properties with flexible alias headers
          const rawId = resolveVal(row, [
            "questionno",
            "no",
            "ลำดับ",
            "เลขข้อ",
            "ข้อที่",
            "qno",
          ]);
          const qNo = rawId ? Number(rawId) : idx + 1;

          const questionText = resolveVal(row, [
            "question",
            "โจทย์",
            "คำถาม",
            "โจทย์คำถาม",
          ]);
          const yearVal = resolveVal(row, ["year", "ปี", "พศ"]);
          const caseNo = resolveVal(row, [
            "case",
            "เคส",
            "เลขเคส",
            "caseid",
          ]);
          const caseDesc = resolveVal(row, [
            "casedescription",
            "description",
            "รายละเอียดเคส",
            "สถานการณ์",
          ]);
          const expVal = resolveVal(row, [
            "explanation",
            "อธิบาย",
            "เฉลยละเอียด",
            "คำอธิบาย",
          ]);

          // Choices mapping aliases
          const choiceA = resolveVal(row, [
            "choicea",
            "a",
            "ตัวเลือกa",
            "ก",
            "ตัวเลือกก",
          ]);
          const choiceB = resolveVal(row, [
            "choiceb",
            "b",
            "ตัวเลือกb",
            "ข",
            "ตัวเลือกข",
          ]);
          const choiceC = resolveVal(row, [
            "choicec",
            "c",
            "ตัวเลือกc",
            "ค",
            "ตัวเลือกค",
          ]);
          const choiceD = resolveVal(row, [
            "choiced",
            "d",
            "ตัวเลือกd",
            "ง",
            "ตัวเลือกง",
          ]);
          const choiceE = resolveVal(row, [
            "choicee",
            "e",
            "ตัวเลือกe",
            "จ",
            "ตัวเลือกจ",
          ]);

          let ans = String(
            resolveVal(row, ["answer", "เฉลย", "คำตอบ"]) || "",
          ).trim();
          if (alphabetMapTranslate[ans]) {
            ans = alphabetMapTranslate[ans];
          }
          ans = ans.toUpperCase();

          return {
            year: yearVal ? String(yearVal).trim() : "2563",
            caseId: caseNo ? String(caseNo).trim() : "",
            caseDescription: caseDesc ? String(caseDesc).trim() : "",
            id: qNo,
            question: questionText || "ไม่มีข้อความโจทย์คำถามวิเคราะห์",
            choices: {
              A: choiceA || "",
              B: choiceB || "",
              C: choiceC || "",
              D: choiceD || "",
              E: choiceE || "",
            },
            answer: ans,
            explanation: expVal || "",
          };
        });

        // Overwrite memory catalogs and wipe progress
        rawQuestionsList = parsed;
        answersState = {};
        bookmarksState = [];
        saveToLocalStorage();

        renderYearLandingPage();
        showLandingPage();
        popCustomToast(
          `นำเข้าสำเร็จ! ตรวจพบข้อมูลข้อสอบทั้งสิ้น ${parsed.length} ข้อ`,
          "success",
        );
      } catch (err) {
        console.error("Reader Error: ", err);
        popCustomToast(
          "โครงสร้างไฟล์ล้มเหลว กรุณาจัดเรียงคอลัมน์ใหม่",
          "error",
        );
      }
    };
    reader.readAsArrayBuffer(file);
  }

  // --- LOCAL STORAGE PERSISTENCE SYSTEMS ---
  function saveToLocalStorage() {
    localStorage.setItem(
      "local_pharmacy_answersState_v4",
      JSON.stringify(answersState),
    );
    localStorage.setItem(
      "local_pharmacy_bookmarksState_v4",
      JSON.stringify(bookmarksState),
    );
    localStorage.setItem(
      "local_pharmacy_shuffleChoices_v4",
      JSON.stringify(shuffleChoicesMode),
    );
  }

  function loadProgressOnlyFromLocalStorage() {
    const answers = localStorage.getItem("local_pharmacy_answersState_v4");
    const bookmarks = localStorage.getItem(
      "local_pharmacy_bookmarksState_v4",
    );
    const shuffle = localStorage.getItem(
      "local_pharmacy_shuffleChoices_v4",
    );

    if (answers) answersState = JSON.parse(answers);
    if (bookmarks) bookmarksState = JSON.parse(bookmarks);

    if (shuffle) {
      shuffleChoicesMode = JSON.parse(shuffle);
    }
  }
  // Wipe database cache and reload defaults
  clearStateBtn.addEventListener("click", async () => {
    const confirmed = confirm(
      "คุณแน่ใจใช่หรือไม่ว่าต้องการล้างประวัติการทำควิซและบุ๊กมาร์กทั้งหมด?",
    );

    if (confirmed) {
      localStorage.removeItem("local_pharmacy_answersState_v4");
      localStorage.removeItem("local_pharmacy_bookmarksState_v4");
      localStorage.removeItem("local_pharmacy_shuffleChoices_v4");
      localStorage.removeItem("local_pharmacy_questionsSet_v4");

      answersState = {};
      bookmarksState = [];
      shuffleChoicesMode = false;

      rawQuestionsList = await autoLoadPLEExcelFiles();
      renderYearLandingPage();
      showLandingPage();

      popCustomToast("ล้างความคืบหน้าและโหลด Excel ใหม่แล้ว", "success");
    }
  });
  // --- MICRO-INTERACTIONS TOAST NOTIFICATION ---
  function popCustomToast(message, type = "neutral") {
    const oldToast = document.getElementById("app-toast");
    if (oldToast) oldToast.remove();

    const toast = document.createElement("div");
    toast.id = "app-toast";
    toast.className =
      "fixed bottom-5 right-5 z-50 px-5 py-3.5 rounded-2xl shadow-xl border text-sm font-bold flex items-center space-x-2.5 animate-bounce transition-all duration-300 ";

    if (type === "success") {
      toast.className += "bg-emerald-500 border-emerald-600 text-white";
      toast.innerHTML = `<i class="fa-solid fa-circle-check text-lg"></i> <span>${message}</span>`;
    } else if (type === "error") {
      toast.className += "bg-rose-500 border-rose-600 text-white";
      toast.innerHTML = `<i class="fa-solid fa-triangle-exclamation text-lg"></i> <span>${message}</span>`;
    } else if (type === "info") {
      toast.className += "bg-indigo-600 border-indigo-700 text-white";
      toast.innerHTML = `<i class="fa-solid fa-circle-info text-lg"></i> <span>${message}</span>`;
    } else {
      toast.className +=
        "bg-slate-800 border-slate-900 text-white dark:bg-slate-700 dark:border-slate-600";
      toast.innerHTML = `<i class="fa-solid fa-bell text-lg"></i> <span>${message}</span>`;
    }

    document.body.appendChild(toast);
    setTimeout(() => {
      toast.classList.add("opacity-0");
      setTimeout(() => toast.remove(), 300);
    }, 3000);
  }

  if (backToYearsBtn) {
    backToYearsBtn.addEventListener("click", showLandingPage);
  }

  // --- DARK MODE THEME CONTROLS ---
  themeToggler.addEventListener("click", () => {
    if (document.documentElement.classList.contains("dark")) {
      document.documentElement.classList.remove("dark");
      localStorage.setItem("theme", "light");
      themeIcon.className = "fa-solid fa-moon text-sm";
      popCustomToast("สลับเข้าสู่โหมดโทนสีสว่าง (Light Mode)", "neutral");
    } else {
      document.documentElement.classList.add("dark");
      localStorage.setItem("theme", "dark");
      themeIcon.className = "fa-solid fa-sun text-sm";
      popCustomToast("สลับเข้าสู่โหมดโทนสีมืด (Dark Mode)", "neutral");
    }
  });

  initApp();
});
