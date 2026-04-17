const SHEET_ID = "1IOY3D3CYQ1R9aYF6QooMucWhYDgSOzGzeHzbSj9i4J8";
const SHEET_NAME = "시트2";
const API_KEY = "AIzaSyDqD4ECDE5xaoBeEoRyicDaqUJ5RN-dUgs";
const CLIENT_ID = "44215227073-q4rnmebvgdklojddc43dsr1r7gm2qlrc.apps.googleusercontent.com";
const CLIENT_SECRET = "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";
const REFRESH_TOKEN = "1//04FjUOlf9xwIiCgYIARAAGAQSNwF-L9Ir5665xGlWI9AO7uWx6DHORfJLypgRnBM44W-BRre85Vkis6UV_NKZHXRjuArPdfwA_wA";

const MEMBERS = [
  "롱게임", "비비드드림", "꿈요셉", "사쁘니", "행복소금",
  "미쁨댁", "우부왕", "금만가", "플리퍼즈", "스테디 릴리",
  "올댓드림", "오늘날씨맑음", "꿀람쥐", "꿈을 현실로", "뚜연",
  "라난", "럭셔리 노블", "레리꼬", "부자될또정", "하예", "뿌래랠래", "자산일조", "산골지야"
];

const CHALLENGE = {
  title: "고전이 답했다: 마땅히 가져야 할 부에 대하여",
  startDate: "2026-03-27",
  chaptersPerDay: 2,
  chapters: [
    "금은 원래 흙이었다",
    "개츠비와 이노크, 두 가지 죽음 앞에서",
    "기회는 반드시 위기 속에서 나온다",
    "10년 주기론",
    "돈이 좋아하는 것들",
    "오늘의 법칙",
    "좋은 돈과 나쁜 돈",
    "부자가 되는 건 일거리가 달라지는 것",
    "자발적 피로를 선택하라",
    "부시맨의 콜라병",
    "무엇을 기다리는가",
    "경기가 좋지 않을 때 해야 하는 것",
    "가장 절망적인 악덕은 무지다",
    "돈을 부르는 자세가 있다",
    "낚시로부터 배운 것",
    "1달러를 벌어보자",
    "경쟁하지 말고 독점하는 법",
    "아무도 거들떠보지 않는 곳에 돈이 있다",
    "생각과 경험을 팔아야 큰돈을 벌 수 있다",
    "소비자가 아닌 생산자가 되어라",
    "위대한 3분의 법칙",
    "결국, 한 단어를 찾는 힘이다",
    "성공을 설명하는 하나의 단어",
    "이기적인 마음을 이용하라",
    "우연한 기회에 발견하는 것",
    "당연하다는 말의 의미",
    "일을 대하는 태도",
    "끈기의 뜻",
    "한 우물을 팔 것인가, 여러 우물을 팔 것인가",
    "투자의 5계명",
    "공부하고, 투자하라, 그리고 기다려라",
    "당신의 ‘곰스크’는 어디에 있는가",
    "변하지 않는 성공의 단 한 가지 법칙",
    "무한히 애쓸 수 있는 능력",
    "돈 버는 습관: 어떤, 행위를, 저절로",
    "일론 머스크에게는 있고, 당신에게는 없는 것",
    "호리병이 아닌 대접에 담을 것",
    "미래를 예측하는 법",
    "근로 소득을 높이는 방법",
    "‘이곳’에서 ‘저곳’으로 넘어가는 원리",
    "뻔하게 사는 게 정답이다",
    "2상상의 거인을 키워라",
    "“누구나 다 그렇게 될 수는 없잖아요?”",
    "고통을 이기지 못하면 고통이 그대를 이길 것이다",
    "비밀의 개수와 부는 비례한다"
  ]
};

const SPECIAL_ROWS = ["챌린지 성공", "챌린지 성공률"];
const DAY_MS = 24 * 60 * 60 * 1000;

let sheetData = {};
let tokenCache = { accessToken: "", tokenExpiry: 0 };
let activeTab = "today";

function clamp(value, min, max) {
  return Math.min(Math.max(value, min), max);
}

function parseLocalDate(value) {
  const [year, month, day] = value.split("-").map(Number);
  return new Date(year, month - 1, day);
}

function addDays(date, days) {
  const next = new Date(date);
  next.setDate(next.getDate() + days);
  return next;
}

function formatDateInput(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function formatShortDate(date) {
  return `${date.getMonth() + 1}/${date.getDate()}`;
}

function formatKoreanDate(date) {
  return `${date.getMonth() + 1}월 ${date.getDate()}일 ${getWeekdayLabel(date)}`;
}

function getWeekdayLabel(date) {
  return ["일요일", "월요일", "화요일", "수요일", "목요일", "금요일", "토요일"][date.getDay()];
}

function getPreviousDateString(dateString) {
  return formatDateInput(addDays(parseLocalDate(dateString), -1));
}

function getTodayLocal() {
  return formatDateInput(new Date());
}

function todayString() {
  return getTodayLocal();
}

function switchTab(tab) {
  activeTab = tab;
  document.querySelectorAll(".tab").forEach(button => {
    button.classList.toggle("active", button.dataset.tab === tab);
  });
  document.querySelectorAll(".tab-content").forEach(section => {
    section.classList.toggle("active", section.id === `tab-${tab}`);
  });
  if (tab === "stats") {
    renderStats();
  }
}

function initTabs() {
  document.querySelectorAll(".tab").forEach(button => {
    button.addEventListener("click", () => switchTab(button.dataset.tab));
  });
}

function getSelectedDate() {
  return document.getElementById("date-picker").value;
}

function normalizeManualDateInput(value) {
  if (!value) return "";
  const text = String(value).trim();

  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return text;

  const fullDateMatch = text.match(/^(\d{4})[./](\d{1,2})[./](\d{1,2})$/);
  if (fullDateMatch) {
    const [, year, month, day] = fullDateMatch;
    return `${year}-${String(Number(month)).padStart(2, "0")}-${String(Number(day)).padStart(2, "0")}`;
  }

  const shortDateMatch = text.match(/^(\d{1,2})[./](\d{1,2})$/);
  if (shortDateMatch) {
    const [, month, day] = shortDateMatch;
    const baseYear = parseLocalDate(getSelectedDate() || todayString()).getFullYear();
    return `${baseYear}-${String(Number(month)).padStart(2, "0")}-${String(Number(day)).padStart(2, "0")}`;
  }

  return "";
}

function openDatePicker(picker) {
  try {
    if (typeof picker.showPicker === "function") {
      picker.showPicker();
      return;
    }
  } catch (error) {
    console.warn("showPicker 호출 실패", error);
  }

  const manualInput = window.prompt(
    "기준 날짜를 입력해 주세요. 예: 2026-04-10 또는 4/10",
    picker.value || todayString()
  );

  if (manualInput === null) return;

  const normalized = normalizeManualDateInput(manualInput);
  if (!normalized) {
    showToast("날짜 형식은 2026-04-10 또는 4/10으로 입력해 주세요.", "error");
    return;
  }

  picker.value = normalized;
  renderAll();
}

function initDatePicker() {
  const picker = document.getElementById("date-picker");
  picker.value = todayString();
  picker.addEventListener("input", renderAll);
  picker.addEventListener("change", renderAll);
}

function renderHeaderChrome() {
  return;
}

function getChallengeState(dateString) {
  const selectedDate = parseLocalDate(dateString);
  const startDate = parseLocalDate(CHALLENGE.startDate);
  const totalChapters = CHALLENGE.chapters.length;
  const plannedDays = Math.ceil(totalChapters / CHALLENGE.chaptersPerDay);
  const finishDate = addDays(startDate, plannedDays - 1);

  const diffDays = Math.floor((selectedDate - startDate) / DAY_MS) + 1;
  const beforeStart = diffDays <= 0;
  const elapsedDays = clamp(diffDays, 0, plannedDays);
  const completedCount = clamp(elapsedDays * CHALLENGE.chaptersPerDay, 0, totalChapters);
  const currentStartIndex = clamp((elapsedDays - 1) * CHALLENGE.chaptersPerDay, 0, totalChapters);
  const currentEndIndex = clamp(currentStartIndex + CHALLENGE.chaptersPerDay, 0, totalChapters);
  const todayChapters = beforeStart || completedCount === totalChapters && selectedDate > finishDate
    ? []
    : CHALLENGE.chapters.slice(currentStartIndex, currentEndIndex);
  const nextChapters = CHALLENGE.chapters.slice(completedCount, completedCount + CHALLENGE.chaptersPerDay);
  const progressPercent = totalChapters === 0 ? 0 : (completedCount / totalChapters) * 100;
  const remainingCount = totalChapters - completedCount;

  let dayLabel = "시작 전";
  let caption = `${formatShortDate(startDate)} 시작 예정`;
  let completedRange = "아직 시작 전";
  let paceTitle = "시작 전 준비";
  let paceSub = `첫 일정은 ${formatShortDate(startDate)}부터`;

  if (!beforeStart) {
    dayLabel = elapsedDays >= plannedDays ? "마지막 구간" : `Day ${elapsedDays}`;
    completedRange = completedCount > 0 ? `1 ~ ${completedCount}챕터 완료` : "아직 완료 없음";
    paceTitle = remainingCount === 0 ? "완독 완료" : "예정대로 진행 중";
    paceSub = remainingCount === 0
      ? `${formatShortDate(finishDate)} 기준 목표를 모두 읽었습니다.`
      : `오늘 포함 하루 ${CHALLENGE.chaptersPerDay}챕터 페이스`;

    if (remainingCount === 0) {
      caption = `🏆 ${formatShortDate(finishDate)} 완독 목표 달성`;
    } else if (todayChapters.length) {
      caption = `오늘은 ${currentStartIndex + 1} ~ ${currentStartIndex + todayChapters.length}챕터 차례입니다.`;
    }
  }

  return {
    selectedDate,
    startDate,
    finishDate,
    totalChapters,
    plannedDays,
    completedCount,
    remainingCount,
    progressPercent,
    todayChapters,
    nextChapters,
    completedRange,
    dayLabel,
    caption,
    paceTitle,
    paceSub
  };
}

function renderChapterList(targetId, items, emptyMessage, tone) {
  const target = document.getElementById(targetId);
  if (!items.length) {
    target.innerHTML = `
      <div class="chapter-chip">
        <div class="chapter-order">안내</div>
        <div class="chapter-title">${emptyMessage}</div>
      </div>
    `;
    return;
  }

  target.innerHTML = items.map(item => `
    <div class="chapter-chip">
      <div class="chapter-order">${tone} ${item.index}</div>
      <div class="chapter-title">${item.title}</div>
      <div class="chapter-note">${item.note}</div>
    </div>
  `).join("");
}

function renderBookProgress() {
  const state = getChallengeState(getSelectedDate());

  document.getElementById("book-title").textContent = CHALLENGE.title;
  document.getElementById("race-book-title").textContent = CHALLENGE.title;
  document.getElementById("book-meta").textContent =
    `${formatShortDate(state.startDate)} 시작 · 하루 ${CHALLENGE.chaptersPerDay}챕터 · 총 ${state.totalChapters}챕터 · 완독 예정 ${formatShortDate(state.finishDate)}`;
  document.getElementById("chapter-badge").textContent = `${state.completedCount} / ${state.totalChapters}`;
  document.getElementById("hero-day-label").textContent = state.dayLabel;
  document.getElementById("hero-remaining").textContent = `${state.remainingCount}개`;
  document.getElementById("hero-finish-date").textContent = formatShortDate(state.finishDate);

  document.getElementById("race-track-fill").style.width = `${state.progressPercent}%`;
  document.getElementById("race-runner").style.left =
    `clamp(18px, calc(${state.progressPercent}% - 12px), calc(100% - 112px))`;
  document.getElementById("race-progress-badge").textContent = `${Math.round(state.progressPercent)}%`;
  document.getElementById("race-caption").textContent = state.caption;

  const todayItems = state.todayChapters.map((title, index) => ({
    index: state.completedCount - state.todayChapters.length + index + 1,
    title,
    note: `${formatShortDate(state.selectedDate)} 기준 오늘 읽을 순서`
  }));

  const nextItems = state.nextChapters.map((title, index) => ({
    index: state.completedCount + index + 1,
    title,
    note: "다음 읽기 순서"
  }));

  renderChapterList("today-chapter-list", todayItems, "선택 날짜가 시작 전이거나 이미 완독했습니다.", "오늘");
  renderChapterList("next-chapter-list", nextItems, "다음 챕터가 없습니다. 완독 상태입니다.", "다음");

  renderRoadmap(state);
}

function renderCaptureSnapshot() {
  const selectedDateString = getSelectedDate();
  const state = getChallengeState(selectedDateString);
  const todayStartIndex = state.completedCount - state.todayChapters.length;

  document.getElementById("snapshot-title-date").textContent = formatKoreanDate(state.selectedDate);
  document.getElementById("snapshot-date-trigger").setAttribute(
    "aria-label",
    `${formatKoreanDate(state.selectedDate)} 선택됨. 기준 날짜 변경`
  );

  const chapterList = document.getElementById("snapshot-chapter-list");
  if (!state.todayChapters.length) {
    chapterList.innerHTML = `<div class="snapshot-chapter-empty">오늘 읽을 챕터가 없습니다.</div>`;
  } else {
    chapterList.innerHTML = state.todayChapters.map((title, index) => `
      <div class="snapshot-chapter">
        <b>${todayStartIndex + index + 1}</b>
        <span>${title}</span>
      </div>
    `).join("");
  }

  const yesterdayDateString = getPreviousDateString(selectedDateString);
  const yesterdayDate = parseLocalDate(yesterdayDateString);
  const yesterdayRecord = sheetData[yesterdayDateString];
  const yesterdayCount = MEMBERS.filter(member => yesterdayRecord?.[member]).length;
  const yesterdayRate = yesterdayRecord ? Math.round((yesterdayCount / MEMBERS.length) * 100) : 0;
  const yesterdayRingProgress = document.getElementById("yesterday-ring-progress");
  const yesterdayRingGlow = document.getElementById("yesterday-ring-glow");

  document.getElementById("yesterday-date-label").textContent =
    `${formatShortDate(yesterdayDate)} 챌린지 성공률`;
  document.getElementById("yesterday-rate-display").textContent =
    yesterdayRecord ? `${yesterdayRate}%` : "-";
  yesterdayRingProgress.style.strokeDashoffset = yesterdayRecord ? String(100 - yesterdayRate) : "100";
  yesterdayRingGlow.style.strokeDashoffset = yesterdayRecord ? String(100 - yesterdayRate) : "100";
  document.getElementById("yesterday-count-display").textContent =
    yesterdayRecord ? `${yesterdayCount} / ${MEMBERS.length}명 인증` : "전날 데이터 대기 중";

  const allDates = Object.keys(sheetData).sort().filter(date => date <= yesterdayDateString);
  const recent7Dates = allDates.slice(-7);
  const comebackList = document.getElementById("snapshot-comeback-list");
  const comebackSub = document.getElementById("snapshot-comeback-sub");

  if (recent7Dates.length < 2) {
    comebackList.innerHTML = `<div class="snapshot-comeback-empty">부활상 데이터 대기 중</div>`;
    comebackSub.textContent = `${formatShortDate(yesterdayDate)} 기준`;
    return;
  }

  const comebackStats = MEMBERS.map(name => {
    let comebackCount = 0;
    let successCount = 0;
    let lastComebackDate = "";

    recent7Dates.forEach((date, index) => {
      const ok = Boolean(sheetData[date]?.[name]);
      if (ok) successCount += 1;
      if (index > 0) {
        const prevOk = Boolean(sheetData[recent7Dates[index - 1]]?.[name]);
        if (!prevOk && ok) {
          comebackCount += 1;
          lastComebackDate = date;
        }
      }
    });

    return { name, comebackCount, successCount, lastComebackDate };
  }).filter(stat => stat.comebackCount > 0)
    .sort((a, b) =>
      b.comebackCount - a.comebackCount ||
      b.successCount - a.successCount ||
      b.lastComebackDate.localeCompare(a.lastComebackDate) ||
      a.name.localeCompare(b.name)
    );

  if (!comebackStats.length) {
    comebackList.innerHTML = `<div class="snapshot-comeback-empty">최근 7일 부활 멤버가 없습니다</div>`;
  } else {
    comebackList.innerHTML = comebackStats.map(stat => `
      <div class="snapshot-comeback-chip">
        <span>💪</span>
        <strong>${stat.name}</strong>
      </div>
    `).join("");
  }

  comebackSub.textContent =
    `${formatShortDate(parseLocalDate(recent7Dates[0]))} ~ ${formatShortDate(yesterdayDate)} 기준`;
}

function renderRoadmap(state) {
  const roadmap = document.getElementById("chapter-roadmap");
  const activeStart = state.completedCount - state.todayChapters.length;
  const activeEnd = state.completedCount;

  roadmap.innerHTML = CHALLENGE.chapters.map((title, index) => {
    let className = "roadmap-item upcoming";
    if (index < activeStart) {
      className = "roadmap-item done";
    } else if (index >= activeStart && index < activeEnd && state.todayChapters.length) {
      className = "roadmap-item current";
    }

    return `<div class="${className}">${index + 1}. ${title}</div>`;
  }).join("");
}

function renderToday() {
  const date = getSelectedDate();
  const record = sheetData[date] || {};
  const checkedCount = MEMBERS.filter(member => record[member]).length;
  const total = MEMBERS.length;
  const rate = total === 0 ? 0 : Math.round((checkedCount / total) * 100);

  document.getElementById("rate-display").textContent = `${rate}%`;
  document.getElementById("count-display").textContent = `${checkedCount} / ${total}명 인증`;
  document.getElementById("progress-fill").style.width = `${rate}%`;

  const grid = document.getElementById("member-grid");
  grid.innerHTML = "";
  MEMBERS.forEach(name => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = `member-btn${record[name] ? " checked" : ""}`;
    button.textContent = name;
    grid.appendChild(button);
  });
}

function renderStats() {
  const today = todayString();
  const dates = Object.keys(sheetData).sort().filter(date => date < today);

  if (!dates.length) {
    document.getElementById("perfect-list").innerHTML = `<div class="chip empty">데이터가 없습니다</div>`;
    document.getElementById("streak-list").innerHTML = "";
    document.getElementById("recent7-list").innerHTML = "";
    document.getElementById("comeback-list").innerHTML = `<div class="chip empty">데이터가 없습니다</div>`;
    document.getElementById("trend-list").innerHTML = "";
    document.getElementById("stat-list").innerHTML = "";
    document.getElementById("hardest-date").textContent = "-";
    document.getElementById("hardest-rate").textContent = "-";
    document.getElementById("hottest-date").textContent = "-";
    document.getElementById("hottest-rate").textContent = "-";
    return;
  }

  const totalDates = dates.length;
  const recent7 = dates.slice(-7);

  const memberStats = MEMBERS.map(name => {
    let success = 0;
    let currentStreak = 0;
    let maxStreak = 0;
    let recent7Success = 0;
    let hasRecentComeback = false;

    dates.forEach(date => {
      const ok = Boolean(sheetData[date]?.[name]);
      if (ok) {
        success += 1;
        currentStreak += 1;
        maxStreak = Math.max(maxStreak, currentStreak);
      } else {
        currentStreak = 0;
      }
    });

    recent7.forEach((date, index) => {
      const ok = Boolean(sheetData[date]?.[name]);
      if (ok) recent7Success += 1;
      if (index > 0) {
        const prevOk = Boolean(sheetData[recent7[index - 1]]?.[name]);
        if (!prevOk && ok) hasRecentComeback = true;
      }
    });

    return {
      name,
      success,
      rate: (success / totalDates) * 100,
      currentStreak,
      maxStreak,
      recent7Success,
      recent7Rate: recent7.length ? (recent7Success / recent7.length) * 100 : 0,
      hasRecentComeback
    };
  });

  const dateStats = dates.map(date => {
    const count = MEMBERS.filter(member => sheetData[date]?.[member]).length;
    const rate = (count / MEMBERS.length) * 100;
    const dt = parseLocalDate(date);
    return {
      date,
      label: formatShortDate(dt),
      count,
      rate
    };
  });

  const perfect = memberStats.filter(stat => stat.success === totalDates);
  document.getElementById("perfect-list").innerHTML = perfect.length
    ? perfect.map(stat => `<div class="chip">${stat.name}</div>`).join("")
    : `<div class="chip empty">아직 없습니다</div>`;

  const topStreak = [...memberStats].sort((a, b) => b.maxStreak - a.maxStreak).slice(0, 5);
  document.getElementById("streak-list").innerHTML = topStreak.map((stat, index) => `
    <div class="stat-item">
      <div class="stat-item-header">
        <span class="stat-item-name">${index + 1}. ${stat.name}</span>
        <div class="stat-item-right">
          <span class="stat-item-rate">🔥 ${stat.maxStreak}일</span>
          ${stat.currentStreak > 1 ? `<span class="stat-item-sub">현재 ${stat.currentStreak}일</span>` : ""}
        </div>
      </div>
      <div class="bar-bg"><div class="bar-fill" style="width:${(stat.maxStreak / totalDates) * 100}%"></div></div>
    </div>
  `).join("");

  const recentSorted = [...memberStats].sort((a, b) => b.recent7Rate - a.recent7Rate);
  document.getElementById("recent7-list").innerHTML = recentSorted.map((stat, index) => `
    <div class="stat-item">
      <div class="stat-item-header">
        <span class="stat-item-name">${index + 1}. ${stat.name}</span>
        <div class="stat-item-right">
          <span class="stat-item-sub">${stat.recent7Success}/${recent7.length}일</span>
          <span class="stat-item-rate">${stat.recent7Rate.toFixed(1)}%</span>
        </div>
      </div>
      <div class="bar-bg"><div class="bar-fill" style="width:${stat.recent7Rate}%"></div></div>
    </div>
  `).join("");

  const hardest = dateStats.reduce((lowest, current) => current.rate < lowest.rate ? current : lowest);
  const hottest = dateStats.reduce((highest, current) => current.rate > highest.rate ? current : highest);
  document.getElementById("hardest-date").textContent = hardest.label;
  document.getElementById("hardest-rate").textContent = `${hardest.rate.toFixed(1)}% · ${hardest.count}명`;
  document.getElementById("hottest-date").textContent = hottest.label;
  document.getElementById("hottest-rate").textContent = `${hottest.rate.toFixed(1)}% · ${hottest.count}명`;

  const comeback = memberStats.filter(stat => stat.hasRecentComeback);
  document.getElementById("comeback-list").innerHTML = comeback.length
    ? comeback.map(stat => `<div class="chip">💪 ${stat.name}</div>`).join("")
    : `<div class="chip empty">해당 없음</div>`;

  document.getElementById("trend-list").innerHTML = dateStats.map(stat => `
    <div class="trend-item">
      <span class="trend-date">${stat.label}</span>
      <div class="trend-bar-bg"><div class="trend-bar-fill" style="width:${stat.rate}%"></div></div>
      <span class="trend-rate">${Math.round(stat.rate)}%</span>
    </div>
  `).join("");

  const sorted = [...memberStats].sort((a, b) => b.success - a.success || b.rate - a.rate);
  let rank = 0;
  let previousSuccess = null;
  document.getElementById("stat-list").innerHTML = sorted.map((stat, index) => {
    if (stat.success !== previousSuccess) {
      rank = index + 1;
      previousSuccess = stat.success;
    }
    return `
      <div class="stat-item">
        <div class="stat-item-header">
          <span class="stat-item-name">${rank}. ${stat.name}</span>
          <div class="stat-item-right">
            ${stat.currentStreak > 1 ? `<span class="badge-streak">🔥${stat.currentStreak}일</span>` : ""}
            <span class="stat-item-sub">${stat.success}/${totalDates}일</span>
            <span class="stat-item-rate">${stat.rate.toFixed(1)}%</span>
          </div>
        </div>
        <div class="bar-bg"><div class="bar-fill" style="width:${stat.rate}%"></div></div>
      </div>
    `;
  }).join("");
}

function parseHeaderToDate(label) {
  if (!label) return null;
  const text = String(label).trim();

  if (/^\d{4}-\d{1,2}-\d{1,2}$/.test(text)) {
    const [year, month, day] = text.split("-");
    return `${year}-${String(Number(month)).padStart(2, "0")}-${String(Number(day)).padStart(2, "0")}`;
  }

  if (/^\d{1,2}\/\d{1,2}$/.test(text)) {
    const [month, day] = text.split("/");
    const baseYear = parseLocalDate(CHALLENGE.startDate).getFullYear();
    return `${baseYear}-${String(Number(month)).padStart(2, "0")}-${String(Number(day)).padStart(2, "0")}`;
  }

  return null;
}

async function getValidAccessToken() {
  if (tokenCache.accessToken && tokenCache.tokenExpiry > Date.now() + 5 * 60 * 1000) {
    return tokenCache.accessToken;
  }

  const response = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id: CLIENT_ID,
      client_secret: CLIENT_SECRET,
      refresh_token: REFRESH_TOKEN,
      grant_type: "refresh_token"
    })
  });

  const data = await response.json();
  if (!data.access_token) {
    throw new Error("토큰 갱신 실패");
  }

  tokenCache.accessToken = data.access_token;
  tokenCache.tokenExpiry = Date.now() + (data.expires_in || 3600) * 1000;
  return tokenCache.accessToken;
}

async function fetchSheetValues() {
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}/values/${encodeURIComponent(SHEET_NAME)}`;

  let apiKeyMessage = "API Key 조회 실패";
  try {
    const apiKeyResponse = await fetch(`${url}?key=${API_KEY}`);
    if (apiKeyResponse.ok) {
      return apiKeyResponse.json();
    }

    apiKeyMessage = `API Key 조회 실패 (${apiKeyResponse.status})`;
    try {
      const apiKeyError = await apiKeyResponse.json();
      apiKeyMessage = apiKeyError.error?.message || apiKeyMessage;
    } catch (error) {
      console.warn("API Key 오류 응답을 해석하지 못했습니다.", error);
    }
  } catch (error) {
    apiKeyMessage = `API Key 조회 실패: ${error.message}`;
    console.warn("API Key 조회가 브라우저에서 실패했습니다.", error);
  }

  try {
    const token = await getValidAccessToken();
    const authorized = await fetch(url, {
      headers: { Authorization: `Bearer ${token}` }
    });
    if (authorized.ok) {
      return authorized.json();
    }
    throw new Error(`토큰 조회 실패 (${authorized.status})`);
  } catch (error) {
    console.warn("토큰 기반 조회도 실패했습니다.", error);
  }

  throw new Error(apiKeyMessage);
}

async function loadFromSheets() {
  setLoading(true);
  try {
    const data = await fetchSheetValues();
    if (!data.values) {
      throw new Error("시트 데이터를 찾지 못했습니다.");
    }

    const rows = data.values;
    const headerRow = rows[0] || [];
    const dateColumns = {};

    for (let index = 1; index < headerRow.length; index += 1) {
      const label = headerRow[index];
      if (!label || SPECIAL_ROWS.includes(label)) continue;
      const dateValue = parseHeaderToDate(label);
      if (dateValue) {
        dateColumns[index] = dateValue;
      }
    }

    const nextData = {};
    for (let rowIndex = 1; rowIndex < rows.length; rowIndex += 1) {
      const row = rows[rowIndex];
      const name = row[0];
      if (!name || SPECIAL_ROWS.includes(name) || !MEMBERS.includes(name)) continue;

      Object.entries(dateColumns).forEach(([colIndex, dateValue]) => {
        if (!nextData[dateValue]) nextData[dateValue] = {};
        nextData[dateValue][name] = row[Number(colIndex)] === "O";
      });
    }

    sheetData = nextData;
    renderAll();

    const now = new Date();
    document.getElementById("last-updated").textContent =
      `마지막 업데이트 ${formatShortDate(now)} ${String(now.getHours()).padStart(2, "0")}:${String(now.getMinutes()).padStart(2, "0")}`;
    showToast("새 데이터를 불러왔습니다.");
  } catch (error) {
    console.error(error);
    document.getElementById("last-updated").textContent =
      `데이터 불러오기 실패: ${error.message}`;
    showToast(`불러오기 실패: ${error.message}`, "error");
  } finally {
    setLoading(false);
  }
}

function renderAll() {
  renderHeaderChrome();
  renderBookProgress();
  renderCaptureSnapshot();
  renderToday();
  if (activeTab === "stats") {
    renderStats();
  }
}

function showToast(message, type = "") {
  const toast = document.getElementById("toast");
  toast.textContent = message;
  toast.className = `toast${type ? ` ${type}` : ""}`;
  toast.classList.add("show");
  window.clearTimeout(showToast._timer);
  showToast._timer = window.setTimeout(() => toast.classList.remove("show"), 2800);
}

function setLoading(enabled) {
  document.getElementById("loading").style.display = enabled ? "flex" : "none";
}

document.getElementById("refresh-btn").addEventListener("click", loadFromSheets);
document.getElementById("refresh-btn-top").addEventListener("click", loadFromSheets);
document.getElementById("jump-today-btn").addEventListener("click", () => {
  switchTab("today");
  document.getElementById("today-chapter-card").scrollIntoView({ behavior: "smooth", block: "center" });
});

initTabs();
initDatePicker();
renderBookProgress();
renderHeaderChrome();
renderCaptureSnapshot();
loadFromSheets();
