const state = {
  meta: null,
  filters: {
    search: "",
    recipient: "",
    status: "all",
    page: 1,
    pageSize: 24,
  },
  result: null,
  selectedId: null,
};

const refs = {
  excelPath: document.querySelector("#excelPath"),
  todayValue: document.querySelector("#todayValue"),
  loadedAt: document.querySelector("#loadedAt"),
  statsGrid: document.querySelector("#statsGrid"),
  searchInput: document.querySelector("#searchInput"),
  recipientSelect: document.querySelector("#recipientSelect"),
  statusSelect: document.querySelector("#statusSelect"),
  recipientChips: document.querySelector("#recipientChips"),
  reminderList: document.querySelector("#reminderList"),
  resultSummary: document.querySelector("#resultSummary"),
  prevPageButton: document.querySelector("#prevPageButton"),
  nextPageButton: document.querySelector("#nextPageButton"),
  pageText: document.querySelector("#pageText"),
  detailEmpty: document.querySelector("#detailEmpty"),
  detailContent: document.querySelector("#detailContent"),
  detailStatus: document.querySelector("#detailStatus"),
  detailCustomerName: document.querySelector("#detailCustomerName"),
  detailProjectName: document.querySelector("#detailProjectName"),
  detailUrgency: document.querySelector("#detailUrgency"),
  detailOverview: document.querySelector("#detailOverview"),
  detailFields: document.querySelector("#detailFields"),
  reloadButton: document.querySelector("#reloadButton"),
  notifyButton: document.querySelector("#notifyButton"),
};

function formatStatus(status) {
  const mapping = {
    overdue: "已逾期",
    due: "应提醒",
    upcoming: "待进入提醒期",
  };
  return mapping[status] || status;
}

function formatDate(value) {
  if (!value) return "-";
  return value;
}

function formatUrgency(days) {
  if (days < 0) {
    return `已过续期日 ${Math.abs(days)} 天`;
  }
  if (days === 0) {
    return "今天到续期日";
  }
  return `距续期 ${days} 天`;
}

async function fetchJSON(url, options) {
  const response = await fetch(url, options);
  if (!response.ok) {
    throw new Error(`请求失败: ${response.status}`);
  }
  return response.json();
}

function applyRecipientFromQuery() {
  const params = new URLSearchParams(window.location.search);
  const recipient = params.get("recipient");
  if (recipient) {
    state.filters.recipient = recipient;
  }
}

async function loadMeta() {
  state.meta = await fetchJSON("/api/meta");
  refs.excelPath.textContent = state.meta.excel_path;
  refs.todayValue.textContent = state.meta.today;
  refs.loadedAt.textContent = state.meta.last_loaded_at || "-";

  refs.recipientSelect.innerHTML = '<option value="">全部</option>';
  state.meta.recipients.forEach((recipient) => {
    const option = document.createElement("option");
    option.value = recipient;
    option.textContent = recipient;
    if (recipient === state.filters.recipient) {
      option.selected = true;
    }
    refs.recipientSelect.appendChild(option);
  });
}

function renderStats(stats) {
  const cards = [
    { label: "客户总数", value: stats.total, tone: "neutral" },
    { label: "已逾期", value: stats.overdue, tone: "danger" },
    { label: "应提醒", value: stats.due, tone: "warning" },
    { label: "待进入提醒期", value: stats.upcoming, tone: "calm" },
  ];

  refs.statsGrid.innerHTML = "";
  cards.forEach((card) => {
    const element = document.createElement("article");
    element.className = `panel stat-card ${card.tone}`;
    element.innerHTML = `
      <p>${card.label}</p>
      <strong>${card.value}</strong>
    `;
    refs.statsGrid.appendChild(element);
  });
}

function renderRecipientChips(topRecipients) {
  refs.recipientChips.innerHTML = "";
  if (!topRecipients.length) {
    refs.recipientChips.innerHTML = '<span class="muted">当前筛选下暂无重点提醒人。</span>';
    return;
  }

  topRecipients.forEach((item) => {
    const button = document.createElement("button");
    button.className = "recipient-chip";
    button.innerHTML = `<span>${item.name}</span><strong>${item.count}</strong>`;
    button.addEventListener("click", () => {
      state.filters.recipient = item.name;
      state.filters.page = 1;
      refs.recipientSelect.value = item.name;
      loadReminders();
    });
    refs.recipientChips.appendChild(button);
  });
}

function renderList(result) {
  refs.reminderList.innerHTML = "";
  refs.resultSummary.textContent = `共 ${result.stats.total} 条结果`;
  refs.pageText.textContent = `${result.filters.page} / ${result.page_count}`;
  refs.prevPageButton.disabled = result.filters.page <= 1;
  refs.nextPageButton.disabled = result.filters.page >= result.page_count;

  if (!result.items.length) {
    refs.reminderList.innerHTML = '<div class="empty-list">当前筛选条件下没有客户记录。</div>';
    return;
  }

  result.items.forEach((item) => {
    const button = document.createElement("button");
    button.className = `reminder-card ${item.status}`;
    button.dataset.id = item.id;
    if (item.id === state.selectedId) {
      button.classList.add("active");
    }
    button.innerHTML = `
      <div class="reminder-head">
        <span class="status-tag ${item.status}">${formatStatus(item.status)}</span>
        <span class="owner-tag">${item.owner_name}</span>
      </div>
      <h3>${item.customer_name}</h3>
      <p class="project-line">${item.project_name || "未填写项目名称"}</p>
      <dl class="mini-grid">
        <div><dt>报告落款日</dt><dd>${formatDate(item.report_date)}</dd></div>
        <div><dt>续期日</dt><dd>${formatDate(item.renewal_date)}</dd></div>
        <div><dt>开始提醒</dt><dd>${formatDate(item.remind_date)}</dd></div>
        <div><dt>剩余时间</dt><dd>${formatUrgency(item.days_until_renewal)}</dd></div>
      </dl>
    `;
    button.addEventListener("click", () => loadDetail(item.id));
    refs.reminderList.appendChild(button);
  });
}

function renderDetail(record) {
  refs.detailEmpty.classList.add("hidden");
  refs.detailContent.classList.remove("hidden");
  refs.detailStatus.textContent = `${formatStatus(record.status)} · 被提醒人 ${record.owner_name}`;
  refs.detailCustomerName.textContent = record.customer_name;
  refs.detailProjectName.textContent = record.project_name || "未填写项目名称";
  refs.detailUrgency.textContent = formatUrgency(record.days_until_renewal);
  refs.detailUrgency.className = `urgency-badge ${record.status}`;

  const overview = [
    ["项目编号", record.project_code || "-"],
    ["部门", record.department || "-"],
    ["负责人", record.responsible_person || "-"],
    ["主笔", record.primary_writer || "-"],
    ["签约人", record.signer || "-"],
    ["报告落款日", record.report_date],
    ["续期日", record.renewal_date],
    ["提醒开始日", record.remind_date],
  ];

  refs.detailOverview.innerHTML = overview
    .map(([label, value]) => `<div class="overview-item"><span>${label}</span><strong>${value}</strong></div>`)
    .join("");

  refs.detailFields.innerHTML = record.fields
    .map(
      (field) => `
        <div class="detail-field">
          <span>${field.label}</span>
          <strong>${field.value || "-"}</strong>
        </div>
      `
    )
    .join("");
}

async function loadDetail(id) {
  state.selectedId = id;
  const record = await fetchJSON(`/api/customers/${id}`);
  renderDetail(record);
  document.querySelectorAll(".reminder-card").forEach((card) => {
    card.classList.toggle("active", card.dataset.id === id);
  });
}

async function loadReminders() {
  const params = new URLSearchParams({
    search: state.filters.search,
    recipient: state.filters.recipient,
    status: state.filters.status,
    page: String(state.filters.page),
    page_size: String(state.filters.pageSize),
  });

  state.result = await fetchJSON(`/api/reminders?${params.toString()}`);
  renderStats(state.result.stats);
  renderRecipientChips(state.result.top_recipients);
  renderList(state.result);

  if (!state.selectedId && state.result.items.length) {
    loadDetail(state.result.items[0].id);
    return;
  }

  if (state.selectedId) {
    const exists = state.result.items.some((item) => item.id === state.selectedId);
    if (!exists && state.result.items.length) {
      loadDetail(state.result.items[0].id);
    }
  }
}

async function reloadWorkbook() {
  refs.reloadButton.disabled = true;
  refs.reloadButton.textContent = "加载中...";
  try {
    await fetchJSON("/api/reload", { method: "POST" });
    await loadMeta();
    await loadReminders();
  } finally {
    refs.reloadButton.disabled = false;
    refs.reloadButton.textContent = "重新加载 Excel";
  }
}

function scheduleNotifications() {
  if (!("Notification" in window)) {
    refs.notifyButton.disabled = true;
    refs.notifyButton.textContent = "浏览器不支持桌面提醒";
    return;
  }

  refs.notifyButton.addEventListener("click", async () => {
    const permission = await Notification.requestPermission();
    if (permission !== "granted") {
      refs.notifyButton.textContent = "桌面提醒未开启";
      return;
    }

    refs.notifyButton.textContent = "桌面提醒已开启";
    pushNotifications();
  });
}

function pushNotifications() {
  if (Notification.permission !== "granted" || !state.result) {
    return;
  }

  const notified = JSON.parse(sessionStorage.getItem("crm-notified-ids") || "[]");
  const cache = new Set(notified);
  const targets = state.result.items.filter((item) => item.status !== "upcoming").slice(0, 3);

  targets.forEach((item) => {
    if (cache.has(item.id)) return;

    const notification = new Notification(`续期提醒：${item.customer_name}`, {
      body: `${item.owner_name} 需要跟进，续期日 ${item.renewal_date}`,
    });

    notification.onclick = () => {
      window.focus();
      loadDetail(item.id);
    };

    cache.add(item.id);
  });

  sessionStorage.setItem("crm-notified-ids", JSON.stringify([...cache]));
}

function bindEvents() {
  let timer = null;
  refs.searchInput.addEventListener("input", (event) => {
    clearTimeout(timer);
    timer = setTimeout(() => {
      state.filters.search = event.target.value.trim();
      state.filters.page = 1;
      loadReminders();
    }, 250);
  });

  refs.recipientSelect.addEventListener("change", (event) => {
    state.filters.recipient = event.target.value;
    state.filters.page = 1;
    loadReminders();
  });

  refs.statusSelect.addEventListener("change", (event) => {
    state.filters.status = event.target.value;
    state.filters.page = 1;
    loadReminders();
  });

  refs.prevPageButton.addEventListener("click", () => {
    if (state.filters.page <= 1) return;
    state.filters.page -= 1;
    loadReminders();
  });

  refs.nextPageButton.addEventListener("click", () => {
    if (!state.result || state.filters.page >= state.result.page_count) return;
    state.filters.page += 1;
    loadReminders();
  });

  refs.reloadButton.addEventListener("click", reloadWorkbook);
}

async function bootstrap() {
  applyRecipientFromQuery();
  bindEvents();
  scheduleNotifications();
  await loadMeta();
  refs.searchInput.value = state.filters.search;
  refs.statusSelect.value = state.filters.status;
  refs.recipientSelect.value = state.filters.recipient;
  await loadReminders();
  pushNotifications();
}

bootstrap().catch((error) => {
  refs.reminderList.innerHTML = `<div class="empty-list">程序启动失败：${error.message}</div>`;
});
