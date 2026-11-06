const sampleRecords = [
  { id: 1, borrowDate: "2026-09-15", date: "15 ก.ย. 69", time: "09:42", job: "WO-2569-0148", requesterName: "วิทยา", partNames: ["Limit Switch Omron D4N", "Limit Switch Omron D4N"], part: "Limit Switch Omron D4N", code: "EL-0321", destination: "เครื่องลูกค้า", sourceSerialNumber: "CUS-MC-1042", qty: 2, unit: "ชิ้น", note: "", status: "borrowed", dueDate: "2026-09-18", due: "18 ก.ย. 69" },
  { id: 2, borrowDate: "2026-09-14", date: "14 ก.ย. 69", time: "15:18", job: "WO-2569-0142", requesterName: "ธนา", part: "สายพาน Timing Belt 5M", code: "MC-0189", destination: "เครื่องใหม่", qty: 1, unit: "ชิ้น", note: "", status: "borrowed", dueDate: "2026-09-20", due: "20 ก.ย. 69" },
  { id: 3, borrowDate: "2026-09-13", date: "13 ก.ย. 69", time: "11:05", job: "SV-2569-0087", requesterName: "พงศกร", part: "Proximity Sensor M18", code: "EL-0177", destination: "เครื่อง Demo", sourceSerialNumber: "DEMO-07", qty: 1, unit: "ชิ้น", note: "", status: "returned", dueDate: "2026-09-16", actualReturnDate: "2026-09-14", returned: "14 ก.ย. 69" },
  { id: 4, borrowDate: "2026-09-12", date: "12 ก.ย. 69", time: "16:30", job: "WO-2569-0136", requesterName: "วิทยา", part: "Pneumatic Cylinder Ø32", code: "PN-0245", destination: "เครื่องลูกค้า", qty: 1, unit: "ชิ้น", note: "", status: "borrowed", dueDate: "2026-09-22", due: "22 ก.ย. 69" },
  { id: 5, borrowDate: "2026-09-10", date: "10 ก.ย. 69", time: "10:12", job: "DEMO-0041", requesterName: "ธนา", part: "Switching Power Supply 24V", code: "EL-0084", destination: "เครื่อง Demo", qty: 1, unit: "ชิ้น", note: "", status: "borrowed", dueDate: "2026-09-30", due: "30 ก.ย. 69" },
  { id: 6, borrowDate: "2026-09-08", date: "8 ก.ย. 69", time: "13:47", job: "WO-2569-0121", requesterName: "พงศกร", part: "Linear Bearing LM20UU", code: "MC-0096", destination: "เครื่องใหม่", qty: 4, unit: "ชิ้น", note: "", status: "returned", dueDate: "2026-09-10", actualReturnDate: "2026-09-11", returned: "11 ก.ย. 69" }
];

let records = [];
const API_URL = String(window.PARTFLOW_CONFIG?.apiUrl || "").trim().replace(/\/$/, "");
const state = {
  filter: "all",
  search: "",
  returnId: null,
  pendingCreateRequestId: null,
  pendingReturnRequestId: null,
  loading: false,
  connection: "connecting",
  formDirty: false,
  saving: false,
  returnSaving: false
};
const labels = { borrowed: "กำลังยืม", returned: "คืนแล้ว" };
const requesterNames = [
  "จีรศักดิ์ คำศรี",
  "ชาติชาย น่าบัณฑิต",
  "ภาสพงษ์ คำศรี",
  "พงษ์ศักดิ์ จัการพิมพ์",
  "สมชาย มูลเงิน",
  "ธนบูรณ์ ศุภชาติสันติ",
  "ณัฐพงษ์ กลิ่นจันทร์",
  "กันต์ศักดิ์ ฉลาดเฉลียว",
  "ประวิทย์ ฉลาดเฉลียว",
  "กษิเดช อิทธิธนาบุญ",
  "เอกเลิศ ตันติยมาศ",
  "อภิลักษณ์ หาไธสง",
  "ธนาธิป ขยัน"
].sort(new Intl.Collator("th").compare);
const recordsBody = document.querySelector("#recordsBody");
const emptyState = document.querySelector("#emptyState");
const historyRecordsBody = document.querySelector("#historyRecordsBody");
const historyPager = createListPager(document.querySelector("#historyRecordsBody").closest(".table-wrap"), renderRecords);
const historyEmptyState = document.querySelector("#historyEmptyState");
const dashboardPage = document.querySelector("#dashboardPage");
const historyPage = document.querySelector("#historyPage");
const transactionPage = document.querySelector("#transactionPage");
const returnModal = document.querySelector("#returnModal");
const form = document.querySelector("#borrowForm");
const partNameList = document.querySelector("#partNameList");
const borrowDateInput = document.querySelector("#borrowDate");
const returnDateInput = document.querySelector("#returnDate");
const actualReturnDateInput = document.querySelector("#actualReturnDate");
const dateFormatter = new Intl.DateTimeFormat("th-TH", { day: "numeric", month: "short", year: "2-digit" });
const timeFormatter = new Intl.DateTimeFormat("th-TH", { hour: "2-digit", minute: "2-digit", hour12: false, timeZone: "Asia/Bangkok" });

function createRequestId() {
  if (window.crypto?.randomUUID) return window.crypto.randomUUID();
  return `pf-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function isApiConfigured() {
  return /^https:\/\/script\.google\.com\/macros\/s\/.+\/exec(?:\?.*)?$/i.test(API_URL);
}

function setConnectionStatus(status, label, detail) {
  state.connection = status;
  const dot = document.querySelector("#connectionDot");
  dot.classList.remove("connecting", "error", "demo");
  if (status !== "ready") dot.classList.add(status);
  document.querySelector("#connectionLabel").textContent = label;
  document.querySelector("#connectionDetail").textContent = detail;
  const notice = document.querySelector("#storageNotice");
  if (notice) {
    notice.textContent = status === "ready"
      ? "ข้อมูลจะบันทึกลง Google Sheet ทันทีเมื่อกด “บันทึกรายการ”"
      : status === "demo"
        ? "โหมดตัวอย่างเก็บข้อมูลเฉพาะระหว่างเปิดหน้าเว็บ ข้อมูลจะหายเมื่อรีเฟรชหน้า"
        : "โปรดรอให้ระบบเชื่อมต่อ Google Sheet ก่อนบันทึกรายการ";
  }
  document.querySelectorAll('[data-action="refresh-data"]').forEach(button => {
    button.disabled = status === "connecting" || status === "demo";
    button.textContent = status === "connecting" ? "กำลังโหลด..." : status === "demo" ? "โหมดตัวอย่าง" : "รีเฟรชข้อมูล";
  });
}

function normalizeRecord(record) {
  const createdAt = record.createdAt ? new Date(record.createdAt) : null;
  const hasCreatedAt = createdAt && !Number.isNaN(createdAt.getTime());
  const partNames = Array.isArray(record.partNames) ? record.partNames.map(String) : getPartNames(record).filter(Boolean);
  return {
    ...record,
    id: String(record.id),
    job: record.jobNumber || record.job || "",
    jobNumber: record.jobNumber || record.job || "",
    customerName: String(record.customerName || ""),
    partNames,
    part: partNames[0] || "-",
    qty: Number(record.quantity ?? record.qty) || partNames.length,
    quantity: Number(record.quantity ?? record.qty) || partNames.length,
    borrowDate: record.borrowDate || "",
    dueDate: record.dueDate || "",
    actualReturnDate: record.actualReturnDate || "",
    date: formatThaiDate(record.borrowDate),
    due: formatThaiDate(record.dueDate),
    returned: record.actualReturnDate ? formatThaiDate(record.actualReturnDate) : "",
    time: hasCreatedAt ? timeFormatter.format(createdAt).replace(".", ":") : (record.time || "--:--"),
    status: record.status === "returned" ? "returned" : "borrowed",
    code: record.code || ""
  };
}

function populateRequesterOptions() {
  const select = document.querySelector("#requesterName");
  const options = requesterNames.map(name => {
    const option = document.createElement("option");
    option.value = name;
    option.textContent = name;
    return option;
  });
  const other = document.createElement("option");
  other.value = "__other__";
  other.textContent = "อื่นๆ (ระบุชื่อผู้เบิก)";
  select.append(...options, other);
}

async function apiRequest(action, payload = {}) {
  const options = action === "list"
    ? undefined
    : {
        method: "POST",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify({ action, ...payload })
      };
  const url = action === "list" ? `${API_URL}?action=list&_ts=${Date.now()}` : API_URL;
  const response = await fetch(url, options);
  const text = await response.text();
  let result;
  try {
    result = JSON.parse(text);
  } catch {
    throw new Error("Apps Script ตอบกลับด้วยข้อมูลที่อ่านไม่ได้ กรุณาตรวจสอบ URL และสิทธิ์การเข้าถึง");
  }
  if (!response.ok || !result.ok) {
    throw new Error(result?.error?.message || `เชื่อมต่อไม่สำเร็จ (${response.status})`);
  }
  return result.data;
}

async function loadRecords() {
  if (!isApiConfigured()) {
    records = sampleRecords.map(normalizeRecord);
    setConnectionStatus("demo", "โหมดตัวอย่าง", "ใส่ Apps Script URL ใน config.js");
    updateMetrics();
    renderRecords();
    return;
  }

  if (state.loading) return;
  state.loading = true;
  setConnectionStatus("connecting", "กำลังเชื่อมต่อ", "กำลังอ่านข้อมูลจาก Google Sheet");
  updateMetrics();
  renderRecords();
  try {
    const data = await apiRequest("list");
    records = (data.records || []).map(normalizeRecord);
    setConnectionStatus("ready", "เชื่อมต่อแล้ว", "Google Sheet พร้อมใช้งาน");
  } catch (error) {
    setConnectionStatus("error", "เชื่อมต่อไม่สำเร็จ", "ตรวจสอบ Apps Script URL และสิทธิ์");
    showToast("โหลดข้อมูลไม่สำเร็จ", error.message);
  } finally {
    state.loading = false;
    updateMetrics();
    renderRecords();
  }
}

function escapeHtml(value) {
  return String(value ?? "").replace(/[&<>'"]/g, character => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&#39;", '"': "&quot;" })[character]);
}

function getPartNames(record) {
  if (record.partNames?.length) return record.partNames;
  if (record.items?.length) return record.items.map(item => item.name);
  return [record.part];
}

function toLocalIsoDate(date = new Date()) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function formatThaiDate(isoDate) {
  return isoDate ? dateFormatter.format(new Date(`${isoDate}T00:00:00`)) : "-";
}

function applyTheme(theme) {
  const normalizedTheme = theme === "dark" ? "dark" : "light";
  document.documentElement.dataset.theme = normalizedTheme;
  document.documentElement.style.colorScheme = normalizedTheme;
  const isDark = normalizedTheme === "dark";
  document.querySelector('meta[name="theme-color"]').setAttribute("content", isDark ? "#0c1625" : "#eef4fb");
}

function renderPartSummary(partNames) {
  if (partNames.length === 1) return `<strong>${escapeHtml(partNames[0])}</strong>`;
  const items = partNames.map((name, index) => `<li><span>${index + 1}</span>${escapeHtml(name)}</li>`).join("");
  return `<details class="part-details"><summary><strong>${escapeHtml(partNames[0])}</strong><span class="multi-part-note">ดูทั้งหมด ${partNames.length} ชิ้น</span></summary><ol>${items}</ol></details>`;
}

function getDueText(record) {
  if (record.status === "returned") return record.returned ? `คืนจริง ${record.returned}` : "คืนแล้ว";
  if (!record.dueDate) return `คืน ${record.due || "-"}`;
  const today = new Date(`${toLocalIsoDate()}T00:00:00`);
  const dueDate = new Date(`${record.dueDate}T00:00:00`);
  const days = Math.round((dueDate - today) / 86400000);
  if (days < 0) return `เกินกำหนด ${Math.abs(days)} วัน`;
  if (days === 0) return "ครบกำหนดวันนี้";
  return `คืน ${record.due || formatThaiDate(record.dueDate)}`;
}

function createRecordRow(record) {
  const partNames = getPartNames(record);
  const itemMeta = [record.code, record.sourceSerialNumber ? `ต้นทาง S/N ${record.sourceSerialNumber}` : ""].filter(Boolean).join(" · ");
  return `
      <tr>
        <td class="date-cell"><strong>${escapeHtml(record.date)}</strong><span>${escapeHtml(record.time)} น.</span></td>
        <td class="job-cell"><strong>${escapeHtml(record.job)}</strong><span>ผู้เบิก ${escapeHtml(record.requesterName || "-")}</span>${record.customerName ? `<span>ลูกค้า ${escapeHtml(record.customerName)}</span>` : ""}</td>
        <td class="part-cell">${renderPartSummary(partNames)}<span>${escapeHtml(itemMeta)}</span></td>
        <td><span class="destination"><i class="destination-dot"></i>${escapeHtml(record.destination)}</span></td>
        <td>${escapeHtml(record.qty)} ${escapeHtml(record.unit)}</td>
        <td><span class="status status-${record.status}">${labels[record.status]}</span><span class="part-cell"><span>${escapeHtml(getDueText(record))}</span></span></td>
        <td><div class="row-actions"><button class="row-action row-action-copy" type="button" data-copy-id="${escapeHtml(record.id)}" aria-label="คัดลอกข้อความติดตามเลขงาน ${escapeHtml(record.job)}">คัดลอก</button>${record.status === "borrowed" ? `<button class="row-action" type="button" data-return-id="${escapeHtml(record.id)}">รับคืน</button>` : ""}</div></td>
      </tr>`;
}

function getFilteredRecords() {
  const term = state.search.trim().toLowerCase();
  const dueLimit = new Date();
  dueLimit.setDate(dueLimit.getDate() + 3);
  const dueLimitIso = toLocalIsoDate(dueLimit);
  return records.filter(record => {
    const searchable = [record.job, record.requesterName || "", record.customerName || "", record.destination, record.sourceSerialNumber || record.serial || "", record.code, ...getPartNames(record)];
    const matchesFilter = state.filter === "all" || record.status === state.filter ||
      (state.filter === "due" && record.status === "borrowed" && record.dueDate && record.dueDate <= dueLimitIso);
    return matchesFilter && (!term || searchable.some(value => String(value || "").toLowerCase().includes(term)));
  });
}

function renderRecords() {
  const recentRecords = records.slice(0, 4);
  recordsBody.innerHTML = recentRecords.map(createRecordRow).join("");
  emptyState.hidden = recentRecords.length > 0;
  emptyState.querySelector("strong").textContent = state.connection === "connecting" ? "กำลังโหลดรายการ" : state.connection === "error" && !records.length ? "โหลดรายการไม่สำเร็จ" : "ยังไม่มีรายการ";
  emptyState.querySelector("span").textContent = state.connection === "connecting" ? "โปรดรอสักครู่" : state.connection === "error" && !records.length ? "ตรวจสอบการเชื่อมต่อ แล้วกดรีเฟรชข้อมูล" : "รายการใหม่จะแสดงในส่วนนี้";
  document.querySelector("#recordCount").textContent = recentRecords.length
    ? `แสดง ${recentRecords.length} รายการล่าสุด จากทั้งหมด ${records.length} รายการ`
    : "ยังไม่มีรายการ";

  const visible = getFilteredRecords();
  const pageRecords = historyPager.page(visible);

  historyRecordsBody.innerHTML = pageRecords.map(createRecordRow).join("");
  historyEmptyState.hidden = visible.length > 0;
  historyEmptyState.querySelector("strong").textContent = state.connection === "connecting" ? "กำลังโหลดรายการ" : state.connection === "error" && !records.length ? "โหลดรายการไม่สำเร็จ" : "ไม่พบรายการที่ค้นหา";
  historyEmptyState.querySelector("span").textContent = state.connection === "connecting" ? "โปรดรอสักครู่" : state.connection === "error" && !records.length ? "ตรวจสอบการเชื่อมต่อ แล้วกดรีเฟรชข้อมูล" : "ลองเปลี่ยนคำค้นหาหรือตัวกรอง";
  document.querySelector("#historyRecordCount").textContent = `พบ ${visible.length} จาก ${records.length} รายการ`;
  document.querySelectorAll(".filter-tab").forEach(tab => {
    const active = tab.dataset.filter === state.filter;
    tab.classList.toggle("active", active);
    tab.setAttribute("aria-selected", active);
  });
}

function syncPartNameInputs() {
  const quantityInput = document.querySelector("#quantity");
  const quantity = Math.min(50, Math.max(1, Math.trunc(Number(quantityInput.value) || 1)));
  quantityInput.value = quantity;
  const existingValues = Array.from(partNameList.querySelectorAll("input")).map(input => input.value);
  partNameList.innerHTML = "";
  for (let index = 0; index < quantity; index += 1) {
    const label = document.createElement("label");
    label.className = "part-name-entry";
    label.innerHTML = `<span>ชื่ออะไหล่ชิ้นที่ ${index + 1} <em>*</em></span><div class="input-icon"><svg aria-hidden="true" viewBox="0 0 24 24"><circle cx="11" cy="11" r="6"/><path d="m16 16 4 4"/></svg><input type="text" placeholder="ค้นหาหรือพิมพ์ชื่ออะไหล่ชิ้นที่ ${index + 1}" autocomplete="off" maxlength="250" required></div>`;
    label.querySelector("input").value = existingValues[index] || "";
    partNameList.appendChild(label);
  }
}

function collectPartNames() {
  return Array.from(partNameList.querySelectorAll("input")).map(input => input.value.trim());
}

function setFilter(filter) {
  state.filter = filter;
  historyPager.reset();
  showHistory();
  renderRecords();
}

function setActiveNavigation(view) {
  document.querySelector("#navDashboard").classList.toggle("active", view === "dashboard");
  document.querySelector("#navNewTransaction").classList.toggle("active", view === "transaction");
  document.querySelector("#navBorrowed").classList.toggle("active", view === "history" && state.filter === "borrowed");
  document.querySelector("#navHistory").classList.toggle("active", view === "history" && state.filter !== "borrowed");
}

function closeMobileNavigation() {
  document.querySelector(".sidebar").classList.remove("open");
  document.querySelectorAll(".mobile-menu").forEach(button => button.setAttribute("aria-expanded", "false"));
}

function showDashboard() {
  closeMobileNavigation();
  transactionPage.hidden = true;
  historyPage.hidden = true;
  dashboardPage.hidden = false;
  setActiveNavigation("dashboard");
  window.scrollTo({ top: 0, behavior: "smooth" });
}

function showHistory() {
  closeMobileNavigation();
  transactionPage.hidden = true;
  dashboardPage.hidden = true;
  historyPage.hidden = false;
  setActiveNavigation("history");
  renderRecords();
  window.scrollTo({ top: 0, behavior: "smooth" });
}

function resetBorrowForm() {
  form.reset();
  const today = toLocalIsoDate();
  borrowDateInput.value = today;
  returnDateInput.value = "";
  returnDateInput.min = today;
  returnDateInput.setCustomValidity("");
  document.querySelector("#otherDestinationField").hidden = true;
  form.elements.otherDestination.required = false;
  document.querySelector("#otherRequesterField").hidden = true;
  form.elements.otherRequesterName.required = false;
  state.pendingCreateRequestId = null;
  state.formDirty = false;
  syncPartNameInputs();
}

function canLeaveForm() {
  if (state.saving) return false;
  return !state.formDirty || window.confirm("ข้อมูลที่กรอกยังไม่ได้บันทึก ต้องการออกจากหน้านี้หรือไม่?");
}

function validateBorrowDates() {
  returnDateInput.min = borrowDateInput.value || "";
  const invalidRange = Boolean(borrowDateInput.value && returnDateInput.value && returnDateInput.value < borrowDateInput.value);
  returnDateInput.setCustomValidity(invalidRange ? "วันที่กำหนดคืนต้องเป็นวันเดียวกับหรือหลังวันที่ยืม" : "");
  return !invalidRange;
}

function openForm() {
  if (state.saving || !transactionPage.hidden) return;
  closeMobileNavigation();
  resetBorrowForm();
  dashboardPage.hidden = true;
  historyPage.hidden = true;
  transactionPage.hidden = false;
  setActiveNavigation("transaction");
  window.scrollTo({ top: 0, behavior: "smooth" });
  setTimeout(() => form.querySelector("input[name=jobNumber]").focus(), 300);
}

function closeForm() {
  if (!canLeaveForm()) return;
  resetBorrowForm();
  showDashboard();
}

function showToast(title, message) {
  const toast = document.querySelector("#toast");
  toast.querySelector("strong").textContent = title;
  toast.querySelector("p").textContent = message;
  toast.classList.add("show");
  window.clearTimeout(showToast.timer);
  showToast.timer = window.setTimeout(() => toast.classList.remove("show"), 3200);
}

function createFollowUpText(record) {
  const partLines = getPartNames(record).map((name, index) => `${index + 1}. ${name}`).join("\n");
  return [
    "ติดตามการยืมอะไหล่",
    `เลขงาน: ${record.job}`,
    `ชื่อผู้เบิก: ${record.requesterName || "-"}`,
    record.customerName ? `ชื่อลูกค้า: ${record.customerName}` : "",
    "รายการอะไหล่:",
    partLines,
    `จำนวนรวม: ${record.qty} ${record.unit}`,
    `นำมาจาก: ${record.destination}`,
    record.sourceSerialNumber ? `หมายเลขเครื่อง (Serial Number): ${record.sourceSerialNumber}` : "",
    `วันที่ยืม: ${record.date}`,
    `วันที่กำหนดคืน: ${record.due || formatThaiDate(record.dueDate)}`,
    `สถานะ: ${labels[record.status]}`,
    record.returned ? `วันที่คืนจริง: ${record.returned}` : "",
    record.note ? `หมายเหตุ: ${record.note}` : "",
    record.status === "borrowed" ? "กรุณาตรวจสอบและแจ้งกำหนดการคืนอะไหล่ ขอบคุณครับ" : ""
  ].filter(Boolean).join("\n");
}

function escapeCsv(value) {
  const text = String(value ?? "");
  const safeText = /^[\s\u0000-\u001f]*[=+\-@]/.test(text) ? `'${text}` : text;
  return `"${safeText.replace(/"/g, '""')}"`;
}

function exportData() {
  const exportRecords = getFilteredRecords();
  if (!exportRecords.length) {
    showToast("ไม่มีข้อมูลให้ส่งออก", "ลองเปลี่ยนคำค้นหาหรือตัวกรองแล้วดำเนินการอีกครั้ง");
    return;
  }

  const headers = [
    "รหัสรายการ", "วันที่ยืม", "เวลา", "วันที่กำหนดคืน", "วันที่คืนจริง",
    "เลขงาน", "ชื่อผู้เบิก", "ชื่อลูกค้า", "ชื่ออะไหล่", "จำนวน", "หน่วย",
    "นำมาจาก", "หมายเลขเครื่อง (Serial Number)", "สถานะ", "หมายเหตุ"
  ];
  const rows = exportRecords.map(record => [
    record.id,
    record.borrowDate || record.date,
    record.time,
    record.dueDate || record.due,
    record.actualReturnDate || "",
    record.job,
    record.requesterName || "",
    record.customerName || "",
    getPartNames(record).join(" | "),
    record.qty,
    record.unit,
    record.destination,
    record.sourceSerialNumber || record.serial || "",
    labels[record.status] || record.status,
    record.note || ""
  ]);
  const csv = "\uFEFF" + [headers, ...rows].map(row => row.map(escapeCsv).join(",")).join("\r\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `PartFlow-${toLocalIsoDate()}.csv`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
  showToast("ส่งออกข้อมูลแล้ว", `ดาวน์โหลด ${exportRecords.length} รายการเป็นไฟล์ CSV`);
}

async function copyText(text) {
  try {
    if (navigator.clipboard && window.isSecureContext) {
      await navigator.clipboard.writeText(text);
      return true;
    }
    const textArea = document.createElement("textarea");
    textArea.value = text;
    textArea.setAttribute("readonly", "");
    textArea.style.position = "fixed";
    textArea.style.opacity = "0";
    document.body.appendChild(textArea);
    textArea.select();
    const copied = document.execCommand("copy");
    textArea.remove();
    return copied;
  } catch {
    return false;
  }
}

function closeReturnModal() {
  returnModal.hidden = true;
  document.body.style.overflow = "";
  actualReturnDateInput.setCustomValidity("");
  state.returnId = null;
  state.pendingReturnRequestId = null;
}

document.addEventListener("click", async event => {
  const control = event.target.closest("button, [data-action]");
  if (!control) return;
  if (control.matches("a")) event.preventDefault();
  if (!transactionPage.hidden && ["show-dashboard", "show-history", "show-borrowed"].includes(control.dataset.action)) {
    if (!canLeaveForm()) return;
    resetBorrowForm();
  }

  if (control.dataset.action === "open-form") openForm();
  if (control.dataset.action === "close-form") closeForm();
  if (control.dataset.action === "show-dashboard") showDashboard();
  if (control.dataset.action === "export-data") exportData();
  if (control.dataset.action === "refresh-data") loadRecords();
  if (control.dataset.action === "show-history") {
    historyPager.reset();
    state.filter = "all";
    state.search = "";
    document.querySelector("#globalSearch").value = "";
    showHistory();
  }
  if (control.dataset.action === "show-borrowed") {
    historyPager.reset();
    state.filter = "borrowed";
    state.search = "";
    document.querySelector("#globalSearch").value = "";
    showHistory();
  }
  if (control.dataset.filter) setFilter(control.dataset.filter);

  if (control.dataset.quantity) {
    const quantityInput = document.querySelector("#quantity");
    const change = control.dataset.quantity === "plus" ? 1 : -1;
    quantityInput.value = Math.min(50, Math.max(1, Math.trunc(Number(quantityInput.value || 1)) + change));
    state.pendingCreateRequestId = null;
    state.formDirty = true;
    syncPartNameInputs();
    if (change > 0) partNameList.lastElementChild.querySelector("input").focus();
  }

  if (control.dataset.returnId) {
    if (state.connection === "connecting" || state.connection === "error") {
      showToast("ยังรับคืนไม่ได้", "กรุณารีเฟรชข้อมูลให้เชื่อมต่อ Google Sheet ก่อน");
      return;
    }
    state.returnId = String(control.dataset.returnId);
    state.pendingReturnRequestId = createRequestId();
    const record = records.find(item => String(item.id) === state.returnId);
    if (!record) return;
    const partNames = getPartNames(record);
    document.querySelector("#returnDescription").textContent = `${partNames.length > 1 ? `${partNames.length} ชิ้น` : partNames[0]} · ${record.job}`;
    actualReturnDateInput.value = toLocalIsoDate();
    actualReturnDateInput.min = record.borrowDate || "";
    actualReturnDateInput.setCustomValidity("");
    returnModal.hidden = false;
    document.body.style.overflow = "hidden";
  }

  if (control.dataset.copyId) {
    const record = records.find(item => String(item.id) === String(control.dataset.copyId));
    if (!record) return;
    const copied = await copyText(createFollowUpText(record));
    showToast(copied ? "คัดลอกข้อความแล้ว" : "คัดลอกไม่สำเร็จ", copied ? `พร้อมนำไปติดตามกับ ${record.requesterName || "ผู้เบิก"}` : "โปรดลองคัดลอกอีกครั้งผ่านเบราว์เซอร์");
  }

  if (control.dataset.action === "cancel-return" && !state.returnSaving) closeReturnModal();

  if (control.dataset.action === "confirm-return") {
    const record = records.find(item => String(item.id) === String(state.returnId));
    if (!actualReturnDateInput.reportValidity()) return;
    if (record?.borrowDate && actualReturnDateInput.value < record.borrowDate) {
      actualReturnDateInput.setCustomValidity("วันที่คืนจริงต้องเป็นวันเดียวกับหรือหลังวันที่ยืม");
      actualReturnDateInput.reportValidity();
      return;
    }
    if (!record) return;

    const confirmButton = control;
    const originalLabel = confirmButton.textContent;
    confirmButton.disabled = true;
    state.returnSaving = true;
    confirmButton.textContent = "กำลังบันทึก...";
    try {
      if (isApiConfigured()) {
        const data = await apiRequest("return", {
          id: record.id,
          actualReturnDate: actualReturnDateInput.value,
          requestId: state.pendingReturnRequestId || createRequestId()
        });
        const index = records.findIndex(item => String(item.id) === String(record.id));
        records[index] = normalizeRecord(data.record);
      } else {
        record.status = "returned";
        record.actualReturnDate = actualReturnDateInput.value;
        record.returned = formatThaiDate(record.actualReturnDate);
        record.updatedAt = new Date().toISOString();
      }
      state.returnSaving = false;
      closeReturnModal();
      updateMetrics();
      renderRecords();
      showToast("รับคืนเรียบร้อย", "บันทึกวันที่คืนจริงและอัปเดตสถานะแล้ว");
    } catch (error) {
      showToast("บันทึกการคืนไม่สำเร็จ", error.message);
    } finally {
      state.returnSaving = false;
      confirmButton.disabled = false;
      confirmButton.textContent = originalLabel;
    }
  }
});

document.querySelector("#quantity").addEventListener("change", syncPartNameInputs);
document.querySelector("#quantity").addEventListener("input", event => {
  if (event.target.value !== "") syncPartNameInputs();
});
document.querySelector("#globalSearch").addEventListener("input", event => {
  state.search = event.target.value;
  historyPager.reset();
  renderRecords();
});
borrowDateInput.addEventListener("change", validateBorrowDates);
returnDateInput.addEventListener("change", validateBorrowDates);
actualReturnDateInput.addEventListener("change", () => {
  actualReturnDateInput.setCustomValidity("");
  state.pendingReturnRequestId = createRequestId();
});

document.addEventListener("keydown", event => {
  if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "k") {
    event.preventDefault();
    if (!transactionPage.hidden) {
      if (!canLeaveForm()) return;
      resetBorrowForm();
    }
    showHistory();
    document.querySelector("#globalSearch").focus();
  }
  if (event.key === "Escape") {
    if (!returnModal.hidden && !state.returnSaving) closeReturnModal();
    else if (!transactionPage.hidden) closeForm();
  }
});

document.querySelector("#requesterName").addEventListener("change", event => {
  const other = event.target.value === "__other__";
  document.querySelector("#otherRequesterField").hidden = !other;
  form.elements.otherRequesterName.required = other;
  if (other) form.elements.otherRequesterName.focus();
});

document.querySelector("#destination").addEventListener("change", event => {
  const other = event.target.value === "อื่นๆ";
  document.querySelector("#otherDestinationField").hidden = !other;
  form.elements.otherDestination.required = other;
});

function markFormEdited() {
  state.formDirty = true;
  state.pendingCreateRequestId = null;
}
form.addEventListener("input", markFormEdited);
form.addEventListener("change", markFormEdited);
window.addEventListener("beforeunload", event => {
  if (!state.formDirty && !state.saving) return;
  event.preventDefault();
  event.returnValue = "";
});

form.addEventListener("submit", async event => {
  event.preventDefault();
  if (state.connection === "connecting" || state.connection === "error") {
    showToast("ยังบันทึกไม่ได้", "กรุณารีเฟรชข้อมูลให้เชื่อมต่อ Google Sheet ก่อน");
    return;
  }
  syncPartNameInputs();
  if (!validateBorrowDates()) {
    returnDateInput.reportValidity();
    return;
  }
  if (!form.reportValidity()) return;

  const data = new FormData(form);
  const now = new Date();
  const partNames = collectPartNames();
  const borrowDateValue = data.get("borrowDate");
  const dueDateValue = data.get("returnDate");
  const createdAt = now.toISOString();
  const destination = data.get("destination") === "อื่นๆ" ? data.get("otherDestination").trim() : data.get("destination");
  const payload = {
    borrowDate: borrowDateValue,
    dueDate: dueDateValue,
    jobNumber: data.get("jobNumber").trim(),
    requesterName: data.get("requesterName") === "__other__" ? data.get("otherRequesterName").trim() : data.get("requesterName").trim(),
    customerName: data.get("customerName").trim(),
    partNames,
    quantity: Number(data.get("quantity")),
    unit: data.get("unit"),
    destination,
    sourceSerialNumber: data.get("sourceSerialNumber").trim(),
    note: data.get("note").trim()
  };

  const submitButton = form.querySelector('button[type="submit"]');
  const originalLabel = submitButton.textContent;
  state.pendingCreateRequestId ||= createRequestId();
  submitButton.disabled = true;
  state.saving = true;
  form.inert = true;
  submitButton.textContent = "กำลังบันทึก...";
  try {
    if (isApiConfigured()) {
      const result = await apiRequest("create", {
        record: payload,
        requestId: state.pendingCreateRequestId
      });
      const saved = normalizeRecord(result.record);
      records = [saved, ...records.filter(record => String(record.id) !== saved.id)];
    } else {
      records.unshift(normalizeRecord({
        id: Date.now(),
        ...payload,
        status: "borrowed",
        createdAt,
        updatedAt: createdAt,
        code: "ข้อมูลตัวอย่าง"
      }));
    }

    resetBorrowForm();
    showDashboard();
    state.filter = "all";
    updateMetrics();
    renderRecords();
    showToast("บันทึกรายการแล้ว", `เพิ่มอะไหล่ ${partNames.length} ชิ้นในประวัติเรียบร้อย`);
  } catch (error) {
    showToast("บันทึกรายการไม่สำเร็จ", error.message);
  } finally {
    state.saving = false;
    form.inert = false;
    submitButton.disabled = false;
    submitButton.textContent = originalLabel;
  }
});

function updateMetrics() {
  const today = new Date(`${toLocalIsoDate()}T00:00:00`);
  const previousMonthDate = new Date(today.getFullYear(), today.getMonth() - 1, 1);
  const isSameMonth = (isoDate, target) => {
    if (!isoDate) return false;
    const date = new Date(`${isoDate}T00:00:00`);
    return date.getFullYear() === target.getFullYear() && date.getMonth() === target.getMonth();
  };

  const thisMonth = records.filter(record => isSameMonth(record.borrowDate, today)).length;
  const previousMonth = records.filter(record => isSameMonth(record.borrowDate, previousMonthDate)).length;
  const borrowed = records.filter(record => record.status === "borrowed").length;
  const returned = records.filter(record => record.status === "returned" && isSameMonth(record.actualReturnDate, today)).length;
  const dueDays = records.filter(record => record.status === "borrowed" && record.dueDate).map(record =>
    Math.round((new Date(`${record.dueDate}T00:00:00`) - today) / 86400000)
  );
  const overdue = dueDays.filter(days => days < 0).length;
  const dueSoon = dueDays.filter(days => days >= 0 && days <= 3).length;
  const difference = thisMonth - previousMonth;

  document.querySelector("#metricTotal").textContent = thisMonth;
  document.querySelector("#metricBorrowed").textContent = borrowed;
  document.querySelector("#metricReturned").textContent = returned;
  document.querySelector("#metricDue").textContent = overdue + dueSoon;
  document.querySelector("#metricDueDetail").textContent = `เกินกำหนด ${overdue} · ภายใน 3 วัน ${dueSoon}`;
  document.querySelector("#metricTrend").textContent = previousMonth === 0 ? "ยังไม่มีข้อมูลเดือนก่อน" : `${difference >= 0 ? "+" : ""}${difference} จากเดือนก่อน`;
  document.querySelector("#navBorrowCount").textContent = borrowed;
}

document.querySelectorAll(".mobile-menu").forEach(mobileMenu => mobileMenu.addEventListener("click", () => {
  const open = document.querySelector(".sidebar").classList.toggle("open");
  document.querySelectorAll(".mobile-menu").forEach(button => button.setAttribute("aria-expanded", String(open)));
}));

document.querySelectorAll(".nav-item").forEach(item => item.addEventListener("click", () => {
  closeMobileNavigation();
}));

document.querySelector("#todayLabel").textContent = new Intl.DateTimeFormat("th-TH", { weekday: "long", day: "numeric", month: "long", year: "numeric" }).format(new Date());
applyTheme(document.documentElement.dataset.theme || "light");
populateRequesterOptions();
resetBorrowForm();
loadRecords();
