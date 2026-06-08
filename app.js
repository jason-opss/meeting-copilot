const STORAGE_KEY = "meetingCopilot.static.v1";
const LIBRARY_KEY = "meetingCopilot.library.v1";
const PROFILE_KEY = "meetingCopilot.profile.v1";
const GEMINI_API_BASE = "https://generativelanguage.googleapis.com/v1beta/models";
const DEFAULT_MODEL = "gemini-3.1-flash-lite";
const FALLBACK_MODEL = "gemini-3.5-flash";
const GENERATION_FALLBACK_MODEL_CANDIDATES = [
  "gemini-3.1-flash-lite",
  "gemini-3.5-flash"
];
const TEXT_GENERATION_ALLOWED_MODELS = [
  "gemini-3.1-flash-lite",
  "gemini-3.5-flash"
];
const SEARCH_GROUNDING_MODEL = "gemini-2.5-flash-lite";
const SEARCH_GROUNDING_MODEL_CANDIDATES = [
  "gemini-2.5-flash-lite",
  "gemini-2.5-flash",
  "gemini-2.0-flash"
];
const LIMIT_MESSAGE = "API Key가 없거나 한도가 제한되었습니다.";
const PDFJS_URL = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.10.38/pdf.min.mjs";
const PDFJS_WORKER_URL = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.10.38/pdf.worker.min.mjs";
const DOCX_FINAL_TEMPLATE_URL = "STYLE_WANTED/DOCX_FINAL_WANTED.docx";

const ASSET_OPTIONS = {
  "인프라": {
    strategies: ["Core", "Core+", "Value-Add", "Opportunistic", "Development / Greenfield"],
    sectors: ["신재생에너지", "디지털 인프라 / 데이터센터", "교통 / 물류 인프라", "유틸리티 / 환경", "사회기반시설 PPP", "기타 인프라"]
  },
  "부동산": {
    strategies: ["Core", "Core+", "Value-Add", "Opportunistic", "Development / PF", "NPL / Distressed"],
    sectors: ["오피스", "물류센터", "주거 / 멀티패밀리", "리테일", "호텔 / 호스피탈리티", "데이터센터", "라이프사이언스", "기타 부동산"]
  },
  "사모투자(PE)": {
    strategies: ["Buyout", "Growth Capital", "Venture Capital", "Secondary", "Co-investment", "Special Situations"],
    sectors: ["테크 / 소프트웨어", "헬스케어 / 바이오", "소비재 / 이커머스", "산업재", "금융서비스", "에너지전환", "기타 PE"],
    fixedCapitalType: "Equity"
  },
  "사모투자(PD)": {
    strategies: ["Direct Lending", "Mezzanine", "Distressed Credit", "Special Situations", "Asset-backed Finance", "Structured Credit"],
    sectors: ["기업 인수금융", "부동산 담보대출", "인프라 대출", "NAV Financing", "운전자금 / Growth Debt", "기타 PD"],
    fixedCapitalType: "Debt"
  },
  "상품금융": {
    strategies: ["Balloon 있음", "Balloon 없음"],
    sectors: ["항공기", "선박", "기타"]
  }
};

const CAPITAL_OPTIONS = {
  Equity: ["Common Equity", "Preferred Equity", "RCPS", "CPS", "Co-investment Equity", "LP Commitment"],
  Debt: ["Senior Debt", "Subordinated Debt", "Unitranche", "Bridge Loan", "Project Finance Loan", "Asset-backed Loan"],
  "Hybrid / Mezzanine": ["Mezzanine", "Convertible Bond", "Bond with Warrant", "Preferred Equity + Debt", "Structured Equity"]
};

const runtimeConfig = {
  apiKey: "",
  model: DEFAULT_MODEL
};

const emptyState = () => ({
  meetingId: null,
  activePhase: "prep",
  selectedQuestionIndex: 0,
  selectedFileName: "",
  selectedFileMeta: null,
  imProcessingResult: null,
  marketContext: null,
  preMeetingBrief: null,
  postMeetingMemo: null,
  meeting: {
    managerName: "",
    fundName: "",
    gpParticipants: "",
    lpParticipants: "",
    contactName: "",
    meetingDate: "",
    locationType: "",
    assetClass: "",
    strategy: "",
    sector: "",
    capitalType: "",
    dealType: "",
    investmentStructure: "",
    keyConcerns: ""
  },
  documentSettings: {
    pageReadMode: "auto",
    customPages: ""
  },
  questionRecords: [],
  meetingNotes: "",
  transcript: "",
  reportTone: "neutral",
  lastSavedAt: null
});

let state = emptyState();
let profile = {
  name: "사용자",
  department: "부서 미설정",
  model: DEFAULT_MODEL,
  meetingStorageEnabled: true
};
let selectedFile = null;
let pdfjsLibPromise = null;
let isBusy = false;
let lastGeminiModelUsed = DEFAULT_MODEL;
let lastSearchGroundingModelUsed = SEARCH_GROUNDING_MODEL;
let lastGeminiDiagnostic = null;
let libraryUpdateTimer = null;
let localMeetingStorageEnabled = true;

const $ = (id) => document.getElementById(id);

document.addEventListener("DOMContentLoaded", init);

function init() {
  loadProfile();
  loadState();
  bindEvents();
  applyStateToFields();
  renderAll();
  refreshIcons();
}

function bindEvents() {
  document.querySelectorAll("[data-phase]").forEach((button) => {
    button.addEventListener("click", () => switchPhase(button.dataset.phase));
  });

  $("newMeetingButton").addEventListener("click", resetMeeting);
  $("openSettingsButton").addEventListener("click", openSettings);
  $("mobileSettingsButton").addEventListener("click", openSettings);
  $("closeSettingsButton").addEventListener("click", closeSettings);
  $("cancelSettingsButton").addEventListener("click", closeSettings);
  $("saveSettingsButton").addEventListener("click", saveSettings);
  $("testSettingsButton").addEventListener("click", testSettingsConnection);
  $("clearLocalStorageButton").addEventListener("click", clearLocalMeetingStorage);
  $("exportAndClearStorageButton").addEventListener("click", exportMeetingStorageAndClear);

  $("saveButton").addEventListener("click", () => {
    saveCurrentMeeting("최근 미팅에 저장했습니다.");
  });
  $("savePrepButton").addEventListener("click", () => {
    saveCurrentMeeting("최근 미팅에 저장했습니다.");
  });
  $("exportButton").addEventListener("click", exportMarkdown);
  $("copyAllButton").addEventListener("click", () => copyText(buildFullMarkdown(), "전체 내용을 복사했습니다."));
  $("copyReportButton").addEventListener("click", () => copyText(buildReportMarkdown(), "보고서를 복사했습니다."));
  $("exportDocxButton").addEventListener("click", exportReportDocx);
  $("regenerateReportButton").addEventListener("click", generateReport);
  $("generateBriefButton").addEventListener("click", generateBrief);
  $("generateReportButton").addEventListener("click", generateReport);
  $("addQuestionButton").addEventListener("click", addQuestion);
  $("timestampButton").addEventListener("click", insertTimestamp);

  const dropZone = $("dropZone");
  const fileInput = $("fileInput");
  dropZone.addEventListener("click", () => fileInput.click());
  fileInput.addEventListener("change", (event) => handleFiles(event.target.files));
  $("removeFileButton").addEventListener("click", removeSelectedFile);
  ["dragenter", "dragover", "dragleave", "drop"].forEach((eventName) => {
    dropZone.addEventListener(eventName, preventDefaults);
  });
  ["dragenter", "dragover"].forEach((eventName) => {
    dropZone.addEventListener(eventName, () => dropZone.classList.add("dragover"));
  });
  ["dragleave", "drop"].forEach((eventName) => {
    dropZone.addEventListener(eventName, () => dropZone.classList.remove("dragover"));
  });
  dropZone.addEventListener("drop", (event) => handleFiles(event.dataTransfer.files));

  [
    "managerName",
    "fundName",
    "gpParticipants",
    "lpParticipants",
    "meetingDate",
    "locationType",
    "assetClass",
    "strategy",
    "sector",
    "capitalType",
    "dealType",
    "investmentStructure",
    "keyConcerns",
    "headerManagerName",
    "headerFundName",
    "pageReadMode",
    "customPages",
    "meetingNotes",
    "transcript",
    "reportTone"
  ].forEach((id) => {
    $(id).addEventListener("input", handleFieldInput);
    $(id).addEventListener("change", handleFieldInput);
  });

  $("assetClass").addEventListener("change", () => {
    populateAssetDependentSelects();
    syncFieldsToState();
    saveState();
    scheduleExistingLibraryUpdate();
  });
  $("capitalType").addEventListener("change", () => {
    populateCapitalStructureSelect();
    syncFieldsToState();
    saveState();
    scheduleExistingLibraryUpdate();
  });
}

function preventDefaults(event) {
  event.preventDefault();
  event.stopPropagation();
}

function handleFiles(files) {
  if (!files || !files.length) return;
  selectedFile = files[0];
  state.selectedFileName = selectedFile.name;
  state.selectedFileMeta = {
    name: selectedFile.name,
    size: selectedFile.size,
    type: selectedFile.type || guessMimeType(selectedFile.name)
  };
  state.imProcessingResult = null;
  state.preMeetingBrief = null;
  state.marketContext = null;
  renderFileState();
  saveState();
  scheduleExistingLibraryUpdate();
}

function removeSelectedFile(event) {
  event?.preventDefault();
  event?.stopPropagation();
  selectedFile = null;
  state.selectedFileName = "";
  state.selectedFileMeta = null;
  state.imProcessingResult = null;
  state.marketContext = null;
  state.preMeetingBrief = null;
  state.postMeetingMemo = null;
  state.questionRecords = [];
  state.selectedQuestionIndex = 0;
  const fileInput = $("fileInput");
  if (fileInput) fileInput.value = "";
  renderAll();
  saveState();
  scheduleExistingLibraryUpdate();
  toast("선택한 IM 파일을 제거했습니다.");
}

function handleFieldInput(event) {
  const { id, value } = event.target;
  if (id === "headerManagerName") $("managerName").value = value;
  if (id === "headerFundName") $("fundName").value = value;
  if (id === "managerName") $("headerManagerName").value = value;
  if (id === "fundName") $("headerFundName").value = value;
  syncFieldsToState();
  saveState();
  scheduleExistingLibraryUpdate();
  renderStatus();
}

function populateAssetDependentSelects(preferredStrategy = "", preferredSector = "") {
  const assetClass = $("assetClass").value;
  const config = ASSET_OPTIONS[assetClass];
  fillSelect($("strategy"), config?.strategies || [], "자산군을 먼저 선택", preferredStrategy);
  fillSelect($("sector"), config?.sectors || [], "자산군을 먼저 선택", preferredSector);
  applyFixedCapitalTypeForAsset(config);
}

function populateCapitalStructureSelect(preferredStructure = "") {
  const capitalType = $("capitalType").value;
  const options = CAPITAL_OPTIONS[capitalType] || [];
  fillSelect($("investmentStructure"), options, "Equity / Debt를 먼저 선택", preferredStructure);
}

function fillSelect(select, options, emptyLabel, preferredValue = "") {
  select.innerHTML = "";
  const empty = document.createElement("option");
  empty.value = "";
  empty.textContent = emptyLabel;
  select.appendChild(empty);
  options.forEach((value) => {
    const option = document.createElement("option");
    option.value = value;
    option.textContent = value;
    select.appendChild(option);
  });
  select.disabled = options.length === 0;
  if (preferredValue && options.includes(preferredValue)) {
    select.value = preferredValue;
  } else {
    select.value = "";
  }
}

function applyFixedCapitalTypeForAsset(config) {
  const capitalSelect = $("capitalType");
  if (config?.fixedCapitalType) {
    capitalSelect.value = config.fixedCapitalType;
    capitalSelect.disabled = true;
    capitalSelect.title = "사모투자 PE/PD는 자산군에서 Equity/Debt 성격이 이미 정해집니다.";
    populateCapitalStructureSelect();
    return;
  }
  capitalSelect.disabled = false;
  capitalSelect.title = "";
  populateCapitalStructureSelect(state.meeting.investmentStructure || "");
}

function inferCapitalType(value = "") {
  const normalized = value.toLowerCase();
  if (!normalized) return "";
  if (["senior", "debt", "loan", "unitranche", "bridge"].some((word) => normalized.includes(word))) return "Debt";
  if (["mezzanine", "convertible", "hybrid", "cb", "bw"].some((word) => normalized.includes(word))) return "Hybrid / Mezzanine";
  if (["equity", "rcps", "cps", "preferred", "common"].some((word) => normalized.includes(word))) return "Equity";
  return "";
}

function syncFieldsToState() {
  state.meeting = {
    managerName: $("managerName").value.trim(),
    fundName: $("fundName").value.trim(),
    gpParticipants: $("gpParticipants").value.trim(),
    lpParticipants: $("lpParticipants").value.trim(),
    contactName: $("gpParticipants").value.trim(),
    meetingDate: $("meetingDate").value,
    locationType: $("locationType").value,
    assetClass: $("assetClass").value,
    strategy: $("strategy").value,
    sector: $("sector").value,
    capitalType: $("capitalType").value,
    dealType: $("dealType").value,
    investmentStructure: $("investmentStructure").value,
    keyConcerns: $("keyConcerns").value.trim()
  };
  state.documentSettings = {
    pageReadMode: $("pageReadMode").value,
    customPages: $("customPages").value.trim()
  };
  state.meetingNotes = $("meetingNotes").value;
  state.transcript = $("transcript").value;
  state.reportTone = $("reportTone").value;
}

function applyStateToFields() {
  $("managerName").value = state.meeting.managerName || "";
  $("fundName").value = state.meeting.fundName || "";
  $("headerManagerName").value = state.meeting.managerName || "";
  $("headerFundName").value = state.meeting.fundName || "";
  $("gpParticipants").value = state.meeting.gpParticipants || state.meeting.contactName || "";
  $("lpParticipants").value = state.meeting.lpParticipants || "";
  $("meetingDate").value = state.meeting.meetingDate || "";
  $("locationType").value = state.meeting.locationType || "";
  $("assetClass").value = state.meeting.assetClass || "";
  populateAssetDependentSelects(state.meeting.strategy || state.meeting.sectorStrategy || "", state.meeting.sector || "");
  if (!ASSET_OPTIONS[$("assetClass").value]?.fixedCapitalType) {
    $("capitalType").value = state.meeting.capitalType || inferCapitalType(state.meeting.investmentStructure || "");
    populateCapitalStructureSelect(state.meeting.investmentStructure || "");
    $("investmentStructure").value = state.meeting.investmentStructure || "";
  }
  $("dealType").value = state.meeting.dealType || "";
  $("keyConcerns").value = state.meeting.keyConcerns || "";
  $("pageReadMode").value = "auto";
  $("customPages").value = "";
  $("meetingNotes").value = state.meetingNotes || "";
  $("transcript").value = state.transcript || "";
  $("reportTone").value = state.reportTone || "neutral";
}

function loadProfile() {
  try {
    profile = { ...profile, ...JSON.parse(localStorage.getItem(PROFILE_KEY) || "{}") };
  } catch {
    profile = { name: "사용자", department: "부서 미설정", model: DEFAULT_MODEL, meetingStorageEnabled: true };
  }
  runtimeConfig.model = normalizeTextGenerationModel(profile.model || DEFAULT_MODEL);
  profile.model = runtimeConfig.model;
  localMeetingStorageEnabled = profile.meetingStorageEnabled !== false;
}

function saveProfile() {
  const safeProfile = {
    name: profile.name || "사용자",
    department: profile.department || "부서 미설정",
    model: normalizeTextGenerationModel(profile.model || DEFAULT_MODEL),
    meetingStorageEnabled: localMeetingStorageEnabled
  };
  localStorage.setItem(PROFILE_KEY, JSON.stringify(safeProfile));
}

function loadState() {
  try {
    const saved = JSON.parse(localStorage.getItem(STORAGE_KEY) || "null");
    if (saved) state = mergeSavedState(saved);
  } catch {
    state = emptyState();
  }
}

function saveState() {
  syncFieldsSilently();
  state.lastSavedAt = new Date().toISOString();
  if (!localMeetingStorageEnabled) return;
  localStorage.setItem(STORAGE_KEY, JSON.stringify(sanitizeStateForStorage(state)));
}

function syncFieldsSilently() {
  if (!$("managerName")) return;
  state.meeting.managerName = $("managerName").value.trim();
  state.meeting.fundName = $("fundName").value.trim();
  state.meeting.gpParticipants = $("gpParticipants").value.trim();
  state.meeting.lpParticipants = $("lpParticipants").value.trim();
  state.meeting.contactName = $("gpParticipants").value.trim();
  state.meeting.meetingDate = $("meetingDate").value;
  state.meeting.locationType = $("locationType").value;
  state.meeting.assetClass = $("assetClass").value;
  state.meeting.strategy = $("strategy").value;
  state.meeting.sector = $("sector").value;
  state.meeting.capitalType = $("capitalType").value;
  state.meeting.dealType = $("dealType").value;
  state.meeting.investmentStructure = $("investmentStructure").value;
  state.meeting.keyConcerns = $("keyConcerns").value.trim();
  state.documentSettings.pageReadMode = "auto";
  state.documentSettings.customPages = "";
  state.meetingNotes = $("meetingNotes").value;
  state.transcript = $("transcript").value;
  state.reportTone = $("reportTone").value;
}

function openSettings() {
  $("settingsName").value = profile.name || "";
  $("settingsDepartment").value = profile.department || "";
  $("settingsModel").value = normalizeTextGenerationModel(runtimeConfig.model || DEFAULT_MODEL);
  $("settingsApiKey").value = "";
  $("settingsApiKey").placeholder = runtimeConfig.apiKey ? "새 API Key를 붙여넣으면 현재 키가 교체됩니다" : "Gemini API Key를 붙여넣으세요";
  $("settingsStorageEnabled").checked = localMeetingStorageEnabled;
  renderSettingsKeyHint();
  renderStorageSettingsStatus();
  renderProviderStatus();
  $("settingsModal").classList.remove("hidden");
  $("settingsModal").classList.add("flex");
  setTimeout(() => {
    $("settingsModal").classList.remove("opacity-0");
    $("settingsPanel").classList.remove("scale-95");
  }, 10);
}

function closeSettings() {
  $("settingsModal").classList.add("opacity-0");
  $("settingsPanel").classList.add("scale-95");
  setTimeout(() => {
    $("settingsModal").classList.add("hidden");
    $("settingsModal").classList.remove("flex");
  }, 180);
}

function saveSettings() {
  applySettingsFromModal();
  closeSettings();
  toast(runtimeConfig.apiKey ? `${runtimeConfig.model} 모델로 Gemini 설정을 현재 탭에 적용했습니다. ${keyFingerprint(runtimeConfig.apiKey)}` : "설정을 현재 탭에 적용했습니다. AI 실행 전 API Key를 입력해야 합니다.");
}

function applySettingsFromModal() {
  profile.name = $("settingsName").value.trim() || "사용자";
  profile.department = $("settingsDepartment").value.trim() || "부서 미설정";
  profile.model = normalizeTextGenerationModel($("settingsModel").value.trim() || DEFAULT_MODEL);
  runtimeConfig.model = profile.model;
  localMeetingStorageEnabled = $("settingsStorageEnabled").checked;
  const nextKey = sanitizeGeminiApiKey($("settingsApiKey").value);
  if (nextKey) {
    runtimeConfig.apiKey = nextKey;
    $("settingsApiKey").value = "";
    $("settingsApiKey").placeholder = "새 API Key를 붙여넣으면 현재 키가 교체됩니다";
  }
  saveProfile();
  renderProfile();
  renderSettingsKeyHint();
  renderStorageSettingsStatus();
  renderProviderStatus();
}

async function testSettingsConnection() {
  const button = $("testSettingsButton");
  try {
    applySettingsFromModal();
    if (!runtimeConfig.apiKey) throw new Error("API Key가 아직 적용되지 않았습니다. 새 키를 붙여넣고 다시 테스트하세요.");
    button.disabled = true;
    button.textContent = "모델 확인 중...";
    const diagnostic = await runGeminiConnectionDiagnostics(normalizeTextGenerationModel(runtimeConfig.model || DEFAULT_MODEL));
    lastGeminiDiagnostic = diagnostic;
    toast(`연결 테스트 성공: ${diagnostic.model} / ${keyFingerprint(runtimeConfig.apiKey)}`);
  } catch (error) {
    lastGeminiDiagnostic = {
      ok: false,
      model: runtimeConfig.model || DEFAULT_MODEL,
      checkedAt: new Date().toISOString(),
      keyFingerprint: keyFingerprint(runtimeConfig.apiKey),
      error: getGeminiErrorInfoFromError(error)
    };
    showError(error);
  } finally {
    button.disabled = false;
    button.textContent = "저장 후 연결 테스트";
    renderSettingsKeyHint();
    renderProviderStatus();
  }
}

async function runGeminiConnectionDiagnostics(model) {
  const checkedAt = new Date().toISOString();
  const models = await fetchGeminiModelList();
  const normalizedModel = normalizeGeminiModelName(model);
  const availableModel = models.find((item) => normalizeGeminiModelName(item.name) === normalizedModel);
  if (!availableModel) {
    throw createGeminiApiError(
      `${model} 모델이 이 API Key/프로젝트에서 조회되지 않습니다. 설정에서 다른 모델을 선택하거나 AI Studio의 모델 권한을 확인하세요.`,
      {
        model,
        status: 404,
        statusText: "MODEL_NOT_LISTED",
        rawMessage: "The selected model was not returned by models.list.",
        reason: "model_not_listed"
      }
    );
  }
  if (Array.isArray(availableModel.supportedActions) && !availableModel.supportedActions.includes("generateContent")) {
    throw createGeminiApiError(
      `${model} 모델은 현재 generateContent 호출을 지원하지 않습니다. 텍스트 생성 가능한 Gemini 모델을 선택하세요.`,
      {
        model,
        status: 400,
        statusText: "UNSUPPORTED_ACTION",
        rawMessage: "The selected model does not support generateContent.",
        reason: "unsupported_action"
      }
    );
  }
  const text = await callGeminiModel(model, {
    contents: [{ role: "user", parts: [{ text: "Reply with only OK." }] }],
    generationConfig: { temperature: 0, maxOutputTokens: 8 }
  });
  return {
    ok: true,
    model,
    checkedAt,
    keyFingerprint: keyFingerprint(runtimeConfig.apiKey),
    modelCount: models.length,
    responsePreview: text.slice(0, 40)
  };
}

async function fetchGeminiModelList() {
  const apiKey = getGeminiApiKeyForRequest();
  const response = await fetch(`${GEMINI_API_BASE}?pageSize=1000`, {
    method: "GET",
    headers: { "x-goog-api-key": apiKey }
  });
  if (!response.ok) {
    const errorText = await response.text().catch(() => "");
    const info = getGeminiErrorInfo(errorText, response.status, "models.list");
    throw createGeminiApiError(info.message, info);
  }
  const data = await response.json();
  return (data.models || []).filter((model) => {
    const actions = model.supportedActions || model.supportedGenerationMethods || model.supported_actions || model.supported_generation_methods || [];
    model.supportedActions = actions;
    return actions.includes("generateContent");
  });
}

function normalizeGeminiModelName(name) {
  return String(name || "").replace(/^models\//, "").trim();
}

function normalizeTextGenerationModel(model) {
  const normalized = normalizeGeminiModelName(model || DEFAULT_MODEL);
  return TEXT_GENERATION_ALLOWED_MODELS.includes(normalized) ? normalized : DEFAULT_MODEL;
}

function renderAll() {
  renderProfile();
  renderFileState();
  renderBrief();
  renderQuestions();
  renderSelectedQuestion();
  renderReport();
  renderRecentMeetings();
  renderProviderStatus();
  renderStatus();
  switchPhase(state.activePhase || "prep", false);
}

function renderProfile() {
  $("profileName").textContent = profile.name || "사용자";
  $("profileDepartment").textContent = profile.department || "부서 미설정";
  $("profileInitial").textContent = makeInitial(profile.name);
}

function makeInitial(name) {
  const cleaned = (name || "LP").trim();
  return cleaned.slice(0, 2).toUpperCase();
}

function renderProviderStatus() {
  if (!$("providerStatus")) return;
  const base = runtimeConfig.apiKey
    ? `Gemini API Key가 현재 탭 메모리에만 적용되어 있습니다. 현재 모델: ${runtimeConfig.model || DEFAULT_MODEL} / 검색 그라운딩: ${lastSearchGroundingModelUsed || SEARCH_GROUNDING_MODEL} (2.5 flash-lite → 2.5 flash → 2.0 flash) / 적용 키: ${keyFingerprint(runtimeConfig.apiKey)}`
    : `Gemini API Key가 없습니다. AI 버튼을 실행하면 "${LIMIT_MESSAGE}" 메시지가 표시됩니다.`;
  const diagnostic = formatGeminiDiagnostic(lastGeminiDiagnostic);
  $("providerStatus").textContent = diagnostic ? `${base}\n${diagnostic}` : base;
}

function formatGeminiDiagnostic(diagnostic) {
  if (!diagnostic) return "";
  const checkedAt = diagnostic.checkedAt ? formatDateTime(diagnostic.checkedAt) : "";
  if (diagnostic.ok) {
    return `최근 연결 테스트: 성공 · ${diagnostic.model} · ${checkedAt}`;
  }
  const error = diagnostic.error || {};
  const code = [error.status, error.statusText || error.reason].filter(Boolean).join(" ");
  return `최근 연결 테스트: 실패 · ${diagnostic.model}${code ? ` · ${code}` : ""} · ${error.message || error.rawMessage || LIMIT_MESSAGE}`;
}

function renderSettingsKeyHint() {
  const hint = $("settingsKeyHint");
  if (!hint) return;
  hint.textContent = runtimeConfig.apiKey
    ? `현재 탭에 적용된 키: ${keyFingerprint(runtimeConfig.apiKey)}. 새 키를 입력하고 저장하면 즉시 교체됩니다.`
    : "현재 적용된 API Key가 없습니다.";
}

function renderStorageSettingsStatus() {
  const box = $("storageStatus");
  if (!box) return;
  const recentCount = getLibrary().length;
  const hasCurrent = Boolean(localStorage.getItem(STORAGE_KEY));
  box.textContent = localMeetingStorageEnabled
    ? `로컬 저장 켜짐: 현재 미팅${hasCurrent ? " 저장됨" : " 미저장"} / 최근 미팅 ${recentCount}개. API Key는 저장하지 않습니다. 이름/부서/모델은 편의를 위해 저장됩니다.`
    : "로컬 저장 꺼짐: 현재 탭에서 작업은 가능하지만 브라우저를 닫으면 미팅 데이터가 남지 않습니다. 이름/부서/모델은 유지됩니다.";
}

function keyFingerprint(key) {
  if (!key) return "키 없음";
  const trimmed = sanitizeGeminiApiKey(key);
  const tail = trimmed.slice(-4);
  return `•••• ${tail}`;
}

function sanitizeGeminiApiKey(value) {
  const raw = String(value || "").trim();
  if (!raw) return "";
  const compact = raw.replace(/[\s"'`<>]/g, "");
  const matched = compact.match(/(?:AQ\.[A-Za-z0-9_-]+|AIza[A-Za-z0-9_-]+)/);
  return matched ? matched[0] : compact;
}

function getGeminiApiKeyForRequest() {
  const key = sanitizeGeminiApiKey(runtimeConfig.apiKey);
  if (!key) throw new Error(LIMIT_MESSAGE);
  if (!/^[A-Za-z0-9_.-]+$/.test(key)) {
    throw new Error("API Key에 키 이외의 문자가 섞여 있습니다. 키 값만 다시 붙여넣고 저장하세요.");
  }
  runtimeConfig.apiKey = key;
  return key;
}

function renderFileState() {
  const removeButton = $("removeFileButton");
  if (state.selectedFileName) {
    $("dropZoneTitle").textContent = state.selectedFileName;
    $("dropZoneDescription").textContent = selectedFile
      ? `선택 완료: ${formatBytes(selectedFile.size)}`
      : "이전 세션의 파일명만 남아 있습니다. 다시 분석하려면 파일을 다시 선택하세요.";
    $("fileModeBadge").textContent = "IM 있음";
    $("fileModeBadge").className = "rounded-full bg-brand-100 px-2.5 py-1 text-xs font-bold text-brand-700";
    removeButton?.classList.remove("hidden");
  } else {
    $("dropZoneTitle").textContent = "IM, Teaser, 이미지 PDF를 업로드하세요";
    $("dropZoneDescription").textContent = "PDF, TXT, PNG, JPG 지원. 큰 문서는 핵심 페이지 중심으로 읽습니다.";
    $("fileModeBadge").textContent = "선택 사항";
    $("fileModeBadge").className = "rounded-full bg-slate-100 px-2.5 py-1 text-xs font-bold text-slate-500";
    removeButton?.classList.add("hidden");
  }
  renderProcessingLog();
}

function renderProcessingLog() {
  const box = $("processingLog");
  if (!state.imProcessingResult) {
    box.classList.add("hidden");
    box.innerHTML = "";
    return;
  }
  const result = state.imProcessingResult;
  const pages = result.selectedPageNumbers?.length ? `선택 페이지: ${result.selectedPageNumbers.join(", ")}` : "선택 페이지 없음";
  box.classList.remove("hidden");
  box.innerHTML = `
    <div class="font-extrabold text-slate-900">문서 처리 결과</div>
    <div class="mt-2 space-y-1 text-sm">
      <div>처리 방식: ${escapeHtml(result.analysisMode || "standard")}</div>
      <div>${escapeHtml(pages)}</div>
      <div>읽은 텍스트: ${formatNumber(result.textCharCount || 0)}자</div>
      ${result.warning ? `<div class="text-amber-700">${escapeHtml(result.warning)}</div>` : ""}
    </div>
  `;
}

function renderStatus() {
  const badge = $("headerStatus");
  if (state.postMeetingMemo) {
    setBadge(badge, "보고서 작성 완료", "emerald");
  } else if (state.preMeetingBrief) {
    setBadge(badge, "브리프 준비 완료", "brand");
  } else if (state.selectedFileName || state.meeting.managerName || state.meeting.fundName) {
    setBadge(badge, "초안 작성 중", "slate");
  } else {
    setBadge(badge, "새 미팅", "slate");
  }
}

function setBadge(element, text, tone) {
  const classes = {
    emerald: "inline-flex items-center rounded-full bg-emerald-100 px-2.5 py-1 text-xs font-bold text-emerald-700",
    brand: "inline-flex items-center rounded-full bg-brand-100 px-2.5 py-1 text-xs font-bold text-brand-700",
    slate: "inline-flex items-center rounded-full bg-slate-100 px-2.5 py-1 text-xs font-bold text-slate-600"
  };
  element.className = classes[tone] || classes.slate;
  element.textContent = text;
}

function switchPhase(phase, persist = true) {
  state.activePhase = phase;
  ["prep", "live", "report"].forEach((name) => {
    document.querySelector(`[data-phase="${name}"]`).classList.toggle("active", name === phase);
    $(`phase${capitalize(name)}`).classList.toggle("hidden", name !== phase);
  });
  if (persist) saveState();
  refreshIcons();
}

async function generateBrief() {
  if (isBusy) {
    toast("이미 AI 작업이 진행 중입니다.");
    return;
  }
  try {
    syncFieldsToState();
    validateBriefInputs();
    requireApiKey();
    showLoader("사전 브리프를 생성하고 있습니다.", "IM과 세팅값을 통합하고 시장, 뉴스, 정책, 리스크 맥락을 보강합니다.");

    if (selectedFile) {
      updateLoader("IM을 읽는 중입니다.", "큰 문서는 핵심 페이지 중심으로 읽습니다.");
      state.imProcessingResult = await processSelectedIm();
      hydrateMeetingFieldsFromAnalysis();
      renderProcessingLog();
    } else if (state.selectedFileName && !selectedFile) {
      state.imProcessingResult = null;
    }

    updateLoader("시장 맥락을 검색하는 중입니다.", "Gemini Google Search grounding으로 최근 시장, 뉴스, 정책 자료를 확인합니다.");
    try {
      state.marketContext = await fetchMarketContext();
    } catch (marketError) {
      console.warn("Market grounding failed; continuing without external market context.", marketError);
      state.marketContext = buildMarketContextFallback(marketError);
      toast("시장검색이 일시 실패해 IM/세팅값 기반으로 브리프를 계속 생성합니다.");
    }

    updateLoader("질문과 리스크를 정리하는 중입니다.", "LP 미팅 관점의 예상 Q&A, Red Flag, 후속 요청자료를 구성합니다.");
    const prompt = buildBriefPrompt();
    const brief = await callGeminiText(prompt, { json: true, temperature: 0.25 });
    state.preMeetingBrief = normalizeBrief(parseGeminiJson(brief));
    updateLoader("Q&A 적합성을 검토하는 중입니다.", "GP에게 실제로 물을 수 있는 질문인지 다시 점검하고, LP 내부 입력 출처 표현은 딜 사실관계 질문으로 바꿉니다.");
    state.preMeetingBrief = await refineBriefQuestionsForGp(state.preMeetingBrief);
    hydrateMeetingFieldsFromAnalysis();
    state.questionRecords = makeQuestionRecords(state.preMeetingBrief);
    state.selectedQuestionIndex = 0;
    state.postMeetingMemo = null;

    saveState();
    saveMeetingToLibrary();
    renderAll();
    switchPhase("live");
    toast("사전 브리프를 생성하고 미팅 노트로 이동했습니다.");
  } catch (error) {
    showError(error);
  } finally {
    hideLoader();
  }
}

function validateBriefInputs() {
  const requiredWithoutIm = [
    ["managerName", "운용사명"],
    ["fundName", "펀드명 / 대출명"],
    ["locationType", "실제 투자지역"],
    ["assetClass", "자산군"],
    ["strategy", "전략"],
    ["sector", "섹터"],
    ["capitalType", "Equity / Debt"],
    ["investmentStructure", "투자구조"]
  ];
  if (selectedFile || state.selectedFileName) return;
  const missing = requiredWithoutIm
    .filter(([key]) => !state.meeting[key])
    .map(([, label]) => label);
  if (missing.length) {
    throw new Error(`IM이 없을 때는 ${missing.join(", ")} 입력이 필요합니다.`);
  }
}

function hydrateMeetingFieldsFromAnalysis() {
  const patch = deriveEffectiveMeetingInfo();
  const writableKeys = [
    "managerName",
    "fundName",
    "locationType",
    "assetClass",
    "strategy",
    "sector",
    "capitalType",
    "dealType",
    "investmentStructure"
  ];
  let changed = false;
  writableKeys.forEach((key) => {
    if (!state.meeting[key] && patch[key]) {
      state.meeting[key] = patch[key];
      changed = true;
    }
  });
  if (changed) {
    applyStateToFields();
    renderStatus();
  }
}

function requireApiKey() {
  getGeminiApiKeyForRequest();
}

async function processSelectedIm() {
  const file = selectedFile;
  const mimeType = file.type || guessMimeType(file.name);
  if (mimeType.startsWith("image/")) return processImageFile(file, mimeType);
  if (mimeType === "text/plain" || file.name.toLowerCase().endsWith(".txt")) return processTextFile(file);
  if (mimeType === "application/pdf" || file.name.toLowerCase().endsWith(".pdf")) return processPdfFile(file);
  throw new Error("지원하지 않는 파일 형식입니다. PDF, TXT, PNG, JPG 파일을 사용해주세요.");
}

async function processTextFile(file) {
  const text = await file.text();
  const policy = chooseFocusedPolicy({ file, text, pageCount: 1 });
  const excerpt = policy.mode === "focused" ? text.slice(0, 70000) : text.slice(0, 120000);
  const analysis = await analyzeImText(excerpt, policy, { fileName: file.name, mimeType: file.type || "text/plain" });
  return {
    uploadedFileMetadata: makeFileMeta(file),
    analysisMode: policy.mode,
    analysisReasons: policy.reasons,
    selectedPageNumbers: [1],
    textExcerpt: excerpt,
    textCharCount: excerpt.length,
    imAnalysis: analysis,
    warning: policy.warning
  };
}

async function processImageFile(file, mimeType) {
  const policy = chooseFocusedPolicy({ file, text: "", pageCount: 1 });
  const base64 = await fileToBase64(file);
  const prompt = buildImVisionPrompt(policy, makeFileMeta(file));
  const text = await callGeminiVision([{ mimeType, data: base64 }], prompt);
  const analysis = parseGeminiJson(text);
  return {
    uploadedFileMetadata: makeFileMeta(file),
    analysisMode: "vision",
    analysisReasons: policy.reasons,
    selectedPageNumbers: [1],
    textExcerpt: "",
    textCharCount: 0,
    imAnalysis: analysis,
    warning: policy.warning
  };
}

async function processPdfFile(file) {
  const pdf = await extractPdfText(file);
  const policy = chooseFocusedPolicy({ file, text: pdf.fullText, pageCount: pdf.pageCount });
  const selectedPages = selectPages(pdf.pages, policy);
  const selectedText = selectedPages.map((page) => `\n\n[Page ${page.pageNumber}]\n${page.text}`).join("").slice(0, policy.mode === "focused" ? 80000 : 140000);

  let analysis;
  let warning = policy.warning;
  if (selectedText.trim().length > 1000) {
    analysis = await analyzeImText(selectedText, policy, makeFileMeta(file));
  } else {
    const images = await renderPdfPages(file, selectedPages.slice(0, 6).map((page) => page.pageNumber));
    if (!images.length) throw new Error("PDF를 브라우저에서 읽지 못했습니다. 페이지 범위를 줄여 다시 시도해주세요.");
    const prompt = buildImVisionPrompt(policy, makeFileMeta(file));
    const text = await callGeminiVision(images, prompt);
    analysis = parseGeminiJson(text);
    warning = warning || "텍스트가 적어 이미지 페이지를 Vision으로 읽었습니다.";
  }

  return {
    uploadedFileMetadata: makeFileMeta(file),
    analysisMode: policy.mode,
    analysisReasons: policy.reasons,
    selectedPageNumbers: selectedPages.map((page) => page.pageNumber),
    textExcerpt: selectedText,
    textCharCount: selectedText.length,
    imAnalysis: analysis,
    warning
  };
}

async function analyzeImText(text, policy, fileMeta) {
  const response = await callGeminiText(buildImAnalysisPrompt(text, policy, fileMeta), { json: true, temperature: 0.15 });
  return parseGeminiJson(response);
}

function chooseFocusedPolicy({ file, text, pageCount }) {
  const reasons = [];
  if (file.size > 8 * 1024 * 1024) reasons.push("파일 용량이 큼");
  if (pageCount > 40) reasons.push("페이지 수가 많음");
  if ((text || "").length > 60000) reasons.push("추출 텍스트가 김");
  if (state.documentSettings.pageReadMode === "focused") reasons.push("사용자가 Focused skim mode 선택");
  if (state.documentSettings.pageReadMode === "custom") reasons.push("사용자 지정 페이지 우선");
  const shouldFocus = state.documentSettings.pageReadMode !== "auto" || reasons.length > 0;
  return {
    mode: shouldFocus ? "focused" : "standard",
    reasons,
    warning: shouldFocus ? "큰 IM 또는 지정 조건에 따라 핵심 정보 중심으로만 읽었습니다." : ""
  };
}

function selectPages(pages, policy) {
  if (!pages.length) return [];
  if (state.documentSettings.pageReadMode === "custom") {
    const custom = parsePageRanges(state.documentSettings.customPages, pages.length);
    if (custom.length) return custom.map((pageNumber) => pages[pageNumber - 1]).filter(Boolean);
  }
  if (policy.mode !== "focused") return pages.slice(0, 60);

  const firstPages = pages.slice(0, Math.min(5, pages.length));
  const keywordPages = [...pages]
    .map((page) => ({ ...page, score: scorePage(page.text) }))
    .sort((a, b) => b.score - a.score)
    .filter((page) => page.score > 0)
    .slice(0, 10);
  const map = new Map();
  [...firstPages, ...keywordPages].forEach((page) => map.set(page.pageNumber, page));
  return [...map.values()].sort((a, b) => a.pageNumber - b.pageNumber).slice(0, 14);
}

function scorePage(text = "") {
  const keywords = [
    "asset", "strategy", "sector", "region", "deal", "overview", "risk", "sponsor", "borrower",
    "manager", "portfolio", "track record", "irr", "moic", "ltv", "dscr", "exit", "fee",
    "자산", "분류", "전략", "섹터", "지역", "딜", "개요", "거래", "관계자", "운용", "스폰서",
    "차주", "담보", "리스크", "위험", "수익률", "회수", "수수료", "트랙레코드", "투자구조"
  ];
  const lowered = text.toLowerCase();
  return keywords.reduce((score, keyword) => score + (lowered.includes(keyword.toLowerCase()) ? 1 : 0), 0);
}

async function extractPdfText(file) {
  const pdfjs = await loadPdfJs();
  const arrayBuffer = await file.arrayBuffer();
  const loadingTask = pdfjs.getDocument({ data: arrayBuffer });
  const doc = await loadingTask.promise;
  const pages = [];
  for (let pageNumber = 1; pageNumber <= doc.numPages; pageNumber += 1) {
    const page = await doc.getPage(pageNumber);
    const content = await page.getTextContent();
    const text = content.items.map((item) => item.str).join(" ").replace(/\s+/g, " ").trim();
    pages.push({ pageNumber, text });
  }
  return {
    pageCount: doc.numPages,
    pages,
    fullText: pages.map((page) => page.text).join("\n")
  };
}

async function renderPdfPages(file, pageNumbers) {
  const pdfjs = await loadPdfJs();
  const arrayBuffer = await file.arrayBuffer();
  const doc = await pdfjs.getDocument({ data: arrayBuffer }).promise;
  const images = [];
  for (const pageNumber of pageNumbers) {
    const page = await doc.getPage(pageNumber);
    const viewport = page.getViewport({ scale: 1.2 });
    const canvas = document.createElement("canvas");
    canvas.width = viewport.width;
    canvas.height = viewport.height;
    const context = canvas.getContext("2d");
    await page.render({ canvasContext: context, viewport }).promise;
    const dataUrl = canvas.toDataURL("image/jpeg", 0.72);
    images.push({ mimeType: "image/jpeg", data: dataUrl.split(",")[1] });
  }
  return images;
}

async function loadPdfJs() {
  if (!pdfjsLibPromise) {
    pdfjsLibPromise = import(PDFJS_URL).then((lib) => {
      lib.GlobalWorkerOptions.workerSrc = PDFJS_WORKER_URL;
      return lib;
    });
  }
  return pdfjsLibPromise;
}

function getAssetClassMarketChecklist() {
  const meeting = state.meeting || {};
  const im = state.imProcessingResult?.imAnalysis || {};
  const assetClass = localizeDisplayTerm(meeting.assetClass || im.autoDetectedFields?.assetClass || "");
  const region = localizeDisplayTerm(meeting.locationType || im.autoDetectedFields?.region || "");
  const strategy = localizeDisplayTerm(meeting.strategy || im.autoDetectedFields?.strategy || "");
  const sector = localizeDisplayTerm(meeting.sector || im.autoDetectedFields?.sector || "");
  const dealName = meeting.fundName || "";
  const managerName = meeting.managerName || "";
  const regionScope = region === "국내" ? "한국" : region === "해외" ? "해외/미국/주요 투자지역" : "해당 투자지역";
  const base = [
    `${regionScope} 거시경제와 금리 환경`,
    `${managerName} 운용사 관련 최근 뉴스와 평판 이슈`,
    `${dealName} 거래명/대출명/펀드명 관련 뉴스`,
    "IM에 언급된 개별 자산, 차주, 시공사, 보증기관, 포트폴리오 회사, 파이프라인 딜 관련 뉴스"
  ];
  let checklist = [];
  let prioritySearchQueries = [];

  if (/사모투자\(PE\)|Private Equity|PE/i.test(assetClass)) {
    checklist = [
      `${regionScope} 거시경제, 기준금리, 인플레이션, 환율 환경`,
      "Buyout/Growth/VC 등 해당 전략의 거래량, 밸류에이션, Exit, IPO/M&A 분위기",
      "레버리지론, 사모대출, 조달비용, 리파이낸싱 시장",
      `${sector || "해당 섹터"} 산업 성장성, 멀티플, 규제/기술 변화`,
      "해당 GP의 펀드레이징, 투자/회수 실적, 포트폴리오 관련 뉴스",
      "IM에 언급된 편입 자산 또는 편입 예정 파이프라인 회사 뉴스"
    ];
    prioritySearchQueries = [
      `${regionScope} private equity buyout market exit M&A IPO valuation`,
      `${regionScope} leveraged loan private equity financing market`,
      `${managerName} ${dealName} portfolio pipeline news`
    ];
  } else if (/사모투자\(PD\)|Private Debt|PD/i.test(assetClass)) {
    checklist = [
      `${regionScope} 금리, 크레딧 스프레드, 부실률, 리파이낸싱 환경`,
      "Direct Lending/Mezzanine/Structured Credit 등 해당 전략 시장 분위기",
      "차주 섹터의 현금흐름, 담보가치, 회수율, 코버넌트 리스크",
      "은행 대출 축소, 사모대출 공급, 채권시장 유동성",
      "해당 GP의 크레딧 운용 실적과 부실/회수 관련 뉴스",
      "IM에 언급된 차주, 담보자산, 후순위/선순위 구조 관련 뉴스"
    ];
    prioritySearchQueries = [
      `${regionScope} private debt direct lending credit spread default refinancing`,
      `${sector || ""} private credit borrower covenant risk`,
      `${managerName} ${dealName} credit loan news`
    ];
  } else if (/부동산|Real Estate/i.test(assetClass)) {
    checklist = [
      `${regionScope} 거시경제, 기준금리, 부동산 금융/PF 조달 환경`,
      `${sector || "해당 섹터"} 부동산 임대, 공실, 거래, 캡레이트, 가격 동향`,
      `${strategy || "해당 전략"} 전략 관련 개발, 인허가, 분양, 매각, 리파이낸싱 시장`,
      "주요 권역 거래 사례와 유사 자산 매매/임대 사례",
      "시공사, 차주, 보증기관, 신탁사, 대주단 관련 뉴스",
      "프로젝트 딜이면 소재지 인근 거래 사례와 해당 자산/부지 관련 뉴스"
    ];
    prioritySearchQueries = [
      `${regionScope} ${sector || "오피스 부동산"} 거래 임대 공실 캡레이트 금리`,
      `${regionScope} 부동산 PF 대출 조달 리파이낸싱 연체`,
      `${dealName} ${managerName} 부동산 PF 프로젝트 뉴스`
    ];
  } else if (/인프라|Infrastructure/i.test(assetClass)) {
    checklist = [
      `${regionScope} 금리, 물가, 환율과 인프라 자산 할인율 환경`,
      `${sector || "해당 섹터"} 수요, 이용량, 요금체계, 규제/정책 변화`,
      "사업권, 인허가, 정부계획, 공공기관 정책, 보조금/요금 조정 이슈",
      "EPC, O&M, 주요 계약상대방, 건설/운영 리스크",
      "유사 인프라 거래 사례와 수익률/매각 사례",
      "해당 GP 및 프로젝트/운영사 관련 뉴스"
    ];
    prioritySearchQueries = [
      `${regionScope} infrastructure ${sector || ""} regulation tariff demand policy`,
      `${regionScope} infrastructure transaction yield financing interest rate`,
      `${managerName} ${dealName} infrastructure project news`
    ];
  } else if (/상품금융|Commodity Finance/i.test(assetClass)) {
    checklist = [
      `${regionScope} 금리, 환율, 물류/교역량, 자산가치 환경`,
      `${sector || "항공기 선박"} 운임, 리스료, 잔존가치, 중고거래 시장`,
      "차주/리스이용자 신용도, 보험, 정비, 재매각/재리스 가능성",
      "Balloon 상환 여부, 만기 잔존가치, 담보 처분 가능성",
      "선박/항공기 규제, 환경규제, 운항/발주/폐선 동향",
      "관련 자산, 차주, 운용사 뉴스"
    ];
    prioritySearchQueries = [
      `${regionScope} ${sector || "aircraft shipping"} lease rate residual value market`,
      `${sector || "aircraft shipping"} finance balloon repayment collateral risk`,
      `${managerName} ${dealName} commodity finance news`
    ];
  } else {
    checklist = [
      `${regionScope} 거시경제와 금리 환경`,
      `${assetClass || "해당 자산군"} 시장 동향과 거래 분위기`,
      `${strategy || "해당 전략"} 전략 관련 리스크와 회수/상환 환경`,
      `${sector || "해당 섹터"} 산업/시장 뉴스`,
      "해당 GP와 IM에 언급된 주요 거래상대방 뉴스"
    ];
    prioritySearchQueries = [
      `${regionScope} ${assetClass} ${strategy} ${sector} market trend risk`,
      `${managerName} ${dealName} news`
    ];
  }

  return {
    assetClass: assetClass || "확인 필요",
    baselineChecklist: [...base, ...checklist].filter(Boolean),
    prioritySearchQueries: prioritySearchQueries.map((query) => query.replace(/\s+/g, " ").trim()).filter(Boolean)
  };
}

function buildMarketSearchQueries() {
  const meeting = deriveEffectiveMeetingInfo();
  const im = state.imProcessingResult?.imAnalysis || {};
  const detected = im.autoDetectedFields || {};
  const checklist = getAssetClassMarketChecklist();
  const specific = extractSpecificMarketSearchSignals();
  const parts = [
    meeting.managerName,
    meeting.fundName,
    meeting.keyConcerns,
    detected.region,
    detected.assetClass,
    detected.strategy,
    detected.sector,
    detected.capitalStructure,
    ...asArray(im.keyRisks).slice(0, 3).map(formatListItemText),
    ...asArray(im.keyNumbersToVerify).slice(0, 3).map(formatListItemText),
    ...asArray(im.verificationItems).slice(0, 3).map(formatListItemText)
  ].filter(Boolean).join(" ");
  const cleaned = parts.replace(/\s+/g, " ").trim();
  const region = localizeDisplayTerm(meeting.locationType) || detected.region || "";
  const assetClass = localizeDisplayTerm(meeting.assetClass) || detected.assetClass || "";
  const strategy = localizeDisplayTerm(meeting.strategy) || detected.strategy || "";
  const sector = localizeDisplayTerm(meeting.sector) || detected.sector || "";
  return [
    `${region} ${assetClass} ${strategy} ${sector} 시장 동향`,
    `${region} ${sector} ${strategy} 최근 뉴스 정책 규제`,
    `${cleaned} 리스크 금리 보증 상환 조건`,
    `${meeting.managerName || ""} ${meeting.fundName || ""} 운용사 거래 뉴스`,
    ...specific.specificSearchQueries,
    ...checklist.prioritySearchQueries
  ].map((query) => query.replace(/\s+/g, " ").trim()).filter(Boolean).slice(0, 12);
}

function buildBaselineMarketSearchQueries() {
  const meeting = deriveEffectiveMeetingInfo();
  const im = state.imProcessingResult?.imAnalysis || {};
  const detected = im.autoDetectedFields || {};
  const region = localizeDisplayTerm(meeting.locationType) || detected.region || "주요 투자지역";
  const assetClass = localizeDisplayTerm(meeting.assetClass) || detected.assetClass || "대체투자";
  const strategy = localizeDisplayTerm(meeting.strategy) || detected.strategy || "";
  const sector = localizeDisplayTerm(meeting.sector) || detected.sector || "";
  const regionScope = region === "국내" ? "한국" : region === "해외" ? "북미 미국 글로벌" : region;
  const queries = [
    `${regionScope} ${assetClass} 시장 동향 금리 거래량 2026`,
    `${regionScope} ${sector || assetClass} 시장 전망 리스크 정책 2026`,
    `${regionScope} ${strategy || assetClass} 투자 시장 밸류에이션 유동성 2026`
  ];
  if (/부동산|Real Estate/i.test(assetClass)) {
    queries.push(`${regionScope} 부동산 경기 동향 PF 대출 연체 공실률 거래량 2026`);
  }
  if (/사모투자\(PE\)|Private Equity|PE/i.test(assetClass)) {
    queries.push(`${regionScope} private equity market fundraising exits valuation M&A IPO 2026`);
  }
  if (/사모투자\(PD\)|Private Debt|PD/i.test(assetClass)) {
    queries.push(`${regionScope} private debt direct lending credit spread refinancing default 2026`);
  }
  if (/인프라|Infrastructure/i.test(assetClass)) {
    queries.push(`${regionScope} infrastructure investment market financing regulation yield 2026`);
  }
  return mergeTextLists(queries.map((query) => query.replace(/\s+/g, " ").trim()).filter(Boolean)).slice(0, 6);
}

function buildMarketSearchFocus() {
  const meeting = deriveEffectiveMeetingInfo();
  const im = state.imProcessingResult?.imAnalysis || {};
  const checklist = getAssetClassMarketChecklist();
  const specific = extractSpecificMarketSearchSignals();
  return {
    gpOrManager: meeting.managerName || "",
    fundLoanOrDealName: meeting.fundName || "",
    location: localizeDisplayTerm(meeting.locationType) || "",
    assetClass: localizeDisplayTerm(meeting.assetClass) || "",
    strategy: localizeDisplayTerm(meeting.strategy) || "",
    sector: localizeDisplayTerm(meeting.sector) || "",
    capitalType: localizeDisplayTerm(meeting.capitalType) || "",
    investmentStructure: localizeDisplayTerm(meeting.investmentStructure) || "",
    dealMemo: meeting.keyConcerns || "",
    imDetectedFields: im.autoDetectedFields || {},
    imKeyRisks: asArray(im.keyRisks).slice(0, 4).map(formatListItemText),
    imNumbersToVerify: asArray(im.keyNumbersToVerify).slice(0, 4).map(formatListItemText),
    imVerificationItems: asArray(im.verificationItems).slice(0, 4).map(formatListItemText),
    assetClassMarketChecklist: checklist.baselineChecklist,
    baselineMarketQueries: buildBaselineMarketSearchQueries(),
    specificSignals: specific.specificSignals,
    specificSearchQueries: specific.specificSearchQueries,
    suggestedSearchQueries: buildMarketSearchQueries()
  };
}

function extractSpecificMarketSearchSignals() {
  const corpus = buildMarketSearchCorpus();
  const meeting = deriveEffectiveMeetingInfo();
  const fundOrDealName = meeting.fundName || "";
  const signals = [];
  const addSignal = (type, value, queries) => {
    const cleanValue = String(value || "").replace(/\s+/g, " ").trim();
    if (!cleanValue) return;
    signals.push({
      type,
      value: cleanValue,
      searchPurpose: describeSpecificSearchPurpose(type),
      queries: mergeTextLists(queries.map((query) => query.replace(/\s+/g, " ").trim()).filter(Boolean)).slice(0, 4)
    });
  };

  extractRegexMatches(corpus, /(?:경기도|서울|인천|부산|대구|대전|광주|울산|세종|강원|충청|전라|경상|제주)?\s*[가-힣A-Za-z0-9]+(?:세교|신도시|지구|역세권|블록|BL|구역|산단|택지|도시개발)[가-힣A-Za-z0-9\s()_-]{0,18}(?:PF|프로젝트\s*파이낸싱|브릿지론|본PF)/gi)
    .forEach((phrase) => {
      const location = phrase.replace(/(?:PF|프로젝트\s*파이낸싱|브릿지론|본PF).*/i, "").trim() || phrase;
      addSignal("domestic_project_pf", phrase, [
        `${phrase} 뉴스 인허가 분양 사업지`,
        `${location} 부동산 경기 미분양 공급 거래량`,
        `${location} PF 대출 연체 리파이낸싱 시공사`,
        `${fundOrDealName} ${phrase} 사업주체 시공사 신탁사`
      ]);
    });

  extractRegexMatches(corpus, /\b(?:Dallas|Austin|Houston|New York|Los Angeles|Seattle|Atlanta|Phoenix|Chicago|Boston|San Francisco|San Jose|Miami|Washington(?:\s*DC)?|Toronto|London|Tokyo|Singapore|Sydney|Melbourne)\b[\w\s/-]{0,32}\b(?:multifamily|multi-family|office|logistics|industrial|data center|hotel|student housing|PF|construction loan|bridge loan)\b/gi)
    .forEach((phrase) => addSignal("global_property_pf", phrase, [
      `${phrase} market rent vacancy cap rate transaction`,
      `${phrase} construction loan refinancing debt market`,
      `${phrase} supply pipeline absorption occupancy`,
      `${fundOrDealName} ${phrase} project sponsor news`
    ]));

  extractRegexMatches(corpus, /\b(?:Asia|Asian|North America|US|U\.S\.|Europe|Global|APAC|아시아|북미|미국|유럽)\b[\w\s/-]{0,36}(?:middle market|mid-market|미들마켓|buyout|growth|private equity|PE|사모투자)/gi)
    .forEach((phrase) => addSignal("regional_pe_strategy", phrase, [
      `${phrase} deal activity exits valuation fundraising`,
      `${phrase} M&A IPO financing market`,
      `${phrase} private equity dry powder portfolio exit`,
      `${fundOrDealName} ${phrase} GP track record`
    ]));

  extractNamedPartiesForSearch(corpus).forEach((name) => addSignal("named_party", name, [
    `${name} 최근 뉴스 보도자료 투자 거래`,
    `${name} 소송 제재 신용등급 부실 리스크`,
    `${name} ${fundOrDealName} 관련 뉴스`,
    `${name} track record portfolio transaction`
  ]));

  const uniqueSignals = [];
  const seen = new Set();
  signals.forEach((signal) => {
    const key = normalizeFactKey(`${signal.type} ${signal.value}`);
    if (!key || seen.has(key)) return;
    seen.add(key);
    uniqueSignals.push(signal);
  });
  return {
    specificSignals: uniqueSignals.slice(0, 10),
    specificSearchQueries: mergeTextLists(uniqueSignals.flatMap((signal) => signal.queries)).slice(0, 16)
  };
}

function buildMarketSearchCorpus() {
  const meeting = deriveEffectiveMeetingInfo();
  const im = state.imProcessingResult?.imAnalysis || {};
  return [
    meeting.managerName,
    meeting.fundName,
    meeting.keyConcerns,
    state.imProcessingResult?.textExcerpt,
    JSON.stringify(im)
  ].filter(Boolean).join("\n").replace(/\s+/g, " ").trim();
}

function extractRegexMatches(text, regex) {
  return mergeTextLists((String(text || "").match(regex) || [])
    .map((item) => item.replace(/\s+/g, " ").trim())
    .filter((item) => item.length >= 3 && item.length <= 80));
}

function extractNamedPartiesForSearch(text) {
  const knownParties = [
    "KKR", "Carlyle", "칼라일", "BlackRock", "블랙락", "Ares", "Apollo", "Brookfield", "EQT", "TPG",
    "이지스자산운용", "마스턴투자운용", "코람코", "미래에셋", "삼성SRA", "신한자산운용", "하나자산신탁",
    "포스코이앤씨", "현대건설", "대우건설", "GS건설", "DL이앤씨", "HDC현대산업개발", "롯데건설"
  ];
  const foundKnown = knownParties.filter((name) => new RegExp(escapeRegExp(name), "i").test(text));
  const companyMatches = extractRegexMatches(text, /[가-힣A-Za-z0-9&.\s-]{2,32}(?:자산운용|투자운용|이앤씨|건설|신탁|증권|캐피탈|파트너스|운용|리츠|REITs|Capital|Partners|Management|Asset Management|Construction)/gi);
  return mergeTextLists([...foundKnown, ...companyMatches])
    .map((item) => item.replace(/^(?:및|또는|그리고|with|and)\s+/i, "").trim())
    .filter((item) => item.length >= 2 && item.length <= 50)
    .slice(0, 10);
}

function describeSpecificSearchPurpose(type) {
  if (type === "domestic_project_pf") return "국내 PF 사업지는 사업지 뉴스, 인허가·분양·미분양, 지역 부동산 경기, PF 대출/시공사/신탁사 이슈를 확인합니다.";
  if (type === "global_property_pf") return "해외 부동산 PF는 도시별 임대료, 공실률, 캡레이트, 공급 파이프라인, 건설대출·리파이낸싱 환경을 확인합니다.";
  if (type === "regional_pe_strategy") return "지역·전략형 PE는 거래량, Exit, 밸류에이션, 펀드레이징, 조달 환경을 확인합니다.";
  if (type === "named_party") return "운용사·관계사명은 최근 뉴스, 보도자료, 소송·제재·신용 이슈, 트랙레코드와 거래 관련성을 확인합니다.";
  return "IM에서 확인된 구체 신호를 별도 검색 축으로 확인합니다.";
}

function formatIsoDate(date) {
  return date.toISOString().slice(0, 10);
}

function addMonths(date, months) {
  const next = new Date(date);
  next.setMonth(next.getMonth() + months);
  return next;
}

async function fetchMarketContext() {
  const searchFocus = buildMarketSearchFocus();
  const searchDate = new Date();
  const searchDateText = formatIsoDate(searchDate);
  const recencyCutoffText = formatIsoDate(addMonths(searchDate, -12));
  const prompt = `
기관 LP의 운용사 미팅 준비를 위해 최신 시장, 뉴스, 정책, 규제, 리스크 맥락을 조사해 JSON으로만 답하세요.
사용자의 "딜 / 자산 메모"에는 자산 개요, 지역, 보증 구조, 과거 이슈, 우려사항, 잠정 판단이 섞여 있을 수 있습니다. 이 메모를 검색 키워드와 리스크 분석의 핵심 단서로 사용하세요.
한국어를 기본 언어로 사용하세요. 운용사명, 펀드명, 대출명, 거래명, 계약명, 약어 등 원문 유지가 필요한 고유명사를 제외하고 설명문, 질문, 답변 요지, 리스크, Follow-up은 반드시 한국어로 작성하세요. 입력이 영어여도 보고서 문장과 표 내용은 한국어로 번역·요약하세요.
입력값의 fundName은 펀드명뿐 아니라 직접대출명, PF 대출명, 단일 자산 거래명일 수 있습니다. 시장 맥락과 검색 키워드를 만들 때 펀드형 투자와 직접대출/단일 거래 가능성을 모두 열어두세요.

검색 기준일: ${searchDateText}
최신 뉴스/정책/시장자료 기준: 검색 기준일(${searchDateText}) 기준 최근 1년, 즉 ${recencyCutoffText} 이후 공개된 자료 사용

검색/분석 방식:
- 아래 "검색 초점"의 자산군별 체크리스트를 기본 검색 흐름으로 사용하고, IM/입력값에서 발견되는 특이사항을 추가 검색 단서로 사용하세요.
- 검색 초점의 specificSignals/specificSearchQueries는 IM에서 확인된 구체 사업지·도시·전략·운용사·관계사 신호입니다. 이 항목들은 기본 시장동향과 별도로 반드시 검색하세요.
- 예: "경기도 오산세교 PF"는 오산세교 사업지 뉴스, 인허가, 분양/미분양, 지역 부동산 경기, PF 대출·시공사·신탁사 이슈를 확인합니다.
- 예: "Dallas multifamily PF"는 Dallas multifamily 임대료, 공실률, 캡레이트, 공급 파이프라인, construction loan/refinancing 환경을 확인합니다.
- 예: "Asia middle market PE"는 Asia middle-market PE의 거래량, Exit, 밸류에이션, 펀드레이징, 조달 환경을 확인합니다.
- KKR, Carlyle, BlackRock, 이지스자산운용, 포스코이앤씨 같은 운용사·시공사·관계사명은 최근 뉴스, 공식 보도자료, 소송·제재·신용 이슈, 관련 거래/트랙레코드를 별도 확인하세요.
- 구체적인 지역, 섹터, 전략, 보증, 금리, 상환, 공사/인허가, 정책 키워드 중심으로 찾으세요.
- 일반적인 자산군 설명은 1문장 이하로 줄이고, 이번 건과 직접 연결되는 시장/뉴스/정책/리스크만 남기세요.
- 운용사명, 펀드명, 대출명, 거래명에 대한 개별 뉴스는 검색 결과에서 실제 기사/공시/보도자료가 확인된 경우에만 작성하세요.
- 검색 결과에 없는 운용사명, 펀드명, 대출명, 거래명 관련 기사 제목·날짜·출처를 절대 만들지 마세요.
- 검색 결과가 확인하지 못한 고유명사는 "검색 결과에서 확인되지 않음"으로만 처리하고 directDealEvents/recentEvents에 넣지 마세요.
- directDealEvents/recentEvents에는 운용사/펀드/대출명/거래명/프로젝트명 또는 IM에서 확인된 특정 자산·사업지와 직접 연결되는 기사·공시·보도자료만 넣으세요.
- 일반 시장 동향, 섹터 전망, 금리/환율/거래량/밸류에이션 자료는 keyMarketTrends 또는 riskSignals로 분류하세요. directDealEvents/recentEvents에 섞지 마세요.
- recentEvents, policyRegulatoryNotes, keyMarketTrends, riskSignals, sources에는 반드시 날짜와 출처를 붙이세요.
- 날짜가 없거나 검색 기준일(${searchDateText}) 기준 최근 1년 범위(${recencyCutoffText} 이후) 밖의 자료이면 최신 뉴스/정책/시장자료처럼 쓰지 말고 배열에서 제외하세요.
- "날짜 확인 필요", "출처 확인 필요", "최근 자료 확인 필요" 같은 문구를 출력하지 마세요. 근거가 없으면 해당 배열을 비우고 sourceQuality에 부족하다고 쓰세요.
- 금리 상승, 인플레이션, 거래 위축, 밸류에이션 변화 같은 판단은 어느 시점의 어떤 출처에 근거한 것인지 date/source/fact로 명시하세요.
- 검색 결과가 부족하면 부족하다고 표시하고, 추정으로 채우지 마세요.
- 각 항목에는 실제 검색 결과에서 확인한 source/title/date만 사용하세요. source/title/date 중 하나라도 불확실하면 해당 항목을 제외하세요.
- 토큰을 아끼기 위해 각 배열은 최대 개수만 지키고 짧게 작성하세요.

검색 초점:
${JSON.stringify(searchFocus, null, 2)}

미팅 세팅값:
${JSON.stringify(state.meeting, null, 2)}

IM 분석 요약:
${JSON.stringify(state.imProcessingResult?.imAnalysis || {}, null, 2)}

JSON 스키마:
{
  "summary": "날짜와 출처가 확인된 최신 근거만 사용한 핵심 시장 맥락 3문장 이내",
  "directDealEvents": [{"date": "YYYY-MM-DD 또는 YYYY-MM", "source": "출처명", "title": "기사/공시/보도자료명", "fact": "운용사/펀드/대출/거래/프로젝트에 직접 연결되는 확인된 사실", "relevance": "이번 건과의 직접 관련성"}],
  "keyMarketTrends": [{"date": "YYYY-MM-DD 또는 YYYY-MM", "source": "출처명", "title": "자료/기사명", "fact": "확인된 시장 동향", "relevance": "이번 건과의 관련성"}],
  "recentEvents": [{"date": "YYYY-MM-DD 또는 YYYY-MM", "source": "출처명", "title": "기사/공시/자료명", "fact": "directDealEvents와 같은 직접 관련 이벤트만 작성", "relevance": "이번 건과의 직접 관련성"}],
  "policyRegulatoryNotes": [{"date": "YYYY-MM-DD 또는 YYYY-MM", "source": "출처명", "title": "정책/규제/자료명", "fact": "확인된 정책/규제 내용", "relevance": "이번 건과의 관련성"}],
  "riskSignals": [{"date": "YYYY-MM-DD 또는 YYYY-MM", "source": "출처명", "title": "근거 자료명", "fact": "이번 건에서 확인해야 할 리스크 신호", "relevance": "LP 확인 포인트"}],
  "lpQuestions": ["최대 5개. 위 최신 근거에서 파생된 LP 질문"],
  "followUpRequests": ["최대 3개. 시장/정책 확인용 추가 요청자료"],
  "sources": [{"date": "YYYY-MM-DD 또는 YYYY-MM", "source": "출처명", "title": "출처 제목", "note": "이번 건과의 관련성"}],
  "sourceQuality": "검색 기준일 기준 최근 1년 내 날짜가 확인된 구체적 자료 충분/부족. 부족하면 어떤 축이 부족한지 설명"
}`;
  let detailedContext = null;
  let detailedError = null;
  try {
    const response = await callGeminiWithSearch(prompt);
    const parsed = parseGeminiJson(response.text);
    detailedContext = sanitizeGroundedMarketContext(parsed, response.groundingMetadata);
    if (hasAnyMarketEvidence(detailedContext)) return detailedContext;
  } catch (error) {
    detailedError = error;
  }

  const baselinePrompt = buildBaselineMarketContextPrompt({
    searchFocus,
    searchDateText,
    recencyCutoffText,
    priorFailure: detailedError?.message || detailedContext?.sourceQuality || ""
  });
  const baselineResponse = await callGeminiWithSearch(baselinePrompt);
  const baselineParsed = parseGeminiJson(baselineResponse.text);
  const baselineContext = sanitizeGroundedMarketContext(baselineParsed, baselineResponse.groundingMetadata);
  return mergeMarketContexts(detailedContext, baselineContext);
}

function buildBaselineMarketContextPrompt({ searchFocus, searchDateText, recencyCutoffText, priorFailure }) {
  return `
기관 LP의 운용사 미팅 준비를 위해 기본 시장동향을 반드시 검색해 JSON으로만 답하세요.
이 요청은 운용사/펀드 개별 뉴스 검색이 아니라, 선택된 지역·자산군·섹터의 거시 시장 맥락을 확보하기 위한 필수 검색입니다.

검색 기준일: ${searchDateText}
최신 시장자료 기준: 검색 기준일(${searchDateText}) 기준 최근 1년, 즉 ${recencyCutoffText} 이후 공개된 자료 사용

반드시 수행할 검색:
${JSON.stringify(searchFocus.baselineMarketQueries || [], null, 2)}

검색/작성 원칙:
- Google Search grounding 도구를 반드시 실행하세요.
- 내부 지식만으로 답하지 말고, 검색 결과에 출처 메타데이터가 붙는 공개 자료만 사용하세요.
- 국내 부동산이면 국내 부동산 경기, PF/대출, 거래량, 공실/임대, 금리 영향을 다룹니다.
- 해외/사모투자이면 북미 또는 주요 글로벌 private equity 시장의 fundraising, deal activity, exits, valuation, financing 동향을 다룹니다.
- 직접 관련 뉴스가 없어도 keyMarketTrends, policyRegulatoryNotes, riskSignals에는 지역·자산군 기준의 기본 시장동향을 채우세요.
- specificSignals가 있으면 해당 사업지·도시·전략·관계자명을 기본 시장동향과 연결해 Q&A에 쓸 수 있는 시사점으로 정리하세요.
- 날짜/source/title 중 하나라도 불확실한 항목은 쓰지 마세요.
- 근거 없는 기사 제목이나 수치를 만들지 마세요.
- 각 배열은 최대 4개로 짧게 작성하세요.
${priorFailure ? `\n직전 상세 검색 상태: ${priorFailure}` : ""}

검색 초점:
${JSON.stringify(searchFocus, null, 2)}

JSON 스키마:
{
  "summary": "날짜와 출처가 확인된 기본 시장동향 3문장 이내",
  "directDealEvents": [],
  "keyMarketTrends": [{"date": "YYYY-MM-DD 또는 YYYY-MM", "source": "출처명", "title": "자료/기사명", "fact": "지역·자산군 기준으로 확인된 시장 동향", "relevance": "이번 건의 Q&A에 주는 시사점"}],
  "recentEvents": [],
  "policyRegulatoryNotes": [{"date": "YYYY-MM-DD 또는 YYYY-MM", "source": "출처명", "title": "정책/규제/시장자료명", "fact": "확인된 정책/규제/시장 환경", "relevance": "이번 건과의 관련성"}],
  "riskSignals": [{"date": "YYYY-MM-DD 또는 YYYY-MM", "source": "출처명", "title": "근거 자료명", "fact": "Q&A에 반영해야 할 시장 리스크 신호", "relevance": "LP 확인 포인트"}],
  "lpQuestions": ["기본 시장동향에서 파생되는 GP 질의 최대 5개"],
  "followUpRequests": ["시장/정책 확인용 추가 요청자료 최대 3개"],
  "sources": [{"date": "YYYY-MM-DD 또는 YYYY-MM", "source": "출처명", "title": "출처 제목", "note": "이번 건과의 관련성"}],
  "sourceQuality": "기본 시장동향 검색 근거 충분/부족"
}`;
}

function mergeMarketContexts(primary = null, fallback = null) {
  if (!primary) return fallback || {};
  if (!fallback || !hasAnyMarketEvidence(fallback)) return primary;
  const merged = {
    ...primary,
    summary: primary.summary || fallback.summary || "",
    directDealEvents: mergeMarketItemLists(primary.directDealEvents, fallback.directDealEvents).slice(0, 5),
    keyMarketTrends: mergeMarketItemLists(primary.keyMarketTrends, fallback.keyMarketTrends).slice(0, 5),
    recentEvents: mergeMarketItemLists(primary.recentEvents, fallback.recentEvents).slice(0, 5),
    policyRegulatoryNotes: mergeMarketItemLists(primary.policyRegulatoryNotes, fallback.policyRegulatoryNotes).slice(0, 5),
    riskSignals: mergeMarketItemLists(primary.riskSignals, fallback.riskSignals).slice(0, 5),
    lpQuestions: mergeTextLists([...asArray(primary.lpQuestions), ...asArray(fallback.lpQuestions)]).slice(0, 5),
    followUpRequests: mergeTextLists([...asArray(primary.followUpRequests), ...asArray(fallback.followUpRequests)]).slice(0, 3),
    sources: mergeMarketItemLists(primary.sources, fallback.sources).slice(0, 8),
    sourceQuality: [primary.sourceQuality, "기본 지역·자산군 시장동향 검색으로 보강됨.", fallback.sourceQuality].filter(Boolean).join(" ")
  };
  merged.groundingDiagnostics = {
    ...(primary.groundingDiagnostics || {}),
    baselineGroundingDiagnostics: fallback.groundingDiagnostics || null
  };
  return merged;
}

function mergeMarketItemLists(...lists) {
  const seen = new Set();
  const merged = [];
  lists.flatMap((list) => asArray(list)).forEach((item) => {
    const key = typeof item === "string"
      ? normalizeFactKey(item)
      : normalizeFactKey([item.date, item.source, item.title, item.fact].filter(Boolean).join(" "));
    if (!key || seen.has(key)) return;
    seen.add(key);
    merged.push(item);
  });
  return merged;
}

function buildMarketContextFallback(error) {
  const message = formatMarketGroundingFailureMessage(error);
  return {
    summary: "",
    directDealEvents: [],
    keyMarketTrends: [],
    recentEvents: [],
    policyRegulatoryNotes: [],
    riskSignals: [],
    lpQuestions: [],
    followUpRequests: [],
    sources: [],
    sourceQuality: message,
    groundingDiagnostics: {
      model: lastSearchGroundingModelUsed || SEARCH_GROUNDING_MODEL,
      sourceCount: 0,
      droppedCount: 0,
      failed: true,
      errorMessage: error?.message || message,
      rawMessage: error?.gemini?.rawMessage || ""
    }
  };
}

function formatMarketGroundingFailureMessage(error) {
  const raw = `${error?.message || ""} ${error?.gemini?.rawMessage || ""}`;
  if (/no_grounding_metadata|without grounding chunks|grounding metadata|출처 메타데이터/i.test(raw)) {
    return "검색 그라운딩을 요청했지만 모델 응답에 출처 메타데이터가 붙지 않아 최신 뉴스/시장 근거를 반영하지 않았습니다. 앱은 같은 2.5 모델에서 검색 강제 재시도 후 후순위 모델로 넘어갑니다.";
  }
  if (/429|RESOURCE_EXHAUSTED|quota/i.test(raw)) {
    return "검색 그라운딩 호출이 API 한도 문제로 실패해 최신 뉴스/시장 근거를 반영하지 않았습니다.";
  }
  if (/401|403|UNAUTHENTICATED|PERMISSION_DENIED|API Key/i.test(raw)) {
    return "API Key 권한 문제로 검색 그라운딩이 실패해 최신 뉴스/시장 근거를 반영하지 않았습니다.";
  }
  return "검색 그라운딩이 일시 실패하여 최신 뉴스/시장 근거를 보고서에 반영하지 않았습니다.";
}

function sanitizeGroundedMarketContext(context = {}, groundingMetadata = null) {
  const chunks = asArray(groundingMetadata?.groundingChunks)
    .map((chunk) => chunk.web || chunk)
    .filter(Boolean);
  const groundedText = chunks.map((chunk) => [
    chunk.title,
    chunk.uri,
    chunk.domain,
    chunk.web?.title,
    chunk.web?.uri
  ].filter(Boolean).join(" ")).join("\n");
  const webSearchQueries = asArray(groundingMetadata?.webSearchQueries);
  const groundingAvailable = chunks.length > 0;
  const protectedNames = [
    state.meeting.managerName,
    state.meeting.fundName,
    state.imProcessingResult?.imAnalysis?.autoDetectedFields?.managerName,
    state.imProcessingResult?.imAnalysis?.autoDetectedFields?.fundName,
    state.imProcessingResult?.imAnalysis?.fundSnapshot?.managerName,
    state.imProcessingResult?.imAnalysis?.fundSnapshot?.fundName
  ].map(normalizeProtectedEntityName).filter((name) => name.length >= 3);
  const directEntityNames = mergeTextLists([
    ...protectedNames,
    ...extractDirectEntityNames()
  ].map(normalizeProtectedEntityName)).filter((name) => name.length >= 3);

  const dropped = [];
  const sanitizeItems = (items, fieldName, options = {}) => asArray(items).filter((item) => {
    const check = validateGroundedMarketItem(item, { groundingAvailable, groundedText, protectedNames, directEntityNames, ...options });
    if (!check.ok) dropped.push(`${fieldName}: ${check.reason}`);
    return check.ok;
  }).slice(0, fieldName === "sources" ? 8 : 5);

  const directDealEvents = sanitizeItems([
    ...asArray(context.directDealEvents),
    ...asArray(context.recentEvents)
  ], "직접 관련 뉴스", { directOnly: true });
  const relatedPartySignals = sanitizeItems([
    ...asArray(context.directDealEvents),
    ...asArray(context.recentEvents),
    ...asArray(context.riskSignals)
  ], "관계자 리스크 신호", { relatedPartyOk: true });
  const sanitized = {
    summary: groundingAvailable ? cleanMarketSummary(context.summary, protectedNames, groundedText) : "",
    directDealEvents,
    keyMarketTrends: sanitizeItems(context.keyMarketTrends || context.trends, "시장 동향"),
    recentEvents: directDealEvents,
    policyRegulatoryNotes: sanitizeItems(context.policyRegulatoryNotes || context.newsPolicy, "정책/규제"),
    riskSignals: mergeMarketItemLists(sanitizeItems(context.riskSignals, "리스크 신호"), relatedPartySignals).slice(0, 5),
    lpQuestions: asArray(context.lpQuestions).slice(0, 5),
    followUpRequests: asArray(context.followUpRequests).slice(0, 3),
    sources: sanitizeItems(context.sources, "sources"),
    sourceQuality: context.sourceQuality || ""
  };

  if (!groundingAvailable) {
    sanitized.sourceQuality = "검색 그라운딩 메타데이터가 없어 최신 뉴스/시장 근거를 보고서에 반영하지 않음.";
  } else if (!hasAnyMarketEvidence(sanitized)) {
    sanitized.summary = "";
    sanitized.sourceQuality = "검색은 실행되었으나 검색 기준일 기준 최근 1년 내 날짜·출처 기준을 통과한 시장/뉴스 근거가 부족해 보고서에 제한적으로 반영함.";
  } else {
    sanitized.sourceQuality = buildSourceQualitySummary(sanitized, dropped);
  }
  sanitized.groundingDiagnostics = {
    model: lastSearchGroundingModelUsed || SEARCH_GROUNDING_MODEL,
    webSearchQueries,
    sourceCount: chunks.length,
    droppedCount: dropped.length,
    directEntityNames
  };
  return sanitized;
}

function buildSourceQualitySummary(context = {}, dropped = []) {
  const counts = {
    direct: asArray(context.directDealEvents).length,
    market: asArray(context.keyMarketTrends).length,
    policy: asArray(context.policyRegulatoryNotes).length,
    risk: asArray(context.riskSignals).length,
    sources: asArray(context.sources).length
  };
  const parts = [
    counts.direct ? `직접 관련 뉴스 ${counts.direct}건` : "",
    counts.market ? `시장 참고자료 ${counts.market}건` : "",
    counts.policy ? `정책/규제 ${counts.policy}건` : "",
    counts.risk ? `리스크 신호 ${counts.risk}건` : "",
    counts.sources ? `출처 ${counts.sources}건` : ""
  ].filter(Boolean);
  const base = parts.length
    ? `검색 기준일 기준 최근 1년 내 날짜와 출처가 확인된 자료를 ${parts.join(", ")} 반영했습니다.`
    : "검색 기준일 기준 최근 1년 내 날짜와 출처가 확인된 자료가 제한적입니다.";
  const droppedText = dropped.length
    ? ` 검증 제외 ${dropped.length}건: 출처/날짜/검색근거가 부족한 항목은 환각 방지를 위해 제외했습니다.`
    : "";
  return `${base}${droppedText}`;
}

function hasAnyMarketEvidence(context = {}) {
  return [
    context.directDealEvents,
    context.keyMarketTrends,
    context.policyRegulatoryNotes,
    context.riskSignals,
    context.sources
  ].some((items) => asArray(items).length > 0);
}

function validateGroundedMarketItem(item, { groundingAvailable, groundedText, protectedNames, directEntityNames = [], directOnly = false, relatedPartyOk = false }) {
  if (!groundingAvailable) return { ok: false, reason: "grounding metadata 없음" };
  const text = formatListItemText(item);
  if (!text.trim()) return { ok: false, reason: "빈 항목" };
  if (/날짜\s*확인\s*필요|출처\s*확인\s*필요|최근\s*자료\s*확인\s*필요/i.test(text)) {
    return { ok: false, reason: "불확실 문구 포함" };
  }
  if (item && typeof item === "object" && !Array.isArray(item)) {
    if (!hasReliableMarketDate(item.date || item.publishedAt || item.asOfDate)) {
      return { ok: false, reason: "날짜 없음 또는 최신 기준 미달" };
    }
    const source = String(item.source || "").trim();
    const title = String(item.title || item.headline || item.name || "").trim();
    if (!source || !title) {
      return { ok: false, reason: "출처 또는 제목 없음" };
    }
    if (isLowReliabilityMarketSource(source, title)) {
      return { ok: false, reason: "출처 신뢰도 낮음" };
    }
    if (!hasGroundingOverlap(source, title, groundedText, item)) {
      return { ok: false, reason: "검색 메타데이터와 출처/제목 불일치" };
    }
  }
  const normalizedText = normalizeProtectedEntityName(text);
  const normalizedGroundedText = normalizeProtectedEntityName(groundedText);
  const mentionedProtectedName = protectedNames.find((name) => normalizedText.includes(name));
  if (mentionedProtectedName && !normalizedGroundedText.includes(mentionedProtectedName)) {
    return { ok: false, reason: `고유명사 '${mentionedProtectedName}' 검색근거 없음` };
  }
  if (directOnly && !hasDirectEntitySupport(text, groundedText, directEntityNames)) {
    return { ok: false, reason: "직접 관련 고유명사 근거 없음" };
  }
  if (relatedPartyOk && !isRelatedPartyRiskSignal(text, groundedText, directEntityNames)) {
    return { ok: false, reason: "관계자 리스크 신호 아님" };
  }
  return { ok: true };
}

function isRelatedPartyRiskSignal(itemText, groundedText, directEntityNames = []) {
  const text = normalizeProtectedEntityName(itemText).toLowerCase();
  const grounded = normalizeProtectedEntityName(groundedText).toLowerCase();
  const mentionsParty = directEntityNames.some((name) => {
    const normalized = normalizeProtectedEntityName(name).toLowerCase();
    return normalized.length >= 3 && text.includes(normalized) && grounded.includes(normalized);
  });
  if (!mentionsParty) return false;
  return /경영권|매각|인수|지분|최대주주|주주|핵심인력|인력|조직|담당|이탈|소송|제재|신용|등급|부실|책임준공|시공|재무|유동성|트랙레코드|trackrecord|ownership|sale|acquisition|litigation|credit/i.test(itemText);
}

function extractDirectEntityNames() {
  const im = state.imProcessingResult?.imAnalysis || {};
  const snapshot = im.fundSnapshot || {};
  const detected = im.autoDetectedFields || {};
  const candidates = [
    state.meeting.managerName,
    state.meeting.fundName,
    snapshot.managerName,
    snapshot.fundName,
    snapshot.projectName,
    snapshot.dealName,
    snapshot.loanName,
    detected.managerName,
    detected.fundName,
    detected.projectName,
    detected.dealName,
    detected.loanName,
    findStructuredFact([snapshot, detected, im, state.meeting], ["projectName", "dealName", "loanName", "siteName", "assetName", "사업지", "프로젝트명", "대출명", "거래명", "자산명"])
  ];
  const memoText = [
    state.meeting.fundName,
    state.meeting.keyConcerns,
    state.imProcessingResult?.textExcerpt,
    JSON.stringify(state.imProcessingResult?.imAnalysis || {})
  ].filter(Boolean).join(" ");
  const siteMatches = memoText.match(/[가-힣A-Za-z0-9]+(?:지구|세교|블록|BL|PF|M\d+BL)[가-힣A-Za-z0-9()_-]*/gi) || [];
  return mergeTextLists([...candidates, ...siteMatches].filter(Boolean));
}

function hasDirectEntitySupport(itemText, groundedText, directEntityNames) {
  const item = normalizeProtectedEntityName(itemText).toLowerCase();
  const grounded = normalizeProtectedEntityName(groundedText).toLowerCase();
  return directEntityNames.some((name) => {
    const normalized = normalizeProtectedEntityName(name).toLowerCase();
    return normalized.length >= 3 && item.includes(normalized) && grounded.includes(normalized);
  });
}

function isLowReliabilityMarketSource(source, title = "") {
  const text = `${source} ${title}`;
  return /youtube|youtu\.be|유튜브|tiktok|instagram|facebook|reddit|blog|블로그|카페|forum|커뮤니티|\bdaum\b|\bnaver\b|다음|네이버/i.test(text);
}

function hasGroundingOverlap(source, title, groundedText, item = null) {
  const grounded = normalizeProtectedEntityName(groundedText).toLowerCase();
  const sourceKey = normalizeProtectedEntityName(source).toLowerCase();
  if (sourceKey.length >= 2 && grounded.includes(sourceKey)) return true;
  const sourceTokens = String(source || "")
    .split(/[\s"'“”‘’()[\]{}<>.,;:|/\\·_-]+/)
    .map(normalizeProtectedEntityName)
    .filter((token) => token.length >= 3);
  if (sourceTokens.some((token) => grounded.includes(token.toLowerCase()))) return true;
  const itemText = formatListItemText(item || {});
  const titleTokens = `${title || ""} ${itemText || ""}`
    .split(/[\s"'“”‘’()[\]{}<>.,;:|/\\·_-]+/)
    .map(normalizeProtectedEntityName)
    .filter((token) => token.length >= 4);
  const matches = titleTokens.filter((token) => grounded.includes(token.toLowerCase()));
  return matches.length >= 1;
}

function cleanMarketSummary(summary, protectedNames, groundedText) {
  const text = String(summary || "").trim();
  if (!text) return "";
  const normalizedGrounded = normalizeProtectedEntityName(groundedText);
  const mentionsUnsupportedName = protectedNames.some((name) => normalizeProtectedEntityName(text).includes(name) && !normalizedGrounded.includes(name));
  return mentionsUnsupportedName ? "" : text;
}

function hasReliableMarketDate(value) {
  const text = String(value || "").trim();
  if (!text || /확인\s*필요|unknown|n\/a|미상|불명/i.test(text)) return false;
  const match = text.match(/(\d{4})(?:[-./년]\s*(\d{1,2}))?(?:[-./월]\s*(\d{1,2}))?/);
  if (!match) return false;
  const year = Number(match[1]);
  const month = Number(match[2] || 1);
  const day = Number(match[3] || 1);
  const date = new Date(year, month - 1, day);
  if (Number.isNaN(date.getTime())) return false;
  return date >= addMonths(new Date(), -12);
}

function normalizeProtectedEntityName(value) {
  return String(value || "")
    .replace(/\s+/g, "")
    .replace(/[㈜주식회사펀드대출거래제호]/g, "")
    .trim();
}

function buildBriefPrompt() {
  return `
당신은 기관 LP의 대체투자 운용사 미팅을 준비하는 AI 코파일럿입니다.
아래 입력값을 바탕으로 사전 준비 결과를 JSON으로만 작성하세요.
한국어를 기본 언어로 사용하세요. 운용사명, 펀드명, 대출명, 거래명, 계약명, 약어 등 원문 유지가 필요한 고유명사를 제외하고 설명문, 질문, 답변 요지, 리스크, Follow-up은 반드시 한국어로 작성하세요. 입력이 영어여도 보고서 문장과 표 내용은 한국어로 번역·요약하세요.

중요 원칙:
- IM이 있으면 IM 확인 내용을 우선하되, 사용자가 입력한 세팅값은 보정값으로 사용합니다.
- fundName 필드는 펀드명일 수도 있고 직접대출명, PF 대출명, 단일 자산 거래명일 수도 있습니다. 모든 케이스를 블라인드풀 펀드로 가정하지 말고, 입력/IM이 직접대출, PF 대출, 단일 차주·단일 자산 대출, 단일 트랜치로 보이면 "본 펀드" 대신 "본 건", "본 대출", "본 거래"라고 표현하세요.
- 직접대출 또는 단일 트랜치 Senior Debt로 보이는 경우에는 "Senior Debt 외 다른 트랜치 존재 여부"를 기본 질문으로 만들지 마세요. 대신 단일 선순위 대출인지, 담보권·보증·상환재원·LTV/DSCR·EOD·약정 조건·대주간 권리관계·후순위/메자닌 동반 여부가 실제로 필요한 경우만 확인 질문으로 작성하세요.
- "딜 / 자산 메모"는 단순 우려사항이 아니라 사용자가 알고 있는 자산 개요, 지역, 보증 구조, 과거 이슈, 잠정 판단을 담는 핵심 맥락입니다.
- 이 메모의 표현을 그대로 믿기보다, 시장/뉴스/정책/리스크 맥락과 대조해 확인 필요 항목과 질문으로 전환합니다.
- 시장/뉴스/정책 맥락은 일반론보다 이번 건의 지역, 자산군, 섹터, 전략, 보증, 금리, 상환 구조와 직접 연결되는 내용만 우선 사용합니다.
- 시장/뉴스/정책 항목은 state.marketContext의 date/source/title/fact/relevance만 사용합니다. state.marketContext에 없는 기사명, 날짜, 출처, 운용사/펀드 관련 뉴스를 새로 만들지 마세요.
- state.marketContext에서 검색 근거가 부족하다고 표시된 경우, marketContext는 비워두거나 sourceQuality에 부족하다고만 쓰세요.
- "날짜 확인 필요", "출처 확인 필요" 같은 문구를 만들지 마세요. 날짜/출처가 없는 뉴스는 최신 근거처럼 쓰지 말고 확인 필요 항목으로 돌리세요.
- IM 내용과 사용자가 입력한 값이 충돌하면 conflicts 또는 verificationItems에 표시합니다.
- Q&A는 최소 5개, 최대 10개로 작성합니다.
- Q&A의 중심은 딜 판단입니다. 투자 thesis, 구조, 상환/Exit, 담보/보증, 현금흐름, track record, alignment, 주요 리스크와 mitigation을 우선 질문하세요.
- IM 내부 숫자/문구 오류 확인 질문은 보조 질문입니다. 같은 항목의 숫자, 금리, 수익률, LTV, DSCR, 기간, 금액, 약정 조건이 서로 다르게 적혀 있는 경우에만 expectedQaList에 포함하고, 전체 Q&A 중 최대 2개까지만 포함하세요.
- 운용사명, 사업주체, 주주, 신탁사, 시공사처럼 역할이 여러 개인 당사자가 섞여 있을 때는 "확인됩니다", "불일치합니다"처럼 단정하지 말고 "본 건 관계자의 역할, 법적 지위, 책임 범위 및 계약상 관계 확인 필요"처럼 GP가 답할 수 있는 질문형으로 낮춰 쓰세요. "사용자 입력", "세팅값", "내부 메모" 같은 LP 내부 출처 표현은 Q&A 질문에 쓰지 마세요.
- 시장검색 또는 IM에서 직접 확인되지 않은 운용사/주관사/거래당사자 관계를 새로 만들어 단정하지 마세요.
- Alignment/IM 정합성 질문은 딜 판단에 중요한 경우에만 넣고, 전체 Q&A 중 1개를 넘기지 마세요.
- 예: 주택도시기금 금리가 한 곳에는 2.8%, 다른 곳에는 2.6%로 보이면 "적용 기준일, 적용 구간, 산식 또는 오기 여부"를 묻는 Q&A를 만드세요.
- 정보가 부족하면 지어내지 말고 "확인 필요"로 표시합니다.
- 질문은 실제 미팅에서 바로 읽을 수 있는 수준으로 구체적으로 작성합니다.

세팅값:
${JSON.stringify(state.meeting, null, 2)}

IM 처리 결과:
${JSON.stringify(state.imProcessingResult || null, null, 2)}

시장/뉴스/정책/리스크 맥락:
${JSON.stringify(state.marketContext || null, null, 2)}

응답 JSON 스키마:
{
  "fundSnapshot": {
    "managerName": "운용사명 또는 확인 필요",
    "fundName": "펀드명, 대출명, 거래명 또는 확인 필요",
    "strategy": "전략",
    "region": "지역",
    "assetClass": "자산군",
    "capitalStructure": "투자구조",
    "targetSize": "펀드 규모 또는 확인 필요",
    "loanSize": "대출 규모 또는 확인 필요",
    "investmentPeriod": "투자 기간 또는 확인 필요",
    "loanMaturity": "대출만기 또는 확인 필요",
    "targetReturn": "목표 수익률 또는 확인 필요",
    "loanRate": "대출금리 또는 확인 필요",
    "commitmentAmount": "당사 검토 약정액. 확인되지 않으면 확인 필요",
    "keyNumbers": ["핵심 숫자"]
  },
  "classificationSummary": {
    "confirmedFromIm": ["IM에서 확인된 정보"],
    "providedByUser": ["사용자가 입력한 정보"],
    "conflicts": ["충돌 정보"],
    "needsVerification": ["확인 필요 정보"]
  },
  "marketContext": {
    "summary": "날짜와 출처가 확인된 시장 맥락 요약",
    "trends": [{"date": "YYYY-MM-DD 또는 YYYY-MM", "source": "출처명", "title": "자료명", "fact": "시장 동향", "relevance": "이번 건과의 관련성"}],
    "newsPolicy": [{"date": "YYYY-MM-DD 또는 YYYY-MM", "source": "출처명", "title": "뉴스/정책명", "fact": "뉴스/정책/규제 내용", "relevance": "이번 건과의 관련성"}],
    "riskSignals": [{"date": "YYYY-MM-DD 또는 YYYY-MM", "source": "출처명", "title": "근거 자료명", "fact": "리스크 신호", "relevance": "LP 확인 포인트"}],
    "sources": [{"date": "YYYY-MM-DD 또는 YYYY-MM", "source": "출처명", "title": "출처 제목"}]
  },
  "expectedQaList": [
    {
      "category": "투자 thesis/구조/상환·Exit/담보·보증/현금흐름/트랙레코드/Alignment/리스크/IM 정합성 중 하나",
      "importance": "High/Medium/Low",
      "question": "질문",
      "rationale": "왜 물어봐야 하는지",
      "source": "IM/사용자 입력/시장 검색/템플릿"
    }
  ],
  "keyRisks": ["핵심 리스크"],
  "redFlags": ["Red Flag 체크리스트"],
  "expectedFollowUpRequests": ["예상 추가 요청자료"],
  "verificationItems": ["원문 대조 또는 추가 확인 필요 항목"]
}`;
}

function buildImAnalysisPrompt(text, policy, fileMeta) {
  return `
기관 LP가 운용사 미팅 전에 IM을 빠르게 파악하려고 합니다.
아래 IM 텍스트를 읽고 JSON으로만 요약하세요.
한국어를 기본 언어로 사용하세요. 운용사명, 펀드명, 대출명, 거래명, 계약명, 약어 등 원문 유지가 필요한 고유명사를 제외하고 설명문, 질문, 답변 요지, 리스크, Follow-up은 반드시 한국어로 작성하세요. 입력이 영어여도 보고서 문장과 표 내용은 한국어로 번역·요약하세요.
IM이 펀드형 투자자료인지, 직접대출/프로젝트 파이낸싱/단일 자산 거래자료인지 먼저 구분하세요. 직접대출 또는 단일 트랜치 대출이면 fundName에는 대출명 또는 거래명을 넣고, 질문은 펀드 운용전략보다 차주, 담보, 보증, 상환재원, 선순위성, 약정 조건 중심으로 정리하세요.
검토 유형은 fund 또는 loan 중 하나로 판단하세요. PF 대출, 담보대출, 직접대출, 대출확약, 차주/대주 구조이면 loan입니다. 블라인드펀드, 프로젝트펀드 출자, LP commitment 중심이면 fund입니다.

분석 모드: ${policy.mode}
파일 정보: ${JSON.stringify(fileMeta)}

Focused skim mode일 때는 다음 항목을 최우선으로 봅니다:
자산분류, 섹터, 지역, 딜 개요, 관계 플레이어, 주요 리스크 사항, 투자구조, 핵심 숫자.
부동산/인프라/복합자산의 sector는 하나로 압축하지 말고 IM에 나온 주요 용도와 섹터를 모두 보존하세요. 예: 공동주택, 오피스텔, 판매시설, 근린생활시설.
PE/PD의 sector도 여러 개가 있으면 모두 보존하세요. 예: 테크/소프트웨어, 헬스케어/바이오, 소비재/이커머스.
숫자/조건 정합성도 자연스럽게 점검하세요. 같은 항목의 금리, 수익률, LTV, DSCR, 대출기간, 상환조건, 보증조건, 수수료, 금액이 페이지별로 다르면 keyNumbersToVerify, conflicts, verificationItems, mustAskQuestionsFromIm에 반영하세요.
당사 검토 약정액은 총 펀드규모/총 대출규모가 아닙니다. "당사", "본 LP", "검토 약정", "출자 검토액" 등으로 명시된 금액이 없으면 commitmentAmount는 빈 값으로 두세요.

IM 텍스트:
${text}

응답 JSON 스키마:
{
  "fundSnapshot": {
    "managerName": "",
    "fundName": "",
    "targetSize": "펀드 규모",
    "loanSize": "대출 규모",
    "investmentPeriod": "투자 기간",
    "loanMaturity": "대출만기",
    "targetReturn": "목표 수익률",
    "loanRate": "대출금리",
    "commitmentAmount": "당사/본 LP의 검토 약정액. 총 펀드규모나 총 대출규모와 혼동하지 말고, 명시되지 않았으면 빈 값"
  },
  "strategySummary": "전략 요약",
  "keyInvestmentMerits": ["투자 포인트"],
  "keyRisks": ["리스크"],
  "keyNumbersToVerify": ["확인해야 할 숫자"],
  "trackRecordCheckpoints": ["트랙레코드 확인사항"],
  "mustAskQuestionsFromIm": ["IM 기반 필수 질문"],
  "followUpRequestsFromIm": ["IM 기반 추가 요청자료"],
  "autoDetectedFields": {
    "managerName": "",
    "fundName": "",
    "assetClass": "",
    "region": "",
    "strategy": "",
    "sector": "",
    "capitalType": "",
    "dealType": "fund 또는 loan",
    "investmentStructure": ""
  },
  "conflicts": [],
  "verificationItems": []
}`;
}

function buildImVisionPrompt(policy, fileMeta) {
  return `
기관 LP가 운용사 미팅 전에 이미지/PDF IM을 읽으려고 합니다.
첨부 이미지를 보고 JSON으로만 답하세요.
한국어를 기본 언어로 사용하세요. 운용사명, 펀드명, 대출명, 거래명, 계약명, 약어 등 원문 유지가 필요한 고유명사를 제외하고 설명문, 질문, 답변 요지, 리스크, Follow-up은 반드시 한국어로 작성하세요. 입력이 영어여도 보고서 문장과 표 내용은 한국어로 번역·요약하세요.
IM이 펀드형 투자자료인지, 직접대출/프로젝트 파이낸싱/단일 자산 거래자료인지 먼저 구분하세요. 직접대출 또는 단일 트랜치 대출이면 fundName에는 대출명 또는 거래명을 넣고, 질문은 펀드 운용전략보다 차주, 담보, 보증, 상환재원, 선순위성, 약정 조건 중심으로 정리하세요.
검토 유형은 fund 또는 loan 중 하나로 판단하세요. PF 대출, 담보대출, 직접대출, 대출확약, 차주/대주 구조이면 loan입니다. 블라인드펀드, 프로젝트펀드 출자, LP commitment 중심이면 fund입니다.

분석 모드: ${policy.mode}
파일 정보: ${JSON.stringify(fileMeta)}

우선 추출 항목:
자산분류, 섹터, 지역, 딜 개요, 관계 플레이어, 주요 리스크 사항, 투자구조, 핵심 숫자, 트랙레코드, 수수료/조건.
sector는 하나로 압축하지 말고 이미지/표에 나온 주요 용도와 섹터를 모두 보존하세요.
숫자/조건 정합성도 자연스럽게 점검하세요. 같은 항목의 금리, 수익률, LTV, DSCR, 대출기간, 상환조건, 보증조건, 수수료, 금액이 페이지별로 다르면 keyNumbersToVerify, conflicts, verificationItems, mustAskQuestionsFromIm에 반영하세요.
당사 검토 약정액은 총 펀드규모/총 대출규모가 아닙니다. "당사", "본 LP", "검토 약정", "출자 검토액" 등으로 명시된 금액이 없으면 commitmentAmount는 빈 값으로 두세요.

응답 JSON 스키마:
{
  "fundSnapshot": {
    "managerName": "",
    "fundName": "",
    "targetSize": "",
    "loanSize": "",
    "investmentPeriod": "",
    "loanMaturity": "",
    "targetReturn": "",
    "loanRate": "",
    "commitmentAmount": ""
  },
  "strategySummary": "",
  "keyInvestmentMerits": [],
  "keyRisks": [],
  "keyNumbersToVerify": [],
  "trackRecordCheckpoints": [],
  "mustAskQuestionsFromIm": [],
  "followUpRequestsFromIm": [],
  "autoDetectedFields": {},
  "conflicts": [],
  "verificationItems": []
}`;
}

async function generateReport() {
  if (isBusy) {
    toast("이미 AI 작업이 진행 중입니다.");
    return;
  }
  try {
    syncFieldsToState();
    requireApiKey();
    if (!state.preMeetingBrief) throw new Error("먼저 사전 브리프를 생성해주세요.");
    const hasMeetingInputs = state.questionRecords.some((q) => hasAnswer(q.answerRecord)) || state.meetingNotes.trim() || state.transcript.trim();
    if (!hasMeetingInputs) throw new Error("미팅 답변, 자유 메모, Transcript 중 하나 이상을 입력해주세요.");

    switchPhase("report");
    showLoader("최종 보고서를 생성하고 있습니다.", "사전 분석과 미팅 중 기록을 논점별로 종합합니다.");
    const prompt = buildReportPrompt();
    const text = await callGeminiText(prompt, { json: true, temperature: 0.2 });
    state.postMeetingMemo = parseGeminiJson(text);
    saveState();
    saveMeetingToLibrary();
    renderReport();
    renderStatus();
    toast("최종 보고서를 생성했습니다.");
  } catch (error) {
    showError(error);
  } finally {
    hideLoader();
  }
}

function buildReportPrompt() {
  const tone = getReportToneConfig(state.reportTone);
  const effectiveMeeting = deriveEffectiveMeetingInfo();
  const reviewType = resolveReviewType(effectiveMeeting, {}, state.imProcessingResult?.imAnalysis || {});
  const metricLabels = buildOverviewMetricRows({}, reviewType).map(([label]) => label);
  return `
[중요 작성 지시]
- 01 요약의 "심사역 의견"과 03 심사역 검토 의견의 "검토 의견 요약"은 반드시 이 딜이 무엇인지부터 알 수 있어야 합니다.
- reviewConclusion에는 딜명/대출명/펀드명, 운용사, 지역·자산군·전략·구조 중 확인 가능한 정보를 넣어 "본 건은 어떤 성격의 검토 건인지"를 먼저 설명한 뒤, 추가 검토/보류/드랍/조건부 진행 판단을 이어서 작성하세요.
- 03 심사역 검토 의견의 각 값에는 "결론:", "투자매력:", "핵심리스크:", "추가확인사항:" 같은 라벨을 쓰지 말고 본문 문장만 작성하세요.
- 정보가 일부 부족해도 "확인 필요"만 쓰지 말고, 확인된 딜 정체성 + 부족한 확인사항을 함께 적으세요.

당신은 기관 LP의 운용사 미팅 후속 정리를 돕는 AI 코파일럿입니다.
사전 분석, 질문별 답변, 내부 메모, 자유 메모, Transcript를 종합해 JSON으로만 답하세요.
확인되지 않은 내용은 단정하지 말고 "확인 필요"로 표시하세요.
한국어를 기본 언어로 사용하세요. 운용사명, 펀드명, 대출명, 거래명, 계약명, 약어 등 원문 유지가 필요한 고유명사를 제외하고 설명문, 질문, 답변 요지, 리스크, Follow-up은 반드시 한국어로 작성하세요. 입력이 영어여도 보고서 문장과 표 내용은 한국어로 번역·요약하세요.
fundName은 펀드명뿐 아니라 직접대출명, PF 대출명, 단일 자산 거래명일 수 있습니다. 직접대출 또는 단일 트랜치 대출이면 보고서 문장에서도 "본 펀드"보다 "본 건", "본 대출", "본 거래"를 우선 사용하세요.

보고서 톤:
- 선택값: ${tone.label}
- 검토의견 기본값: ${tone.reviewOpinion}
- 작성 지침: ${tone.promptInstruction}
- 최종 DOCX는 DOCX_FINAL_WANTED.docx 양식에 들어갑니다. 아래 6개 섹션에 맞는 내용만 생성하세요.
  01 요약: 펀드명/대출명/거래명, 운용사, 미팅일자, 검토 결과, 심사역 의견에 들어갈 핵심 문장
  02 미팅 개요: 미팅 정보와 지역/자산군/전략/구조를 간결히 정리
  03 심사역 검토 의견: 반드시 아래 4개 bullet을 이 순서로 작성
    1. 결론: 추가 검토 / 보류 / 드랍 / 조건부 진행 중 하나가 바로 보이도록 작성
    2. 투자 매력 / Why This Deal: IRR, 담보, 입지, GP 역량, 구조적 보호장치, 시장 타이밍 등 중 핵심만 작성
    3. 핵심 리스크 / Key Risks: 딜을 깨뜨릴 수 있는 가장 중요한 위험 2~3개만 작성
    4. 추가 확인사항 / Next DD Items: 미팅 후 바로 받아야 할 자료와 확인 액션 작성
  04 핵심 DDQ: 최대 5개 질문과 GP 답변 요지
  05 주요 예상 리스크: 최대 6개 리스크, 구분/내용/심각도/당사 View로 나뉘어 들어갈 수 있게 작성
  06 향후 Follow-up: 1~3개 핵심 후속 확인사항만 작성. 항목 내용만 쓰고 담당/기한/상태는 생성하지 마세요.
- 한 줄 결론과 내부 보고용 간이보고서도 이 DOCX 양식에 맞춰 작성하세요.
- 단, 긍정적 톤이어도 원문 대조 필요 항목과 추가 DD 필요 사항은 반드시 남기세요.
- 모든 문장은 보고서형 음슴체로 작성하세요. 예: "~입니다", "~합니다", "~됩니다" 금지. "~임", "~함", "~됨", "~필요", "~검토 필요"처럼 끝내세요.
- DDQ 질문과 GP 답변 요지도 한국어로 작성하세요. 영어 원문 답변이 있으면 그대로 복사하지 말고 한국어로 요약하세요.
- 내부 보고용 간이보고서는 5문장 이내로만 작성하세요.
- internalReportSummary는 DOCX 03 "심사역 검토 의견" 섹션에 직접 들어갑니다. 반드시 결론 → 투자 매력 → 핵심 리스크 → 추가 확인사항 순서로 작성하세요.
  - reviewConclusion: 이 건을 추가 검토할지, 보류할지, 드랍할지 한 문장으로 작성. 값 안에 "결론:" 같은 라벨 금지
  - whyThisDeal: 왜 볼 만한 딜인지 핵심 투자 매력만 작성. 값 안에 "투자 매력:" 같은 라벨 금지
  - keyRisks: 이 딜을 깨뜨릴 수 있는 핵심 리스크 2~3개만 작성. 값 안에 "핵심 리스크:" 같은 라벨 금지
  - nextDdItems: 그래서 바로 무엇을 더 받아야 하는지 작성. 값 안에 "추가 확인사항:" 같은 라벨 금지
  - investmentMemo: 위 4개 항목을 한 번 더 자연어로 요약하되, 네 항목의 순서를 유지
  - DOCX 03 섹션에는 항목명 라벨을 표시하지 않고 각 bullet 본문만 들어갑니다.
  - 중립적 리뷰: 장점, 확인 필요 항목, 주요 리스크를 균형 있게 작성
  - 긍정적 검토: 투자 적합성과 긍정 요소를 먼저 쓰되 조건부 확인사항을 명확히 작성
  - 보수적 리스크 강조: downside, 원문 대조 필요 항목, 미해결 쟁점을 앞세워 작성
- 불필요한 일반론, 긴 배경 설명, 중복 문장은 제외하세요.
- 미팅 개요의 아래 4개 항목은 IM, 사전 브리프, 질문별 답변, 자유 메모, Transcript에서 확인된 값이 있으면 반드시 추출하세요. 확인되지 않은 항목은 지어내지 말고 "확인 필요"로 두세요.
  - ${metricLabels[0]}
  - ${metricLabels[1]}
  - ${metricLabels[2]}
  - 당사 검토 약정액
검토 유형: ${reviewType || "자동 판단"}
검토 유형이 loan이면 펀드 규모/투자 기간/목표 수익률은 쓰지 말고 대출 규모/대출만기/대출금리만 쓰세요.
검토 유형이 fund이면 대출 규모/대출만기/대출금리는 쓰지 말고 펀드 규모/투자 기간/목표 수익률만 쓰세요.

세팅값:
${JSON.stringify(state.meeting, null, 2)}

사전 브리프:
${JSON.stringify(state.preMeetingBrief, null, 2)}

시장 맥락:
${JSON.stringify(state.marketContext, null, 2)}

질문별 미팅 기록:
${JSON.stringify(getQuestionRecordsForReport(), null, 2)}

자유 메모:
${state.meetingNotes}

Transcript:
${state.transcript}

응답 JSON 스키마:
{
  "issueBasedMeetingNotes": [{"issue": "논점", "summary": "회의록 요약", "evidence": "근거/발언"}],
  "unresolvedIssues": ["미해결 쟁점"],
  "followUpRequestList": ["1~3개. 핵심 후속 요청자료 또는 액션. 항목 내용만 작성하고 담당/기한/상태는 제외"],
  "followUpEmailDraft": "운용사에게 보낼 이메일 초안",
  "internalReportSummary": {
    "oneLineView": "한 줄 검토 의견",
    "reviewConclusion": "추가 검토/보류/드랍/조건부 진행 중 하나가 바로 보이는 본문만 작성. '결론:' 라벨 금지",
    "whyThisDeal": "IRR, 담보, 입지, GP 역량, 구조적 보호장치, 시장 타이밍 등 핵심 투자 매력 본문만 작성. '투자 매력:' 라벨 금지",
    "keyRisks": "가장 중요한 위험 2~3개 본문만 작성. '핵심 리스크:' 라벨 금지",
    "nextDdItems": "미팅 후 바로 받아야 할 자료와 확인 액션 본문만 작성. '추가 확인사항:' 라벨 금지",
    "investmentMemo": "심사역 검토 의견. 결론 → 투자 매력 → 핵심 리스크 → 추가 확인사항 순서로 4문장 내외",
    "riskView": "리스크 관점"
  },
  "meetingOverview": {
    "reviewType": "${reviewType || "fund 또는 loan"}",
    "location": "장소. 확인되지 않으면 확인 필요",
    "regionAssetClass": "지역 / 자산군 / 섹터 / 전략 / Equity·Debt 요약",
    "sizeOrLoanAmount": "${metricLabels[0]}. IM 또는 Q&A에서 확인된 값",
    "periodOrMaturity": "${metricLabels[1]}. IM 또는 Q&A에서 확인된 값",
    "returnOrLoanRate": "${metricLabels[2]}. IM 또는 Q&A에서 확인된 값",
    "commitmentAmount": "당사 검토 약정액. 확인되지 않으면 확인 필요"
  },
  "nextActionItems": ["다음 액션"],
  "qualityChecks": ["품질 점검"],
  "sourceVerificationItems": ["원문 대조 필요 항목"]
}`;
}

async function callGeminiText(prompt, options = {}) {
  const body = {
    contents: [{ role: "user", parts: [{ text: prompt }] }],
    generationConfig: {
      temperature: options.temperature ?? 0.2
    }
  };
  if (options.json) body.generationConfig.responseMimeType = "application/json";
  return callGeminiGenerate(body);
}

async function callGeminiVision(files, prompt) {
  const parts = [{ text: prompt }];
  files.forEach((file) => {
    parts.push({ inline_data: { mime_type: file.mimeType, data: file.data } });
  });
  return callGeminiGenerate({
    contents: [{ role: "user", parts }],
    generationConfig: { temperature: 0.15, responseMimeType: "application/json" }
  });
}

async function callGeminiWithSearch(prompt) {
  requireApiKey();
  const candidates = buildSearchGroundingModelCandidates();
  const failures = [];

  for (const model of candidates) {
    try {
      let response = await callGeminiModel(model, buildGeminiSearchRequestBody(prompt), { response: "full" });
      if (!hasUsableGroundingMetadata(response.groundingMetadata)) {
        failures.push({
          model,
          reason: "no_grounding_metadata",
          rawMessage: "The first search-grounding response did not include grounding chunks. Retrying with an explicit search instruction."
        });
        response = await callGeminiModel(model, buildGeminiSearchRequestBody(buildGroundingRequiredPrompt(prompt)), { response: "full" });
      }
      if (hasUsableGroundingMetadata(response.groundingMetadata)) {
        lastSearchGroundingModelUsed = model;
        return response;
      }
      failures.push({
        model,
        reason: "no_grounding_metadata",
        rawMessage: "The model returned a response without grounding chunks."
      });
    } catch (error) {
      const info = getGeminiErrorInfoFromError(error);
      failures.push({
        model,
        reason: info.reason || "unknown",
        status: info.status || null,
        statusText: info.statusText || "",
        rawMessage: info.rawMessage || error?.message || ""
      });
      if (!shouldTryNextSearchGroundingModel(error)) throw error;
    }
  }

  const info = buildSearchGroundingFailureInfo(failures);
  throw createGeminiApiError(info.message, info);
}

function buildGeminiSearchRequestBody(prompt) {
  return {
    contents: [{ role: "user", parts: [{ text: prompt }] }],
    tools: [{ google_search: {} }]
  };
}

function buildGroundingRequiredPrompt(prompt) {
  return `
반드시 Google Search grounding 도구를 실행한 뒤 답하세요.
- 내부 지식이나 추정만으로 답하지 마세요.
- 최근 뉴스/시장/정책/리스크 자료를 공개 웹에서 검색하고, 검색으로 확인된 출처가 있는 항목만 JSON에 넣으세요.
- 검색 결과가 부족하면 배열을 비우되, 응답 자체는 grounding metadata가 붙도록 검색 쿼리를 실행하세요.
- 검색에 사용할 우선 쿼리는 아래 요청의 suggestedSearchQueries, 지역, 섹터, 운용사명, 펀드명, 대출명, 프로젝트명입니다.

${prompt}`;
}

function buildSearchGroundingModelCandidates() {
  return [
    SEARCH_GROUNDING_MODEL,
    ...SEARCH_GROUNDING_MODEL_CANDIDATES
  ]
    .map(normalizeGeminiModelName)
    .filter(Boolean)
    .filter(isLikelySearchGroundingModel)
    .filter((model, index, models) => models.indexOf(model) === index);
}

function isLikelySearchGroundingModel(model) {
  const name = String(model || "").toLowerCase();
  if (!name.startsWith("gemini-")) return false;
  if (/image|tts|robotics|computer-use|customtools/.test(name)) return false;
  return true;
}

function hasUsableGroundingMetadata(metadata) {
  return asArray(metadata?.groundingChunks).length > 0;
}

function shouldTryNextSearchGroundingModel(error) {
  const info = getGeminiErrorInfoFromError(error);
  const raw = `${info.status || ""} ${info.statusText || ""} ${info.rawMessage || ""} ${info.message || ""}`;
  if (["unavailable", "rate_limit", "model_not_found", "model_not_listed", "unsupported_action"].includes(info.reason)) return true;
  return /503|UNAVAILABLE|429|RESOURCE_EXHAUSTED|404|NOT_FOUND|not found|not supported|unsupported|google_search|tool/i.test(raw);
}

function buildSearchGroundingFailureInfo(failures = []) {
  const details = failures.map((failure) => {
    const code = [failure.status, failure.statusText || failure.reason].filter(Boolean).join(" ");
    const raw = String(failure.rawMessage || "").replace(/\s+/g, " ").trim();
    return `${failure.model}${code ? ` (${code})` : ""}${raw ? `: ${raw.slice(0, 180)}` : ""}`;
  }).join(" | ");
  return {
    model: failures.map((failure) => failure.model).filter(Boolean).join(", "),
    status: null,
    statusText: "SEARCH_GROUNDING_FAILED",
    rawMessage: details,
    reason: "search_grounding_failed",
    message: "검색 그라운딩을 실행했지만 출처 메타데이터를 확보하지 못했습니다. 2.5 모델은 검색 강제 재시도까지 수행했고, 후순위 모델도 실패했습니다."
  };
}

async function callGeminiGenerate(body, options = {}) {
  requireApiKey();
  const models = options.noFallback
    ? [normalizeTextGenerationModel(options.model || runtimeConfig.model || DEFAULT_MODEL)]
    : buildGenerationModelCandidates(options.model || runtimeConfig.model || DEFAULT_MODEL);
  const failures = [];

  for (const model of models) {
    try {
      return await callGeminiModel(model, body, options);
    } catch (error) {
      const info = getGeminiErrorInfoFromError(error);
      failures.push(info);
      if (options.noFallback || !shouldTryNextTextGenerationModel(model, error)) throw error;
      logGeminiDiagnostic({
        model,
        status: info.status || null,
        statusText: info.statusText || "",
        reason: "fallback_for_this_request",
        rawMessage: info.rawMessage || info.message || ""
      });
    }
  }

  const info = buildGenerationFailureInfo(failures);
  throw createGeminiApiError(info.message, info);
}

function buildGenerationModelCandidates(model) {
  return [
    normalizeTextGenerationModel(model),
    ...GENERATION_FALLBACK_MODEL_CANDIDATES
  ]
    .map(normalizeGeminiModelName)
    .filter(Boolean)
    .filter(isLikelyTextGenerationModel)
    .filter(isAllowedTextGenerationModel)
    .filter((item, index, items) => items.indexOf(item) === index);
}

function isAllowedTextGenerationModel(model) {
  return TEXT_GENERATION_ALLOWED_MODELS.includes(normalizeGeminiModelName(model));
}

function isLikelyTextGenerationModel(model) {
  const name = String(model || "").toLowerCase();
  if (!name.startsWith("gemini-")) return false;
  if (/image|tts|robotics|computer-use|customtools/.test(name)) return false;
  return true;
}

function buildGenerationFailureInfo(failures = []) {
  const details = failures.map((info) => {
    const code = [info.status, info.statusText || info.reason].filter(Boolean).join(" ");
    const raw = String(info.rawMessage || info.message || "").replace(/\s+/g, " ").trim();
    return `${info.model}${code ? ` (${code})` : ""}${raw ? `: ${raw.slice(0, 160)}` : ""}`;
  }).join(" | ");
  return {
    model: failures.map((info) => info.model).filter(Boolean).join(", "),
    status: null,
    statusText: "GENERATION_FAILED",
    rawMessage: details,
    reason: "generation_failed",
    message: `Gemini 텍스트 생성 후보가 모두 실패했습니다. ${details}`
  };
}

async function callGeminiModel(model, body, options = {}) {
  lastGeminiModelUsed = model;
  const url = `${GEMINI_API_BASE}/${encodeURIComponent(model)}:generateContent`;
  try {
    const apiKey = getGeminiApiKeyForRequest();
    const requestBody = sanitizeGeminiRequestBodyForModel(body, model);
    const response = await fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-goog-api-key": apiKey
      },
      body: JSON.stringify(requestBody)
    });
    if (!response.ok) {
      const errorText = await response.text().catch(() => "");
      const info = getGeminiErrorInfo(errorText, response.status, model);
      logGeminiDiagnostic(info);
      throw createGeminiApiError(info.message, info);
    }
    const data = await response.json();
    const text = data.candidates?.[0]?.content?.parts?.map((part) => part.text || "").join("\n").trim();
    if (!text) throw new Error(LIMIT_MESSAGE);
    if (options.response === "full") {
      return {
        text,
        raw: data,
        groundingMetadata: data.candidates?.[0]?.groundingMetadata || null
      };
    }
    return text;
  } catch (error) {
    throw error instanceof Error ? error : new Error(LIMIT_MESSAGE);
  }
}

function sanitizeGeminiRequestBodyForModel(body, model) {
  const next = JSON.parse(JSON.stringify(body || {}));
  if (/^gemini-3/i.test(model) && next.generationConfig) {
    delete next.generationConfig.temperature;
    delete next.generationConfig.topP;
    delete next.generationConfig.topK;
    if (!Object.keys(next.generationConfig).length) delete next.generationConfig;
  }
  return next;
}

function shouldTryNextTextGenerationModel(model, error = "") {
  if (model === FALLBACK_MODEL) return false;
  const message = typeof error === "string" ? error : error?.message || "";
  const info = typeof error === "string" ? null : error?.gemini;
  if (info?.status === 429 || info?.status === 503) return true;
  if (["rate_limit", "unavailable"].includes(info?.reason)) return true;
  return /503|UNAVAILABLE|high demand|quota|429|RESOURCE_EXHAUSTED|GenerateContentRequestsPerDay|사용량이 많습니다/i.test(message);
}

function getGeminiErrorInfo(errorText, status, model) {
  let parsed = null;
  try {
    parsed = JSON.parse(errorText);
  } catch {
    parsed = null;
  }
  const rawMessage = parsed?.error?.message || errorText || "";
  const statusText = parsed?.error?.status || "";
  const base = {
    model,
    status,
    statusText,
    rawMessage,
    reason: "unknown"
  };
  if (/503|UNAVAILABLE|high demand/i.test(`${status} ${statusText} ${rawMessage}`)) {
    return {
      ...base,
      reason: "unavailable",
      message: `${model} 모델이 일시적으로 혼잡합니다. 잠시 후 다시 시도하거나 Flash Lite를 사용하세요.`
    };
  }
  if (/429|RESOURCE_EXHAUSTED|quota|rate limit|GenerateContentRequestsPerDay|RPM|TPM|RPD/i.test(`${status} ${statusText} ${rawMessage}`)) {
    return {
      ...base,
      reason: "rate_limit",
      message: `${model} 모델의 프로젝트 한도 또는 분당 한도에 걸렸습니다. 잠시 후 다시 시도하거나 다른 프로젝트/API Key 또는 Flash Lite를 사용하세요.`
    };
  }
  if (/bound service account is deleted or disabled|service account .*deleted|service account .*disabled/i.test(`${status} ${statusText} ${rawMessage}`)) {
    return {
      ...base,
      reason: "service_account_disabled",
      message: "이 API Key에 연결된 서비스 계정이 삭제되었거나 비활성화되어 Google API가 401로 거절했습니다. 같은 프로젝트에서 서비스 계정을 활성화하거나 새 Gemini API Key를 발급해 적용하세요."
    };
  }
  if (/API key not valid|permission|PERMISSION_DENIED|unauth|403|401/i.test(`${status} ${statusText} ${rawMessage}`)) {
    return {
      ...base,
      reason: "auth",
      message: "API Key가 유효하지 않거나 이 모델을 사용할 권한이 없습니다. 설정에서 새 키를 저장했는지 확인하세요."
    };
  }
  if (/not found|404|model/i.test(`${status} ${statusText} ${rawMessage}`)) {
    return {
      ...base,
      reason: "model_not_found",
      message: `${model} 모델을 사용할 수 없습니다. 설정에서 다른 Gemini 모델을 선택하세요.`
    };
  }
  return {
    ...base,
    message: rawMessage || LIMIT_MESSAGE
  };
}

function createGeminiApiError(message, info = {}) {
  const error = new Error(message || LIMIT_MESSAGE);
  error.gemini = info;
  return error;
}

function getGeminiErrorInfoFromError(error) {
  if (error?.gemini) return error.gemini;
  return {
    model: lastGeminiModelUsed,
    status: null,
    statusText: "",
    rawMessage: error?.message || "",
    reason: "client_error",
    message: error?.message || LIMIT_MESSAGE
  };
}

function logGeminiDiagnostic(info) {
  if (!window.console?.warn) return;
  console.warn("[Gemini diagnostic]", {
    model: info.model,
    status: info.status,
    statusText: info.statusText,
    reason: info.reason,
    rawMessage: info.rawMessage
  });
}

function parseGeminiJson(text) {
  if (typeof text !== "string") return text;
  const stripped = text
    .replace(/^```(?:json)?/i, "")
    .replace(/```$/i, "")
    .trim();
  try {
    return JSON.parse(stripped);
  } catch {
    const jsonObject = extractFirstJsonObject(stripped);
    if (jsonObject) return JSON.parse(jsonObject);
    throw new Error(LIMIT_MESSAGE);
  }
}

function extractFirstJsonObject(text) {
  const start = text.indexOf("{");
  if (start < 0) return "";
  let depth = 0;
  let inString = false;
  let escaping = false;
  for (let index = start; index < text.length; index += 1) {
    const char = text[index];
    if (inString) {
      if (escaping) {
        escaping = false;
      } else if (char === "\\") {
        escaping = true;
      } else if (char === "\"") {
        inString = false;
      }
      continue;
    }
    if (char === "\"") {
      inString = true;
      continue;
    }
    if (char === "{") depth += 1;
    if (char === "}") {
      depth -= 1;
      if (depth === 0) return text.slice(start, index + 1);
    }
  }
  return "";
}

function normalizeBrief(brief) {
  const expectedQaList = Array.isArray(brief.expectedQaList) ? brief.expectedQaList.map(normalizeQuestionForDealType) : [];
  const imIssueQuestions = buildImIssueQuestions();
  const normalizedQaList = normalizeExpectedQaList(expectedQaList, imIssueQuestions);
  return {
    fundSnapshot: normalizeBriefSnapshot(brief.fundSnapshot || {}),
    classificationSummary: brief.classificationSummary || {},
    marketContext: normalizeBriefMarketContext(brief.marketContext || {}),
    expectedQaList: normalizedQaList,
    keyRisks: asArray(brief.keyRisks),
    redFlags: asArray(brief.redFlags),
    expectedFollowUpRequests: asArray(brief.expectedFollowUpRequests),
    verificationItems: mergeTextLists([...asArray(brief.verificationItems), ...imIssueQuestions.map((item) => item.rationale)])
  };
}

async function refineBriefQuestionsForGp(brief) {
  const originalQuestions = asArray(brief?.expectedQaList);
  if (!originalQuestions.length) return brief;
  try {
    const text = await callGeminiText(buildQuestionSuitabilityPrompt(brief), { json: true, temperature: 0.1 });
    const parsed = parseGeminiJson(text);
    const refinedQuestions = normalizeGpQuestionReviewList(parsed.expectedQaList || parsed.questions || [], originalQuestions);
    if (refinedQuestions.length >= 5) {
      return {
        ...brief,
        expectedQaList: refinedQuestions,
        verificationItems: mergeTextLists([
          ...asArray(brief.verificationItems),
          ...asArray(parsed.removedOrMovedItems).map(formatListItemText)
        ])
      };
    }
  } catch (error) {
    console.warn("Q&A suitability review failed; applying local question cleanup.", error);
  }
  return {
    ...brief,
    expectedQaList: normalizeGpQuestionReviewList(originalQuestions, originalQuestions)
  };
}

function buildQuestionSuitabilityPrompt(brief) {
  return `
당신은 기관 LP가 GP 미팅에서 읽을 Q&A 리스트를 최종 편집하는 검토자입니다.
아래 Q&A가 "GP에게 직접 물어볼 수 있는 질문"인지 다시 판단하고, 부적절한 표현은 질문을 버리지 말고 GP가 답할 수 있는 딜 사실관계 질문으로 고쳐 JSON으로만 답하세요.

핵심 원칙:
- 검색/시장동향 생성은 이미 별도 단계에서 끝났습니다. 여기서는 Q&A 문장만 검토합니다.
- "사용자 입력", "LP 내부 입력", "세팅값", "앱 입력값", "내부 메모"처럼 GP가 알 수 없는 출처 표현은 question에 절대 쓰지 마세요.
- 단, 사용자가 입력한 메모나 우려사항에 담긴 주요 리스크는 버리지 마세요. GP가 답할 수 있는 사실관계, 계약 조건, 증빙자료, 리스크 완화 질문으로 재작성하세요.
- IM과 입력값이 충돌해 보일 때도 "사용자 입력과 IM이 불일치"라고 묻지 말고, "제공 자료 기준으로 각 당사자의 역할, 법적 지위, 책임 범위, 계약상 관계를 확인해 달라"처럼 물으세요.
- GP가 모를 LP 내부 판단 과정, UI 입력 출처, 모델 판단 근거를 묻는 질문은 금지입니다.
- GP가 알 수 있는 항목은 사업주체/차주/주주/운용사/시공사/신탁사/대주단 관계, 계약 책임, 전력/인허가, 임대/테넌트, EPC, 담보/보증, DSCR/LTV, 상환/리파이낸싱, 트랙레코드, alignment입니다.
- 운용사 경영권 매각, 주주 변경, 핵심인력 이탈/변동, 담당팀 변경, 시공사·신탁사·관계사 재무/소송/신용 이슈는 "GP가 확답하기 어렵다"는 이유로 삭제하지 마세요. 본 건 담당 조직, 의사결정 라인, 핵심인력 유지, 이해상충 관리, LP 보고 체계, 계약상 책임과 리스크 완화 조치에 대한 질문으로 재작성하세요.
- 질문은 최소 5개, 최대 10개를 유지하세요. 중요 리스크가 있으면 삭제보다 재작성하세요.
- 내부 사고 과정은 출력하지 마세요.

세팅값:
${JSON.stringify(state.meeting, null, 2)}

시장 맥락:
${JSON.stringify(state.marketContext || null, null, 2)}

브리프:
${JSON.stringify(brief, null, 2)}

응답 JSON 스키마:
{
  "expectedQaList": [
    {
      "category": "투자 thesis/구조/상환·Exit/담보·보증/현금흐름/트랙레코드/Alignment/리스크/IM 정합성 중 하나",
      "importance": "High/Medium/Low",
      "question": "GP에게 그대로 물을 수 있는 질문",
      "rationale": "LP 관점에서 확인해야 하는 이유",
      "source": "Q&A 적합성 검토"
    }
  ],
  "removedOrMovedItems": ["질문에서 제외하거나 확인 필요 항목으로 이동한 내부 출처/표현"]
}`;
}

function normalizeGpQuestionReviewList(reviewedQuestions, originalQuestions = []) {
  const cleaned = mergeQuestionLists(asArray(reviewedQuestions)
    .map(normalizeQuestionForDealType)
    .map(rewriteQuestionForGpAudience)
    .filter(isQuestionSuitableForGp));
  const marketQuestions = buildMarketDerivedGpQuestions();
  const withFallback = cleaned.length >= 5
    ? cleaned
    : mergeQuestionLists([...cleaned, ...marketQuestions, ...asArray(originalQuestions).map(rewriteQuestionForGpAudience), ...buildFallbackDealQuestions()])
      .filter(isQuestionSuitableForGp);
  return mergeQuestionLists([...marketQuestions, ...withFallback]).filter(isQuestionSuitableForGp).slice(0, 10);
}

function buildMarketDerivedGpQuestions() {
  const marketText = JSON.stringify(state.marketContext || {});
  const questions = [];
  if (/경영권|매각|인수|지분|최대주주|주주\s*변경|ownership|sale|acquisition/i.test(marketText)) {
    questions.push({
      category: "Alignment",
      importance: "High",
      question: "운용사 또는 주요 관계자의 경영권 매각·주주 변경 이슈가 본 건 담당 조직, IC 의사결정 라인, 핵심인력 유지 조건, LP 커뮤니케이션 체계에 미치는 영향은 무엇이며, 변동 발생 시 어떤 통지·승인 절차가 적용됩니까?",
      rationale: "운용 안정성, 핵심인력 유지, 의사결정 연속성 및 LP 보호 장치 확인 필요",
      source: "시장/관계자 뉴스"
    });
  }
  if (/핵심인력|인력|조직|담당팀|이탈|변동|key person|team/i.test(marketText)) {
    questions.push({
      category: "트랙레코드",
      importance: "High",
      question: "본 건을 담당하는 핵심 운용역과 실무팀 구성, 최근 12개월 내 인력 변동 여부, key person 또는 담당자 변경 시 LP에게 제공되는 보고·승인·보완 절차는 어떻게 정리되어 있습니까?",
      rationale: "담당팀 안정성과 운용 연속성 확인 필요",
      source: "시장/관계자 뉴스"
    });
  }
  if (/시공사|건설|책임준공|재무|신용|등급|유동성|소송|제재|부실|construction|credit|litigation/i.test(marketText)) {
    questions.push({
      category: "리스크",
      importance: "High",
      question: "시공사, 신탁사, 스폰서 등 주요 관계자의 재무·신용·소송·책임준공 관련 이슈가 본 건의 공정, 담보가치, 보증 이행 및 대주단 권리에 미치는 영향과 보완 장치는 무엇입니까?",
      rationale: "관계자 리스크가 PF 또는 단일 자산 거래의 실행력과 회수 가능성에 미치는 영향 확인 필요",
      source: "시장/관계자 뉴스"
    });
  }
  return questions;
}

function rewriteQuestionForGpAudience(item) {
  if (!item || typeof item !== "object") return item;
  let question = String(item.question || "").trim();
  let rationale = String(item.rationale || "").trim();
  const internalSourcePattern = /사용자\s*입력|LP\s*내부\s*입력|내부\s*메모|세팅값|앱\s*입력값|화면\s*입력값/gi;
  if (internalSourcePattern.test(question)) {
    question = question
      .replace(/사용자\s*입력\s*운용사와\s*IM상\s*사업주체·주주·운용\s*관련\s*당사자의\s*역할이\s*혼재되어\s*보입니다\.?/gi, "제공 자료 기준으로 본 건의 운용사, 사업주체, 주주 및 운용 관련 당사자의 역할 구분이 필요합니다.")
      .replace(/사용자\s*입력\s*내용과\s*IM상\s*주요\s*관계자의\s*역할\s*구분/gi, "제공 자료상 주요 관계자의 역할 구분")
      .replace(internalSourcePattern, "제공 자료");
  }
  if (internalSourcePattern.test(rationale)) {
    rationale = rationale.replace(internalSourcePattern, "제공 자료");
  }
  question = question
    .replace(/GP가\s*알\s*수\s*없는\s*/gi, "")
    .replace(/입력\s*출처/gi, "자료 출처")
    .replace(/\s+/g, " ")
    .trim();
  return {
    ...item,
    question,
    rationale,
    source: item.source === "사용자 입력" ? "Q&A 적합성 검토" : item.source || "Q&A 적합성 검토"
  };
}

function isQuestionSuitableForGp(item) {
  if (!item?.question) return false;
  const question = String(item.question);
  if (/사용자\s*입력|LP\s*내부\s*입력|내부\s*메모|세팅값|앱\s*입력값|화면\s*입력값/i.test(question)) return false;
  if (/왜\s*사용자가|어떤\s*필드|모델이\s*판단|LLM|프롬프트/i.test(question)) return false;
  return question.length >= 12;
}

function normalizeBriefSnapshot(snapshot = {}) {
  const reviewType = resolveReviewType(deriveEffectiveMeetingInfo(), snapshot, state.imProcessingResult?.imAnalysis || {});
  const normalized = { ...snapshot };
  const meeting = state.meeting || {};
  const userFacts = deriveUserProvidedDealFacts();
  if (isMeaningfulFact(meeting.managerName)) normalized.managerName = meeting.managerName;
  if (isMeaningfulFact(meeting.fundName)) normalized.fundName = meeting.fundName;
  if (isMeaningfulFact(meeting.strategy)) normalized.strategy = meeting.strategy;
  if (isMeaningfulFact(meeting.investmentStructure || meeting.capitalType)) normalized.capitalStructure = meeting.investmentStructure || meeting.capitalType;
  if (isMeaningfulFact(userFacts.region)) normalized.region = userFacts.region;
  if (isMeaningfulFact(meeting.assetClass) || isMeaningfulFact(meeting.sector)) {
    normalized.assetClass = [meeting.assetClass, meeting.sector].filter(isMeaningfulFact).join(" / ");
  }
  if (reviewType === "loan") {
    delete normalized.targetSize;
    delete normalized.investmentPeriod;
    delete normalized.targetReturn;
    if (isMeaningfulFact(userFacts.loanSize)) normalized.loanSize = userFacts.loanSize;
    if (isMeaningfulFact(userFacts.loanMaturity)) normalized.loanMaturity = userFacts.loanMaturity;
    if (isMeaningfulFact(userFacts.loanRate)) normalized.loanRate = userFacts.loanRate;
  } else if (reviewType === "fund") {
    delete normalized.loanSize;
    delete normalized.loanMaturity;
    delete normalized.loanRate;
  }
  return normalized;
}

function deriveUserProvidedDealFacts() {
  const text = [state.meeting?.fundName, state.meeting?.keyConcerns].filter(Boolean).join("\n");
  return {
    region: extractUserRegion(text),
    loanSize: extractUserLoanSize(text),
    loanMaturity: extractUserLoanMaturity(text),
    loanRate: extractUserLoanRate(text)
  };
}

function extractUserRegion(text = "") {
  const value = String(text || "");
  const provinceCity = value.match(/([가-힣]+도\s*[가-힣]+시)/);
  if (provinceCity) return provinceCity[1].replace(/\s+/g, " ");
  const cityDistrict = value.match(/([가-힣]+시\s*[가-힣0-9]+지구)/);
  if (cityDistrict) return cityDistrict[1].replace(/\s+/g, " ");
  const district = value.match(/([가-힣0-9]+지구|[가-힣]+시|[가-힣]+군|[가-힣]+구)/);
  return district ? district[1] : "";
}

function extractUserLoanSize(text = "") {
  const value = String(text || "");
  const match = value.match(/(?:총\s*)?(?:대출\s*규모|대출액|합계)[^.\n]{0,80}?([\d,]+)\s*억/) || value.match(/([\d,]+)\s*억(?:원)?\s*(?:합계|규모)/);
  return match ? `${match[1]}억원` : "";
}

function extractUserLoanMaturity(text = "") {
  const value = String(text || "");
  const match = value.match(/(?:대출기간|대출\s*만기|만기)[^.\n]{0,60}?(\d+)\s*개월/) || value.match(/최초\s*인출일로부터\s*(\d+)\s*개월/);
  return match ? `최초 인출일로부터 ${match[1]}개월` : "";
}

function extractUserLoanRate(text = "") {
  const value = String(text || "");
  const trA = value.match(/Tr\.?\s*A[^.\n]{0,60}?(\d+(?:\.\d+)?)\s*%/i);
  const allIn = value.match(/All-?in[^.\n]{0,40}?(\d+(?:\.\d+)?)\s*~\s*(\d+(?:\.\d+)?)\s*%/i);
  const fund = value.match(/(?:주택도시기금|기금)[^.\n]{0,50}?(\d+(?:\.\d+)?)\s*%/);
  const parts = [];
  if (trA) parts.push(`Tr.A ${trA[1]}%`);
  if (allIn) parts.push(`All-in ${allIn[1]}~${allIn[2]}%`);
  if (fund) parts.push(`주택도시기금 ${fund[1]}%`);
  return parts.join(", ");
}

function normalizeBriefMarketContext(context = {}) {
  const external = state.marketContext || {};
  if (external.groundingDiagnostics || asArray(external.keyMarketTrends).length || asArray(external.recentEvents).length || asArray(external.policyRegulatoryNotes).length || asArray(external.riskSignals).length) {
    return {
      summary: external.summary || "",
      directDealEvents: asArray(external.directDealEvents || external.recentEvents),
      trends: asArray(external.keyMarketTrends),
      newsPolicy: asArray(external.policyRegulatoryNotes),
      riskSignals: asArray(external.riskSignals),
      sources: asArray(external.sources),
      sourceQuality: external.sourceQuality || ""
    };
  }
  return {
    summary: context.summary || external.summary || "",
    directDealEvents: asArray(context.directDealEvents).length ? context.directDealEvents : asArray(external.directDealEvents || external.recentEvents),
    trends: asArray(context.trends).length ? context.trends : asArray(external.keyMarketTrends),
    newsPolicy: asArray(context.newsPolicy).length
      ? context.newsPolicy
      : asArray(external.policyRegulatoryNotes),
    riskSignals: asArray(context.riskSignals).length ? context.riskSignals : asArray(external.riskSignals),
    sources: asArray(context.sources).length ? context.sources : asArray(external.sources),
    sourceQuality: context.sourceQuality || external.sourceQuality || ""
  };
}

function normalizeExpectedQaList(aiQuestions = [], imIssueQuestions = []) {
  const normalizedAi = mergeQuestionLists(asArray(aiQuestions).map(normalizeQuestionForDealType));
  const dealQuestions = normalizedAi.filter((item) => !isImConsistencyQuestion(item));
  const aiImQuestions = normalizedAi.filter(isImConsistencyQuestion);
  const imQuestions = mergeQuestionLists([...aiImQuestions, ...asArray(imIssueQuestions)]).slice(0, 2);
  const dealLimit = imQuestions.length ? 8 : 10;
  const base = mergeQuestionLists([...dealQuestions.slice(0, dealLimit), ...imQuestions]).slice(0, 10);
  if (base.length >= 5) return base;
  const fallbackQuestions = buildFallbackDealQuestions();
  return mergeQuestionLists([...base, ...fallbackQuestions]).slice(0, 10);
}

function isImConsistencyQuestion(item) {
  const text = [
    item?.category,
    item?.question,
    item?.rationale,
    item?.source
  ].filter(Boolean).join(" ");
  return /IM\s*숫자|IM\s*정합성|원문\s*대조|오기|다르게|복수\s*기재|불일치|conflict|verification|keyNumbersToVerify/i.test(text);
}

function buildImIssueQuestions() {
  const analysis = state.imProcessingResult?.imAnalysis || {};
  const aiIssues = [
    ...asArray(analysis.conflicts),
    ...asArray(analysis.keyNumbersToVerify)
  ].map((item) => makeImIssueQuestion(formatListItemText(item), "AI/IM 분석"));
  const detectedIssues = detectNumericInconsistencies(state.imProcessingResult?.textExcerpt || "")
    .map((item) => makeImIssueQuestion(item, "IM 숫자 정합성"));
  return mergeQuestionLists([...detectedIssues, ...aiIssues])
    .filter(isMaterialImIssueQuestion)
    .slice(0, 2);
}

function makeImIssueQuestion(issueText, source = "IM 원문 대조") {
  const text = String(issueText || "").trim();
  if (!text) return null;
  const numberLike = /%|bp|bps|금리|수익률|LTV|DSCR|기간|만기|금액|억|원|배|개월|년|상환|보증|약정|수수료/i.test(text);
  if (!numberLike && !/충돌|불일치|상이|다르게|오류|오기|누락|확인 필요/i.test(text)) return null;
  const isActualConflict = /충돌|불일치|상이|다르게|복수\s*기재|오류|오기/i.test(text) || source === "IM 숫자 정합성";
  return {
    category: numberLike ? "IM 숫자/조건 정합성" : "IM 원문 대조",
    importance: numberLike ? "High" : "Medium",
    question: numberLike
      ? isActualConflict
        ? `IM에서 ${text} 항목의 기준 또는 수치가 원문상 다르게 보입니다. 정확한 적용 기준, 산식, 기준일 및 오기 여부를 확인해주실 수 있습니까?`
        : `IM에서 ${text} 항목은 투자 조건 판단에 중요합니다. 정확한 적용 기준, 산식, 기준일 및 증빙 자료를 확인해주실 수 있습니까?`
      : `IM에서 ${text} 항목은 원문 대조 또는 추가 설명이 필요해 보입니다. 정확한 사실관계와 판단 근거를 설명해주실 수 있습니까?`,
    rationale: `IM 원문 확인 필요: ${text}`,
    source
  };
}

function isMaterialImIssueQuestion(item) {
  if (!item) return false;
  const text = [item.question, item.rationale].filter(Boolean).join(" ");
  return /%|bp|bps|금리|수익률|LTV|DSCR|기간|만기|금액|억|원|배|개월|년|상환|보증|약정|수수료|충돌|불일치|상이|오류|오기/i.test(text);
}

function buildFallbackDealQuestions() {
  const effectiveMeeting = deriveEffectiveMeetingInfo();
  const directLoan = isDirectLoanLikeCase();
  const base = directLoan
    ? [
        ["구조", "본 건의 차주, 담보, 선순위성, 대주간 권리관계 및 주요 약정 조건을 설명해주실 수 있습니까?", "대출 구조와 회수 가능성 판단의 기본 전제 확인 필요"],
        ["상환재원", "Base/Downside 시나리오별 상환재원, DSCR/LTV 완충 여력 및 만기 대응 계획은 어떻게 됩니까?", "상환 안정성과 downside 방어력 확인 필요"],
        ["담보·보증", "담보권, 보증, 책임준공 또는 기타 신용보강의 실행 요건과 예외 조항은 무엇입니까?", "계약상 보호 장치의 실효성 확인 필요"],
        ["리스크", "본 건에서 GP가 보는 핵심 downside 리스크와 이를 통제하기 위한 covenant 또는 모니터링 체계는 무엇입니까?", "리스크 관리 체계 확인 필요"],
        ["Exit / Refinancing", "만기 전 refinancing 또는 take-out 가능성, 실패 시 대안 시나리오는 무엇입니까?", "회수 경로와 유동성 리스크 확인 필요"]
      ]
    : [
        ["투자 thesis", "본 건의 핵심 투자 thesis와 현 시점에 해당 전략을 집행해야 하는 근거는 무엇입니까?", "투자 판단의 핵심 논리 확인 필요"],
        ["트랙레코드", "동일 전략/섹터에서 실현 또는 회수 완료된 track record와 이번 건에 적용 가능한 교훈은 무엇입니까?", "GP 실행 역량 검증 필요"],
        ["포트폴리오 / 자산", "초기 포트폴리오 또는 주요 투자대상별 수익 창출 경로와 downside 방어 장치는 무엇입니까?", "현금흐름과 가치상승 경로 확인 필요"],
        ["Exit", "주요 Exit 경로, 예상 buyer pool, holding period 연장 시 대응 방안은 무엇입니까?", "회수 가능성과 기간 리스크 확인 필요"],
        ["Alignment", "GP commitment, key man, 보수/성과보수 구조가 LP와 이해관계를 어떻게 정렬합니까?", "LP 보호와 이해상충 가능성 확인 필요"]
      ];
  return base.map(([category, question, rationale]) => ({
    category,
    importance: category === "투자 thesis" || category === "구조" ? "High" : "Medium",
    question: rewriteFundOnlyQuestion(question),
    rationale,
    source: effectiveMeeting.fundName ? "기본 DDQ" : "IM/기본 DDQ"
  }));
}

function detectNumericInconsistencies(text) {
  const compact = String(text || "").replace(/\s+/g, " ");
  if (!compact) return [];
  const metrics = [
    { label: "주택도시기금 금리", pattern: /(?:주택도시기금|기금)[^.%]{0,80}?금리[^.%]{0,80}?(\d+(?:\.\d+)?)\s*%|금리[^.%]{0,80}?(?:주택도시기금|기금)[^.%]{0,80}?(\d+(?:\.\d+)?)\s*%/gi, unit: "%" },
    { label: "HUG 보증 관련 금리", pattern: /HUG[^.%]{0,80}?금리[^.%]{0,80}?(\d+(?:\.\d+)?)\s*%|금리[^.%]{0,80}?HUG[^.%]{0,80}?(\d+(?:\.\d+)?)\s*%/gi, unit: "%" },
    { label: "대출 금리", pattern: /(?:대출|차입|PF)[^.%]{0,80}?금리[^.%]{0,80}?(\d+(?:\.\d+)?)\s*%|금리[^.%]{0,80}?(?:대출|차입|PF)[^.%]{0,80}?(\d+(?:\.\d+)?)\s*%/gi, unit: "%" },
    { label: "LTV", pattern: /LTV[^.%]{0,60}?(\d+(?:\.\d+)?)\s*%/gi, unit: "%" },
    { label: "DSCR", pattern: /DSCR[^.\n]{0,60}?(\d+(?:\.\d+)?)\s*(?:x|배)?/gi, unit: "x" }
  ];
  return metrics.flatMap(({ label, pattern, unit }) => {
    const values = collectMetricValues(compact, pattern, unit);
    return values.length > 1 ? [`${label}가 ${values.join(", ")}로 복수 기재됨`] : [];
  });
}

function collectMetricValues(text, pattern, unit = "") {
  const values = new Set();
  for (const match of text.matchAll(pattern)) {
    const value = match.slice(1).find(Boolean);
    if (value) values.add(`${value}${unit}`);
  }
  return [...values];
}

function mergeQuestionLists(items) {
  const seen = new Set();
  return items.filter(Boolean).filter((item) => {
    const key = String(item.question || item).replace(/\s+/g, " ").trim().toLowerCase();
    if (!key || seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function mergeTextLists(items) {
  const seen = new Set();
  return items.filter(Boolean).map((item) => String(item).trim()).filter((item) => {
    const key = item.replace(/\s+/g, " ").toLowerCase();
    if (!key || seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function normalizeQuestionForDealType(item) {
  if (!item || typeof item !== "object") return item;
  const normalized = {
    ...item,
    question: isDirectLoanLikeCase() ? rewriteFundOnlyQuestion(item.question) : item.question,
    rationale: isDirectLoanLikeCase() ? rewriteFundOnlyQuestion(item.rationale) : item.rationale
  };
  return softenUnsupportedRoleAssertion(normalized);
}

function rewriteFundOnlyQuestion(value) {
  if (!value) return value;
  return String(value)
    .replace(/본 펀드/g, "본 건")
    .replace(/펀드명/g, "펀드명 / 대출명")
    .replace(/정확한 명칭과 투자 구조\(Senior Debt 외 다른 트랜치 존재 여부 등\)는 어떻게 됩니까\?/g, "정확한 대출명, 차주/담보 구조, 선순위성 및 주요 약정 조건은 어떻게 됩니까?")
    .replace(/Senior Debt 외 다른 트랜치 존재 여부/g, "후순위/메자닌 동반 여부가 실제로 필요한 경우 그 사유");
}

function softenUnsupportedRoleAssertion(item) {
  const text = [item.category, item.question, item.rationale, item.source].filter(Boolean).join(" ");
  const roleAssertion = /(운용사|GP|사업주체|주관|주요주주|시공사|신탁사).*(불일치|확인됩니다|확인됨|주관으로|기재되어 있으나|역할)/i.test(text);
  const namedPartyMix = /(월넛|이지스|자산운용|PFV|피에프브이|사업주체|주요주주)/i.test(text) && /(역할|법적\s*지위|불일치|주관)/i.test(text);
  if (!roleAssertion && !namedPartyMix) return item;
  return {
    ...item,
    category: "Alignment",
    importance: item.importance === "Low" ? "Medium" : item.importance || "Medium",
    question: "본 건의 운용사, 사업주체, 주주 및 운용 관련 당사자의 역할 구분이 필요합니다. 각 당사자의 역할, 법적 지위, 책임 범위 및 대주단과의 계약상 관계는 어떻게 정리됩니까?",
    rationale: "관계자 역할을 단정하지 않고 계약상 책임 주체와 이해관계 정렬 여부를 확인할 필요가 있음",
    source: item.source || "Q&A 적합성 검토"
  };
}

function isDirectLoanLikeCase() {
  const text = [
    state.meeting.fundName,
    state.meeting.assetClass,
    state.meeting.strategy,
    state.meeting.sector,
    state.meeting.capitalType,
    state.meeting.investmentStructure,
    state.meeting.keyConcerns,
    state.imProcessingResult?.imAnalysis?.strategySummary,
    JSON.stringify(state.imProcessingResult?.imAnalysis?.autoDetectedFields || {})
  ].join(" ").toLowerCase();
  return /대출|loan|debt|credit|pf|project finance|senior|차주|담보|보증|hug|단일|tranche|트랜치|private debt|pd|direct lending/.test(text);
}

function makeQuestionRecords(brief) {
  const questions = brief.expectedQaList.length
    ? brief.expectedQaList
    : asArray(state.imProcessingResult?.imAnalysis?.mustAskQuestionsFromIm).map((question) => ({ question, category: "IM", importance: "High" }));
  return questions.map((item, index) => ({
    id: crypto.randomUUID ? crypto.randomUUID() : `q-${Date.now()}-${index}`,
    category: item.category || "질문",
    importance: item.importance || "Medium",
    question: item.question || String(item),
    rationale: item.rationale || "",
    source: item.source || "",
    answerRecord: {
      status: "미확인",
      answer: "",
      internalMemo: "",
      followUpNeeded: false,
      importantQuote: ""
    }
  }));
}

function renderBrief() {
  const container = $("briefOutput");
  if (!state.preMeetingBrief) {
    container.innerHTML = `
      <div class="rounded-xl border border-dashed border-slate-300 bg-white p-8 text-center text-sm text-slate-500">
        브리프를 생성하면 Fund Snapshot, 시장 맥락, 예상 Q&A, 리스크, Red Flag가 여기에 표시됩니다.
      </div>`;
    return;
  }
  const brief = state.preMeetingBrief;
  container.innerHTML = `
    <div class="grid gap-4 lg:grid-cols-2">
      ${card("Fund Snapshot", renderObjectList(brief.fundSnapshot))}
      ${card("Investment Classification Summary", renderClassification(brief.classificationSummary))}
      ${card("AI Market Context", renderMarketContext(brief.marketContext), "lg:col-span-2")}
      ${card("예상 Q&A", renderQuestionsPreview(brief.expectedQaList), "lg:col-span-2")}
      ${card("핵심 리스크", renderList(brief.keyRisks))}
      ${card("Red Flag Checklist", renderList(brief.redFlags))}
      ${card("Expected Follow-up Requests", renderList(brief.expectedFollowUpRequests))}
      ${card("확인 필요 항목", renderList(brief.verificationItems))}
    </div>
    <div class="mt-5 flex justify-center">
      <button id="goLiveButton" class="inline-flex items-center justify-center gap-2 rounded-md bg-slate-900 px-5 py-2.5 text-sm font-extrabold text-white shadow-sm transition hover:bg-slate-800">
        미팅 노트로 이동
        <i data-lucide="arrow-right" class="h-4 w-4"></i>
      </button>
    </div>`;
  $("goLiveButton").addEventListener("click", () => switchPhase("live"));
  refreshIcons();
}

function card(title, body, extraClass = "") {
  return `<div class="brief-card ${extraClass}"><h3>${escapeHtml(title)}</h3>${body}</div>`;
}

function renderObjectList(object = {}) {
  const entries = Object.entries(object).filter(([, value]) => value !== undefined && value !== null && String(value).trim() !== "");
  if (!entries.length) return emptyText();
  return `<dl class="space-y-2 text-sm">${entries.map(([key, value]) => `
    <div class="grid grid-cols-[130px_1fr] gap-3">
      <dt class="font-bold text-slate-500">${escapeHtml(labelize(key))}</dt>
      <dd class="text-slate-800">${Array.isArray(value) ? renderInlineList(value) : escapeHtml(String(value))}</dd>
    </div>`).join("")}</dl>`;
}

function renderClassification(summary = {}) {
  return [
    ["IM에서 확인", summary.confirmedFromIm],
    ["사용자 입력", summary.providedByUser],
    ["충돌 정보", summary.conflicts],
    ["확인 필요", summary.needsVerification]
  ].map(([title, items]) => `<div class="mb-3"><div class="mb-1 text-xs font-extrabold text-slate-500">${title}</div>${renderList(items)}</div>`).join("");
}

function hasWeakMarketDate(value) {
  const text = String(value || "").trim();
  return !text || /확인\s*필요|unknown|n\/a|미상|불명/i.test(text);
}

function isWeakMarketItem(item) {
  const text = formatListItemText(item);
  if (/날짜\s*확인\s*필요|출처\s*확인\s*필요|최근\s*자료\s*확인\s*필요/i.test(text)) return true;
  if (item && typeof item === "object" && !Array.isArray(item)) {
    return hasWeakMarketDate(item.date || item.publishedAt || item.asOfDate) || !String(item.source || "").trim();
  }
  return false;
}

function renderMarketEvidenceList(items) {
  const filtered = asArray(items).filter((item) => item && !isWeakMarketItem(item));
  if (!filtered.length) return `<p class="text-sm leading-6 text-slate-400">검색일 기준 최근 1년 내 날짜와 출처가 확인된 근거 없음</p>`;
  return `<div class="space-y-2">${filtered.map(renderMarketEvidenceItem).join("")}</div>`;
}

function renderMarketEvidenceItem(item) {
  if (!item || typeof item !== "object" || Array.isArray(item)) {
    return `<div class="rounded-md border border-slate-200 bg-slate-50 p-3 text-sm leading-6 text-slate-700">${escapeHtml(formatListItemText(item))}</div>`;
  }
  const date = item.date || item.publishedAt || item.asOfDate || "";
  const source = item.source || "";
  const title = item.title || item.headline || item.name || "";
  const fact = item.fact || item.note || item.summary || item.content || item.text || "";
  const relevance = item.relevance || item.implication || item.lpView || "";
  return `
    <div class="rounded-md border border-slate-200 bg-slate-50 p-3">
      <div class="mb-1 flex flex-wrap items-center gap-2 text-[11px] font-bold text-slate-500">
        <span class="rounded bg-white px-2 py-0.5 text-slate-700">${escapeHtml(date)}</span>
        <span>${escapeHtml(source)}</span>
      </div>
      ${title ? `<div class="text-sm font-extrabold leading-5 text-slate-900">${escapeHtml(title)}</div>` : ""}
      ${fact ? `<div class="mt-1 text-sm leading-6 text-slate-700">${escapeHtml(fact)}</div>` : ""}
      ${relevance ? `<div class="mt-1 text-xs leading-5 text-slate-500">관련성: ${escapeHtml(relevance)}</div>` : ""}
    </div>`;
}

function renderMarketContext(context = {}) {
  const directDealEvents = asArray(context.directDealEvents || context.recentEvents);
  const trends = asArray(context.trends || context.keyMarketTrends);
  const newsPolicy = asArray(context.newsPolicy).length
    ? context.newsPolicy
    : asArray(context.policyRegulatoryNotes);
  const sources = asArray(context.sources);
  return `
    ${context.summary ? `<p class="mb-3 text-sm leading-6 text-slate-700">${escapeHtml(context.summary)}</p>` : ""}
    <div class="grid gap-3 md:grid-cols-2 xl:grid-cols-4">
      <div><div class="mb-2 text-xs font-extrabold text-slate-500">직접 관련 뉴스</div>${renderMarketEvidenceList(directDealEvents)}</div>
      <div><div class="mb-2 text-xs font-extrabold text-slate-500">시장 참고자료</div>${renderMarketEvidenceList(trends)}</div>
      <div><div class="mb-2 text-xs font-extrabold text-slate-500">정책/규제</div>${renderMarketEvidenceList(newsPolicy)}</div>
      <div><div class="mb-2 text-xs font-extrabold text-slate-500">리스크 신호</div>${renderMarketEvidenceList(context.riskSignals)}</div>
    </div>
    ${sources.length ? `<div class="mt-4 border-t border-slate-100 pt-3">
      <div class="mb-2 text-xs font-extrabold text-slate-500">출처</div>
      ${renderMarketEvidenceList(sources)}
    </div>` : ""}
    ${context.sourceQuality ? `<p class="mt-3 text-xs leading-5 text-slate-500">출처 품질: ${escapeHtml(context.sourceQuality)}</p>` : ""}`;
}

function renderQuestionsPreview(questions = []) {
  if (!questions.length) return emptyText();
  return `<div class="space-y-3">${questions.map((item, index) => `
    <div class="rounded-lg border border-slate-200 bg-slate-50 p-3">
      <div class="mb-1 text-xs font-extrabold text-brand-700">Q${index + 1}. ${escapeHtml(item.category || "질문")} · ${escapeHtml(item.importance || "Medium")}</div>
      <div class="text-sm font-bold leading-6 text-slate-900">${escapeHtml(item.question || "")}</div>
      ${item.rationale ? `<div class="mt-1 text-xs leading-5 text-slate-500">${escapeHtml(item.rationale)}</div>` : ""}
    </div>`).join("")}</div>`;
}

function renderQuestions() {
  const list = $("questionList");
  const questions = state.questionRecords || [];
  $("questionCount").textContent = `${questions.length}개`;
  if (!questions.length) {
    list.innerHTML = `<div class="rounded-lg border border-dashed border-slate-300 bg-white p-5 text-center text-sm text-slate-500">사전 브리프를 만들면 질문 카드가 자동 생성됩니다.</div>`;
    return;
  }
  list.innerHTML = questions.map((q, index) => `
    <button class="question-card w-full text-left ${index === state.selectedQuestionIndex ? "selected" : ""}" data-question-index="${index}">
      <div class="meta">Q${index + 1}. ${escapeHtml(q.category)} · ${escapeHtml(q.importance)}</div>
      <p>${escapeHtml(q.question)}</p>
      <div class="mt-3 flex items-center gap-2 text-xs text-slate-500">
        ${q.answerRecord.answer?.trim() ? `<span class="rounded bg-slate-100 px-2 py-1 font-bold">답변 기록</span>` : ""}
        ${q.answerRecord.followUpNeeded ? `<span class="rounded bg-amber-100 px-2 py-1 font-bold text-amber-700">추가 검토</span>` : ""}
      </div>
    </button>`).join("");
  list.querySelectorAll("[data-question-index]").forEach((button) => {
    button.addEventListener("click", () => {
      state.selectedQuestionIndex = Number(button.dataset.questionIndex);
      renderQuestions();
      renderSelectedQuestion();
      saveState();
    });
  });
}

function renderSelectedQuestion() {
  const panel = $("selectedQuestionPanel");
  const q = state.questionRecords[state.selectedQuestionIndex];
  if (!q) {
    panel.innerHTML = `<div class="py-12 text-center text-sm text-slate-500">선택된 질문이 없습니다.</div>`;
    return;
  }
  panel.innerHTML = `
    <div class="mb-4 flex flex-wrap items-start justify-between gap-3">
      <div>
        <div class="text-xs font-extrabold uppercase text-brand-700">Q${state.selectedQuestionIndex + 1}. ${escapeHtml(q.category)} · ${escapeHtml(q.importance)}</div>
        <h3 class="mt-2 text-lg font-extrabold leading-7 text-slate-900">${escapeHtml(q.question)}</h3>
        ${q.rationale ? `<p class="mt-2 text-sm leading-6 text-slate-500">${escapeHtml(q.rationale)}</p>` : ""}
      </div>
    </div>
    <div class="grid gap-4">
      <label class="field-label">운용사 답변
        <textarea id="gpAnswer" class="field-control min-h-36" placeholder="운용사 답변을 질문별로 기록하세요.">${escapeHtml(q.answerRecord.answer)}</textarea>
      </label>
      <label class="field-label">내부 메모
        <textarea id="internalMemo" class="field-control min-h-28" placeholder="LP 내부 판단, 추가 검토 의견, 우려사항">${escapeHtml(q.answerRecord.internalMemo)}</textarea>
      </label>
      <label class="field-label">중요 발언
        <input id="importantQuote" class="field-control" type="text" value="${escapeAttribute(q.answerRecord.importantQuote)}" placeholder="인용하거나 원문 대조할 발언">
      </label>
      <label class="inline-flex items-center gap-2 text-sm font-bold text-slate-700">
        <input id="followUpNeeded" type="checkbox" class="h-4 w-4 rounded border-slate-300 text-brand-600" ${q.answerRecord.followUpNeeded ? "checked" : ""}>
        추가 검토 필요
      </label>
    </div>`;

  ["gpAnswer", "internalMemo", "importantQuote", "followUpNeeded"].forEach((id) => {
    $(id).addEventListener("input", updateSelectedQuestionRecord);
    $(id).addEventListener("change", updateSelectedQuestionRecord);
  });
}

function updateSelectedQuestionRecord() {
  const q = state.questionRecords[state.selectedQuestionIndex];
  if (!q) return;
  const nextRecord = {
    answer: $("gpAnswer").value,
    internalMemo: $("internalMemo").value,
    importantQuote: $("importantQuote").value,
    followUpNeeded: $("followUpNeeded").checked
  };
  q.answerRecord = {
    status: deriveAnswerRecordStatus(nextRecord),
    ...nextRecord
  };
  saveState();
  scheduleExistingLibraryUpdate();
  renderQuestions();
}

function addQuestion() {
  state.questionRecords.push({
    id: crypto.randomUUID ? crypto.randomUUID() : `q-${Date.now()}`,
    category: "사용자 추가",
    importance: "Medium",
    question: "새 질문을 입력하세요.",
    rationale: "",
    source: "사용자",
    answerRecord: {
      status: "미확인",
      answer: "",
      internalMemo: "",
      followUpNeeded: false,
      importantQuote: ""
    }
  });
  state.selectedQuestionIndex = state.questionRecords.length - 1;
  saveState();
  scheduleExistingLibraryUpdate();
  renderQuestions();
  renderSelectedQuestion();
}

function renderReport() {
  const container = $("reportOutput");
  if (!state.postMeetingMemo) {
    container.innerHTML = `
      <div class="flex flex-col items-center justify-center py-16 text-center text-slate-400">
        <i data-lucide="file-signature" class="mb-3 h-12 w-12 opacity-50"></i>
        <p class="text-sm">미팅 기록을 완료한 뒤 최종 보고서를 생성하세요.</p>
      </div>`;
    refreshIcons();
    return;
  }
  container.innerHTML = markdownToHtml(buildReportMarkdown());
  refreshIcons();
}

function buildReportMarkdown() {
  const memo = state.postMeetingMemo;
  if (!memo) return "";
  const effectiveMeeting = deriveEffectiveMeetingInfo(memo);
  const summary = memo.internalReportSummary || {};
  const oneLineView = formatListItemText(summary.oneLineView) || "확인 필요";
  const investmentMemo = formatListItemText(summary.investmentMemo) || "확인 필요";
  const riskView = formatListItemText(summary.riskView) || "확인 필요";
  return `# ${effectiveMeeting.managerName || "운용사"} / ${effectiveMeeting.fundName || "건명"} 미팅 정리

## 1. 논점별 회의록
${asArray(memo.issueBasedMeetingNotes).map((item) => `- **${formatListItemText(item.issue) || "논점"}**: ${formatListItemText(item.summary) || ""}${item.evidence ? `\n  - 근거: ${formatListItemText(item.evidence)}` : ""}`).join("\n") || "- 확인 필요"}

## 2. 미해결 쟁점
${listMarkdown(memo.unresolvedIssues)}

## 3. Follow-up Request
${listMarkdown(asArray(memo.followUpRequestList).slice(0, 3))}

## 4. 운용사 이메일 초안
${memo.followUpEmailDraft || "확인 필요"}

## 5. 내부 보고용 간이보고서
**한 줄 의견:** ${oneLineView}

${investmentMemo}

**리스크 관점:** ${riskView}

## 6. Next Action Items
${listMarkdown(memo.nextActionItems)}

## 7. Quality Checks
${listMarkdown(memo.qualityChecks)}

## 8. 원문 대조 필요 항목
${listMarkdown(memo.sourceVerificationItems)}
`;
}

function getReportToneConfig(toneValue = "neutral") {
  const configs = {
    neutral: {
      label: "중립적 리뷰",
      reviewOpinion: "추가 DD 필요",
      status: "추가 DD 필요",
      promptInstruction: "장점과 리스크를 균형 있게 쓰고, 판단을 유보할 부분은 확인 필요로 남깁니다."
    },
    positive: {
      label: "긍정적 검토",
      reviewOpinion: "조건부 긍정",
      status: "조건부 검토 지속",
      promptInstruction: "투자 적합성과 긍정 요소를 먼저 정리하되, 투심 상정 전 확인할 조건을 명확히 둡니다."
    },
    conservative: {
      label: "보수적 리스크 강조",
      reviewOpinion: "보수적 검토",
      status: "추가 DD 필요",
      promptInstruction: "downside, 원문 대조 필요 항목, 미해결 쟁점을 앞세우고 보수적인 문장으로 작성합니다."
    }
  };
  return configs[toneValue] || configs.neutral;
}

function getQuestionRecordsForReport() {
  return (state.questionRecords || []).map((question, index) => ({
    no: index + 1,
    category: question.category || "",
    importance: question.importance || "",
    question: question.question || "",
    rationale: question.rationale || "",
    gpAnswer: question.answerRecord?.answer || "",
    internalMemo: question.answerRecord?.internalMemo || "",
    importantQuote: question.answerRecord?.importantQuote || "",
    needsFurtherReview: Boolean(question.answerRecord?.followUpNeeded)
  }));
}

function buildDocxMemoData() {
  const memo = state.postMeetingMemo || {};
  const tone = getReportToneConfig(state.reportTone);
  const summary = memo.internalReportSummary || {};
  const questions = (state.questionRecords || []).slice(0, 5);
  const risks = [
    ...asArray(state.preMeetingBrief?.keyRisks),
    ...asArray(memo.unresolvedIssues)
  ].slice(0, 6);
  const followUps = asArray(memo.followUpRequestList).length ? memo.followUpRequestList : state.preMeetingBrief?.expectedFollowUpRequests;
  const snapshot = state.preMeetingBrief?.fundSnapshot || {};
  const imAnalysis = state.imProcessingResult?.imAnalysis || {};
  const effectiveMeeting = deriveEffectiveMeetingInfo(memo);
  const rawOverviewFacts = deriveMeetingOverviewFacts(memo, snapshot, imAnalysis);
  const reviewType = resolveReviewType(effectiveMeeting, rawOverviewFacts, imAnalysis);
  const overviewFacts = normalizeOverviewFactsForReviewType(rawOverviewFacts, reviewType);
  const overviewMetricRows = buildOverviewMetricRows(overviewFacts, reviewType);
  const assetName = deriveAssetNameForMemo(effectiveMeeting, snapshot, imAnalysis, memo);
  const classificationParts = [
    localizeDisplayTerm(effectiveMeeting.locationType),
    localizeDisplayTerm(effectiveMeeting.assetClass),
    localizeDisplayTerm(effectiveMeeting.sector),
    localizeDisplayTerm(effectiveMeeting.strategy),
    localizeDisplayTerm(effectiveMeeting.capitalType)
  ].filter(Boolean);
  const data = {
    title: "LP Meeting Review Memo",
    reviewType,
    assetName,
    subjectLabel: reviewType === "fund" ? "펀드명" : reviewType === "loan" ? "대출명" : "펀드명 / 대출명",
    managerName: effectiveMeeting.managerName || "운용사 확인 필요",
    fundName: effectiveMeeting.fundName || assetName || "펀드명 / 대출명 확인 필요",
    contactName: effectiveMeeting.gpParticipants || effectiveMeeting.contactName || "",
    gpAttendees: effectiveMeeting.gpParticipants || effectiveMeeting.contactName || "",
    lpAttendees: effectiveMeeting.lpParticipants || "",
    meetingDate: effectiveMeeting.meetingDate || "",
    meetingLocation: overviewFacts.location || "",
    classification: classificationParts.join(" · ") || "분류 확인 필요",
    regionAssetClass: overviewFacts.regionAssetClass || "",
    sizeOrLoanAmount: overviewFacts.sizeOrLoanAmount || "",
    periodOrMaturity: overviewFacts.periodOrMaturity || "",
    returnOrLoanRate: overviewFacts.returnOrLoanRate || "",
    commitmentAmount: overviewFacts.commitmentAmount || "",
    overviewMetricRows,
    status: tone.status,
    reviewOpinion: tone.reviewOpinion,
    oneLineView: formatListItemText(summary.oneLineView) || "핵심 결론 확인 필요",
    reviewConclusion: formatListItemText(summary.reviewConclusion) || formatListItemText(summary.oneLineView) || "",
    whyThisDeal: formatListItemText(summary.whyThisDeal || summary.investmentMerits || summary.dealMerits || summary.whyDeal) || "",
    keyRisksForOpinion: formatListItemText(summary.keyRisks || summary.keyRiskSummary || summary.riskView) || "",
    nextDdItemsForOpinion: formatListItemText(summary.nextDdItems || summary.nextDdItemsSummary || summary.followUpSummary) || "",
    evidence: deriveExecutiveEvidence(memo),
    merits: asArray(state.preMeetingBrief?.keyInvestmentMerits).slice(0, 4),
    overview: [
      ["GP / 건명", `${effectiveMeeting.managerName || "확인 필요"} / ${effectiveMeeting.fundName || "확인 필요"}`],
      ["투자 지역", localizeDisplayTerm(effectiveMeeting.locationType) || "확인 필요"],
      ["자산군 / 전략", `${localizeDisplayTerm(effectiveMeeting.assetClass) || "확인 필요"} / ${localizeDisplayTerm(effectiveMeeting.strategy) || "확인 필요"}`],
      ["섹터 / 구조", `${localizeDisplayTerm(effectiveMeeting.sector) || "확인 필요"} / ${localizeDisplayTerm(effectiveMeeting.investmentStructure || effectiveMeeting.capitalType) || "확인 필요"}`],
      ["딜 / 자산 메모", effectiveMeeting.keyConcerns || "확인 필요"],
      ...overviewMetricRows,
      ["당사 검토 약정액", overviewFacts.commitmentAmount || "직접 입력"]
    ],
    ddq: questions.map((question, index) => ({
      label: `Q${index + 1}. ${question.category || "DDQ"}`,
      question: question.question || "질문 확인 필요",
      answer: question.answerRecord?.answer || question.answerRecord?.internalMemo || question.rationale || "미답변 / 추가 확인 필요"
    })),
    risks: risks.length ? risks : ["원문 계약서 및 추가 DD 전까지 핵심 리스크 확인 필요"],
    followUps: asArray(followUps).slice(0, 3),
    finalOpinion: formatListItemText(summary.investmentMemo) || "본건은 추가 DD 결과 확인 후 투심 상정 여부 판단 필요",
    riskView: formatListItemText(summary.riskView) || "",
    nextActions: asArray(memo.nextActionItems).slice(0, 6),
    sourceChecks: asArray(memo.sourceVerificationItems).slice(0, 6)
  };
  return normalizeDocxMemoTone(data);
}

function deriveAssetNameForMemo(effectiveMeeting = {}, snapshot = {}, imAnalysis = {}, memo = {}) {
  const detected = imAnalysis.autoDetectedFields || {};
  const imSnapshot = imAnalysis.fundSnapshot || {};
  const reportOverview = memo.meetingOverview || memo.investmentOverview || {};
  const name = pickFirstMeaningful([
    effectiveMeeting.assetName,
    effectiveMeeting.fundName,
    snapshot.assetName,
    snapshot.projectName,
    snapshot.dealName,
    snapshot.loanName,
    snapshot.fundName,
    imSnapshot.assetName,
    imSnapshot.projectName,
    imSnapshot.dealName,
    imSnapshot.loanName,
    imSnapshot.fundName,
    detected.assetName,
    detected.projectName,
    detected.dealName,
    detected.loanName,
    detected.fundName,
    reportOverview.assetName,
    reportOverview.projectName,
    reportOverview.dealName,
    reportOverview.loanName
  ]);
  if (name) return assetNameForMemoTitle(name);
  const fallback = [
    effectiveMeeting.locationType,
    effectiveMeeting.sector,
    effectiveMeeting.strategy,
    effectiveMeeting.capitalType
  ].map(localizeDisplayTerm).filter(Boolean).join(" ");
  return assetNameForMemoTitle(fallback || "자산명 확인 필요");
}

function resolveReviewType(effectiveMeeting = {}, overviewFacts = {}, imAnalysis = {}) {
  const detected = imAnalysis.autoDetectedFields || {};
  const imSnapshot = imAnalysis.fundSnapshot || {};
  const explicit = normalizeReviewType(state.meeting?.dealType || effectiveMeeting.dealType || detected.dealType || imSnapshot.dealType);
  if (explicit) return explicit;
  const text = [
    state.meeting?.fundName,
    state.meeting?.keyConcerns,
    effectiveMeeting.fundName,
    effectiveMeeting.assetClass,
    effectiveMeeting.strategy,
    effectiveMeeting.capitalType,
    effectiveMeeting.investmentStructure,
    overviewFacts.sizeOrLoanAmount,
    overviewFacts.periodOrMaturity,
    overviewFacts.returnOrLoanRate,
    JSON.stringify(imAnalysis || {}),
    state.imProcessingResult?.textExcerpt
  ].filter(Boolean).join(" ");
  return inferReviewTypeFromText(text);
}

function normalizeReviewType(value) {
  const text = String(value || "").trim().toLowerCase();
  if (!text || text === "auto" || text === "자동" || text === "자동 판단") return "";
  if (/loan|debt|credit|대출|여신|pf|담보대출|프로젝트파이낸싱|facility/.test(text)) return "loan";
  if (/fund|펀드|출자|commitment|블라인드|blind|separate account|세컨더리|secondary|buyout|growth/.test(text)) return "fund";
  return "";
}

function inferReviewTypeFromText(value) {
  const text = String(value || "").toLowerCase();
  const loanScore = [
    /대출/, /pf/, /프로젝트\s*파이낸싱/, /담보대출/, /대출금리/, /대출만기/, /차주/, /대주/, /tr\.\s*a/,
    /loan/, /facility/, /senior debt/, /credit/, /borrower/, /lender/, /ltv/
  ].reduce((score, pattern) => score + (pattern.test(text) ? 1 : 0), 0);
  const fundScore = [
    /펀드/, /블라인드/, /출자/, /약정액/, /목표\s*수익률/, /투자\s*기간/, /net\s*irr/, /gross\s*irr/, /moic/,
    /fund/, /lp commitment/, /capital call/, /buyout/, /growth capital/, /venture/
  ].reduce((score, pattern) => score + (pattern.test(text) ? 1 : 0), 0);
  if (loanScore >= 2 && loanScore >= fundScore) return "loan";
  if (fundScore >= 2) return "fund";
  return "";
}

function normalizeOverviewFactsForReviewType(facts = {}, reviewType = "") {
  const normalized = {
    ...facts,
    sizeOrLoanAmount: selectMetricValueForReviewType(facts.sizeOrLoanAmount, reviewType, "size"),
    periodOrMaturity: selectMetricValueForReviewType(facts.periodOrMaturity, reviewType, "period"),
    returnOrLoanRate: selectMetricValueForReviewType(facts.returnOrLoanRate, reviewType, "return")
  };
  const userFacts = deriveUserProvidedDealFacts();
  const meeting = state.meeting || {};
  const userRegionAssetClass = [
    userFacts.region,
    meeting.assetClass,
    meeting.sector,
    meeting.strategy,
    meeting.capitalType
  ].filter(isMeaningfulFact).join(" / ");
  if (isMeaningfulFact(userRegionAssetClass)) normalized.regionAssetClass = userRegionAssetClass;
  if (reviewType === "loan") {
    if (isMeaningfulFact(userFacts.loanSize)) normalized.sizeOrLoanAmount = userFacts.loanSize;
    if (isMeaningfulFact(userFacts.loanMaturity)) normalized.periodOrMaturity = userFacts.loanMaturity;
    if (isMeaningfulFact(userFacts.loanRate)) normalized.returnOrLoanRate = userFacts.loanRate;
  }
  return normalized;
}

function buildOverviewMetricRows(facts = {}, reviewType = "") {
  if (reviewType === "loan") {
    return [
      ["대출 규모", facts.sizeOrLoanAmount || "확인 필요"],
      ["대출만기", facts.periodOrMaturity || "확인 필요"],
      ["대출금리", facts.returnOrLoanRate || "확인 필요"]
    ];
  }
  if (reviewType === "fund") {
    return [
      ["펀드 규모", facts.sizeOrLoanAmount || "확인 필요"],
      ["투자 기간", facts.periodOrMaturity || "확인 필요"],
      ["목표 수익률", facts.returnOrLoanRate || "확인 필요"]
    ];
  }
  return [
    ["펀드 규모 / 대출 규모", facts.sizeOrLoanAmount || "확인 필요"],
    ["투자 기간 / 대출만기", facts.periodOrMaturity || "확인 필요"],
    ["목표 수익률 / 대출금리", facts.returnOrLoanRate || "확인 필요"]
  ];
}

function selectMetricValueForReviewType(value, reviewType, metricKind) {
  const text = String(value || "").replace(/\s+/g, " ").trim();
  if (!text || !reviewType) return text;
  const parts = splitPairedMetricValue(text);
  const selected = reviewType === "loan" ? parts.loan : parts.fund;
  if (isMeaningfulMetricValue(selected)) return selected;
  if (parts.wasSplit) return selected || "확인 필요";
  if (reviewType === "fund" && metricLooksLoanOnly(text, metricKind)) return "확인 필요";
  if (reviewType === "loan" && metricLooksFundOnly(text, metricKind)) return "확인 필요";
  return text;
}

function splitPairedMetricValue(value) {
  const text = String(value || "").trim();
  const separators = /\s+\/\s+|\s*;\s*/;
  const parts = text.split(separators).map((item) => item.trim()).filter(Boolean);
  if (parts.length >= 2) return { fund: parts[0], loan: parts.slice(1).join(" / "), wasSplit: true };
  return { fund: text, loan: text, wasSplit: false };
}

function isMeaningfulMetricValue(value) {
  const text = String(value || "").trim();
  return Boolean(text) && !/^(확인\s*필요|미확인|n\/a|na|null|undefined|-|없음)$/i.test(text);
}

function metricLooksLoanOnly(value, metricKind) {
  const text = String(value || "");
  if (metricKind === "period") return /대출만기|loan maturity|tenor|최초\s*인출|만기/i.test(text) && !/투자\s*기간|holding/i.test(text);
  if (metricKind === "return") return /대출금리|loan rate|interest|coupon|tr\.\s*a|주택도시기금|금리/i.test(text) && !/irr|moic|목표\s*수익률|target return/i.test(text);
  return /대출\s*규모|loan size|facility|대출금액/i.test(text) && !/펀드\s*규모|fund size/i.test(text);
}

function metricLooksFundOnly(value, metricKind) {
  const text = String(value || "");
  if (metricKind === "period") return /투자\s*기간|holding period|fund life/i.test(text) && !/대출만기|loan maturity|최초\s*인출/i.test(text);
  if (metricKind === "return") return /irr|moic|목표\s*수익률|target return|net return/i.test(text) && !/대출금리|loan rate|interest|tr\.\s*a/i.test(text);
  return /펀드\s*규모|fund size|target size|모집\s*규모/i.test(text) && !/대출\s*규모|loan size|facility/i.test(text);
}

function deriveEffectiveMeetingInfo(memo = state.postMeetingMemo || {}) {
  const meeting = state.meeting || {};
  const briefSnapshot = state.preMeetingBrief?.fundSnapshot || {};
  const imAnalysis = state.imProcessingResult?.imAnalysis || {};
  const imSnapshot = imAnalysis.fundSnapshot || {};
  const detected = imAnalysis.autoDetectedFields || {};
  const reportOverview = memo.meetingOverview || memo.investmentOverview || {};
  const structuredSources = [meeting, reportOverview, briefSnapshot, imSnapshot, detected, imAnalysis, state.preMeetingBrief, memo].filter(Boolean);
  const textSource = [
    state.meetingNotes,
    state.transcript,
    state.meeting?.keyConcerns,
    state.imProcessingResult?.textExcerpt,
    JSON.stringify(getQuestionRecordsForReport()),
    JSON.stringify(state.preMeetingBrief || {}),
    JSON.stringify(imAnalysis || {}),
    JSON.stringify(memo || {})
  ].filter(Boolean).join("\n");

  const raw = {
    ...meeting,
    managerName: pickFirstMeaningful([
      meeting.managerName,
      briefSnapshot.managerName,
      imSnapshot.managerName,
      detected.managerName,
      findStructuredFact(structuredSources, ["managerName", "gpName", "manager", "sponsor", "운용사", "gp", "GP"]),
      findLabeledFact(textSource, ["운용사", "GP", "GP명", "자산운용사", "manager", "sponsor"])
    ]),
    fundName: pickFirstMeaningful([
      meeting.fundName,
      briefSnapshot.fundName,
      imSnapshot.fundName,
      detected.fundName,
      findStructuredFact(structuredSources, ["fundName", "dealName", "loanName", "projectName", "transactionName", "펀드명", "대출명", "거래명", "프로젝트명"]),
      findLabeledFact(textSource, ["펀드명", "대출명", "거래명", "프로젝트명", "건명", "fund name", "deal name", "loan name"])
    ]),
    locationType: normalizeLocationType(pickFirstMeaningful([
      meeting.locationType,
      reportOverview.region,
      briefSnapshot.region,
      imSnapshot.region,
      detected.region,
      findStructuredFact(structuredSources, ["locationType", "region", "investmentRegion", "투자지역", "지역"])
    ])),
    assetClass: normalizeAssetClass(pickFirstMeaningful([
      meeting.assetClass,
      briefSnapshot.assetClass,
      imSnapshot.assetClass,
      detected.assetClass,
      findStructuredFact(structuredSources, ["assetClass", "asset", "자산군", "자산분류"])
    ])),
    strategy: pickKnownOption("strategy", pickFirstMeaningful([
      meeting.strategy,
      briefSnapshot.strategy,
      imSnapshot.strategy,
      detected.strategy,
      findStructuredFact(structuredSources, ["strategy", "전략", "investmentStrategy"])
    ])),
    sector: buildSectorValue([
      meeting.sector,
      briefSnapshot.sector,
      imSnapshot.sector,
      detected.sector,
      findStructuredFact(structuredSources, ["sector", "섹터", "industry"])
    ], textSource),
    capitalType: normalizeCapitalType(pickFirstMeaningful([
      meeting.capitalType,
      briefSnapshot.capitalType,
      briefSnapshot.capitalStructure,
      imSnapshot.capitalType,
      imSnapshot.capitalStructure,
      detected.capitalType,
      detected.investmentStructure,
      findStructuredFact(structuredSources, ["capitalType", "capitalStructure", "equityDebt", "투자구조", "구조"])
    ])),
    dealType: normalizeReviewType(pickFirstMeaningful([
      meeting.dealType,
      reportOverview.dealType,
      reportOverview.reviewType,
      briefSnapshot.dealType,
      imSnapshot.dealType,
      detected.dealType,
      detected.reviewType,
      findStructuredFact(structuredSources, ["dealType", "reviewType", "transactionType", "검토유형", "거래유형"])
    ])),
    investmentStructure: pickFirstMeaningful([
      meeting.investmentStructure,
      briefSnapshot.investmentStructure,
      briefSnapshot.capitalStructure,
      imSnapshot.investmentStructure,
      imSnapshot.capitalStructure,
      detected.investmentStructure,
      findStructuredFact(structuredSources, ["investmentStructure", "capitalStructure", "structure", "투자구조", "상세투자구조"])
    ])
  };
  if (!raw.capitalType) raw.capitalType = inferCapitalType(raw.investmentStructure || "");
  if (!raw.dealType) raw.dealType = inferReviewTypeFromText(textSource);
  if (!raw.investmentStructure && raw.capitalType) {
    raw.investmentStructure = pickKnownCapitalStructure(raw.capitalType, raw.investmentStructure);
  }
  return raw;
}

function normalizeLocationType(value) {
  const text = localizeDisplayTerm(value);
  if (/국내|한국|korea|domestic/i.test(text)) return "국내";
  if (/해외|미국|유럽|일본|중국|global|overseas|international|us|usa|europe|japan|china/i.test(text)) return "해외";
  return text;
}

function normalizeAssetClass(value) {
  const text = localizeDisplayTerm(value);
  const normalizedText = normalizeFactKey(text);
  if (!normalizedText) return text;
  const options = Object.keys(ASSET_OPTIONS);
  return options.find((option) => normalizeFactKey(option) === normalizedText)
    || options.find((option) => normalizedText.includes(normalizeFactKey(option)) || normalizeFactKey(option).includes(normalizedText))
    || text;
}

function normalizeSectorValue(value) {
  const text = localizeDisplayTerm(value);
  if (!text) return "";
  const assetClass = normalizeAssetClass(state.meeting.assetClass || state.preMeetingBrief?.fundSnapshot?.assetClass || state.imProcessingResult?.imAnalysis?.autoDetectedFields?.assetClass || "");
  const options = ASSET_OPTIONS[assetClass]?.sectors || Object.values(ASSET_OPTIONS).flatMap((config) => config.sectors || []);
  const normalizedText = normalizeFactKey(text);
  if (!normalizedText) return text;
  const matches = options.filter((option) => {
    const key = normalizeFactKey(option);
    return key && (normalizedText.includes(key) || key.includes(normalizedText));
  });
  if (matches.length) return mergeTextLists(matches).join(", ");
  return text;
}

function buildSectorValue(candidates = [], sourceText = "") {
  const explicitSectors = candidates
    .filter(Boolean)
    .flatMap((value) => String(value).split(/\s*,\s*|\s*;\s*|\n+/))
    .map((value) => normalizeSectorValue(value))
    .filter(Boolean);
  const imUseSectors = extractAssetUseSectorsFromText(sourceText)
    .split(/\s*,\s*/)
    .filter(Boolean);
  return cleanSectorList([...explicitSectors, ...imUseSectors], sourceText).join(", ");
}

function extractAssetUseSectorsFromText(text = "") {
  const source = String(text || "");
  if (!source.trim()) return "";
  const sectors = [];
  if (/공동주택|아파트|민간임대|임대주택|주거|세대/i.test(source)) sectors.push("공동주택 / 주거");
  if (/오피스텔/i.test(source)) sectors.push("오피스텔");
  if (/판매시설|상업시설|리테일|상가/i.test(source)) sectors.push("판매시설 / 리테일");
  if (/근린생활시설|근생/i.test(source)) sectors.push("근린생활시설");
  if (/업무시설|오피스(?!텔)/i.test(source)) sectors.push("오피스");
  if (/물류센터|물류/i.test(source)) sectors.push("물류센터");
  if (/데이터센터/i.test(source)) sectors.push("데이터센터");
  if (/호텔|호스피탈리티/i.test(source)) sectors.push("호텔 / 호스피탈리티");
  if (/바이오|헬스케어/i.test(source)) sectors.push("헬스케어 / 바이오");
  if (/테크\s*\/\s*소프트웨어|소프트웨어|SaaS|IT\s*서비스|테크 기업/i.test(source)) sectors.push("테크 / 소프트웨어");
  if (/이커머스|소비재|커머스/i.test(source)) sectors.push("소비재 / 이커머스");
  return mergeTextLists(sectors).join(", ");
}

function cleanSectorList(sectors = [], sourceText = "") {
  const source = String(sourceText || "");
  const list = mergeTextLists(sectors)
    .map((sector) => String(sector).trim())
    .filter(Boolean);
  const hasDetailedResidential = list.some((sector) => /공동주택\s*\/\s*주거/i.test(sector));
  const hasDetailedRetail = list.some((sector) => /판매시설\s*\/\s*리테일/i.test(sector));
  const hasOfficeBuilding = /업무시설|오피스(?!텔)/i.test(source);
  const realEstateContext = /부동산|PF|공동주택|오피스텔|판매시설|근린생활시설|민간임대|분양|주택/i.test(source);
  const filtered = list.filter((sector) => {
    const normalized = normalizeFactKey(sector);
    if (hasDetailedResidential && /^(공동주택|주거|주거멀티패밀리)$/.test(normalized)) return false;
    if (hasDetailedRetail && /^(판매시설|리테일)$/.test(normalized)) return false;
    if (normalized === normalizeFactKey("오피스") && list.some((item) => normalizeFactKey(item) === normalizeFactKey("오피스텔")) && !hasOfficeBuilding) return false;
    if (realEstateContext && ["테크 / 소프트웨어", "헬스케어 / 바이오", "소비재 / 이커머스"].some((item) => normalizeFactKey(item) === normalized)) return false;
    return true;
  });
  const priority = [
    "공동주택 / 주거",
    "주거 / 멀티패밀리",
    "오피스텔",
    "판매시설 / 리테일",
    "리테일",
    "근린생활시설",
    "오피스",
    "물류센터",
    "데이터센터",
    "호텔 / 호스피탈리티",
    "테크 / 소프트웨어",
    "헬스케어 / 바이오",
    "소비재 / 이커머스"
  ].map(normalizeFactKey);
  return filtered
    .map((sector, index) => ({ sector, index, priority: priority.indexOf(normalizeFactKey(sector)) }))
    .sort((a, b) => {
      const rankA = a.priority === -1 ? 999 : a.priority;
      const rankB = b.priority === -1 ? 999 : b.priority;
      return rankA - rankB || a.index - b.index;
    })
    .map((item) => item.sector);
}

function normalizeCapitalType(value) {
  const text = localizeDisplayTerm(value);
  if (/equity|지분|소수지분|출자|common|preferred|rcps|cps/i.test(text)) return "Equity";
  if (/debt|대출|loan|senior|subordinated|unitranche|bond|credit/i.test(text)) return "Debt";
  if (/hybrid|mezzanine|메자닌|혼합|convertible|전환/i.test(text)) return "Hybrid / Mezzanine";
  if (/펀드형|fund/i.test(text) && /사모투자|PE|Private Equity/i.test(state.meeting.assetClass || state.preMeetingBrief?.fundSnapshot?.assetClass || "")) return "Equity";
  return text;
}

function pickKnownOption(kind, value) {
  const text = String(value || "").trim();
  if (!text) return "";
  const normalizedText = normalizeFactKey(text);
  if (!normalizedText) return text;
  const assetClass = normalizeAssetClass(state.meeting.assetClass || state.preMeetingBrief?.fundSnapshot?.assetClass || state.imProcessingResult?.imAnalysis?.autoDetectedFields?.assetClass || "");
  const options = ASSET_OPTIONS[assetClass]?.[kind === "strategy" ? "strategies" : "sectors"] || [];
  return options.find((option) => normalizeFactKey(option) === normalizedText)
    || options.find((option) => normalizedText.includes(normalizeFactKey(option)) || normalizeFactKey(option).includes(normalizedText))
    || text;
}

function pickKnownCapitalStructure(capitalType, value) {
  const options = CAPITAL_OPTIONS[capitalType] || [];
  const text = String(value || "").trim();
  const normalizedText = normalizeFactKey(text);
  if (!normalizedText) return text;
  return options.find((option) => normalizeFactKey(option) === normalizedText)
    || options.find((option) => normalizedText.includes(normalizeFactKey(option)) || normalizeFactKey(option).includes(normalizedText))
    || text;
}

function deriveMeetingOverviewFacts(memo = {}, snapshot = {}, imAnalysis = {}) {
  const reportOverview = memo.meetingOverview || memo.investmentOverview || {};
  const imSnapshot = imAnalysis.fundSnapshot || {};
  const detected = imAnalysis.autoDetectedFields || {};
  const structuredSources = [reportOverview, snapshot, imSnapshot, detected, state.meeting, memo].filter(Boolean);
  const textSource = [
    state.meeting.keyConcerns,
    state.meetingNotes,
    state.transcript,
    state.imProcessingResult?.textExcerpt,
    JSON.stringify(getQuestionRecordsForReport()),
    JSON.stringify(reportOverview),
    JSON.stringify(snapshot),
    JSON.stringify(imAnalysis),
    JSON.stringify(memo.issueBasedMeetingNotes || []),
    JSON.stringify(memo.sourceVerificationItems || [])
  ].filter(Boolean).join("\n");

  const region = pickFirstMeaningful([
    reportOverview.region,
    state.meeting.locationType,
    snapshot.region,
    detected.region,
    findStructuredFact(structuredSources, ["region", "locationType", "투자지역", "지역"])
  ]);
  const assetClass = pickFirstMeaningful([
    state.meeting.assetClass,
    snapshot.assetClass,
    detected.assetClass,
    findStructuredFact(structuredSources, ["assetClass", "자산군", "asset"])
  ]);
  const sector = buildSectorValue([
    state.meeting.sector,
    snapshot.sector,
    detected.sector,
    findStructuredFact(structuredSources, ["sector", "섹터"])
  ], textSource);
  const strategy = pickFirstMeaningful([
    state.meeting.strategy,
    snapshot.strategy,
    detected.strategy,
    findStructuredFact(structuredSources, ["strategy", "전략"])
  ]);
  const capitalType = pickFirstMeaningful([
    state.meeting.capitalType,
    state.meeting.investmentStructure,
    snapshot.capitalStructure,
    snapshot.capitalType,
    detected.capitalType,
    detected.investmentStructure,
    findStructuredFact(structuredSources, ["capitalStructure", "capitalType", "investmentStructure", "투자구조", "구조"])
  ]);

  const regionAssetClass = pickFirstMeaningful([
    reportOverview.regionAssetClass,
    reportOverview["지역 / 자산군"],
    reportOverview["지역 / 분류"],
    [region, assetClass, sector, strategy, capitalType].map(localizeDisplayTerm).filter(Boolean).join(" / ")
  ]);

  const sizeOrLoanAmount = pickFirstMeaningful([
    reportOverview.sizeOrLoanAmount,
    reportOverview["펀드 규모 / 대출 규모"],
    findStructuredFact(structuredSources, [
      "sizeOrLoanAmount", "targetSize", "fundSize", "loanSize", "debtSize", "facilitySize",
      "financingAmount", "commitmentTarget", "offeringSize", "펀드규모", "대출규모", "모집규모", "대출금액"
    ]),
    findLabeledFact(textSource, ["펀드 규모", "대출 규모", "모집 규모", "목표 규모", "대출금액", "약정 총액", "facility size", "loan amount", "fund size"])
  ]);
  const periodOrMaturity = pickFirstMeaningful([
    reportOverview.periodOrMaturity,
    reportOverview["투자 기간 / 대출만기"],
    findStructuredFact(structuredSources, [
      "periodOrMaturity", "investmentPeriod", "loanMaturity", "maturity", "tenor",
      "term", "duration", "holdingPeriod", "투자기간", "대출만기", "만기"
    ]),
    findLabeledFact(textSource, ["투자 기간", "대출만기", "대출 만기", "만기", "투자기간", "tenor", "maturity", "term", "duration"])
  ]);
  const returnOrLoanRate = pickFirstMeaningful([
    reportOverview.returnOrLoanRate,
    reportOverview["목표 수익률 / 대출금리"],
    findStructuredFact(structuredSources, [
      "returnOrLoanRate", "targetReturn", "targetIrr", "netIrr", "grossIrr", "moic",
      "returns", "loanRate", "interestRate", "coupon", "margin", "spread", "목표수익률", "대출금리", "금리"
    ]),
    findLabeledFact(textSource, ["목표 수익률", "대출금리", "대출 금리", "금리", "Net IRR", "Gross IRR", "MOIC", "target return", "interest rate", "coupon", "margin", "spread"])
  ]);
  const commitmentAmount = pickConfirmedCommitmentAmount([
    reportOverview.commitmentAmount,
    reportOverview["당사 검토 약정액"],
    findStructuredFact(structuredSources, [
      "commitmentAmount", "lpCommitment", "proposedCommitment", "ourCommitment",
      "reviewCommitment", "당사검토약정액", "출자검토액"
    ]),
    findLabeledFact(textSource, ["당사 검토 약정액", "검토 약정액", "출자 검토액", "당사 약정액", "LP commitment", "our commitment", "proposed commitment"])
  ], textSource, sizeOrLoanAmount);

  return {
    location: pickFirstMeaningful([
      reportOverview.location,
      reportOverview.venue,
      reportOverview.place,
      reportOverview["장소"],
      findStructuredFact(structuredSources, ["location", "venue", "place", "장소"])
    ]),
    regionAssetClass,
    sizeOrLoanAmount,
    periodOrMaturity,
    returnOrLoanRate,
    commitmentAmount
  };
}

function pickConfirmedCommitmentAmount(values, sourceText = "", sizeOrLoanAmount = "") {
  const value = pickFirstMeaningful(values);
  if (!value) return "";
  const source = String(sourceText || "");
  const explicitCommitment = /당사\s*검토\s*약정액|검토\s*약정액|출자\s*검토액|당사\s*약정액|본\s*LP\s*약정|LP\s*commitment|our\s*commitment|proposed\s*commitment/i.test(source)
    || (/(당사|본\s*LP|our|LP)/i.test(value) && /(약정|commitment|출자|투자액)/i.test(value));
  if (!explicitCommitment && amountsLookSame(value, sizeOrLoanAmount)) return "";
  if (!explicitCommitment && /(총\s*)?(펀드|대출|모집|조달|facility|loan)\s*(규모|금액|amount|size)/i.test(value)) return "";
  return value;
}

function amountsLookSame(left, right) {
  const leftKey = normalizeAmountKey(left);
  const rightKey = normalizeAmountKey(right);
  return Boolean(leftKey && rightKey && leftKey === rightKey);
}

function normalizeAmountKey(value) {
  const text = String(value || "");
  const match = text.match(/(\d{1,3}(?:,\d{3})+|\d+(?:\.\d+)?)\s*(조|억|만|원|bn|billion|m|million)?/i);
  if (!match) return "";
  return `${match[1].replace(/,/g, "")}${String(match[2] || "").toLowerCase()}`;
}

function pickFirstMeaningful(values) {
  for (const value of values.flat()) {
    const cleaned = cleanFactValue(value);
    if (isMeaningfulFact(cleaned)) return cleaned;
  }
  return "";
}

function cleanFactValue(value) {
  if (value === undefined || value === null) return "";
  if (Array.isArray(value)) return value.map(cleanFactValue).filter(isMeaningfulFact).join(", ");
  if (typeof value === "object") return formatListItemText(value);
  return String(value)
    .replace(/^["'`]+|["'`,.]+$/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function isMeaningfulFact(value) {
  const text = String(value || "").trim();
  return Boolean(text) && !/^(확인\s*필요|미확인|n\/a|na|null|undefined|-|없음)$/i.test(text);
}

function normalizeFactKey(key) {
  return String(key || "").toLowerCase().replace(/[^a-z0-9가-힣]/g, "");
}

function findStructuredFact(sources, keys) {
  const normalizedKeys = new Set(keys.map(normalizeFactKey));
  for (const source of asArray(sources)) {
    const value = findStructuredFactInValue(source, normalizedKeys);
    if (isMeaningfulFact(value)) return value;
  }
  return "";
}

function findStructuredFactInValue(value, normalizedKeys, seen = new Set()) {
  if (!value || typeof value !== "object") return "";
  if (seen.has(value)) return "";
  seen.add(value);
  if (Array.isArray(value)) {
    for (const item of value) {
      const found = findStructuredFactInValue(item, normalizedKeys, seen);
      if (isMeaningfulFact(found)) return found;
    }
    return "";
  }
  for (const [key, item] of Object.entries(value)) {
    if (normalizedKeys.has(normalizeFactKey(key))) {
      const cleaned = cleanFactValue(item);
      if (isMeaningfulFact(cleaned)) return cleaned;
    }
  }
  for (const item of Object.values(value)) {
    const found = findStructuredFactInValue(item, normalizedKeys, seen);
    if (isMeaningfulFact(found)) return found;
  }
  return "";
}

function findLabeledFact(text, labels) {
  const source = String(text || "").replace(/\r/g, "\n");
  for (const label of labels) {
    const escaped = escapeRegExp(label).replace(/\\\s+/g, "\\s*");
    const pattern = new RegExp(`${escaped}\\s*(?:은|는|:|：|=|-)?\\s*([^\\n;]{2,120})`, "i");
    const match = source.match(pattern);
    if (match) {
      const value = cleanFactValue(match[1].replace(/[}"\]]+$/g, ""));
      if (isMeaningfulFact(value)) return value;
    }
  }
  return "";
}

function escapeRegExp(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function localizeDisplayTerm(value) {
  const text = String(value || "").trim();
  if (!text) return "";
  const exact = DISPLAY_TERM_MAP[text.toLowerCase()];
  if (exact) return exact;
  let localized = text;
  DISPLAY_TERM_REPLACEMENTS.forEach(([pattern, replacement]) => {
    localized = localized.replace(pattern, replacement);
  });
  return localized;
}

const DISPLAY_TERM_MAP = {
  "domestic": "국내",
  "domestic korea": "국내",
  "korea": "국내",
  "kr": "국내",
  "overseas": "해외",
  "global": "해외",
  "international": "해외",
  "real estate": "부동산",
  "infrastructure": "인프라",
  "private equity": "사모투자(PE)",
  "pe": "사모투자(PE)",
  "private debt": "사모투자(PD)",
  "pd": "사모투자(PD)",
  "commodity finance": "상품금융",
  "development / pf": "개발 / PF",
  "development/pf": "개발 / PF",
  "residential / multifamily": "주거 / 멀티패밀리",
  "residential/multifamily": "주거 / 멀티패밀리",
  "senior debt": "선순위 대출",
  "subordinated debt": "후순위 대출",
  "guarantee / contract": "보증 / 계약",
  "exit / repayment": "Exit / 상환",
  "construction / completion": "공사 / 준공",
  "business plan": "사업성",
  "stakeholders": "이해관계자",
  "partially answered": "부분답변",
  "needs follow-up material": "추가자료 필요",
  "unconfirmed": "미확인",
  "equity": "지분",
  "debt": "대출",
  "hybrid / mezzanine": "혼합 / 메자닌"
};

const DISPLAY_TERM_REPLACEMENTS = [
  [/\bDomestic Korea\b/gi, "국내"],
  [/\bDomestic\b/gi, "국내"],
  [/\bKorea\b/gi, "국내"],
  [/\bOverseas\b/gi, "해외"],
  [/\bGlobal\b/gi, "해외"],
  [/\bInternational\b/gi, "해외"],
  [/\bReal Estate\b/gi, "부동산"],
  [/\bInfrastructure\b/gi, "인프라"],
  [/\bPrivate Equity\b/gi, "사모투자(PE)"],
  [/\bPrivate Debt\b/gi, "사모투자(PD)"],
  [/\bCommodity Finance\b/gi, "상품금융"],
  [/\bDevelopment\s*\/\s*PF\b/gi, "개발 / PF"],
  [/\bResidential\s*\/\s*Multifamily\b/gi, "주거 / 멀티패밀리"],
  [/\bSenior Debt\b/gi, "선순위 대출"],
  [/\bSubordinated Debt\b/gi, "후순위 대출"],
  [/\bGuarantee\s*\/\s*contract\b/gi, "보증 / 계약"],
  [/\bExit\s*\/\s*repayment\b/gi, "Exit / 상환"],
  [/\bConstruction\s*\/\s*completion\b/gi, "공사 / 준공"],
  [/\bBusiness plan\b/gi, "사업성"],
  [/\bStakeholders\b/gi, "이해관계자"],
  [/\bPartially answered\b/gi, "부분답변"],
  [/\bNeeds follow-up material\b/gi, "추가자료 필요"],
  [/\bUnconfirmed\b/gi, "미확인"],
  [/\btrigger conditions\b/gi, "발동 조건"],
  [/\bexclusions\b/gi, "제외 조항"],
  [/\btake-out loan\b/gi, "Take-out 대출"],
  [/\btake-out commitment\b/gi, "Take-out 확약"],
  [/\blease-up\b/gi, "임대율"],
  [/\bstress case\b/gi, "스트레스 케이스"],
  [/\bfollow-up package\b/gi, "후속자료 패키지"],
  [/\blegal review package\b/gi, "법률 검토 패키지"],
  [/\bconstruction restart\b/gi, "공사 재개"],
  [/\bcost contingency\b/gi, "비용 예비비"],
  [/\boriginal agreements\b/gi, "원문 계약서"],
  [/\bHybrid\s*\/\s*Mezzanine\b/gi, "혼합 / 메자닌"],
  [/\bEquity\b/gi, "지분"],
  [/\bDebt\b/gi, "대출"]
];

function deriveExecutiveEvidence(memo) {
  const reasons = [];
  const snapshot = state.preMeetingBrief?.fundSnapshot || {};
  const effectiveMeeting = deriveEffectiveMeetingInfo(memo);
  if (effectiveMeeting.assetClass || effectiveMeeting.strategy) {
    reasons.push(`${localizeDisplayTerm(effectiveMeeting.assetClass) || "자산군"} / ${localizeDisplayTerm(effectiveMeeting.strategy) || "전략"} 검토 건으로 분류됨`);
  }
  if (snapshot.strategy || snapshot.region || snapshot.assetClass) {
    reasons.push(`IM 기준 투자대상 분류: ${[snapshot.region, snapshot.assetClass, snapshot.strategy].map(localizeDisplayTerm).filter(Boolean).join(" · ")}`);
  }
  if (memo.internalReportSummary?.riskView) {
    reasons.push(formatListItemText(memo.internalReportSummary.riskView));
  }
  asArray(state.preMeetingBrief?.redFlags).slice(0, 2).forEach((item) => reasons.push(formatListItemText(item)));
  return reasons.slice(0, 3).length ? reasons.slice(0, 3) : ["사전 분석 및 미팅 답변 기준으로 추가 DD 필요 항목 확인"];
}

function normalizeDocxMemoTone(value) {
  if (Array.isArray(value)) return value.map((item) => normalizeDocxMemoTone(item));
  if (value && typeof value === "object") {
    return Object.fromEntries(Object.entries(value).map(([key, item]) => [key, normalizeDocxMemoTone(item)]));
  }
  if (typeof value !== "string") return value;
  return normalizeReportLanguage(value);
}

function normalizeReportLanguage(value) {
  return toEumseumStyle(localizeDisplayTerm(value));
}

function toEumseumStyle(value) {
  let text = String(value || "").trim();
  if (!text) return text;
  const replacements = [
    [/공유해 주실 수 있습니까\?/g, "공유 가능한지?"],
    [/확인할 수 있습니까\?/g, "확인 가능한지?"],
    [/설명해 주실 수 있습니까\?/g, "설명 가능한지?"],
    [/할 수 있습니까\?/g, "가능한지?"],
    [/제시해 주십시오/g, "제시 요청"],
    [/설명해 주십시오/g, "설명 요청"],
    [/공유해 주십시오/g, "공유 요청"],
    [/필요합니다(?=\.|,|\n|$)/g, "필요"],
    [/필요합니다/g, "필요함"],
    [/확인해야 합니다(?=\.|,|\n|$)/g, "확인 필요"],
    [/검토해야 합니다(?=\.|,|\n|$)/g, "검토 필요"],
    [/해야 합니다(?=\.|,|\n|$)/g, "필요"],
    [/하였습니다(?=\.|,|\n|$)/g, "하였음"],
    [/했습니다(?=\.|,|\n|$)/g, "했음"],
    [/합니다(?=\.|,|\n|$)/g, "함"],
    [/되었습니다(?=\.|,|\n|$)/g, "되었음"],
    [/됩니다(?=\.|,|\n|$)/g, "됨"],
    [/있습니다(?=\.|,|\n|$)/g, "있음"],
    [/없습니다(?=\.|,|\n|$)/g, "없음"],
    [/입니다(?=\.|,|\n|$)/g, "임"],
    [/입니다/g, "임"],
    [/합니다/g, "함"],
    [/됩니다/g, "됨"]
  ];
  replacements.forEach(([pattern, replacement]) => {
    text = text.replace(pattern, replacement);
  });
  return text;
}

function limitSentences(value, limit = 5) {
  const text = toEumseumStyle(value);
  const sentences = text.match(/[^.!?\n]+[.!?]?/g)?.map((item) => item.trim()).filter(Boolean) || [];
  return (sentences.length ? sentences.slice(0, limit).join(" ") : text).trim();
}

function formatKoreanDate(value) {
  if (!value) return "확인 필요";
  const date = new Date(`${value}T00:00:00`);
  if (Number.isNaN(date.getTime())) return value;
  const weekdays = ["일", "월", "화", "수", "목", "금", "토"];
  return `${date.getFullYear()}년 ${date.getMonth() + 1}월 ${date.getDate()}일(${weekdays[date.getDay()]})`;
}

function deriveRiskRows(data) {
  const rows = asArray(data.risks).map((item, index) => {
    const text = formatListItemText(item);
    const category = inferRiskCategory(text, index);
    const severity = index < 2 ? "High" : index < 5 ? "Medium" : "Low";
    return {
      category,
      description: text,
      severity,
      view: inferRiskView(text)
    };
  });
  return rows.length ? rows : [{
    category: "확인 필요",
    description: "주요 리스크 확인 필요",
    severity: "High",
    view: "추가 DD 전 투자 판단 보류 필요"
  }];
}

function inferRiskCategory(text, index) {
  if (/금리|DSCR|이자|금융비용/i.test(text)) return "금리 리스크";
  if (/Exit|매각|분양|상환|유동성/i.test(text)) return "Exit / 상환";
  if (/Track|트랙|운용|GP|실적/i.test(text)) return "트랙레코드";
  if (/공사|준공|시공|건설/i.test(text)) return "공사 / 준공";
  if (/보증|HUG|담보|계약/i.test(text)) return "계약 / 보증";
  return ["사업성", "현금흐름", "Alignment", "Key Man", "법무"][index] || "기타";
}

function inferRiskView(text) {
  if (/확인|미확인|검증|수령|자료/i.test(text)) return "원문 자료 수령 전 투자 판단 보류 필요";
  if (/금리|DSCR|이자|금융비용/i.test(text)) return "Stress 시나리오 재검토 요청";
  if (/Exit|매각|분양|상환/i.test(text)) return "상환 / Exit 경로 구체화 필요";
  if (/보증|HUG|담보/i.test(text)) return "계약 조건 및 실행 요건 원문 대조 필요";
  return "추가 DDQ 회신 후 재검토 필요";
}

function stripFollowUpMetadata(text) {
  return String(text || "")
    .replace(/\s*(담당|owner|기한|due|상태|status)\s*[:：]\s*[^/|,\n;]+/gi, "")
    .replace(/\s*\/\s*(담당|owner|기한|due|상태|status)\s*[:：]?\s*[^/|,\n;]+/gi, "")
    .replace(/\s{2,}/g, " ")
    .replace(/\s+([,.;])/g, "$1")
    .trim();
}

function deriveFollowUpRows(data) {
  const primaryItems = asArray(data.followUps).filter(Boolean);
  const fallbackItems = [
    ...asArray(data.nextActions),
    ...asArray(data.sourceChecks).map((item) => `원문 대조: ${formatListItemText(item)}`)
  ].filter(Boolean);
  const items = (primaryItems.length ? primaryItems : fallbackItems).slice(0, 3);
  const rows = items.map((item) => ({
    item: stripFollowUpMetadata(formatListItemText(item)),
    owner: "",
    due: ""
  })).filter((row) => row.item);
  return rows.length ? rows : [{
    item: "원문 계약서 및 추가 DD 자료 확인",
    owner: "",
    due: ""
  }];
}

function inferFollowUpOwner(text) {
  if (/내부|당사|자체|법무법인|준법|결재/i.test(text)) return "당사";
  if (/협의|조건|약정|조항/i.test(text)) return "양측 협의";
  return "GP";
}

function buildFullMarkdown() {
  const effectiveMeeting = deriveEffectiveMeetingInfo();
  return `# LP Meeting Copilot

## 세팅값
${Object.entries({ ...state.meeting, ...effectiveMeeting }).map(([key, value]) => `- ${labelize(key)}: ${value || "확인 필요"}`).join("\n")}

## 사전 브리프
${state.preMeetingBrief ? JSON.stringify(state.preMeetingBrief, null, 2) : "아직 생성되지 않았습니다."}

## 미팅 중 기록
${state.questionRecords.map((q, index) => `### Q${index + 1}. ${q.question}
- 운용사 답변: ${q.answerRecord.answer || "미입력"}
- 내부 메모: ${q.answerRecord.internalMemo || "미입력"}
- 추가 검토 필요: ${q.answerRecord.followUpNeeded ? "예" : "아니오"}`).join("\n\n")}

## 자유 메모
${state.meetingNotes || "미입력"}

## Transcript
${state.transcript || "미입력"}

## 최종 보고서
${buildReportMarkdown() || "아직 생성되지 않았습니다."}
`;
}

function exportMarkdown() {
  const content = buildFullMarkdown();
  const effectiveMeeting = deriveEffectiveMeetingInfo();
  const blob = new Blob([content], { type: "text/markdown;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = `${safeFileName(effectiveMeeting.managerName || "meeting")}-${safeFileName(effectiveMeeting.fundName || "copilot")}.md`;
  anchor.click();
  URL.revokeObjectURL(url);
}

async function exportReportDocx() {
  if (!state.postMeetingMemo) {
    toast("먼저 최종 보고서를 생성해주세요.");
    return;
  }
  try {
    const data = buildDocxMemoData();
    const blob = await createDocxBlob(data);
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = `${safeFileName(data.managerName)}-${safeFileName(data.fundName)}-review-memo.docx`;
    anchor.click();
    URL.revokeObjectURL(url);
    toast("DOCX 보고서를 내보냈습니다.");
  } catch (error) {
    toast(`DOCX 생성 실패: ${error.message}`);
  }
}

async function createDocxBlob(data) {
  const zip = await createDocxFromFinalTemplate(data);
  return new Blob([zip], {
    type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
  });
}

async function createDocxFromFinalTemplate(data) {
  try {
    const templateBytes = await loadFinalDocxTemplateBytes();
    const files = await readZipArchive(templateBytes);
    const decoder = new TextDecoder();
    const encoder = new TextEncoder();
    const documentXmlText = decoder.decode(files["word/document.xml"]);
    files["word/document.xml"] = encoder.encode(applyFinalWantedTemplate(documentXmlText, data));
    return createZipArchive(files);
  } catch (error) {
    console.warn("DOCX template fetch failed. Falling back to generated DOCX.", error);
    return createDocxFromGeneratedDocument(data);
  }
}

async function loadFinalDocxTemplateBytes() {
  try {
    const response = await fetch(DOCX_FINAL_TEMPLATE_URL);
    if (!response.ok) throw new Error(`템플릿 응답 오류: ${response.status}`);
    return new Uint8Array(await response.arrayBuffer());
  } catch (error) {
    throw new Error(`DOCX_FINAL_WANTED.docx 템플릿을 불러오지 못했습니다. ${error.message || error}`);
  }
}

async function createDocxFromGeneratedDocument(data) {
  return createZipArchive({
    "[Content_Types].xml": contentTypesXml(),
    "_rels/.rels": packageRelsXml(),
    "word/document.xml": documentXml(data),
    "word/styles.xml": stylesXml(),
    "word/_rels/document.xml.rels": documentRelsXml()
  });
}

async function copyText(text, message) {
  if (!text) return toast("복사할 내용이 없습니다.");
  await navigator.clipboard.writeText(text);
  toast(message);
}

async function createZipArchive(files) {
  const encoder = new TextEncoder();
  const entries = Object.entries(files).map(([name, content]) => {
    const data = typeof content === "string" ? encoder.encode(content) : content;
    const crc = crc32(data);
    return { name, data, crc };
  });
  let offset = 0;
  const localParts = [];
  const centralParts = [];
  entries.forEach((entry) => {
    const nameBytes = encoder.encode(entry.name);
    const localHeader = new Uint8Array(30 + nameBytes.length);
    const localView = new DataView(localHeader.buffer);
    localView.setUint32(0, 0x04034b50, true);
    localView.setUint16(4, 20, true);
    localView.setUint16(6, 0, true);
    localView.setUint16(8, 0, true);
    localView.setUint16(10, 0, true);
    localView.setUint16(12, 0, true);
    localView.setUint32(14, entry.crc, true);
    localView.setUint32(18, entry.data.length, true);
    localView.setUint32(22, entry.data.length, true);
    localView.setUint16(26, nameBytes.length, true);
    localHeader.set(nameBytes, 30);
    localParts.push(localHeader, entry.data);

    const centralHeader = new Uint8Array(46 + nameBytes.length);
    const centralView = new DataView(centralHeader.buffer);
    centralView.setUint32(0, 0x02014b50, true);
    centralView.setUint16(4, 20, true);
    centralView.setUint16(6, 20, true);
    centralView.setUint16(8, 0, true);
    centralView.setUint16(10, 0, true);
    centralView.setUint16(12, 0, true);
    centralView.setUint16(14, 0, true);
    centralView.setUint32(16, entry.crc, true);
    centralView.setUint32(20, entry.data.length, true);
    centralView.setUint32(24, entry.data.length, true);
    centralView.setUint16(28, nameBytes.length, true);
    centralView.setUint32(42, offset, true);
    centralHeader.set(nameBytes, 46);
    centralParts.push(centralHeader);
    offset += localHeader.length + entry.data.length;
  });
  const centralSize = centralParts.reduce((sum, part) => sum + part.length, 0);
  const end = new Uint8Array(22);
  const endView = new DataView(end.buffer);
  endView.setUint32(0, 0x06054b50, true);
  endView.setUint16(8, entries.length, true);
  endView.setUint16(10, entries.length, true);
  endView.setUint32(12, centralSize, true);
  endView.setUint32(16, offset, true);
  return concatUint8Arrays([...localParts, ...centralParts, end]);
}

async function readZipArchive(bytes) {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const decoder = new TextDecoder();
  const eocdOffset = findEndOfCentralDirectory(view);
  const entryCount = view.getUint16(eocdOffset + 10, true);
  const centralOffset = view.getUint32(eocdOffset + 16, true);
  const files = {};
  let offset = centralOffset;
  for (let index = 0; index < entryCount; index += 1) {
    if (view.getUint32(offset, true) !== 0x02014b50) throw new Error("DOCX 중앙 디렉터리를 읽지 못했습니다.");
    const method = view.getUint16(offset + 10, true);
    const compressedSize = view.getUint32(offset + 20, true);
    const fileNameLength = view.getUint16(offset + 28, true);
    const extraLength = view.getUint16(offset + 30, true);
    const commentLength = view.getUint16(offset + 32, true);
    const localOffset = view.getUint32(offset + 42, true);
    const nameBytes = bytes.slice(offset + 46, offset + 46 + fileNameLength);
    const name = decoder.decode(nameBytes);

    if (view.getUint32(localOffset, true) !== 0x04034b50) throw new Error("DOCX 로컬 파일 헤더를 읽지 못했습니다.");
    const localNameLength = view.getUint16(localOffset + 26, true);
    const localExtraLength = view.getUint16(localOffset + 28, true);
    const dataStart = localOffset + 30 + localNameLength + localExtraLength;
    const compressed = bytes.slice(dataStart, dataStart + compressedSize);
    files[name] = method === 0 ? compressed : await inflateRawBytes(compressed, method);
    offset += 46 + fileNameLength + extraLength + commentLength;
  }
  return files;
}

function findEndOfCentralDirectory(view) {
  const minOffset = Math.max(0, view.byteLength - 65558);
  for (let offset = view.byteLength - 22; offset >= minOffset; offset -= 1) {
    if (view.getUint32(offset, true) === 0x06054b50) return offset;
  }
  throw new Error("DOCX 파일 구조를 읽지 못했습니다.");
}

async function inflateRawBytes(bytes, method) {
  if (method !== 8) throw new Error(`지원하지 않는 DOCX 압축 방식입니다: ${method}`);
  if (typeof require === "function") {
    const zlib = require("zlib");
    return new Uint8Array(zlib.inflateRawSync(Buffer.from(bytes)));
  }
  if (typeof DecompressionStream !== "undefined") {
    const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
    return new Uint8Array(await new Response(stream).arrayBuffer());
  }
  throw new Error("현재 브라우저에서 DOCX 압축 해제를 지원하지 않습니다.");
}

function concatUint8Arrays(parts) {
  const total = parts.reduce((sum, part) => sum + part.length, 0);
  const merged = new Uint8Array(total);
  let offset = 0;
  parts.forEach((part) => {
    merged.set(part, offset);
    offset += part.length;
  });
  return merged;
}

function crc32(bytes) {
  const table = crc32.table || (crc32.table = Array.from({ length: 256 }, (_, index) => {
    let value = index;
    for (let bit = 0; bit < 8; bit += 1) value = value & 1 ? 0xedb88320 ^ (value >>> 1) : value >>> 1;
    return value >>> 0;
  }));
  let crc = 0xffffffff;
  for (let index = 0; index < bytes.length; index += 1) crc = table[(crc ^ bytes[index]) & 0xff] ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}

function contentTypesXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`;
}

function packageRelsXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;
}

function documentRelsXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;
}

function stylesXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Malgun Gothic"/><w:color w:val="1E293B"/><w:sz w:val="19"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Title">
    <w:name w:val="Title"/>
    <w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
    <w:rPr><w:b/><w:color w:val="0F172A"/><w:sz w:val="32"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="SmallMuted">
    <w:name w:val="SmallMuted"/>
    <w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Malgun Gothic"/><w:color w:val="64748B"/><w:sz w:val="17"/></w:rPr>
  </w:style>
</w:styles>`;
}

function applyFinalWantedTemplate(templateXml, data) {
  let xml = templateXml;
  const assetName = assetNameForMemoTitle(data.assetName || data.fundName);
  xml = replaceAssetNamePlaceholder(xml, assetName);
  xml = replaceTemplateLiteral(xml, "{{자산명}} 검토 메모", `${assetName} 검토 메모`);
  xml = replaceTemplateLiteral(xml, "본 메모는 투자 결정을 위한 내부 검토 자료이며, 외부 유출이 엄격히 금지됩니다.", "본 메모는 투자 결정을 위한 내부 검토 자료이며, 외부 유출이 엄격히 금지됨.");

  const tables = getXmlBlocks(xml, "tbl");
  if (tables.length < 4) throw new Error("DOCX_FINAL_WANTED.docx 템플릿 표 구조를 확인하지 못했습니다.");

  const updatedTables = [...tables];
  const summaryIndex = findTemplateTableIndex(tables, ["01", "요약"], 1);
  const overviewIndex = findTemplateTableIndex(tables, ["02", "미팅개요"], 3);
  const opinionIndex = findTemplateTableIndex(tables, ["03", "심사역검토의견"], -1);
  const ddqIndex = findTemplateTableIndex(tables, ["04", "핵심DDQ"], 6);
  const riskIndex = findTemplateTableIndex(tables, ["구분", "리스크내용", "심각도", "당사View"], 8);
  const followUpIndex = findTemplateTableIndex(tables, ["No.", "Follow-up항목", "담당", "기한"], 10);

  updateTemplateTable(updatedTables, tables, summaryIndex, fillTemplateSummaryTable, data);
  updateTemplateTable(updatedTables, tables, overviewIndex, fillTemplateOverviewTable, data);
  updateTemplateTable(updatedTables, tables, opinionIndex, fillTemplateOpinionTable, data);
  updateTemplateTable(updatedTables, tables, ddqIndex, fillTemplateDdqTable, data, 5);
  updateTemplateTable(updatedTables, tables, riskIndex, fillTemplateRiskTable, data, 6);
  updateTemplateTable(updatedTables, tables, followUpIndex, fillTemplateFollowUpTable, data, 8);

  updatedTables.forEach((table, index) => {
    if (table !== tables[index]) xml = replaceXmlBlockAt(xml, "tbl", index, table);
  });

  if (opinionIndex < 0) xml = replaceTemplateOpinionSection(xml, data);
  const opinions = deriveTemplateOpinionLines(data, 5);
  TEMPLATE_OPINION_TEXTS.forEach((text, index) => {
    xml = replaceTemplateLiteral(xml, text, opinions[index] || "");
  });
  return xml;
}

function updateTemplateTable(updatedTables, originalTables, index, filler, ...args) {
  if (index < 0 || !originalTables[index]) return;
  updatedTables[index] = filler(originalTables[index], ...args);
}

function findTemplateTableIndex(tables, requiredTexts, fallbackIndex = -1) {
  const normalizedNeedles = requiredTexts.map(normalizeWordSectionText);
  const found = tables.findIndex((table) => {
    const text = normalizeWordSectionText(getWordText(table));
    return normalizedNeedles.every((needle) => text.includes(needle));
  });
  if (found >= 0) return found;
  return tables[fallbackIndex] ? fallbackIndex : -1;
}

function assetNameForMemoTitle(value) {
  return String(value || "자산명 확인 필요").replace(/\s*검토\s*$/g, "").trim() || "자산명 확인 필요";
}

function buildMeetingOverviewRowsForDocx(data) {
  const overview = Object.fromEntries(data.overview || []);
  const metricRows = asArray(data.overviewMetricRows).length
    ? data.overviewMetricRows
    : buildOverviewMetricRows({
      sizeOrLoanAmount: data.sizeOrLoanAmount || overview["펀드 규모 / 대출 규모"] || overview["펀드 규모"] || overview["대출 규모"],
      periodOrMaturity: data.periodOrMaturity || overview["투자 기간 / 대출만기"] || overview["투자 기간"] || overview["대출만기"],
      returnOrLoanRate: data.returnOrLoanRate || overview["목표 수익률 / 대출금리"] || overview["목표 수익률"] || overview["대출금리"]
    }, data.reviewType);
  return [
    ["미팅 일시", formatKoreanDate(data.meetingDate)],
    ["장소", data.meetingLocation || "확인 필요"],
    ["GP 참석자", data.gpAttendees || data.contactName || "확인 필요"],
    ["당사 참석자", data.lpAttendees || `${profile.name || "담당자 확인 필요"} / ${profile.department || "부서 확인 필요"}`],
    ["지역 / 자산군", data.regionAssetClass || [overview["투자 지역"], data.classification].filter(Boolean).join(" / ") || "확인 필요"],
    ...metricRows,
    ["당사 검토 약정액", data.commitmentAmount || overview["당사 검토 약정액"] || "직접 입력"]
  ];
}

function fillTemplateSummaryTable(tableXml, data) {
  const summaryOpinion = buildSummaryOpinionText(data);
  const lines = [
    `${data.subjectLabel || "펀드명 / 대출명"}: ${data.fundName || data.assetName || "확인 필요"}`,
    `운용사(GP): ${data.managerName || "확인 필요"}`,
    `미팅일자: ${formatKoreanDate(data.meetingDate)}`,
    `검토 결과: ${data.reviewOpinion || "확인 필요"}`,
    `심사역 의견: ${summaryOpinion || "확인된 딜 정체성과 추가 DD 필요사항 확인 필요"}`
  ];
  return hasSectionHeaderRow(tableXml, "01")
    ? setExistingTableRows(tableXml, [[lines]], { startRow: 1, normalizeValueCells: true })
    : setSingleRowTableCells(tableXml, [lines]);
}

function fillTemplateOverviewTable(tableXml, data) {
  const rows = buildMeetingOverviewRowsForDocx(data);
  return setExistingTableRows(tableXml, rows, { startRow: hasSectionHeaderRow(tableXml, "02") ? 1 : 0, normalizeValueCells: true });
}

function fillTemplateOpinionTable(tableXml, data) {
  const lines = deriveStructuredReviewOpinionLines(data);
  const rows = [
    ["검토 의견 요약", lines[0] || "검토 의견 확인 필요"],
    ["투자 Highlight", lines[1] || "투자 매력 확인 필요"],
    ["핵심 리스크", lines[2] || "핵심 리스크 확인 필요"],
    ["추가 확인사항", lines[3] || "추가 확인사항 확인 필요"]
  ];
  return setExistingTableRows(tableXml, rows, { startRow: hasSectionHeaderRow(tableXml, "03") ? 1 : 0, normalizeValueCells: true });
}

function fillTemplateDdqTable(tableXml, data, limit) {
  const rows = asArray(data.ddq).slice(0, limit);
  const ddqRows = (rows.length ? rows : [{ question: "핵심 질의 확인 필요", answer: "미답변 / 추가 확인 필요" }]).map((item, index) => [
    `Q${index + 1}`,
    normalizeReportLanguage(item.question || "질문 확인 필요"),
    normalizeReportLanguage(item.answer || "미답변 / 추가 확인 필요")
  ]);
  return rebuildTableWithDataRows(tableXml, ddqRows, { headerRows: hasSectionHeaderRow(tableXml, "04") ? 2 : 1 });
}

function fillTemplateRiskTable(tableXml, data, limit) {
  const rows = deriveRiskRows(data).slice(0, limit).map((row) => [
    row.category,
    row.description,
    row.severity,
    row.view
  ]);
  return rebuildTableWithDataRows(tableXml, rows);
}

function fillTemplateFollowUpTable(tableXml, data, limit) {
  const rows = deriveFollowUpRows(data).slice(0, limit).map((row, index) => [
    `F-${String(index + 1).padStart(2, "0")}`,
    row.item,
    row.owner,
    row.due,
    row.status
  ]);
  return rebuildTableWithDataRows(tableXml, rows);
}

function hasSectionHeaderRow(tableXml, sectionNo) {
  const rows = getXmlBlocks(tableXml, "tr");
  if (!rows.length) return false;
  const firstText = normalizeWordSectionText(getWordText(rows[0]));
  return firstText.startsWith(String(sectionNo));
}

function deriveTemplateOpinionLines(data, limit) {
  const structured = deriveStructuredReviewOpinionLines(data);
  if (structured.length) return structured.slice(0, limit);
  const lines = mergeTextLists([
    ...splitOpinionText(data.finalOpinion),
    ...splitOpinionText(data.oneLineView),
    ...splitOpinionText(data.riskView),
    ...asArray(data.evidence).flatMap(splitOpinionText)
  ]).map((item) => limitSentences(formatListItemText(item), 2)).filter(Boolean);
  return lines.slice(0, limit);
}

function deriveStructuredReviewOpinionLines(data = {}) {
  const conclusion = ensureDealIdentityInOpinion(firstUsefulText([
    data.reviewConclusion,
    data.oneLineView,
    inferConclusionFromTone(data)
  ]), data);
  const whyThisDeal = firstUsefulText([
    data.whyThisDeal,
    compactListText(data.merits, 3),
    compactListText(data.evidence, 2),
    "투자 매력 확인 필요"
  ]);
  const keyRisks = firstUsefulText([
    data.keyRisksForOpinion,
    data.riskView,
    compactListText(data.risks, 3),
    "핵심 리스크 확인 필요"
  ]);
  const nextDdItems = firstUsefulText([
    data.nextDdItemsForOpinion,
    compactListText(data.followUps, 3),
    compactListText(data.sourceChecks, 3),
    "추가 확인자료 및 원문 대조 필요"
  ]);
  return [
    [conclusion, 2],
    [whyThisDeal, 2],
    [keyRisks, 2],
    [nextDdItems, 2]
  ].map(([value, limit]) => cleanReviewOpinionBody(limitSentences(formatListItemText(value), limit)))
    .filter(isMeaningfulFact);
}

function cleanReviewOpinionBullet(value) {
  return String(value || "")
    .replace(/^\s*(?:[-•·]\s*)?(?:결론|투자\s*매력\s*(?:\/\s*Why\s*This\s*Deal)?|Why\s*This\s*Deal|핵심\s*리스크\s*(?:\/\s*Key\s*Risks)?|Key\s*Risks|추가\s*확인\s*사항\s*(?:\/\s*Next\s*DD\s*Items)?|Next\s*DD\s*Items)\s*[:：-]\s*/i, "")
    .replace(/\s+/g, " ")
    .trim();
}

function firstUsefulText(values) {
  return values.map(formatListItemText).find(isMeaningfulFact) || "";
}

function compactListText(items, limit = 3) {
  return asArray(items)
    .map(formatListItemText)
    .filter(isMeaningfulFact)
    .slice(0, limit)
    .join("; ");
}

function buildSummaryOpinionText(data = {}) {
  const baseOpinion = firstUsefulText([
    data.finalOpinion,
    data.reviewConclusion,
    data.oneLineView,
    inferConclusionFromTone(data)
  ]);
  return cleanReviewOpinionBody(limitSentences(ensureDealIdentityInOpinion(baseOpinion, data), 3));
}

function ensureDealIdentityInOpinion(value, data = {}) {
  const base = cleanReviewOpinionBody(formatListItemText(value));
  const identity = deriveDealIdentitySentence(data);
  if (!identity) return base;
  if (!isMeaningfulFact(base)) return identity;
  if (opinionHasDealIdentity(base, data)) return base;
  return `${identity} ${base}`;
}

function deriveDealIdentitySentence(data = {}) {
  const dealName = firstUsefulDealText([data.assetName, data.fundName]);
  const managerName = firstUsefulDealText([data.managerName]);
  const classification = firstUsefulDealText([
    data.regionAssetClass,
    data.classification,
    compactListText(data.overviewMetricRows, 2)
  ]);
  const reviewType = normalizeReviewType(data.reviewType || "");
  const typeLabel = reviewType === "loan" ? "대출" : reviewType === "fund" ? "펀드" : "투자";
  const genericSubject = reviewType === "loan" ? "대출 건" : reviewType === "fund" ? "펀드 건" : "투자 건";
  const subject = [
    managerName ? `${managerName}이 제안한` : "",
    dealName || (managerName ? genericSubject : "")
  ].filter(Boolean).join(" ");
  if (subject && classification) {
    return `본 건은 ${subject}${koreanEuroParticle(subject)}, ${classification} 성격의 ${typeLabel} 검토 건임.`;
  }
  if (subject) return `본 건은 ${subject}에 대한 ${typeLabel} 검토 건임.`;
  if (classification) return `본 건은 ${classification} 성격의 ${typeLabel} 검토 건임.`;
  return "";
}

function koreanEuroParticle(value) {
  const chars = [...String(value || "").trim()];
  const last = chars[chars.length - 1] || "";
  const code = last.charCodeAt(0);
  if (code < 0xac00 || code > 0xd7a3) return "로";
  const jong = (code - 0xac00) % 28;
  return jong === 0 || jong === 8 ? "로" : "으로";
}

function opinionHasDealIdentity(value, data = {}) {
  const text = String(value || "");
  const identityParts = [
    data.assetName,
    data.fundName,
    data.managerName,
    ...String(data.regionAssetClass || data.classification || "").split(/[\/,;·•|]+/g)
  ].map(cleanFactValue).filter((item) => isUsefulDealText(item) && item.length >= 2);
  if (identityParts.some((item) => text.includes(item))) return true;
  return /본\s*건|본건|검토\s*건|대출\s*검토|펀드\s*검토|투자\s*검토/i.test(text)
    && /대출|펀드|PF|부동산|인프라|Equity|Debt|Credit|주거|멀티패밀리|바이오|헬스케어|소프트웨어|소비재|이커머스/i.test(text);
}

function firstUsefulDealText(values) {
  return asArray(values).map(cleanFactValue).find(isUsefulDealText) || "";
}

function isUsefulDealText(value) {
  const text = String(value || "").trim();
  return isMeaningfulFact(text)
    && !/^(직접\s*입력|미입력)$/i.test(text)
    && !/(확인\s*필요|자산명\s*확인|펀드명\s*\/\s*대출명|대출명\s*확인|펀드명\s*확인)/i.test(text);
}

function cleanReviewOpinionBody(value) {
  return cleanReviewOpinionBullet(value)
    .replace(/^\s*(?:[-•◈]\s*)?(?:결론|검토\s*의견\s*요약|투자\s*(?:매력|Highlight)|Why\s*This\s*Deal|핵심\s*리스크|Key\s*Risks|추가\s*확인\s*사항|추가확인사항|Next\s*DD\s*Items)\s*[:：]\s*/i, "")
    .replace(/\s+/g, " ")
    .trim();
}

function inferConclusionFromTone(data = {}) {
  const opinion = String(data.reviewOpinion || data.status || "");
  if (/드랍|중단|부적합/i.test(opinion)) return "드랍 또는 검토 중단 검토 필요";
  if (/보류/i.test(opinion)) return "보류 후 핵심 쟁점 해소 여부 확인 필요";
  if (/긍정|진행/i.test(opinion)) return "조건부 추가 검토 진행 가능";
  return "추가 검토 필요";
}

function splitOpinionText(value) {
  const text = formatListItemText(value);
  if (!text) return [];
  const bulletItems = text
    .split(/\n+|(?:^|\s)[•·]\s+|(?:^|\s)-\s+/)
    .map((item) => item.replace(/^\d+[.)]\s*/, "").trim())
    .filter(Boolean);
  const source = bulletItems.length > 1 ? bulletItems : [text];
  return source.flatMap((item) => {
    const sentences = item.match(/[^.!?\n]+[.!?]?/g)?.map((part) => part.trim()).filter(Boolean) || [];
    return sentences.length > 1 ? sentences : [item.trim()];
  }).filter(Boolean);
}

function replaceTemplateOpinionSection(xml, data) {
  const blocks = getWordBodyChildBlocks(xml);
  const startIndex = blocks.findIndex((block) => normalizeWordSectionText(block.text).includes("03심사역검토의견"));
  if (startIndex < 0) return xml;
  const endIndex = blocks.findIndex((block, index) => index > startIndex && normalizeWordSectionText(block.text).includes("04핵심DDQ"));
  if (endIndex < 0) return xml;
  const replacement = finalWantedOpinionList(data, { limit: 5 }) || paragraph("심사역 검토 의견 확인 필요", { indentLeft: 360, hanging: 220, size: 19 });
  return `${xml.slice(0, blocks[startIndex].end)}${replacement}${xml.slice(blocks[endIndex].start)}`;
}

function getWordBodyChildBlocks(xml) {
  const bodyOpen = xml.match(/<w:body\b[^>]*>/);
  const bodyCloseIndex = xml.lastIndexOf("</w:body>");
  if (!bodyOpen || bodyCloseIndex < 0) return [];
  const bodyStart = bodyOpen.index + bodyOpen[0].length;
  const bodyXml = xml.slice(bodyStart, bodyCloseIndex);
  const blocks = [];
  const pattern = /<w:(p|tbl)\b[\s\S]*?<\/w:\1>/g;
  let match;
  while ((match = pattern.exec(bodyXml))) {
    blocks.push({
      block: match[0],
      text: getWordText(match[0]),
      start: bodyStart + match.index,
      end: bodyStart + match.index + match[0].length
    });
  }
  return blocks;
}

function getWordText(xml) {
  return [...String(xml || "").matchAll(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g)]
    .map((match) => unescapeXml(match[1]))
    .join("");
}

function normalizeWordSectionText(value) {
  return String(value || "").replace(/\s+/g, "").trim();
}

const TEMPLATE_OPINION_TEXTS = [
  "GP의 인프라 섹터 운용 경험은 업계 평균 대비 양호한 수준으로 판단되나, Fund III 기준 트랙레코드(Realized Deal 실적)가 미흡하여 Net IRR 추정치의 신뢰성 검증 필요.",
  "목표 LTV 60% 수준은 현행 고금리 환경(기준금리 3.5%)하 이자비용 부담이 상당하며, Debt Service Coverage Ratio(DSCR) 안전마진이 과거 펀드 대비 축소된 것으로 추정됨.",
  "주요 포트폴리오 자산 3건 중 2건이 사회인프라(가용성 기반, Availability-based)이나, 나머지 1건은 Merchant 방식으로 현금흐름 변동성 요인이 혼재 — 섹터 일관성 우려.",
  "Carried Interest 20% / Hurdle Rate 7% 구조는 업계 표준 범위이나, GP-LP 이해관계 정렬(Alignment of Interest) 측면에서 GP 공동투자(Co-investment) 규모 확인 필요.",
  "전반적으로 당사 대체투자 가이드라인 부합 여부 검토 중이며, 추가 DDQ 회신 및 독립적 재무모델 검증 완료 시 Pre-IC 상정 가능한 것으로 판단됨."
];

function getXmlBlocks(xml, tagName) {
  return xml.match(new RegExp(`<w:${tagName}\\b[\\s\\S]*?</w:${tagName}>`, "g")) || [];
}

function replaceXmlBlockAt(xml, tagName, index, replacement) {
  let current = -1;
  return xml.replace(new RegExp(`<w:${tagName}\\b[\\s\\S]*?</w:${tagName}>`, "g"), (match) => {
    current += 1;
    return current === index ? replacement : match;
  });
}

function replaceTemplateLiteral(xml, sourceText, nextText) {
  return xml.replace(escapeXml(sourceText), escapeXml(toEumseumStyle(nextText)));
}

function replaceAssetNamePlaceholder(xml, assetName) {
  const safeName = escapeXml(toEumseumStyle(assetName || "자산명 확인 필요"));
  let next = xml.replace(escapeXml("{{자산명}}"), safeName);
  next = next.replace(
    /(<w:t\b[^>]*>)\{\{(<\/w:t>[\s\S]*?<w:t\b[^>]*>)자산명(<\/w:t>[\s\S]*?<w:t\b[^>]*>)\}\}/,
    `$1${safeName}$2$3`
  );
  return next;
}

function setSingleRowTableCells(tableXml, cellValues) {
  const rows = getXmlBlocks(tableXml, "tr");
  if (!rows.length) return tableXml;
  const row = setRowCellTexts(rows[0], cellValues);
  return setTableRows(tableXml, [row]);
}

function setExistingTableRows(tableXml, rowValues, options = {}) {
  const rows = getXmlBlocks(tableXml, "tr");
  const updatedRows = rows.map((row, index) => {
    const values = rowValues[index - (options.startRow || 0)];
    return values ? setRowCellTexts(row, values, options) : row;
  });
  return setTableRows(tableXml, updatedRows);
}

function rebuildTableWithDataRows(tableXml, rowValues, options = {}) {
  const rows = getXmlBlocks(tableXml, "tr");
  if (!rows.length) return tableXml;
  const headerRows = Math.max(1, options.headerRows || 1);
  const headers = rows.slice(0, Math.min(headerRows, rows.length));
  const dataTemplate = rows[headerRows] || rows[rows.length - 1] || rows[0];
  const nextRows = [...headers, ...rowValues.map((values) => setRowCellTexts(dataTemplate, values))];
  return setTableRows(tableXml, nextRows);
}

function setTableRows(tableXml, rows) {
  const existingRows = getXmlBlocks(tableXml, "tr");
  if (!existingRows.length) return tableXml;
  const firstRow = existingRows[0];
  const lastRow = existingRows[existingRows.length - 1];
  const prefix = tableXml.slice(0, tableXml.indexOf(firstRow));
  const suffix = tableXml.slice(tableXml.indexOf(lastRow) + lastRow.length);
  return `${prefix}${rows.join("")}${suffix}`;
}

function setRowCellTexts(rowXml, values, options = {}) {
  const cells = getXmlBlocks(rowXml, "tc");
  let next = rowXml;
  cells.forEach((cell, index) => {
    const value = values[index] ?? "";
    const cellOptions = options.normalizeValueCells && index > 0 ? { normalizeRun: true } : {};
    next = next.replace(cell, setCellText(cell, value, cellOptions));
  });
  return next;
}

function setCellText(cellXml, value, options = {}) {
  const start = cellXml.match(/^<w:tc\b[^>]*>/)?.[0] || "<w:tc>";
  const tcPr = cellXml.match(/<w:tcPr[\s\S]*?<\/w:tcPr>/)?.[0] || "";
  const paragraphs = getXmlBlocks(cellXml, "p");
  const templateParagraph = paragraphs[0] || "<w:p><w:r><w:t></w:t></w:r></w:p>";
  const lines = Array.isArray(value) ? value : String(value || "").split("\n");
  const paragraphXml = (lines.length ? lines : [""]).map((line) => setParagraphXmlText(templateParagraph, line, options)).join("");
  return `${start}${tcPr}${paragraphXml}</w:tc>`;
}

function setParagraphXmlText(paragraphXml, value, options = {}) {
  const text = toEumseumStyle(typeof value === "object" ? value.text : value);
  const start = paragraphXml.match(/^<w:p\b[^>]*>/)?.[0] || "<w:p>";
  const pPr = paragraphXml.match(/<w:pPr[\s\S]*?<\/w:pPr>/)?.[0] || "";
  const rPr = options.normalizeRun ? normalizedTableValueRunPr() : paragraphXml.match(/<w:rPr[\s\S]*?<\/w:rPr>/)?.[0] || "";
  if (!String(text || "").trim()) return `${start}${pPr}</w:p>`;
  return `${start}${pPr}<w:r>${rPr}<w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`;
}

function normalizedTableValueRunPr() {
  return '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Malgun Gothic"/><w:color w:val="1F2937"/><w:sz w:val="18"/></w:rPr>';
}

function documentXml(data) {
  const body = fullReviewMemoBody(data);
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`;
}

function fullReviewMemoBody(data) {
  const body = [
    finalWantedTitleBlock(data),
    finalWantedSectionHeader("01", "요약 (Executive Summary)"),
    finalWantedSummaryTable(data),
    finalWantedSectionHeader("02", "미팅 개요 (Meeting Overview)"),
    finalWantedMeetingOverviewTable(data),
    finalWantedSectionHeader("03", "심사역 검토 의견 (Reviewer's Opinion)"),
    finalWantedOpinionTable(data),
    finalWantedSectionHeader("04", "핵심 DDQ — 미팅 Q&A 기록"),
    finalWantedDdqTable(data, 5),
    finalWantedSectionHeader("05", "주요 예상 리스크 (Key Risk Factors)"),
    finalWantedRiskTable(data, 6),
    finalWantedSectionHeader("06", "향후 Follow-up 사항"),
    finalWantedFollowUpTable(data, 3),
    finalWantedFooter(),
    sectionProperties()
  ].join("");
  return body;
}

function documentMetaBlock(data) {
  return `<w:tbl><w:tblPr>${tableWidth()}<w:tblBorders><w:bottom w:val="single" w:sz="10" w:color="CBD5E1"/></w:tblBorders><w:tblCellMar>${tableMargins(0, 0, 0, 0)}</w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="5600"/><w:gridCol w:w="3760"/></w:tblGrid>
    <w:tr>
      ${tableCell("Strictly Confidential · 투자심사본부 대체투자팀", { width: 5600, borders: "none", padding: 0, color: "334155", bold: true, size: 17 })}
      ${tableCell(`작성일 ${formatDocxDate(new Date())} · ${data.status || "Draft 초안"}`, { width: 3760, borders: "none", align: "right", color: "64748B", size: 17, padding: 0 })}
    </w:tr>
  </w:tbl>`;
}

function titleBlock(data) {
  const subtitle = [data.managerName, data.classification].filter(Boolean).join(" · ");
  return [
    paragraph("LP Meeting Review Memo", { style: "Title", align: "center", bold: true, size: 31, color: "0F172A" }),
    paragraph(data.fundName || "펀드명 / 대출명 확인 필요", { align: "center", bold: true, size: 23, color: "334155" }),
    paragraph(subtitle || "운용사 및 투자 분류 확인 필요", { align: "center", size: 18, color: "64748B" })
  ].join("");
}

function finalWantedTitleBlock(data, options = {}) {
  const assetName = assetNameForMemoTitle(data.assetName || data.fundName);
  return [
    paragraph(`${assetName} 검토 메모`, { align: "center", bold: true, size: options.compact ? 28 : 31, color: "0F172A" }),
    paragraph("GP Meeting Review Memo", { align: "center", bold: true, size: 18, color: "64748B" })
  ].join("");
}

function finalWantedSectionHeader(number, title) {
  return finalWantedTable([520, 8840], `
    <w:tr>
      ${tableCell(number, { width: 520, fill: "0F172A", color: "FFFFFF", bold: true, align: "center", size: 18 })}
      ${tableCell(title, { width: 8840, fill: "F1F5F9", color: "0F172A", bold: true, size: 20 })}
    </w:tr>
  `, { topSize: 0, bottomSize: 0, cellMargin: tableMargins(80, 120, 80, 120) });
}

function finalWantedSummaryTable(data) {
  const summaryOpinion = buildSummaryOpinionText(data);
  const lines = [
    `${data.subjectLabel || "펀드명 / 대출명"}: ${data.fundName || data.assetName || "확인 필요"}`,
    `운용사(GP): ${data.managerName || "확인 필요"}`,
    `미팅일자: ${formatKoreanDate(data.meetingDate)}`,
    `검토 결과: ${data.reviewOpinion || "확인 필요"}`,
    `심사역 의견: ${summaryOpinion || "확인된 딜 정체성과 추가 DD 필요사항 확인 필요"}`
  ].map((text) => ({ text }));
  return finalWantedTable([9360], `<w:tr>${tableCell(lines, { width: 9360, fill: "F8FAFC" })}</w:tr>`, {
    cellMargin: tableMargins(120, 150, 120, 150)
  });
}

function finalWantedMeetingOverviewTable(data, options = {}) {
  const rows = buildMeetingOverviewRowsForDocx(data);
  const visibleRows = options.compact ? rows.slice(0, 6) : rows;
  return finalWantedTable([1700, 7660], visibleRows.map(([label, value]) => `
    <w:tr>
      ${tableCell(label, { width: 1700, fill: "F8FAFC", bold: true, color: "334155", align: "center" })}
      ${tableCell(value, { width: 7660 })}
    </w:tr>
  `).join(""));
}

function finalWantedOpinionList(data, options = {}) {
  const limit = options.limit || 5;
  const opinionItems = deriveTemplateOpinionLines(data, limit)
    .map((item) => `• ${limitSentences(formatListItemText(item), 2)}`);
  return opinionItems.slice(0, limit).map((item) => paragraph(item, { indentLeft: 360, hanging: 220, size: 19 })).join("");
}

function finalWantedOpinionTable(data) {
  const lines = deriveStructuredReviewOpinionLines(data);
  const rows = [
    ["검토 의견 요약", lines[0] || "검토 의견 확인 필요"],
    ["투자 Highlight", lines[1] || "투자 매력 확인 필요"],
    ["핵심 리스크", lines[2] || "핵심 리스크 확인 필요"],
    ["추가 확인사항", lines[3] || "추가 확인사항 확인 필요"]
  ];
  return finalWantedTable([1700, 7660], rows.map(([label, value]) => `
    <w:tr>
      ${tableCell(label, { width: 1700, fill: "F8FAFC", bold: true, color: "334155", align: "center" })}
      ${tableCell(value, { width: 7660 })}
    </w:tr>
  `).join(""));
}

function finalWantedDdqTable(data, limit = 5) {
  const rows = asArray(data.ddq).slice(0, limit);
  const ddqRows = rows.length ? rows : [{ question: "핵심 질의 확인 필요", answer: "미답변 / 추가 확인 필요" }];
  return finalWantedTable([620, 3900, 4840], `
    <w:tr>
      ${tableCell("No.", { width: 620, fill: "F1F5F9", bold: true, align: "center" })}
      ${tableCell("GP 측 질문", { width: 3900, fill: "F1F5F9", bold: true, align: "center" })}
      ${tableCell("GP 답변 요지", { width: 4840, fill: "F1F5F9", bold: true, align: "center" })}
    </w:tr>
    ${ddqRows.map((item, index) => `<w:tr>
      ${tableCell(`Q${index + 1}`, { width: 620, fill: "F8FAFC", bold: true, align: "center" })}
      ${tableCell(item.question || "질문 확인 필요", { width: 3900 })}
      ${tableCell(item.answer || "미답변 / 추가 확인 필요", { width: 4840 })}
    </w:tr>`).join("")}
  `);
}

function finalWantedRiskTable(data, limit = 6) {
  const rows = deriveRiskRows(data).slice(0, limit);
  return finalWantedTable([1200, 4700, 1050, 2410], `
    <w:tr>
      ${tableCell("구분", { width: 1200, fill: "F1F5F9", bold: true, align: "center" })}
      ${tableCell("리스크 내용", { width: 4700, fill: "F1F5F9", bold: true, align: "center" })}
      ${tableCell("심각도", { width: 1050, fill: "F1F5F9", bold: true, align: "center" })}
      ${tableCell("당사 View", { width: 2410, fill: "F1F5F9", bold: true, align: "center" })}
    </w:tr>
    ${rows.map((row) => `<w:tr>
      ${tableCell(row.category, { width: 1200, bold: true, align: "center" })}
      ${tableCell(row.description, { width: 4700 })}
      ${tableCell(row.severity, { width: 1050, bold: true, align: "center" })}
      ${tableCell(row.view, { width: 2410 })}
    </w:tr>`).join("")}
  `);
}

function finalWantedFollowUpTable(data, limit = 3) {
  const rows = deriveFollowUpRows(data).slice(0, limit);
  return finalWantedTable([700, 5660, 1500, 1500], `
    <w:tr>
      ${tableCell("No.", { width: 700, fill: "F1F5F9", bold: true, align: "center" })}
      ${tableCell("Follow-up 항목", { width: 5660, fill: "F1F5F9", bold: true, align: "center" })}
      ${tableCell("담당", { width: 1500, fill: "F1F5F9", bold: true, align: "center" })}
      ${tableCell("기한", { width: 1500, fill: "F1F5F9", bold: true, align: "center" })}
    </w:tr>
    ${rows.map((row, index) => `<w:tr>
      ${tableCell(`F-${String(index + 1).padStart(2, "0")}`, { width: 700, bold: true, align: "center" })}
      ${tableCell(row.item, { width: 5660 })}
      ${tableCell(row.owner, { width: 1500, align: "center" })}
      ${tableCell(row.due, { width: 1500, align: "center" })}
    </w:tr>`).join("")}
  `, { bottomSize: 18 });
}

function finalWantedFooter() {
  return [
    paragraph("본 메모는 투자 결정을 위한 내부 검토 자료이며, 외부 유출이 엄격히 금지됨.", { align: "center", size: 16, color: "64748B" }),
    paragraph("This document is strictly confidential and intended solely for internal use.", { align: "center", size: 16, color: "64748B" })
  ].join("");
}

function finalWantedTable(columns, rowsXml, options = {}) {
  const cellMargin = options.cellMargin || tableMargins(100, 130, 100, 130);
  return `<w:tbl><w:tblPr>${tableWidth()}<w:tblLayout w:type="fixed"/><w:tblBorders>${borderXml("CBD5E1", options)}</w:tblBorders><w:tblCellMar>${cellMargin}</w:tblCellMar></w:tblPr><w:tblGrid>${columns.map((width) => `<w:gridCol w:w="${width}"/>`).join("")}</w:tblGrid>${rowsXml}</w:tbl>${spacer(0)}`;
}

function paragraph(text, options = {}) {
  const style = options.style ? `<w:pStyle w:val="${options.style}"/>` : "";
  const align = options.align ? `<w:jc w:val="${options.align}"/>` : "";
  const indent = options.indentLeft ? `<w:ind w:left="${options.indentLeft}"${options.hanging ? ` w:hanging="${options.hanging}"` : ""}/>` : "";
  const keepLines = options.keepLines ? "<w:keepLines/>" : "";
  return `<w:p><w:pPr>${style}${align}${indent}${keepLines}<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>${run(text, Boolean(options.bold), options)}</w:p>`;
}

function bulletSection(title, items) {
  return `${paragraph(title)}${bulletList(items)}`;
}

function bulletList(items) {
  return asArray(items).filter(Boolean).map((item) => paragraph(`• ${formatListItemText(item)}`)).join("");
}

function executiveSummaryTable(data) {
  const evidence = asArray(data.evidence).slice(0, 4).map((item) => ({ text: `• ${formatListItemText(item)}` }));
  return sectionTable("1. Executive Summary", `
    <w:tr>${tableCell("검토의견", labelCell())}${tableCell(data.reviewOpinion || "확인 필요", { width: 7660, bold: true, color: "0F172A" })}</w:tr>
    <w:tr>${tableCell("한 줄 결론", labelCell())}${tableCell(data.oneLineView || "확인 필요", { width: 7660, fill: "F8FAFC", bold: true })}</w:tr>
    <w:tr>${tableCell("핵심 근거", labelCell())}${tableCell(evidence.length ? evidence : "확인 필요", { width: 7660 })}</w:tr>
  `);
}

function investmentOverviewTable(data) {
  const overview = Object.fromEntries(data.overview || []);
  const classification = data.classification || "분류 확인 필요";
  const rows = [
    ["운용사", data.managerName],
    ["펀드명 / 대출명", data.fundName],
    ["자산군", classification],
    ["투자지역", overview["투자 지역"]],
    ["섹터 / 구조", overview["섹터 / 구조"]],
    ["딜 / 자산 메모", overview["딜 / 자산 메모"]]
  ];
  return sectionTable("2. 투자 개요", rows.map(([label, value]) => `
    <w:tr>${tableCell(label, labelCell())}${tableCell(value || "확인 필요", { width: 7660, bold: label !== "딜 / 자산 메모" })}</w:tr>
  `).join(""));
}

function opinionTable(data) {
  return sectionTable("3. 의견", `
    <w:tr>
      ${tableCell("종합 의견", labelCell())}
      ${tableCell(limitSentences(data.finalOpinion || "추가 DD 결과 확인 후 투심 상정 여부 판단 필요", 5), { width: 7660 })}
    </w:tr>
  `);
}

function ddqInsightsTable(data) {
  const rows = asArray(data.ddq).slice(0, 4);
  const ddqRows = rows.length ? rows : [{ label: "Q1. DDQ", question: "핵심 질의 확인 필요", answer: "미답변 / 추가 확인 필요" }];
  return ddqSectionTable("4. 핵심 DDQ", ddqRows.map((item, index) => `
    <w:tr>
      ${tableCell(`Q${index + 1}`, ddqLabelCell())}
      ${tableCell([
        { text: cleanDocxLabel(item.label || "DDQ"), bold: true, color: "334155", size: 18 },
        { text: "LP 질의", bold: true, color: "0F172A" },
        { text: item.question || "질문 확인 필요" },
        { text: "GP 답변 / 미팅 기록", bold: true, color: "0F172A" },
        { text: item.answer || "미답변 / 추가 확인 필요", color: "475569" }
      ], { width: 8600 })}
    </w:tr>
  `).join(""));
}

function riskTableOnly(data) {
  const risks = asArray(data.risks).slice(0, 6);
  const riskLines = (risks.length ? risks : ["핵심 리스크 확인 필요"]).map((item, index) => {
    const level = index < 2 ? "High" : "Medium";
    return { text: `[${level}] ${formatListItemText(item)}` };
  });
  return sectionTable("5. 주요 예상 리스크", `
    <w:tr>
      ${tableCell("리스크", labelCell())}
      ${tableCell(riskLines, { width: 7660 })}
    </w:tr>
  `);
}

function followUpActionTable(data) {
  const actionItems = [
    ...asArray(data.followUps),
    ...asArray(data.nextActions),
    ...asArray(data.sourceChecks).map((item) => `원문 대조: ${formatListItemText(item)}`)
  ].slice(0, 8);
  const followUpLines = (actionItems.length ? actionItems : ["원문 계약서 및 추가 DD 자료 확인"]).map((item) => ({ text: `□ ${formatListItemText(item)}` }));
  return sectionTable("6. 향후 Follow-up", `
    <w:tr>
      ${tableCell("요청 / 액션", labelCell())}
      ${tableCell(followUpLines, { width: 7660, fill: "F8FAFC" })}
    </w:tr>
  `, { bottomSize: 18 });
}

function sectionTable(title, rowsXml, options = {}) {
  const cellMargin = options.compact ? tableMargins(70, 120, 70, 120) : tableMargins(105, 145, 105, 145);
  return `<w:tbl><w:tblPr>${tableWidth()}<w:tblLayout w:type="fixed"/><w:tblBorders>${borderXml("CBD5E1", options)}</w:tblBorders><w:tblCellMar>${cellMargin}</w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="1700"/><w:gridCol w:w="7660"/></w:tblGrid>
    <w:tr>${tableCell(title, { width: 9360, gridSpan: 2, fill: "F1F5F9", bold: true, color: "0F172A", size: 20 })}</w:tr>
    ${rowsXml}
  </w:tbl>${spacer(0)}`;
}

function ddqSectionTable(title, rowsXml, options = {}) {
  const cellMargin = options.compact ? tableMargins(70, 120, 70, 120) : tableMargins(105, 145, 105, 145);
  return `<w:tbl><w:tblPr>${tableWidth()}<w:tblLayout w:type="fixed"/><w:tblBorders>${borderXml("CBD5E1", options)}</w:tblBorders><w:tblCellMar>${cellMargin}</w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="760"/><w:gridCol w:w="8600"/></w:tblGrid>
    <w:tr>${tableCell(title, { width: 9360, gridSpan: 2, fill: "F1F5F9", bold: true, color: "0F172A", size: 20 })}</w:tr>
    ${rowsXml}
  </w:tbl>${spacer(0)}`;
}

function twoColumnSectionTable(title, rowsXml, options = {}) {
  const cellMargin = options.compact ? tableMargins(70, 120, 70, 120) : tableMargins(105, 145, 105, 145);
  return `<w:tbl><w:tblPr>${tableWidth()}<w:tblLayout w:type="fixed"/><w:tblBorders>${borderXml("CBD5E1", options)}</w:tblBorders><w:tblCellMar>${cellMargin}</w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="4680"/><w:gridCol w:w="4680"/></w:tblGrid>
    <w:tr>${tableCell(title, { width: 9360, gridSpan: 2, fill: "F1F5F9", bold: true, color: "0F172A", size: 20 })}</w:tr>
    ${rowsXml}
  </w:tbl>${spacer(0)}`;
}

function labelCell(extra = {}) {
  return { width: 1700, fill: "F8FAFC", bold: true, color: "334155", align: "center", ...extra };
}

function ddqLabelCell(extra = {}) {
  return { width: 760, fill: "F8FAFC", bold: true, color: "334155", align: "center", ...extra };
}

function tableCell(content, options = {}) {
  const width = options.width || 3900;
  const fill = options.fill || "FFFFFF";
  const gridSpan = options.gridSpan ? `<w:gridSpan w:val="${options.gridSpan}"/>` : "";
  const borders = options.borders === "none" ? "<w:tcBorders><w:top w:val=\"nil\"/><w:left w:val=\"nil\"/><w:bottom w:val=\"nil\"/><w:right w:val=\"nil\"/></w:tcBorders>" : "";
  const padding = options.padding === 0 ? tableMargins(0, 0, 0, 0) : tableMargins(100, 140, 100, 140);
  return `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/>${gridSpan}<w:shd w:fill="${fill}"/>${borders}<w:tcMar>${padding}</w:tcMar><w:vAlign w:val="center"/></w:tcPr>${cellParagraphs(content, options)}</w:tc>`;
}

function cellParagraphs(content, options = {}) {
  if (Array.isArray(content)) {
    return content.map((item) => {
      if (typeof item === "string") return paragraph(toEumseumStyle(item), paragraphOptionsForLine(item, options));
      const text = item.text || "확인 필요";
      return paragraph(toEumseumStyle(text), paragraphOptionsForLine(text, { ...options, ...item }));
    }).join("");
  }
  return String(content || "확인 필요").split("\n").map((line) => paragraph(toEumseumStyle(line), paragraphOptionsForLine(line, {
    bold: options.bold,
    color: options.color,
    size: options.size,
    align: options.align
  }))).join("");
}

function paragraphOptionsForLine(line, options = {}) {
  const trimmed = String(line || "").trim();
  const next = { ...options };
  if (/^[•□]/.test(trimmed)) {
    next.indentLeft = next.indentLeft || 360;
    next.hanging = next.hanging || 220;
  } else if (/^\[(High|Medium|Low)\]/i.test(trimmed)) {
    next.indentLeft = next.indentLeft || 300;
  }
  return next;
}

function tableWidth() {
  return `<w:tblW w:w="9360" w:type="dxa"/><w:tblInd w:w="0" w:type="dxa"/>`;
}

function tableMargins(top, left, bottom, right) {
  return `<w:top w:w="${top}" w:type="dxa"/><w:left w:w="${left}" w:type="dxa"/><w:bottom w:w="${bottom}" w:type="dxa"/><w:right w:w="${right}" w:type="dxa"/>`;
}

function borderXml(color = "CBD5E1", options = {}) {
  const topSize = options.topSize ?? 14;
  const bottomSize = options.bottomSize ?? 14;
  return `<w:top w:val="single" w:sz="${topSize}" w:color="0F172A"/><w:left w:val="single" w:sz="6" w:color="${color}"/><w:bottom w:val="single" w:sz="${bottomSize}" w:color="0F172A"/><w:right w:val="single" w:sz="6" w:color="${color}"/><w:insideH w:val="single" w:sz="6" w:color="${color}"/><w:insideV w:val="single" w:sz="6" w:color="${color}"/>`;
}

function run(text, bold = false, options = {}) {
  const size = options.size || 19;
  const color = options.color || "1E293B";
  return `<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Malgun Gothic"/>${bold ? "<w:b/>" : ""}<w:color w:val="${color}"/><w:sz w:val="${size}"/></w:rPr><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>`;
}

function spacer(after = 120) {
  if (!after) return "";
  return `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:p>`;
}

function sectionProperties() {
  return `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1080" w:right="1080" w:bottom="1080" w:left="1080"/></w:sectPr>`;
}

function formatDocxDate(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}. ${month}. ${day}`;
}

function cleanDocxLabel(value) {
  return String(value || "DDQ").replace(/^Q\d+\.\s*/, "").replace(/\s+/g, " ").trim();
}

function escapeXml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function unescapeXml(value) {
  return String(value ?? "")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, "\"")
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

function insertTimestamp() {
  const textarea = $("meetingNotes");
  const now = new Date();
  const stamp = `[${String(now.getHours()).padStart(2, "0")}:${String(now.getMinutes()).padStart(2, "0")}] `;
  const insertion = `${textarea.value ? "\n" : ""}${stamp}`;
  textarea.value += insertion;
  textarea.focus();
  syncFieldsToState();
  saveState();
  scheduleExistingLibraryUpdate();
}

function saveCurrentMeeting(message = "최근 미팅에 저장했습니다.") {
  syncFieldsToState();
  saveState();
  if (!localMeetingStorageEnabled) {
    showRecentMeetingFeedback("로컬 저장이 꺼져 있어 현재 미팅을 브라우저에 저장하지 않았습니다.");
    toast("로컬 저장이 꺼져 있습니다.");
    return;
  }
  saveMeetingToLibrary();
  showRecentMeetingFeedback(message);
  toast(message);
}

function saveMeetingToLibrary() {
  if (!localMeetingStorageEnabled) return;
  const library = getLibrary();
  if (!state.meetingId) state.meetingId = createMeetingId();
  const id = state.meetingId;
  const item = {
    id,
    managerName: state.meeting.managerName || "운용사 미정",
    fundName: state.meeting.fundName || "펀드명 / 대출명 미정",
    updatedAt: new Date().toISOString(),
    status: state.postMeetingMemo ? "보고서 완료" : state.preMeetingBrief ? "브리프 완료" : "작성 중",
    selectedFileName: state.selectedFileName || "",
    snapshot: sanitizeStateForStorage({ ...state, meetingId: id })
  };
  const next = [item, ...library.filter((entry) => entry.id !== id)].slice(0, 8);
  localStorage.setItem(LIBRARY_KEY, JSON.stringify(next));
  saveState();
  renderRecentMeetings();
  showRecentMeetingFeedback(`${item.fundName} 저장됨 · ${formatDateTime(item.updatedAt)}`);
}

function getLibrary() {
  try {
    return JSON.parse(localStorage.getItem(LIBRARY_KEY) || "[]")
      .filter((item) => item && item.id)
      .sort((a, b) => new Date(b.updatedAt || 0) - new Date(a.updatedAt || 0));
  } catch {
    return [];
  }
}

function renderRecentMeetings() {
  const list = $("recentMeetings");
  const library = getLibrary();
  if (!library.length) {
    list.innerHTML = `<li class="px-2 py-3 text-sm leading-5 text-slate-500">저장된 미팅이 없습니다. 상단 저장 버튼을 누르면 최근 미팅에 추가됩니다.</li>`;
    return;
  }
  list.innerHTML = library.map((item) => `
    <li class="group flex items-stretch gap-1 rounded-md ${item.id === state.meetingId ? "bg-slate-800 ring-1 ring-slate-700" : "hover:bg-slate-800"}">
      <button class="min-w-0 flex-1 rounded-md px-3 py-2 text-left transition hover:text-white" title="이 미팅 불러오기" data-load-meeting="${escapeAttribute(item.id)}">
        <span class="block truncate text-sm font-semibold text-slate-200">${escapeHtml(item.fundName)}</span>
        <span class="block truncate text-xs text-slate-500 group-hover:text-slate-400">${escapeHtml(item.managerName)} · ${escapeHtml(item.status)} · ${formatDateTime(item.updatedAt)}</span>
        ${item.selectedFileName ? `<span class="mt-1 block truncate text-[11px] text-slate-600 group-hover:text-slate-500">원본 파일 재선택 필요: ${escapeHtml(item.selectedFileName)}</span>` : ""}
      </button>
      <button class="mr-1 flex w-8 shrink-0 items-center justify-center rounded-md text-slate-500 transition hover:bg-slate-700 hover:text-white" title="최근 미팅 삭제" data-delete-meeting="${escapeAttribute(item.id)}">
        <i data-lucide="trash-2" class="h-4 w-4"></i>
      </button>
    </li>`).join("");
  list.querySelectorAll("[data-load-meeting]").forEach((button) => {
    button.addEventListener("click", () => loadMeetingFromLibrary(button.dataset.loadMeeting));
  });
  list.querySelectorAll("[data-delete-meeting]").forEach((button) => {
    button.addEventListener("click", (event) => {
      event.stopPropagation();
      deleteMeetingFromLibrary(button.dataset.deleteMeeting);
    });
  });
  refreshIcons();
}

function scheduleExistingLibraryUpdate() {
  if (!localMeetingStorageEnabled) return;
  if (!state.meetingId || !isMeetingInLibrary(state.meetingId)) return;
  clearTimeout(libraryUpdateTimer);
  libraryUpdateTimer = setTimeout(() => {
    saveMeetingToLibrary();
  }, 800);
}

function isMeetingInLibrary(id) {
  return getLibrary().some((entry) => entry.id === id);
}

function showRecentMeetingFeedback(message) {
  const box = $("recentMeetingFeedback");
  if (!box) return;
  box.textContent = message;
  box.classList.remove("hidden");
  clearTimeout(showRecentMeetingFeedback.timer);
  showRecentMeetingFeedback.timer = setTimeout(() => {
    box.classList.add("hidden");
  }, 4200);
}

function loadMeetingFromLibrary(id) {
  const item = getLibrary().find((entry) => entry.id === id);
  if (!item?.snapshot) {
    toast("이전 저장 항목은 불러올 데이터가 없습니다. 다시 저장해주세요.");
    return;
  }
  selectedFile = null;
  state = mergeSavedState(item.snapshot);
  state.meetingId = item.id;
  applyStateToFields();
  renderAll();
  switchPhase(state.activePhase || "prep");
  toast("최근 미팅을 불러왔습니다. 원본 파일은 다시 선택해야 합니다.");
}

function deleteMeetingFromLibrary(id) {
  const item = getLibrary().find((entry) => entry.id === id);
  const label = item ? `${item.managerName} / ${item.fundName}` : "선택한 미팅";
  if (!confirm(`${label}을(를) 최근 미팅에서 삭제할까요? 현재 작성 중인 화면은 유지됩니다.`)) return;
  const next = getLibrary().filter((entry) => entry.id !== id);
  localStorage.setItem(LIBRARY_KEY, JSON.stringify(next));
  if (state.meetingId === id) {
    state.meetingId = null;
    saveState();
  }
  renderRecentMeetings();
  showRecentMeetingFeedback("최근 미팅에서 삭제했습니다. 현재 화면 내용은 유지됩니다.");
  toast("최근 미팅에서 삭제했습니다.");
}

function getMeetingStorageBackup() {
  return {
    exportedAt: new Date().toISOString(),
    app: "LP Meeting Copilot",
    note: "Gemini API Key는 포함하지 않습니다. 이름/부서/모델 프로필은 별도 localStorage 키에 저장됩니다.",
    currentMeeting: sanitizeStateForStorage(state),
    recentMeetings: getLibrary()
  };
}

function downloadJsonFile(data, fileName) {
  const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  anchor.click();
  URL.revokeObjectURL(url);
}

function resetCurrentMeetingAfterStorageClear() {
  selectedFile = null;
  state = emptyState();
  applyStateToFields();
  renderAll();
  switchPhase("prep");
}

function removeMeetingStorageKeys() {
  localStorage.removeItem(STORAGE_KEY);
  localStorage.removeItem(LIBRARY_KEY);
}

function clearLocalMeetingStorage() {
  if (!confirm("브라우저에 저장된 현재 미팅과 최근 미팅 목록을 모두 삭제할까요? API Key는 원래 저장되어 있지 않고, 이름/부서/모델 설정은 유지됩니다.")) return;
  removeMeetingStorageKeys();
  resetCurrentMeetingAfterStorageClear();
  renderStorageSettingsStatus();
  showRecentMeetingFeedback("브라우저 저장소를 비웠습니다.");
  toast("로컬 미팅 저장소를 모두 삭제했습니다.");
}

function exportMeetingStorageAndClear() {
  const backup = getMeetingStorageBackup();
  downloadJsonFile(backup, `meeting-copilot-backup-${new Date().toISOString().slice(0, 10)}.json`);
  removeMeetingStorageKeys();
  resetCurrentMeetingAfterStorageClear();
  renderStorageSettingsStatus();
  showRecentMeetingFeedback("백업 파일을 내려받고 브라우저 저장소를 비웠습니다.");
  toast("백업 후 로컬 저장소를 삭제했습니다.");
}

function resetMeeting() {
  if (!confirm("현재 미팅 내용을 새 미팅으로 초기화할까요? 저장된 최근 미팅 목록은 유지됩니다.")) return;
  selectedFile = null;
  state = emptyState();
  applyStateToFields();
  saveState();
  renderAll();
  toast("새 미팅을 시작합니다.");
}

function mergeSavedState(saved) {
  const meeting = { ...emptyState().meeting, ...(saved.meeting || {}) };
  if (!meeting.gpParticipants && meeting.contactName) meeting.gpParticipants = meeting.contactName;
  return {
    ...emptyState(),
    ...saved,
    meeting,
    documentSettings: { ...emptyState().documentSettings, ...(saved.documentSettings || {}) },
    questionRecords: Array.isArray(saved.questionRecords) ? saved.questionRecords : []
  };
}

function sanitizeStateForStorage(sourceState) {
  const safe = JSON.parse(JSON.stringify(sourceState || state));
  delete safe.apiKey;
  delete safe.geminiApiKey;
  delete safe.selectedFile;
  safe.selectedFileBlob = null;
  return safe;
}

function createMeetingId() {
  if (window.crypto?.randomUUID) return window.crypto.randomUUID();
  return `meeting-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function showLoader(title, message) {
  setBusy(true);
  $("loaderTitle").textContent = title;
  $("loaderMessage").textContent = message;
  $("globalLoader").classList.remove("hidden");
  $("globalLoader").classList.add("flex");
  refreshIcons();
}

function updateLoader(title, message) {
  $("loaderTitle").textContent = title;
  $("loaderMessage").textContent = message;
}

function hideLoader() {
  $("globalLoader").classList.add("hidden");
  $("globalLoader").classList.remove("flex");
  setBusy(false);
}

function setBusy(nextBusy) {
  isBusy = nextBusy;
  [
    "generateBriefButton",
    "generateReportButton",
    "regenerateReportButton",
    "saveButton",
    "savePrepButton",
    "exportButton",
    "exportDocxButton",
    "copyAllButton"
  ].forEach((id) => {
    const element = $(id);
    if (element) element.disabled = nextBusy;
  });
}

function showError(error) {
  const message = normalizeGeminiError(error);
  toast(message);
}

function normalizeGeminiError(error) {
  const message = error?.message || "";
  const info = error?.gemini;
  if (info?.message) return info.message;
  if (/API Key가 아직 적용되지 않았습니다/i.test(message)) return message;
  if (/기본 모델과 Flash Lite 재시도 모두 실패했습니다/i.test(message)) {
    return message;
  }
  if (/무료 한도|분당 한도|유효하지 않거나|권한|모델을 사용할 수 없습니다/i.test(message)) {
    return message;
  }
  if (/503|UNAVAILABLE|high demand|사용량이 많습니다/i.test(message)) {
    return `${lastGeminiModelUsed} 모델이 일시적으로 혼잡합니다. 잠시 후 다시 시도하거나 Flash Lite 모델을 선택하세요.`;
  }
  if (
    message === LIMIT_MESSAGE ||
    /quota|429|api key|permission|unauth|model|billing|resource_exhausted/i.test(message)
  ) {
    return LIMIT_MESSAGE;
  }
  return message || LIMIT_MESSAGE;
}

function toast(message) {
  const box = $("toast");
  box.textContent = message;
  box.classList.remove("hidden");
  clearTimeout(toast.timer);
  toast.timer = setTimeout(() => box.classList.add("hidden"), 3200);
}

function refreshIcons() {
  if (window.lucide) window.lucide.createIcons();
}

function asArray(value) {
  if (Array.isArray(value)) return value;
  if (!value) return [];
  return [value];
}

function renderList(items) {
  const array = asArray(items).filter(Boolean);
  if (!array.length) return emptyText();
  return `<ul>${array.map((item) => `<li>${formatListItemHtml(item)}</li>`).join("")}</ul>`;
}

function renderInlineList(items) {
  return asArray(items).map((item) => escapeHtml(String(item))).join(", ");
}

function listMarkdown(items) {
  const array = asArray(items).filter(Boolean);
  return array.length ? array.map((item) => `- ${formatListItemText(item)}`).join("\n") : "- 확인 필요";
}

function formatListItemHtml(item) {
  return escapeHtml(formatListItemText(item));
}

function formatNestedListValue(value) {
  if (Array.isArray(value)) return value.map(formatListItemText).filter(Boolean).join(", ");
  if (value && typeof value === "object") return formatListItemText(value);
  return String(value ?? "");
}

function formatListItemText(item) {
  if (Array.isArray(item)) return item.map(formatListItemText).filter(Boolean).join(", ");
  if (typeof item === "string") return item;
  if (!item || typeof item !== "object") return String(item ?? "");
  if (item.riskType || item.assessment || item.details) {
    return [
      formatNestedListValue(item.riskType),
      item.assessment ? `(${formatNestedListValue(item.assessment)})` : "",
      item.details ? `- ${formatNestedListValue(item.details)}` : ""
    ].filter(Boolean).join(" ");
  }
  if (item.request || item.rationale || item.category) {
    return [
      item.category ? `[${formatNestedListValue(item.category)}]` : "",
      formatNestedListValue(item.request || item.title || item.question || ""),
      item.importance ? `(${formatNestedListValue(item.importance)})` : "",
      item.rationale ? `- ${formatNestedListValue(item.rationale)}` : ""
    ].filter(Boolean).join(" ");
  }
  if (item.title || item.note) {
    return [formatNestedListValue(item.title), item.note ? `- ${formatNestedListValue(item.note)}` : ""].filter(Boolean).join(" ");
  }
  return Object.entries(item)
    .filter(([, value]) => value !== undefined && value !== null && String(formatNestedListValue(value)).trim())
    .map(([key, value]) => `${labelize(key)}: ${formatNestedListValue(value)}`)
    .join(" / ");
}

function emptyText() {
  return `<p class="text-sm text-slate-400">확인 필요</p>`;
}

function hasAnswer(record) {
  return record.answer?.trim() || record.internalMemo?.trim() || record.importantQuote?.trim() || record.followUpNeeded;
}

function deriveAnswerRecordStatus(record = {}) {
  if (record.followUpNeeded) return "추가 검토 필요";
  if (record.answer?.trim()) return "답변됨";
  return "미확인";
}

function capitalize(value) {
  return value.charAt(0).toUpperCase() + value.slice(1);
}

function labelize(key) {
  const labels = {
    managerName: "운용사명",
    fundName: "펀드명 / 대출명",
    gpParticipants: "GP측 참가자",
    lpParticipants: "당사 참가자",
    contactName: "GP측 참가자",
    meetingDate: "미팅일",
    locationType: "투자지역",
    assetClass: "자산군",
    strategy: "전략",
    sector: "섹터",
    capitalType: "Equity / Debt",
    investmentStructure: "투자구조",
    keyConcerns: "딜 / 자산 메모",
    capitalStructure: "투자구조",
    region: "지역",
    targetSize: "펀드 규모",
    loanSize: "대출 규모",
    investmentPeriod: "투자 기간",
    loanMaturity: "대출만기",
    targetReturn: "목표 수익률",
    loanRate: "대출금리",
    commitmentAmount: "당사 검토 약정액",
    keyNumbers: "핵심 숫자"
  };
  return labels[key] || key;
}

function markdownToHtml(markdown) {
  const lines = markdown.split("\n");
  let html = "";
  let inList = false;
  lines.forEach((line) => {
    if (line.startsWith("# ")) {
      if (inList) { html += "</ul>"; inList = false; }
      html += `<h2>${escapeHtml(line.slice(2))}</h2>`;
    } else if (line.startsWith("## ")) {
      if (inList) { html += "</ul>"; inList = false; }
      html += `<h3>${escapeHtml(line.slice(3))}</h3>`;
    } else if (line.startsWith("- ")) {
      if (!inList) { html += "<ul>"; inList = true; }
      html += `<li>${formatBold(escapeHtml(line.slice(2)))}</li>`;
    } else if (line.trim()) {
      if (inList) { html += "</ul>"; inList = false; }
      html += `<p>${formatBold(escapeHtml(line))}</p>`;
    }
  });
  if (inList) html += "</ul>";
  return html;
}

function formatBold(text) {
  return text.replace(/\*\*(.*?)\*\*/g, "<strong>$1</strong>");
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function escapeAttribute(value) {
  return escapeHtml(value).replace(/`/g, "&#096;");
}

function formatBytes(bytes) {
  if (!bytes) return "0 B";
  const units = ["B", "KB", "MB", "GB"];
  const index = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1);
  return `${(bytes / Math.pow(1024, index)).toFixed(index ? 1 : 0)} ${units[index]}`;
}

function formatNumber(value) {
  return new Intl.NumberFormat("ko-KR").format(value || 0);
}

function formatDateTime(value) {
  if (!value) return "시간 미상";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "시간 미상";
  return new Intl.DateTimeFormat("ko-KR", {
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit"
  }).format(date);
}

function safeFileName(value) {
  return String(value).replace(/[\\/:*?"<>|]/g, "_").trim() || "meeting";
}

function makeFileMeta(file) {
  return {
    fileName: file.name,
    size: file.size,
    sizeLabel: formatBytes(file.size),
    mimeType: file.type || guessMimeType(file.name)
  };
}

function guessMimeType(fileName) {
  const lower = fileName.toLowerCase();
  if (lower.endsWith(".pdf")) return "application/pdf";
  if (lower.endsWith(".txt")) return "text/plain";
  if (lower.endsWith(".png")) return "image/png";
  if (lower.endsWith(".jpg") || lower.endsWith(".jpeg")) return "image/jpeg";
  return "application/octet-stream";
}

function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result).split(",")[1]);
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

function parsePageRanges(input, maxPage) {
  const pages = new Set();
  String(input || "").split(",").forEach((part) => {
    const trimmed = part.trim();
    if (!trimmed) return;
    if (trimmed.includes("-")) {
      const [start, end] = trimmed.split("-").map((value) => Number(value.trim()));
      if (!Number.isFinite(start) || !Number.isFinite(end)) return;
      for (let page = Math.max(1, start); page <= Math.min(maxPage, end); page += 1) pages.add(page);
    } else {
      const page = Number(trimmed);
      if (Number.isFinite(page) && page >= 1 && page <= maxPage) pages.add(page);
    }
  });
  return [...pages].sort((a, b) => a - b);
}
