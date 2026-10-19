require("dotenv").config();
const express = require("express");
const multer = require("multer");
const xlsx = require("xlsx");
const axios = require("axios");
const path = require("path");
const crypto = require("crypto");
const iconv = require("iconv-lite");

const app = express();
const PORT = process.env.PORT || 3000;
const EXCEL_MIME_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const JUSO_API_TIMEOUT_MS = 3000;
const DIAGNOSTIC_OLD_SEARCH = process.env.DIAGNOSTIC_OLD_SEARCH === "true";
const DIAGNOSTIC_VERBOSE = process.env.DIAGNOSTIC_VERBOSE === "true";
const DIAGNOSTIC_SLOW_ROW_MS = 3000;
const DIAGNOSTIC_RETRY_HEAVY_COUNT = 4;
const PROCESS_STATUS_TOP_ROWS_LIMIT = 3;
const processJobs = new Map();
let latestJobId = null;

const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 10 * 1024 * 1024, // 10MB 제한
  },
});

app.use(express.static("public"));
app.use(express.json());

const JUSO_API_URL = "https://business.juso.go.kr/addrlink/addrLinkApi.do";
const JUSO_API_KEY = process.env.JUSO_CONFIRM_KEY;

function asyncHandler(handler) {
  return (req, res, next) => {
    Promise.resolve(handler(req, res, next)).catch(next);
  };
}

function sendApiError(res, status, message, code) {
  const payload = { error: message };
  if (code) {
    payload.code = code;
  }
  return res.status(status).json(payload);
}

function logApiError(context, error, details = {}) {
  const errorName = error?.name || "Error";
  const message = error?.message || String(error);
  console.error(`[${context}]`, {
    name: errorName,
    message,
    ...details,
  });
}

function preprocessAddress(address) {
  // "외N필지" 또는 "외 N필지" 패턴이 있으면 그 앞부분만 남김
  return String(address)
    .replace(/외\s*\d+\s*필지.*$/, "")
    .trim();
}

function tokenizeForBackoff(text) {
  const raw = String(text).trim().split(/\s+/);
  const tokens = [];

  for (const t of raw) {
    const m = t.match(/^(\d+(?:-\d+)?)([^\d].+)$/);
    if (m) {
      tokens.push(m[1], m[2]);
    } else {
      tokens.push(t);
    }
  }

  return tokens;
}

// 원본 주소에 명시된 "동/읍/면/리" 지역명 후보를 추출 (예: "청라동", "가좌동")
// 숫자가 섞인 토큰("120동", "7동" 같은 아파트 동/호)은 제외한다.
function extractDongCandidates(text) {
  return String(text)
    .trim()
    .split(/\s+/)
    .filter((t) => /^[가-힣]{2,}$/.test(t) && /(동|읍|면|리)$/.test(t));
}

// backoff로 검색어가 짧아지는 과정에서 원본 동 정보가 이미 잘려나갔다면,
// 검색 결과가 실제로 그 동과 일치하는지 확인한다.
// (구 이름 등 남은 짧은 검색어가 우연히 건물명 텍스트와 매치되는 오탐을 걸러내기 위함)
function isPlausibleBackoffMatch(dongCandidates, shortenedKeyword, result) {
  if (dongCandidates.length === 0) return true;
  if (dongCandidates.some((d) => shortenedKeyword.includes(d))) return true;

  return dongCandidates.some(
    (d) =>
      result.emdNm === d ||
      result.jibunAddr?.includes(d) ||
      result.roadAddr?.includes(d)
  );
}

// 시/군/구 개편(예: 인천 서구 -> 서해구)으로 옛 구 이름이 더 이상 검색되지 않는 경우를 위해
// 구/군 토큰 하나를 제거한 검색어를 만든다. 시/도("~시","~도") 토큰과 동/읍/면/리 토큰은 건드리지 않는다.
function stripSingleGuGunToken(text) {
  const tokens = String(text).trim().split(/\s+/);
  const idx = tokens.findIndex((t) => /^[가-힣]{2,}$/.test(t) && /(구|군)$/.test(t));
  if (idx === -1) return null;

  return [...tokens.slice(0, idx), ...tokens.slice(idx + 1)].join(" ").trim();
}

// 인천 주소는 구 개편(서구->서해구, 중구/동구->영종구/제물포구/검단구 등)이 잦아
// tail-backoff가 번지까지 지우기 전에 미리 구 이름을 제거해둔다.
// ("인천광역시"/"인천" 토큰 바로 다음에 오는 구/군 토큰만 정확히 겨냥해서 제거)
function stripIncheonGuToken(text) {
  const tokens = String(text).trim().split(/\s+/);
  const siIdx = tokens.findIndex((t) => t === "인천광역시" || t === "인천");
  if (siIdx === -1) return text;

  const guToken = tokens[siIdx + 1];
  if (!guToken || !(/^[가-힣]{2,}$/.test(guToken) && /(구|군)$/.test(guToken))) {
    return text;
  }

  return [...tokens.slice(0, siIdx + 1), ...tokens.slice(siIdx + 2)]
    .join(" ")
    .trim();
}

/**
 * ✅ 추가: "제1동" 처럼 '제'+숫자1개+'동' 이면 '제'만 제거
 * - 예: "화곡 제1동" -> "화곡 1동"
 * - 예: "제2동" -> "2동"
 */
function removeJePrefixIfNeeded(value) {
  if (value === null || value === undefined) return "";
  const s = String(value).trim();
  if (s === "") return "";

  // 단어 경계 기준으로 "제1동" 형태만 "1동"으로 치환
  // (숫자 1개만 허용: 제0동~제9동)
  return s.replace(/제([0-9])동/g, "$1동");
}

function addDongSuffix(value) {
  const trimmed = String(value ?? "").trim();
  if (!trimmed || /[동읍면]$/.test(trimmed)) return trimmed;
  return trimmed + "동";
}

function normalizeDongOfficeHeaders(worksheet) {
  if (!worksheet["!ref"]) return;
  const range = xlsx.utils.decode_range(worksheet["!ref"]);
  const matches = [];

  for (let column = range.s.c; column <= range.e.c; column++) {
    const cell = worksheet[xlsx.utils.encode_cell({ r: range.s.r, c: column })];
    if (!cell) continue;
    const normalized = String(cell.v ?? "").replace(/\s+/g, "");
    if (normalized.includes("열람동사무소") && !normalized.endsWith("_원본")) {
      matches.push({ cell, normalized });
    }
  }

  if (matches.length > 1) {
    throw Object.assign(new Error("열람 동사무소 컬럼이 중복되어 있습니다. 하나만 남겨주세요."), {
      status: 400,
      code: "DONG_OFFICE_COLUMN_DUPLICATED",
    });
  }

  for (const { cell, normalized } of matches) {
    cell.v = normalized;
    cell.t = "s";
    delete cell.w;
  }
}

function getAddressErrorMessage(error) {
  if (error?.code === "ECONNABORTED") {
    return "주소 조회 시간 초과";
  }

  return error?.message || "주소 조회 중 오류가 발생했습니다";
}

function normalizeSearchResult(searchResult) {
  if (searchResult.success) {
    return "성공";
  }

  const error = String(searchResult.error || "").trim();

  if (!error || error === "정상" || error === "주소를 찾을 수 없습니다") {
    return "미매칭";
  }

  if (error.includes("검색어가 너무 깁니다")) {
    return "검색어 초과";
  }

  if (error === "주소 조회 시간 초과") {
    return "조회 시간 초과";
  }

  if (error === "주소가 비어 있습니다") {
    return "주소 없음";
  }

  return "조회 오류";
}

function createDiagnosticContext(req, totalRows) {
  const requestId = `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;

  return {
    enabled: DIAGNOSTIC_OLD_SEARCH,
    verbose: DIAGNOSTIC_VERBOSE,
    requestId,
    startedAt: Date.now(),
    totalRows,
    fileName: req.file?.originalname || "",
    lastStartedRow: null,
    lastCompletedRow: null,
    rowSummaries: [],
    failedRowsCount: 0,
  };
}

function logDiagnostic(context, message, details = {}) {
  if (!context?.enabled) {
    return;
  }

  console.log(`[diagnostic:${context.requestId}] ${message}`, details);
}

function getTopRows(rowSummaries, field, limit = 3) {
  return [...rowSummaries]
    .filter((row) => Number(row[field]) > 0)
    .sort((a, b) => b[field] - a[field])
    .slice(0, limit)
    .map(
      ({
        rowNumber,
        sourceRowNumber,
        status,
        elapsedMs,
        attemptCount,
        address,
      }) => ({
        rowNumber,
        sourceRowNumber,
        status,
        elapsedMs,
        attemptCount,
        address,
      })
    );
}

function cloneTopRows(rowSummaries) {
  return {
    slowestRows: getTopRows(
      rowSummaries,
      "elapsedMs",
      PROCESS_STATUS_TOP_ROWS_LIMIT
    ),
    mostRetriedRows: getTopRows(
      rowSummaries,
      "attemptCount",
      PROCESS_STATUS_TOP_ROWS_LIMIT
    ),
  };
}

function createJobId() {
  return crypto.randomUUID();
}

function createJobState({ fileName, mode }) {
  const jobId = createJobId();

  return {
    jobId,
    status: "queued",
    phase: "queued",
    mode,
    fileName,
    totalRows: 0,
    completedRows: 0,
    fastCompletedRows: 0,
    preciseTargetRows: 0,
    preciseCompletedRows: 0,
    startedAt: null,
    elapsedMs: 0,
    lastStartedRow: null,
    lastCompletedRow: null,
    lastAddress: "",
    failedRowsCount: 0,
    rowSummaries: [],
    resultBuffer: null,
    resultFileName: null,
    error: null,
  };
}

function inferLikelyCause(status) {
  if (!status) {
    return "처리 상태 정보가 없습니다.";
  }

  const elapsedMs = status.elapsedMs || 0;
  const mostRetried = status.mostRetriedRows || [];
  const slowestRows = status.slowestRows || [];
  const lastCompletedRow = status.lastCompletedRow || 0;

  if (
    elapsedMs >= 30000 &&
    mostRetried.some((row) => (row.attemptCount || 0) >= DIAGNOSTIC_RETRY_HEAVY_COUNT)
  ) {
    return "주소 재검색이 누적되어 시간 초과가 발생한 것으로 보입니다.";
  }

  if (
    slowestRows.some((row) => (row.elapsedMs || 0) >= DIAGNOSTIC_SLOW_ROW_MS * 2)
  ) {
    return "일부 주소 행 처리 시간이 길어져 병목이 발생한 것으로 보입니다.";
  }

  if (status.status === "failed" && lastCompletedRow <= 5) {
    return "파일 형식이나 초기 데이터 문제로 초반에 중단된 것으로 보입니다.";
  }

  if (status.status === "running" || elapsedMs >= 30000) {
    return "요청 처리 시간이 길어 앞단 프록시에서 502를 반환했을 가능성이 높습니다.";
  }

  return "일시적인 서버 또는 네트워크 문제일 수 있습니다.";
}

function buildJobSnapshot(job, overrides = {}) {
  const rowSummaries = overrides.rowSummaries || job.rowSummaries || [];
  const snapshot = {
    ...job,
    ...overrides,
    rowSummaries,
    elapsedMs:
      overrides.elapsedMs ??
      ((overrides.startedAt ?? job.startedAt)
        ? Date.now() - (overrides.startedAt ?? job.startedAt)
        : 0),
  };

  const topRows = cloneTopRows(rowSummaries);
  snapshot.slowestRows = topRows.slowestRows;
  snapshot.mostRetriedRows = topRows.mostRetriedRows;
  snapshot.likelyCause = inferLikelyCause(snapshot);
  return snapshot;
}

function saveJob(job) {
  processJobs.set(job.jobId, job);
  latestJobId = job.jobId;
  return job;
}

function updateJob(jobId, overrides = {}) {
  const job = processJobs.get(jobId);
  if (!job) {
    return null;
  }

  const nextJob = buildJobSnapshot(job, overrides);
  processJobs.set(jobId, nextJob);
  latestJobId = jobId;
  return nextJob;
}

function getJob(jobId) {
  return processJobs.get(jobId) || null;
}

function getLatestJob() {
  return latestJobId ? getJob(latestJobId) : null;
}

function clearPreviousJobs() {
  processJobs.clear();
  latestJobId = null;

  if (typeof global.gc === "function") {
    global.gc();
  }
}

function serializeJob(job) {
  if (!job) {
    return null;
  }

  return {
    jobId: job.jobId,
    status: job.status,
    phase: job.phase,
    mode: job.mode,
    fileName: job.fileName,
    totalRows: job.totalRows,
    completedRows: job.completedRows,
    fastCompletedRows: job.fastCompletedRows,
    preciseTargetRows: job.preciseTargetRows,
    preciseCompletedRows: job.preciseCompletedRows,
    elapsedMs: job.elapsedMs,
    lastStartedRow: job.lastStartedRow,
    lastCompletedRow: job.lastCompletedRow,
    lastAddress: job.lastAddress,
    failedRowsCount: job.failedRowsCount,
    slowestRows: job.slowestRows,
    mostRetriedRows: job.mostRetriedRows,
    likelyCause: job.likelyCause,
    error: job.error,
    downloadReady: Boolean(job.resultBuffer),
  };
}

function shouldRetryInPrecise(status) {
  return (
    status === "미매칭" ||
    status === "검색어 초과" ||
    status === "조회 시간 초과" ||
    status === "조회 오류"
  );
}

function buildResultRow(row, searchResult, dongOfficeColumn) {
  const resultRow = { ...row };

  if (searchResult) {
    resultRow.관할동 = searchResult.success ? searchResult.emdNm : "";
    resultRow.시군구 = searchResult.success ? searchResult.sggNm : "";
    resultRow.시도 = searchResult.success ? searchResult.siNm : "";
    resultRow.도로명주소 = searchResult.success ? searchResult.roadAddr : "";

    const dongOfficeRaw = searchResult.success
      ? (searchResult.dong || "").split(/\s+/).pop() || ""
      : "";

    if (/제동$/.test(resultRow.관할동)) {
      resultRow.관할동사무소 = dongOfficeRaw.replace(
        /제제([0-9])동/g,
        "제$1동"
      );
    } else {
      resultRow.관할동사무소 = removeJePrefixIfNeeded(dongOfficeRaw);
    }

    resultRow.검색결과 = normalizeSearchResult(searchResult);
  } else {
    resultRow.관할동 = "";
    resultRow.시군구 = "";
    resultRow.시도 = "";
    resultRow.도로명주소 = "";
    resultRow.관할동사무소 = "";
    resultRow.검색결과 = "주소 없음";
  }

  if (dongOfficeColumn) {
    const originalValue = row[dongOfficeColumn];
    const officeValue = addDongSuffix(originalValue);
    const gwanhalDong = resultRow.관할동 || "";

    resultRow[`${dongOfficeColumn}_원본`] = originalValue;

    if (/제동$/.test(gwanhalDong)) {
      resultRow[dongOfficeColumn] = officeValue.replace(
        /제제([0-9])동/g,
        "제$1동"
      );
    } else {
      resultRow[dongOfficeColumn] = removeJePrefixIfNeeded(officeValue);
    }
  }

  return resultRow;
}

// 단일 API 호출
async function searchAddressOnce(keyword, countPerPage = 1) {
  try {
    const response = await axios.get(JUSO_API_URL, {
      params: {
        confmKey: JUSO_API_KEY,
        currentPage: 1,
        countPerPage,
        keyword: keyword,
        resultType: "json",
        addInfoYn: "Y",
      },
      timeout: JUSO_API_TIMEOUT_MS,
    });

    const result = response.data?.results;
    const totalCount = Number(result?.common?.totalCount || 0);
    if (result?.common?.errorCode === "0" && result?.juso?.length > 0) {
      const juso = result.juso[0];
      return {
        success: true,
        emdNm: juso.emdNm,
        sggNm: juso.sggNm,
        siNm: juso.siNm,
        roadAddr: juso.roadAddr,
        jibunAddr: juso.jibunAddr,
        dong: juso.hemdNm,
        totalCount,
      };
    }
    return {
      success: false,
      error: result?.common?.errorMessage || "주소를 찾을 수 없습니다",
      totalCount,
    };
  } catch (error) {
    return { success: false, error: getAddressErrorMessage(error) };
  }
}

// baseKeyword로 시작해 뒤 토큰을 하나씩 잘라가며 재검색한다.
// requireUnique가 true면(구/군 정보를 뺀 검색) countPerPage를 늘려 후보가 유일할 때만 채택한다.
async function runBackoffChain(
  baseKeyword,
  dongCandidates,
  diagnosticContext,
  startAttempt,
  { requireUnique = false, phase = "with-gu" } = {}
) {
  const countPerPage = requireUnique ? 5 : 1;
  let attempt = startAttempt;

  const tryKeyword = async (keyword) => {
    const result = await searchAddressOnce(keyword, countPerPage);
    const plausible =
      result.success &&
      (!requireUnique || result.totalCount === 1) &&
      isPlausibleBackoffMatch(dongCandidates, keyword, result);

    if (diagnosticContext?.verbose) {
      logDiagnostic(diagnosticContext, "search attempt", {
        attempt,
        phase,
        keyword,
        success: result.success,
        totalCount: result.totalCount,
        rejectedAsImplausible: result.success && !plausible,
        error: result.success ? null : result.error,
      });
    }

    return { result, plausible };
  };

  let { result, plausible } = await tryKeyword(baseKeyword);
  if (plausible) {
    return { success: true, result, attempt };
  }

  const tokens = tokenizeForBackoff(baseKeyword);
  while (tokens.length > 1) {
    tokens.pop();
    attempt += 1;
    const shortened = tokens.join(" ").trim();
    ({ result, plausible } = await tryKeyword(shortened));
    if (plausible) {
      return { success: true, result, attempt };
    }
  }

  return { success: false, result, attempt };
}

async function searchAddressWithBackoff(keyword, diagnosticContext) {
  const preprocessed = preprocessAddress(keyword);
  if (!preprocessed) {
    return { success: false, error: "주소가 비어 있습니다", attemptCount: 0 };
  }

  // 인천 주소는 번지를 지우기 전에 구 이름부터 제거해, tail-backoff가
  // "인천서구청" 같은 옛 구 이름이 들어간 건물명에 낚이는 것을 원천 차단한다.
  const processed = stripIncheonGuToken(preprocessed);

  const dongCandidates = extractDongCandidates(processed);

  const primary = await runBackoffChain(processed, dongCandidates, diagnosticContext, 1, {
    requireUnique: false,
    phase: "with-gu",
  });

  if (primary.success) {
    return { ...primary.result, attemptCount: primary.attempt };
  }

  // 구/군 이름이 개편되어 더 이상 검색되지 않는 경우를 대비해,
  // 구/군 토큰을 제거하고 한 번 더 시도한다. 정보 손실로 인한 오매칭을 막기 위해
  // 검색 결과가 유일한 후보일 때(totalCount === 1)만 채택한다.
  const guDropped = stripSingleGuGunToken(processed);
  const finalAttempt = guDropped
    ? await runBackoffChain(guDropped, dongCandidates, diagnosticContext, primary.attempt + 1, {
        requireUnique: true,
        phase: "gu-dropped",
      })
    : primary;

  if (finalAttempt.success) {
    return { ...finalAttempt.result, attemptCount: finalAttempt.attempt };
  }

  return {
    success: false,
    error: "주소를 찾을 수 없습니다",
    attemptCount: finalAttempt.attempt,
  };
}

async function searchAddress(keyword, diagnosticContext) {
  const processed = preprocessAddress(keyword);
  if (!processed) {
    return { success: false, error: "주소가 비어 있습니다", attemptCount: 0 };
  }

  if (diagnosticContext?.mode === "precise") {
    return searchAddressWithBackoff(processed, diagnosticContext);
  }

  const result = await searchAddressOnce(processed);
  return { ...result, attemptCount: 1 };
}

function buildDownloadFileName(originalName) {
  return `result_${originalName.replace(/\.(xlsx|xls|csv)$/i, ".xlsx")}`;
}

function sanitizeDownloadFileName(name) {
  const normalized = String(name || "result.xlsx")
    .replace(/[^\x20-\x7E]/g, "_")
    .replace(/[\"\\;,\r\n]/g, "_")
    .replace(/\s+/g, "_")
    .replace(/_+/g, "_")
    .replace(/^_+|_+$/g, "");

  return normalized || "result.xlsx";
}

function encodeRFC5987ValueChars(value) {
  return encodeURIComponent(String(value || "result.xlsx"))
    .replace(/['()]/g, escape)
    .replace(/\*/g, "%2A");
}

async function processExcelJob(jobId, fileBuffer, originalName, mode) {
  const job = getJob(jobId);
  if (!job) {
    return;
  }

  const diagnosticContext = {
    enabled: DIAGNOSTIC_OLD_SEARCH,
    verbose: DIAGNOSTIC_VERBOSE,
    requestId: jobId,
    startedAt: Date.now(),
    totalRows: 0,
    fileName: originalName,
    lastStartedRow: null,
    lastCompletedRow: null,
    rowSummaries: [],
    failedRowsCount: 0,
    mode: "fast",
  };

  const statusRowSummaries = [];

  try {
    updateJob(jobId, {
      status: "running",
      phase: "fast",
      startedAt: diagnosticContext.startedAt,
      elapsedMs: 0,
      fastCompletedRows: 0,
      preciseTargetRows: 0,
      preciseCompletedRows: 0,
      rowSummaries: statusRowSummaries,
      error: null,
    });

    let normalizedBuffer = fileBuffer;
    if (originalName.toLowerCase().endsWith(".csv")) {
      const content = iconv.decode(fileBuffer, "cp949");
      normalizedBuffer = Buffer.from(content, "utf-8");
    }

    const workbook = xlsx.read(normalizedBuffer, { type: "buffer" });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    normalizeDongOfficeHeaders(worksheet);
    const data = xlsx.utils.sheet_to_json(worksheet);
    diagnosticContext.totalRows = data.length;

    updateJob(jobId, {
      totalRows: data.length,
      rowSummaries: statusRowSummaries,
    });

    if (data.length === 0) {
      throw Object.assign(new Error("파일에 데이터가 없습니다"), {
        status: 400,
        code: "EMPTY_FILE",
      });
    }

    const addressColumn = Object.keys(data[0]).find(
      (key) =>
        key.toLowerCase().includes("주소") ||
        key.toLowerCase().includes("address")
    );

    if (!addressColumn) {
      throw Object.assign(new Error("'주소' 컬럼을 찾을 수 없습니다"), {
        status: 400,
        code: "ADDRESS_COLUMN_MISSING",
      });
    }

    const dongOfficeColumn = Object.keys(data[0]).find((key) =>
      key.includes("열람동사무소") && !key.endsWith("_원본")
    );

    const results = [];
    const retryIndexes = [];
    const addressCache = new Map();
    let currentFailedRowsCount = 0;
    logDiagnostic(diagnosticContext, "request started", {
      totalRows: data.length,
      fileName: diagnosticContext.fileName,
      mode: "auto-2pass",
    });

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const address = row[addressColumn];
      const rowNumber = i + 1;
      const sourceRowNumber = row["연번"] ?? null;
      const rowStartedAt = Date.now();
      let attemptCount = 0;

      diagnosticContext.lastStartedRow = rowNumber;
      updateJob(jobId, {
        phase: "fast",
        lastStartedRow: rowNumber,
        lastAddress: String(address || ""),
        fastCompletedRows: i,
        rowSummaries: statusRowSummaries,
      });

      let searchResult = null;
      if (address) {
        const cacheKey = `fast:${preprocessAddress(address)}`;
        searchResult = addressCache.get(cacheKey);

        if (!searchResult) {
          searchResult = await searchAddress(address, { ...diagnosticContext, mode: "fast" });
          addressCache.set(cacheKey, searchResult);
        }
        attemptCount = searchResult.attemptCount || 0;
      }

      const resultRow = buildResultRow(row, searchResult, dongOfficeColumn);

      results.push(resultRow);
      diagnosticContext.lastCompletedRow = rowNumber;
      const rowSummary = {
        rowNumber,
        sourceRowNumber,
        address,
        status: resultRow.검색결과,
        elapsedMs: Date.now() - rowStartedAt,
        attemptCount,
      };
      diagnosticContext.rowSummaries.push(rowSummary);
      statusRowSummaries.push(rowSummary);

      if (shouldRetryInPrecise(rowSummary.status)) {
        retryIndexes.push(i);
      }

      if (rowSummary.status !== "성공") {
        diagnosticContext.failedRowsCount += 1;
        currentFailedRowsCount += 1;
      }

      updateJob(jobId, {
        completedRows: rowNumber,
        fastCompletedRows: rowNumber,
        lastCompletedRow: rowNumber,
        lastAddress: String(address || ""),
        failedRowsCount: currentFailedRowsCount,
        rowSummaries: statusRowSummaries,
      });

      if (rowSummary.elapsedMs >= DIAGNOSTIC_SLOW_ROW_MS) {
        logDiagnostic(diagnosticContext, "slow row", rowSummary);
      } else if (rowSummary.attemptCount >= DIAGNOSTIC_RETRY_HEAVY_COUNT) {
        logDiagnostic(diagnosticContext, "retry-heavy row", rowSummary);
      } else if (
        rowSummary.status !== "성공" &&
        rowSummary.status !== "주소 없음"
      ) {
        logDiagnostic(diagnosticContext, "failed row", rowSummary);
      }
    }

    updateJob(jobId, {
      phase: "precise",
      preciseTargetRows: retryIndexes.length,
      preciseCompletedRows: 0,
      rowSummaries: statusRowSummaries,
    });

    for (let j = 0; j < retryIndexes.length; j++) {
      const index = retryIndexes[j];
      const row = data[index];
      const address = row[addressColumn];
      const rowNumber = index + 1;
      const rowStartedAt = Date.now();
      let attemptCount = 0;

      diagnosticContext.lastStartedRow = rowNumber;
      updateJob(jobId, {
        phase: "precise",
        lastStartedRow: rowNumber,
        lastAddress: String(address || ""),
        preciseCompletedRows: j,
        rowSummaries: statusRowSummaries,
      });

      const cacheKey = `precise:${preprocessAddress(address)}`;
      let searchResult = addressCache.get(cacheKey);

      if (!searchResult) {
        searchResult = await searchAddress(address, {
          ...diagnosticContext,
          mode: "precise",
        });
        addressCache.set(cacheKey, searchResult);
      }
      attemptCount = searchResult.attemptCount || 0;

      const resultRow = buildResultRow(row, searchResult, dongOfficeColumn);
      const previousStatus = results[index].검색결과;
      results[index] = resultRow;
      diagnosticContext.lastCompletedRow = rowNumber;

      const rowSummary = {
        rowNumber,
        sourceRowNumber: row["연번"] ?? null,
        address,
        status: `${resultRow.검색결과} (2차)`,
        elapsedMs: Date.now() - rowStartedAt,
        attemptCount,
      };
      diagnosticContext.rowSummaries.push(rowSummary);
      statusRowSummaries.push(rowSummary);

      if (previousStatus !== "성공" && resultRow.검색결과 === "성공") {
        currentFailedRowsCount = Math.max(0, currentFailedRowsCount - 1);
      }

      updateJob(jobId, {
        completedRows: data.length + (j + 1),
        preciseCompletedRows: j + 1,
        lastCompletedRow: rowNumber,
        lastAddress: String(address || ""),
        failedRowsCount: currentFailedRowsCount,
        rowSummaries: statusRowSummaries,
      });

      if (rowSummary.elapsedMs >= DIAGNOSTIC_SLOW_ROW_MS) {
        logDiagnostic(diagnosticContext, "slow row", rowSummary);
      } else if (rowSummary.attemptCount >= DIAGNOSTIC_RETRY_HEAVY_COUNT) {
        logDiagnostic(diagnosticContext, "retry-heavy row", rowSummary);
      } else if (resultRow.검색결과 !== "성공") {
        logDiagnostic(diagnosticContext, "failed row", rowSummary);
      }
    }

    const newWorkbook = xlsx.utils.book_new();
    const newWorksheet = xlsx.utils.json_to_sheet(results);
    xlsx.utils.book_append_sheet(newWorkbook, newWorksheet, "Result");

    const resultBuffer = xlsx.write(newWorkbook, {
      type: "buffer",
      bookType: "xlsx",
    });

    logDiagnostic(diagnosticContext, "request completed summary", {
      totalRows: data.length,
      totalElapsedMs: Date.now() - diagnosticContext.startedAt,
      lastStartedRow: diagnosticContext.lastStartedRow,
      lastCompletedRow: diagnosticContext.lastCompletedRow,
      failedRowsCount: diagnosticContext.failedRowsCount,
      slowestRows: getTopRows(diagnosticContext.rowSummaries, "elapsedMs"),
      mostRetriedRows: getTopRows(
        diagnosticContext.rowSummaries,
        "attemptCount"
      ),
    });

    updateJob(jobId, {
      status: "completed",
      phase: "completed",
      totalRows: data.length,
      completedRows: data.length + retryIndexes.length,
      fastCompletedRows: data.length,
      preciseTargetRows: retryIndexes.length,
      preciseCompletedRows: retryIndexes.length,
      elapsedMs: Date.now() - diagnosticContext.startedAt,
      lastStartedRow: diagnosticContext.lastStartedRow,
      lastCompletedRow: diagnosticContext.lastCompletedRow,
      lastAddress:
        String(data[data.length - 1]?.[addressColumn] || "") || job.lastAddress,
      failedRowsCount: currentFailedRowsCount,
      rowSummaries: statusRowSummaries,
      resultBuffer,
      resultFileName: buildDownloadFileName(originalName),
    });
  } catch (error) {
    updateJob(jobId, {
      status: "failed",
      elapsedMs: diagnosticContext.startedAt
        ? Date.now() - diagnosticContext.startedAt
        : 0,
      error: error.message || "파일 처리 중 오류가 발생했습니다.",
      rowSummaries: statusRowSummaries,
    });
    logApiError("job", error, {
      jobId,
      fileName: originalName,
      mode,
    });
  }
}

app.post(
  "/api/process-excel-jobs",
  upload.single("file"),
  asyncHandler(async (req, res) => {
    if (!req.file) {
      return sendApiError(
        res,
        400,
        "파일이 업로드되지 않았습니다",
        "FILE_MISSING"
      );
    }

    clearPreviousJobs();

    const job = saveJob(
      createJobState({
        fileName: req.file.originalname,
        mode: "auto",
      })
    );

    setImmediate(() => {
      processExcelJob(job.jobId, req.file.buffer, req.file.originalname, "auto");
    });

    res.status(202).json({
      jobId: job.jobId,
      status: job.status,
      phase: job.phase,
    });
  })
);

// multer 에러 핸들러
app.use((error, req, res, next) => {
  if (error instanceof multer.MulterError) {
    logApiError("upload", error, {
      path: req.originalUrl,
      fileName: req.file?.originalname || req.body?.fileName || null,
    });

    if (error.code === "LIMIT_FILE_SIZE") {
      return sendApiError(
        res,
        413,
        "파일이 너무 큽니다. 10MB 이하의 파일을 사용해주세요.",
        "FILE_TOO_LARGE"
      );
    }
    return sendApiError(
      res,
      400,
      `파일 업로드 오류: ${error.message}`,
      error.code || "UPLOAD_ERROR"
    );
  }
  next(error);
});

app.post(
  "/api/search",
  asyncHandler(async (req, res) => {
    const { keyword } = req.body;
    if (!keyword) {
      return sendApiError(res, 400, "검색어가 필요합니다", "KEYWORD_MISSING");
    }
    const mode = req.body.mode === "precise" ? "precise" : "fast";
    const result = await searchAddress(keyword, { mode });
    res.json(result);
  })
);

app.get("/api/process-status/latest", (req, res) => {
  const latestJob = getLatestJob();
  if (!latestJob) {
    return sendApiError(
      res,
      404,
      "최근 처리 상태 정보가 없습니다.",
      "PROCESS_STATUS_NOT_FOUND"
    );
  }

  res.json(serializeJob(latestJob));
});

app.get("/api/process-excel-jobs/:jobId", (req, res) => {
  const job = getJob(req.params.jobId);
  if (!job) {
    return sendApiError(res, 404, "작업을 찾을 수 없습니다.", "JOB_NOT_FOUND");
  }

  res.json(serializeJob(job));
});

app.get("/api/process-excel-jobs/:jobId/download", (req, res) => {
  const job = getJob(req.params.jobId);
  if (!job) {
    return sendApiError(res, 404, "작업을 찾을 수 없습니다.", "JOB_NOT_FOUND");
  }

  if (job.status !== "completed" || !job.resultBuffer) {
    return sendApiError(
      res,
      409,
      "결과 파일이 아직 준비되지 않았습니다.",
      "JOB_NOT_COMPLETED"
    );
  }

  res.setHeader("Content-Type", EXCEL_MIME_TYPE);
  try {
    const originalName = job.resultFileName || "result.xlsx";
    const safeAsciiName = sanitizeDownloadFileName(originalName);
    const utf8Name = encodeRFC5987ValueChars(originalName);
    res.setHeader(
      "Content-Disposition",
      `attachment; filename="${safeAsciiName}"; filename*=UTF-8''${utf8Name}`
    );
  } catch (error) {
    res.setHeader(
      "Content-Disposition",
      'attachment; filename="result.xlsx"'
    );
  }
  res.send(job.resultBuffer);
});

app.use("/api", (req, res) => {
  sendApiError(
    res,
    404,
    `요청한 API를 찾을 수 없습니다: ${req.method} ${req.originalUrl}`,
    "API_NOT_FOUND"
  );
});

app.use((error, req, res, next) => {
  const status = error.status || error.statusCode || 500;
  const code = error.code || "INTERNAL_SERVER_ERROR";
  const isApiRequest = req.originalUrl.startsWith("/api");

  if (req.originalUrl === "/api/process-excel-jobs" && latestJobId) {
    updateJob(latestJobId, {
      status: "failed",
      phase: "failed",
      elapsedMs: getLatestJob()?.startedAt
        ? Date.now() - getLatestJob().startedAt
        : getLatestJob()?.elapsedMs || 0,
      error: error.message || "파일 처리 중 오류가 발생했습니다.",
    });
  }

  logApiError("api", error, {
    path: req.originalUrl,
    method: req.method,
    status,
    fileName: req.file?.originalname || null,
  });

  if (isApiRequest) {
    return sendApiError(
      res,
      status,
      status >= 500
        ? "파일 처리 중 오류가 발생했습니다. 잠시 후 다시 시도해주세요."
        : error.message,
      code
    );
  }

  next(error);
});

app.listen(PORT, () => {
  console.log(`서버가 http://localhost:${PORT} 에서 실행 중입니다`);
});
