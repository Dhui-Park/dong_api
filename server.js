require("dotenv").config();
const express = require("express");
const multer = require("multer");
const xlsx = require("xlsx");
const axios = require("axios");
const path = require("path");
const iconv = require("iconv-lite");

const app = express();
const PORT = process.env.PORT || 3000;
const EXCEL_MIME_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

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

/**
 * backoff 전에 토큰화를 더 촘촘히:
 * - 공백 기준 단어로 나눈 뒤
 * - 단어가 숫자로 시작하고 뒤에 문자가 붙으면 "숫자" + "문자"로 분리
 */
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

// 단일 API 호출
async function searchAddressOnce(keyword) {
  try {
    const response = await axios.get(JUSO_API_URL, {
      params: {
        confmKey: JUSO_API_KEY,
        currentPage: 1,
        countPerPage: 1,
        keyword: keyword,
        resultType: "json",
        addInfoYn: "Y",
      },
    });

    const result = response.data?.results;
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
      };
    }
    return {
      success: false,
      error: result?.common?.errorMessage || "주소를 찾을 수 없습니다",
    };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

// 검색 실패 시 마지막 토큰부터 제거하며 재시도
async function searchAddress(keyword) {
  // 1. 전처리 (외N필지 패턴 제거)
  let processed = preprocessAddress(keyword);

  // 2. 첫 번째 시도
  let result = await searchAddressOnce(processed);
  if (result.success) {
    return result;
  }

  // 3. 실패 시: 토큰화 후 마지막 토큰부터 하나씩 제거하며 재시도
  let tokens = tokenizeForBackoff(processed);

  while (tokens.length > 1) {
    tokens.pop();
    const shortened = tokens.join(" ").trim();
    result = await searchAddressOnce(shortened);
    if (result.success) {
      return result;
    }
  }

  return { success: false, error: "주소를 찾을 수 없습니다" };
}

app.post(
  "/api/process-excel",
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

    let fileBuffer = req.file.buffer;
    const fileName = req.file.originalname.toLowerCase();

    // CSV 파일인 경우 EUC-KR에서 UTF-8로 변환 시도
    if (fileName.endsWith(".csv")) {
      const content = iconv.decode(fileBuffer, "cp949");
      fileBuffer = Buffer.from(content, "utf-8");
    }

    const workbook = xlsx.read(fileBuffer, { type: "buffer" });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(worksheet);

    if (data.length === 0) {
      return sendApiError(
        res,
        400,
        "파일에 데이터가 없습니다",
        "EMPTY_FILE"
      );
    }

    const addressColumn = Object.keys(data[0]).find(
      (key) =>
        key.toLowerCase().includes("주소") ||
        key.toLowerCase().includes("address")
    );

    if (!addressColumn) {
      return sendApiError(
        res,
        400,
        "'주소' 컬럼을 찾을 수 없습니다",
        "ADDRESS_COLUMN_MISSING"
      );
    }

    // '열람 동사무소' 컬럼 존재 여부 확인
    const dongOfficeColumn = Object.keys(data[0]).find((key) =>
      key.includes("열람 동사무소")
    );

    // '열람 동사무소' 값에 '동' 붙이기 함수
    function addDongSuffix(value) {
      if (!value) return "";
      const trimmed = String(value).trim();
      if (trimmed === "") return "";
      // '동', '읍', '면'으로 끝나면 그대로 반환
      if (/[동읍면]$/.test(trimmed)) {
        return trimmed;
      }
      return trimmed + "동";
    }

    const results = [];
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const address = row[addressColumn];

      const resultRow = { ...row };

      if (address) {
        const searchResult = await searchAddress(address);

        resultRow.관할동 = searchResult.success ? searchResult.emdNm : "";
        resultRow.시군구 = searchResult.success ? searchResult.sggNm : "";
        resultRow.시도 = searchResult.success ? searchResult.siNm : "";
        resultRow.도로명주소 = searchResult.success
          ? searchResult.roadAddr
          : "";

        // 관할동사무소: 마지막 단어만 저장 (예: "강서구 화곡1동" → "화곡1동")
        const dongOfficeRaw = searchResult.success
          ? (searchResult.dong || "").split(/\s+/).pop() || ""
          : "";

        // 관할동이 "제동"으로 끝나는 경우 (예: "화곡제동")
        // → "제제{숫자}동" 패턴일 때만 '제' 한 글자 제거 (제제1동 → 제1동)
        if (/제동$/.test(resultRow.관할동)) {
          resultRow.관할동사무소 = dongOfficeRaw.replace(
            /제제([0-9])동/g,
            "제$1동"
          );
        } else {
          // 기존 로직: "제{숫자}동" → "{숫자}동"
          resultRow.관할동사무소 = removeJePrefixIfNeeded(dongOfficeRaw);
        }

        resultRow.검색결과 = searchResult.success ? "성공" : searchResult.error;
      } else {
        resultRow.관할동 = "";
        resultRow.시군구 = "";
        resultRow.시도 = "";
        resultRow.도로명주소 = "";
        resultRow.관할동사무소 = "";
        resultRow.검색결과 = "주소 없음";
      }

      // '열람 동사무소' 컬럼이 있으면 '_api용' 컬럼 추가 (+ "제" 처리)
      if (dongOfficeColumn) {
        const officeApiValue = addDongSuffix(row[dongOfficeColumn]);
        const gwanhalDong = resultRow.관할동 || "";

        // 관할동이 "제동"으로 끝나는 경우 (예: "화곡제동")
        // → "제제{숫자}동" 패턴일 때만 '제' 한 글자 제거 (제제1동 → 제1동)
        if (/제동$/.test(gwanhalDong)) {
          resultRow["열람 동사무소_api용"] = officeApiValue.replace(
            /제제([0-9])동/g,
            "제$1동"
          );
        } else {
          // 기존 로직: "제{숫자}동" → "{숫자}동"
          resultRow["열람 동사무소_api용"] =
            removeJePrefixIfNeeded(officeApiValue);
        }
      }

      results.push(resultRow);
    }

    const newWorkbook = xlsx.utils.book_new();
    const newWorksheet = xlsx.utils.json_to_sheet(results);
    xlsx.utils.book_append_sheet(newWorkbook, newWorksheet, "Result");

    const buffer = xlsx.write(newWorkbook, {
      type: "buffer",
      bookType: "xlsx",
    });

    res.setHeader("Content-Type", EXCEL_MIME_TYPE);
    res.setHeader("Content-Disposition", "attachment; filename=result.xlsx");
    res.send(buffer);
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

app.post("/api/search", asyncHandler(async (req, res) => {
  const { keyword } = req.body;
  if (!keyword) {
    return sendApiError(res, 400, "검색어가 필요합니다", "KEYWORD_MISSING");
  }
  const result = await searchAddress(keyword);
  res.json(result);
}));

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
