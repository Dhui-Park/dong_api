const https = require("node:https");
const { performance } = require("node:perf_hooks");

function probeConnection({ apiUrl, apiKey, request = https.request, timeoutMs = 15000 }) {
  return new Promise((resolve) => {
    const start = performance.now();
    const timings = {};
    let req;
    let timer;
    let settled = false;
    let httpStatus = null;
    const mark = (name) => { timings[name] = performance.now() - start; };
    const finish = (errorCode = null) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      resolve({ timings, elapsedMs: performance.now() - start, httpStatus, errorCode });
    };
    try {
      const url = new URL(apiUrl);
      url.search = new URLSearchParams({ confmKey: apiKey, currentPage: "1",
        countPerPage: "1", keyword: "서울특별시 중구 세종대로 110", resultType: "json", addInfoYn: "Y" }).toString();
      req = request(url, { method: "GET", agent: false }, (res) => {
        httpStatus = res.statusCode;
        mark("firstResponseMs");
        res.on("error", (error) => finish(safeCode(error)));
        res.on("aborted", () => finish("RESPONSE_ABORTED"));
        res.on("end", () => { mark("responseCompleteMs"); finish(); });
        res.resume();
      });
      req.on("socket", (socket) => {
        socket.once("lookup", () => mark("dnsMs"));
        socket.once("connect", () => mark("tcpMs"));
        socket.once("secureConnect", () => mark("tlsMs"));
      });
      req.on("error", (error) => finish(safeCode(error)));
      timer = setTimeout(() => {
        finish("DIAGNOSTIC_TIMEOUT");
        req.destroy();
      }, timeoutMs);
      req.end();
    } catch (error) {
      finish(safeCode(error));
      if (req) req.destroy();
    }
  });
}

function safeCode(error) {
  const code = String(error?.code || "DIAGNOSTIC_ERROR");
  return /^[A-Z][A-Z0-9_]{0,63}$/.test(code) ? code : "DIAGNOSTIC_ERROR";
}

async function runConnectionDiagnostics({ apiUrl, apiKey, log = console.log, probe = probeConnection }) {
  if (!apiKey) {
    log(JSON.stringify({ event: "connection.diagnostic", errorCode: "KEY_MISSING" }));
    return;
  }
  for (let attempt = 1; attempt <= 3; attempt++) {
    const result = await probe({ apiUrl, apiKey });
    log(JSON.stringify({ event: "connection.diagnostic", time: new Date().toISOString(),
      transport: "node:https-fresh-connection", attempt, ...result }));
  }
}

module.exports = { probeConnection, runConnectionDiagnostics };
