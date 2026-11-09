const test = require("node:test");
const assert = require("node:assert/strict");
const { EventEmitter } = require("node:events");
const { probeConnection, runConnectionDiagnostics } = require("../connection-diagnostic");

test("records connection phases without returning request secrets", async () => {
  const request = (url, options, callback) => {
    const req = new EventEmitter();
    req.end = () => {
      const socket = new EventEmitter();
      req.emit("socket", socket);
      for (const event of ["lookup", "connect", "secureConnect"]) socket.emit(event);
      const res = new EventEmitter();
      res.statusCode = 200;
      res.resume = () => res.emit("end");
      callback(res);
    };
    return req;
  };
  const result = await probeConnection({ apiUrl: "https://example.com", apiKey: "secret-key", request });
  assert.equal(result.httpStatus, 200);
  assert.equal(result.errorCode, null);
  assert.equal(Object.keys(result.timings).length, 5);
  assert.equal(JSON.stringify(result).includes("secret-key"), false);
});

test("overall deadline destroys stalled requests", async () => {
  let destroyed = false;
  const request = () => Object.assign(new EventEmitter(), {
    end() {}, destroy() { destroyed = true; },
  });
  const result = await probeConnection({ apiUrl: "https://example.com", apiKey: "key", request, timeoutMs: 5 });
  assert.equal(result.errorCode, "DIAGNOSTIC_TIMEOUT");
  assert.equal(destroyed, true);
});

test("connection failures resolve safely", async () => {
  const request = () => {
    const req = new EventEmitter();
    req.end = () => req.emit("error", Object.assign(new Error("secret-key"), { code: "ENOTFOUND" }));
    return req;
  };
  const result = await probeConnection({ apiUrl: "https://example.com", apiKey: "key", request });
  assert.equal(result.errorCode, "ENOTFOUND");
  assert.equal(JSON.stringify(result).includes("secret-key"), false);
});

test("runs exactly three sequential probes and skips missing keys", async () => {
  let calls = 0;
  const logs = [];
  const options = { apiUrl: "https://example.com", apiKey: "key", log: line => logs.push(line),
    probe: async () => { calls++; return { errorCode: null }; } };
  await runConnectionDiagnostics(options);
  assert.equal(calls, 3);
  assert.equal(logs.length, 3);
  await runConnectionDiagnostics({ ...options, apiKey: undefined });
  assert.equal(calls, 3);
});
