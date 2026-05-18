import { detectFieldMap, normalizeRows, parseRoomsFromWorkbook, parseWorkbookRows } from "./scheduler.js";

export async function parseStudentImportSource(source, expectedPool) {
  if (isTextSource(source)) {
    return parseStudentTextSource(await readTextSource(source), expectedPool);
  }
  return parseWorkbookRows(await readBinarySource(source), expectedPool);
}

export async function parseRoomImportSource(source) {
  if (isTextSource(source)) {
    return parseRoomTextSource(await readTextSource(source));
  }
  return parseRoomsFromWorkbook(await readBinarySource(source));
}

function isTextSource(source) {
  if (typeof source === "string") return true;
  if (typeof File !== "undefined" && source instanceof File) {
    return /\.(csv|tsv|txt)$/i.test(source.name || "") || String(source.type || "").startsWith("text/");
  }
  return false;
}

async function readTextSource(source) {
  if (typeof source === "string") return source;
  if (typeof source?.text === "function") return source.text();
  if (typeof source?.arrayBuffer === "function") {
    return new TextDecoder("utf-8").decode(await source.arrayBuffer());
  }
  return String(source || "");
}

async function readBinarySource(source) {
  if (source instanceof ArrayBuffer) return source;
  if (ArrayBuffer.isView(source)) return source.buffer.slice(source.byteOffset, source.byteOffset + source.byteLength);
  if (typeof source?.arrayBuffer === "function") return source.arrayBuffer();
  throw new Error("无法读取导入文件内容");
}

function parseStudentTextSource(text, expectedPool) {
  const matrix = parseTextMatrix(text);
  const extracted = extractTextTable(matrix, 16, (headers) => detectFieldMap(headers), ["班级", "姓名", "考号", "首选科目", "选科组合", "总分市排名", "总分", "数学"]);
  const fieldMap = detectFieldMap(extracted.headers);
  return {
    ...normalizeRows(extracted.rows, fieldMap, expectedPool),
    importMeta: { sheetName: "文本导入", headerRowNumber: extracted.headerRowNumber, sourceType: "text" },
  };
}

function parseRoomTextSource(text) {
  const matrix = parseTextMatrix(text);
  const extracted = extractTextTable(matrix, 12, detectRoomHeaderMap, ["考场号", "门牌号", "教室"]);
  const headerMap = detectRoomHeaderMap(extracted.headers);
  const rooms = extracted.rows
    .map((row) => ({
      roomNo: readRoomField(row, headerMap, "考场号"),
      doorNo: readRoomField(row, headerMap, "门牌号"),
      roomName: readRoomField(row, headerMap, "教室"),
      capacity: Math.max(1, Math.trunc(readNumberField(row, headerMap, ["人数", "容量"], 40) || 40)),
      enabled: String(readRoomField(row, headerMap, "考场号") || "").trim() !== "",
    }))
    .filter((room) => room.roomNo && room.enabled)
    .sort((a, b) => String(a.roomNo).localeCompare(String(b.roomNo), "zh-Hans-CN", { numeric: true }));
  return rooms;
}

function detectRoomHeaderMap(headers) {
  const aliases = {
    考场号: ["考场号", "考场", "场号", "房号"],
    门牌号: ["门牌号", "门牌", "门号", "标牌号"],
    教室: ["教室", "室号", "教室名", "房间"],
    人数: ["人数", "容量", "可容纳", "座位数"],
  };
  const normalized = headers.map(normalizeHeader);
  const map = {};
  for (const [field, list] of Object.entries(aliases)) {
    map[field] = findHeader(headers, list, normalized);
  }
  return map;
}

function readRoomField(row, headerMap, field) {
  const header = headerMap[field];
  return header ? row[header] ?? "" : "";
}

function readNumberField(row, headerMap, fields, fallback = "") {
  for (const field of fields) {
    const header = headerMap[field];
    const value = header ? row[header] : "";
    const parsed = Number.parseFloat(String(value ?? "").replace(/,/g, ""));
    if (Number.isFinite(parsed)) return parsed;
  }
  return fallback;
}

function extractTextTable(matrix, scanLimit, fieldMapper, requiredFields) {
  if (!matrix.length) return { headers: [], rows: [], headerRowNumber: 1 };
  const limit = Math.min(matrix.length, scanLimit);
  let best = { index: 0, score: -1, headers: uniquifyHeaders(matrix[0] || []) };
  for (let index = 0; index < limit; index += 1) {
    const headers = uniquifyHeaders(matrix[index] || []);
    const nonEmpty = headers.filter((header) => !header.startsWith("__EMPTY_")).length;
    if (nonEmpty < 2) continue;
    const map = fieldMapper(headers);
    const requiredScore = requiredFields.filter((field) => map[field]).length * 3;
    const score = requiredScore + Math.min(nonEmpty, 12) / 12;
    if (score > best.score) best = { index, score, headers };
  }
  const rows = matrix.slice(best.index + 1).map((values, rowIndex) => {
    const row = {};
    best.headers.forEach((header, columnIndex) => {
      row[header] = values[columnIndex] ?? "";
    });
    row.__rowNumber = best.index + rowIndex + 2;
    return row;
  });
  return { headers: best.headers, rows, headerRowNumber: best.index + 1 };
}

function parseTextMatrix(text) {
  const normalized = String(text ?? "").replace(/\r\n/g, "\n").replace(/\r/g, "\n").trim();
  if (!normalized) return [];
  const lines = normalized.split("\n").filter((line) => String(line).trim() !== "");
  if (!lines.length) return [];
  const delimiter = detectDelimiter(lines.slice(0, 8));
  return lines.map((line) => parseDelimitedLine(line, delimiter));
}

function detectDelimiter(lines) {
  const joined = lines.join("\n");
  const candidates = ["\t", ",", "；", ";", "|"];
  let best = "\t";
  let max = -1;
  for (const delimiter of candidates) {
    const count = countOccurrences(joined, delimiter);
    if (count > max) {
      max = count;
      best = delimiter;
    }
  }
  return best;
}

function parseDelimitedLine(line, delimiter) {
  const cells = [];
  let current = "";
  let quoted = false;
  for (let index = 0; index < line.length; index += 1) {
    const char = line[index];
    if (quoted) {
      if (char === '"' && line[index + 1] === '"') {
        current += '"';
        index += 1;
      } else if (char === '"') {
        quoted = false;
      } else {
        current += char;
      }
    } else if (char === '"') {
      quoted = true;
    } else if (char === delimiter) {
      cells.push(current.trim());
      current = "";
    } else {
      current += char;
    }
  }
  cells.push(current.trim());
  return cells;
}

function countOccurrences(text, token) {
  if (!token) return 0;
  return text.split(token).length - 1;
}

function normalizeHeader(value) {
  return String(value ?? "").trim().replace(/\s+/g, "");
}

function findHeader(headers, aliases, normalized = headers.map(normalizeHeader)) {
  for (const alias of aliases) {
    const index = normalized.indexOf(normalizeHeader(alias));
    if (index >= 0) return headers[index];
  }
  return "";
}

function uniquifyHeaders(values) {
  const seen = new Map();
  return values.map((value, index) => {
    const base = String(value ?? "").trim() || `__EMPTY_${index + 1}`;
    const count = seen.get(base) || 0;
    seen.set(base, count + 1);
    return count ? `${base}_${count + 1}` : base;
  });
}
