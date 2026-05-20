import JSZip from "jszip";
import XLSX from "xlsx-js-style";

export const DEFAULT_COMBO_ORDER = [
  "化学地理",
  "化学生物",
  "化学政治",
  "政治地理",
  "生物地理",
  "生物政治",
];

export const LANGUAGE_SUBJECTS = ["日语", "俄语", "西班牙语", "法语", "德语"];
export const ELECTIVE_SUBJECTS = ["化学", "地理", "政治", "生物"];
export const SCHEDULE_MODES = {
  THREE_DAY_SPLIT: "threeDaySplit",
  TWO_DAY_COMBO: "twoDayCombo",
};

export const VALIDATION_LEVELS = {
  BLOCKER: "阻断错误",
  WARNING: "风险提醒",
  REVIEW: "人工复核",
};

export const EXPORT_GROUPS = {
  ALL: "all",
  ADMIN: "admin",
  CLASS: "class",
  ROOM: "room",
  SUBJECT: "subject",
  ROOM_SUMMARY: "roomSummary",
  DOOR_SUMMARY: "doorSummary",
  TIME: "timeSheet",
  VALIDATION: "validation",
};

const REQUIRED_FIELDS = ["班级", "姓名", "考号", "首选科目", "选科组合", "总分市排名", "总分", "数学"];

const FIELD_ALIASES = {
  班级: ["班级", "行政班", "班别", "班名", "教学班", "原班级", "所在班级", "学生班级"],
  姓名: ["姓名", "学生姓名", "考生姓名", "姓名.", "名字"],
  考号: ["考号", "准考证号", "学号", "学生号", "考籍号", "报名号", "考试号", "座号"],
  首选科目: ["首选科目", "首选", "科类", "类别", "选考类别", "学生类别", "物历方向"],
  选科组合: ["选科组合", "组合", "选科", "科目组合", "选考科目", "再选科目", "再选组合"],
  总分市排名: ["总分市排名", "总分排名", "市排名", "排名", "总名次", "总分名次", "总排名", "市名次"],
  总分: ["总分", "总成绩", "总分成绩", "赋分总分"],
  数学: ["数学", "数学成绩", "数学分数"],
  外语: ["外语", "英语", "外语成绩", "英语成绩", "外语分数", "英语分数"],
  化学: ["化学", "化学成绩", "化学分数"],
  地理: ["地理", "地理成绩", "地理分数"],
  政治: ["政治", "思想政治", "政治成绩", "思想政治成绩", "政治分数"],
  生物: ["生物", "生物成绩", "生物学", "生物学成绩", "生物分数"],
};

const COMBO_CANONICAL = new Map([
  ["化学|地理", "化学地理"],
  ["化学|生物", "化学生物"],
  ["化学|政治", "化学政治"],
  ["地理|政治", "政治地理"],
  ["地理|生物", "生物地理"],
  ["政治|生物", "生物政治"],
]);

const COMBO_SHORT_LABELS = {
  化学地理: "化地",
  化学生物: "化生",
  化学政治: "化政",
  政治地理: "政地",
  生物地理: "生地",
  生物政治: "生政",
};

const safeString = (value) => (value == null ? "" : String(value).trim());

const toNumber = (value, fallback = 0) => {
  if (typeof value === "number") return Number.isFinite(value) ? value : fallback;
  const parsed = Number.parseFloat(safeString(value).replace(/,/g, ""));
  return Number.isFinite(parsed) ? parsed : fallback;
};

const normalizeHeader = (value) => safeString(value).replace(/\s+/g, "");

const findHeader = (headers, aliases) => {
  const normalized = headers.map(normalizeHeader);
  for (const alias of aliases) {
    const index = normalized.indexOf(normalizeHeader(alias));
    if (index >= 0) return headers[index];
  }
  return "";
};

export function detectFieldMap(headers) {
  const map = {};
  for (const [field, aliases] of Object.entries(FIELD_ALIASES)) {
    map[field] = findHeader(headers, aliases);
  }
  for (const language of LANGUAGE_SUBJECTS) {
    map[language] = findHeader(headers, [language, `${language}成绩`, `${language}分数`]);
  }
  for (const subject of ELECTIVE_SUBJECTS) {
    map[subject] = map[subject] || findHeader(headers, [subject, `${subject}成绩`, `${subject}分数`]);
  }
  return map;
}

export function parseWorkbookRows(arrayBuffer, expectedPool) {
  const workbook = XLSX.read(arrayBuffer, { type: "array", cellDates: false });
  const candidates = workbook.SheetNames.map((sheetName) => {
    const rawRows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: "", raw: true, blankrows: false });
    return { sheetName, ...extractStudentRows(rawRows) };
  });
  const selected = candidates.sort((a, b) => b.score - a.score)[0] || { sheetName: workbook.SheetNames[0] || "", rows: [], headers: [], headerRowNumber: 1 };
  const { sheetName, rows, headers, headerRowNumber } = selected;
  const fieldMap = detectFieldMap(headers);
  return {
    ...normalizeRows(rows, fieldMap, expectedPool),
    importMeta: { sheetName, headerRowNumber },
  };
}

function extractStudentRows(rawRows) {
  if (!rawRows.length) return { headers: [], rows: [], headerRowNumber: 1, score: -1 };
  const scanLimit = Math.min(rawRows.length, 16);
  let best = { index: 0, score: -1, headers: uniquifyHeaders(rawRows[0] || []) };
  for (let index = 0; index < scanLimit; index += 1) {
    const headers = uniquifyHeaders(rawRows[index] || []);
    const nonEmpty = headers.filter((header) => !header.startsWith("__EMPTY_")).length;
    if (nonEmpty < 2) continue;
    const map = detectFieldMap(headers);
    const requiredScore = REQUIRED_FIELDS.filter((field) => map[field]).length * 3;
    const subjectScore = [...ELECTIVE_SUBJECTS, ...LANGUAGE_SUBJECTS, "外语"].filter((field) => map[field]).length;
    const score = requiredScore + subjectScore + Math.min(nonEmpty, 12) / 12;
    if (score > best.score) best = { index, score, headers };
  }
  const rows = rawRows.slice(best.index + 1).map((values, rowIndex) => {
    const row = {};
    best.headers.forEach((header, columnIndex) => {
      row[header] = values[columnIndex] ?? "";
    });
    row.__rowNumber = best.index + rowIndex + 2;
    return row;
  });
  return { headers: best.headers, rows, headerRowNumber: best.index + 1, score: best.score };
}

function uniquifyHeaders(values) {
  const seen = new Map();
  return values.map((value, index) => {
    const base = safeString(value) || `__EMPTY_${index + 1}`;
    const count = seen.get(base) || 0;
    seen.set(base, count + 1);
    return count ? `${base}_${count + 1}` : base;
  });
}

export function normalizeRows(rows, fieldMap, expectedPool) {
  const students = rows
    .filter((row) => Object.entries(row).some(([key, value]) => !key.startsWith("__") && safeString(value)))
    .map((row, index) => {
      const get = (field) => row[fieldMap[field]] ?? "";
      const language = detectLanguage(row, fieldMap);
      const firstSubject = safeString(get("首选科目")) || expectedPool;
      const subjectScores = Object.fromEntries(ELECTIVE_SUBJECTS.map((subject) => [subject, toNumber(row[fieldMap[subject]], Number.NaN)]));
      return {
        id: safeString(get("考号")),
        name: safeString(get("姓名")),
        className: safeString(get("班级")),
        firstSubject,
        pool: expectedPool || firstSubject,
        comboRaw: safeString(get("选科组合")),
        totalRank: toNumber(get("总分市排名"), Number.POSITIVE_INFINITY),
        totalScore: toNumber(get("总分")),
        mathScore: toNumber(get("数学")),
        foreignScore: toNumber(get("外语")),
        language,
        languageScore: language === "英语" ? toNumber(get("外语")) : toNumber(row[fieldMap[language]]),
        subjectScores,
        original: row,
        rowNumber: row.__rowNumber || index + 2,
      };
    });

  return {
    fieldMap,
    students,
    errors: validateImportedStudents(students, fieldMap, expectedPool),
  };
}

function detectLanguage(row, fieldMap) {
  for (const language of LANGUAGE_SUBJECTS) {
    const header = fieldMap[language];
    if (header && safeString(row[header]) !== "") return language;
  }
  return "英语";
}

function validateImportedStudents(students, fieldMap, expectedPool) {
  const errors = [];
  for (const field of REQUIRED_FIELDS) {
    if (!fieldMap[field]) errors.push(`缺少必填字段：${field}`);
  }
  const seen = new Map();
  for (const student of students) {
    if (!student.id) errors.push(`第 ${student.rowNumber} 行缺少考号`);
    if (!student.name) errors.push(`第 ${student.rowNumber} 行缺少姓名`);
    if (!student.className) errors.push(`第 ${student.rowNumber} 行缺少班级`);
    if (!Number.isFinite(student.totalRank)) errors.push(`第 ${student.rowNumber} 行缺少总分市排名`);
    if (expectedPool && student.firstSubject && normalizePoolLabel(student.firstSubject) !== expectedPool) {
      errors.push(`第 ${student.rowNumber} 行首选科目为 ${student.firstSubject}，不是 ${expectedPool}`);
    }
    if (student.id) {
      if (seen.has(student.id)) errors.push(`考号重复：${student.id}（第 ${seen.get(student.id)} 行、第 ${student.rowNumber} 行）`);
      seen.set(student.id, student.rowNumber);
    }
  }
  return [...new Set(errors)];
}

function normalizePoolLabel(value) {
  const text = safeString(value);
  if (text.includes("物理")) return "物理";
  if (text.includes("历史")) return "历史";
  return text;
}

export function parseRoomsFromWorkbook(arrayBuffer) {
  const workbook = XLSX.read(arrayBuffer, { type: "array", cellDates: false });
  const sheetName = workbook.SheetNames.find((name) => name.includes("考场门牌人数")) || workbook.SheetNames[0];
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: "", raw: true });
  return rows
    .map((row) => ({
      roomNo: safeString(row["考场号"]),
      doorNo: safeString(row["门牌号"]),
      roomName: safeString(row["教室"]),
      capacity: Math.max(1, Math.trunc(toNumber(row["人数"], toNumber(row["容量"], 40)) || 40)),
      enabled: safeString(row["考场号"]) !== "",
    }))
    .filter((room) => room.roomNo && room.enabled && room.roomNo !== "501")
    .sort(compareRoomNo);
}

export function defaultRooms() {
  const floors = [
    [201, "高三01班"],
    [202, "高三02班"],
    [203, "高三03班"],
    [204, "高三04班"],
    [211, "高三05班"],
    [212, "高三06班"],
    [213, "高三07班"],
    [214, "高三08班"],
    [301, "高三09班"],
    [302, "高三10班"],
    [303, "高三11班"],
    [304, "高三12班"],
    [311, "高三13班"],
    [312, "高三14班"],
    [313, "高三15班"],
    [414, "414"],
    [401, "高三17班"],
    [402, "高三18班"],
    [403, "高三19班"],
    [404, "404"],
    [412, "412"],
    [413, "413"],
  ];
  return floors.map(([doorNo, roomName], index) => ({
    roomNo: String(index + 1),
    doorNo: String(doorNo),
    roomName,
    capacity: index === 21 ? 34 : 40,
    enabled: true,
  }));
}

function compareRoomNo(a, b) {
  const an = toNumber(a.roomNo, Number.NaN);
  const bn = toNumber(b.roomNo, Number.NaN);
  if (Number.isFinite(an) && Number.isFinite(bn)) return an - bn;
  return safeString(a.roomNo).localeCompare(safeString(b.roomNo), "zh-Hans-CN", { numeric: true });
}

export function buildSchedule({ physicsStudents, historyStudents, rooms, minorLanguageRooms, examTimes, mode = SCHEDULE_MODES.THREE_DAY_SPLIT }) {
  const enabledRooms = rooms.filter((room) => room.enabled).sort(compareRoomNo);
  const allStudents = [...physicsStudents, ...historyStudents];
  const errors = validateBeforeSchedule({ physicsStudents, historyStudents, enabledRooms, minorLanguageRooms, mode });
  if (errors.length) {
    return { errors, summary: [], mainAssignments: [], foreignAssignments: [], electiveAssignments: [], subjectAssignments: [], allRows: [], mode, examTimes };
  }

  const mainAssignments = assignMain(physicsStudents, historyStudents, enabledRooms);
  const foreignAssignments = assignForeign(mainAssignments, enabledRooms, minorLanguageRooms);
  const electiveAssignments = mode === SCHEDULE_MODES.TWO_DAY_COMBO ? assignElectives(allStudents, enabledRooms) : [];
  const subjectAssignments = mode === SCHEDULE_MODES.THREE_DAY_SPLIT ? assignSplitSubjects(allStudents, enabledRooms) : [];
  const allRows = mergeStudentRows(allStudents, mainAssignments, foreignAssignments, electiveAssignments, subjectAssignments, examTimes, mode);
  const summary = buildSummary(mainAssignments, foreignAssignments, electiveAssignments, subjectAssignments, enabledRooms, mode);
  const generatedErrors = validateGenerated({ mainAssignments, foreignAssignments, electiveAssignments, subjectAssignments, allStudents, mode, rooms: enabledRooms, examTimes });

  return {
    errors: generatedErrors,
    summary,
    mode,
    mainAssignments,
    foreignAssignments,
    electiveAssignments,
    subjectAssignments,
    allRows,
    examTimes,
  };
}

function validateBeforeSchedule({ physicsStudents, historyStudents, enabledRooms, minorLanguageRooms, mode }) {
  const errors = [];
  if (!physicsStudents.length) errors.push("缺少物理类学生成绩单");
  if (!historyStudents.length) errors.push("缺少历史类学生成绩单");
  if (!enabledRooms.length) errors.push("缺少普通考场清单");
  const ids = new Map();
  for (const student of [...physicsStudents, ...historyStudents]) {
    if (ids.has(student.id)) errors.push(`物理/历史成绩单存在重复考号：${student.id}`);
    ids.set(student.id, true);
    const subjects = getElectiveSubjects(student);
    if (subjects.length !== 2) errors.push(`${student.name || student.id} 的再选科不是两门：${student.comboRaw}`);
    if (mode === SCHEDULE_MODES.THREE_DAY_SPLIT) {
      for (const subject of subjects) {
        if (!Number.isFinite(student.subjectScores?.[subject])) {
          errors.push(`${student.name || student.id} 选了${subject}，但${subject}成绩为空`);
        }
      }
    }
  }
  if (!enabledRooms.length) {
    return [...new Set(errors)];
  }
  const foreignPlanErrors = validateForeignRoomPlan([...physicsStudents, ...historyStudents], enabledRooms, minorLanguageRooms);
  errors.push(...foreignPlanErrors);
  if (!foreignPlanErrors.length) {
    const foreignNeeded = countForeignNeededRooms([...physicsStudents, ...historyStudents], enabledRooms, minorLanguageRooms);
    if (!Number.isFinite(foreignNeeded)) errors.push(`外语普通考场容量不足：请新增考场或调整外语考场容量`);
    if (foreignNeeded > enabledRooms.length) errors.push(`外语普通考场容量不足：需要 ${foreignNeeded} 个考场，当前 ${enabledRooms.length} 个`);
  }
  const physicsRooms = countNeededRooms(physicsStudents, enabledRooms, 0);
  const totalMainRooms = physicsRooms + countNeededRooms(historyStudents, enabledRooms, physicsRooms);
  if (totalMainRooms > enabledRooms.length) errors.push(`普通考场容量不足：需要 ${totalMainRooms} 个考场，当前 ${enabledRooms.length} 个`);
  if (mode === SCHEDULE_MODES.TWO_DAY_COMBO) {
    const electiveNeeded = countNeededRooms([...physicsStudents, ...historyStudents], enabledRooms, 0);
    if (electiveNeeded > enabledRooms.length) errors.push(`四选二普通考场容量不足：需要 ${electiveNeeded} 个考场，当前 ${enabledRooms.length} 个`);
  } else {
    for (const subject of ELECTIVE_SUBJECTS) {
      const allStudents = [...physicsStudents, ...historyStudents];
      const examStudents = allStudents.filter((student) => getElectiveSubjects(student).includes(subject));
      const selfStudyStudents = allStudents.filter((student) => !getElectiveSubjects(student).includes(subject));
      const examNeeded = countNeededRooms(examStudents, enabledRooms, 0);
      const selfNeeded = countNeededRooms(selfStudyStudents, enabledRooms, examNeeded);
      const totalNeeded = examNeeded + selfNeeded;
      if (totalNeeded > enabledRooms.length) {
        const availableSelfSeats = capacityFrom(enabledRooms, examNeeded);
        errors.push(`${subject}考试+自习普通考场容量不足：考试需 ${examNeeded} 个考场，自习还缺 ${Math.max(1, selfStudyStudents.length - availableSelfSeats)} 个座位`);
      }
    }
  }
  return [...new Set(errors)];
}

function countNeededRooms(students, rooms, startIndex) {
  let count = 0;
  let remaining = students.length;
  for (let index = startIndex; index < rooms.length && remaining > 0; index += 1) {
    remaining -= rooms[index].capacity;
    count += 1;
  }
  return remaining > 0 ? Number.POSITIVE_INFINITY : count;
}

function capacityFrom(rooms, startIndex) {
  return rooms.slice(startIndex).reduce((sum, room) => sum + room.capacity, 0);
}

function countForeignNeededRooms(students, rooms, foreignRoomPlan = {}) {
  const languageGroups = groupForeignStudentsByLanguage(students);
  const manualIndexesByLanguage = buildManualForeignRoomIndexes(languageGroups, rooms, foreignRoomPlan);
  const reservedManualIndexes = new Set([...manualIndexesByLanguage.values()].flat());
  const occupied = new Set();
  let cursor = 0;
  let usedCount = 0;
  for (const { language, students: languageStudents } of languageGroups) {
    const preferredIndexes = manualIndexesByLanguage.get(language) || [];
    const indexes = preferredIndexes.length ? preferredIndexes : allocateRoomIndexes(languageStudents.length, rooms, cursor, new Set([...occupied, ...reservedManualIndexes]));
    if (!indexes.length || capacityOfIndexes(rooms, indexes) < languageStudents.length) return Number.POSITIVE_INFINITY;
    indexes.forEach((index) => occupied.add(index));
    usedCount = Math.max(usedCount, ...indexes.map((index) => index + 1));
    cursor = Math.max(cursor, Math.max(...indexes) + 1);
  }
  return usedCount;
}

function groupForeignStudentsByLanguage(students) {
  return ["英语", ...LANGUAGE_SUBJECTS]
    .map((language) => ({
      language,
      students: students.filter((student) => (student.language || "英语") === language),
    }))
    .filter((group) => group.students.length);
}

function getPlannedRoomIndexes(plan = {}, rooms = []) {
  return resolveForeignRoomPlan(plan, rooms).indexes;
}

function resolveForeignRoomPlan(plan = {}, rooms = []) {
  const tokens = safeString([plan.roomNos, plan.roomNo, plan.doorNo, plan.roomName].filter(Boolean).join(" "))
    .split(/[,，、/ ]+/)
    .map((value) => value.trim())
    .filter(Boolean);
  const indexes = [];
  const unmatched = [];
  for (const token of tokens) {
    const index = findRoomIndexByToken(token, rooms);
    if (index >= 0) {
      if (!indexes.includes(index)) indexes.push(index);
    } else {
      unmatched.push(token);
    }
  }
  return { indexes, tokens, unmatched };
}

function findRoomIndexByToken(token, rooms) {
  const raw = safeString(token);
  const normalizedToken = normalizeRoomToken(raw);
  return rooms.findIndex((room) => {
    const candidates = [
      room.roomNo,
      room.doorNo,
      room.roomName,
      `${room.roomNo}考场`,
      `第${room.roomNo}考场`,
      room.doorNo ? `${room.doorNo}教室` : "",
    ].map((value) => normalizeRoomToken(value));
    return candidates.some((candidate) => candidate && candidate === normalizedToken);
  });
}

function normalizeRoomToken(value) {
  return safeString(value)
    .replace(/^第/, "")
    .replace(/(考场|教室|门牌|号)$/g, "")
    .replace(/\s+/g, "")
    .toLowerCase();
}

function buildManualForeignRoomIndexes(languageGroups, rooms, foreignRoomPlan = {}) {
  const result = new Map();
  for (const { language } of languageGroups) {
    const resolved = resolveForeignRoomPlan(foreignRoomPlan[language], rooms);
    if (resolved.tokens.length) result.set(language, resolved.indexes);
  }
  return result;
}

function validateForeignRoomPlan(students, rooms, foreignRoomPlan = {}) {
  const errors = [];
  const languageGroups = groupForeignStudentsByLanguage(students);
  const manualRoomOwners = new Map();
  for (const { language, students: languageStudents } of languageGroups) {
    const resolved = resolveForeignRoomPlan(foreignRoomPlan[language], rooms);
    if (!resolved.tokens.length) continue;
    if (resolved.unmatched.length) {
      errors.push(`${language}指定外语考场未找到：${resolved.unmatched.join("、")}。请在“确认考场”里新增或改成已有考场号/门牌号`);
      continue;
    }
    if (!resolved.indexes.length) {
      errors.push(`${language}指定外语考场为空：请填写已有考场号，或留空让系统自动接在英语后面`);
      continue;
    }
    const capacity = capacityOfIndexes(rooms, resolved.indexes);
    if (capacity < languageStudents.length) {
      const roomText = resolved.indexes.map((index) => rooms[index]?.roomNo).filter(Boolean).join("、");
      errors.push(`${language}外语考场容量不足：指定 ${roomText} 考场容量 ${capacity}，学生 ${languageStudents.length} 人，还缺 ${languageStudents.length - capacity} 座位`);
    }
    for (const index of resolved.indexes) {
      const room = rooms[index];
      const owner = manualRoomOwners.get(index);
      if (owner && owner !== language) {
        errors.push(`外语语种考场冲突：${owner}和${language}都指定了${room?.roomNo || index + 1}考场`);
      } else {
        manualRoomOwners.set(index, language);
      }
    }
  }
  return [...new Set(errors)];
}

function allocateRoomIndexes(studentCount, rooms, startIndex, occupied = new Set()) {
  const indexes = [];
  let remaining = studentCount;
  for (let index = startIndex; index < rooms.length && remaining > 0; index += 1) {
    if (occupied.has(index)) continue;
    indexes.push(index);
    remaining -= Number(rooms[index].capacity) || 40;
  }
  return remaining > 0 ? [] : indexes;
}

function capacityOfIndexes(rooms, indexes) {
  return indexes.reduce((sum, index) => sum + (Number(rooms[index]?.capacity) || 40), 0);
}

function studentSort(a, b) {
  return (
    a.totalRank - b.totalRank ||
    b.mathScore - a.mathScore ||
    b.foreignScore - a.foreignScore ||
    a.id.localeCompare(b.id, "zh-Hans-CN", { numeric: true })
  );
}

function assignMain(physicsStudents, historyStudents, rooms) {
  const assignments = [];
  let roomIndex = 0;
  for (const [pool, students] of [
    ["物理", [...physicsStudents].sort(studentSort)],
    ["历史", [...historyStudents].sort(studentSort)],
  ]) {
    let seat = 1;
    for (const student of students) {
      const room = rooms[roomIndex];
      assignments.push(makeAssignment(student, room, seat, "主考", pool));
      seat += 1;
      if (seat > room.capacity) {
        roomIndex += 1;
        seat = 1;
      }
    }
    if (seat > 1) roomIndex += 1;
  }
  return assignments;
}

function makeAssignment(student, room, seatNo, plan, subjectLabel) {
  return {
    studentId: student.id,
    name: student.name,
    className: student.className,
    firstSubject: student.firstSubject,
    comboRaw: student.comboRaw,
    language: student.language,
    languageScore: student.languageScore,
    original: student.original,
    plan,
    subjectLabel,
    roomNo: room.roomNo,
    doorNo: room.doorNo,
    roomName: room.roomName,
    seatNo,
  };
}

function assignForeign(mainAssignments, rooms, minorLanguageRooms) {
  const assignments = [];
  const languageGroups = groupForeignStudentsByLanguage(mainAssignments);
  const manualIndexesByLanguage = buildManualForeignRoomIndexes(languageGroups, rooms, minorLanguageRooms);
  const reservedManualIndexes = new Set([...manualIndexesByLanguage.values()].flat());
  const occupied = new Set();
  let cursor = 0;

  for (const { language, students } of languageGroups) {
    const sorted = [...students].sort((a, b) => b.languageScore - a.languageScore || a.studentId.localeCompare(b.studentId, "zh-Hans-CN", { numeric: true }));
    const preferredIndexes = manualIndexesByLanguage.get(language) || [];
    const roomIndexes = preferredIndexes.length ? preferredIndexes : allocateRoomIndexes(sorted.length, rooms, cursor, new Set([...occupied, ...reservedManualIndexes]));
    let studentIndex = 0;
    for (const roomIndex of roomIndexes) {
      const room = rooms[roomIndex];
      const capacity = Number(room.capacity) || 40;
      for (let seat = 1; seat <= capacity && studentIndex < sorted.length; seat += 1) {
        assignments.push({
          ...sorted[studentIndex],
          plan: "外语",
          subjectLabel: language,
          roomNo: room.roomNo,
          doorNo: room.doorNo,
          roomName: room.roomName,
          seatNo: seat,
        });
        studentIndex += 1;
      }
      occupied.add(roomIndex);
    }
    if (roomIndexes.length) cursor = Math.max(cursor, Math.max(...roomIndexes) + 1);
  }

  return assignments;
}

export function getElectiveCombo(student) {
  const parts = safeString(student.comboRaw)
    .split(/[,，、/ ]+/)
    .map((part) => part.trim())
    .filter(Boolean);
  const combo = parts.filter((part) => part !== student.firstSubject);
  if (combo.length !== 2) return "";
  const key = [...combo].sort().join("|");
  return COMBO_CANONICAL.get(key) || combo.join("");
}

function assignElectives(students, rooms) {
  const grouped = new Map(DEFAULT_COMBO_ORDER.map((combo) => [combo, []]));
  for (const student of students) {
    const combo = getElectiveCombo(student);
    if (!grouped.has(combo)) grouped.set(combo, []);
    grouped.get(combo).push(student);
  }
  const ordered = [...grouped.entries()].flatMap(([combo, comboStudents]) =>
    comboStudents.sort(studentSort).map((student) => ({ student, combo })),
  );
  const assignments = [];
  let roomIndex = 0;
  let seat = 1;
  for (const item of ordered) {
    const room = rooms[roomIndex];
    assignments.push(makeAssignment(item.student, room, seat, "四选二", item.combo));
    seat += 1;
    if (seat > room.capacity) {
      roomIndex += 1;
      seat = 1;
    }
  }
  return assignments;
}

function assignSplitSubjects(students, rooms) {
  return ELECTIVE_SUBJECTS.flatMap((subject) => assignOneSubject(students, rooms, subject));
}

function assignOneSubject(students, rooms, subject) {
  const examStudents = students
    .filter((student) => getElectiveSubjects(student).includes(subject))
    .sort((a, b) => subjectStudentSort(a, b, subject));
  const selfStudyStudents = students
    .filter((student) => !getElectiveSubjects(student).includes(subject))
    .sort(selfStudySort);

  const assignments = [];
  let roomIndex = 0;
  let seat = 1;
  for (const student of examStudents) {
    const room = rooms[roomIndex];
    assignments.push(makeSubjectAssignment(student, room, seat, subject, "考试"));
    seat += 1;
    if (seat > room.capacity) {
      roomIndex += 1;
      seat = 1;
    }
  }
  if (seat > 1) {
    roomIndex += 1;
    seat = 1;
  }
  for (const student of selfStudyStudents) {
    const room = rooms[roomIndex];
    assignments.push(makeSubjectAssignment(student, room, seat, subject, "自习"));
    seat += 1;
    if (seat > room.capacity) {
      roomIndex += 1;
      seat = 1;
    }
  }
  return assignments;
}

function makeSubjectAssignment(student, room, seatNo, subject, status) {
  return {
    ...makeAssignment(student, room, seatNo, subject, status),
    plan: subject,
    subjectLabel: subject,
    status,
    subjectScore: student.subjectScores?.[subject] ?? "",
  };
}

function subjectStudentSort(a, b, subject) {
  return (
    (b.subjectScores?.[subject] ?? Number.NEGATIVE_INFINITY) - (a.subjectScores?.[subject] ?? Number.NEGATIVE_INFINITY) ||
    b.mathScore - a.mathScore ||
    b.foreignScore - a.foreignScore ||
    a.id.localeCompare(b.id, "zh-Hans-CN", { numeric: true })
  );
}

function selfStudySort(a, b) {
  return a.className.localeCompare(b.className, "zh-Hans-CN", { numeric: true }) || a.id.localeCompare(b.id, "zh-Hans-CN", { numeric: true });
}

export function getElectiveSubjects(student) {
  const parts = safeString(student.comboRaw)
    .split(/[,，、/ ]+/)
    .map((part) => part.trim())
    .filter(Boolean);
  return parts.filter((part) => ELECTIVE_SUBJECTS.includes(part)).sort((a, b) => ELECTIVE_SUBJECTS.indexOf(a) - ELECTIVE_SUBJECTS.indexOf(b));
}

function mergeStudentRows(students, mainAssignments, foreignAssignments, electiveAssignments, subjectAssignments, examTimes, mode) {
  const mainMap = new Map(mainAssignments.map((item) => [item.studentId, item]));
  const foreignMap = new Map(foreignAssignments.map((item) => [item.studentId, item]));
  const electiveMap = new Map(electiveAssignments.map((item) => [item.studentId, item]));
  const subjectMaps = Object.fromEntries(ELECTIVE_SUBJECTS.map((subject) => [subject, new Map(subjectAssignments.filter((item) => item.subjectLabel === subject).map((item) => [item.studentId, item]))]));
  return [...students]
    .sort((a, b) => a.className.localeCompare(b.className, "zh-Hans-CN", { numeric: true }) || a.id.localeCompare(b.id, "zh-Hans-CN"))
    .map((student) => {
      const main = mainMap.get(student.id);
      const foreign = foreignMap.get(student.id);
      const elective = electiveMap.get(student.id);
      const baseRow = {
        班级: student.className,
        姓名: student.name,
        考号: student.id,
        首选科目: student.firstSubject,
        选科组合: student.comboRaw,
        外语语种: student.language,
        总分: student.totalScore,
        总分市排名: student.totalRank,
        语数物历考场: formatRoom(main),
        语数物历座位: main?.seatNo ?? "",
        外语考场: formatRoom(foreign),
        外语座位: foreign?.seatNo ?? "",
        考试时间: formatExamTimes(examTimes),
      };
      if (mode === SCHEDULE_MODES.TWO_DAY_COMBO) {
        return {
          ...baseRow,
          四选二考场: formatRoom(elective),
          四选二座位: elective?.seatNo ?? "",
        };
      }
      return {
        ...baseRow,
        ...Object.fromEntries(
          ELECTIVE_SUBJECTS.flatMap((subject) => {
            const assignment = subjectMaps[subject].get(student.id);
            return [
              [`${subject}状态`, assignment?.status ?? ""],
              [`${subject}考场座位`, assignment ? `${formatRoom(assignment)} ${assignment.seatNo}座` : ""],
            ];
          }),
        ),
      };
    });
}

function formatRoom(assignment) {
  if (!assignment) return "";
  return [assignment.roomNo ? `第${assignment.roomNo}考场` : "", assignment.doorNo, assignment.roomName].filter(Boolean).join(" ");
}

function formatExamTimes(examTimes = []) {
  return examTimes
    .filter((item) => item.subject && (item.date || item.start || item.end))
    .map((item) => `${item.subject} ${[item.date, item.start && item.end ? `${item.start}-${item.end}` : item.start || item.end].filter(Boolean).join(" ")}`)
    .join("；");
}

function buildSummary(mainAssignments, foreignAssignments, electiveAssignments, subjectAssignments, rooms, mode) {
  const lines = [];
  lines.push(...summarizePlan("语数物/历", mainAssignments));
  lines.push(...summarizePlan("外语", foreignAssignments));
  if (mode === SCHEDULE_MODES.TWO_DAY_COMBO) {
    lines.push(...summarizePlan("四选二", electiveAssignments));
  } else {
    for (const subject of ELECTIVE_SUBJECTS) {
      const rows = subjectAssignments.filter((item) => item.subjectLabel === subject);
      const examRows = rows.filter((item) => item.status === "考试");
      const selfRows = rows.filter((item) => item.status === "自习");
      lines.push(`${subject}：考试 ${examRows.length} 人，自习 ${selfRows.length} 人，考试考场 ${roomRange(examRows)}，自习考场 ${roomRange(selfRows)}`);
    }
  }
  lines.push(...summarizeRooms(mainAssignments, rooms).map((item) => `语数物/历第${item.roomNo}考场：${item.count}/${item.capacity} 人，空余 ${Math.max(0, item.capacity - item.count)}`));
  if (mode === SCHEDULE_MODES.TWO_DAY_COMBO) {
    lines.push(...summarizeRooms(electiveAssignments, rooms).map((item) => `四选二第${item.roomNo}考场：${describeAssignments(electiveAssignments.filter((a) => a.roomNo === item.roomNo))}`));
  }
  return lines;
}

function roomRange(assignments) {
  if (!assignments.length) return "无";
  const rooms = [...new Set(assignments.map((item) => item.roomNo))].sort((a, b) => safeString(a).localeCompare(safeString(b), "zh-Hans-CN", { numeric: true }));
  return rooms.length === 1 ? `第${rooms[0]}考场` : `第${rooms[0]}-${rooms[rooms.length - 1]}考场`;
}

function summarizePlan(label, assignments) {
  const counter = new Map();
  for (const item of assignments) {
    const key = item.subjectLabel || label;
    counter.set(key, (counter.get(key) || 0) + 1);
  }
  return [...counter.entries()].map(([key, count]) => `${label}-${key}：${count} 人`);
}

function summarizeRooms(assignments, rooms) {
  const counts = new Map();
  assignments.forEach((item) => counts.set(item.roomNo, (counts.get(item.roomNo) || 0) + 1));
  return rooms.map((room) => ({ ...room, count: counts.get(room.roomNo) || 0 })).filter((item) => item.count);
}

function validateGenerated({ mainAssignments, foreignAssignments, electiveAssignments, subjectAssignments, allStudents, mode, rooms = [], examTimes = [] }) {
  const errors = [];
  const basicPlans = [
    ["主考", mainAssignments],
    ["外语", foreignAssignments],
  ];
  if (mode === SCHEDULE_MODES.TWO_DAY_COMBO) basicPlans.push(["四选二", electiveAssignments]);
  for (const [label, assignments] of basicPlans) {
    const ids = new Set(assignments.map((item) => item.studentId));
    if (ids.size !== allStudents.length || assignments.length !== allStudents.length) {
      errors.push(`${label}存在漏排或重复排：应排 ${allStudents.length} 人，实际 ${assignments.length} 条、唯一学生 ${ids.size} 人`);
    }
  }
  const mainByRoom = new Map();
  for (const item of mainAssignments) {
    if (!mainByRoom.has(item.roomNo)) mainByRoom.set(item.roomNo, new Set());
    mainByRoom.get(item.roomNo).add(item.firstSubject);
  }
  for (const [roomNo, subjects] of mainByRoom.entries()) {
    if (subjects.has("物理") && subjects.has("历史")) errors.push(`主考第${roomNo}考场混入物理和历史`);
  }
  if (mode === SCHEDULE_MODES.THREE_DAY_SPLIT) {
    for (const subject of ELECTIVE_SUBJECTS) {
      const rows = subjectAssignments.filter((item) => item.subjectLabel === subject);
      const ids = new Set(rows.map((item) => item.studentId));
      if (rows.length !== allStudents.length || ids.size !== allStudents.length) {
        errors.push(`${subject}存在漏排或重复排：应排 ${allStudents.length} 人，实际 ${rows.length} 条、唯一学生 ${ids.size} 人`);
      }
      const byRoom = new Map();
      for (const item of rows) {
        if (!byRoom.has(item.roomNo)) byRoom.set(item.roomNo, new Set());
        byRoom.get(item.roomNo).add(item.status);
      }
      for (const [roomNo, statuses] of byRoom.entries()) {
        if (statuses.has("考试") && statuses.has("自习")) errors.push(`${subject}第${roomNo}考场混入考试和自习`);
      }
    }
    for (const student of allStudents) {
      const rows = subjectAssignments.filter((item) => item.studentId === student.id);
      const examCount = rows.filter((item) => item.status === "考试").length;
      const selfCount = rows.filter((item) => item.status === "自习").length;
      if (examCount !== 2 || selfCount !== 2) errors.push(`${student.name || student.id} 再选科安排异常：考试${examCount}门，自习${selfCount}门`);
    }
  }
  errors.push(...validateCapacity([...mainAssignments, ...foreignAssignments, ...electiveAssignments, ...subjectAssignments], rooms));
  errors.push(...validateSeatConflicts([...mainAssignments, ...foreignAssignments, ...electiveAssignments, ...subjectAssignments]));
  errors.push(...validateTimeRoomConflicts({ mainAssignments, foreignAssignments, electiveAssignments, subjectAssignments, mode, examTimes }));
  return errors;
}

function validateCapacity(assignments, rooms) {
  const errors = [];
  const roomMap = new Map(rooms.map((room) => [safeString(room.roomNo), room]));
  const counts = new Map();
  for (const item of assignments) {
    if (isMinorLanguageAssignment(item)) continue;
    const key = `${item.plan}|${item.subjectLabel}|${item.status || ""}|${item.roomNo}`;
    counts.set(key, (counts.get(key) || 0) + 1);
  }
  for (const [key, count] of counts.entries()) {
    const [plan, subjectLabel, status, roomNo] = key.split("|");
    const room = roomMap.get(roomNo);
    if (room && count > room.capacity) errors.push(`${planLabel({ plan, subjectLabel, status })}第${roomNo}考场超容量：${count}/${room.capacity}`);
  }
  return errors;
}

function validateSeatConflicts(assignments) {
  const errors = [];
  const seats = new Map();
  for (const item of assignments) {
    const key = `${item.plan}|${item.subjectLabel}|${item.status || ""}|${item.roomNo}|${item.seatNo}`;
    if (!seats.has(key)) seats.set(key, []);
    seats.get(key).push(item);
  }
  for (const [key, rows] of seats.entries()) {
    if (rows.length > 1) {
      const [plan, subjectLabel, status, roomNo, seatNo] = key.split("|");
      errors.push(`${planLabel({ plan, subjectLabel, status })}第${roomNo}考场${seatNo}座重复：${rows.map((row) => `${row.name || row.studentId}`).join("、")}`);
    }
  }
  return errors;
}

function validateTimeRoomConflicts({ mainAssignments, foreignAssignments, electiveAssignments, subjectAssignments, mode, examTimes }) {
  const errors = [];
  const sessions = buildSessionGroups({ mainAssignments, foreignAssignments, electiveAssignments, subjectAssignments, mode, examTimes });
  for (const session of sessions) {
    const byLocation = new Map();
    for (const group of session.groups) {
      const locationKey = getLocationKey(group.room);
      if (!locationKey) continue;
      if (!byLocation.has(locationKey)) byLocation.set(locationKey, []);
      byLocation.get(locationKey).push(group);
    }
    for (const [locationKey, groups] of byLocation.entries()) {
      const occupied = groups.filter((group) => group.assignments.length);
      if (occupied.length > 1) {
        const uniqueSources = [...new Set(occupied.map((group) => group.label))];
        if (uniqueSources.length > 1) {
          errors.push(`${locationKey} ${session.label}同时用于 ${uniqueSources.join(" 和 ")}`);
        }
      }
    }
  }
  return [...new Set(errors)];
}

export function buildValidationReport({ schedule, rooms = [], importErrors = [] }) {
  const blockers = [...new Set([...(importErrors || []), ...(schedule.errors || [])])].map((detail) =>
    validationRow(VALIDATION_LEVELS.BLOCKER, "硬错误", "需处理", detail, suggestValidationAction(detail)),
  );
  const warnings = buildValidationWarnings(schedule, rooms);
  const reviews = buildValidationReviews(schedule, rooms);
  const okRow = blockers.length
    ? []
    : [validationRow(VALIDATION_LEVELS.REVIEW, "总体状态", "可用于导出/打印", warnings.length ? `通过硬校验，另有 ${warnings.length} 条风险提醒` : "通过硬校验，未发现风险提醒", "导出前快速查看下方复核项")];
  return [...okRow, ...blockers, ...warnings, ...reviews];
}

export function summarizeValidationReport(rows) {
  return {
    blockers: rows.filter((row) => row.级别 === VALIDATION_LEVELS.BLOCKER).length,
    warnings: rows.filter((row) => row.级别 === VALIDATION_LEVELS.WARNING).length,
    reviews: rows.filter((row) => row.级别 === VALIDATION_LEVELS.REVIEW).length,
  };
}

function buildValidationWarnings(schedule, rooms) {
  const warnings = [];
  const enabledRooms = rooms.filter((room) => room.enabled !== false);
  const usedRoomNos = new Set(buildRoomDetailRows(schedule).map((row) => safeString(row.考场号)));
  for (const room of enabledRooms) {
    if (!usedRoomNos.has(safeString(room.roomNo))) continue;
    if (!safeString(room.doorNo)) warnings.push(validationRow(VALIDATION_LEVELS.WARNING, "门牌信息", "需复核", `第${room.roomNo}考场门牌号为空`, "在“确认考场”里补充门牌号"));
    if (!safeString(room.roomName)) warnings.push(validationRow(VALIDATION_LEVELS.WARNING, "教室信息", "需复核", `第${room.roomNo}考场教室名为空`, "在“确认考场”里补充教室名"));
  }
  const printHeaders = Object.keys(buildPrintRows(schedule.allRows)[0] || {});
  if (printHeaders.length > 14) {
    warnings.push(validationRow(VALIDATION_LEVELS.WARNING, "打印宽度", "需复核", `年级打印表共有 ${printHeaders.length} 列，A4 横向可能较紧`, "优先打印班主任表或科目表"));
  }
  for (const [className, count] of countByValue(schedule.allRows, "班级").entries()) {
    if (count > 45) warnings.push(validationRow(VALIDATION_LEVELS.WARNING, "班主任表", "需复核", `${className} 共 ${count} 人，单班打印可能超过一页`, "导出后检查该班 sheet 分页"));
  }
  warnings.push(validationRow(VALIDATION_LEVELS.WARNING, "敏感列", "已隔离", "管理总表含总分/排名；教师打印表不含总分、排名、单科分数", "对外发送时使用年级打印表、班主任表、科目表或考场信息表"));
  return warnings;
}

function buildValidationReviews(schedule, rooms) {
  const rows = [];
  rows.push(validationRow(VALIDATION_LEVELS.REVIEW, "学生人数", "复核", `总人数 ${schedule.allRows.length}，物理 ${schedule.mainAssignments.filter((item) => item.firstSubject === "物理").length}，历史 ${schedule.mainAssignments.filter((item) => item.firstSubject === "历史").length}`, "核对是否等于本次考试报名人数"));
  rows.push(...summarizePlan("语数物/历", schedule.mainAssignments).map((detail) => validationRow(VALIDATION_LEVELS.REVIEW, "语数物/历人数", "复核", detail, "核对语文、数学、物理/历史人数")));
  rows.push(...summarizePlan("外语", schedule.foreignAssignments).map((detail) => validationRow(VALIDATION_LEVELS.REVIEW, "外语人数", "复核", detail, "核对英语和其他语种人数")));
  if (schedule.mode === SCHEDULE_MODES.THREE_DAY_SPLIT) {
    for (const subject of ELECTIVE_SUBJECTS) {
      const subjectRows = schedule.subjectAssignments.filter((item) => item.subjectLabel === subject);
      const examRows = subjectRows.filter((item) => item.status === "考试");
      const selfRows = subjectRows.filter((item) => item.status === "自习");
      rows.push(validationRow(VALIDATION_LEVELS.REVIEW, `${subject}安排`, "复核", `考试 ${examRows.length} 人 ${roomRange(examRows)}；自习 ${selfRows.length} 人 ${roomRange(selfRows)}`, "核对考试人数、自习人数和考场范围"));
    }
  }
  for (const item of summarizeRooms(schedule.mainAssignments, rooms)) {
    rows.push(validationRow(VALIDATION_LEVELS.REVIEW, "语数物/历空座", "复核", `第${item.roomNo}考场 ${item.count}/${item.capacity} 人，空余 ${Math.max(0, item.capacity - item.count)}`, "核对语数物/历考场人数"));
  }
  for (const [language, assignments] of groupBy(schedule.foreignAssignments.filter((item) => item.subjectLabel !== "英语"), "subjectLabel").entries()) {
    rows.push(validationRow(VALIDATION_LEVELS.REVIEW, "外语语种座位", "复核", `${language} ${assignments.length} 人，考场 ${roomRange(assignments)}，座位 1-${assignments.length}`, "核对该语种门牌和座位范围"));
  }
  return rows;
}

function validationRow(level, item, result, detail, action) {
  return { 级别: level, 项目: item, 结果: result, 详情: detail, 建议处理: action };
}

function suggestValidationAction(message) {
  if (message.includes("同时用于")) return "调整同一时段冲突的考场号/门牌号/外语语种考场";
  if (message.includes("座重复")) return "重新生成或检查该考场座位号";
  if (message.includes("超容量") || message.includes("容量不足")) return "新增考场或调整容量";
  if (message.includes("漏排") || message.includes("重复排")) return "检查学生名单和排考规则";
  if (message.includes("混入") || message.includes("混场")) return "检查首选科目、再选科和考场边界";
  if (message.includes("小语种") || message.includes("未指定")) return "设置外语语种考场";
  if (message.includes("缺少") || message.includes("缺")) return "补齐导入字段或学生信息";
  return "按提示修正后重新生成";
}

function isMinorLanguageAssignment(item) {
  return item.plan === "外语" && item.subjectLabel && item.subjectLabel !== "英语";
}

function planLabel({ plan, subjectLabel, status }) {
  if (plan === subjectLabel && status) return `${subjectLabel}${status}`;
  if (subjectLabel && subjectLabel !== plan) return `${plan}-${subjectLabel}`;
  return plan || subjectLabel || "排考";
}

function buildSessionGroups({ mainAssignments, foreignAssignments, electiveAssignments, subjectAssignments, mode, examTimes }) {
  const sessions = [
    {
      label: getSessionLabel(examTimes, "语文", "语数物历时段"),
      groups: groupAssignmentsByRoom(mainAssignments, "主考"),
    },
    {
      label: getSessionLabel(examTimes, "外语", "外语时段"),
      groups: groupAssignmentsByRoom(foreignAssignments, "外语"),
    },
  ];
  if (mode === SCHEDULE_MODES.TWO_DAY_COMBO) {
    sessions.push({
      label: getSessionLabel(examTimes, "四选二", "四选二时段"),
      groups: groupAssignmentsByRoom(electiveAssignments, "四选二"),
    });
  } else {
    for (const subject of ELECTIVE_SUBJECTS) {
      sessions.push({
        label: getSessionLabel(examTimes, subject, `${subject}时段`),
        groups: groupAssignmentsByRoom(subjectAssignments.filter((item) => item.subjectLabel === subject), subject),
      });
    }
  }
  return sessions;
}

function groupAssignmentsByRoom(assignments, label) {
  const roomGroups = new Map();
  for (const item of assignments) {
    const key = safeString(item.roomNo);
    if (!roomGroups.has(key)) {
      roomGroups.set(key, {
        label: `${item.subjectLabel || label}考场${item.roomNo}`,
        room: {
          roomNo: item.roomNo,
          doorNo: item.doorNo,
          roomName: item.roomName,
        },
        assignments: [],
      });
    }
    roomGroups.get(key).assignments.push(item);
  }
  return [...roomGroups.values()];
}

function getSessionLabel(examTimes = [], subject, fallback) {
  const match = examTimes.find((item) => item.subject === subject || (subject === "语文" && item.subject === "数学") || (subject === "四选二" && ELECTIVE_SUBJECTS.includes(item.subject)));
  if (!match) return fallback;
  const time = [match.date, match.start && match.end ? `${match.start}-${match.end}` : match.start || match.end].filter(Boolean).join(" ");
  return time ? `${match.subject} ${time}` : fallback;
}

function getLocationKey(room) {
  if (safeString(room.doorNo)) return `门牌${safeString(room.doorNo)}`;
  if (safeString(room.roomNo)) return `考场${safeString(room.roomNo)}`;
  if (safeString(room.roomName)) return `教室${safeString(room.roomName)}`;
  return "";
}

function countByValue(rows, field) {
  const counter = new Map();
  for (const row of rows) {
    const key = safeString(row[field]) || "未填写";
    counter.set(key, (counter.get(key) || 0) + 1);
  }
  return counter;
}

function countMinorLanguages(assignments) {
  const counter = new Map();
  for (const item of assignments) {
    if (item.subjectLabel && item.subjectLabel !== "英语") {
      counter.set(item.subjectLabel, (counter.get(item.subjectLabel) || 0) + 1);
    }
  }
  return counter;
}

function groupBy(rows, field) {
  const groups = new Map();
  for (const row of rows) {
    const key = row[field];
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(row);
  }
  return groups;
}

export async function buildWorkbookFile({ examName, examDate, schedule, rooms, selected = {}, group = EXPORT_GROUPS.ALL, fileNameSuffix = "", printSettingsBySheet = {} }) {
  const workbook = buildWorkbook({ schedule, rooms, selected, group, printSettingsBySheet });
  ensureWorkbookHasSheet(workbook);
  const fileName = buildExportFileName(examName, examDate, fileNameSuffix || defaultExportSuffix(group));
  const rawData = XLSX.write(workbook, { bookType: "xlsx", type: "array", compression: true });
  const data = await patchWorkbookData(rawData, workbook);
  return { workbook, fileName, data };
}

function ensureWorkbookHasSheet(workbook) {
  if (workbook.SheetNames.length) return;
  appendSheet(workbook, "暂无数据", [{ 提示: "暂无可导出的内容，请先完成导入和生成排考。" }], { orientation: "portrait", profile: "print" });
}

export async function exportWorkbook({ examName, examDate, schedule, rooms, selected = {}, group = EXPORT_GROUPS.ALL, fileNameSuffix = "", printSettingsBySheet = {} }) {
  return buildWorkbookFile({ examName, examDate, schedule, rooms, selected, group, fileNameSuffix, printSettingsBySheet });
}

async function patchWorkbookData(arrayBuffer, workbook) {
  const zip = await JSZip.loadAsync(arrayBuffer);
  for (const [index, sheetName] of workbook.SheetNames.entries()) {
    const sheet = workbook.Sheets[sheetName];
    const settings = sheet?.["!pageSetup"];
    if (!settings) continue;
    const path = `xl/worksheets/sheet${index + 1}.xml`;
    const file = zip.file(path);
    if (!file) continue;
    const xml = await file.async("string");
    zip.file(path, injectPageSetup(xml, { ...settings, margins: sheet?.["!margins"], printOptions: sheet?.["!printOptions"] }));
  }
  const patched = await zip.generateAsync({ type: "uint8array" });
  return patched;
}

function injectPageSetup(xml, settings = {}) {
  const fitToWidth = settings.fitToWidth ?? 1;
  const fitToHeight = settings.fitToHeight ?? 0;
  const scale = settings.scale ?? undefined;
  const orientation = settings.orientation || "portrait";
  const pageSetup = `<pageSetup paperSize="${settings.paperSize || 9}" orientation="${orientation}" fitToWidth="${fitToWidth}" fitToHeight="${fitToHeight}"${scale ? ` scale="${scale}"` : ""}/>`;
  const margins = settings.margins || {};
  const pageMargins = `<pageMargins left="${margins.left ?? 0.25}" right="${margins.right ?? 0.25}" top="${margins.top ?? 0.35}" bottom="${margins.bottom ?? 0.35}" header="${margins.header ?? 0.1}" footer="${margins.footer ?? 0.1}"/>`;
  const printOptions = settings.printOptions?.horizontalCentered ? `<printOptions horizontalCentered="1" verticalCentered="${settings.printOptions?.verticalCentered ? 1 : 0}"/>` : "";
  const sheetPr = `<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>`;
  let next = xml;
  if (!next.includes("<sheetPr")) {
    next = next.replace(/(<worksheet[^>]*>)/, `$1${sheetPr}`);
  } else if (!next.includes("<pageSetUpPr")) {
    next = next.replace(/<sheetPr>(.*?)<\/sheetPr>/s, "<sheetPr><pageSetUpPr fitToPage=\"1\"/></sheetPr>");
  }
  if (next.includes("<pageSetup")) {
    next = next.replace(/<pageSetup[^>]*\/>/, pageSetup);
  } else if (next.includes("<pageMargins")) {
    next = next.replace(/<pageMargins[^>]*\/>/, (match) => `${pageSetup}${match}`);
  } else {
    next = next.replace(/<\/worksheet>$/, `${pageSetup}</worksheet>`);
  }
  if (next.includes("<pageMargins")) {
    next = next.replace(/<pageMargins[^>]*\/>/, pageMargins);
  } else {
    next = next.replace(/<pageSetup[^>]*\/>/, (match) => `${match}${pageMargins}`);
  }
  if (printOptions && !next.includes("<printOptions")) {
    next = next.replace(/<pageMargins[^>]*\/>/, (match) => `${printOptions}${match}`);
  }
  return next;
}

export function buildWorkbook({ schedule, rooms, selected = {}, group = EXPORT_GROUPS.ALL, printSettingsBySheet = {} }) {
  const workbook = XLSX.utils.book_new();
  const enabled = {
    grade: true,
    classes: true,
    roomSummary: false,
    roomSheets: true,
    doorSummary: true,
    subjectSheets: false,
    ...selected,
  };

  if (group === EXPORT_GROUPS.ALL || group === EXPORT_GROUPS.VALIDATION) {
    appendSheet(workbook, "校验报告", buildValidationReport({ schedule, rooms }), { orientation: "landscape", profile: "validation", title: "校验报告" });
  }
  if ((group === EXPORT_GROUPS.ALL && enabled.doorSummary) || group === EXPORT_GROUPS.DOOR_SUMMARY) {
    if (enabled.doorSummary) appendSheet(workbook, "门牌人数总览", buildDoorRows(schedule, rooms), { orientation: "landscape", profile: "overview" });
  }
  if ((group === EXPORT_GROUPS.ALL && enabled.grade) || group === EXPORT_GROUPS.ADMIN) {
    appendSheet(workbook, "管理总表", buildAdminRows(schedule.allRows), { orientation: "landscape", profile: "admin" });
  }
  if ((group === EXPORT_GROUPS.ALL && enabled.classes) || group === EXPORT_GROUPS.CLASS) {
    if (enabled.classes) appendClassSheets(workbook, buildPrintRows(schedule.allRows), printSettingsBySheet);
  }
  if ((group === EXPORT_GROUPS.ALL && enabled.subjectSheets !== false) || group === EXPORT_GROUPS.SUBJECT) {
    if (schedule.mode === SCHEDULE_MODES.THREE_DAY_SPLIT) appendSubjectSheets(workbook, schedule);
  }
  if ((group === EXPORT_GROUPS.ALL && enabled.roomSheets) || group === EXPORT_GROUPS.ROOM) {
    if (enabled.roomSheets) appendRoomSheets(workbook, schedule, printSettingsBySheet);
  }
  if ((group === EXPORT_GROUPS.ALL && enabled.timeSheet !== false) || group === EXPORT_GROUPS.TIME) {
    appendSheet(workbook, "考试时间", buildTimeSheetRows(schedule.examTimes || []), { orientation: "portrait", profile: "timePrint", title: "考试时间表" });
  }
  return workbook;
}

export function buildExportJobs({ examName, examDate, schedule, rooms, selected = {}, exportMode = EXPORT_GROUPS.ALL }) {
  if (exportMode === EXPORT_GROUPS.ALL) {
    return [{ group: EXPORT_GROUPS.VALIDATION, fileNameSuffix: "校验报告" }, { group: EXPORT_GROUPS.ALL, fileNameSuffix: "考场安排" }];
  }
  const jobs = [{ group: EXPORT_GROUPS.VALIDATION, fileNameSuffix: "校验报告" }];
  if (selected.grade) jobs.push({ group: EXPORT_GROUPS.ADMIN, fileNameSuffix: "管理表" });
  if (selected.classes) jobs.push({ group: EXPORT_GROUPS.CLASS, fileNameSuffix: "班主任表" });
  if (selected.subjectSheets && schedule.mode === SCHEDULE_MODES.THREE_DAY_SPLIT) jobs.push({ group: EXPORT_GROUPS.SUBJECT, fileNameSuffix: "科目表" });
  if (selected.roomSheets) jobs.push({ group: EXPORT_GROUPS.ROOM, fileNameSuffix: "考场信息表" });
  if (selected.doorSummary) jobs.push({ group: EXPORT_GROUPS.DOOR_SUMMARY, fileNameSuffix: "门牌人数" });
  if (selected.timeSheet) jobs.push({ group: EXPORT_GROUPS.TIME, fileNameSuffix: "考试时间" });
  return jobs;
}

export function buildExportFileName(examName, examDate, suffix) {
  const datePart = examDate || new Date().toISOString().slice(0, 10);
  return `${examName || "考试"}_${datePart}_${suffix}.xlsx`;
}

function defaultExportSuffix(group) {
  if (group === EXPORT_GROUPS.ADMIN) return "管理表";
  if (group === EXPORT_GROUPS.CLASS) return "班主任表";
  if (group === EXPORT_GROUPS.ROOM) return "考场信息表";
  if (group === EXPORT_GROUPS.SUBJECT) return "科目表";
  if (group === EXPORT_GROUPS.DOOR_SUMMARY) return "门牌人数";
  if (group === EXPORT_GROUPS.TIME) return "考试时间";
  if (group === EXPORT_GROUPS.VALIDATION) return "校验报告";
  return "考场安排";
}

function appendSheet(workbook, name, rows, options = {}) {
  const bodyRows = rows.length ? rows : [{ 提示: "暂无数据" }];
  const headers = Object.keys(bodyRows[0]).filter((header) => !header.startsWith("__"));
  const visibleRows = bodyRows.map((row) => Object.fromEntries(headers.map((header) => [header, row[header]])));
  const titleRows = options.title ? [[options.title], ...(options.note ? [[options.note]] : [])] : [];
  const sheet = options.title ? XLSX.utils.aoa_to_sheet(titleRows) : XLSX.utils.json_to_sheet(visibleRows);
  if (options.title) {
    XLSX.utils.sheet_add_json(sheet, visibleRows, { origin: `A${titleRows.length + 1}`, skipHeader: false });
  }
  setSheetPrintDefaults(sheet, headers, bodyRows, { ...options, titleRowCount: titleRows.length });
  if (options.title) ensureTitleMerge(sheet, headers, options.title, options.note);
  XLSX.utils.book_append_sheet(workbook, sheet, safeSheetName(name));
}

function ensureTitleMerge(sheet, headers, title, note = "") {
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");
  range.e.r = Math.max(range.e.r, note ? 2 : 1);
  range.e.c = Math.max(range.e.c, headers.length - 1);
  sheet["!ref"] = XLSX.utils.encode_range(range);
  sheet["!merges"] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: Math.max(0, headers.length - 1) } },
    ...(note ? [{ s: { r: 1, c: 0 }, e: { r: 1, c: Math.max(0, headers.length - 1) } }] : []),
  ];
  const titleAddress = XLSX.utils.encode_cell({ r: 0, c: 0 });
  sheet[titleAddress] = sheet[titleAddress] || { t: "s", v: title };
  sheet[titleAddress].v = title;
  sheet[titleAddress].s = {
    font: { name: "宋体", sz: 16, bold: true, color: { rgb: "1D1D1F" } },
    alignment: { horizontal: "center", vertical: "center", wrapText: true },
    fill: { fgColor: { rgb: "EAF2FF" } },
  };
  if (note) {
    const noteAddress = XLSX.utils.encode_cell({ r: 1, c: 0 });
    sheet[noteAddress] = sheet[noteAddress] || { t: "s", v: note };
    sheet[noteAddress].v = note;
    sheet[noteAddress].s = {
      font: { name: "宋体", sz: 9, color: { rgb: "3A3A3C" } },
      alignment: { horizontal: "center", vertical: "center", wrapText: true },
      fill: { fgColor: { rgb: "F7FAFF" } },
    };
  }
}

function appendClassSheets(workbook, rows, printSettingsBySheet = {}) {
  const classes = [...new Set(rows.map((row) => row.__className))].sort((a, b) => a.localeCompare(b, "zh-Hans-CN", { numeric: true }));
  if (!classes.length) {
    appendSheet(workbook, "班主任表", [{ 提示: "暂无班级数据，请先导入学生并生成排考。" }], { orientation: "portrait", profile: "print", title: "班主任表" });
    return;
  }
  for (const className of classes) {
    const classRows = rows
      .filter((row) => row.__className === className)
      .map(({ __className, ...row }) => row);
    const defaultOrientation = recommendExportOrientation(classRows, "classPrint");
    appendSheet(workbook, safeSheetName(className), classRows, {
      orientation: getExportOrientation(printSettingsBySheet, "classes", className, defaultOrientation),
      profile: "classPrint",
      title: `${className}考场安排`,
      note: "说明：语数物/座位号、语数历/座位号、外语、化学、地理、政治、生物均为“考场/座位号”；黄色底色表示该科为自习安排。",
    });
  }
}

function appendSubjectSheets(workbook, schedule) {
  for (const subject of ELECTIVE_SUBJECTS) {
    const rows = buildSubjectPrintRows(schedule, subject);
    appendSheet(workbook, subject, rows, { orientation: "portrait" });
  }
}

function appendRoomSheets(workbook, schedule, printSettingsBySheet = {}) {
  const groups = buildRoomSheetGroups(schedule);
  if (!groups.length) {
    appendSheet(workbook, "考场信息表", [{ 提示: "暂无考场安排，请先完成排考。" }], { orientation: "portrait", profile: "print", title: "考场信息表" });
    return;
  }
  for (const group of groups) {
    appendSheet(workbook, group.name, group.rows, { orientation: getRoomSheetExportOrientation(printSettingsBySheet, group, recommendExportOrientation(group.rows, "roomPrint")), profile: "roomPrint", title: `${group.name}考场信息表` });
  }
}

function getExportOrientation(settingsBySheet, tabKey, pageKey, fallback) {
  return settingsBySheet?.[`${tabKey}:${pageKey}`]?.orientation || fallback;
}

function getRoomSheetExportOrientation(settingsBySheet, group, fallback) {
  return settingsBySheet?.[`roomDetails:${group.name}`]?.orientation ||
    (group.foreignPreviewKey ? settingsBySheet?.[`foreign:${group.foreignPreviewKey}`]?.orientation : "") ||
    fallback;
}

function recommendExportOrientation(rows, profile = "") {
  const headers = Object.keys(rows[0] || {}).filter((header) => !header.startsWith("__"));
  if (profile === "roomPrint" && headers.length <= 6 && rows.length <= 46) return "portrait";
  if (profile === "classPrint" && headers.length <= 8 && rows.length <= 42) return "portrait";
  return headers.length >= 8 ? "landscape" : "portrait";
}

function statusSort(a, b) {
  const order = { 考试: 0, 自习: 1 };
  return (order[a.status] ?? 9) - (order[b.status] ?? 9);
}

export function buildAdminRows(rows) {
  return rows.map((row) => ({
    班级: row.班级,
    姓名: row.姓名,
    考号: row.考号,
    首选科目: row.首选科目,
    选科组合: row.选科组合,
    外语语种: row.外语语种,
    总分: row.总分,
    总分市排名: row.总分市排名,
    语数物历考场: parseRoomNo(row.语数物历考场),
    语数物历座位: row.语数物历座位,
    外语考场: parseRoomNo(row.外语考场),
    外语座位: row.外语座位,
    化学考场: parseSubjectRoom(row, "化学"),
    化学座位: parseSubjectSeat(row, "化学"),
    地理考场: parseSubjectRoom(row, "地理"),
    地理座位: parseSubjectSeat(row, "地理"),
    政治考场: parseSubjectRoom(row, "政治"),
    政治座位: parseSubjectSeat(row, "政治"),
    生物考场: parseSubjectRoom(row, "生物"),
    生物座位: parseSubjectSeat(row, "生物"),
  }));
}

export function buildPrintRows(rows) {
  return rows.map((row) => {
    const mainHeader = mainExamHeader(row);
    return {
      姓名: row.姓名,
      考号: row.考号,
      选科: compactComboLabel(row.选科组合),
      [mainHeader]: compactRoomSeat(row.语数物历考场, row.语数物历座位),
      外语: compactRoomSeat(row.外语考场, row.外语座位),
      化学: compactSubjectCell(row, "化学"),
      地理: compactSubjectCell(row, "地理"),
      政治: compactSubjectCell(row, "政治"),
      生物: compactSubjectCell(row, "生物"),
      "__className": row.班级,
      "__selfStudy:化学": row.化学状态 === "自习",
      "__selfStudy:地理": row.地理状态 === "自习",
      "__selfStudy:政治": row.政治状态 === "自习",
      "__selfStudy:生物": row.生物状态 === "自习",
    };
  });
}

export function buildSubjectPrintRows(schedule, subject) {
  return schedule.subjectAssignments
    .filter((item) => item.subjectLabel === subject)
    .sort((a, b) => statusSort(a, b) || safeString(a.roomNo).localeCompare(safeString(b.roomNo), "zh-Hans-CN", { numeric: true }) || a.seatNo - b.seatNo)
    .map((item) => ({
      科目: subject,
      姓名: item.name,
      考号: item.studentId,
      班级: item.className,
      考场号: item.roomNo,
      座位号: item.seatNo,
      当科: subject,
    }));
}

function parseSubjectRoom(row, subject) {
  const value = safeString(row[`${subject}考场座位`]);
  if (!value) return "";
  const roomNo = parseRoomNo(value);
  if (!roomNo) return "";
  return row[`${subject}状态`] === "自习" ? `${roomNo}自习室` : roomNo;
}

function parseSubjectSeat(row, subject) {
  const value = safeString(row[`${subject}考场座位`]);
  const match = value.match(/(\d+)\s*座/);
  return match ? Number(match[1]) : "";
}

function parseRoomNo(value) {
  const match = safeString(value).match(/第?(\d+)考场/);
  return match ? Number(match[1]) : safeString(value);
}

function extractRoomNumber(value) {
  const text = safeString(value);
  const match = text.match(/(\d+)/);
  return match ? Number(match[1]) : text.replace(/自习室$/, "");
}

export function toInvigilatorRow(row, options = {}) {
  const isSelfStudy = row.状态 === "自习";
  const subjectName = displayRoomSheetSubject(row);
  const base = {
    姓名: row.姓名,
    考号: row.考号,
    班级: row.班级,
    考场号: row.考场号,
    座位号: row.座位号,
    当科: isSelfStudy ? "自习" : subjectName,
  };
  const withPlan = options.includePlan ? { 科目: subjectName, ...base } : base;
  return withPlan;
}

export function buildRoomPrintRows(schedule) {
  return buildRoomDetailRows(schedule).map((row) => toInvigilatorRow(row, { includePlan: true }));
}

function buildRoomSheetGroups(schedule) {
  const rows = buildRoomDetailRows(schedule);
  const groupMap = new Map();
  const groupMeta = new Map();
  for (const row of rows) {
    const isSelfStudy = row.状态 === "自习";
    const suffix = isSelfStudy ? `${row.考场号}自习室` : row.考场号;
    const subjectName = displayRoomSheetSubject(row);
    const name = `${subjectName}-${suffix}`;
    if (!groupMap.has(name)) groupMap.set(name, []);
    if (!groupMeta.has(name) && row.考试类型 === "外语") {
      groupMeta.set(name, { foreignPreviewKey: `${subjectName}|${row.考场号}` });
    }
    groupMap.get(name).push(toInvigilatorRow(row));
  }
  return [...groupMap.entries()]
    .sort(([a], [b]) => safeString(a).localeCompare(safeString(b), "zh-Hans-CN", { numeric: true }))
    .map(([name, groupRows]) => ({
      name,
      rows: groupRows.sort((a, b) => toNumber(a.座位号) - toNumber(b.座位号)),
      profile: "roomPrint",
      ...(groupMeta.get(name) || {}),
    }));
}

function displayPlanName(plan) {
  return plan === "主考" ? "语数物历" : plan;
}

function displayRoomSheetSubject(row) {
  if (row.考试类型 === "外语") return row.外语语种 || "外语";
  return displayPlanName(row.考试类型);
}

export function buildRoomSummaryRows(schedule, rooms) {
  const rows = [];
  const subjectLabels = ELECTIVE_SUBJECTS;
  for (const room of buildDoorSummaryRooms(schedule, rooms)) {
    const main = findAssignmentsInRoom(schedule.mainAssignments, room);
    const foreign = findAssignmentsInRoom(schedule.foreignAssignments, room);
    const elective = findAssignmentsInRoom(schedule.electiveAssignments, room);
    const subjectRows = subjectLabels.flatMap((subject) => findAssignmentsInRoom(schedule.subjectAssignments, room).filter((item) => item.subjectLabel === subject));
    if (!main.length && !foreign.length && !elective.length) continue;
    rows.push({
      考场号: room.roomNo,
      门牌号: room.doorNo,
      教室: room.roomName,
      主考人数: main.length,
      外语人数: describeForeignAssignments(foreign),
      四选二内容: describeAssignments(elective),
      三天内容: describeSubjectAssignments(subjectRows),
      备注: "",
    });
  }
  return rows;
}

export function buildDoorRows(schedule, rooms) {
  return buildDoorSummaryRooms(schedule, rooms).map((room) => {
    const main = findAssignmentsInRoom(schedule.mainAssignments, room);
    const foreign = findAssignmentsInRoom(schedule.foreignAssignments, room);
    const elective = findAssignmentsInRoom(schedule.electiveAssignments, room);
    const subjectRows = (subject) => findAssignmentsInRoom(schedule.subjectAssignments, room).filter((item) => item.subjectLabel === subject);
    const isSelfStudySubjectRoom = (subject) => subjectRows(subject).some((item) => item.status === "自习");
    return {
      门牌号: room.doorNo,
      教室: room.roomName,
      考场号: room.roomNo,
      语数物历: main.length || "",
      外语: describeForeignAssignments(foreign),
      化学: subjectRows("化学").length || "",
      地理: subjectRows("地理").length || "",
      政治: subjectRows("政治").length || "",
      生物: subjectRows("生物").length || "",
      四选二: describeAssignments(elective),
      人数: room.capacity,
      "__comboSummary:四选二": describeAssignments(elective),
      "__selfStudy:化学": isSelfStudySubjectRoom("化学"),
      "__selfStudy:地理": isSelfStudySubjectRoom("地理"),
      "__selfStudy:政治": isSelfStudySubjectRoom("政治"),
      "__selfStudy:生物": isSelfStudySubjectRoom("生物"),
    };
  }).filter((row) => row.语数物历 || row.外语 || row.化学 || row.地理 || row.政治 || row.生物 || row.四选二);
}

function buildDoorSummaryRooms(schedule, rooms) {
  const rows = new Map();
  for (const room of rooms || []) {
    rows.set(roomKey(room), { ...room });
  }
  for (const item of [
    ...(schedule.mainAssignments || []),
    ...(schedule.foreignAssignments || []),
    ...(schedule.electiveAssignments || []),
    ...(schedule.subjectAssignments || []),
  ]) {
    const key = roomKey(item);
    if (!key || rows.has(key)) continue;
    rows.set(key, {
      roomNo: item.roomNo,
      doorNo: item.doorNo || "",
      roomName: item.roomName || "",
      capacity: "",
      enabled: true,
    });
  }
  return [...rows.values()].sort(compareRoomNo);
}

function findAssignmentsInRoom(assignments = [], room) {
  return assignments.filter((item) => roomKey(item) === roomKey(room));
}

function roomKey(value = {}) {
  const doorNo = safeString(value.doorNo);
  if (doorNo) return `door:${doorNo}`;
  const roomNo = safeString(value.roomNo);
  return roomNo ? `room:${roomNo}` : "";
}

function describeForeignAssignments(assignments = []) {
  return describeAssignments(assignments);
}

function buildTimeSheetRows(examTimes = []) {
  return examTimes
    .filter((item) => item.subject || item.date || item.start || item.end)
    .map((item, index) => ({
      序号: index + 1,
      科目: item.subject,
      日期: item.date,
      时间: [item.start, item.end].filter(Boolean).join("-"),
    }));
}

export function buildRoomDetailRows(schedule) {
  return [...schedule.mainAssignments, ...schedule.foreignAssignments, ...schedule.electiveAssignments, ...schedule.subjectAssignments]
    .sort((a, b) => safeString(a.roomNo).localeCompare(safeString(b.roomNo), "zh-Hans-CN", { numeric: true }) || a.plan.localeCompare(b.plan, "zh-Hans-CN") || a.seatNo - b.seatNo)
    .map((item) => ({
      考试类型: item.plan,
      考场号: item.roomNo,
      门牌号: item.doorNo,
      教室: item.roomName,
      座位号: item.seatNo,
      班级: item.className,
      姓名: item.name,
      考号: item.studentId,
      首选科目: item.firstSubject,
      选科组合: item.comboRaw,
      外语语种: item.language,
      状态: item.status || "",
      该科分数: item.subjectScore ?? "",
    }));
}

function describeAssignments(assignments) {
  const counter = new Map();
  for (const item of assignments) {
    const key = item.subjectLabel || item.plan;
    counter.set(key, (counter.get(key) || 0) + 1);
  }
  return [...counter.entries()].map(([key, count]) => `${key}${count}`).join("+");
}

function describeSubjectAssignments(assignments) {
  if (!assignments.length) return "";
  const counter = new Map();
  for (const item of assignments) {
    const key = `${item.subjectLabel}${item.status}`;
    counter.set(key, (counter.get(key) || 0) + 1);
  }
  return [...counter.entries()].map(([key, count]) => `${key}${count}`).join("+");
}

function compactComboLabel(value) {
  const combo = safeString(value);
  const subjects = ELECTIVE_SUBJECTS.filter((subject) => combo.includes(subject));
  if (subjects.length === 2) {
    const key = [...subjects].sort().join("|");
    const canonical = COMBO_CANONICAL.get(key) || subjects.join("");
    return COMBO_SHORT_LABELS[canonical] || canonical;
  }
  return COMBO_SHORT_LABELS[combo] || combo.replace(/\s+/g, "");
}

function compactRoomSeat(room, seat) {
  const roomLabel = extractRoomNumber(room);
  if (!roomLabel && !seat) return "";
  return [roomLabel || "", seat ? String(seat) : ""].filter(Boolean).join("/");
}

function compactSubjectCell(row, subject) {
  const roomText = parseSubjectRoom(row, subject);
  const seatText = parseSubjectSeat(row, subject);
  if (!roomText && !seatText) return "";
  return [extractRoomNumber(roomText) || roomText, seatText || ""].filter(Boolean).join("/");
}

function mainExamHeader(row) {
  return safeString(row.首选科目).includes("物理") ? "语数物/座位号" : "语数历/座位号";
}

function setSheetPrintDefaults(sheet, headers = [], rows = [], options = {}) {
  const profile = options.profile || "default";
  const compact = profile === "classPrint" || profile === "roomPrint" || profile === "print" || profile === "timePrint" || profile === "overview";
  sheet["!margins"] = compact
    ? { left: 0.2, right: 0.2, top: 0.28, bottom: 0.28, header: 0.1, footer: 0.1 }
    : { left: 0.18, right: 0.18, top: 0.22, bottom: 0.22, header: 0.1, footer: 0.1 };
  sheet["!cols"] = headers.map((header) => ({ wch: estimateColumnWidth(header, rows, options) }));
  sheet["!freeze"] = { xSplit: 0, ySplit: (options.titleRowCount || 0) + 1 };
  sheet["!pageSetup"] = {
    paperSize: 9,
    orientation: options.orientation || (headers.length > 8 || rows.length > 24 ? "landscape" : "portrait"),
    fitToWidth: 1,
    fitToHeight: 1,
  };
  sheet["!printOptions"] = { horizontalCentered: true, verticalCentered: false };
  sheet["!autofilter"] = sheet["!ref"] ? { ref: sheet["!ref"] } : undefined;
  applyCellStyle(sheet, headers, rows, options);
}

function estimateColumnWidth(header, rows = [], options = {}) {
  const sampleWidth = rows
    .slice(0, 160)
    .reduce((max, row) => Math.max(max, safeString(row[header]).length), safeString(header).length);
  const compact = options.profile === "print" || options.profile === "overview" || options.profile === "timePrint";
  const veryCompact = options.profile === "classPrint" || options.profile === "roomPrint";
  if (options.profile === "timePrint") {
    if (header === "序号") return 8;
    if (header === "科目") return 14;
    if (header === "日期") return 18;
    if (header === "时间") return 24;
  }
  if (options.profile === "overview") {
    if (header === "门牌号") return 10;
    if (header === "教室") return 14;
    if (header === "考场号") return 9;
    return 8;
  }
  if (options.profile === "validation") {
    if (header === "级别") return 10;
    if (header === "项目") return 12;
    if (header === "结果") return 12;
    if (header === "详情") return 38;
    if (header === "建议处理") return 24;
  }
  if (header === "科目") return 10;
  if (header === "姓名") return veryCompact ? 9 : Math.min(Math.max(sampleWidth + 2, 8), 10);
  if (header === "班级") return veryCompact ? 9 : 8;
  if (header === "考号") return veryCompact ? 13 : Math.min(Math.max(sampleWidth + 1, 12), 15);
  if (header === "选科") return veryCompact ? 7 : 10;
  if (header.includes("座位")) return compact ? 6 : 8;
  if (header === "语数物/座位号" || header === "语数历/座位号") return compact ? 10 : 12;
  if (header === "外语" || header === "当科") return compact ? 10 : 12;
  if (header.includes("考场")) return compact ? 8 : 10;
  if (header.includes("组合")) return compact ? 10 : 12;
  if (header.includes("四选二")) return compact ? 18 : 20;
  if (header.includes("时间")) return 34;
  if (header.includes("排名")) return 12;
  if (compact) return Math.min(Math.max(sampleWidth + 1, 6), 12);
  return Math.min(Math.max(sampleWidth + 1, 9), 18);
}

function applyCellStyle(sheet, headers = [], rows = [], options = {}) {
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");
  const titleRowCount = options.titleRowCount || (options.title ? 1 : 0);
  const headerRowIndex = titleRowCount;
  const dataStartRowIndex = titleRowCount + 1;
  const profile = options.profile || "default";
  for (let row = range.s.r; row <= range.e.r; row += 1) {
    for (let col = range.s.c; col <= range.e.c; col += 1) {
      const address = XLSX.utils.encode_cell({ r: row, c: col });
      if (!sheet[address]) continue;
      const header = headers[col] || "";
      const dataRow = rows[row - dataStartRowIndex];
      sheet[address].s = {
        font: {
          name: "宋体",
          sz: getCellFontSize({ row, headerRowIndex, profile, titleRowCount }),
          bold: row < titleRowCount ? row === 0 : row === headerRowIndex,
          color: { rgb: "1D1D1F" },
        },
        alignment: { horizontal: "center", vertical: "center", wrapText: true },
        border: {
          top: { style: "thin", color: { rgb: "000000" } },
          bottom: { style: "thin", color: { rgb: "000000" } },
          left: { style: "thin", color: { rgb: "000000" } },
          right: { style: "thin", color: { rgb: "000000" } },
        },
      };
      if (row === 0 && options.title) {
        sheet[address].s.fill = { fgColor: { rgb: "DCEBFF" } };
      } else if (row > 0 && row < titleRowCount) {
        sheet[address].s.fill = { fgColor: { rgb: "F7FAFF" } };
      } else if (row === headerRowIndex) {
        sheet[address].s.fill = { fgColor: { rgb: "F2F2F7" } };
      } else if (sheet[address].v === "自习" || String(sheet[address].v).includes("自习") || dataRow?.[`__selfStudy:${header}`]) {
        sheet[address].s.fill = { fgColor: { rgb: "FFF1C9" } };
      }
    }
  }
  const rowHeight = getBodyRowHeight(profile, rows.length);
  sheet["!rows"] = Array.from({ length: range.e.r + 1 }, (_, index) => ({
    hpt: getRowHeight({ index, titleRowCount, headerRowIndex, profile, rowHeight }),
  }));
  if (sheet["!merges"]?.length && options.title) {
    sheet["!rows"][0] = { hpt: profile === "timePrint" ? 42 : 28 };
    if (titleRowCount > 1) sheet["!rows"][1] = { hpt: 24 };
  }
  const titleRow = options.title ? 0 : null;
  if (titleRow === 0) {
    for (let titleIndex = 0; titleIndex < titleRowCount; titleIndex += 1) {
      for (let col = range.s.c; col <= range.e.c; col += 1) {
        const address = XLSX.utils.encode_cell({ r: titleIndex, c: col });
        if (!sheet[address]) continue;
        sheet[address].s = {
          font: { name: "宋体", sz: titleIndex === 0 ? (profile === "timePrint" ? 20 : 14) : 9, bold: titleIndex === 0, color: { rgb: "1D1D1F" } },
          alignment: { horizontal: "center", vertical: "center", wrapText: true },
          fill: { fgColor: { rgb: titleIndex === 0 ? "EAF2FF" : "F7FAFF" } },
          border: {
            top: { style: "thin", color: { rgb: "000000" } },
            bottom: { style: "thin", color: { rgb: "000000" } },
            left: { style: "thin", color: { rgb: "000000" } },
            right: { style: "thin", color: { rgb: "000000" } },
          },
        };
      }
    }
  }
}

function getCellFontSize({ row, headerRowIndex, profile, titleRowCount }) {
  if (row === 0 && titleRowCount) return profile === "timePrint" ? 20 : 14;
  if (row > 0 && row < titleRowCount) return 9;
  if (row === headerRowIndex) return profile === "timePrint" ? 13 : 10;
  if (profile === "timePrint") return 14;
  if (profile === "classPrint" || profile === "roomPrint") return 9;
  return 10;
}

function getBodyRowHeight(profile, rowCount) {
  if (profile === "timePrint") return rowCount <= 8 ? 42 : 30;
  if (profile === "classPrint") return rowCount > 45 ? 14 : 15.5;
  if (profile === "roomPrint") return rowCount > 42 ? 13.5 : 15.5;
  if (profile === "overview") return 18;
  if (profile === "validation") return 18;
  if (profile === "print") return 16;
  return 18;
}

function getRowHeight({ index, titleRowCount, headerRowIndex, profile, rowHeight }) {
  if (index === 0 && titleRowCount) return profile === "timePrint" ? 42 : 28;
  if (index > 0 && index < titleRowCount) return 24;
  if (index === headerRowIndex) return profile === "timePrint" ? 28 : 20;
  return rowHeight;
}

function safeSheetName(name) {
  return safeString(name).replace(/[\\/?*[\]:]/g, "").slice(0, 31) || "Sheet";
}

export function saveExamRecord(record) {
  const records = loadExamRecords();
  const versionBase = `${record.examName || "考试"}_${record.examDate || ""}`;
  const version = records.filter((item) => item.versionBase === versionBase).length + 1;
  const fullRecord = {
    ...sanitizeRecord(record),
    id: crypto.randomUUID(),
    versionBase,
    version,
    createdAt: new Date().toISOString(),
  };
  localStorage.setItem("exam-room-records", JSON.stringify([fullRecord, ...records]));
  return fullRecord;
}

export function loadExamRecords() {
  try {
    return JSON.parse(localStorage.getItem("exam-room-records") || "[]");
  } catch {
    return [];
  }
}

export function deleteExamRecord(id) {
  localStorage.setItem("exam-room-records", JSON.stringify(loadExamRecords().filter((record) => record.id !== id)));
}

function sanitizeRecord(record) {
  return {
    ...record,
    physics: stripPool(record.physics),
    history: stripPool(record.history),
    schedule: {
      ...record.schedule,
      mainAssignments: stripAssignments(record.schedule?.mainAssignments || []),
      foreignAssignments: stripAssignments(record.schedule?.foreignAssignments || []),
      electiveAssignments: stripAssignments(record.schedule?.electiveAssignments || []),
      subjectAssignments: stripAssignments(record.schedule?.subjectAssignments || []),
    },
  };
}

function stripPool(pool = {}) {
  return {
    ...pool,
    students: (pool.students || []).map(({ original, ...student }) => student),
  };
}

function stripAssignments(assignments) {
  return assignments.map(({ original, ...assignment }) => assignment);
}
