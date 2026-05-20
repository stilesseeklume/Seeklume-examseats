import {
  buildPrintRows,
  buildDoorRows,
  buildRoomSummaryRows,
  buildRoomPrintRows,
  buildSchedule,
  buildSubjectPrintRows,
  buildValidationReport,
  ELECTIVE_SUBJECTS,
  SCHEDULE_MODES,
  summarizeValidationReport,
} from "../src/scheduler.js";
import { parseRoomImportSource, parseStudentImportSource } from "../src/importers.js";

const physicsText = [
  "班级\t姓名\t考号\t首选科目\t选科组合\t总分市排名\t总分\t数学\t英语\t化学\t地理\t政治\t生物",
  "1班\t物化生01\tP001\t物理\t物理 化学 生物\t1\t690\t145\t138\t96\t\t\t94",
  "1班\t物化地02\tP002\t物理\t物理 化学 地理\t2\t680\t140\t132\t93\t91\t\t",
  "2班\t物政地03\tP003\t物理\t物理 政治 地理\t3\t670\t136\t128\t\t89\t92\t",
  "2班\t物生政04\tP004\t物理\t物理 生物 政治\t4\t660\t134\t126\t\t\t88\t90",
].join("\n");

const historyText = [
  "班级\t姓名\t考号\t首选科目\t选科组合\t总分市排名\t总分\t数学\t英语\t化学\t地理\t政治\t生物\t日语",
  "3班\t历化生01\tH001\t历史\t历史 化学 生物\t5\t650\t130\t\t90\t\t\t89\t126",
  "3班\t历政地02\tH002\t历史\t历史 政治 地理\t6\t640\t128\t\t\t86\t91\t\t124",
  "4班\t历化政03\tH003\t历史\t历史 化学 政治\t7\t630\t126\t120\t87\t\t88\t",
  "4班\t历生地04\tH004\t历史\t历史 生物 地理\t8\t620\t124\t118\t\t84\t\t86",
].join("\n");

const roomText = [
  "考场号,门牌号,教室,容量",
  "1,201,高三01班,2",
  "2,202,高三02班,2",
  "3,203,高三03班,2",
  "4,204,高三04班,2",
  "5,205,高三05班,2",
  "6,206,高三06班,2",
  "7,207,高三07班,2",
  "8,208,高三08班,2",
].join("\n");

const physics = await parseStudentImportSource(physicsText, "物理");
const history = await parseStudentImportSource(historyText, "历史");
const rooms = await parseRoomImportSource(roomText);

assertEqual(physics.errors.length, 0, `物理导入错误：${physics.errors.join(";")}`);
assertEqual(history.errors.length, 0, `历史导入错误：${history.errors.join(";")}`);
assertEqual(rooms.length, 8, "考场导入数量");

const comboSchedule = buildSchedule({
  physicsStudents: physics.students,
  historyStudents: history.students,
  rooms,
  minorLanguageRooms: {},
  examTimes: [],
  mode: SCHEDULE_MODES.TWO_DAY_COMBO,
});

const splitSchedule = buildSchedule({
  physicsStudents: physics.students,
  historyStudents: history.students,
  rooms,
  minorLanguageRooms: {},
  examTimes: [],
  mode: SCHEDULE_MODES.THREE_DAY_SPLIT,
});

const manualJapaneseSchedule = buildSchedule({
  physicsStudents: physics.students,
  historyStudents: history.students,
  rooms,
  minorLanguageRooms: { 日语: { roomNos: "6" } },
  examTimes: [],
  mode: SCHEDULE_MODES.THREE_DAY_SPLIT,
});

const manualJapaneseDoorSchedule = buildSchedule({
  physicsStudents: physics.students,
  historyStudents: history.students,
  rooms,
  minorLanguageRooms: { 日语: { roomNos: "206" } },
  examTimes: [],
  mode: SCHEDULE_MODES.THREE_DAY_SPLIT,
});

const insufficientJapaneseSchedule = buildSchedule({
  physicsStudents: physics.students,
  historyStudents: history.students,
  rooms: rooms.map((room) => (room.roomNo === "1" ? { ...room, capacity: 1 } : room)),
  minorLanguageRooms: { 日语: { roomNos: "1" } },
  examTimes: [],
  mode: SCHEDULE_MODES.THREE_DAY_SPLIT,
});

assertEqual(comboSchedule.errors.length, 0, `两天组合生成错误：${comboSchedule.errors.join(";")}`);
assertEqual(splitSchedule.errors.length, 0, `三天拆分生成错误：${splitSchedule.errors.join(";")}`);
assertEqual(manualJapaneseSchedule.errors.length, 0, `手动日语考场生成错误：${manualJapaneseSchedule.errors.join(";")}`);
assertEqual(manualJapaneseDoorSchedule.errors.length, 0, `手动日语门牌生成错误：${manualJapaneseDoorSchedule.errors.join(";")}`);
assertNoMixedPhysicsHistory(comboSchedule.mainAssignments);
assertTwoDayComboSummary(comboSchedule);

for (const subject of ELECTIVE_SUBJECTS) {
  const rows = splitSchedule.subjectAssignments.filter((item) => item.subjectLabel === subject);
  assertEqual(rows.length, 8, `${subject}总安排人数`);
  assertNoMixedExamSelf(rows, subject);
}

assertTwoExamTwoSelf(splitSchedule.subjectAssignments, [...physics.students, ...history.students]);
assertForeignLanguagesDoNotMixRooms(splitSchedule.foreignAssignments);
assertLanguageUsesOnlyRooms(manualJapaneseSchedule.foreignAssignments, "日语", ["6"]);
assertLanguageUsesOnlyRooms(manualJapaneseDoorSchedule.foreignAssignments, "日语", ["6"]);
assertManualForeignRoomBlocksWhenInsufficient(insufficientJapaneseSchedule, "日语");
assertPrintRowsHideScores(buildPrintRows(splitSchedule.allRows), "班主任表");
assertPrintRowsHideScores(buildSubjectPrintRows(splitSchedule, "化学"), "科目表");
assertPrintRowsHideScores(buildRoomPrintRows(splitSchedule), "考场信息表");
assertRoomPrintRowsShowForeignLanguage(buildRoomPrintRows(splitSchedule));
assertSelfStudyRoomsAreNumeric(buildPrintRows(splitSchedule.allRows));
assertForeignDoorSummary(buildDoorRows(splitSchedule, rooms));

const validationReport = buildValidationReport({ schedule: splitSchedule, rooms, importErrors: [] });
const validationSummary = summarizeValidationReport(validationReport);
assertEqual(validationSummary.blockers, 0, "公开样例阻断错误数");
console.log("规则验证通过");
console.log({
  physics: physics.students.length,
  history: history.students.length,
  rooms: rooms.length,
  mainRooms: Object.fromEntries(countBy(comboSchedule.mainAssignments, "roomNo")),
  splitSubjects: Object.fromEntries(ELECTIVE_SUBJECTS.map((subject) => [subject, Object.fromEntries(countBy(splitSchedule.subjectAssignments.filter((item) => item.subjectLabel === subject), "status"))])),
});

function countBy(rows, key) {
  const counter = new Map();
  for (const row of rows) {
    counter.set(row[key], (counter.get(row[key]) || 0) + 1);
  }
  return counter;
}

function assertEqual(actual, expected, label) {
  if (actual !== expected) {
    throw new Error(`${label}：期望 ${expected}，实际 ${actual}`);
  }
}

function assertNoMixedPhysicsHistory(assignments) {
  const byRoom = new Map();
  for (const item of assignments) {
    if (!byRoom.has(item.roomNo)) byRoom.set(item.roomNo, new Set());
    byRoom.get(item.roomNo).add(item.firstSubject);
  }
  for (const [roomNo, subjects] of byRoom.entries()) {
    if (subjects.has("物理") && subjects.has("历史")) {
      throw new Error(`主考第${roomNo}考场混入物理和历史`);
    }
  }
}

function assertNoMixedExamSelf(assignments, subject) {
  const byRoom = new Map();
  for (const item of assignments) {
    if (!byRoom.has(item.roomNo)) byRoom.set(item.roomNo, new Set());
    byRoom.get(item.roomNo).add(item.status);
  }
  for (const [roomNo, statuses] of byRoom.entries()) {
    if (statuses.has("考试") && statuses.has("自习")) {
      throw new Error(`${subject}第${roomNo}考场混入考试和自习`);
    }
  }
}

function assertTwoExamTwoSelf(assignments, students) {
  for (const student of students) {
    const rows = assignments.filter((item) => item.studentId === student.id);
    const examCount = rows.filter((item) => item.status === "考试").length;
    const selfCount = rows.filter((item) => item.status === "自习").length;
    assertEqual(examCount, 2, `${student.name} 考试科目数`);
    assertEqual(selfCount, 2, `${student.name} 自习科目数`);
  }
}

function assertPrintRowsHideScores(rows, label) {
  const forbidden = ["总分", "总分市排名", "该科分数", "化学分数", "地理分数", "政治分数", "生物分数", "状态"];
  const headers = Object.keys(rows[0] || {});
  for (const field of forbidden) {
    if (headers.includes(field)) {
      throw new Error(`${label} 不应包含敏感列：${field}`);
    }
  }
}

function assertRoomPrintRowsShowForeignLanguage(rows) {
  const japaneseRows = rows.filter((row) => row.科目 === "日语" || row.当科 === "日语");
  if (!japaneseRows.length) {
    throw new Error("考场信息表里的外语考场应显示具体语种，例如日语，而不是笼统写外语");
  }
}

function assertSelfStudyRoomsAreNumeric(rows) {
  const selfStudyCells = rows
    .flatMap((row) => [row.化学, row.地理, row.政治, row.生物])
    .filter((value) => String(value).includes("自习"));
  if (selfStudyCells.length) {
    throw new Error("教师打印表自习考场应只显示数字，不应保留“自习室”字样");
  }
}

function assertForeignDoorSummary(rows) {
  const foreignCells = rows.map((row) => String(row.外语 || ""));
  if (!foreignCells.some((cell) => cell.includes("英语"))) {
    throw new Error("门牌人数总览应显示英语人数明细");
  }
  if (!foreignCells.some((cell) => cell.includes("日语"))) {
    throw new Error("门牌人数总览应显示小语种人数明细");
  }
}

function assertForeignLanguagesDoNotMixRooms(assignments) {
  const byRoom = new Map();
  for (const item of assignments) {
    if (!byRoom.has(item.roomNo)) byRoom.set(item.roomNo, new Set());
    byRoom.get(item.roomNo).add(item.subjectLabel);
  }
  for (const [roomNo, languages] of byRoom.entries()) {
    if (languages.size > 1) {
      throw new Error(`外语第${roomNo}考场混入多个语种：${[...languages].join("、")}`);
    }
  }
}

function assertLanguageUsesOnlyRooms(assignments, language, expectedRooms) {
  const actualRooms = [...new Set(assignments.filter((item) => item.subjectLabel === language).map((item) => String(item.roomNo)))];
  const expected = expectedRooms.map(String).sort().join(",");
  const actual = actualRooms.sort().join(",");
  if (actual !== expected) {
    throw new Error(`${language}应严格使用手动指定考场 ${expected}，实际 ${actual || "无"}`);
  }
}

function assertManualForeignRoomBlocksWhenInsufficient(schedule, language) {
  if (!schedule.errors.some((message) => message.includes(`${language}外语考场容量不足`))) {
    throw new Error(`${language}手动指定考场容量不足时应阻断生成，实际错误：${schedule.errors.join(";")}`);
  }
}

function assertTwoDayComboSummary(schedule) {
  const summaryRows = buildRoomSummaryRows(schedule, rooms);
  const mixed = summaryRows.find((row) => String(row.四选二内容 || "").includes("+"));
  if (!mixed) {
    throw new Error("两天组合模式下应能在考场汇总/门牌说明中显示混合组合人数");
  }
  if (!/化学|地理|政治|生物/.test(String(mixed.四选二内容))) {
    throw new Error(`四选二混合组合说明不清楚：${mixed.四选二内容}`);
  }
}
