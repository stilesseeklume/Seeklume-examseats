import React, { useEffect, useMemo, useState } from "react";
import { createRoot } from "react-dom/client";
import JSZip from "jszip";
import {
  AlertTriangle,
  ArrowLeftRight,
  ArrowRight,
  BookOpenCheck,
  CalendarDays,
  Clock3,
  FilePlus2,
  Filter,
  History,
  Maximize2,
  Minimize2,
  House,
  Search,
  CheckCircle2,
  ClipboardList,
  Download,
  FileDown,
  FileSpreadsheet,
  Globe2,
  Image,
  Plus,
  School,
  Trash2,
  Upload,
  UsersRound,
  Zap,
} from "lucide-react";
import {
  buildRoomDetailRows,
  buildAdminRows,
  buildDoorRows,
  buildPrintRows,
  buildRoomPrintRows,
  buildRoomSummaryRows,
  buildSubjectPrintRows,
  buildValidationReport,
  buildWorkbookFile,
  buildSchedule,
  buildExportJobs,
  buildExportFileName,
  DEFAULT_COMBO_ORDER,
  defaultRooms,
  deleteExamRecord,
  ELECTIVE_SUBJECTS,
  EXPORT_GROUPS,
  exportWorkbook,
  getElectiveSubjects,
  LANGUAGE_SUBJECTS,
  loadExamRecords,
  saveExamRecord,
  SCHEDULE_MODES,
  summarizeValidationReport,
} from "./scheduler.js";
import { parseRoomImportSource, parseStudentImportSource } from "./importers.js";
import { HomeLanding } from "./landing/HomeLanding.jsx";
import { AboutModal } from "./landing/AboutModal.jsx";
import "./styles.css";

const TEMPLATE_TIMES = [
  { subject: "语文", dayOffset: 0, start: "09:00", end: "11:30" },
  { subject: "数学", dayOffset: 0, start: "15:00", end: "17:00" },
  { subject: "物理/历史", dayOffset: 1, start: "09:00", end: "10:15" },
  { subject: "外语", dayOffset: 1, start: "15:00", end: "17:00" },
  { subject: "化学", dayOffset: 2, start: "08:30", end: "09:45" },
  { subject: "地理", dayOffset: 2, start: "11:00", end: "12:15" },
  { subject: "政治", dayOffset: 2, start: "14:30", end: "15:45" },
  { subject: "生物", dayOffset: 2, start: "17:00", end: "18:15" },
];

const LOCAL_STORAGE_KEYS = [
  "physics-import",
  "history-import",
  "exam-rooms",
  "minor-language-rooms",
  "exam-times",
  "exam-date",
  "schedule-mode",
  "export-mode",
  "exam-room-records",
  "exam-workspace-draft",
];

function shouldOpenHomeGuide() {
  return new URLSearchParams(window.location.search).get("view") === "home";
}

function App() {
  const workspaceDraft = loadWorkspaceDraft();
  const [examName, setExamName] = useState(workspaceDraft?.examName || "高三大型考试");
  const [examDate, setExamDate] = useState(workspaceDraft?.examDate || loadExamDate());
  const [physics, setPhysics] = useState(() => workspaceDraft?.physics ? deserializePool(workspaceDraft.physics) : loadImportedPool("physics-import"));
  const [history, setHistory] = useState(() => workspaceDraft?.history ? deserializePool(workspaceDraft.history) : loadImportedPool("history-import"));
  const [rooms, setRooms] = useState(workspaceDraft?.rooms || loadRooms());
  const [minorRooms, setMinorRooms] = useState(workspaceDraft?.minorRooms || loadMinorRooms());
  const [scheduleMode, setScheduleMode] = useState(workspaceDraft?.scheduleMode || loadScheduleMode());
  const [examTimes, setExamTimes] = useState(workspaceDraft?.examTimes || loadExamTimes(workspaceDraft?.examDate || loadExamDate()));
  const [selected, setSelected] = useState({
    grade: true,
    classes: true,
    subjectSheets: false,
    roomSummary: false,
    roomSheets: true,
    doorSummary: true,
    timeSheet: true,
    ...(workspaceDraft?.selected || {}),
  });
  const [exportMode, setExportMode] = useState(workspaceDraft?.exportMode || loadExportMode());
  const [records, setRecords] = useState(loadExamRecords());
  const [savedMessage, setSavedMessage] = useState("");
  const [previewKey, setPreviewKey] = useState(workspaceDraft?.previewKey || "print");
  const [activeStep, setActiveStep] = useState(workspaceDraft?.activeStep ?? 0);
  const [studentQuery, setStudentQuery] = useState("");
  const [view, setView] = useState(shouldOpenHomeGuide() ? "home" : workspaceDraft ? "workspace" : "home");
  const [recordFilter, setRecordFilter] = useState("");
  const [showAbout, setShowAbout] = useState(false);

  const importErrors = [
    ...physics.errors.map((error) => `物理类：${error}`),
    ...history.errors.map((error) => `历史类：${error}`),
  ];
  const schedule = useMemo(
    () =>
      buildSchedule({
        physicsStudents: physics.students,
        historyStudents: history.students,
        rooms,
        minorLanguageRooms: minorRooms,
        examTimes,
        mode: scheduleMode,
      }),
    [physics.students, history.students, rooms, minorRooms, examTimes, scheduleMode],
  );
  const allErrors = [...importErrors, ...schedule.errors];
  const validationReport = useMemo(() => buildValidationReport({ schedule, rooms, importErrors }), [schedule, rooms, importErrors]);
  const validationSummary = useMemo(() => summarizeValidationReport(validationReport), [validationReport]);
  const exportSummaryItems = useMemo(() => {
    const items = ["校验报告（单独给自己）"];
    if (selected.doorSummary) items.push("门牌人数总览");
    if (selected.grade) items.push("管理总表");
    if (selected.classes) items.push("班主任表");
    if (selected.roomSheets) items.push("考场信息表");
    if (selected.timeSheet) items.push("考试时间表");
    if (selected.subjectSheets && schedule.mode === SCHEDULE_MODES.THREE_DAY_SPLIT) items.push("科目总表");
    return items;
  }, [selected, schedule.mode]);
  const studentSearchRows = useMemo(() => searchStudents(schedule.allRows, studentQuery), [schedule.allRows, studentQuery]);
  const minorLanguages = useMemo(() => {
    const langs = new Set([...physics.students, ...history.students].map((student) => student.language).filter((language) => language !== "英语"));
    return LANGUAGE_SUBJECTS.filter((language) => langs.has(language));
  }, [physics.students, history.students]);
  const minorRoomRecommendation = useMemo(
    () => buildMinorRoomRecommendation({ students: [...physics.students, ...history.students], rooms, minorLanguages }),
    [physics.students, history.students, rooms, minorLanguages],
  );
  const foreignLanguages = useMemo(() => {
    const langs = new Set([...physics.students, ...history.students].map((student) => student.language || "英语"));
    if (physics.students.length || history.students.length) langs.add("英语");
    return ["英语", ...LANGUAGE_SUBJECTS].filter((language) => langs.has(language));
  }, [physics.students, history.students]);
  const previewTabs = useMemo(() => buildPreviewTabs(schedule, rooms, scheduleMode, validationReport), [schedule, rooms, scheduleMode, validationReport]);
  const totalStudents = physics.students.length + history.students.length;
  const enabledRooms = rooms.filter((room) => room.enabled).length;
  const hasStarted = totalStudents > 0 || rooms.length > 0 || Object.keys(minorRooms).length > 0;
  const visibleErrors = hasStarted ? allErrors : [];
  const routedIssues = useMemo(() => visibleErrors.map(createIssueRoute), [visibleErrors]);
  const activeStepIssues = useMemo(() => routedIssues.filter((issue) => issue.stepIndex === activeStep), [routedIssues, activeStep]);
  const configuredMinorRooms = minorLanguages.filter((language) => minorRooms[language]?.roomNos || minorRooms[language]?.roomNo).length;
  const languageStats = useMemo(() => buildLanguageStats([...physics.students, ...history.students], minorRooms), [physics.students, history.students, minorRooms]);
  const roomStats = useMemo(() => buildRoomStats(rooms), [rooms]);
  const duplicateDoorNos = useMemo(() => findDuplicateDoorNos(rooms), [rooms]);
  const roomConflictErrors = useMemo(() => visibleErrors.filter((error) => error.includes("同时用于") || error.includes("门牌")), [visibleErrors]);
  const conflictDoorNos = useMemo(() => extractDoorNosFromErrors(roomConflictErrors), [roomConflictErrors]);
  const roomIssueDoorNos = useMemo(() => new Set([...duplicateDoorNos, ...conflictDoorNos]), [duplicateDoorNos, conflictDoorNos]);
  const progressValue = Math.round(((activeStep + 1) / 6) * 100);
  const filteredRecords = useMemo(() => {
    const term = recordFilter.trim().toLowerCase();
    if (!term) return records;
    return records.filter((record) => [record.examName, record.examDate, record.version].some((value) => String(value || "").toLowerCase().includes(term)));
  }, [records, recordFilter]);

  useEffect(() => {
    if (view !== "workspace") return;
    const timer = setTimeout(() => {
      saveWorkspaceDraft({
        examName,
        examDate,
        physics,
        history,
        rooms,
        minorRooms,
        scheduleMode,
        examTimes,
        selected,
        exportMode,
        activeStep,
        previewKey,
      });
    }, 400);
    return () => clearTimeout(timer);
  }, [view, examName, examDate, physics, history, rooms, minorRooms, scheduleMode, examTimes, selected, exportMode, activeStep, previewKey]);

  const handleStudentFile = async (file, pool) => {
    const parsed = await parseStudentImportSource(file, pool);
    if (pool === "物理") setPhysicsAndStore(parsed);
    if (pool === "历史") setHistoryAndStore(parsed);
  };

  const setPhysicsAndStore = (next) => {
    setPhysics((prev) => {
      const resolved = typeof next === "function" ? next(prev) : next;
      saveImportedPool("physics-import", resolved);
      return resolved;
    });
  };

  const setHistoryAndStore = (next) => {
    setHistory((prev) => {
      const resolved = typeof next === "function" ? next(prev) : next;
      saveImportedPool("history-import", resolved);
      return resolved;
    });
  };

  const updateStudent = (pool, index, patch) => {
    const setter = pool === "物理" ? setPhysicsAndStore : setHistoryAndStore;
    setter((prev) => ({
      ...prev,
      students: prev.students.map((student, rowIndex) => (rowIndex === index ? patchStudent(student, patch, pool) : student)),
    }));
  };

  const moveStudent = (fromPool, index, toPool) => {
    const sourceSetter = fromPool === "物理" ? setPhysicsAndStore : setHistoryAndStore;
    const targetSetter = toPool === "物理" ? setPhysicsAndStore : setHistoryAndStore;
    sourceSetter((prev) => {
      const moving = prev.students[index];
      if (!moving) return prev;
      targetSetter((targetPrev) => ({ ...targetPrev, students: [...targetPrev.students, patchStudent({ ...moving, pool: toPool }, { firstSubject: toPool }, toPool)] }));
      return { ...prev, students: prev.students.filter((_, rowIndex) => rowIndex !== index) };
    });
  };

  const handleRoomFile = async (file) => {
    const parsedRooms = await parseRoomImportSource(file);
    setRoomsAndStore(parsedRooms.length ? parsedRooms : rooms);
  };

  const setRoomsAndStore = (nextRooms) => {
    setRooms(nextRooms);
    localStorage.setItem("exam-rooms", JSON.stringify(nextRooms));
  };

  const updateRoom = (index, patch) => {
    const nextRooms = rooms.map((room, roomIndex) => (roomIndex === index ? { ...room, ...patch } : room));
    setRoomsAndStore(nextRooms);
  };

  const addRoom = () => {
    const enabledRooms = rooms.filter((room) => room.enabled && room.roomNo);
    const numericRoomNos = enabledRooms.map((room) => Number(room.roomNo)).filter(Number.isFinite);
    const nextRoomNo = numericRoomNos.length ? Math.max(...numericRoomNos) + 1 : rooms.length + 1;
    setRoomsAndStore([
      ...rooms,
      {
        roomNo: String(nextRoomNo),
        doorNo: "",
        roomName: "",
        capacity: 40,
        enabled: true,
      },
    ]);
  };

  const removeRoom = (index) => {
    setRoomsAndStore(rooms.filter((_, roomIndex) => roomIndex !== index));
  };

  const updateMinorRoom = (language, patch) => {
    const next = { ...minorRooms, [language]: { ...minorRooms[language], ...patch } };
    setMinorRooms(next);
    localStorage.setItem("minor-language-rooms", JSON.stringify(next));
  };

  const commitMinorRoom = (language, field, value) => {
    updateMinorRoom(language, { [field]: value });
  };

  const applyMinorRoomRecommendation = () => {
    const next = { ...minorRooms };
    for (const item of minorRoomRecommendation.items) {
      if (!item.rooms.length) continue;
      next[item.language] = {
        ...next[item.language],
        roomNos: item.rooms.map((room) => room.roomNo).join(","),
      };
    }
    setMinorRooms(next);
    localStorage.setItem("minor-language-rooms", JSON.stringify(next));
  };

  const updateExamTime = (index, patch) => {
    const next = examTimes.map((item, itemIndex) => (itemIndex === index ? { ...item, ...patch } : item));
    setExamTimes(next);
    localStorage.setItem("exam-times", JSON.stringify(next));
  };

  const applyTimeTemplate = (firstDate = examDate, mode = scheduleMode) => {
    const next = mode === SCHEDULE_MODES.TWO_DAY_COMBO ? buildTwoDayTimes(firstDate) : buildThreeDayTimes(firstDate);
    setExamDate(firstDate);
    setExamTimes(next);
    localStorage.setItem("exam-date", firstDate);
    localStorage.setItem("exam-times", JSON.stringify(next));
  };

  const applyParsedExamTimes = (parsedTimes) => {
    const next = mergeParsedExamTimes(examTimes, parsedTimes);
    const firstDate = getEarliestExamDate(next) || examDate;
    setExamDate(firstDate);
    setExamTimes(next);
    localStorage.setItem("exam-date", firstDate);
    localStorage.setItem("exam-times", JSON.stringify(next));
  };

  const updateScheduleMode = (mode) => {
    setScheduleMode(mode);
    localStorage.setItem("schedule-mode", mode);
    applyTimeTemplate(examDate, mode);
  };

  const saveRecord = () => {
    if (allErrors.length) return;
    const record = saveExamRecord({
      examName,
      examDate,
      physics,
      history,
      rooms,
      minorRooms,
      examTimes,
      selected,
      schedule,
    });
    setRecords(loadExamRecords());
    setSavedMessage(`已保存：${record.examName} 第 ${record.version} 版`);
    setTimeout(() => setSavedMessage(""), 2600);
  };

  const removeRecord = (id) => {
    deleteExamRecord(id);
    setRecords(loadExamRecords());
  };

  const clearLocalData = () => {
    const confirmed = window.confirm("清空后，本机浏览器里的学生名单、考场配置、考试时间和历史记录都会删除。已导出的 Excel 不受影响。确定清空吗？");
    if (!confirmed) return;
    LOCAL_STORAGE_KEYS.forEach((key) => localStorage.removeItem(key));
    const today = todayLocalDateString();
    setExamName("高三大型考试");
    setExamDate(today);
    setPhysics({ students: [], errors: [], fieldMap: {} });
    setHistory({ students: [], errors: [], fieldMap: {} });
    setRooms(defaultRooms());
    setMinorRooms({});
    setExamTimes(buildThreeDayTimes(today));
    setScheduleMode(SCHEDULE_MODES.THREE_DAY_SPLIT);
    setExportMode(EXPORT_GROUPS.ALL);
    setSelected({ grade: true, classes: true, subjectSheets: false, roomSummary: false, roomSheets: true, doorSummary: true, timeSheet: true });
    setRecords([]);
    setRecordFilter("");
    setActiveStep(0);
    setSavedMessage("已清空本机数据");
    setTimeout(() => setSavedMessage(""), 2400);
  };

  const exportNow = async () => {
    if (allErrors.length) return;
    const jobs = buildExportJobs({ examName, examDate, schedule, rooms, selected, exportMode });
    const printSettingsBySheet = loadPreviewPrintSettings();
    if (exportMode === "package") {
      await exportZipPackage({ examName, examDate, schedule, rooms, selected, jobs, printSettingsBySheet });
      return;
    }
    for (const job of jobs) {
      const { fileName, data } = await buildWorkbookFile({ examName, examDate, schedule, rooms, selected, printSettingsBySheet, ...job });
      downloadBlob(new Blob([data], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), fileName);
    }
  };

  const exportRecord = async (record) => {
    const recordSelected = { grade: true, classes: true, subjectSheets: false, roomSummary: false, roomSheets: true, doorSummary: true, timeSheet: true, ...(record.selected || {}) };
    const jobs = buildExportJobs({
      examName: record.examName,
      examDate: record.examDate,
      schedule: record.schedule,
      rooms: record.rooms || defaultRooms(),
      selected: recordSelected,
      exportMode,
    });
    if (exportMode === "package") {
      await exportZipPackage({
        examName: record.examName,
        examDate: record.examDate,
        schedule: record.schedule,
        rooms: record.rooms || defaultRooms(),
        selected: recordSelected,
        jobs,
      });
      return;
    }
    for (const job of jobs) {
      const { fileName, data } = await buildWorkbookFile({
        examName: record.examName,
        examDate: record.examDate,
        schedule: record.schedule,
        rooms: record.rooms || defaultRooms(),
        selected: recordSelected,
        ...job,
      });
      downloadBlob(new Blob([data], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), fileName);
    }
  };

  const updateExportMode = (mode) => {
    setExportMode(mode);
    localStorage.setItem("export-mode", mode);
  };

  const newBlankExam = () => {
    const today = todayLocalDateString();
    const blankPhysics = { students: [], errors: [], fieldMap: {} };
    const blankHistory = { students: [], errors: [], fieldMap: {} };
    const nextRooms = [];
    const nextTimes = buildThreeDayTimes(today);
    setExamName("高三大型考试");
    setExamDate(today);
    setPhysicsAndStore(blankPhysics);
    setHistoryAndStore(blankHistory);
    setRoomsAndStore(nextRooms);
    setMinorRooms({});
    setExamTimes(nextTimes);
    setScheduleMode(SCHEDULE_MODES.THREE_DAY_SPLIT);
    setSelected({ grade: true, classes: true, subjectSheets: false, roomSummary: false, roomSheets: true, doorSummary: true, timeSheet: true });
    setActiveStep(0);
    localStorage.setItem("minor-language-rooms", JSON.stringify({}));
    localStorage.setItem("exam-times", JSON.stringify(nextTimes));
    localStorage.setItem("schedule-mode", SCHEDULE_MODES.THREE_DAY_SPLIT);
    localStorage.removeItem("exam-workspace-draft");
    setSavedMessage("已新建空白考试");
    setTimeout(() => setSavedMessage(""), 2200);
    setView("workspace");
  };

  const loadRecordIntoWorkspace = (record) => {
    setExamName(record.examName || "高三大型考试");
    setExamDate(record.examDate || todayLocalDateString());
    if (record.physics?.students || record.history?.students) {
      setPhysicsAndStore(deserializePool(record.physics || { students: [], errors: [], fieldMap: {} }));
      setHistoryAndStore(deserializePool(record.history || { students: [], errors: [], fieldMap: {} }));
    }
    setRoomsAndStore(record.rooms || defaultRooms());
    setMinorRooms(record.minorRooms || {});
    setExamTimes(record.examTimes || (record.schedule?.mode === SCHEDULE_MODES.TWO_DAY_COMBO ? buildTwoDayTimes(record.examDate || todayLocalDateString()) : buildThreeDayTimes(record.examDate || todayLocalDateString())));
    setSelected({ grade: true, classes: true, subjectSheets: false, roomSummary: false, roomSheets: true, doorSummary: true, timeSheet: true, ...(record.selected || {}) });
    setScheduleMode(record.schedule?.mode || SCHEDULE_MODES.THREE_DAY_SPLIT);
    localStorage.setItem("minor-language-rooms", JSON.stringify(record.minorRooms || {}));
    localStorage.setItem("exam-times", JSON.stringify(record.examTimes || []));
    localStorage.setItem("schedule-mode", record.schedule?.mode || SCHEDULE_MODES.THREE_DAY_SPLIT);
    localStorage.removeItem("exam-workspace-draft");
    setActiveStep(4);
    setSavedMessage(record.physics?.students || record.history?.students ? `已载入：${record.examName} 第 ${record.version} 版` : "已载入旧记录设置；旧记录不含可编辑学生池");
    setTimeout(() => setSavedMessage(""), 2600);
    setView("workspace");
  };

  const openRecord = (record) => {
    loadRecordIntoWorkspace(record);
  };

  const goHome = () => {
    setView("home");
  };

  const goIssue = (issue) => {
    setActiveStep(issue.stepIndex);
    setView("workspace");
  };

  const loadRecordById = (id) => {
    const record = records.find((item) => item.id === id);
    if (record) loadRecordIntoWorkspace(record);
  };

  const stepItems = [
    {
      title: "导入成绩单",
      shortTitle: "导入",
      emoji: "⬆️",
      icon: <Upload size={18} />,
      status: physics.students.length && history.students.length && !importErrors.length ? "done" : importErrors.length ? "error" : "idle",
      note: "导入物理类、历史类成绩单。",
      content: (
        <>
          <div className="upload-grid">
            <FilePicker title="物理类成绩单" hint={`${physics.students.length} 人`} onFile={(file) => handleStudentFile(file, "物理")} />
            <FilePicker title="历史类成绩单" hint={`${history.students.length} 人`} onFile={(file) => handleStudentFile(file, "历史")} />
          </div>
          <ImportErrorPanel errors={importErrors} />
          <FieldMap title="物理类字段识别" map={physics.fieldMap} meta={physics.importMeta} />
          <FieldMap title="历史类字段识别" map={history.fieldMap} meta={history.importMeta} />
          <GradeRosterOverview physicsStudents={physics.students} historyStudents={history.students} />
          <StudentRoster
            title="物理类名单"
            pool="物理"
            students={physics.students}
            onAdd={() => setPhysicsAndStore((prev) => ({ ...prev, students: [...prev.students, createBlankStudent("物理")] }))}
            onChange={(index, patch) => updateStudent("物理", index, patch)}
            onDelete={(index) => setPhysicsAndStore((prev) => ({ ...prev, students: prev.students.filter((_, rowIndex) => rowIndex !== index) }))}
            onMove={(index) => moveStudent("物理", index, "历史")}
            showSubjectScores={scheduleMode === SCHEDULE_MODES.THREE_DAY_SPLIT}
          />
          <StudentRoster
            title="历史类名单"
            pool="历史"
            students={history.students}
            onAdd={() => setHistoryAndStore((prev) => ({ ...prev, students: [...prev.students, createBlankStudent("历史")] }))}
            onChange={(index, patch) => updateStudent("历史", index, patch)}
            onDelete={(index) => setHistoryAndStore((prev) => ({ ...prev, students: prev.students.filter((_, rowIndex) => rowIndex !== index) }))}
            onMove={(index) => moveStudent("历史", index, "物理")}
            showSubjectScores={scheduleMode === SCHEDULE_MODES.THREE_DAY_SPLIT}
          />
        </>
      ),
    },
    {
      title: "确认考场",
      shortTitle: "考场",
      emoji: "🏫",
      icon: <ClipboardList size={18} />,
      status: hasStarted && rooms.filter((room) => room.enabled).length ? "done" : "idle",
      note: "先准备本次可用教室；尾场不满40人是正常排座结果。",
      content: (
        <>
          <RoomFixPanel errors={roomConflictErrors} duplicateDoorNos={duplicateDoorNos} conflictDoorNos={conflictDoorNos} onGoMinor={() => setActiveStep(2)} />
          <div className="toolbar">
            <FilePicker compact title="导入考场模板" hint="读取考场号、门牌、教室、容量" onFile={handleRoomFile} />
            <button type="button" onClick={() => setRoomsAndStore(defaultRooms())}>生成22个40人考场</button>
            <button type="button" onClick={addRoom}>新增考场</button>
          </div>
          <p className="step-help">考场容量表示最多可坐人数。排座时最后一个考场可能不满员，这是正常的；容量不足时才需要新增考场或调整容量。</p>
          <div className="table-wrap room-table">
            <table>
              <thead>
                <tr>
                  <th>启用</th>
                  <th>考场号</th>
                  <th>门牌号</th>
                  <th>教室</th>
                  <th>容量</th>
                  <th>操作</th>
                </tr>
              </thead>
              <tbody>
                {rooms.map((room, index) => (
                  <tr key={`${room.roomNo}-${index}`} className={roomIssueDoorNos.has(String(room.doorNo || "").trim()) ? "room-duplicate-row" : ""}>
                    <td><input type="checkbox" checked={room.enabled} onChange={(event) => updateRoom(index, { enabled: event.target.checked })} /></td>
                    <td><input value={room.roomNo} onChange={(event) => updateRoom(index, { roomNo: event.target.value })} /></td>
                    <td><input value={room.doorNo} onChange={(event) => updateRoom(index, { doorNo: event.target.value })} /></td>
                    <td><input value={room.roomName} onChange={(event) => updateRoom(index, { roomName: event.target.value })} /></td>
                    <td><DraftNumberInput value={room.capacity} min={1} fallback={40} onChange={(capacity) => updateRoom(index, { capacity })} /></td>
                    <td><button type="button" onClick={() => removeRoom(index)}>删除</button></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </>
      ),
    },
    {
      title: "外语安排",
      shortTitle: "语种",
      emoji: "💬",
      icon: <FileSpreadsheet size={18} />,
      status: hasStarted ? "done" : "idle",
      note: "英语先排，其他语种接着排；语种之间不混考场。",
      content: (
        <>
          <p className="muted">外语按语种分组排座：英语先用普通考场，日语、俄语等接着用后续考场。手动填写后会锁定该语种考场，不会自动改到别处；容量不够会提示补考场。</p>
          <div className="minor-recommend-card">
            <div>
              <strong>外语考场建议</strong>
              <span>{minorRoomRecommendation.summary}</span>
            </div>
            <button type="button" onClick={applyMinorRoomRecommendation} disabled={!minorRoomRecommendation.items.some((item) => item.rooms.length)}>
              按建议锁定考场
            </button>
          </div>
          {minorRoomRecommendation.items.length > 0 && (
            <div className="minor-recommend-list">
              {minorRoomRecommendation.items.map((item) => (
                <span key={item.language}>
                  {item.language} {item.count} 人：{item.rooms.length ? item.rooms.map((room) => `${room.roomNo}考场`).join("、") : "暂无可推荐考场"}
                </span>
              ))}
            </div>
          )}
          {minorLanguages.length === 0 ? (
            <div className="empty">当前成绩单只有英语或尚未导入外语语种；英语会自动从第 1 考场开始排。</div>
          ) : (
            <div className="language-grid">
              {minorLanguages.map((language) => (
                <div className="language-row" key={language}>
                  <strong>{language}</strong>
                  <MinorRoomInput placeholder="考场号/门牌，可填 23 或 23,24" value={minorRooms[language]?.roomNos || minorRooms[language]?.roomNo || ""} onCommit={(value) => commitMinorRoom(language, "roomNos", value)} />
                  <span className="language-row-note">留空自动续排；填写后严格锁定，容量不够会报错。</span>
                </div>
              ))}
            </div>
          )}
        </>
      ),
    },
    {
      title: "考试时间",
      shortTitle: "时间",
      emoji: "🕘",
      icon: <History size={18} />,
      status: hasStarted && examTimes.length ? "done" : "idle",
      note: scheduleMode === SCHEDULE_MODES.TWO_DAY_COMBO ? "套用两天组合时间。" : "套用三天固定时间。",
      content: (
        <>
          <div className="time-template-panel">
            <label>
              <span>第一天日期</span>
              <input type="date" value={examDate} onChange={(event) => {
                applyTimeTemplate(event.target.value, scheduleMode);
              }} />
            </label>
            <div className="time-mode-switch" role="group" aria-label="考试天数">
              <button type="button" className={scheduleMode === SCHEDULE_MODES.THREE_DAY_SPLIT ? "active" : ""} onClick={() => updateScheduleMode(SCHEDULE_MODES.THREE_DAY_SPLIT)}>三天</button>
              <button type="button" className={scheduleMode === SCHEDULE_MODES.TWO_DAY_COMBO ? "active" : ""} onClick={() => updateScheduleMode(SCHEDULE_MODES.TWO_DAY_COMBO)}>两天</button>
            </div>
            <span className="time-template-hint">改日期或天数会自动铺开时间表。</span>
          </div>
          <TimeNoticeImporter examDate={examDate} examTimes={examTimes} onApply={applyParsedExamTimes} />
          <div className="table-wrap compact-table">
            <table>
              <thead>
                <tr>
                  <th>科目</th>
                  <th>日期</th>
                  <th>开始</th>
                  <th>结束</th>
                </tr>
              </thead>
              <tbody>
                {examTimes.map((item, index) => (
                  <tr key={item.subject}>
                    <td><input value={item.subject} onChange={(event) => updateExamTime(index, { subject: event.target.value })} /></td>
                    <td><input type="date" value={item.date} onChange={(event) => updateExamTime(index, { date: event.target.value })} /></td>
                    <td><input type="time" value={item.start} onChange={(event) => updateExamTime(index, { start: event.target.value })} /></td>
                    <td><input type="time" value={item.end} onChange={(event) => updateExamTime(index, { end: event.target.value })} /></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </>
      ),
    },
    {
      title: "全面预览",
      shortTitle: "预览",
      emoji: "🔎",
      icon: <AlertTriangle size={18} />,
      status: visibleErrors.length ? "error" : schedule.allRows.length ? "done" : "idle",
      note: "核对异常、总表和各科安排。",
      content: (
        <div className="preview-reader-layout">
          {visibleErrors.length > 0 && (
            <div className="error-box validation-error-box">
              <strong>暂不能导出</strong>
              <p>有 {visibleErrors.length} 条阻断错误。先点左侧“校验报告”查看原因和建议。</p>
            </div>
          )}
          <PreviewPanel tabs={previewTabs} activeKey={previewKey} onChange={setPreviewKey} />
          <details className="preview-support-panel">
            <summary>排考摘要与学生查询</summary>
            <div className="summary-grid compact">
              <Metric label="学生总数" value={schedule.allRows.length} />
              <Metric label="物理类" value={physics.students.length} />
              <Metric label="历史类" value={history.students.length} />
              <Metric label="普通考场" value={rooms.filter((room) => room.enabled).length} />
              <Metric label="排考规则" value={scheduleMode === SCHEDULE_MODES.THREE_DAY_SPLIT ? "三天拆分" : "两天组合"} />
              <ScheduleSummaryDetails lines={schedule.summary} />
              <StudentSearch query={studentQuery} onQuery={setStudentQuery} rows={studentSearchRows} />
            </div>
          </details>
        </div>
      ),
    },
    {
      title: "导出",
      shortTitle: "导出",
      emoji: "⬇️",
      icon: <Download size={18} />,
      status: "idle",
      note: "生成 Excel 材料。",
      content: (
        <>
          <div className="export-options export-options-list">
            <span className="check-option export-fixed-option"><input type="checkbox" checked readOnly /> 校验报告</span>
            <CheckOption label="门牌人数总览" checked={selected.doorSummary} onChange={(doorSummary) => setSelected({ ...selected, doorSummary })} />
            <CheckOption label="管理总表" checked={selected.grade} onChange={(grade) => setSelected({ ...selected, grade })} />
            <CheckOption label="班主任表" checked={selected.classes} onChange={(classes) => setSelected({ ...selected, classes })} />
            <CheckOption label="考场信息表" checked={selected.roomSheets} onChange={(roomSheets) => setSelected({ ...selected, roomSheets })} />
            <CheckOption label="考试时间表" checked={selected.timeSheet} onChange={(timeSheet) => setSelected({ ...selected, timeSheet })} />
            <CheckOption label="科目总表" checked={selected.subjectSheets} onChange={(subjectSheets) => setSelected({ ...selected, subjectSheets })} />
          </div>
          <div className="export-mode">
            <span>导出方式</span>
            <label><input type="radio" name="exportMode" checked={exportMode === EXPORT_GROUPS.ALL} onChange={() => updateExportMode(EXPORT_GROUPS.ALL)} /> 合成一个 Excel</label>
            <label><input type="radio" name="exportMode" checked={exportMode === "package"} onChange={() => updateExportMode("package")} /> 分别导出多个 Excel</label>
          </div>
          <div className="export-summary">
            <span>本次会生成</span>
            <div className="export-summary-chips">
              {exportSummaryItems.map((item) => <span key={item}>{item}</span>)}
            </div>
            <small>校验报告会单独导出给你复核；班主任表和考场信息表不夹带校验报告。</small>
          </div>
          <div className="actions">
            <button type="button" className="primary" disabled={allErrors.length > 0} onClick={exportNow}><Download size={17} /> 导出Excel</button>
            <button type="button" disabled={allErrors.length > 0} onClick={saveRecord}>保存到本机历史</button>
            {savedMessage && <span className="saved">{savedMessage}</span>}
          </div>
        </>
      ),
    },
  ];

  const activeStepItem = stepItems[activeStep];
  const isPreviewStep = activeStepItem.shortTitle === "预览";

  return view === "home" ? (
    <>
      <HomeLanding
        records={records}
        filteredRecords={filteredRecords}
        recordFilter={recordFilter}
        setRecordFilter={setRecordFilter}
        clearLocalData={clearLocalData}
        formatRecordTime={formatRecordTime}
        openRecord={openRecord}
        exportRecord={exportRecord}
        removeRecord={removeRecord}
        onStart={newBlankExam}
        onOpenAbout={() => setShowAbout(true)}
      />
      {showAbout && <AboutModal onClose={() => setShowAbout(false)} />}
    </>
  ) : (
    <main className="app-shell">
      <header className="topbar">
        <div className="topbar-brand">
          <a href="https://seeklume.work/" style={{ color: "inherit", textDecoration: "none" }}>
            <span className="brand-logo">Seeklume ExamSeats</span>
          </a>
        </div>
        <div className={`dynamic-island ${visibleErrors.length ? "warning" : "ready"}`} role="status">
          <span className="island-emoji" aria-hidden="true">{visibleErrors.length ? "⚠️" : activeStepItem.emoji}</span>
          <span className="island-title">{activeStepItem.shortTitle}</span>
          <span className="island-detail">{visibleErrors.length ? `${visibleErrors.length} 项待处理` : "就绪"}</span>
          <div className="island-popover" role="tooltip">
            {visibleErrors.length ? (
              <div className="issue-popover-list">
                <strong>阻断问题</strong>
                {routedIssues.slice(0, 4).map((issue) => (
                  <button type="button" className="issue-popover-row" key={issue.message} onClick={() => goIssue(issue)}>
                    <span>{issue.emoji}</span>
                    <b>{issue.stepTitle}</b>
                    <small>{issue.detail || issue.summary}</small>
                  </button>
                ))}
                {visibleErrors.length > 4 && <p>还有 {visibleErrors.length - 4} 条，请看当前步骤卡片。</p>}
              </div>
            ) : (
              <>
                <strong>{hasStarted ? "当前可生成" : "就绪"}</strong>
                <p>{hasStarted ? "硬校验通过。导出前可在预览页查看完整校验报告。" : "可以开始导入成绩单和考场。"}</p>
              </>
            )}
          </div>
        </div>
        <button type="button" className="privacy" onClick={() => setShowAbout(true)}><CheckCircle2 size={18} /> 仅本机保存 · 不上传云端</button>
      </header>

      <section className="exam-meta">
        <label>
          考试名称
          <input value={examName} onChange={(event) => setExamName(event.target.value)} />
        </label>
        <label>
          考试日期
          <input type="date" value={examDate} onChange={(event) => {
            setExamDate(event.target.value);
            applyTimeTemplate(event.target.value, scheduleMode);
          }} />
        </label>
      </section>

      <div className={`workbench-layout ${isPreviewStep ? "with-sidebar" : "single-panel"}`}>
        {isPreviewStep && (
          <aside className="workbench-sidebar">
            <div className="sidebar-exam">
              <span>预览导航</span>
              <strong>{examName || "未命名考试"}</strong>
              <small>{examDate || "未设置日期"} · 请选择要查看的表</small>
            </div>
            <nav className="sidebar-steps" aria-label="预览导航">
              {previewTabs.map((tabItem) => (
                <button
                  key={tabItem.key}
                  type="button"
                  className={tabItem.key === previewKey ? "active" : ""}
                  onClick={() => setPreviewKey(tabItem.key)}
                >
                  <b>{tabItem.label}</b>
                  <em>{tabItem.rows.length}</em>
                </button>
              ))}
            </nav>
            <div className="preview-sidebar-summary">
              <StatCard emoji="👥" icon={<UsersRound size={21} />} label="学生" value={`${totalStudents}`} hint="物理 + 历史" popup={<StudentStatPopup physics={physics.students.length} history={history.students.length} total={totalStudents} />} />
              <StatCard emoji="🏫" icon={<School size={21} />} label="考场" value={`${enabledRooms}`} hint={`配置 ${rooms.length}`} popup={<RoomStatPopup stats={roomStats} />} />
              <StatCard emoji="💬" icon={<Globe2 size={21} />} label="语种" value={foreignLanguages.length ? `${foreignLanguages.length}种` : "无"} hint={minorLanguages.length ? `非英语 ${configuredMinorRooms}/${minorLanguages.length}` : "仅英语"} popup={<LanguageStatPopup stats={languageStats} />} />
              <StatCard emoji={visibleErrors.length ? "⚠️" : "✅"} icon={visibleErrors.length ? <AlertTriangle size={21} /> : <CheckCircle2 size={21} />} label="校验" value={visibleErrors.length ? `${visibleErrors.length}` : hasStarted ? "通过" : "待开始"} hint={visibleErrors.length ? "待处理" : hasStarted ? "可生成" : "未导入"} tone={visibleErrors.length ? "danger" : "success"} popup={<ValidationStatPopup issues={routedIssues} summary={validationSummary} hasStarted={hasStarted} onIssueClick={goIssue} />} />
            </div>
          </aside>
        )}

        <section className="workspace-panel">
          <Step number={activeStep + 1} title={activeStepItem.title} emoji={activeStepItem.emoji} icon={activeStepItem.icon} status={activeStepItem.status} note={activeStepItem.note} issues={activeStepIssues} onIssueClick={goIssue}>
            {activeStepItem.content}
          </Step>
        </section>
      </div>

      <div className="floating-step-actions">
        <button type="button" disabled={activeStep === 0} onClick={() => setActiveStep(activeStep - 1)}>上一步</button>
        <button type="button" className="primary" disabled={activeStep === stepItems.length - 1} onClick={() => setActiveStep(activeStep + 1)}>下一步</button>
      </div>

      <nav className="dock" aria-label="排考步骤">
        <button
          type="button"
          className="dock-item home-dock-item"
          onClick={goHome}
          title="Home"
          data-label="返回首页"
        >
          <span className="dock-emoji" aria-hidden="true">🏠</span>
          <span className="dock-state" aria-label="返回首页" />
        </button>
        {stepItems.map((item, index) => {
          const stepIssueCount = routedIssues.filter((issue) => issue.stepIndex === index).length;
          return (
          <button
            key={item.title}
            type="button"
            className={`dock-item ${index === activeStep ? "active" : ""} ${item.status} ${stepIssueCount ? "has-issue" : ""}`}
            onClick={() => setActiveStep(index)}
            title={item.title}
            data-label={stepIssueCount ? `${dockTipLabel(item.title)} · ${stepIssueCount}个问题` : dockTipLabel(item.title)}
          >
            <span className="dock-emoji" aria-hidden="true">{item.emoji}</span>
            {stepIssueCount > 0 && <span className="dock-issue-badge">{stepIssueCount}</span>}
            <span className="dock-state" aria-label={getStatusLabel(item.status)} />
          </button>
          );
        })}
      </nav>
      {showAbout && <AboutModal onClose={() => setShowAbout(false)} />}
    </main>
  );
}

async function exportZipPackage({ examName, examDate, schedule, rooms, selected, jobs, printSettingsBySheet }) {
  const zip = new JSZip();
  for (const job of jobs) {
    const { fileName, data } = await buildWorkbookFile({ examName, examDate, schedule, rooms, selected, printSettingsBySheet, ...job });
    assertWorkbookDownloadable(fileName, data);
    zip.file(fileName, data);
  }
  const blob = await zip.generateAsync({ type: "blob" });
  downloadBlob(blob, buildExportFileName(examName, examDate, "分开材料").replace(/\.xlsx$/, ".zip"));
}

function assertWorkbookDownloadable(fileName, data) {
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
  const isZip = bytes[0] === 0x50 && bytes[1] === 0x4b && bytes[2] === 0x03 && bytes[3] === 0x04;
  if (!isZip || bytes.byteLength < 4000) {
    throw new Error(`${fileName} 生成异常，请先检查是否已经生成排考内容。`);
  }
}

function downloadBlob(blob, fileName) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 60_000);
}

function loadRooms() {
  try {
    return JSON.parse(localStorage.getItem("exam-rooms") || "null") || defaultRooms();
  } catch {
    return defaultRooms();
  }
}

function loadMinorRooms() {
  try {
    return JSON.parse(localStorage.getItem("minor-language-rooms") || "{}");
  } catch {
    return {};
  }
}

function loadExamTimes() {
  try {
    return JSON.parse(localStorage.getItem("exam-times") || "null") || buildThreeDayTimes(todayLocalDateString());
  } catch {
    return buildThreeDayTimes(todayLocalDateString());
  }
}

function loadExamDate() {
  return localStorage.getItem("exam-date") || todayLocalDateString();
}

function loadScheduleMode() {
  return localStorage.getItem("schedule-mode") || SCHEDULE_MODES.THREE_DAY_SPLIT;
}

function loadExportMode() {
  return localStorage.getItem("export-mode") || EXPORT_GROUPS.ALL;
}

function loadWorkspaceDraft() {
  try {
    const stored = JSON.parse(localStorage.getItem("exam-workspace-draft") || "null");
    if (!stored) return null;
    return {
      ...stored,
      physics: stored.physics ? deserializePool(stored.physics) : null,
      history: stored.history ? deserializePool(stored.history) : null,
      rooms: stored.rooms || null,
      minorRooms: stored.minorRooms || null,
      examTimes: stored.examTimes || null,
      selected: stored.selected || null,
    };
  } catch {
    return null;
  }
}

function saveWorkspaceDraft(draft) {
  try {
    localStorage.setItem("exam-workspace-draft", JSON.stringify({
      ...draft,
      physics: serializePoolForDraft(draft.physics),
      history: serializePoolForDraft(draft.history),
      rooms: draft.rooms,
      minorRooms: draft.minorRooms,
      examTimes: draft.examTimes,
      selected: draft.selected,
    }));
  } catch {
    // keep working in-memory even if storage is full
  }
}

function serializePoolForDraft(pool) {
  if (!pool) return null;
  return {
    ...pool,
    students: (pool.students || []).map(serializeStudent),
  };
}

function loadImportedPool(key) {
  try {
    const stored = JSON.parse(localStorage.getItem(key) || "null");
    if (!stored) return { students: [], errors: [], fieldMap: {} };
    return {
      ...stored,
      errors: stored.errors || [],
      fieldMap: stored.fieldMap || {},
      students: (stored.students || []).map(deserializeStudent),
    };
  } catch {
    return { students: [], errors: [], fieldMap: {} };
  }
}

function deserializePool(pool) {
  return {
    ...pool,
    errors: pool.errors || [],
    fieldMap: pool.fieldMap || {},
    students: (pool.students || []).map(deserializeStudent),
  };
}

function saveImportedPool(key, value) {
  try {
    localStorage.setItem(key, JSON.stringify({
      ...value,
      students: (value.students || []).map(serializeStudent),
    }));
  } catch {
    // localStorage can fail if the browser quota is full; keep the in-memory edit usable.
  }
}

function serializeStudent(student) {
  return {
    ...student,
    totalRank: encodeNumber(student.totalRank),
    totalScore: encodeNumber(student.totalScore),
    mathScore: encodeNumber(student.mathScore),
    foreignScore: encodeNumber(student.foreignScore),
    languageScore: encodeNumber(student.languageScore),
    subjectScores: Object.fromEntries(Object.entries(student.subjectScores || {}).map(([subject, value]) => [subject, encodeNumber(value)])),
  };
}

function deserializeStudent(student) {
  return {
    ...student,
    totalRank: decodeNumber(student.totalRank),
    totalScore: decodeNumber(student.totalScore),
    mathScore: decodeNumber(student.mathScore),
    foreignScore: decodeNumber(student.foreignScore),
    languageScore: decodeNumber(student.languageScore),
    subjectScores: Object.fromEntries(Object.entries(student.subjectScores || {}).map(([subject, value]) => [subject, decodeNumber(value)])),
  };
}

function encodeNumber(value) {
  if (Number.isNaN(value)) return "__NaN__";
  if (value === Number.POSITIVE_INFINITY) return "__Infinity__";
  if (value === Number.NEGATIVE_INFINITY) return "__-Infinity__";
  return value;
}

function decodeNumber(value) {
  if (value === "__NaN__") return Number.NaN;
  if (value === "__Infinity__") return Number.POSITIVE_INFINITY;
  if (value === "__-Infinity__") return Number.NEGATIVE_INFINITY;
  return value;
}

function buildThreeDayTimes(firstDate) {
  return TEMPLATE_TIMES.map((item) => {
    return {
      subject: item.subject,
      date: addDaysToDateString(firstDate, item.dayOffset),
      start: item.start,
      end: item.end,
    };
  });
}

function buildTwoDayTimes(firstDate) {
  return [
    { subject: "语文", dayOffset: 0, start: "09:00", end: "11:30" },
    { subject: "数学", dayOffset: 0, start: "15:00", end: "17:00" },
    { subject: "物理/历史", dayOffset: 1, start: "09:00", end: "10:15" },
    { subject: "外语", dayOffset: 1, start: "15:00", end: "17:00" },
    { subject: "四选二", dayOffset: 1, start: "17:20", end: "18:35" },
  ].map((item) => ({
    subject: item.subject,
    date: addDaysToDateString(firstDate, item.dayOffset),
    start: item.start,
    end: item.end,
  }));
}

function addDaysToDateString(dateString, days) {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(dateString || "");
  if (!match) return "";
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const date = new Date(year, month - 1, day + days);
  const yyyy = date.getFullYear();
  const mm = String(date.getMonth() + 1).padStart(2, "0");
  const dd = String(date.getDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
}

function todayLocalDateString() {
  const date = new Date();
  const yyyy = date.getFullYear();
  const mm = String(date.getMonth() + 1).padStart(2, "0");
  const dd = String(date.getDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
}

function TimeNoticeImporter({ examDate, examTimes, onApply }) {
  const [text, setText] = useState("");
  const [imageName, setImageName] = useState("");
  const [recognizing, setRecognizing] = useState(false);
  const [message, setMessage] = useState("可粘贴通知文字，或上传微信截图、图片通知。");
  const [preview, setPreview] = useState([]);
  const fileInputId = "exam-time-notice-input";

  const parseCurrentText = (rawText) => {
    const parsed = parseExamNoticeText(rawText, examDate);
    setPreview(parsed);
    if (!parsed.length) {
      setMessage("没有识别到可用的时间片段，试着把通知里的日期和“科目 时间”一起贴进来。");
      return;
    }
    setMessage(`识别到 ${parsed.length} 条时间信息，确认后可直接写入考试时间表。`);
  };

  const handleTextChange = (event) => {
    const next = event.target.value;
    setText(next);
    parseCurrentText(next);
  };

  const handleFile = async (file) => {
    if (!file) return;
    setImageName(file.name);
    setMessage("已收到图片，正在离线识别文字；第一次可能需要等几秒。");
    setRecognizing(true);
    try {
      const extracted = await extractTextFromImage(file);
      const merged = [text, extracted].filter(Boolean).join("\n");
      if (!merged.trim()) {
        setMessage("没有提取到清晰文字。你可以先把截图里的通知文字复制出来再贴这里。");
        setPreview([]);
        return;
      }
      setText(merged);
      parseCurrentText(merged);
    } catch (error) {
      setMessage(`暂时无法直接识别这张图：${error?.message || "请把截图里的文字复制后贴入"}`);
    } finally {
      setRecognizing(false);
    }
  };

  return (
    <details className="notice-importer" open>
      <summary>
        <span>导入通知时间</span>
        <small>收到微信截图或文字通知时再打开</small>
      </summary>
      <div className="notice-import-grid">
        <label className="notice-input-card">
          <div className="notice-input-head">
            <Image size={16} />
            <strong>通知文字 / OCR 结果</strong>
          </div>
          <textarea value={text} placeholder="例如：5月30日 上午9:00-11:30 语文；下午15:00-17:00 数学 ..." onChange={handleTextChange} />
        </label>
        <div className="notice-side-card">
          <label className="notice-upload">
            <input id={fileInputId} type="file" accept="image/*" onChange={(event) => event.target.files?.[0] && handleFile(event.target.files[0])} />
            <span className="notice-upload-icon" aria-hidden="true">🖼️</span>
            <strong>{recognizing ? "正在识别图片" : "上传通知截图"}</strong>
            <small>{imageName || "微信截图、照片、通知图片都可以"}</small>
          </label>
          <div className="notice-message">{message}</div>
          {preview.length > 0 && (
            <>
              <div className="notice-preview-head">
                <span>识别预览</span>
                <button type="button" className="primary" onClick={() => onApply(preview)}>写入考试时间</button>
              </div>
              <div className="notice-preview-list">
                {preview.map((item) => (
                  <span key={`${item.subject}-${item.date}-${item.start}-${item.end}`}>
                    {item.subject} · {item.date} · {item.start}-{item.end}
                  </span>
                ))}
              </div>
            </>
          )}
        </div>
      </div>
      <p className="notice-tip">支持“5月30日 语文 9:00-11:30”这类文字，也支持通知里分行写法。图片识别如果不清楚，会提示你直接贴文字。</p>
    </details>
  );
}

async function extractTextFromImage(file) {
  if (typeof window === "undefined") throw new Error("当前环境不可用");
  const textRecognizer = window.OCR?.recognize || window.Tesseract?.recognize;
  if (typeof textRecognizer === "function") {
    const result = await textRecognizer(file, "chi_sim+eng");
    return String(result?.data?.text || result?.text || "");
  }
  const tesseract = await import("tesseract.js");
  const result = await tesseract.recognize(file, "chi_sim+eng");
  return String(result?.data?.text || result?.text || "");
}

function parseExamNoticeText(rawText, examDate) {
  const text = String(rawText || "").replace(/\r/g, "\n");
  if (!text.trim()) return [];
  const normalized = text
    .replace(/[：:]/g, ":")
    .replace(/[—–~～]/g, "-")
    .replace(/[至到]/g, "-")
    .replace(/[，,;]/g, "；");
  const baseDate = getNoticeDate(normalized) || examDate || todayLocalDateString();
  const sentences = normalized
    .split(/[\n。！？!?.；;]/)
    .map((line) => line.trim())
    .filter(Boolean);
  const items = [];
  let currentDate = baseDate;
  for (const line of sentences) {
    const lineDate = getNoticeDate(line);
    if (lineDate) currentDate = lineDate;
    const subjectMatches = extractNoticeSubjectTimes(line);
    for (const match of subjectMatches) {
      if (!match.subject || !match.start || !match.end) continue;
      items.push({
        subject: match.subject,
        date: currentDate,
        start: match.start,
        end: match.end,
      });
    }
  }
  return dedupeExamTimes(items);
}

function extractNoticeSubjectTimes(line) {
  const result = [];
  const subjectPattern = "语文|数学|物理\\/历史|物理|历史|外语|英语|化学|地理|政治|生物|日语|俄语|法语|西班牙语|德语";
  const subjectFirst = new RegExp(`(${subjectPattern})[^0-9]{0,12}([0-2]?\\d:[0-5]\\d)\\s*-\\s*([0-2]?\\d:[0-5]\\d)`, "g");
  const timeFirst = new RegExp(`([0-2]?\\d:[0-5]\\d)\\s*-\\s*([0-2]?\\d:[0-5]\\d)[^\\u4e00-\\u9fa5]{0,8}(${subjectPattern})`, "g");
  let match;
  while ((match = subjectFirst.exec(line))) {
    result.push({
      subject: normalizeNoticeSubject(match[1]),
      start: normalizeNoticeTime(match[2]),
      end: normalizeNoticeTime(match[3]),
    });
  }
  while ((match = timeFirst.exec(line))) {
    const before = line.slice(Math.max(0, match.index - 12), match.index);
    if (new RegExp(`(${subjectPattern})`).test(before)) continue;
    result.push({
      subject: normalizeNoticeSubject(match[3]),
      start: normalizeNoticeTime(match[1]),
      end: normalizeNoticeTime(match[2]),
    });
  }
  return result;
}

function normalizeNoticeSubject(subject) {
  const text = String(subject || "").trim();
  if (!text) return "";
  if (text.includes("语文")) return "语文";
  if (text.includes("数学")) return "数学";
  if (text.includes("物理") && text.includes("历史")) return "物理/历史";
  if (text === "物理") return "物理/历史";
  if (text === "历史") return "物理/历史";
  if (text.includes("英语") || text.includes("外语")) return "外语";
  if (text.includes("化学")) return "化学";
  if (text.includes("地理")) return "地理";
  if (text.includes("政治")) return "政治";
  if (text.includes("生物")) return "生物";
  if (text.includes("日语")) return "日语";
  if (text.includes("俄语")) return "俄语";
  if (text.includes("法语")) return "法语";
  if (text.includes("西班牙语")) return "西班牙语";
  if (text.includes("德语")) return "德语";
  return text;
}

function normalizeNoticeTime(time) {
  const normalized = String(time || "").trim().replace(/[点时]/g, ":").replace(/::+/g, ":").replace(/^(\d)$/g, "0$1:00").replace(/^(\d{1,2})$/, "$1:00").replace(/^(\d{1,2}):(\d)$/g, "$1:0$2");
  const match = /^(\d{1,2}):(\d{2})$/.exec(normalized);
  if (!match) return normalized;
  return `${String(match[1]).padStart(2, "0")}:${match[2]}`;
}

function getNoticeDate(text) {
  const normalized = String(text || "");
  const fullDate = /(\d{4})年(\d{1,2})月(\d{1,2})日/.exec(normalized);
  if (fullDate) return `${fullDate[1]}-${String(fullDate[2]).padStart(2, "0")}-${String(fullDate[3]).padStart(2, "0")}`;
  const monthDay = /(\d{1,2})月(\d{1,2})日/.exec(normalized);
  if (!monthDay) return "";
  const year = new Date().getFullYear();
  return `${year}-${String(monthDay[1]).padStart(2, "0")}-${String(monthDay[2]).padStart(2, "0")}`;
}

function dedupeExamTimes(items) {
  const seen = new Set();
  const normalized = [];
  for (const item of items) {
    const key = [item.subject, item.date, item.start, item.end].join("|");
    if (seen.has(key)) continue;
    seen.add(key);
    normalized.push(item);
  }
  return normalized;
}

function mergeParsedExamTimes(existing, parsed) {
  const map = new Map();
  for (const item of existing || []) {
    const key = String(item.subject || "");
    map.set(key, { ...item });
  }
  for (const item of parsed || []) {
    const key = String(item.subject || "");
    map.set(key, { ...map.get(key), ...item });
  }
  return [...map.values()].filter((item) => item.subject || item.date || item.start || item.end);
}

function getEarliestExamDate(items) {
  const dates = (items || []).map((item) => String(item.date || "")).filter(Boolean).sort();
  return dates[0] || "";
}

function Step({ number, title, emoji, icon, status, note, issues = [], onIssueClick, children }) {
  return (
    <section className="step">
      <div className="step-title">
        <span className="step-number">{number}</span>
        <div className="step-copy">
          <div className="step-head">
            <span className="step-emoji" aria-hidden="true">{emoji}</span>
            <h2>{title}</h2>
            <span className={`step-badge ${status}`}>{getStatusLabel(status)}</span>
          </div>
          <p>{note}</p>
        </div>
      </div>
      <StepIssueGuide issues={issues} onIssueClick={onIssueClick} />
      {children}
    </section>
  );
}

function StepIssueGuide({ issues, onIssueClick }) {
  if (!issues.length) return null;
  return (
    <div className="step-issue-guide">
      <div>
        <strong>这一步有 {issues.length} 个问题</strong>
        <span>{issues[0]?.detail || "先处理这里，校验会自动刷新。"}</span>
      </div>
      <div className="step-issue-list">
        {issues.slice(0, 6).map((issue) => (
          <button type="button" key={issue.message} onClick={() => onIssueClick?.(issue)}>
            <span>{issue.emoji}</span>
            <b>{issue.summary}</b>
            <small>{issue.detail}</small>
          </button>
        ))}
        {issues.length > 6 && <span className="issue-more">还有 {issues.length - 6} 个</span>}
      </div>
    </div>
  );
}

function FilePicker({ title, hint, onFile, compact = false }) {
  const emoji = title.includes("物理") ? "⚛️" : title.includes("历史") ? "📜" : title.includes("门牌") ? "🪧" : "📄";
  return (
    <label className={compact ? "file-picker compact-picker" : "file-picker"}>
      <span className="file-emoji" aria-hidden="true">{emoji}</span>
      <span>{title}</span>
      {hint && <small>{hint}</small>}
      <input type="file" accept=".xls,.xlsx" onChange={(event) => event.target.files?.[0] && onFile(event.target.files[0])} />
    </label>
  );
}

function ImportErrorPanel({ errors }) {
  if (!errors.length) return null;
  return (
    <div className="error-box import-error-panel">
      <strong>导入需要处理</strong>
      <ul>
        {errors.slice(0, 8).map((error) => <li key={error}>{error}</li>)}
      </ul>
      {errors.length > 8 && <p>还有 {errors.length - 8} 条，请先修正导入文件或学生类别。</p>}
    </div>
  );
}

function FieldMap({ title, map, meta }) {
  const entries = Object.entries(map || {}).filter(([, value]) => value);
  if (!entries.length) return null;
  return (
    <details className="field-map">
      <summary>{title}</summary>
      {meta && (
        <p className="field-map-meta">
          自动读取：{meta.sheetName || "第一个工作表"} · 表头第 {meta.headerRowNumber || 1} 行
        </p>
      )}
      <div>{entries.map(([key, value]) => <span key={key}>{key} → {value}</span>)}</div>
    </details>
  );
}

function MinorRoomInput({ placeholder, value, onCommit }) {
  const [draft, setDraft] = useState(value || "");
  useEffect(() => setDraft(value || ""), [value]);
  const commit = () => {
    if (draft !== (value || "")) onCommit(draft);
  };
  return (
    <input
      placeholder={placeholder}
      value={draft}
      onChange={(event) => setDraft(event.target.value)}
      onBlur={commit}
      onKeyDown={(event) => {
        if (event.key === "Enter") event.currentTarget.blur();
      }}
    />
  );
}

function DraftNumberInput({ value, min = 1, fallback = 1, onChange }) {
  const [draft, setDraft] = useState(String(value ?? ""));
  useEffect(() => { setDraft(String(value ?? "")); }, [value]);
  const commit = () => {
    const num = Number(draft);
    if (Number.isFinite(num) && num >= min) {
      onChange(num);
    } else {
      setDraft(String(value ?? fallback));
    }
  };
  return (
    <input
      type="number"
      min={min}
      value={draft}
      onChange={(event) => setDraft(event.target.value)}
      onBlur={commit}
      onKeyDown={(event) => { if (event.key === "Enter") event.currentTarget.blur(); }}
    />
  );
}

function Metric({ label, value }) {
  return (
    <div className="metric">
      <span>{label}</span>
      <strong>{value}</strong>
    </div>
  );
}

function ScheduleSummaryDetails({ lines }) {
  if (!lines?.length) return null;
  return (
    <details className="summary-details">
      <summary>
        <span>排考生成摘要</span>
        <small>{lines.length} 条统计</small>
      </summary>
      <div className="summary-list">
        {lines.map((line) => <span key={line}>{line}</span>)}
      </div>
    </details>
  );
}

function StatCard({ emoji, icon, label, value, hint, tone = "neutral", popup }) {
  return (
    <article className={`stat-card ${tone}`}>
      <span className="stat-watermark" aria-hidden="true">{emoji}</span>
      <div className="stat-top">
        <span className="stat-icon" aria-hidden="true">{emoji}</span>
        <span className="stat-label">{label}</span>
      </div>
      <strong className="stat-value">{value}</strong>
      <span className="stat-hint">{hint}</span>
      {popup && <div className="stat-popup" role="tooltip">{popup}</div>}
    </article>
  );
}

function StudentStatPopup({ physics, history, total }) {
  return (
    <div className="popup-list">
      <strong>学生概览</strong>
      <span>物理类 <b>{physics}</b> 人</span>
      <span>历史类 <b>{history}</b> 人</span>
      <span>合计 <b>{total}</b> 人</span>
    </div>
  );
}

function RoomStatPopup({ stats }) {
  return (
    <div className="popup-list">
      <strong>考场概览</strong>
      <span>启用 <b>{stats.enabled}</b> 个</span>
      <span>总座位 <b>{stats.capacity}</b> 个</span>
      <span>最大容量 <b>{stats.maxCapacity}</b> 人/场</span>
    </div>
  );
}

function LanguageStatPopup({ stats }) {
  return (
    <div className="popup-list">
      <strong>外语语种</strong>
      {stats.map((item) => (
        <span key={item.language}>
          {item.language} <b>{item.count}</b> 人{item.language === "英语" ? "" : ` · ${item.roomNo ? `${item.roomNo}考场` : "未设考场"}`}
        </span>
      ))}
    </div>
  );
}

function ValidationStatPopup({ issues = [], summary, hasStarted, onIssueClick }) {
  return (
    <div className="popup-list">
      <strong>{issues.length ? "待处理问题" : hasStarted ? "校验通过" : "等待开始"}</strong>
      {!hasStarted && !issues.length ? <span>导入成绩单和考场后再显示校验结果。</span> : issues.length ? issues.slice(0, 4).map((issue) => (
        <button type="button" className="popup-route-button" key={issue.message} onClick={() => onIssueClick?.(issue)}>
          <span>{issue.emoji}</span>
          <b>{issue.stepTitle}</b>
          <small>{issue.detail || issue.summary}</small>
        </button>
      )) : <span>阻断 0 · 风险 {summary.warnings} · 复核 {summary.reviews}</span>}
      {issues.length > 4 && <span>还有 {issues.length - 4} 条，请看当前步骤卡片</span>}
    </div>
  );
}

function HeroVisual({ activeStep, progressValue, totalStudents, enabledRooms, allErrors, examDate }) {
  return (
    <aside className="hero-visual" aria-label="排考状态预览">
      <div className="visual-top">
        <span><CalendarDays size={15} /> {formatShortDate(examDate)}</span>
        <span className={allErrors ? "visual-alert" : "visual-ok"}>{allErrors ? `${allErrors} 项待处理` : "Ready"}</span>
      </div>
      <div className="visual-card main-visual-card">
        <div className="visual-card-head">
          <span>当前</span>
          <strong>{activeStep}</strong>
        </div>
        <div className="visual-progress">
          <span style={{ width: `${progressValue}%` }} />
        </div>
        <div className="seat-map" aria-hidden="true">
          {Array.from({ length: 48 }).map((_, index) => (
            <span key={index} className={index % 7 === 0 ? "seat accent" : index % 5 === 0 ? "seat cool" : "seat"} />
          ))}
        </div>
      </div>
      <div className="visual-mini-grid">
        <div className="visual-card mini">
          <UsersRound size={18} />
          <strong>{totalStudents}</strong>
          <span>Students</span>
        </div>
        <div className="visual-card mini">
          <School size={18} />
          <strong>{enabledRooms}</strong>
          <span>Rooms</span>
        </div>
      </div>
    </aside>
  );
}

function CheckOption({ label, checked, onChange }) {
  return (
    <label className="check-option">
      <input type="checkbox" checked={checked} onChange={(event) => onChange(event.target.checked)} />
      {label}
    </label>
  );
}

function StudentRoster({ title, pool, students, onAdd, onChange, onDelete, onMove, showSubjectScores }) {
  const [query, setQuery] = useState("");
  const [classFilter, setClassFilter] = useState("");
  const [comboFilter, setComboFilter] = useState("");
  const [languageFilter, setLanguageFilter] = useState("");
  const [expanded, setExpanded] = useState(false);
  const classOptions = useMemo(() => uniqueValues(students.map((student) => student.className)), [students]);
  const comboOptions = useMemo(() => buildComboOptions(students), [students]);
  const languageOptions = useMemo(() => uniqueValues(students.map((student) => student.language)), [students]);
  const visibleStudents = useMemo(
    () => filterRosterStudents(students, { query, classFilter, comboFilter, languageFilter }),
    [students, query, classFilter, comboFilter, languageFilter],
  );
  return (
    <section className={`roster-card ${expanded ? "expanded" : ""}`}>
      <div className="roster-header">
        <div>
          <h3>{title}</h3>
          <span>{students.length ? `${visibleStudents.length}/${students.length} 人` : "未导入"}</span>
        </div>
        <div className="roster-tools">
          <label className="roster-search">
            <Search size={14} />
            <input value={query} placeholder="搜姓名/考号/班级" onChange={(event) => setQuery(event.target.value)} />
          </label>
          <label className="roster-filter">
            <Filter size={14} />
            <select value={classFilter} onChange={(event) => setClassFilter(event.target.value)}>
              <option value="">全部班级</option>
              {classOptions.map((item) => <option key={item} value={item}>{item}</option>)}
            </select>
          </label>
          <label className="roster-filter">
            <Filter size={14} />
            <select value={comboFilter} onChange={(event) => setComboFilter(event.target.value)}>
              <option value="">全部组合</option>
              {comboOptions.map((item) => <option key={item.value} value={item.value}>{item.label}</option>)}
            </select>
          </label>
          <label className="roster-filter">
            <Filter size={14} />
            <select value={languageFilter} onChange={(event) => setLanguageFilter(event.target.value)}>
              <option value="">全部语种</option>
              {languageOptions.map((item) => <option key={item} value={item}>{item}</option>)}
            </select>
          </label>
          <button type="button" className="icon-button" onClick={() => setExpanded((value) => !value)} title={expanded ? "收起名单" : "展开名单"}>
            {expanded ? <Minimize2 size={16} /> : <Maximize2 size={16} />}
          </button>
          <button type="button" onClick={onAdd}><Plus size={16} /> 新增学生</button>
        </div>
      </div>
      {!students.length ? (
        <div className="roster-empty">
          <div className="roster-empty-icon" aria-hidden="true">{pool === "物理" ? "⚛️" : "📜"}</div>
          <strong>{pool}类名单等待导入</strong>
          <span>导入成绩单后，这里会显示可编辑学生名单。</span>
        </div>
      ) : (
      <div className={`table-wrap roster-table ${expanded ? "expanded" : ""}`}>
        <table>
          <thead>
            <tr>
              <th>姓名</th>
              <th>考号</th>
              <th>班级</th>
              <th>首选科目</th>
              <th>选科组合</th>
              <th>总分</th>
              <th>总分排名</th>
              <th>数学</th>
              <th>外语</th>
              {showSubjectScores && <th>化学</th>}
              {showSubjectScores && <th>地理</th>}
              {showSubjectScores && <th>政治</th>}
              {showSubjectScores && <th>生物</th>}
              <th>操作</th>
            </tr>
          </thead>
          <tbody>
            {visibleStudents.map(({ student, index }) => (
              <tr key={student.id || `${pool}-${index}`}>
                <td><input value={student.name || ""} onChange={(event) => onChange(index, { name: event.target.value })} /></td>
                <td><input value={student.id || ""} onChange={(event) => onChange(index, { id: event.target.value })} /></td>
                <td><input value={student.className || ""} onChange={(event) => onChange(index, { className: event.target.value })} /></td>
                <td><input value={student.firstSubject || pool} onChange={(event) => onChange(index, { firstSubject: event.target.value })} /></td>
                <td><input value={student.comboRaw || ""} onChange={(event) => onChange(index, { comboRaw: event.target.value })} /></td>
                <td><input type="number" value={student.totalScore ?? ""} onChange={(event) => onChange(index, { totalScore: Number(event.target.value) || 0 })} /></td>
                <td><input type="number" value={Number.isFinite(student.totalRank) ? student.totalRank : ""} onChange={(event) => onChange(index, { totalRank: Number(event.target.value) || Number.POSITIVE_INFINITY })} /></td>
                <td><input type="number" value={student.mathScore ?? ""} onChange={(event) => onChange(index, { mathScore: Number(event.target.value) || 0 })} /></td>
                <td><input type="number" value={student.foreignScore ?? ""} onChange={(event) => onChange(index, { foreignScore: Number(event.target.value) || 0 })} /></td>
                {showSubjectScores && <td><input type="number" value={student.subjectScores?.化学 ?? ""} onChange={(event) => onChange(index, { subjectScores: { ...student.subjectScores, 化学: Number(event.target.value) || 0 } })} /></td>}
                {showSubjectScores && <td><input type="number" value={student.subjectScores?.地理 ?? ""} onChange={(event) => onChange(index, { subjectScores: { ...student.subjectScores, 地理: Number(event.target.value) || 0 } })} /></td>}
                {showSubjectScores && <td><input type="number" value={student.subjectScores?.政治 ?? ""} onChange={(event) => onChange(index, { subjectScores: { ...student.subjectScores, 政治: Number(event.target.value) || 0 } })} /></td>}
                {showSubjectScores && <td><input type="number" value={student.subjectScores?.生物 ?? ""} onChange={(event) => onChange(index, { subjectScores: { ...student.subjectScores, 生物: Number(event.target.value) || 0 } })} /></td>}
                <td>
                  <div className="record-actions">
                    <button type="button" onClick={() => onMove(index)}><ArrowLeftRight size={16} /> 切换</button>
                    <button type="button" onClick={() => onDelete(index)}><Trash2 size={16} /> 删除</button>
                  </div>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      )}
    </section>
  );
}

function filterRosterStudents(students, filters) {
  const query = String(filters?.query || "").trim().toLowerCase();
  const classFilter = String(filters?.classFilter || "").trim();
  const comboFilter = String(filters?.comboFilter || "").trim();
  const languageFilter = String(filters?.languageFilter || "").trim();
  return students
    .map((student, index) => ({ student, index }))
    .filter(({ student }) => {
      const searchable = [student.name, student.id, student.className, student.comboRaw, student.language];
      const matchesQuery = !query || searchable.some((value) => String(value || "").toLowerCase().includes(query));
      const matchesClass = !classFilter || String(student.className || "") === classFilter;
      const matchesCombo = !comboFilter || comboKeyForStudent(student) === comboFilter;
      const matchesLanguage = !languageFilter || String(student.language || "") === languageFilter;
      return matchesQuery && matchesClass && matchesCombo && matchesLanguage;
    });
}

function GradeRosterOverview({ physicsStudents, historyStudents }) {
  const [query, setQuery] = useState("");
  const [poolFilter, setPoolFilter] = useState("");
  const [subjectFilter, setSubjectFilter] = useState("");
  const [comboFilter, setComboFilter] = useState("");
  const [classFilter, setClassFilter] = useState("");
  const [open, setOpen] = useState(false);
  const students = useMemo(
    () => [
      ...physicsStudents.map((student) => ({ ...student, sourcePool: "物理" })),
      ...historyStudents.map((student) => ({ ...student, sourcePool: "历史" })),
    ],
    [physicsStudents, historyStudents],
  );
  const classOptions = useMemo(() => uniqueValues(students.map((student) => student.className)), [students]);
  const comboOptions = useMemo(() => buildComboOptions(students), [students]);
  const visible = useMemo(() => {
    const term = query.trim().toLowerCase();
    return students.filter((student) => {
      const subjects = getElectiveSubjects(student);
      const matchesQuery = !term || [student.name, student.id, student.className, student.comboRaw, student.language].some((value) => String(value || "").toLowerCase().includes(term));
      const matchesPool = !poolFilter || student.sourcePool === poolFilter;
      const matchesSubject = !subjectFilter || subjects.includes(subjectFilter);
      const matchesCombo = !comboFilter || comboKeyForStudent(student) === comboFilter;
      const matchesClass = !classFilter || String(student.className || "") === classFilter;
      return matchesQuery && matchesPool && matchesSubject && matchesCombo && matchesClass;
    });
  }, [students, query, poolFilter, subjectFilter, comboFilter, classFilter]);

  if (!students.length) return null;
  const rows = visible;
  return (
    <>
      <section className="overview-launch-card">
        <div>
          <h3>全年级筛选总览</h3>
          <span>跨物理类/历史类核对单科、组合、班级。</span>
        </div>
        <button type="button" className="primary" onClick={() => setOpen(true)}><Maximize2 size={16} /> 打开总览</button>
      </section>
      {open && (
        <div className="fullscreen-overlay" role="dialog" aria-modal="true">
          <section className="fullscreen-card">
            <div className="fullscreen-head">
              <div>
                <h3>全年级筛选总览</h3>
                <span>{visible.length}/{students.length} 人</span>
              </div>
              <button type="button" onClick={() => setOpen(false)}><Minimize2 size={16} /> 收起</button>
            </div>
            <div className="roster-tools modal-tools">
              <label className="roster-search">
                <Search size={14} />
                <input value={query} placeholder="搜姓名/考号/班级" onChange={(event) => setQuery(event.target.value)} />
              </label>
              <label className="roster-filter">
                <Filter size={14} />
                <select value={poolFilter} onChange={(event) => setPoolFilter(event.target.value)}>
                  <option value="">物理+历史</option>
                  <option value="物理">物理类</option>
                  <option value="历史">历史类</option>
                </select>
              </label>
              <label className="roster-filter">
                <Filter size={14} />
                <select value={subjectFilter} onChange={(event) => setSubjectFilter(event.target.value)}>
                  <option value="">全部再选科</option>
                  {ELECTIVE_SUBJECTS.map((subject) => <option key={subject} value={subject}>{subject}</option>)}
                </select>
              </label>
              <label className="roster-filter">
                <Filter size={14} />
                <select value={comboFilter} onChange={(event) => setComboFilter(event.target.value)}>
                  <option value="">全部组合</option>
                  {comboOptions.map((item) => <option key={item.value} value={item.value}>{item.label}</option>)}
                </select>
              </label>
              <label className="roster-filter">
                <Filter size={14} />
                <select value={classFilter} onChange={(event) => setClassFilter(event.target.value)}>
                  <option value="">全部班级</option>
                  {classOptions.map((item) => <option key={item} value={item}>{item}</option>)}
                </select>
              </label>
            </div>
            <div className="table-wrap roster-table fullscreen-table">
              <table>
                <thead>
                  <tr>
                    <th>类别</th>
                    <th>班级</th>
                    <th>姓名</th>
                    <th>考号</th>
                    <th>规范组合</th>
                    <th>原始组合</th>
                    <th>语种</th>
                  </tr>
                </thead>
                <tbody>
                  {rows.map((student) => (
                    <tr key={`${student.sourcePool}-${student.id}`}>
                      <td>{student.sourcePool}</td>
                      <td>{student.className}</td>
                      <td>{student.name}</td>
                      <td>{student.id}</td>
                      <td>{comboLabelForStudent(student)}</td>
                      <td>{student.comboRaw}</td>
                      <td>{student.language}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </section>
        </div>
      )}
    </>
  );
}

function buildComboOptions(students) {
  const options = new Map();
  for (const student of students || []) {
    const value = comboKeyForStudent(student);
    if (!value) continue;
    options.set(value, comboLabelForStudent(student));
  }
  return [...options.entries()]
    .map(([value, label]) => ({ value, label }))
    .sort((a, b) => comboSortRank(a.value) - comboSortRank(b.value) || a.label.localeCompare(b.label, "zh-Hans-CN", { numeric: true }));
}

function comboKeyForStudent(student) {
  const first = normalizeFirstSubject(student.firstSubject || student.pool);
  const electives = getElectiveSubjects(student);
  if (!first && electives.length !== 2) return "";
  const pair = normalizeElectivePair(electives);
  return [first || "未识别", pair || String(student.comboRaw || "").trim()].filter(Boolean).join("|");
}

function comboLabelForStudent(student) {
  const key = comboKeyForStudent(student);
  if (!key) return String(student.comboRaw || "");
  const [first, pair] = key.split("|");
  return `${shortFirstSubject(first)}${shortElectivePair(pair)}`;
}

function normalizeFirstSubject(value) {
  const text = String(value || "");
  if (text.includes("物理")) return "物理";
  if (text.includes("历史")) return "历史";
  return "";
}

function normalizeElectivePair(subjects) {
  const key = [...subjects].sort((a, b) => ELECTIVE_SUBJECTS.indexOf(a) - ELECTIVE_SUBJECTS.indexOf(b)).join("");
  return DEFAULT_COMBO_ORDER.includes(key) ? key : key;
}

function shortFirstSubject(subject) {
  if (subject === "物理") return "物";
  if (subject === "历史") return "史";
  return subject;
}

function shortElectivePair(pair) {
  return String(pair || "")
    .replace("化学", "化")
    .replace("地理", "地")
    .replace("政治", "政")
    .replace("生物", "生");
}

function comboSortRank(value) {
  const [first, pair] = String(value || "").split("|");
  const firstRank = first === "物理" ? 0 : first === "历史" ? 1 : 2;
  const pairRank = DEFAULT_COMBO_ORDER.indexOf(pair);
  return firstRank * 100 + (pairRank >= 0 ? pairRank : 90);
}

function uniqueValues(values) {
  return Array.from(new Set((values || []).map((value) => String(value || "").trim()).filter(Boolean))).sort((a, b) =>
    a.localeCompare(b, "zh-Hans-CN", { numeric: true }),
  );
}

function buildLanguageStats(students, minorRooms) {
  const counts = new Map();
  for (const student of students) {
    const language = student.language || "英语";
    counts.set(language, (counts.get(language) || 0) + 1);
  }
  return ["英语", ...LANGUAGE_SUBJECTS]
    .filter((language) => counts.has(language))
    .map((language) => ({
      language,
      count: counts.get(language) || 0,
      roomNo: minorRooms[language]?.roomNos || minorRooms[language]?.roomNo || "",
      roomName: "",
    }));
}

function buildMinorRoomRecommendation({ students, rooms, minorLanguages }) {
  const enabledRooms = (rooms || [])
    .filter((room) => room.enabled && room.roomNo)
    .sort((a, b) => String(a.roomNo).localeCompare(String(b.roomNo), "zh-Hans-CN", { numeric: true }));
  const englishCount = students.filter((student) => (student.language || "英语") === "英语").length;
  let remaining = englishCount;
  let englishRoomCount = 0;
  for (const room of enabledRooms) {
    if (remaining <= 0) break;
    remaining -= Number(room.capacity) || 40;
    englishRoomCount += 1;
  }
  const englishEndRoom = englishRoomCount ? enabledRooms[englishRoomCount - 1]?.roomNo : "无";
  let cursor = Math.min(englishRoomCount, enabledRooms.length);
  const items = minorLanguages.map((language) => {
    const count = students.filter((student) => student.language === language).length;
    const roomsForLanguage = [];
    let remainingLanguage = count;
    while (cursor < enabledRooms.length && remainingLanguage > 0) {
      const room = enabledRooms[cursor];
      roomsForLanguage.push(room);
      remainingLanguage -= Number(room.capacity) || 40;
      cursor += 1;
    }
    return {
      language,
      rooms: remainingLanguage > 0 ? [] : roomsForLanguage,
      count,
    };
  });
  const summary = !students.length
    ? "导入成绩单后自动计算英语占用范围。"
    : remaining > 0
      ? `英语 ${englishCount} 人，当前普通考场不足，暂不能推荐后续语种考场。`
      : `英语 ${englishCount} 人，预计使用 1-${englishEndRoom} 考场；其他语种从后续考场连续安排。`;
  return { englishCount, englishRoomCount, englishEndRoom, items, summary };
}

function buildRoomStats(rooms) {
  const enabled = rooms.filter((room) => room.enabled);
  const capacities = enabled.map((room) => Number(room.capacity) || 0);
  return {
    enabled: enabled.length,
    capacity: capacities.reduce((sum, value) => sum + value, 0),
    maxCapacity: capacities.length ? Math.max(...capacities) : 0,
  };
}

function findDuplicateDoorNos(rooms) {
  const counts = new Map();
  for (const room of rooms || []) {
    if (room.enabled === false) continue;
    const doorNo = String(room.doorNo || "").trim();
    if (!doorNo) continue;
    counts.set(doorNo, (counts.get(doorNo) || 0) + 1);
  }
  return new Set([...counts.entries()].filter(([, count]) => count > 1).map(([doorNo]) => doorNo));
}

function extractDoorNosFromErrors(errors = []) {
  const doorNos = new Set();
  for (const error of errors) {
    for (const match of String(error).matchAll(/门牌\s*([^\s，,。；;]+)/g)) {
      if (match[1]) doorNos.add(match[1].trim());
    }
  }
  return doorNos;
}

function RoomFixPanel({ errors, duplicateDoorNos, conflictDoorNos, onGoMinor }) {
  const hasDuplicateDoors = duplicateDoorNos.size > 0;
  const hasConflictDoors = conflictDoorNos.size > 0;
  if (!errors.length && !hasDuplicateDoors) return null;
  return (
    <div className="error-box room-fix-panel">
      <strong>考场信息需要修正</strong>
      <p>这通常不是你手填错了，而是导入的门牌模板、外语安排和当前排考规则发生了冲突。</p>
      {hasDuplicateDoors && <p>普通考场里重复的门牌号：{[...duplicateDoorNos].join("、")}。下方对应行已高亮，直接改门牌号即可。</p>}
      {hasConflictDoors && <p>本次报错涉及门牌：{[...conflictDoorNos].join("、")}。如果是外语语种占用了同一时段教室，请去“外语安排”页调整。</p>}
      {errors.length > 0 && (
        <ul>
          {errors.slice(0, 4).map((error) => <li key={error}>{error}</li>)}
        </ul>
      )}
      <div className="actions">
        <button type="button" onClick={onGoMinor}>去改外语安排</button>
      </div>
    </div>
  );
}

function dockTipLabel(title) {
  if (title === "导入成绩单") return "导入成绩单";
  if (title === "确认考场") return "管理考场";
  if (title === "外语安排" || title === "小语种") return "设置外语语种考场";
  if (title === "考试时间") return "填写时间";
  if (title === "全面预览") return "查看预览";
  if (title === "导出") return "导出材料";
  return title;
}

function createBlankStudent(pool) {
  return {
    id: "",
    name: "",
    className: "",
    firstSubject: pool,
    pool,
    comboRaw: "",
    totalRank: Number.POSITIVE_INFINITY,
    totalScore: 0,
    mathScore: 0,
    foreignScore: 0,
    language: "英语",
    languageScore: 0,
    subjectScores: { 化学: Number.NaN, 地理: Number.NaN, 政治: Number.NaN, 生物: Number.NaN },
    original: {},
    rowNumber: 0,
  };
}

function patchStudent(student, patch, pool) {
  const next = {
    ...student,
    ...patch,
    pool,
    subjectScores: {
      化学: student.subjectScores?.化学 ?? Number.NaN,
      地理: student.subjectScores?.地理 ?? Number.NaN,
      政治: student.subjectScores?.政治 ?? Number.NaN,
      生物: student.subjectScores?.生物 ?? Number.NaN,
      ...(patch.subjectScores || {}),
    },
  };
  return next;
}

function PreviewPanel({ tabs, activeKey, onChange }) {
  const active = tabs.find((tab) => tab.key === activeKey) || tabs[0];
  const [query, setQuery] = useState("");
  const [filters, setFilters] = useState({});
  const [currentPageIndex, setCurrentPageIndex] = useState(0);
  const [pageSettings, setPageSettings] = useState({});
  const [expanded, setExpanded] = useState(false);
  useEffect(() => {
    setCurrentPageIndex(0);
    setQuery("");
    setFilters({});
  }, [activeKey]);

  if (!active) return null;

  const activeRows = active.rows || [];
  const allColumns = activeRows.length ? Object.keys(activeRows[0]) : [];
  const filterColumns = getPreviewFilterColumns(allColumns);
  const filteredRows = useMemo(() => filterTableRows(activeRows, query, filters), [activeRows, query, filters]);
  const displayColumns = getPreviewDisplayColumns(active.label, filteredRows.length ? Object.keys(filteredRows[0]) : allColumns);
  const previewMeta = getPreviewSheetMeta(active.label, displayColumns.length, filteredRows.length);

  const allPages = useMemo(
    () => buildAllPreviewPages(active, filteredRows, displayColumns, previewMeta, pageSettings),
    [active, filteredRows, displayColumns, previewMeta, pageSettings],
  );

  const totalPages = allPages.length;
  const safeIndex = Math.min(currentPageIndex, Math.max(0, totalPages - 1));
  const currentPage = allPages[safeIndex] || null;
  const currentSettings = currentPage?.settings || { orientation: "portrait", fontSize: 10, rowHeight: 18, zoom: 92 };
  const currentPrintKey = currentPage ? getPreviewPrintKey(active.key, currentPage.sheetKey) : "";

  const savePageSettings = (newSettings) => {
    if (!currentPage) return;
    const next = { ...pageSettings, [currentPrintKey]: newSettings };
    setPageSettings(next);
    localStorage.setItem("preview-print-settings", JSON.stringify(next));
  };

  const patchSettings = (patch) => savePageSettings({ ...currentSettings, ...patch });

  const applyToAll = () => {
    if (!currentPage) return;
    const next = { ...pageSettings };
    const seen = new Set();
    for (const page of allPages) {
      if (seen.has(page.sheetKey)) continue;
      seen.add(page.sheetKey);
      next[getPreviewPrintKey(active.key, page.sheetKey)] = currentSettings;
    }
    setPageSettings(next);
    localStorage.setItem("preview-print-settings", JSON.stringify(next));
  };

  const fitToOnePage = () => {
    if (!currentPage) return;
    savePageSettings(getOnePagePreviewSettings(active.label, displayColumns, currentPage.sheetRows, Boolean(currentPage.note)));
  };

  const fitA4 = () => {
    if (!currentPage) return;
    savePageSettings(getA4PreviewSettings(active.label, displayColumns, currentPage.sheetRows));
  };

  const resetToAuto = () => {
    if (!currentPage) return;
    const next = { ...pageSettings };
    delete next[currentPrintKey];
    setPageSettings(next);
    localStorage.setItem("preview-print-settings", JSON.stringify(next));
  };

  const navTo = (index) => setCurrentPageIndex(Math.max(0, Math.min(totalPages - 1, index)));
  const isPortrait = currentSettings.orientation === "portrait";
  const rowsPerPage = currentPage ? getRowsPerPreviewPage(active.label, currentSettings, Boolean(currentPage.note)) : 30;
  const fitStatus = currentPage ? getPageFitStatus(currentPage.totalChunks, currentPage.sheetRows.length, rowsPerPage) : null;

  return (
    <div className={`preview-panel ${expanded ? "excel-mode" : ""}`}>
      {/* 顶栏：切表 + 翻页 + 状态 + 查人 */}
      <div className="preview-topbar">
        <label className="table-filter">
          <select value={active.key} onChange={(e) => onChange(e.target.value)}>
            {tabs.map((t) => <option key={t.key} value={t.key}>{t.label}</option>)}
          </select>
        </label>

        {totalPages > 1 && (
          <div className="preview-page-nav">
            <button type="button" className="icon-button" onClick={() => navTo(safeIndex - 1)} disabled={safeIndex === 0}>◀</button>
            <span className="preview-page-counter">{safeIndex + 1}<em>/{totalPages}</em></span>
            <button type="button" className="icon-button" onClick={() => navTo(safeIndex + 1)} disabled={safeIndex >= totalPages - 1}>▶</button>
            {totalPages > 3 && (
              <label className="table-filter">
                <select value={safeIndex} onChange={(e) => navTo(Number(e.target.value))}>
                  {allPages.map((page, i) => (
                    <option key={page.key} value={i}>
                      {page.title}{page.totalChunks > 1 ? ` (${page.pageNumber}/${page.totalChunks})` : ""}
                    </option>
                  ))}
                </select>
              </label>
            )}
          </div>
        )}

        {fitStatus && (
          <span className={`preview-fit-badge preview-fit-${fitStatus.status}`}>{fitStatus.label}</span>
        )}

        <div className="preview-search-row">
          <label className="table-search">
            <Search size={14} />
            <input value={query} placeholder="查人/考场/班级" onChange={(e) => setQuery(e.target.value)} />
          </label>
          {filterColumns.map((col) => (
            <label key={col} className="table-filter">
              <Filter size={13} />
              <select value={filters[col] || ""} onChange={(e) => setFilters((f) => ({ ...f, [col]: e.target.value }))}>
                <option value="">{col}</option>
                {uniqueValues(activeRows.map((r) => r[col])).map((v) => <option key={v} value={v}>{v}</option>)}
              </select>
            </label>
          ))}
          {(query || Object.values(filters).some(Boolean)) && (
            <button type="button" className="link-button" onClick={() => { setQuery(""); setFilters({}); }}>清空</button>
          )}
        </div>

        <span className="muted preview-row-count">{filteredRows.length}/{activeRows.length} 行</span>

        <button type="button" className="icon-button" onClick={() => setExpanded((v) => !v)} title={expanded ? "退出全屏" : "全屏预览"}>
          {expanded ? <Minimize2 size={16} /> : <Maximize2 size={16} />}
        </button>
      </div>

      {/* 调纸工具栏 */}
      <div className="preview-print-toolbar">
        <div className="print-orient-group">
          <button type="button" className={isPortrait ? "active" : ""} onClick={() => patchSettings({ orientation: "portrait" })}>竖向</button>
          <button type="button" className={!isPortrait ? "active" : ""} onClick={() => patchSettings({ orientation: "landscape" })}>横向</button>
        </div>
        <PreviewStepper label="字号" value={currentSettings.fontSize} min={7} max={18}
          onDecrease={() => patchSettings({ fontSize: Math.max(7, currentSettings.fontSize - 1) })}
          onIncrease={() => patchSettings({ fontSize: Math.min(18, currentSettings.fontSize + 1) })}
          onChange={(v) => patchSettings({ fontSize: clampNumber(v, 7, 18, currentSettings.fontSize) })}
        />
        <PreviewStepper label="行高" value={currentSettings.rowHeight} min={10} max={60}
          onDecrease={() => patchSettings({ rowHeight: Math.max(10, currentSettings.rowHeight - 1) })}
          onIncrease={() => patchSettings({ rowHeight: Math.min(60, currentSettings.rowHeight + 1) })}
          onChange={(v) => patchSettings({ rowHeight: clampNumber(v, 10, 60, currentSettings.rowHeight) })}
        />
        <PreviewStepper label="缩放" value={currentSettings.zoom} min={55} max={120} suffix="%"
          onDecrease={() => patchSettings({ zoom: Math.max(55, currentSettings.zoom - 5) })}
          onIncrease={() => patchSettings({ zoom: Math.min(120, currentSettings.zoom + 5) })}
          onChange={(v) => patchSettings({ zoom: clampNumber(v, 55, 120, currentSettings.zoom) })}
        />
        <div className="print-action-group">
          <button type="button" onClick={fitToOnePage} className="soft-button">压成一页</button>
          <button type="button" onClick={fitA4}>适配 A4</button>
          {totalPages > 1 && (
            <button type="button" onClick={applyToAll} className="soft-button">应用到全部</button>
          )}
          <button type="button" onClick={resetToAuto} className="link-button">重置</button>
          <button type="button" className="link-button" onClick={() => window.print()}>打印</button>
        </div>
      </div>

      {/* 纸张预览 */}
      <div className={`excel-canvas ${currentSettings.orientation}`}>
        <div className="excel-ribbon">
          <span>{currentPage?.title || active.label}</span>
          <small>
            {isPortrait ? "竖向" : "横向"} · 字 {currentSettings.fontSize} · 行高 {currentSettings.rowHeight} · {currentSettings.zoom}%
            {fitStatus && <> · <span className={`ribbon-fit-${fitStatus.status}`}>{fitStatus.label}</span></>}
          </small>
        </div>
        <div className="excel-pages">
          {currentPage ? (
            <div
              className={`excel-page ${currentSettings.orientation}`}
              key={currentPage.key}
              style={{
                "--preview-font-size": `${currentSettings.fontSize}px`,
                "--preview-row-height": `${currentSettings.rowHeight}px`,
                "--preview-zoom": currentSettings.zoom / 100,
              }}
              onDoubleClick={() => setExpanded(true)}
            >
              <div className="preview-print-title">{currentPage.title}</div>
              {currentPage.note && <div className="preview-print-note">{currentPage.note}</div>}
              <div className={`table-wrap preview-table ${expanded ? "expanded" : ""}`}>
                <table>
                  <thead>
                    <tr>{displayColumns.map((col) => <th key={col}>{col}</th>)}</tr>
                  </thead>
                  <tbody>
                    {currentPage.rows.map((row, i) => (
                      <tr key={row.考号 ? `${row.考号}-${i}` : i}>
                        {displayColumns.map((col) => (
                          <td key={col} className={row.__selfStudy || row[`__selfStudy:${col}`] ? "preview-self-study" : ""}>{row[col]}</td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              <div className="preview-page-footer">
                {currentPage.groupRowCount > 0 && <span>共 {currentPage.groupRowCount} 人</span>}
                {currentPage.groupPageCount > 1
                  ? <span className="preview-page-warn">本组第 {currentPage.groupPageIndex} / {currentPage.groupPageCount} 页 · 建议双面打印</span>
                  : <span>1 张 A4 铺开</span>}
                <span>总第 {safeIndex + 1} / {totalPages} 页</span>
              </div>
            </div>
          ) : (
            <div className="excel-page portrait">
              <div className="preview-empty-sheet">暂无数据</div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

function getPaperPageOptions(pages, query, rowQuery = "", filters = {}) {
  const sheetTerm = String(query || "").trim().toLowerCase();
  const rowTerm = String(rowQuery || "").trim().toLowerCase();
  const activeFilters = Object.entries(filters).filter(([, value]) => String(value || "").trim());
  return pages.filter((page) => {
    const pageText = [page.title, page.note, page.key, page.sheetKey].some((value) => String(value || "").toLowerCase().includes(sheetTerm));
    const matchesSheet = !sheetTerm || pageText || page.rows.some((row) => Object.values(row).some((value) => String(value || "").toLowerCase().includes(sheetTerm)));
    const matchesRowQuery = !rowTerm || page.rows.some((row) => Object.values(row).some((value) => String(value || "").toLowerCase().includes(rowTerm)));
    const matchesFilters = activeFilters.every(([column, value]) => page.rows.some((row) => String(row[column] ?? "") === String(value)));
    return matchesSheet && matchesRowQuery && matchesFilters;
  });
}

function getPreviewGroupRows(label, rows, sheetKey) {
  if (!sheetKey) return rows;
  const group = groupRowsForPreviewPages(label, rows).find((item) => item.key === sheetKey);
  return group?.rows || rows;
}

function loadPreviewPrintSettings() {
  try {
    return JSON.parse(localStorage.getItem("preview-print-settings") || "{}");
  } catch {
    return {};
  }
}

function getPreviewPrintKey(tabKey, pageKey) {
  return `${tabKey}:${pageKey}`;
}

function getPagePrintSettings(tabKey, page, fallback, settings = {}) {
  return settings[getPreviewPrintKey(tabKey, page.sheetKey || page.key)] || fallback;
}

function samePrintSettings(a, b) {
  return a?.orientation === b?.orientation && a?.fontSize === b?.fontSize && a?.rowHeight === b?.rowHeight && a?.zoom === b?.zoom;
}

function isSplitSheetPreview(label) {
  const text = String(label || "");
  return !["校验报告", "考试时间", "门牌人数总览", "管理总表"].some((item) => text.includes(item));
}

function clampNumber(value, min, max, fallback) {
  const number = Number(value);
  const base = Number.isFinite(number) ? number : fallback;
  return Math.min(max, Math.max(min, base));
}

function PreviewStepper({ label, value, min, max, suffix = "", onDecrease, onIncrease, onChange }) {
  const handleClick = (event, action) => {
    event.preventDefault();
    event.stopPropagation();
    action();
  };
  return (
    <div className="preview-stepper" aria-label={label}>
      <span>{label}</span>
      <button type="button" onClick={(event) => handleClick(event, onDecrease)} disabled={value <= min} aria-label={`${label}减小`}>-</button>
      <input
        type="number"
        min={min}
        max={max}
        value={value}
        onChange={(event) => onChange(event.target.value)}
      />
      {suffix && <em>{suffix}</em>}
      <button type="button" onClick={(event) => handleClick(event, onIncrease)} disabled={value >= max} aria-label={`${label}增大`}>+</button>
    </div>
  );
}

function getPreviewDisplayColumns(label, columns) {
  return columns
    .filter((column) => !column.startsWith("__"))
    .filter((column) => !(String(label || "").includes("班主任") && column === "班级分组"))
    .filter((column) => !(String(label || "").includes("考场信息") && column === "科目"));
}

function getPreviewSheetMeta(label, columnCount, rowCount) {
  if (String(label || "").includes("班主任")) {
    return {
      title: "班主任考场安排",
      note: "说明：语数物/座位号、语数历/座位号、外语、化学、地理、政治、生物均为“考场/座位号”；黄色底色表示该科为自习安排。",
    };
  }
  if (String(label || "").includes("外语")) {
    return {
      title: "外语安排",
      note: "英语和其他语种分开看；当前页的方向只作用于这一种语种。",
    };
  }
  if (String(label || "").includes("考场信息")) {
    return {
      title: "考场信息表",
      note: `${rowCount} 行 · ${columnCount} 列 · 预览按 A4 自动估算，导出 Excel 会再次写入打印页边距和缩放。`,
    };
  }
  if (String(label || "").includes("门牌")) {
    return {
      title: "门牌人数总览",
      note: "黄色底色表示该科该考场为自习室；无单独“自习室”列，避免重复占位。",
    };
  }
  if (String(label || "").includes("考试时间")) {
    return {
      title: "考试时间表",
      note: "导出时按 A4 纵向铺开，适合单独打印张贴或发给老师。",
    };
  }
  return {
    title: label || "预览表",
    note: "",
  };
}

function buildPaperPreviewPages(label, rows, columns, settings, meta, settingsBySheet = {}, tabKey = "__none__") {
  const grouped = groupRowsForPreviewPages(label, rows);
  const pages = [];
  for (const group of grouped) {
    const groupDefaultSettings = getA4PreviewSettings(label, columns, group.rows);
    const groupSettings = getPagePrintSettings(tabKey, { sheetKey: group.key, key: group.key }, groupDefaultSettings || settings, settingsBySheet);
    const rowsPerPage = getRowsPerPreviewPage(label, groupSettings, Boolean(group.note || meta.note));
    const chunks = chunkRows(group.rows, rowsPerPage);
    chunks.forEach((chunk, chunkIndex) => {
      pages.push({
        key: `${group.key}-${chunkIndex}`,
        sheetKey: group.key,
        title: group.title || meta.title,
        note: group.note || meta.note,
        rows: chunk,
        sheetRows: group.rows,
        printSettings: groupSettings,
        pageNumber: pages.length + 1,
        columns,
      });
    });
  }
  return pages.length ? pages : [{ key: "empty", sheetKey: "empty", title: meta.title, note: meta.note, rows: [], sheetRows: [], printSettings: settings, pageNumber: 1, columns }];
}

function groupRowsForPreviewPages(label, rows) {
  const text = String(label || "");
  if (text.includes("语数物/历")) {
    return groupRowsBy(rows, "考场号", (group) => ({
      title: `语数物历-${group}考场`,
      note: "主考按考场拆纸，当前页设置只影响这一考场。",
    }));
  }
  if (text.includes("外语")) {
    return groupRowsBy(rows, "__paperGroup", (group) => {
      const [language = "外语", room = ""] = String(group).split("|");
      return {
        title: `${language}-${room}考场`,
        note: language === "英语" ? "英语按考场拆纸。" : `${language}单独安排；当前页只影响这一考场。`,
      };
    });
  }
  if (["化学", "地理", "政治", "生物"].some((subject) => text === subject)) {
    return groupRowsBy(rows, "__paperGroup", (group) => {
      const [subject = text, type = "", room = ""] = String(group).split("|");
      return {
        title: `${subject}-${room}${type === "自习" ? "自习" : "考场"}`,
        note: type === "自习" ? "这是一张自习安排纸，可单独调打印样式。" : "这是一张考试安排纸，可单独调打印样式。",
      };
    });
  }
  if (text.includes("班主任")) {
    return groupRowsBy(rows, "__pageGroup", (group) => ({
      title: `${group}考场安排`,
      note: "说明：语数物/座位号、语数历/座位号、外语、化学、地理、政治、生物均为“考场/座位号”；黄色底色表示该科为自习安排。",
    }));
  }
  if (text.includes("考场信息")) {
    return groupRowsBy(rows, "__pageGroup", (group) => ({
      title: `${group}考场信息表`,
      note: "",
    }));
  }
  return [{ key: "all", title: "", note: "", rows }];
}

function groupRowsBy(rows, field, metaFactory) {
  const groups = new Map();
  for (const row of rows) {
    const key = String(row[field] || row.__pageGroup || row.__paperGroup || "未分组");
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(row);
  }
  return [...groups.entries()].map(([key, groupRows]) => ({ key, rows: groupRows, ...metaFactory(key) }));
}

function getRowsPerPreviewPage(label, settings, hasNote) {
  const pageHeight = settings.orientation === "landscape" ? 794 : 1123;
  const verticalPadding = 60;
  const titleHeight = 34;
  const noteHeight = hasNote ? 26 : 0;
  const headerHeight = 24;
  const footerHeight = 24;
  const available = pageHeight - verticalPadding - titleHeight - noteHeight - headerHeight - footerHeight;
  const baseRows = Math.floor(available / Math.max(8, settings.rowHeight || 18));
  if (String(label || "").includes("考场信息")) return Math.max(22, baseRows);
  if (String(label || "").includes("班主任")) return Math.max(28, baseRows);
  return Math.max(12, baseRows);
}

function chunkRows(rows, size) {
  const chunks = [];
  for (let index = 0; index < rows.length; index += size) {
    chunks.push(rows.slice(index, index + size));
  }
  return chunks;
}

function buildAllPreviewPages(active, filteredRows, displayColumns, previewMeta, pageSettings) {
  const label = active?.label || "";
  const tabKey = active?.key || "__none__";
  const grouped = groupRowsForPreviewPages(label, filteredRows);
  const pages = [];
  for (const group of grouped) {
    const printKey = getPreviewPrintKey(tabKey, group.key);
    const savedSettings = pageSettings[printKey];
    const autoSettings = getA4PreviewSettings(label, displayColumns, group.rows);
    const settings = savedSettings || autoSettings;
    const hasNote = Boolean(group.note || previewMeta.note);
    const rowsPerPage = getRowsPerPreviewPage(label, settings, hasNote);
    const chunks = chunkRows(group.rows, rowsPerPage);
    const totalChunks = Math.max(1, chunks.length);
    (chunks.length ? chunks : [[]]).forEach((chunk, chunkIndex) => {
      pages.push({
        key: `${group.key}-${chunkIndex}`,
        sheetKey: group.key,
        title: group.title || previewMeta.title,
        note: group.note || previewMeta.note || "",
        rows: chunk,
        sheetRows: group.rows,
        settings,
        pageNumber: chunkIndex + 1,
        totalChunks,
        columns: displayColumns,
        groupRowCount: group.rows.length,
        groupPageCount: totalChunks,
        groupPageIndex: chunkIndex + 1,
      });
    });
  }
  return pages.length
    ? pages
    : [{
        key: "empty",
        sheetKey: "empty",
        title: previewMeta.title,
        note: previewMeta.note || "",
        rows: [],
        sheetRows: [],
        settings: { orientation: "portrait", fontSize: 10, rowHeight: 18, zoom: 92 },
        pageNumber: 1,
        totalChunks: 1,
        columns: displayColumns,
        groupRowCount: 0,
        groupPageCount: 1,
        groupPageIndex: 1,
      }];
}

function getPageFitStatus(totalChunks, rowCount, rowsPerPage) {
  if (totalChunks <= 1) {
    const fillRatio = rowCount / Math.max(1, rowsPerPage);
    if (fillRatio > 0.9) return { status: "tight", label: "偏紧" };
    return { status: "fit", label: "适配" };
  }
  return { status: "over", label: `${totalChunks} 页` };
}

function getA4PreviewSettings(label, columns = [], rows = []) {
  if (String(label || "").includes("考试时间")) {
    return {
      orientation: "portrait",
      fontSize: 18,
      rowHeight: 46,
      zoom: 100,
    };
  }
  return autoFitOnePage(label, columns, rows);
}

function getOnePagePreviewSettings(label, columns = [], rows = [], hasNote = false) {
  const auto = autoFitOnePage(label, columns, rows);
  if (fitsOnePage(label, columns, rows, auto.orientation, auto, hasNote)) return auto;
  const rowCount = Math.max(1, rows.length);
  const usablePortrait = getUsablePageDimensions("portrait", hasNote);
  const usableLandscape = getUsablePageDimensions("landscape", hasNote);
  const portraitRowHeight = Math.max(8, Math.floor(usablePortrait.height / rowCount));
  const landscapeRowHeight = Math.max(8, Math.floor(usableLandscape.height / rowCount));
  if (portraitRowHeight >= landscapeRowHeight) {
    return {
      orientation: "portrait",
      rowHeight: portraitRowHeight,
      fontSize: clampNumber(portraitRowHeight - 4, 7, 12, 9),
      zoom: 96,
    };
  }
  return {
    orientation: "landscape",
    rowHeight: landscapeRowHeight,
    fontSize: clampNumber(landscapeRowHeight - 4, 7, 12, 9),
    zoom: 88,
  };
}

function autoFitOnePage(label, columns, rows) {
  const portraitTiers = [
    { fontSize: 12, rowHeight: 26, zoom: 96 },
    { fontSize: 11, rowHeight: 22, zoom: 96 },
    { fontSize: 10, rowHeight: 18, zoom: 96 },
    { fontSize: 9, rowHeight: 15, zoom: 96 },
  ];
  for (const tier of portraitTiers) {
    if (fitsOnePage(label, columns, rows, "portrait", tier, true)) {
      return { orientation: "portrait", ...tier };
    }
  }
  const landscapeTiers = [
    { fontSize: 11, rowHeight: 22, zoom: 88 },
    { fontSize: 10, rowHeight: 18, zoom: 88 },
    { fontSize: 9, rowHeight: 15, zoom: 88 },
  ];
  for (const tier of landscapeTiers) {
    if (fitsOnePage(label, columns, rows, "landscape", tier, true)) {
      return { orientation: "landscape", ...tier };
    }
  }
  return { orientation: "portrait", fontSize: 9, rowHeight: 15, zoom: 96 };
}

function fitsOnePage(label, columns, rows, orientation, settings, hasNote = true) {
  const usable = getUsablePageDimensions(orientation, hasNote);
  const widthNeeded = estimatePreviewWidth(label, columns, rows, orientation, settings.fontSize);
  const heightNeeded = (rows.length || 1) * settings.rowHeight;
  return widthNeeded <= usable.width && heightNeeded <= usable.height;
}

function getUsablePageDimensions(orientation, hasNote = true) {
  const pageWidth = orientation === "portrait" ? 595 : 842;
  const pageHeight = orientation === "portrait" ? 1123 : 794;
  const horizontalMargin = 40;
  const verticalPadding = 60;
  const titleHeight = 34;
  const noteHeight = hasNote ? 26 : 0;
  const headerHeight = 24;
  const footerHeight = 24;
  return {
    width: pageWidth - horizontalMargin * 2,
    height: pageHeight - verticalPadding - titleHeight - noteHeight - headerHeight - footerHeight,
  };
}

function estimatePreviewWidth(label, columns = [], rows = [], orientation, fontSize = 11) {
  const textHeavy = String(label || "").includes("班主任") || String(label || "").includes("考场信息");
  const baseCharWidth = orientation === "portrait" ? 7.2 : 7.0;
  const perChar = baseCharWidth * (Math.max(8, fontSize) / 11);
  const padding = orientation === "portrait" ? 20 : 18;
  const columnPadding = orientation === "portrait" ? 8 : 6;
  let total = 0;
  for (const column of columns) {
    const sample = rows.slice(0, 36).reduce((max, row) => Math.max(max, String(row?.[column] ?? "").length), String(column).length);
    const weight = getColumnWeight(column, textHeavy);
    const contentWidth = Math.max(weight, sample) * perChar + columnPadding;
    total += contentWidth;
  }
  return Math.max(160, total + padding * 2);
}

function getColumnWeight(column, textHeavy) {
  const label = String(column || "");
  if (label.includes("序号")) return 4;
  if (label.includes("日期")) return 10;
  if (label.includes("时间")) return 14;
  if (label.includes("考号")) return 14;
  if (label.includes("姓名")) return 6;
  if (label.includes("班级")) return 6;
  if (label.includes("科目")) return 8;
  if (label.includes("座位")) return 7;
  if (label.includes("考场")) return 8;
  if (label.includes("门牌")) return 6;
  if (label.includes("状态")) return 6;
  if (label.includes("外语")) return 8;
  if (label.includes("组合")) return 10;
  if (label.includes("排名")) return 7;
  if (label.includes("总分")) return 7;
  if (textHeavy) return 9;
  return 8;
}

function getPreviewFilterColumns(columns) {
  const preferred = ["级别", "班级", "班级分组", "考场号", "门牌号", "考试类型", "科目", "首选科目", "外语语种"];
  return preferred.filter((column) => columns.includes(column)).slice(0, 4);
}

function filterTableRows(rows, query, filters = {}) {
  const term = String(query || "").trim().toLowerCase();
  const activeFilters = Object.entries(filters).filter(([, value]) => String(value || "").trim());
  return rows.filter((row) => {
    const matchesQuery = !term || Object.values(row).some((value) => String(value || "").toLowerCase().includes(term));
    const matchesFilters = activeFilters.every(([column, value]) => String(row[column] ?? "") === String(value));
    return matchesQuery && matchesFilters;
  });
}

function buildPreviewTabs(schedule, rooms, scheduleMode, validationReport = []) {
  if (!schedule.allRows.length) return [];
  const roomDetails = buildRoomDetailRows(schedule);
  const classRows = buildClassPreviewRows(schedule.allRows);
  const tabs = [
    { key: "validation", label: "校验报告", rows: validationReport },
    { key: "times", label: "考试时间", rows: buildTimeRowsForPreview(schedule.examTimes || []) },
    { key: "doors", label: "门牌人数总览", rows: buildDoorRows(schedule, rooms) },
    { key: "admin", label: "管理总表", rows: buildAdminRows(schedule.allRows) },
    { key: "classes", label: "班主任表", rows: classRows },
    { key: "roomDetails", label: "考场信息表", rows: buildRoomPreviewRows(schedule) },
    { key: "main", label: "语数物/历", rows: assignmentRows(schedule.mainAssignments, { publicOnly: true }) },
    { key: "foreign", label: "外语", rows: assignmentRows(schedule.foreignAssignments, { publicOnly: true, paperGroup: "foreign" }) },
  ];
  if (scheduleMode === SCHEDULE_MODES.THREE_DAY_SPLIT) {
    for (const subject of ["化学", "地理", "政治", "生物"]) {
      tabs.push({ key: subject, label: subject, rows: buildSubjectRowsForPreview(schedule, subject) });
    }
  } else {
    tabs.push({ key: "elective", label: "四选二", rows: assignmentRows(schedule.electiveAssignments, { publicOnly: true }) });
  }
  return tabs.filter((tab) => tab.rows.length);
}

function buildTimeRowsForPreview(examTimes = []) {
  return examTimes
    .filter((item) => item.subject || item.date || item.start || item.end)
    .map((item, index) => ({
      序号: index + 1,
      科目: item.subject,
      日期: item.date,
      时间: [item.start, item.end].filter(Boolean).join("-"),
    }));
}

function describePreviewAssignments(assignments = []) {
  const counter = new Map();
  for (const item of assignments) {
    const key = item.subjectLabel || item.plan || "";
    if (!key) continue;
    counter.set(key, (counter.get(key) || 0) + 1);
  }
  return [...counter.entries()].map(([key, count]) => `${key}${count}`).join("+");
}

function buildClassPreviewRows(rows) {
  const printRows = rows
    .slice()
    .sort((a, b) => String(a.班级).localeCompare(String(b.班级), "zh-Hans-CN", { numeric: true }) || String(a.考号).localeCompare(String(b.考号), "zh-Hans-CN", { numeric: true }))
    .map((row) => ({ __pageGroup: row.班级, 班级分组: row.班级, ...buildPrintRows([row])[0] }));
  const rowsByClass = new Map();
  for (const row of printRows) {
    if (!rowsByClass.has(row.__pageGroup)) rowsByClass.set(row.__pageGroup, []);
    rowsByClass.get(row.__pageGroup).push(row);
  }
  return printRows.map((row) => {
    const classRows = rowsByClass.get(row.__pageGroup) || [];
    const hasPhysics = classRows.some((item) => String(item["语数物/座位号"] || "").trim());
    const hasHistory = classRows.some((item) => String(item["语数历/座位号"] || "").trim());
    const next = { ...row };
    if (!hasPhysics) delete next["语数物/座位号"];
    if (!hasHistory) delete next["语数历/座位号"];
    return next;
  });
}

function buildRoomPreviewRows(schedule) {
  return buildRoomPrintRows(schedule).map((row) => {
    const suffix = row.当科 === "自习" ? `${row.考场号}自习室` : row.考场号;
    const group = `${row.科目}-${suffix}`;
    const { 科目, ...printRow } = row;
    return {
      __pageGroup: group,
      __selfStudy: row.当科 === "自习",
      ...printRow,
    };
  });
}

function buildSubjectRowsForPreview(schedule, subject) {
  return buildSubjectPrintRows(schedule, subject).map((row) => {
    const raw = schedule.subjectAssignments.find((item) =>
      item.subjectLabel === subject &&
      item.studentId === row.考号 &&
      String(item.roomNo) === String(row.考场号) &&
      Number(item.seatNo) === Number(row.座位号)
    );
    const type = raw?.status || row.当科 || "考试";
    return {
      __paperGroup: `${subject}|${type}|${row.考场号}`,
      __selfStudy: type === "自习",
      ...row,
      当科: type === "自习" ? "自习" : subject,
    };
  });
}

function ValidationReport({ rows, summary }) {
  return (
    <section className="issue-panel">
      <div className="section-heading">
        <h3>校验报告</h3>
        <span>阻断 {summary.blockers} · 风险 {summary.warnings} · 复核 {summary.reviews}</span>
      </div>
      <div className="validation-badges">
        <span className={summary.blockers ? "danger" : "success"}>{summary.blockers ? "有硬错误，暂不能导出" : "硬校验通过"}</span>
        <span className={summary.warnings ? "warning" : "success"}>{summary.warnings ? `${summary.warnings} 条风险提醒` : "无风险提醒"}</span>
        <span>{summary.reviews} 条人工复核</span>
      </div>
      <div className="table-wrap issue-table">
        <table>
          <thead>
            <tr>
              <th>级别</th>
              <th>项目</th>
              <th>结果</th>
              <th>详情</th>
              <th>建议处理</th>
            </tr>
          </thead>
          <tbody>
            {rows.slice(0, 80).map((row, index) => (
              <tr key={`${row.级别}-${row.项目}-${index}`} className={validationRowClass(row.级别)}>
                <td>{row.级别}</td>
                <td>{row.项目}</td>
                <td>{row.结果}</td>
                <td>{row.详情}</td>
                <td>{row.建议处理}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </section>
  );
}

function validationRowClass(level) {
  if (level === "阻断错误") return "validation-blocker";
  if (level === "风险提醒") return "validation-warning";
  return "validation-review";
}

function StudentSearch({ query, onQuery, rows }) {
  const printRows = buildPrintRows(rows);
  return (
    <section className="search-panel">
      <div className="section-heading">
        <h3>学生查询</h3>
        <span>按姓名、考号或班级快速核对</span>
      </div>
      <label className="search-box">
        <Search size={18} />
        <input value={query} placeholder="输入姓名、考号、班级..." onChange={(event) => onQuery(event.target.value)} />
      </label>
      {query.trim() && (
        <div className="table-wrap issue-table">
          <table>
            <thead>
              <tr>
                <th>班级</th>
                <th>姓名</th>
                <th>考号</th>
                <th>语数物/历</th>
                <th>外语</th>
                <th>化学</th>
                <th>地理</th>
                <th>政治</th>
                <th>生物</th>
              </tr>
            </thead>
            <tbody>
              {printRows.map((row) => (
                <tr key={row.考号}>
                  <td>{row.__className || ""}</td>
                  <td>{row.姓名}</td>
                  <td>{row.考号}</td>
                  <td>{row["语数物/座位号"] || row["语数历/座位号"] || ""}</td>
                  <td>{row.外语 || ""}</td>
                  <td>{row.化学 || ""}</td>
                  <td>{row.地理 || ""}</td>
                  <td>{row.政治 || ""}</td>
                  <td>{row.生物 || ""}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </section>
  );
}

function buildIssueRows(errors) {
  return errors.map((message) => ({
    level: "阻断",
    message,
    action: suggestAction(message),
  }));
}

function createIssueRoute(message) {
  const text = String(message || "");
  if (text.includes("缺少物理类学生成绩单")) {
    return issueRoute(0, "导入", "⬆️", text, "再导入一份物理类成绩单", "当前只有历史类或空名单");
  }
  if (text.includes("缺少历史类学生成绩单")) {
    return issueRoute(0, "导入", "⬆️", text, "再导入一份历史类成绩单", "当前只有物理类或空名单");
  }
  if (text.includes("重复考号")) {
    return issueRoute(0, "导入", "⬆️", text, "在名单里搜索这个考号并删除/改正重复项", "同一个考号不能出现两次");
  }
  if (text.includes("再选科不是两门") || text.includes("选了") || text.includes("成绩为空")) {
    return issueRoute(0, "导入", "⬆️", text, "去名单里改选科组合或补齐单科成绩", "这类错误都在学生名单里修");
  }
  const language = LANGUAGE_SUBJECTS.find((item) => text.includes(item));
  if (language && text.includes("未指定外语考试考场")) {
    return issueRoute(2, "外语安排", "💬", text, `设置${language}考场`, `${language}学生需要指定或自动匹配外语考场`);
  }
  if (text.includes("小语种") || text.includes("日语") || text.includes("俄语") || text.includes("西班牙语") || text.includes("法语") || text.includes("德语") || text.includes("未指定") || text.includes("外语时段")) {
    return issueRoute(2, "外语安排", "💬", text, "设置外语语种考场", text || "这里调整各语种使用的考场号");
  }
  if (text.includes("缺少普通考场清单")) {
    return issueRoute(1, "考场", "🏫", text, "先导入考场模板或生成默认考场", "当前没有启用的普通考场，所以还不能计算容量和座位");
  }
  if (text.includes("英语普通考场容量不足")) {
    return issueRoute(1, "考场", "🏫", text, "新增英语考场或扩大现有容量", text);
  }
  if (text.includes("四选二普通考场容量不足")) {
    return issueRoute(1, "考场", "🏫", text, "补充四选二可用考场", text);
  }
  if (text.includes("考试+自习普通考场容量不足")) {
    const subject = ["化学", "地理", "政治", "生物"].find((item) => text.includes(item)) || "该科";
    return issueRoute(1, "考场", "🏫", text, `给${subject}补考场或加容量`, text);
  }
  if (text.includes("普通考场容量不足")) {
    return issueRoute(1, "考场", "🏫", text, "新增考场或提升容量", text);
  }
  if (text.includes("门牌") || text.includes("考场") || text.includes("容量") || text.includes("座位") || text.includes("同时用于")) {
    if (text.includes("同时用于")) {
      return issueRoute(1, "考场", "🏫", text, "处理门牌/教室冲突", text);
    }
    if (text.includes("座重复")) {
      return issueRoute(1, "考场", "🏫", text, "处理同考场座位重复", text);
    }
    if (text.includes("门牌")) {
      return issueRoute(1, "考场", "🏫", text, "修正门牌号或教室", text);
    }
    return issueRoute(1, "考场", "🏫", text, "调整考场号或容量", text);
  }
  if (text.includes("日期") || text.includes("时间") || text.includes("时段")) {
    return issueRoute(3, "时间", "🕘", text, "去检查考试时间", "这里改日期和每场开始/结束时间");
  }
  if (text.includes("漏排") || text.includes("重复排") || text.includes("混入") || text.includes("混场")) {
    return issueRoute(4, "预览", "🔎", text, "去看校验报告", "这里看完整校验和导出前摘要");
  }
  return issueRoute(4, "预览", "🔎", text, "去看校验报告", "这里看完整校验和导出前摘要");
}

function issueRoute(stepIndex, stepTitle, emoji, message, summary, detail) {
  return {
    stepIndex,
    stepTitle,
    emoji,
    message,
    summary,
    detail: detail || message,
  };
}

function getStatusLabel(status) {
  if (status === "done") return "已就绪";
  if (status === "error") return "需处理";
  return "待进行";
}

function suggestAction(message) {
  if (message.includes("容量不足") || (message.includes("缺") && message.includes("座位"))) return "在“确认考场”里新增考场或调整容量";
  if (message.includes("小语种") || message.includes("未指定")) return "在“外语安排”步骤里填写或调整语种考场";
  if (message.includes("再选科") || message.includes("选了")) return "在“导入成绩单”的名单表里修正选科组合或单科成绩";
  if (message.includes("重复考号")) return "在名单表中搜索该考号，删除或修正重复学生";
  if (message.includes("缺少")) return "检查导入表头，或在名单表中补齐学生信息";
  return "按提示修正后重新查看预览";
}

function searchStudents(rows, query) {
  const term = query.trim().toLowerCase();
  if (!term) return [];
  return rows
    .filter((row) => [row.姓名, row.考号, row.班级].some((value) => String(value || "").toLowerCase().includes(term)))
    .slice(0, 50);
}

function assignmentRows(assignments, options = {}) {
  return assignments.map((item) => ({
    ...(options.paperGroup === "foreign" ? { __paperGroup: `${item.subjectLabel || item.language || "外语"}|${item.roomNo}` } : {}),
    考试类型: item.plan === "主考" ? "语数物历" : item.plan,
    考场号: item.status === "自习" ? `${item.roomNo}自习室` : item.roomNo,
    门牌号: item.doorNo,
    教室: item.roomName,
    座位号: item.seatNo,
    班级: item.className,
    姓名: item.name,
    考号: item.studentId,
    首选科目: item.firstSubject,
    选科组合: item.comboRaw,
    外语语种: item.language,
    __selfStudy: item.status === "自习",
    ...(options.publicOnly ? {} : { 状态: item.status || "", 该科分数: item.subjectScore ?? "" }),
  }));
}

function formatShortDate(dateString) {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(dateString || "");
  if (!match) return "未设置";
  return `${Number(match[2])}.${Number(match[3])}`;
}

function formatRecordTime(value) {
  if (!value) return "本机保存";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "本机保存";
  const yyyy = date.getFullYear();
  const mm = String(date.getMonth() + 1).padStart(2, "0");
  const dd = String(date.getDate()).padStart(2, "0");
  const hour = String(date.getHours()).padStart(2, "0");
  const minute = String(date.getMinutes()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd} ${hour}:${minute}`;
}

createRoot(document.getElementById("root")).render(<App />);
