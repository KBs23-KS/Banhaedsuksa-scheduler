import React, { useState, useEffect, useMemo, useRef } from 'react';
import * as XLSX from 'xlsx';
import { initializeApp } from 'firebase/app';
import { getAuth, signInAnonymously, onAuthStateChanged } from 'firebase/auth';
import { getFirestore, collection, doc, setDoc, onSnapshot, deleteDoc, getDocs, getDocsFromServer, getDocFromServer, collectionGroup, query, where, writeBatch, runTransaction } from 'firebase/firestore';
import { 
  FileJson, Settings, X, AlertCircle, CheckCircle2, Plus, Edit, Trash2, ChevronUp, ChevronDown, 
  ChevronsUpDown, Users, BookOpen, Home, Briefcase, CalendarDays, Printer, LayoutDashboard, 
  CalendarX, Menu, RefreshCw, Download, UploadCloud, Pin, Clock, BarChart3, AlertTriangle, Check,
  Search, Filter, FileSpreadsheet, Activity, Info
} from 'lucide-react';

const getFirebaseConfig = () => {
  return {
    apiKey: "AIzaSyBen33CMqdqdvppNTUSw99WBPv_AsKtVEQ",
    authDomain: "school-scheduler-bhs.firebaseapp.com",
    projectId: "school-scheduler-bhs",
    storageBucket: "school-scheduler-bhs.firebasestorage.app",
    messagingSenderId: "148498085376",
    appId: "1:148498085376:web:f966a0d92fb2d663db4ae5",
    measurementId: "G-TD0R1D5S9V"
  };
};

const app = initializeApp(getFirebaseConfig());
const auth = getAuth(app);
const db = getFirestore(app);
const appId = 'bhs-school-scheduler-v3'; 

// STEP 04: Scheduler lock is acquired/released atomically and belongs to one run.
// The lease duration remains 5 minutes (matching the old lock), with periodic renewal.
const SCHEDULER_LOCK_TTL_MS = 5 * 60 * 1000;
const SCHEDULER_LOCK_RENEW_MS = 60 * 1000;

const acquireSchedulerLock = (lockRef, ownerId) => runTransaction(db, async tx => {
  const snap = await tx.get(lockRef);
  const data = snap.exists() ? snap.data() : null;
  const timestamp = Number(data?.timestamp);
  // Treat a lock with a missing/invalid timestamp as occupied, rather than
  // guessing it is safe to steal a possibly active lock.
  const active = data?.locked === true &&
    (!Number.isFinite(timestamp) || timestamp <= 0 || Date.now() - timestamp < SCHEDULER_LOCK_TTL_MS);
  if (active) return false;
  tx.set(lockRef, { locked: true, timestamp: Date.now(), ownerId });
  return true;
});

const renewSchedulerLock = (lockRef, ownerId) => runTransaction(db, async tx => {
  const snap = await tx.get(lockRef);
  if (!snap.exists() || snap.data().locked !== true || snap.data().ownerId !== ownerId) return false;
  tx.update(lockRef, { timestamp: Date.now() });
  return true;
});

const releaseSchedulerLock = (lockRef, ownerId) => runTransaction(db, async tx => {
  const snap = await tx.get(lockRef);
  if (!snap.exists() || snap.data().locked !== true || snap.data().ownerId !== ownerId) return false;
  tx.update(lockRef, { locked: false, timestamp: Date.now() });
  return true;
});

const DEFAULT_PERIODS = [
  { id: '1', name: 'คาบ 1', start: '08:40', end: '09:30', isBreak: false },
  { id: '2', name: 'คาบ 2', start: '09:30', end: '10:20', isBreak: false },
  { id: '3', name: 'คาบ 3', start: '10:20', end: '11:10', isBreak: false },
  { id: '4', name: 'คาบ 4', start: '11:10', end: '12:00', isBreak: false },
  { id: '5', name: 'พักเที่ยง', start: '12:00', end: '13:00', isBreak: true },
  { id: '6', name: 'คาบ 5', start: '13:00', end: '13:50', isBreak: false },
  { id: '7', name: 'คาบ 6', start: '13:50', end: '14:40', isBreak: false },
  { id: '8', name: 'คาบ 7', start: '14:40', end: '15:30', isBreak: false },
];

const DAYS = ['จันทร์', 'อังคาร', 'พุธ', 'พฤหัสบดี', 'ศุกร์'];
// STEP 10: กิจกรรมล็อกทุกห้องเก็บเป็นเอกสารเดียว (ครูผู้ดูแลไม่นับซ้ำหลายห้อง)
const ALL_CLASSROOMS_ID = '__ALL_CLASSROOMS__';
const fixedAppliesToClassroom = (fixedItem, classroomId) =>
  fixedItem.classroomId === classroomId || fixedItem.classroomId === ALL_CLASSROOMS_ID;
const SUBJECT_AREAS = [
  "ไม่มีกลุ่มสาระการเรียนรู้", "ภาษาไทย", "คณิตศาสตร์", "วิทยาศาสตร์และเทคโนโลยี", 
  "สังคมศึกษา ศาสนา และวัฒนธรรม", "สุขศึกษาและพลศึกษา", "ศิลปะ", "การงานอาชีพ", "ภาษาต่างประเทศ"
];

const SUBJECT_COLORS = {
  "ไม่มีกลุ่มสาระการเรียนรู้": "text-slate-500", "ภาษาไทย": "text-rose-600", "คณิตศาสตร์": "text-blue-600",
  "วิทยาศาสตร์และเทคโนโลยี": "text-emerald-600", "สังคมศึกษา ศาสนา และวัฒนธรรม": "text-amber-600", 
  "สุขศึกษาและพลศึกษา": "text-orange-500", "ศิลปะ": "text-purple-600", "การงานอาชีพ": "text-cyan-600",
  "ภาษาต่างประเทศ": "text-indigo-600"
};

const SCHEDULER_CONFIG = { candidateCount: 50, maxConsecutiveTeacherPeriods: 3, maxTeacherPeriodsPerDay: 5, candidateChoicePool: 4, placementJitter: 2.5, yieldEveryCandidates: 5 };
const SCHEDULE_VERSION_LIMIT = 10;
const SCHEDULER_CANDIDATE_KEEP = 3;
const WEIGHTS = { BASE_SCORE: 10000, SUBJECT_SAME_DAY: -25, TEACHER_CONSECUTIVE: -35, TEACHER_GAP: -8, TEACHER_OVERLOAD: -40, WRONG_TIME_PREFERENCE: -8, DAILY_IMBALANCE: -4, GOOD_SPREAD: +12, NO_GAP: +4 };

const getShortTeacherName = (fullName) => {
  if (!fullName || fullName === 'ไม่ทราบ') return '-';
  let name = String(fullName).replace(/^(คุณครู|นาย|นางสาว|นาง|ผอ\.|รองฯ\s?)/, '').trim();
  return 'ครู' + name.split(' ')[0];
};

const sanitizeExcelInput = (text) => {
  if (text === null || text === undefined) return '';
  const str = String(text);
  return str.match(/^[=@+-]/) ? `'${str}` : str;
};

const getPeriodDisplayName = (periods, periodId) => {
  const period = Array.isArray(periods) ? periods.find(p => String(p.id) === String(periodId)) : null;
  return period?.name || `คาบ ${periodId ?? '-'}`;
};

const isAutoGeneratedSchedule = (scheduleItem) => {
  if (!scheduleItem) return false;
  if (scheduleItem.source === 'auto') return true;
  return /^sch_auto_/.test(String(scheduleItem.id || ''));
};

const getTeachingLoadKey = (item) => `${item?.teacherId || ''}_${item?.subjectId || ''}_${item?.classroomId || ''}`;

const cleanScheduleSnapshotItem = (item) => {
  const clean = { ...(item || {}) };
  delete clean.__source;
  delete clean._col;
  return clean;
};

const stableSerializeData = (value) => {
  if (value === undefined) return 'undefined';
  if (value === null || typeof value !== 'object') return JSON.stringify(value);
  if (Array.isArray(value)) return `[${value.map(stableSerializeData).join(',')}]`;
  return `{${Object.keys(value).filter(key => value[key] !== undefined).sort().map(key => `${JSON.stringify(key)}:${stableSerializeData(value[key])}`).join(',')}}`;
};

const areFirestoreRecordsEqual = (a, b) => stableSerializeData(a) === stableSerializeData(b);

const useSortableData = (items, config = null) => {
  const [sortConfig, setSortConfig] = useState(config || { key: null, direction: 'asc' });

  const sortedItems = useMemo(() => {
    if (!Array.isArray(items)) return [];
    let sortableItems = [...items];
    if (sortConfig && sortConfig.key) {
      sortableItems.sort((a, b) => {
        let valA = a[sortConfig.key] || '';
        let valB = b[sortConfig.key] || '';
        if (typeof valA === 'string' && typeof valB === 'string') return sortConfig.direction === 'asc' ? valA.localeCompare(valB, 'th') : valB.localeCompare(valA, 'th');
        if (valA < valB) return sortConfig.direction === 'asc' ? -1 : 1;
        if (valA > valB) return sortConfig.direction === 'asc' ? 1 : -1;
        return 0;
      });
    }
    return sortableItems;
  }, [items, sortConfig]);

  const requestSort = (key) => {
    let direction = 'asc';
    if (sortConfig && sortConfig.key === key && sortConfig.direction === 'asc') direction = 'desc';
    setSortConfig({ key, direction });
  };
  return { items: sortedItems, requestSort, sortConfig };
};

const SortIcon = ({ column, sortConfig }) => {
  if (sortConfig?.key !== column) return <ChevronsUpDown size={14} className="opacity-40" />;
  return sortConfig.direction === 'asc' ? <ChevronUp size={14} /> : <ChevronDown size={14} />;
};

const calculateScheduleHealth = (schedules, loads, unavail, cUnavail, fixed, teachers, subjects, classrooms, periods) => {
  const health = {
    status: 'ready', totalRequired: 0, scheduledForLoads: 0, missing: 0, overScheduled: 0, completionPercent: 0,
    teacherConflicts: [], roomConflicts: [], unavailConflicts: [], cUnavailConflicts: [], breakConflicts: [], blockConflicts: [], teacherLimitConflicts: [], orphans: [], fixedConflicts: [], missingLoads: [], overloads: [],
    teacherStats: {}, roomStats: {}, 
  };

  const teacherMap = new Map(teachers.map(t => [t.id, t]));
  const subjectMap = new Map(subjects.map(s => [s.id, s]));
  const roomMap = new Map(classrooms.map(c => [c.id, c]));
  const periodContext = buildPeriodContext(periods);
  const validPeriodIds = periodContext.periodById;
  const breakPeriods = periodContext.breakIds;

  const loadTracker = {}; 
  loads.forEach(l => {
    const req = Number(l.periods) || 0;
    if (!Number.isInteger(Number(l.periods)) || req < 1 || l.periods === '' || l.periods == null) {
      health.orphans.push(`ภาระสอนครู ${l.teacherId || '-'} วิชา ${l.subjectId || '-'} มีจำนวนคาบไม่ถูกต้อง`);
    }
    health.totalRequired += req;
    const lKey = `${l.teacherId}_${l.subjectId}_${l.classroomId}`;
    if (loadTracker[lKey]) {
      // ห้ามให้รายการซ้ำเขียนทับจำนวนคาบเดิม เพราะจะทำให้ Health Check แสดงผลผิด
      loadTracker[lKey].required += req;
      health.orphans.push(`พบภาระสอนซ้ำสำหรับครู ${l.teacherId} วิชา ${l.subjectId} ห้อง ${l.classroomId}`);
    } else {
      loadTracker[lKey] = { ...l, required: req, scheduled: 0 };
    }
    if (!health.teacherStats[l.teacherId]) health.teacherStats[l.teacherId] = { req: 0, sch: 0, days: {}, gaps: 0, maxConsec: 0, subjects: new Set() };
    health.teacherStats[l.teacherId].req += req;
    health.teacherStats[l.teacherId].subjects.add(l.subjectId);
  });

  const tPlacement = {}; const rPlacement = {};
  
  const checkPlacement = (sch, isFixed = false) => {
    const teacherExists = !sch.teacherId || teacherMap.has(sch.teacherId);
    const roomExists = roomMap.has(sch.classroomId) || (isFixed && sch.classroomId === ALL_CLASSROOMS_ID && classrooms.length > 0);
    const subjectExists = subjectMap.has(sch.subjectId);
    const dayExists = DAYS.includes(sch.day);
    const periodExists = validPeriodIds.has(sch.periodId);

    if (sch.teacherId && !teacherExists) health.orphans.push(`ไม่พบข้อมูลครู รหัส ${sch.teacherId}`);
    if (!roomExists) health.orphans.push(`ไม่พบข้อมูลห้อง รหัส ${sch.classroomId}`);
    if (!subjectExists) health.orphans.push(`ไม่พบข้อมูลวิชา รหัส ${sch.subjectId}`);
    if (!dayExists) health.orphans.push(`ไม่พบวันในระบบ: ${sch.day || '-'}`);
    if (!periodExists) health.orphans.push(`ไม่พบคาบในระบบ: ${sch.periodId || '-'}`);
    if (periodExists && breakPeriods.has(sch.periodId)) health.breakConflicts.push(`${subjectMap.get(sch.subjectId)?.name || sch.subjectId} ลงเวลาพัก (วัน${sch.day} ${getPeriodDisplayName(periods, sch.periodId)})`);

    if (sch.teacherId) {
      const tKey = `${sch.teacherId}_${sch.day}_${sch.periodId}`;
      if (!tPlacement[tKey]) tPlacement[tKey] = [];
      tPlacement[tKey].push(sch);
    }

    const roomIds = isFixed && sch.classroomId === ALL_CLASSROOMS_ID
      ? classrooms.map(c => c.id) : (sch.classroomId ? [sch.classroomId] : []);
    roomIds.forEach(roomId => {
      const rKey = `${roomId}_${sch.day}_${sch.periodId}`;
      if (!rPlacement[rKey]) rPlacement[rKey] = [];
      rPlacement[rKey].push(sch);
    });

    const hasCompleteTeachingRef = Boolean(sch.teacherId && sch.subjectId && sch.classroomId !== ALL_CLASSROOMS_ID && sch.classroomId);
    const lKey = hasCompleteTeachingRef ? `${sch.teacherId}_${sch.subjectId}_${sch.classroomId}` : '';
    const matchedLoad = lKey ? loadTracker[lKey] : null;

    if (matchedLoad) {
      matchedLoad.scheduled += 1;
      if (!health.teacherStats[sch.teacherId]) health.teacherStats[sch.teacherId] = { req: 0, sch: 0, days: {}, gaps: 0, maxConsec: 0, subjects: new Set() };
      health.teacherStats[sch.teacherId].sch += 1;
      if (!health.teacherStats[sch.teacherId].days[sch.day]) health.teacherStats[sch.teacherId].days[sch.day] = [];
      health.teacherStats[sch.teacherId].days[sch.day].push(sch.periodId);
      health.teacherStats[sch.teacherId].subjects.add(sch.subjectId);
    } else if (!isFixed) {
      if (!sch.teacherId) health.orphans.push(`ตารางสอน ${subjectMap.get(sch.subjectId)?.name || sch.subjectId || '-'} ไม่มีครูผู้สอน`);
      else if (teacherExists && roomExists && subjectExists) health.orphans.push(`ตารางสอนไม่มีภาระสอนรองรับ: ${subjectMap.get(sch.subjectId)?.name || sch.subjectId} / ${roomMap.get(sch.classroomId)?.name || sch.classroomId} / ${teacherMap.get(sch.teacherId)?.name || sch.teacherId}`);
    }
  };

  schedules.forEach(s => checkPlacement(s, false));
  fixed.forEach(f => checkPlacement(f, true));

  Object.entries(tPlacement).filter(([_, arr]) => arr.length > 1).forEach(([_, arr]) => health.teacherConflicts.push(`ครู ${teacherMap.get(arr[0].teacherId)?.name || arr[0].teacherId} ซ้อนทับ ${arr.length} วิชา (วัน${arr[0].day} ${getPeriodDisplayName(periods, arr[0].periodId)})`));
  Object.entries(rPlacement).filter(([_, arr]) => arr.length > 1).forEach(([_, arr]) => health.roomConflicts.push(`ห้อง ${roomMap.get(arr[0].classroomId)?.name || arr[0].classroomId} ซ้อนทับ ${arr.length} วิชา (วัน${arr[0].day} ${getPeriodDisplayName(periods, arr[0].periodId)})`));

  unavail.forEach(u => { if (tPlacement[`${u.teacherId}_${u.day}_${u.periodId}`]) health.unavailConflicts.push(`ครู ${teacherMap.get(u.teacherId)?.name || u.teacherId} ถูกจัดในเวลาไม่สะดวก (วัน${u.day} ${getPeriodDisplayName(periods, u.periodId)})`); });
  cUnavail.forEach(u => { if (rPlacement[`${u.classroomId}_${u.day}_${u.periodId}`]) health.cUnavailConflicts.push(`ห้อง ${roomMap.get(u.classroomId)?.name || u.classroomId} ถูกจัดในเวลาไม่ว่าง (วัน${u.day} ${getPeriodDisplayName(periods, u.periodId)})`); });

  let cappedScheduled = 0;
  Object.values(loadTracker).forEach(l => {
    const diff = l.required - l.scheduled;
    if (diff > 0) { health.missingLoads.push({ ...l, missing: diff }); health.missing += diff; cappedScheduled += l.scheduled; } 
    else if (diff < 0) { health.overloads.push(`จัดเกินโควตา: ${subjectMap.get(l.subjectId)?.name || l.subjectId} ห้อง ${roomMap.get(l.classroomId)?.name || l.classroomId} (เกิน ${Math.abs(diff)} คาบ)`); health.overScheduled += Math.abs(diff); cappedScheduled += l.required; } 
    else { cappedScheduled += l.required; }
  });

  // ตรวจความสมบูรณ์ของคาบคู่ แม้ยอดคาบจะตรงโควตาแล้วก็ตาม
  Object.values(loadTracker).forEach(l => {
    if (Number(subjectMap.get(l.subjectId)?.periodBlock) !== 2) return;
    const relevant = [...schedules, ...fixed].filter(s =>
      s.teacherId === l.teacherId && s.subjectId === l.subjectId && s.classroomId === l.classroomId
    );
    const singles = analyzeBlockPairing(relevant, periodContext).singles;
    const allowedSingles = l.required % 2;
    if (singles.length > allowedSingles) {
      health.blockConflicts.push(`วิชา ${subjectMap.get(l.subjectId)?.name || l.subjectId} / ห้อง ${roomMap.get(l.classroomId)?.name || l.classroomId} มีคาบเดี่ยว ${singles.length} คาบ (อนุญาต ${allowedSingles})`);
    }
  });

  health.scheduledForLoads = cappedScheduled;
  health.completionPercent = health.totalRequired > 0 ? Math.min(100, Math.round((cappedScheduled / health.totalRequired) * 100)) : 0;

  Object.keys(health.teacherStats).forEach(tId => {
    const stat = health.teacherStats[tId];
    let totalGaps = 0; let globalMaxConsec = 0;
    Object.values(stat.days).forEach(dayPeriods => {
      const shape = calculateDayTeachingShape(dayPeriods, periodContext);
      totalGaps += shape.gaps;
      globalMaxConsec = Math.max(globalMaxConsec, shape.maxConsec);
    });
    stat.gaps = totalGaps; stat.maxConsec = globalMaxConsec;
  });

  health.teacherLimitConflicts = findTeacherTeachingLimitConflicts([...schedules, ...fixed], periods, teachers);
  const hasHardConflict = health.teacherConflicts.length > 0 || health.roomConflicts.length > 0 || health.unavailConflicts.length > 0 || health.cUnavailConflicts.length > 0 || health.breakConflicts.length > 0 || health.blockConflicts.length > 0 || health.teacherLimitConflicts.length > 0 || health.orphans.length > 0;
  if (hasHardConflict) health.status = 'error';
  else if (loads.length === 0) health.status = 'empty';
  else if (health.missing > 0 || health.overScheduled > 0) health.status = 'warning';
  else health.status = 'ready';

  return health;
};

// ข้อบังคับครู: นับคาบที่ครูมีสอน/กิจกรรมล็อกจริงรวมกัน และไม่นับคาบพักเป็นคาบติดกัน
const teacherDayKey = (teacherId, day) => JSON.stringify([teacherId, day]);
const registerTeacherPeriod = (maps, item) => {
  if (!item.teacherId) return;
  const key = teacherDayKey(item.teacherId, item.day);
  if (!maps.teacherDayPeriods.has(key)) maps.teacherDayPeriods.set(key, new Set());
  maps.teacherDayPeriods.get(key).add(item.periodId);
};

const canPlaceTeacherEntries = (teacherId, entries, maps) => {
  if (!teacherId) return { ok: true, reason: '' };
  const byDay = new Map();
  entries.forEach(item => {
    if (!byDay.has(item.day)) byDay.set(item.day, []);
    byDay.get(item.day).push(item.periodId);
  });
  for (const [day, additions] of byDay) {
    const existing = maps.teacherDayPeriods.get(teacherDayKey(teacherId, day)) || new Set();
    const result = new Set([...existing, ...additions]);
    if (result.size > SCHEDULER_CONFIG.maxTeacherPeriodsPerDay) {
      return { ok: false, reason: 'TEACHER_DAILY_LIMIT' };
    }
    if (calculateDayTeachingShape([...result], maps.periodContext).maxConsec > SCHEDULER_CONFIG.maxConsecutiveTeacherPeriods) {
      return { ok: false, reason: 'TEACHER_CONSECUTIVE_LIMIT' };
    }
  }
  return { ok: true, reason: '' };
};

const buildConstraintMaps = (schedules, tUnavail, cUnavail, fixedSch, periods = DEFAULT_PERIODS) => {
  const maps = { tUnavailMap: new Set(), cUnavailMap: new Set(), fixedMap: new Set(), teacherPlacements: new Set(), roomPlacements: new Set(), teacherDayPeriods: new Map(), periodContext: buildPeriodContext(periods) };
  tUnavail.forEach(u => maps.tUnavailMap.add(`${u.teacherId}_${u.day}_${u.periodId}`));
  cUnavail.forEach(u => maps.cUnavailMap.add(`${u.classroomId}_${u.day}_${u.periodId}`));
  fixedSch.forEach(f => {
    maps.fixedMap.add(`${f.day}_${f.periodId}_${f.classroomId}`);
    if (f.teacherId) maps.teacherPlacements.add(`${f.teacherId}_${f.day}_${f.periodId}`);
    registerTeacherPeriod(maps, f);
    maps.roomPlacements.add(`${f.classroomId}_${f.day}_${f.periodId}`);
  });
  schedules.forEach(s => {
    if (s.teacherId) maps.teacherPlacements.add(`${s.teacherId}_${s.day}_${s.periodId}`);
    registerTeacherPeriod(maps, s);
    maps.roomPlacements.add(`${s.classroomId}_${s.day}_${s.periodId}`);
  });
  return maps;
};

const canPlaceSchedule = (teacherId, classroomId, day, periodId, maps, isBreak) => {
  if (isBreak) return { ok: false, reason: "BREAK_PERIOD" };
  if (maps.cUnavailMap.has(`${classroomId}_${day}_${periodId}`)) return { ok: false, reason: "ROOM_UNAVAILABLE" };
  if (teacherId && maps.tUnavailMap.has(`${teacherId}_${day}_${periodId}`)) return { ok: false, reason: "TEACHER_UNAVAILABLE" };
  if (maps.fixedMap.has(`${day}_${periodId}_${classroomId}`) || maps.fixedMap.has(`${day}_${periodId}_${ALL_CLASSROOMS_ID}`)) return { ok: false, reason: "FIXED_SCHEDULE_CONFLICT" };
  if (maps.roomPlacements.has(`${classroomId}_${day}_${periodId}`)) return { ok: false, reason: "ROOM_CONFLICT" };
  if (teacherId && maps.teacherPlacements.has(`${teacherId}_${day}_${periodId}`)) return { ok: false, reason: "TEACHER_CONFLICT" };
  const teachingLimit = canPlaceTeacherEntries(teacherId, [{ day, periodId }], maps);
  if (!teachingLimit.ok) return teachingLimit;
  return { ok: true, reason: "" };
};

// STEP 04: ตรวจความพร้อมของข้อมูลก่อนสร้างตาราง (ไม่มีการเขียน Firebase)
// ตรวจเฉพาะสิ่งที่ทำให้ตารางไม่ถูกต้อง โดยไม่ปฏิเสธเพียงเพราะช่องเวลามีไม่พอ
const validateAutoScheduleInputs = ({ teachingLoads, teachers, subjects, classrooms, periods,
  schedules, fixedSchedules, unavailabilities, classroomUnavailabilities, semester }) => {
  const problems = [];
  const addProblem = message => { if (problems.length < 30) problems.push(message); };
  const teacherIds = new Set(teachers.map(t => String(t.id)));
  const classroomIds = new Set(classrooms.map(c => String(c.id)));
  const subjectById = new Map(subjects.map(s => [String(s.id), s]));
  const periodById = new Map(periods.map(p => [String(p.id), p]));
  const validSubject = id => {
    const subject = subjectById.get(String(id || ''));
    return !!(subject && (!subject.term || subject.term === 'all' || String(subject.term) === String(semester)));
  };
  const loadKey = item => JSON.stringify([String(item.teacherId || ''), String(item.subjectId || ''), String(item.classroomId || '')]);
  const loadByKey = new Map();

  if (!teachingLoads.length) addProblem('ยังไม่มีข้อมูลภาระสอน กรุณาเพิ่มภาระสอนก่อนจัดตาราง');
  if (!periods.some(p => !p.isBreak)) addProblem('ไม่พบคาบเรียนที่ใช้จัดตารางได้ กรุณาตรวจเมนูคาบเรียน');
  if (periodById.size !== periods.length) addProblem('รหัสคาบเรียนซ้ำกัน กรุณาตรวจการตั้งค่าคาบเรียน');

  teachingLoads.forEach((load, index) => {
    const label = `ภาระสอนรายการที่ ${index + 1}`;
    if (!load.teacherId || !teacherIds.has(String(load.teacherId))) addProblem(`${label}: ไม่พบรหัสครู ${load.teacherId || '-'}`);
    if (!load.subjectId || !validSubject(load.subjectId)) addProblem(`${label}: ไม่พบรหัสวิชา ${load.subjectId || '-'} ในภาคเรียนนี้`);
    if (!load.classroomId || !classroomIds.has(String(load.classroomId))) addProblem(`${label}: ไม่พบรหัสห้อง ${load.classroomId || '-'}`);
    if (!Number.isInteger(Number(load.periods)) || Number(load.periods) < 1 || load.periods === '' || load.periods == null) {
      addProblem(`${label}: จำนวนคาบต่อสัปดาห์ต้องเป็นจำนวนเต็มตั้งแต่ 1 ขึ้นไป`);
    }
    const key = loadKey(load);
    if (loadByKey.has(key)) addProblem(`${label}: ภาระสอนครู–วิชา–ห้องซ้ำกับรายการที่ ${loadByKey.get(key).index}`);
    else loadByKey.set(key, { load, index: index + 1 });
  });

  // ตารางที่เครื่องสร้างอัตโนมัติจะถูกสร้างใหม่ จึงตรวจเฉพาะตารางที่รักษาไว้กับกิจกรรมล็อก
  const anchors = [
    ...schedules.filter(s => !isAutoGeneratedSchedule(s)).map(s => ({ ...s, __kind: 'ตารางที่จัดด้วยมือ' })),
    ...fixedSchedules.map(f => ({ ...f, __kind: 'กิจกรรมล็อก' })),
  ];
  const usedTeachers = new Set();
  const usedRooms = new Set();
  const anchorCounts = new Map();
  const unavailableTeachers = new Set(unavailabilities.map(u => JSON.stringify([u.teacherId, u.day, String(u.periodId)])));
  const unavailableRooms = new Set(classroomUnavailabilities.map(u => JSON.stringify([u.classroomId, u.day, String(u.periodId)])));

  anchors.forEach((item, index) => {
    const label = `${item.__kind} รายการที่ ${index + 1}`;
    if (item.teacherId && !teacherIds.has(String(item.teacherId))) addProblem(`${label}: ไม่พบครู ${item.teacherId}`);
    if (!validSubject(item.subjectId)) addProblem(`${label}: ไม่พบรายวิชา ${item.subjectId || '-'} ในภาคเรียนนี้`);
    const appliesToAll = item.__kind === 'กิจกรรมล็อก' && item.classroomId === ALL_CLASSROOMS_ID;
    if (appliesToAll ? classrooms.length === 0 : !classroomIds.has(String(item.classroomId))) {
      addProblem(`${label}: ไม่พบห้อง ${item.classroomId || '-'}`);
    }
    if (!DAYS.includes(item.day)) addProblem(`${label}: วันไม่ถูกต้อง (${item.day || '-'})`);
    const period = periodById.get(String(item.periodId));
    if (!period || period.isBreak) addProblem(`${label}: คาบเรียนไม่ถูกต้องหรือเป็นคาบพัก (${item.periodId || '-'})`);
    const slotTeacher = JSON.stringify([item.teacherId, item.day, String(item.periodId)]);
    if (item.teacherId) {
      if (usedTeachers.has(slotTeacher)) addProblem(`${label}: ครู ${item.teacherId} มีตารางซ้อนในวัน${item.day} คาบ ${item.periodId}`);
      usedTeachers.add(slotTeacher);
      if (unavailableTeachers.has(slotTeacher)) addProblem(`${label}: ครู ${item.teacherId} ถูกจัดในเวลาไม่สะดวก`);
    }
    const affectedRoomIds = appliesToAll ? classrooms.map(c => c.id) : [item.classroomId];
    affectedRoomIds.forEach(roomId => {
      const slotRoom = JSON.stringify([roomId, item.day, String(item.periodId)]);
      if (usedRooms.has(slotRoom)) addProblem(`${label}: ห้อง ${roomId} มีตารางซ้อนในวัน${item.day} คาบ ${item.periodId}`);
      usedRooms.add(slotRoom);
      if (unavailableRooms.has(slotRoom)) addProblem(`${label}: ห้อง ${roomId} ถูกจัดในเวลาที่ไม่ว่าง`);
    });
    if (item.__kind === 'ตารางที่จัดด้วยมือ' && !item.teacherId) addProblem(`${label}: ไม่ระบุครูผู้สอน`);
    if (item.teacherId && !appliesToAll) {
      const key = loadKey(item);
      if (loadByKey.has(key)) anchorCounts.set(key, (anchorCounts.get(key) || 0) + 1);
      else if (item.__kind === 'ตารางที่จัดด้วยมือ') addProblem(`${label}: ไม่มีภาระสอนรองรับครู ${item.teacherId} วิชา ${item.subjectId} ห้อง ${item.classroomId}`);
      // กิจกรรมล็อกสามารถเป็นกิจกรรมทั่วไปที่ไม่อยู่ในภาระสอน จึงไม่บังคับให้มีภาระสอน
    }
  });
  anchorCounts.forEach((count, key) => {
    const { load } = loadByKey.get(key);
    if (Number.isFinite(Number(load.periods)) && count > Number(load.periods)) {
      addProblem(`ครู ${load.teacherId} วิชา ${load.subjectId} ห้อง ${load.classroomId}: คาบที่จัดเอง/ล็อก ${count} คาบ มากกว่าภาระสอน ${load.periods} คาบ`);
    }
  });
  findTeacherTeachingLimitConflicts(anchors, periods, teachers).forEach(addProblem);
  return problems;
};

const createSeededRandom = (seed) => {
  let value = (Number(seed) || 1) >>> 0;
  return () => {
    value = (value + 0x6D2B79F5) >>> 0;
    let t = value;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
};

const cloneConstraintMaps = (maps) => ({
  // Constraint สามชุดแรกเป็น read-only ระหว่างสร้าง candidate จึงแชร์ reference ได้
  // clone เฉพาะ placement ที่ถูกแก้ระหว่างการทดลองแต่ละ candidate เพื่อลด allocation จำนวนมาก
  tUnavailMap: maps.tUnavailMap,
  cUnavailMap: maps.cUnavailMap,
  fixedMap: maps.fixedMap,
  teacherPlacements: new Set(maps.teacherPlacements),
  roomPlacements: new Set(maps.roomPlacements),
  teacherDayPeriods: new Map([...maps.teacherDayPeriods].map(([day, values]) => [day, new Set(values)])),
  periodContext: maps.periodContext,
});

const buildPeriodContext = (periods) => {
  const safePeriods = Array.isArray(periods) ? periods : [];
  const activePeriods = safePeriods.filter(p => !p.isBreak);
  const activeOrder = activePeriods.map(p => p.id);
  const activeIndex = new Map(activeOrder.map((id, index) => [id, index]));
  const realIndex = new Map(safePeriods.map((p, index) => [p.id, index]));
  const periodById = new Map(safePeriods.map(p => [p.id, p]));
  const breakIds = new Set(safePeriods.filter(p => p.isBreak).map(p => p.id));
  const breakIndex = safePeriods.findIndex(p => p.isBreak);
  const morningIds = breakIndex >= 0
    ? new Set(safePeriods.slice(0, breakIndex).filter(p => !p.isBreak).map(p => p.id))
    : new Set(activeOrder.slice(0, Math.ceil(activeOrder.length / 2)));
  const doublePairs = [];
  const neighborIdsByPeriod = new Map();

  for (let i = 0; i < safePeriods.length; i++) {
    const period = safePeriods[i];
    if (period.isBreak) continue;
    const neighbors = [];
    const previous = safePeriods[i - 1];
    const next = safePeriods[i + 1];
    if (previous && !previous.isBreak) neighbors.push(previous.id);
    if (next && !next.isBreak) neighbors.push(next.id);
    neighborIdsByPeriod.set(period.id, neighbors);
    if (next && !next.isBreak) doublePairs.push([period, next]);
  }

  return { safePeriods, activePeriods, activeOrder, activeIndex, realIndex, periodById, breakIds, morningIds, doublePairs, neighborIdsByPeriod };
};

const getPeriodContext = (periodsOrContext) => {
  if (periodsOrContext?.activeIndex instanceof Map && periodsOrContext?.realIndex instanceof Map) return periodsOrContext;
  return buildPeriodContext(periodsOrContext);
};

const getMorningPeriodIds = (periodsOrContext) => getPeriodContext(periodsOrContext).morningIds;

const calculateDayTeachingShape = (dayPeriods, periodsOrContext) => {
  const context = getPeriodContext(periodsOrContext);
  const ordered = [...new Set(dayPeriods)]
    .filter(id => context.activeIndex.has(id) && context.realIndex.has(id))
    .map(id => ({ id, active: context.activeIndex.get(id), real: context.realIndex.get(id) }))
    .sort((a, b) => a.real - b.real);

  if (ordered.length === 0) return { gaps: 0, maxConsec: 0 };
  let gaps = 0; let currentConsec = 1; let maxConsec = 1;
  for (let i = 0; i < ordered.length - 1; i++) {
    const activeDiff = ordered[i + 1].active - ordered[i].active;
    const realDiff = ordered[i + 1].real - ordered[i].real;
    if (activeDiff === 1 && realDiff === 1) {
      currentConsec += 1;
      maxConsec = Math.max(maxConsec, currentConsec);
    } else {
      if (activeDiff > 1) gaps += activeDiff - 1;
      currentConsec = 1;
    }
  }
  return { gaps, maxConsec };
};

const findTeacherTeachingLimitConflicts = (items, periods, teachers = []) => {
  const byTeacherDay = new Map();
  const names = new Map(teachers.map(t => [t.id, t.name]));
  items.forEach(item => {
    if (!item.teacherId || !DAYS.includes(item.day)) return;
    const key = teacherDayKey(item.teacherId, item.day);
    if (!byTeacherDay.has(key)) byTeacherDay.set(key, { teacherId: item.teacherId, day: item.day, periodIds: new Set() });
    byTeacherDay.get(key).periodIds.add(item.periodId);
  });
  const violations = [];
  byTeacherDay.forEach(({ teacherId, day, periodIds }) => {
    const periodsOnDay = [...periodIds];
    const teacherName = names.get(teacherId) || teacherId;
    if (periodsOnDay.length > SCHEDULER_CONFIG.maxTeacherPeriodsPerDay) {
      violations.push(`ครู ${teacherName} วัน${day} มี ${periodsOnDay.length} คาบ (สูงสุด ${SCHEDULER_CONFIG.maxTeacherPeriodsPerDay} คาบ/วัน)`);
    }
    const maxConsec = calculateDayTeachingShape(periodsOnDay, periods).maxConsec;
    if (maxConsec > SCHEDULER_CONFIG.maxConsecutiveTeacherPeriods) {
      violations.push(`ครู ${teacherName} วัน${day} สอนติดกัน ${maxConsec} คาบ (สูงสุด ${SCHEDULER_CONFIG.maxConsecutiveTeacherPeriods} คาบติดกัน)`);
    }
  });
  return violations;
};

const getTeacherDayShapeScore = (dayPeriods, periodsOrContext) => {
  if (!dayPeriods || dayPeriods.length === 0) return 0;
  const shape = calculateDayTeachingShape(dayPeriods, periodsOrContext);
  let score = shape.gaps * WEIGHTS.TEACHER_GAP;
  if (shape.maxConsec > SCHEDULER_CONFIG.maxConsecutiveTeacherPeriods) {
    score += (shape.maxConsec - SCHEDULER_CONFIG.maxConsecutiveTeacherPeriods) * WEIGHTS.TEACHER_CONSECUTIVE;
  }
  if (shape.gaps === 0 && dayPeriods.length > 1) score += WEIGHTS.NO_GAP;
  return score;
};

const buildPlacementState = (seedSchedules) => {
  const state = { teacherDayPeriods: new Map(), teacherDayCounts: new Map(), roomSubjectDayCounts: new Map() };
  seedSchedules.forEach(sch => {
    if (sch.teacherId) {
      const tdKey = `${sch.teacherId}_${sch.day}`;
      if (!state.teacherDayPeriods.has(tdKey)) state.teacherDayPeriods.set(tdKey, new Set());
      state.teacherDayPeriods.get(tdKey).add(sch.periodId);
      state.teacherDayCounts.set(tdKey, (state.teacherDayCounts.get(tdKey) || 0) + 1);
    }
    const rsdKey = `${sch.classroomId}_${sch.subjectId}_${sch.day}`;
    state.roomSubjectDayCounts.set(rsdKey, (state.roomSubjectDayCounts.get(rsdKey) || 0) + 1);
  });
  return state;
};

const registerPlacementState = (state, scheduleItem) => {
  if (scheduleItem.teacherId) {
    const tdKey = `${scheduleItem.teacherId}_${scheduleItem.day}`;
    if (!state.teacherDayPeriods.has(tdKey)) state.teacherDayPeriods.set(tdKey, new Set());
    state.teacherDayPeriods.get(tdKey).add(scheduleItem.periodId);
    state.teacherDayCounts.set(tdKey, (state.teacherDayCounts.get(tdKey) || 0) + 1);
  }
  const rsdKey = `${scheduleItem.classroomId}_${scheduleItem.subjectId}_${scheduleItem.day}`;
  state.roomSubjectDayCounts.set(rsdKey, (state.roomSubjectDayCounts.get(rsdKey) || 0) + 1);
};

const scorePlacementOption = (entries, load, state, periodsOrContext, getSubject) => {
  if (!entries.length) return -Infinity;
  let score = 0;
  const subject = getSubject(load.subjectId);
  const periodContext = getPeriodContext(periodsOrContext);
  const morningIds = periodContext.morningIds;
  const entriesByDay = new Map();
  entries.forEach(entry => {
    if (!entriesByDay.has(entry.day)) entriesByDay.set(entry.day, []);
    entriesByDay.get(entry.day).push(entry.periodId);
  });

  entriesByDay.forEach((periodIds, day) => {
    if (load.teacherId) {
      const tdKey = `${load.teacherId}_${day}`;
      const oldPeriods = Array.from(state.teacherDayPeriods.get(tdKey) || []);
      const newPeriods = [...oldPeriods, ...periodIds];
      score += getTeacherDayShapeScore(newPeriods, periodContext) - getTeacherDayShapeScore(oldPeriods, periodContext);

      const oldCount = state.teacherDayCounts.get(tdKey) || 0;
      const newCount = oldCount + periodIds.length;
      const oldOver = Math.max(0, oldCount - SCHEDULER_CONFIG.maxTeacherPeriodsPerDay);
      const newOver = Math.max(0, newCount - SCHEDULER_CONFIG.maxTeacherPeriodsPerDay);
      score += (newOver - oldOver) * WEIGHTS.TEACHER_OVERLOAD;

      const dayCounts = DAYS.map(d => state.teacherDayCounts.get(`${load.teacherId}_${d}`) || 0);
      const minDayCount = Math.min(...dayCounts);
      score -= Math.max(0, oldCount - minDayCount) * 3 * periodIds.length;
    }

    const rsdKey = `${load.classroomId}_${load.subjectId}_${day}`;
    const oldSubjectCount = state.roomSubjectDayCounts.get(rsdKey) || 0;
    const newSubjectCount = oldSubjectCount + periodIds.length;
    const oldPenalty = Math.max(0, oldSubjectCount - 2) * WEIGHTS.SUBJECT_SAME_DAY;
    const newPenalty = Math.max(0, newSubjectCount - 2) * WEIGHTS.SUBJECT_SAME_DAY;
    score += newPenalty - oldPenalty;
    if (oldSubjectCount === 0) score += 6;

    if (subject && subject.preferredTime && subject.preferredTime !== 'any') {
      periodIds.forEach(periodId => {
        const isMorning = morningIds.has(periodId);
        const wrongTime = (subject.preferredTime === 'morning' && !isMorning) || (subject.preferredTime === 'afternoon' && isMorning);
        if (wrongTime) score += WEIGHTS.WRONG_TIME_PREFERENCE;
        else score += 2;
      });
    }
  });

  return score;
};

const choosePlacementOption = (options, load, state, periodsOrContext, getSubject, random) => {
  if (!options.length) return null;
  const ranked = options.map(entries => ({
    entries,
    score: scorePlacementOption(entries, load, state, periodsOrContext, getSubject) + (random() * SCHEDULER_CONFIG.placementJitter),
  })).sort((a, b) => b.score - a.score);
  const poolSize = Math.min(SCHEDULER_CONFIG.candidateChoicePool, ranked.length);
  const pickIndex = Math.min(poolSize - 1, Math.floor(Math.pow(random(), 2) * poolSize));
  return ranked[pickIndex]?.entries || ranked[0].entries;
};

const analyzeBlockPairing = (items, periodsOrContext) => {
  const periodContext = getPeriodContext(periodsOrContext);
  const realIndex = periodContext.realIndex;
  const periodMap = periodContext.periodById;
  const byDay = new Map();
  items.forEach(item => {
    if (!item || !item.day || !periodMap.has(item.periodId) || periodMap.get(item.periodId)?.isBreak) return;
    if (!byDay.has(item.day)) byDay.set(item.day, []);
    byDay.get(item.day).push(item);
  });

  const pairs = []; const singles = [];
  byDay.forEach(dayItems => {
    const ordered = [...dayItems].sort((a, b) => (realIndex.get(a.periodId) ?? 999) - (realIndex.get(b.periodId) ?? 999));
    let i = 0;
    while (i < ordered.length) {
      const current = ordered[i];
      const next = ordered[i + 1];
      const currentIdx = realIndex.get(current.periodId);
      const nextIdx = next ? realIndex.get(next.periodId) : null;
      if (next && currentIdx != null && nextIdx === currentIdx + 1) { pairs.push([current, next]); i += 2; }
      else { singles.push(current); i += 1; }
    }
  });
  return { pairs, singles };
};

const getUnpairedFixedSchedules = (fixedItems, periods) => analyzeBlockPairing(fixedItems, periods).singles;

const findBlockPartner = (target, items, periods) => {
  if (!target) return null;
  const analysis = analyzeBlockPairing(items, periods);
  const pair = analysis.pairs.find(([a, b]) => a.id === target.id || b.id === target.id);
  if (!pair) return null;
  return pair[0].id === target.id ? pair[1] : pair[0];
};

const calculateScheduleScore = (schedules, teachers, subjects, periodsOrContext, getSubject) => {
  let score = WEIGHTS.BASE_SCORE;
  const teacherStats = {}; const roomSubjectStats = {};
  const periodContext = getPeriodContext(periodsOrContext);
  const morningPeriods = periodContext.morningIds;

  schedules.forEach(sch => {
    if (sch.teacherId) {
      if (!teacherStats[sch.teacherId]) teacherStats[sch.teacherId] = { days: {} };
      if (!teacherStats[sch.teacherId].days[sch.day]) teacherStats[sch.teacherId].days[sch.day] = [];
      teacherStats[sch.teacherId].days[sch.day].push(sch.periodId);
    }
    const rsKey = `${sch.classroomId}_${sch.subjectId}`;
    if (!roomSubjectStats[rsKey]) roomSubjectStats[rsKey] = { days: {} };
    if (!roomSubjectStats[rsKey].days[sch.day]) roomSubjectStats[rsKey].days[sch.day] = 0;
    roomSubjectStats[rsKey].days[sch.day]++;

    const subj = getSubject(sch.subjectId);
    if (subj && subj.preferredTime && subj.preferredTime !== 'any') {
      const isMorning = morningPeriods.has(sch.periodId);
      if (subj.preferredTime === 'morning' && !isMorning) score += WEIGHTS.WRONG_TIME_PREFERENCE;
      if (subj.preferredTime === 'afternoon' && isMorning) score += WEIGHTS.WRONG_TIME_PREFERENCE;
    }
  });

  Object.values(teacherStats).forEach(tStat => {
    let totalPeriods = 0;
    Object.values(tStat.days).forEach(dayPeriods => {
      totalPeriods += dayPeriods.length;
      const overload = Math.max(0, dayPeriods.length - SCHEDULER_CONFIG.maxTeacherPeriodsPerDay);
      score += overload * WEIGHTS.TEACHER_OVERLOAD;
      score += getTeacherDayShapeScore(dayPeriods, periodContext);
    });
    if (totalPeriods > 0) {
      const avg = totalPeriods / DAYS.length;
      let variance = 0;
      DAYS.forEach(d => { variance += Math.abs((tStat.days[d] ? tStat.days[d].length : 0) - avg); });
      score += Math.max(0, variance - 5) * WEIGHTS.DAILY_IMBALANCE;
    }
  });

  Object.values(roomSubjectStats).forEach(rsStat => {
    const daysUsed = Object.keys(rsStat.days).length;
    const totalClasses = Object.values(rsStat.days).reduce((a, b) => a + b, 0);
    Object.values(rsStat.days).forEach(countInDay => {
      if (countInDay > 2) score += (countInDay - 2) * WEIGHTS.SUBJECT_SAME_DAY;
    });
    if (totalClasses >= 3 && daysUsed >= 3) score += WEIGHTS.GOOD_SPREAD;
  });
  return score;
};

const Modal = ({ isOpen, onClose, title, children }) => {
  if (!isOpen) return null;
  return (
    <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 flex items-center justify-center p-4">
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-lg overflow-hidden animate-in zoom-in-95 duration-200 border border-[#081a39]/10 flex flex-col max-h-[90vh]">
        <div className="flex justify-between items-center p-6 border-b border-[#081a39]/10 bg-[#081a39] text-white shrink-0">
          <h3 className="text-xl font-bold">{title}</h3>
          <button onClick={onClose} className="text-slate-300 hover:text-rose-400 transition-colors bg-white/10 p-2 rounded-full"><X size={20}/></button>
        </div>
        <div className="p-6 overflow-y-auto custom-scrollbar">{children}</div>
      </div>
    </div>
  );
};

const ConfirmModal = ({ confirmData, setConfirmData }) => {
  if (!confirmData) return null;
  return (
    <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[60] flex items-center justify-center p-4">
      <div className="bg-white p-8 rounded-2xl shadow-2xl max-w-sm w-full border border-slate-100 animate-in zoom-in-95 duration-200 text-center">
        <div className={`inline-flex items-center justify-center w-16 h-16 rounded-full mb-6 shadow-inner ${confirmData.type === 'restore' ? 'bg-amber-50 text-amber-500' : 'bg-rose-50 text-rose-500'}`}>
          {confirmData.type === 'restore' ? <UploadCloud size={32} /> : <AlertCircle size={32} />}
        </div>
        <h3 className="text-2xl font-bold text-[#081a39] mb-2">{confirmData.title || 'ยืนยันการทำรายการ'}</h3>
        <p className="text-slate-600 mb-8 whitespace-pre-line text-sm">{confirmData.message || 'คุณต้องการดำเนินการต่อหรือไม่?'}</p>
        <div className="flex justify-center gap-3">
          <button onClick={() => setConfirmData(null)} disabled={confirmData.isLoading} className="flex-1 px-4 py-3 text-[#081a39] bg-slate-100 hover:bg-slate-200 rounded-xl transition-colors font-medium disabled:opacity-50">ยกเลิก</button>
          <button onClick={async () => { await confirmData.action(); }} disabled={confirmData.isLoading} className={`flex-1 px-4 py-3 text-white rounded-xl transition-all shadow-md font-medium disabled:opacity-50 ${confirmData.type === 'restore' ? 'bg-gradient-to-r from-amber-500 to-amber-600 hover:from-amber-600 hover:to-amber-700' : 'bg-gradient-to-r from-rose-500 to-rose-600 hover:from-rose-600 hover:to-rose-700'}`}>
            {confirmData.isLoading ? 'กำลังดำเนินการ...' : (confirmData.confirmText || 'ยืนยัน')}
          </button>
        </div>
      </div>
    </div>
  );
};

const Dashboard = ({ schoolSettings, teachers, activeSubjects, classrooms, health }) => {
  const statusColor = health.status === 'ready' ? 'text-emerald-500' : health.status === 'warning' ? 'text-amber-500' : 'text-rose-500';
  const statusBg = health.status === 'ready' ? 'bg-emerald-500' : health.status === 'warning' ? 'bg-amber-500' : health.status === 'empty' ? 'bg-slate-400' : 'bg-rose-500';
  const statusText = health.status === 'ready' ? 'พร้อมใช้งาน' : health.status === 'warning' ? 'มีคำเตือน' : health.status === 'empty' ? 'ยังไม่มีข้อมูลภาระสอน' : 'มีข้อผิดพลาด (Conflict)';

  return (
    <div className="space-y-6 animate-in fade-in duration-500">
      <div className="bg-gradient-to-br from-[#081a39] via-[#122b5e] to-[#081a39] rounded-3xl p-8 text-white shadow-xl relative overflow-hidden border border-[#d4af37]/30">
        <div className="absolute -top-20 -right-20 p-16 opacity-10 rotate-12 text-[#d4af37]"><Activity size={350} /></div>
        <div className="relative z-10 flex flex-col md:flex-row justify-between md:items-end">
          <div>
            <h1 className="text-4xl font-bold mb-3 tracking-tight">ภาพรวมระบบ</h1>
            <p className="text-[#d4af37] text-xl font-medium flex items-center gap-2">
              <span className="w-2 h-2 rounded-full bg-[#d4af37]"></span>
              {String(schoolSettings?.schoolName || 'โรงเรียน')} | ภาคเรียนที่ {String(schoolSettings?.semester || '1')}/{String(schoolSettings?.academicYear || '2569')}
            </p>
          </div>
          <div className="mt-6 md:mt-0 bg-white/10 backdrop-blur-md p-4 rounded-2xl border border-white/20 text-center min-w-[200px]">
            <div className="text-sm font-bold text-white mb-1 flex justify-center items-center gap-2"><div className={`w-3 h-3 rounded-full ${statusBg} animate-pulse`}></div> สถานะ: {statusText}</div>
            <div className="text-4xl font-bold text-[#d4af37]">{health.completionPercent}%</div>
            <div className="text-xs mt-1 text-slate-300">จัดสำเร็จ {health.scheduledForLoads} / {health.totalRequired} คาบ</div>
          </div>
        </div>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
        <div className="bg-white p-5 rounded-2xl shadow-sm border border-slate-100 flex flex-col justify-between">
           <div className="text-slate-500 font-bold mb-2 flex items-center gap-2"><Briefcase size={18}/> ภาระสอนรวม</div>
           <div className="text-3xl font-bold text-[#081a39]">{health.totalRequired} <span className="text-sm font-normal text-slate-500">คาบ</span></div>
        </div>
        <div className="bg-white p-5 rounded-2xl shadow-sm border border-slate-100 flex flex-col justify-between">
           <div className="text-slate-500 font-bold mb-2 flex items-center gap-2"><CheckCircle2 size={18} className="text-emerald-500"/> จัดตารางแล้ว</div>
           <div className="text-3xl font-bold text-emerald-700">{health.scheduledForLoads} <span className="text-sm font-normal opacity-70">คาบ</span></div>
        </div>
        <div className={`bg-white p-5 rounded-2xl shadow-sm border flex flex-col justify-between ${health.missing > 0 ? 'border-amber-200 bg-amber-50/30' : 'border-slate-100'}`}>
           <div className={`font-bold mb-2 flex items-center gap-2 ${health.missing > 0 ? 'text-amber-600' : 'text-slate-500'}`}><AlertTriangle size={18}/> ยังจัดไม่ครบ</div>
           <div className={`text-3xl font-bold ${health.missing > 0 ? 'text-amber-700' : 'text-[#081a39]'}`}>{health.missing} <span className="text-sm font-normal opacity-70">คาบ</span></div>
        </div>
        <div className={`bg-white p-5 rounded-2xl shadow-sm border flex flex-col justify-between ${health.status === 'error' ? 'border-rose-200 bg-rose-50/30' : 'border-slate-100'}`}>
           <div className={`font-bold mb-2 flex items-center gap-2 ${health.status === 'error' ? 'text-rose-600' : 'text-slate-500'}`}><AlertCircle size={18}/> ข้อขัดแย้ง (Conflicts)</div>
           <div className={`text-3xl font-bold ${health.status === 'error' ? 'text-rose-700' : 'text-[#081a39]'}`}>
             {health.teacherConflicts.length + health.roomConflicts.length + health.unavailConflicts.length + health.cUnavailConflicts.length + health.breakConflicts.length + health.blockConflicts.length + health.teacherLimitConflicts.length + health.orphans.length}
           </div>
        </div>
      </div>
    </div>
  );
};

const SettingsView = ({ schoolSettings, db, user, appId, showToast, getTermKey, fullDataForBackup, handleRestoreBackup }) => {
  const [form, setForm] = useState(schoolSettings);
  const fileInputRef = useRef(null);
  useEffect(() => { setForm(schoolSettings); }, [schoolSettings]);

  const handleSave = async () => {
    if (!user) return;
    try { await setDoc(doc(db, `school_data/${appId}/config/settings`), form); showToast("บันทึกการตั้งค่าแล้ว"); } 
    catch (e) { showToast("เกิดข้อผิดพลาดในการบันทึก", "error"); }
  };

  const handleBackup = () => {
    const backupData = { backupVersion: 2, appId, academicYear: schoolSettings.academicYear, semester: schoolSettings.semester, termKey: getTermKey(schoolSettings), exportedAt: new Date().toISOString(), schoolSettings, ...fullDataForBackup };
    const blob = new Blob([JSON.stringify(backupData, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a"); link.href = url; link.download = `BHS-Backup-${backupData.termKey}-${new Date().toISOString().split('T')[0]}.json`;
    document.body.appendChild(link); link.click(); document.body.removeChild(link); URL.revokeObjectURL(url);
    showToast("ดาวน์โหลดไฟล์สำรองข้อมูลสำเร็จ");
  };

  const handleFileChange = (e) => {
    const file = e.target.files[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = async (event) => {
      try {
        const jsonData = JSON.parse(event.target.result);
        if(!jsonData.backupVersion || !jsonData.appId) throw new Error("Invalid Format");
        handleRestoreBackup(jsonData);
      } catch (err) { showToast("ไฟล์ไม่ถูกต้อง", "error"); }
    };
    reader.readAsText(file); e.target.value = null;
  };

  return (
    <div className="space-y-6">
      <div className="bg-white rounded-2xl shadow-sm border border-slate-100 p-8 max-w-2xl mx-auto animate-in fade-in">
        <h2 className="text-2xl font-bold text-[#081a39] mb-8 flex items-center gap-3"><div className="p-2 bg-[#081a39] rounded-lg shadow-md"><Settings className="text-[#d4af37]"/></div> ตั้งค่าระบบหลัก</h2>
        <div className="space-y-6">
          <div><label className="block text-sm font-bold text-[#081a39] mb-2">ชื่อโรงเรียน</label><input type="text" value={form.schoolName || ''} onChange={e => setForm({...form, schoolName: e.target.value})} className="w-full p-3 border border-slate-200 rounded-xl focus:ring-2 focus:ring-[#081a39] bg-slate-50 focus:bg-white" /></div>
          <div className="grid grid-cols-2 gap-6 p-4 bg-amber-50 rounded-xl border border-amber-100">
            <div className="col-span-2"><p className="text-xs font-bold text-amber-600 mb-2">* ข้อมูลตารางสอนจะแยกตามภาคเรียน หากแก้ไขค่าด้านล่างระบบจะเปลี่ยนไปโหลดข้อมูลใหม่ทันที</p></div>
            <div><label className="block text-sm font-bold text-[#081a39] mb-2">ภาคเรียนที่</label><select value={form.semester || '1'} onChange={e => setForm({...form, semester: e.target.value})} className="w-full p-3 border border-slate-200 rounded-xl focus:ring-2 focus:ring-[#081a39] bg-white font-bold"><option value="1">ภาคเรียนที่ 1</option><option value="2">ภาคเรียนที่ 2</option></select></div>
            <div><label className="block text-sm font-bold text-[#081a39] mb-2">ปีการศึกษา</label><input type="text" value={form.academicYear || ''} onChange={e => setForm({...form, academicYear: e.target.value})} className="w-full p-3 border border-slate-200 rounded-xl focus:ring-2 focus:ring-[#081a39] bg-white font-medium" /></div>
          </div>
          <div><label className="block text-sm font-bold text-[#081a39] mb-2">ผู้อำนวยการโรงเรียน</label><input type="text" value={form.directorName || ''} onChange={e => setForm({...form, directorName: e.target.value})} className="w-full p-3 border border-slate-200 rounded-xl focus:ring-2 focus:ring-[#081a39] bg-slate-50 focus:bg-white" /></div>
          <div><label className="block text-sm font-bold text-[#081a39] mb-2">รองผู้อำนวยการ</label><input type="text" value={form.deputyName || ''} onChange={e => setForm({...form, deputyName: e.target.value})} className="w-full p-3 border border-slate-200 rounded-xl focus:ring-2 focus:ring-[#081a39] bg-slate-50 focus:bg-white" /></div>
          <div className="pt-6 border-t border-slate-100"><button onClick={handleSave} className="w-full py-4 bg-gradient-to-r from-[#081a39] to-[#153063] text-white font-bold text-lg rounded-xl shadow-md hover:shadow-lg transition-all">บันทึกการตั้งค่า</button></div>
        </div>
      </div>
      <div className="bg-white rounded-2xl shadow-sm border border-slate-100 p-8 max-w-2xl mx-auto animate-in fade-in">
        <h2 className="text-xl font-bold text-[#081a39] mb-4 flex items-center gap-2">การจัดการข้อมูลสำรอง (Backup & Restore)</h2>
        <div className="grid grid-cols-2 gap-4">
          <button onClick={handleBackup} className="flex items-center justify-center gap-2 py-3 bg-slate-50 border border-slate-200 text-[#081a39] rounded-xl hover:bg-slate-100 font-bold transition-all shadow-sm"><Download size={18} className="text-emerald-600" /> สำรองข้อมูล (JSON)</button>
          <button onClick={() => fileInputRef.current.click()} className="flex items-center justify-center gap-2 py-3 bg-slate-50 border border-slate-200 text-[#081a39] rounded-xl hover:bg-slate-100 font-bold transition-all shadow-sm"><UploadCloud size={18} className="text-amber-600" /> นำเข้าข้อมูลสำรอง</button>
          <input type="file" accept=".json" ref={fileInputRef} onChange={handleFileChange} className="hidden" />
        </div>
      </div>
    </div>
  );
};

const TeachersView = ({ teachers, dbAction, showToast, handleRequestDelete }) => {
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [form, setForm] = useState({ id: '', name: '', subjectArea: 'ไม่มีกลุ่มสาระการเรียนรู้' });
  const [isEditing, setIsEditing] = useState(false);
  const { items: sortedTeachers, requestSort, sortConfig } = useSortableData(teachers);

  const saveForm = async () => {
    if (!form.id || !form.name) return showToast("กรอกข้อมูลให้ครบ", "error");
    if (!isEditing && teachers.some(t => t.id === form.id)) return showToast("รหัสครูนี้มีอยู่แล้ว", "error");
    const ok = await dbAction('teachers', form.id, form);
    if (ok) { setIsModalOpen(false); showToast(isEditing ? "บันทึกการแก้ไขสำเร็จ" : "เพิ่มข้อมูลครูสำเร็จ"); }
  };

  return (
    <div className="animate-in fade-in space-y-6">
      <div className="flex flex-wrap items-center justify-between gap-4 bg-white p-6 rounded-2xl shadow-sm border border-slate-100">
        <h2 className="text-xl font-bold text-[#081a39] flex items-center gap-2"><Users className="text-[#d4af37]"/> ข้อมูลครู</h2>
        <button onClick={() => {setForm({id:'', name:'', subjectArea:'ไม่มีกลุ่มสาระการเรียนรู้'}); setIsEditing(false); setIsModalOpen(true);}} className="flex items-center gap-2 px-4 py-2 bg-gradient-to-r from-[#081a39] to-[#122b5e] text-white rounded-xl shadow-md font-bold"><Plus size={18} className="text-[#d4af37]"/> เพิ่มครูใหม่</button>
      </div>
      <div className="bg-white rounded-2xl shadow-sm border border-slate-100 overflow-hidden">
        <table className="w-full text-left border-collapse">
          <thead className="bg-[#081a39] text-white border-b-4 border-[#d4af37]">
            <tr>
              <th className="p-4 font-bold cursor-pointer" onClick={() => requestSort('id')}>รหัส <SortIcon column="id" sortConfig={sortConfig}/></th>
              <th className="p-4 font-bold cursor-pointer" onClick={() => requestSort('name')}>ชื่อ-นามสกุล <SortIcon column="name" sortConfig={sortConfig}/></th>
              <th className="p-4 font-bold cursor-pointer" onClick={() => requestSort('subjectArea')}>กลุ่มสาระฯ <SortIcon column="subjectArea" sortConfig={sortConfig}/></th>
              <th className="p-4 font-bold text-center w-28">จัดการ</th>
            </tr>
          </thead>
          <tbody>
            {sortedTeachers.map(t => (
              <tr key={t.id} className="border-b border-slate-50 hover:bg-[#081a39]/5">
                <td className="p-4 text-slate-500 font-mono text-sm">{t.id}</td><td className="p-4 font-medium text-[#081a39]">{t.name}</td><td className={`p-4 font-bold ${SUBJECT_COLORS[t.subjectArea]}`}>{t.subjectArea}</td>
                <td className="p-4 text-center"><button onClick={() => { setForm(t); setIsEditing(true); setIsModalOpen(true); }} className="p-2 text-slate-400 hover:text-[#081a39]"><Edit size={18}/></button><button onClick={() => handleRequestDelete('teacher', t)} className="p-2 text-slate-400 hover:text-rose-500"><Trash2 size={18}/></button></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <Modal isOpen={isModalOpen} onClose={() => setIsModalOpen(false)} title={isEditing ? 'แก้ไขข้อมูลครู' : 'เพิ่มครูใหม่'}>
        <div className="space-y-4">
          <div><label className="block text-sm font-bold text-[#081a39] mb-1">รหัสครู</label><input value={form.id} onChange={e => setForm({...form, id: e.target.value})} disabled={isEditing} className="w-full p-3 border border-slate-200 rounded-xl bg-slate-50 focus:bg-white" /></div>
          <div><label className="block text-sm font-bold text-[#081a39] mb-1">ชื่อ-นามสกุล</label><input value={form.name} onChange={e => setForm({...form, name: e.target.value})} className="w-full p-3 border border-slate-200 rounded-xl bg-slate-50 focus:bg-white" /></div>
          <div><label className="block text-sm font-bold text-[#081a39] mb-1">กลุ่มสาระ</label><select value={form.subjectArea} onChange={e => setForm({...form, subjectArea: e.target.value})} className="w-full p-3 border border-slate-200 rounded-xl bg-slate-50 focus:bg-white">{SUBJECT_AREAS.map(s => <option key={s} value={s}>{s}</option>)}</select></div>
          <button onClick={saveForm} className="w-full mt-4 bg-gradient-to-r from-[#081a39] to-[#122b5e] text-white py-3 rounded-xl font-bold">บันทึก</button>
        </div>
      </Modal>
    </div>
  );
};

// STEP 08: Only subjects without CURRENT-TERM teaching/schedule references are bulk-deletable.
// The master subjects collection is shared across terms. Other-term references are not
// available to this view and must be checked separately before a school-wide cleanup.
const getBulkDeletableSubjectIds = (subjectItems, teachingLoads, schedules, fixedSchedules) => {
  const usedIds = new Set([
    ...teachingLoads.map(item => item.subjectId),
    ...schedules.map(item => item.subjectId),
    ...fixedSchedules.map(item => item.subjectId),
  ]);
  return subjectItems.filter(item => !usedIds.has(item.id)).map(item => item.id);
};

const SubjectsView = ({ activeSubjects, subjects, teachingLoads, schedules, fixedSchedules, bulkDeleteSubjects, renameSubjectId, dbAction, showToast, handleRequestDelete, schoolSettings }) => {
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [form, setForm] = useState({ id: '', name: '', abbr: '', periodsPerWeek: 1, term: '1', periodBlock: 1, preferredTime: 'any' });
  const [isEditing, setIsEditing] = useState(false);
  const [originalSubjectId, setOriginalSubjectId] = useState('');
  const { items: sortedSubjects, requestSort, sortConfig } = useSortableData(activeSubjects);
  const [selectedSubjectIds, setSelectedSubjectIds] = useState([]);
  const deletableIds = useMemo(() => getBulkDeletableSubjectIds(activeSubjects, teachingLoads, schedules, fixedSchedules), [activeSubjects, teachingLoads, schedules, fixedSchedules]);
  const deletableSet = useMemo(() => new Set(deletableIds), [deletableIds]);
  const selectedValidIds = selectedSubjectIds.filter(id => deletableSet.has(id));
  const allSelected = deletableIds.length > 0 && deletableIds.every(id => selectedValidIds.includes(id));

  // Do not carry selection into another term or keep deleted/now-in-use subjects selected.
  useEffect(() => { setSelectedSubjectIds([]); }, [schoolSettings.academicYear, schoolSettings.semester]);
  useEffect(() => { setSelectedSubjectIds(prev => prev.filter(id => deletableSet.has(id))); }, [deletableSet]);
  const toggleSubject = id => setSelectedSubjectIds(prev => prev.includes(id) ? prev.filter(x => x !== id) : [...prev, id]);
  const toggleAllSubjects = () => setSelectedSubjectIds(allSelected ? [] : [...deletableIds]);

  const saveForm = async () => {
    const nextId = String(form.id || '').trim();
    const periodsPerWeek = Number(form.periodsPerWeek);
    if (!nextId || !String(form.name || '').trim()) return showToast("กรอกข้อมูลให้ครบ", "error");
    if (nextId.includes('/') || nextId === '.' || nextId === '..') return showToast("รหัสวิชาไม่สามารถมี / หรือเป็น . หรือ .. ได้", "error");
    if (!Number.isInteger(periodsPerWeek) || periodsPerWeek < 1) return showToast("จำนวนคาบ/สัปดาห์ต้องเป็นจำนวนเต็มบวก", "error");
    if (subjects.some(s => s.id === nextId && (!isEditing || s.id !== originalSubjectId))) return showToast("รหัสวิชานี้มีอยู่แล้ว", "error");
    const updated = { ...form, id: nextId, name: form.name.trim(), periodsPerWeek, periodBlock: Number(form.periodBlock) };
    if (isEditing && originalSubjectId !== nextId) {
      // Changing a subject code is a document-ID migration, not a simple field update.
      // Only close this editor after the confirmed atomic migration succeeds.
      await renameSubjectId(originalSubjectId, updated, () => setIsModalOpen(false));
      return;
    }
    const ok = await dbAction('subjects', nextId, updated);
    if (ok) { setIsModalOpen(false); showToast(isEditing ? "บันทึกการแก้ไขสำเร็จ" : "เพิ่มรายวิชาสำเร็จ"); }
  };

  return (
    <div className="animate-in fade-in space-y-6">
      <div className="flex flex-wrap items-center justify-between gap-4 bg-white p-6 rounded-2xl shadow-sm border border-slate-100">
        <h2 className="text-xl font-bold text-[#081a39] flex items-center gap-2"><BookOpen className="text-[#d4af37]"/> รายวิชา <span className="text-sm font-normal text-slate-500 ml-2">(ภาคเรียนที่ {schoolSettings.semester})</span></h2>
        <button onClick={() => {setForm({id:'', name:'', abbr:'', periodsPerWeek: 1, term: schoolSettings.semester, periodBlock: 1, preferredTime: 'any'}); setIsEditing(false); setOriginalSubjectId(''); setIsModalOpen(true);}} className="flex items-center gap-2 px-4 py-2 bg-gradient-to-r from-[#081a39] to-[#122b5e] text-white rounded-xl shadow-md font-bold"><Plus size={18} className="text-[#d4af37]"/> เพิ่มรายวิชา</button>
      </div>
      <div className="flex flex-wrap items-center gap-3 bg-white border border-slate-100 rounded-xl px-4 py-3">
        <button type="button" onClick={toggleAllSubjects} disabled={deletableIds.length === 0} className="px-3 py-2 rounded-lg border border-slate-200 text-sm font-bold text-[#081a39] disabled:opacity-40 hover:bg-slate-50">{allSelected ? 'ยกเลิกเลือกทั้งหมด' : 'เลือกทั้งหมดที่ลบได้'}</button>
        <span className="text-sm text-slate-600">เลือก {selectedValidIds.length} รายการ</span>
        <button type="button" onClick={() => bulkDeleteSubjects(selectedValidIds)} disabled={selectedValidIds.length === 0} className="ml-auto flex items-center gap-2 px-4 py-2 rounded-lg bg-rose-600 text-white font-bold disabled:opacity-40 hover:bg-rose-700"><Trash2 size={16}/> ลบรายการที่เลือก ({selectedValidIds.length})</button>
        <p className="w-full text-xs text-slate-500">วิชาที่มีภาระสอน ตารางสอน หรือกิจกรรมล็อกในภาคเรียนนี้จะเลือกเพื่อลบไม่ได้ รายวิชาเป็นข้อมูลส่วนกลางที่ใช้ข้ามภาคเรียน</p>
      </div>
      <div className="bg-white rounded-2xl shadow-sm border border-slate-100 overflow-hidden">
        <table className="w-full text-left border-collapse">
          <thead className="bg-[#081a39] text-white border-b-4 border-[#d4af37]">
            <tr>
              <th className="p-4 w-14 text-center"><input type="checkbox" aria-label="เลือกทั้งหมดที่ลบได้" title="เลือกทั้งหมดที่ลบได้" checked={allSelected} disabled={deletableIds.length === 0} onChange={toggleAllSubjects} className="w-4 h-4 accent-[#081a39] cursor-pointer" /></th>
              <th className="p-4 font-bold cursor-pointer" onClick={() => requestSort('id')}>รหัส <SortIcon column="id" sortConfig={sortConfig}/></th>
              <th className="p-4 font-bold cursor-pointer" onClick={() => requestSort('name')}>ชื่อวิชา <SortIcon column="name" sortConfig={sortConfig}/></th>
              <th className="p-4 font-bold text-center">คาบ/สัปดาห์</th><th className="p-4 font-bold text-center w-28">จัดการ</th>
            </tr>
          </thead>
          <tbody>
            {sortedSubjects.map(d => (
              <tr key={d.id} className={`border-b border-slate-50 hover:bg-[#081a39]/5 ${selectedValidIds.includes(d.id) ? 'bg-amber-50/40' : ''}`}>
                <td className="p-4 text-center"><input type="checkbox" aria-label={`เลือกรายวิชา ${d.name} (${d.id})`} title={deletableSet.has(d.id) ? 'เลือกเพื่อลบ' : 'วิชานี้ถูกใช้งานอยู่ ไม่สามารถลบได้'} checked={selectedValidIds.includes(d.id)} disabled={!deletableSet.has(d.id)} onChange={() => toggleSubject(d.id)} className="w-4 h-4 accent-[#081a39] cursor-pointer disabled:cursor-not-allowed" /></td>
                <td className="p-4 text-slate-500 font-mono text-sm">{d.id}</td><td className="p-4 font-medium text-[#081a39]">{d.name} <span className="text-xs text-slate-400">({d.abbr})</span>{!deletableSet.has(d.id) && <span className="ml-2 text-xs text-amber-700">(ถูกใช้งานอยู่)</span>}</td><td className="p-4 text-center font-bold text-[#d4af37]">{d.periodsPerWeek}</td>
                <td className="p-4 text-center"><button onClick={() => { setForm({periodBlock: 1, preferredTime: 'any', ...d}); setOriginalSubjectId(d.id); setIsEditing(true); setIsModalOpen(true); }} className="p-2 text-slate-400 hover:text-[#081a39]"><Edit size={18}/></button><button onClick={() => handleRequestDelete('subject', d)} className="p-2 text-slate-400 hover:text-rose-500"><Trash2 size={18}/></button></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <Modal isOpen={isModalOpen} onClose={() => setIsModalOpen(false)} title={isEditing ? 'แก้ไขรายวิชา' : 'เพิ่มรายวิชา'}>
        <div className="space-y-4">
          <div><label className="block text-sm font-bold text-[#081a39] mb-1">รหัสวิชา</label><input value={form.id} onChange={e => setForm({...form, id: e.target.value})} className="w-full p-3 border border-slate-200 rounded-xl bg-slate-50" /></div>
          {isEditing && originalSubjectId !== String(form.id || '').trim() && <p className="text-xs text-amber-700 bg-amber-50 border border-amber-200 rounded-lg p-3">การเปลี่ยนรหัสวิชาจะขอยืนยันก่อน และอัปเดตภาระสอน ตารางสอน และกิจกรรมล็อกของทุกภาคเรียนพร้อมกัน หากตรวจข้อมูลข้ามภาคเรียนไม่ได้ ระบบจะไม่แก้รหัส</p>}
          <div><label className="block text-sm font-bold text-[#081a39] mb-1">ชื่อวิชา</label><input value={form.name} onChange={e => setForm({...form, name: e.target.value})} className="w-full p-3 border border-slate-200 rounded-xl bg-slate-50" /></div>
          <div className="grid grid-cols-2 gap-4">
            <div><label className="block text-sm font-bold mb-1">ชื่อย่อ</label><input value={form.abbr} onChange={e => setForm({...form, abbr: e.target.value})} className="w-full p-3 border rounded-xl bg-slate-50" /></div>
            <div><label className="block text-sm font-bold mb-1">คาบ/สัปดาห์</label><input type="number" min="1" value={form.periodsPerWeek} onChange={e => setForm({...form, periodsPerWeek: parseInt(e.target.value)})} className="w-full p-3 border rounded-xl bg-slate-50" /></div>
          </div>
          <div className="grid grid-cols-2 gap-4">
            <div><label className="block text-xs font-bold mb-1">คาบต่อเนื่อง (Block)</label><select value={form.periodBlock} onChange={e => setForm({...form, periodBlock: e.target.value})} className="w-full p-3 border rounded-xl bg-slate-50"><option value="1">1 คาบ</option><option value="2">2 คาบติดกัน</option></select></div>
            <div><label className="block text-xs font-bold mb-1">เวลาเหมาะสม</label><select value={form.preferredTime} onChange={e => setForm({...form, preferredTime: e.target.value})} className="w-full p-3 border rounded-xl bg-slate-50"><option value="any">เวลาใดก็ได้</option><option value="morning">ช่วงเช้า</option><option value="afternoon">ช่วงบ่าย</option></select></div>
          </div>
          <div><label className="block text-sm font-bold mb-1">ภาคเรียน</label><select value={form.term} onChange={e => setForm({...form, term: e.target.value})} className="w-full p-3 border rounded-xl bg-slate-50"><option value="1">ภาค 1</option><option value="2">ภาค 2</option><option value="all">ทั้งสองเทอม</option></select></div>
          <button onClick={saveForm} className="w-full mt-4 bg-[#081a39] text-white py-3 rounded-xl font-bold">บันทึก</button>
        </div>
      </Modal>
    </div>
  );
};

const ClassroomsView = ({ classrooms, dbAction, showToast, handleRequestDelete, autoGenerateClassrooms }) => {
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [form, setForm] = useState({ id: '', name: '' });
  const [isEditing, setIsEditing] = useState(false);
  const { items: sortedClassrooms, requestSort, sortConfig } = useSortableData(classrooms);

  const saveForm = async () => {
    if (!form.id || !form.name) return showToast("กรอกข้อมูลให้ครบ", "error");
    if (!isEditing && classrooms.some(c => c.id === form.id)) return showToast("รหัสห้องเรียนนี้มีอยู่แล้ว", "error");
    const ok = await dbAction('classrooms', form.id, form);
    if (ok) { setIsModalOpen(false); showToast(isEditing ? "บันทึกการแก้ไขสำเร็จ" : "เพิ่มห้องเรียนสำเร็จ"); }
  };

  return (
    <div className="animate-in fade-in space-y-6">
      <div className="flex flex-wrap items-center justify-between gap-4 bg-white p-6 rounded-2xl shadow-sm border border-slate-100">
        <h2 className="text-xl font-bold text-[#081a39] flex items-center gap-2"><Home className="text-[#d4af37]"/> ห้องเรียน</h2>
        <div className="flex gap-3">
           <button onClick={autoGenerateClassrooms} className="flex items-center gap-2 px-4 py-2 bg-slate-50 border border-slate-200 text-[#081a39] rounded-xl hover:bg-slate-100 font-medium shadow-sm transition-colors"><Plus size={18} className="text-[#d4af37]" /> สร้างอัตโนมัติ (17 ห้อง)</button>
           <button onClick={() => {setForm({id:'', name:''}); setIsEditing(false); setIsModalOpen(true);}} className="flex items-center gap-2 px-4 py-2 bg-gradient-to-r from-[#081a39] to-[#122b5e] text-white rounded-xl shadow-md font-bold"><Plus size={18} className="text-[#d4af37]"/> เพิ่มห้องเรียน</button>
        </div>
      </div>
      <div className="bg-white rounded-2xl shadow-sm border border-slate-100 overflow-hidden">
        <table className="w-full text-left border-collapse">
          <thead className="bg-[#081a39] text-white border-b-4 border-[#d4af37]">
            <tr>
              <th className="p-4 font-bold cursor-pointer" onClick={() => requestSort('id')}>รหัสอ้างอิง <SortIcon column="id" sortConfig={sortConfig}/></th>
              <th className="p-4 font-bold cursor-pointer" onClick={() => requestSort('name')}>ชื่อห้อง <SortIcon column="name" sortConfig={sortConfig}/></th>
              <th className="p-4 font-bold text-center w-28">จัดการ</th>
            </tr>
          </thead>
          <tbody>
            {sortedClassrooms.map(d => (
              <tr key={d.id} className="border-b border-slate-50 hover:bg-[#081a39]/5">
                <td className="p-4 text-slate-500 font-mono text-sm">{d.id}</td><td className="p-4 font-medium text-[#081a39]">{d.name}</td>
                <td className="p-4 text-center"><button onClick={() => { setForm(d); setIsEditing(true); setIsModalOpen(true); }} className="p-2 text-slate-400 hover:text-[#081a39]"><Edit size={18}/></button><button onClick={() => handleRequestDelete('classroom', d)} className="p-2 text-slate-400 hover:text-rose-500"><Trash2 size={18}/></button></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <Modal isOpen={isModalOpen} onClose={() => setIsModalOpen(false)} title={isEditing ? 'แก้ไขห้องเรียน' : 'เพิ่มห้องเรียน'}>
        <div className="space-y-4">
          <div><label className="block text-sm font-bold text-[#081a39] mb-1">รหัสอ้างอิง</label><input value={form.id} onChange={e => setForm({...form, id: e.target.value})} disabled={isEditing} className="w-full p-3 border border-slate-200 rounded-xl bg-slate-50" /></div>
          <div><label className="block text-sm font-bold text-[#081a39] mb-1">ชื่อห้อง</label><input value={form.name} onChange={e => setForm({...form, name: e.target.value})} className="w-full p-3 border border-slate-200 rounded-xl bg-slate-50" /></div>
          <button onClick={saveForm} className="w-full mt-4 bg-[#081a39] text-white py-3 rounded-xl font-bold">บันทึก</button>
        </div>
      </Modal>
    </div>
  );
};

const LOAD_IMPORT_HEADERS = ['รหัสครู', 'ชื่อครู', 'รหัสวิชา', 'ชื่อวิชา', 'ชื่อห้อง', 'จำนวนคาบ/สัปดาห์'];

const normalizeExcelText = (value) => String(value ?? '').replace(/\s+/g, ' ').trim();

const LoadsView = ({ teachingLoads, teachers, activeSubjects, classrooms, dbAction, bulkUpsertTeachingLoads, showToast, handleRequestDelete, schoolSettings, getTeacher, getSubject, getClassroom }) => {
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [isEditing, setIsEditing] = useState(false);
  const [isImporting, setIsImporting] = useState(false);
  const [form, setForm] = useState({ id: '', teacherId: '', subjectId: '', classroomId: '', periods: 1 });
  const excelInputRef = useRef(null);
  
  const handleSave = async () => {
    const periodsValue = Number(form.periods);
    if (!form.teacherId || !form.subjectId || !form.classroomId || !Number.isInteger(periodsValue) || periodsValue < 1) return showToast("กรุณากรอกข้อมูลให้ครบ และกำหนดจำนวนคาบเป็นจำนวนเต็มบวก", "error");
    const isDuplicate = teachingLoads.some(l => l.teacherId === form.teacherId && l.subjectId === form.subjectId && l.classroomId === form.classroomId && l.id !== form.id );
    if (isDuplicate) return showToast("มีภาระสอนรายการนี้อยู่แล้ว กรุณาแก้ไขจำนวนคาบจากรายการเดิม", "error");

    const id = isEditing && form.id ? form.id : `L_${Date.now()}`;
    const ok = await dbAction('teachingLoads', id, { ...form, periods: periodsValue, id });
    if (ok) { setForm({ id: '', teacherId: '', subjectId: '', classroomId: '', periods: 1 }); setIsModalOpen(false); showToast(isEditing ? "บันทึกการแก้ไขสำเร็จ" : "เพิ่มภาระสอนสำเร็จ"); }
  };

  const downloadExcelTemplate = () => {
    try {
      const workbook = XLSX.utils.book_new();

      const loadSheet = XLSX.utils.aoa_to_sheet([LOAD_IMPORT_HEADERS]);
      loadSheet['!cols'] = [
        { wch: 14 }, { wch: 28 }, { wch: 16 }, { wch: 30 }, { wch: 18 }, { wch: 18 }
      ];
      XLSX.utils.book_append_sheet(workbook, loadSheet, 'ภาระสอน');

      const guideRows = [
        ['แม่แบบนำเข้าภาระสอน BHS Scheduler'],
        ['ภาคเรียน', String(schoolSettings.semester || '')],
        ['ปีการศึกษา', String(schoolSettings.academicYear || '')],
        [],
        ['วิธีกรอก'],
        ['1', 'กรอกข้อมูลในชีต "ภาระสอน" โดยห้ามเปลี่ยนชื่อหัวคอลัมน์'],
        ['2', '1 แถว = ครู 1 คน + วิชา 1 วิชา + ห้องเรียน 1 ห้อง'],
        ['3', 'รหัสครูและชื่อครูต้องตรงกับชีต "รายชื่อครู"'],
        ['4', 'รหัสวิชาและชื่อวิชาต้องตรงกับชีต "รายวิชา" ของภาคเรียนปัจจุบัน'],
        ['5', 'ชื่อห้องต้องตรงกับชีต "ห้องเรียน" (ชื่อห้องในระบบต้องไม่ซ้ำกัน)'],
        ['6', 'จำนวนคาบ/สัปดาห์ต้องเป็นจำนวนเต็มตั้งแต่ 1 ขึ้นไป'],
        ['7', 'ห้ามมีครู+วิชา+ห้องซ้ำกันมากกว่า 1 แถวในไฟล์เดียว'],
        ['8', 'ถ้ารายการมีอยู่แล้ว ระบบจะอัปเดตจำนวนคาบจาก Excel โดยไม่สร้างรายการซ้ำ'],
        [],
        ['ตัวอย่าง'],
        ['T001', 'นายตัวอย่าง ครูผู้สอน', 'ค21101', 'คณิตศาสตร์ 1', 'ม.1/1', '4']
      ];
      const guideSheet = XLSX.utils.aoa_to_sheet(guideRows);
      guideSheet['!cols'] = [{ wch: 18 }, { wch: 80 }, { wch: 18 }, { wch: 28 }, { wch: 18 }, { wch: 18 }];
      XLSX.utils.book_append_sheet(workbook, guideSheet, 'คำแนะนำ');

      const teacherSheet = XLSX.utils.aoa_to_sheet([
        ['รหัสครู', 'ชื่อครู'],
        ...teachers.map(t => [String(t.id || ''), String(t.name || '')])
      ]);
      teacherSheet['!cols'] = [{ wch: 16 }, { wch: 32 }];
      XLSX.utils.book_append_sheet(workbook, teacherSheet, 'รายชื่อครู');

      const subjectSheet = XLSX.utils.aoa_to_sheet([
        ['รหัสวิชา', 'ชื่อวิชา', 'ชื่อย่อ'],
        ...activeSubjects.map(s => [String(s.id || ''), String(s.name || ''), String(s.abbr || '')])
      ]);
      subjectSheet['!cols'] = [{ wch: 18 }, { wch: 36 }, { wch: 16 }];
      XLSX.utils.book_append_sheet(workbook, subjectSheet, 'รายวิชา');

      const roomSheet = XLSX.utils.aoa_to_sheet([
        ['ชื่อห้อง'],
        ...classrooms.map(c => [String(c.name || '')])
      ]);
      roomSheet['!cols'] = [{ wch: 24 }];
      XLSX.utils.book_append_sheet(workbook, roomSheet, 'ห้องเรียน');

      const fileName = `BHS-TeachingLoads-Template-${schoolSettings.academicYear || ''}-${schoolSettings.semester || ''}.xlsx`;
      XLSX.writeFile(workbook, fileName);
      showToast("ดาวน์โหลดแม่แบบ Excel แล้ว");
    } catch (error) {
      console.error(error);
      showToast("สร้างแม่แบบ Excel ไม่สำเร็จ", "error");
    }
  };

  const handleExcelImport = async (event) => {
    const files = Array.from(event.target.files || []);
    event.target.value = '';
    if (files.length === 0) return;

    setIsImporting(true);
    try {
      const collectedRows = [];

      for (const file of files) {
        const buffer = await file.arrayBuffer();
        const workbook = XLSX.read(buffer, { type: 'array' });
        const worksheet = workbook.Sheets['ภาระสอน'] || workbook.Sheets[workbook.SheetNames[0]];
        if (!worksheet) throw new Error(`${file.name}: ไม่พบชีตภาระสอน`);

        const rawMatrix = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '', raw: false });
        const actualHeaders = (rawMatrix[0] || []).map(normalizeExcelText);
        const missingHeaders = LOAD_IMPORT_HEADERS.filter(header => !actualHeaders.includes(header));
        if (missingHeaders.length > 0) {
          throw new Error(`${file.name}: หัวคอลัมน์ไม่ครบ ${missingHeaders.join(', ')}`);
        }

        const headerIndexes = new Map(actualHeaders.map((header, index) => [header, index]));
        rawMatrix.slice(1).forEach((cells, index) => {
          const hasData = LOAD_IMPORT_HEADERS.some(header => normalizeExcelText(cells[headerIndexes.get(header)]) !== '');
          if (hasData) collectedRows.push({ fileName: file.name, cells, rowNumber: index + 2, headerIndexes });
        });
      }

      if (collectedRows.length === 0) {
        showToast("ไม่พบข้อมูลภาระสอนในไฟล์ Excel", "error");
        return;
      }

      const teacherById = new Map(teachers.map(t => [normalizeExcelText(t.id), t]));
      const subjectById = new Map(activeSubjects.map(s => [normalizeExcelText(s.id), s]));
      const roomsByName = new Map();
      classrooms.forEach(c => {
        const name = normalizeExcelText(c.name);
        if (!roomsByName.has(name)) roomsByName.set(name, []);
        roomsByName.get(name).push(c);
      });
      const existingByKey = new Map(teachingLoads.map(load => [getTeachingLoadKey(load), load]));
      const seenKeys = new Map();
      const errors = [];
      const importedItems = [];
      let createCount = 0;
      let updateCount = 0;
      const importToken = Date.now();

      collectedRows.forEach(({ fileName, cells, rowNumber, headerIndexes }, index) => {
        const prefix = `${fileName} แถว ${rowNumber}`;
        const getCell = (header) => normalizeExcelText(cells[headerIndexes.get(header)]);
        const teacherId = getCell('รหัสครู');
        const teacherName = getCell('ชื่อครู');
        const subjectId = getCell('รหัสวิชา');
        const subjectName = getCell('ชื่อวิชา');
        const classroomName = getCell('ชื่อห้อง');
        const periodsText = getCell('จำนวนคาบ/สัปดาห์');
        const periodsValue = Number(periodsText);

        if (!teacherId || !teacherName || !subjectId || !subjectName || !classroomName || !periodsText) {
          errors.push(`${prefix}: กรอกข้อมูลไม่ครบ`);
          return;
        }

        const teacher = teacherById.get(teacherId);
        if (!teacher) {
          errors.push(`${prefix}: ไม่พบรหัสครู ${teacherId}`);
          return;
        }
        if (normalizeExcelText(teacher.name) !== teacherName) {
          errors.push(`${prefix}: ชื่อครูไม่ตรงกับรหัส ${teacherId}`);
          return;
        }

        const subject = subjectById.get(subjectId);
        if (!subject) {
          errors.push(`${prefix}: ไม่พบรหัสวิชา ${subjectId} ในภาคเรียนนี้`);
          return;
        }
        if (normalizeExcelText(subject.name) !== subjectName) {
          errors.push(`${prefix}: ชื่อวิชาไม่ตรงกับรหัส ${subjectId}`);
          return;
        }

        const matchingRooms = roomsByName.get(classroomName) || [];
        if (matchingRooms.length === 0) {
          errors.push(`${prefix}: ไม่พบชื่อห้อง ${classroomName}`);
          return;
        }
        if (matchingRooms.length > 1) {
          errors.push(`${prefix}: ชื่อห้อง ${classroomName} ซ้ำในข้อมูลห้องเรียน กรุณาแก้ชื่อห้องในระบบให้ไม่ซ้ำกัน`);
          return;
        }
        const classroom = matchingRooms[0];
        const classroomId = classroom.id;

        if (!Number.isInteger(periodsValue) || periodsValue < 1) {
          errors.push(`${prefix}: จำนวนคาบ/สัปดาห์ต้องเป็นจำนวนเต็มตั้งแต่ 1 ขึ้นไป`);
          return;
        }

        const key = `${teacherId}_${subjectId}_${classroomId}`;
        if (seenKeys.has(key)) {
          errors.push(`${prefix}: ครู/วิชา/ห้องซ้ำกับ ${seenKeys.get(key)}`);
          return;
        }
        seenKeys.set(key, prefix);

        const existing = existingByKey.get(key);
        const id = existing?.id || `L_IMPORT_${importToken}_${index + 1}`;
        importedItems.push({ id, teacherId, subjectId, classroomId, periods: periodsValue });
        if (existing) updateCount += 1;
        else createCount += 1;
      });

      if (errors.length > 0) {
        const preview = errors.slice(0, 3).join(' | ');
        const more = errors.length > 3 ? ` และอีก ${errors.length - 3} จุด` : '';
        showToast(`นำเข้าไม่ได้ พบ ${errors.length} จุด: ${preview}${more}`, 'error');
        return;
      }

      const ok = await bulkUpsertTeachingLoads(importedItems);
      if (!ok) return;
      showToast(`นำเข้า Excel ${files.length} ไฟล์สำเร็จ ${importedItems.length} รายการ (เพิ่ม ${createCount}, อัปเดต ${updateCount})`);
    } catch (error) {
      console.error(error);
      showToast(`อ่านไฟล์ Excel ไม่สำเร็จ: ${error?.message || 'ไฟล์ไม่ถูกต้อง'}`, 'error');
    } finally {
      setIsImporting(false);
    }
  };

  return (
    <div className="animate-in fade-in space-y-6">
      <div className="flex flex-wrap items-center justify-between gap-4 bg-white p-6 rounded-2xl shadow-sm border border-slate-100">
        <h2 className="text-xl font-bold text-[#081a39] flex items-center gap-2"><Briefcase className="text-[#d4af37]"/> ภาระสอน (เทอม {schoolSettings.semester}/{schoolSettings.academicYear})</h2>
        <div className="flex flex-wrap gap-3">
          <button onClick={downloadExcelTemplate} className="flex items-center gap-2 px-4 py-2 bg-slate-50 border border-slate-200 text-[#081a39] rounded-xl hover:bg-slate-100 font-medium shadow-sm transition-colors"><FileSpreadsheet size={18} className="text-emerald-600" /> ดาวน์โหลดแม่แบบ Excel</button>
          <button onClick={() => excelInputRef.current?.click()} disabled={isImporting} className="flex items-center gap-2 px-4 py-2 bg-slate-50 border border-slate-200 text-[#081a39] rounded-xl hover:bg-slate-100 font-medium shadow-sm transition-colors disabled:opacity-50"><UploadCloud size={18} className="text-amber-600" /> {isImporting ? 'กำลังนำเข้า...' : 'นำเข้า Excel'}</button>
          <input ref={excelInputRef} type="file" accept=".xlsx,.xls" multiple onChange={handleExcelImport} className="hidden" />
          <button onClick={() => { setForm({ id: '', teacherId: '', subjectId: '', classroomId: '', periods: 1 }); setIsEditing(false); setIsModalOpen(true); }} className="flex items-center gap-2 px-4 py-2 bg-gradient-to-r from-[#081a39] to-[#122b5e] text-white rounded-xl shadow-md font-bold"><Plus size={18} className="text-[#d4af37]"/> เพิ่มภาระสอน</button>
        </div>
      </div>
      <div className="bg-white rounded-2xl shadow-sm border border-slate-100 overflow-hidden flex flex-col">
         <div className="overflow-x-auto custom-scrollbar">
           <table className="w-full text-left border-collapse">
              <thead className="bg-[#081a39] text-white border-b-4 border-[#d4af37]">
                <tr><th className="p-4 font-bold">ครูผู้สอน</th><th className="p-4 font-bold">รายวิชา</th><th className="p-4 font-bold text-center">ห้อง</th><th className="p-4 font-bold text-center">คาบ/สัปดาห์</th><th className="p-4 font-bold text-center w-28">จัดการ</th></tr>
              </thead>
              <tbody>
                {teachingLoads.map(l => (
                  <tr key={l.id} className="border-b border-slate-50 hover:bg-[#081a39]/5">
                    <td className="p-4 font-medium text-[#081a39]">{getTeacher(l.teacherId).name}</td><td className="p-4 text-slate-600"><span className="text-xs font-mono bg-slate-100 text-slate-500 px-2 py-1 rounded mr-2">{l.subjectId}</span>{getSubject(l.subjectId).name}</td><td className="p-4 text-center font-bold text-[#d4af37]">{getClassroom(l.classroomId).name}</td><td className="p-4 text-center font-medium">{l.periods}</td>
                    <td className="p-4 text-center"><div className="flex items-center justify-center gap-2"><button onClick={() => { setForm(l); setIsEditing(true); setIsModalOpen(true); }} className="p-2 text-slate-400 hover:text-[#081a39] rounded-lg"><Edit size={18}/></button><button onClick={() => handleRequestDelete('load', l)} className="p-2 text-slate-400 hover:text-rose-500 rounded-lg"><Trash2 size={18}/></button></div></td>
                  </tr>
                ))}
              </tbody>
           </table>
         </div>
      </div>
      <Modal isOpen={isModalOpen} onClose={() => setIsModalOpen(false)} title={isEditing ? "แก้ไขภาระการสอน" : "เพิ่มภาระการสอน"}>
        <div className="space-y-4">
          <div><label className="block text-sm font-bold text-[#081a39] mb-1">ครูผู้สอน</label><select value={form.teacherId} onChange={e => setForm({...form, teacherId: e.target.value})} className="w-full p-3 border border-slate-200 rounded-xl bg-slate-50"><option value="">-- เลือกครูผู้สอน --</option>{teachers.map(t => <option key={t.id} value={t.id}>{t.name}</option>)}</select></div>
          <div><label className="block text-sm font-bold text-[#081a39] mb-1">รายวิชา</label><select value={form.subjectId} onChange={e => setForm({...form, subjectId: e.target.value})} className="w-full p-3 border border-slate-200 rounded-xl bg-slate-50"><option value="">-- เลือกรายวิชา --</option>{activeSubjects.map(s => <option key={s.id} value={s.id}>{s.name} ({s.id})</option>)}</select></div>
          <div><label className="block text-sm font-bold text-[#081a39] mb-1">ห้องเรียน</label><select value={form.classroomId} onChange={e => setForm({...form, classroomId: e.target.value})} className="w-full p-3 border border-slate-200 rounded-xl bg-slate-50"><option value="">-- เลือกห้องเรียน --</option>{classrooms.map(c => <option key={c.id} value={c.id}>{c.name}</option>)}</select></div>
          <div><label className="block text-sm font-bold text-[#081a39] mb-1">จำนวนคาบ/สัปดาห์</label><input type="number" min="1" value={form.periods} onChange={e => setForm({...form, periods: parseInt(e.target.value)})} className="w-full p-3 border border-slate-200 rounded-xl bg-slate-50" /></div>
          <button onClick={handleSave} className="w-full mt-2 bg-[#081a39] text-white py-3 rounded-xl font-bold">บันทึกภาระสอน</button>
        </div>
      </Modal>
    </div>
  );
};

const ConstraintsView = ({ unavailabilities, classroomUnavailabilities, fixedSchedules, schedules, teachers, classrooms, activeSubjects, periods, dbAction, showToast, handleRequestDelete, getTeacher, getClassroom, getSubject }) => {
  const [activeSubTab, setActiveSubTab] = useState('teacher');
  const [tForm, setTForm] = useState({ teacherId: '', day: 'จันทร์', periodId: '1' });
  const [cForm, setCForm] = useState({ classroomId: '', day: 'จันทร์', periodId: '1', reason: '' });
  const [fForm, setFForm] = useState({ subjectId: '', classroomId: '', teacherId: '', day: 'จันทร์', periodId: '1' });

  const handleAdd = async (type) => {
    if (type === 'teacher') {
      if (!tForm.teacherId) return showToast("กรุณาเลือกครู", "error");
      if (unavailabilities.some(u => u.teacherId === tForm.teacherId && u.day === tForm.day && u.periodId === tForm.periodId)) return showToast("มีเวลาที่ไม่สะดวกรายการนี้อยู่แล้ว", "error");
      const id = `U_${Date.now()}`;
      const ok = await dbAction('unavailabilities', id, { ...tForm, id });
      if (ok) showToast("บันทึกเวลาไม่ว่างครูสำเร็จ");
    } else if (type === 'room') {
      if (!cForm.classroomId) return showToast("กรุณาเลือกห้อง", "error");
      if (classroomUnavailabilities.some(u => u.classroomId === cForm.classroomId && u.day === cForm.day && u.periodId === cForm.periodId)) return showToast("มีข้อจำกัดรายการนี้อยู่แล้ว", "error");
      const id = `CU_${Date.now()}`;
      const ok = await dbAction('classroomUnavailabilities', id, { ...cForm, id });
      if (ok) showToast("บันทึกเวลาห้องไม่ว่างสำเร็จ");
    } else if (type === 'fixed') {
      if (!fForm.subjectId || !fForm.classroomId) return showToast("กรุณาเลือกวิชาและห้อง", "error");
      const maps = buildConstraintMaps(schedules, unavailabilities, classroomUnavailabilities, fixedSchedules, periods);
      const targetRooms = fForm.classroomId === ALL_CLASSROOMS_ID ? classrooms : classrooms.filter(c => c.id === fForm.classroomId);
      if (targetRooms.length === 0) return showToast("ยังไม่มีห้องเรียนให้ล็อกกิจกรรม", "error");
      const selectedPeriod = periods.find(p => p.id === fForm.periodId);
      if (!selectedPeriod || selectedPeriod.isBreak) return showToast("ไม่สามารถล็อกกิจกรรมในคาบพักหรือคาบที่ไม่มีอยู่", "error");
      for (const room of targetRooms) {
        const valid = canPlaceSchedule(fForm.teacherId, room.id, fForm.day, fForm.periodId, maps, false);
        if (!valid.ok) {
          const reasons = { "ROOM_UNAVAILABLE": "ห้องไม่ว่างในเวลานี้", "TEACHER_UNAVAILABLE": "ครูไม่สะดวกในเวลานี้", "FIXED_SCHEDULE_CONFLICT": "ห้องนี้มีกิจกรรมล็อกแล้ว", "ROOM_CONFLICT": "ห้องมีตารางอยู่แล้ว", "TEACHER_CONFLICT": "ครูมีตารางอยู่แล้ว", "TEACHER_DAILY_LIMIT": "ครูสอนเกิน 5 คาบต่อวันไม่ได้", "TEACHER_CONSECUTIVE_LIMIT": "ครูสอนเกิน 3 คาบติดต่อกันไม่ได้" };
          return showToast(`${room.name}: ${reasons[valid.reason] || "ไม่สามารถล็อกตารางในเวลานี้ได้"}`, "error");
        }
      }
      const id = `F_${Date.now()}`;
      const ok = await dbAction('fixedSchedules', id, { ...fForm, id });
      if (ok) showToast("บันทึกล็อกตารางสำเร็จ");
    }
  };

  const SubTabBtn = ({ id, label, icon: Icon }) => (
    <button onClick={() => setActiveSubTab(id)} className={`flex items-center gap-2 px-4 py-2 rounded-lg font-bold text-sm transition-all ${activeSubTab === id ? 'bg-[#081a39] text-[#d4af37] shadow-md' : 'bg-white text-slate-500 hover:bg-slate-100 border border-slate-200'}`}><Icon size={16} /> {label}</button>
  );

  return (
    <div className="animate-in fade-in space-y-6">
      <div className="flex gap-3 bg-slate-50 p-2 rounded-xl w-fit border border-slate-200"><SubTabBtn id="teacher" label="เวลาที่ครูไม่สะดวก" icon={Users} /><SubTabBtn id="room" label="เวลาที่ห้องไม่ว่าง" icon={Home} /><SubTabBtn id="fixed" label="กิจกรรมล็อกตายตัว" icon={Pin} /></div>
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        <div className="lg:col-span-1 bg-white rounded-2xl shadow-sm border border-slate-100 p-6 h-fit sticky top-6">
          <h2 className="text-lg font-bold text-[#081a39] mb-4 flex items-center gap-2">เพิ่มข้อมูลใหม่</h2>
          {activeSubTab === 'teacher' && (
            <div className="space-y-4">
              <select value={tForm.teacherId} onChange={e => setTForm({...tForm, teacherId: e.target.value})} className="w-full p-3 border rounded-xl bg-slate-50"><option value="">-- เลือกครูผู้สอน --</option>{teachers.map(t => <option key={t.id} value={t.id}>{t.name}</option>)}</select>
              <div className="grid grid-cols-2 gap-3"><select value={tForm.day} onChange={e => setTForm({...tForm, day: e.target.value})} className="w-full p-3 border rounded-xl bg-slate-50">{DAYS.map(d => <option key={d} value={d}>{d}</option>)}</select><select value={tForm.periodId} onChange={e => setTForm({...tForm, periodId: e.target.value})} className="w-full p-3 border rounded-xl bg-slate-50">{periods.filter(p=>!p.isBreak).map(p => <option key={p.id} value={p.id}>{p.name}</option>)}</select></div>
              <button onClick={() => handleAdd('teacher')} className="w-full bg-[#081a39] text-white py-3 rounded-xl font-bold">บันทึก</button>
            </div>
          )}
          {activeSubTab === 'room' && (
            <div className="space-y-4">
              <select value={cForm.classroomId} onChange={e => setCForm({...cForm, classroomId: e.target.value})} className="w-full p-3 border rounded-xl bg-slate-50"><option value="">-- เลือกห้องเรียน --</option>{classrooms.map(t => <option key={t.id} value={t.id}>{t.name}</option>)}</select>
              <div className="grid grid-cols-2 gap-3"><select value={cForm.day} onChange={e => setCForm({...cForm, day: e.target.value})} className="w-full p-3 border rounded-xl bg-slate-50">{DAYS.map(d => <option key={d} value={d}>{d}</option>)}</select><select value={cForm.periodId} onChange={e => setCForm({...cForm, periodId: e.target.value})} className="w-full p-3 border rounded-xl bg-slate-50">{periods.filter(p=>!p.isBreak).map(p => <option key={p.id} value={p.id}>{p.name}</option>)}</select></div>
              <input type="text" placeholder="เหตุผล (เช่น กิจกรรมระดับชั้น)" value={cForm.reason} onChange={e => setCForm({...cForm, reason: e.target.value})} className="w-full p-3 border rounded-xl bg-slate-50" />
              <button onClick={() => handleAdd('room')} className="w-full bg-[#081a39] text-white py-3 rounded-xl font-bold">บันทึก</button>
            </div>
          )}
          {activeSubTab === 'fixed' && (
            <div className="space-y-4">
              <select value={fForm.subjectId} onChange={e => setFForm({...fForm, subjectId: e.target.value})} className="w-full p-3 border rounded-xl bg-slate-50"><option value="">-- เลือกวิชา / กิจกรรม --</option>{activeSubjects.map(s => <option key={s.id} value={s.id}>{s.name}</option>)}</select>
              <select value={fForm.classroomId} onChange={e => setFForm({...fForm, classroomId: e.target.value})} className="w-full p-3 border rounded-xl bg-slate-50"><option value="">-- เลือกห้องเรียน --</option><option value={ALL_CLASSROOMS_ID}>ทุกห้องเรียน</option>{classrooms.map(c => <option key={c.id} value={c.id}>{c.name}</option>)}</select>
              {fForm.classroomId === ALL_CLASSROOMS_ID && <p className="text-xs text-amber-700">กิจกรรมนี้จะปรากฏในตารางทุกห้อง หากห้องใดมีคาบอื่นอยู่แล้ว ระบบจะไม่บันทึก</p>}
              <select value={fForm.teacherId} onChange={e => setFForm({...fForm, teacherId: e.target.value})} className="w-full p-3 border rounded-xl bg-slate-50"><option value="">-- ครูที่ดูแล (ไม่ระบุได้) --</option>{teachers.map(t => <option key={t.id} value={t.id}>{t.name}</option>)}</select>
              <div className="grid grid-cols-2 gap-3"><select value={fForm.day} onChange={e => setFForm({...fForm, day: e.target.value})} className="w-full p-3 border rounded-xl bg-slate-50">{DAYS.map(d => <option key={d} value={d}>{d}</option>)}</select><select value={fForm.periodId} onChange={e => setFForm({...fForm, periodId: e.target.value})} className="w-full p-3 border rounded-xl bg-slate-50">{periods.filter(p=>!p.isBreak).map(p => <option key={p.id} value={p.id}>{p.name}</option>)}</select></div>
              <button onClick={() => handleAdd('fixed')} className="w-full bg-[#081a39] text-white py-3 rounded-xl font-bold">บันทึก</button>
            </div>
          )}
        </div>
        <div className="lg:col-span-2 bg-white rounded-2xl shadow-sm border border-slate-100 overflow-hidden flex flex-col min-h-[400px]">
           <table className="w-full text-left border-collapse">
              <thead className="bg-[#081a39] text-white border-b-4 border-[#d4af37]"><tr><th className="p-4 font-bold">ข้อมูลรายการ</th><th className="p-4 font-bold text-center">วัน</th><th className="p-4 font-bold text-center">คาบที่</th><th className="p-4 font-bold text-center w-20">ลบ</th></tr></thead>
              <tbody>
                {activeSubTab === 'teacher' && unavailabilities.map(u => (
                  <tr key={u.id} className="border-b border-slate-50 hover:bg-[#081a39]/5"><td className="p-4 font-medium text-[#081a39]">{getTeacher(u.teacherId).name}</td><td className="p-4 text-center">{u.day}</td><td className="p-4 text-center font-bold text-rose-500">{getPeriodDisplayName(periods, u.periodId)}</td><td className="p-4 text-center"><button onClick={() => handleRequestDelete('unavail', u)} className="p-2 text-slate-400 hover:text-rose-500"><Trash2 size={18}/></button></td></tr>
                ))}
                {activeSubTab === 'room' && classroomUnavailabilities.map(u => (
                  <tr key={u.id} className="border-b border-slate-50 hover:bg-[#081a39]/5"><td className="p-4 font-medium text-[#081a39]">{getClassroom(u.classroomId).name} <span className="text-xs text-slate-400 ml-2">{u.reason}</span></td><td className="p-4 text-center">{u.day}</td><td className="p-4 text-center font-bold text-rose-500">{getPeriodDisplayName(periods, u.periodId)}</td><td className="p-4 text-center"><button onClick={() => handleRequestDelete('cUnavail', u)} className="p-2 text-slate-400 hover:text-rose-500"><Trash2 size={18}/></button></td></tr>
                ))}
                {activeSubTab === 'fixed' && fixedSchedules.map(f => (
                  <tr key={f.id} className="border-b border-slate-50 hover:bg-[#081a39]/5"><td className="p-4 font-medium text-[#081a39]"><span className="text-amber-600 font-bold">{getSubject(f.subjectId).name}</span> <span className="text-xs text-slate-500 ml-2">ห้อง: {getClassroom(f.classroomId).name}</span></td><td className="p-4 text-center">{f.day}</td><td className="p-4 text-center font-bold text-rose-500">{getPeriodDisplayName(periods, f.periodId)}</td><td className="p-4 text-center"><button onClick={() => handleRequestDelete('fixed', f)} className="p-2 text-slate-400 hover:text-rose-500"><Trash2 size={18}/></button></td></tr>
                ))}
              </tbody>
           </table>
        </div>
      </div>
    </div>
  );
};

const ScheduleView = ({ schedules, teachingLoads, classrooms, teachers, activeSubjects, periods, unavailabilities, classroomUnavailabilities, fixedSchedules, getSubject, getShortTeacherName, getClassroom, getTeacher, runSmartSchedule, dbAction, showToast, handleRequestDelete, isScheduling, db, appId, getTermKey, schoolSettings, saveScheduleVersion }) => {
  const [viewMode, setViewMode] = useState('room');
  const [selectedId, setSelectedId] = useState('');
  const [isAddModalOpen, setIsAddModalOpen] = useState(false);
  const [addForm, setAddForm] = useState({ teacherId: '', subjectId: '', classroomId: '', day: 'จันทร์', periodId: '1' });

  // จำกัดรายวิชาในหน้าต่างเพิ่มตารางสอนด้วยมือ ให้ตรงตามภาระสอนของครู/ห้องที่เลือก
  const availableManualSubjects = useMemo(() => {
    if (!addForm.teacherId) return [];

    const relevantLoads = teachingLoads.filter(load =>
      load.teacherId === addForm.teacherId &&
      (!addForm.classroomId || load.classroomId === addForm.classroomId)
    );
    const loadBySubject = new Map(relevantLoads.map(load => [load.subjectId, load]));

    return activeSubjects.filter(subject => loadBySubject.has(subject.id)).map(subject => {
      const load = loadBySubject.get(subject.id);
      // ยังไม่ได้เลือกห้อง: แสดงเฉพาะวิชาที่ครูมีภาระสอน แต่ไม่รวมยอดข้ามห้อง
      if (!addForm.classroomId) return { ...subject, optionLabel: `${subject.name} (${subject.id})` };

      const isSameLoad = item => item.teacherId === addForm.teacherId &&
        item.subjectId === subject.id && item.classroomId === addForm.classroomId;
      const assigned = schedules.filter(isSameLoad).length + fixedSchedules.filter(isSameLoad).length;
      const required = Number(load.periods) || 0;
      const remaining = Math.max(0, required - assigned);
      const isComplete = remaining === 0;
      return { ...subject, isComplete, optionLabel: isComplete
        ? `${subject.name} (${subject.id}) — จัดครบแล้ว (${required}/${required} คาบ)`
        : `${subject.name} (${subject.id}) — เหลือ ${remaining}/${required} คาบ` };
    });
  }, [addForm.teacherId, addForm.classroomId, teachingLoads, activeSubjects, schedules, fixedSchedules]);
  const selectableManualSubjects = availableManualSubjects.filter(subject => !subject.isComplete);

  const handleCellClick = (day, periodId) => {
    if (!selectedId) return;
    setAddForm({ teacherId: viewMode === 'teacher' ? selectedId : '', subjectId: '', classroomId: viewMode === 'room' ? selectedId : '', day, periodId });
    setIsAddModalOpen(true);
  };

  const handleManualAdd = async () => {
    if (!addForm.teacherId || !addForm.subjectId || !addForm.classroomId) return showToast("กรุณากรอกข้อมูลให้ครบ", "error");
    if (!DAYS.includes(addForm.day)) return showToast("วันที่เลือกไม่ถูกต้อง", "error");

    const selectedPeriod = periods.find(p => p.id === addForm.periodId);
    if (!selectedPeriod || selectedPeriod.isBreak) return showToast("คาบที่เลือกไม่สามารถจัดการเรียนได้", "error");

    const matchingLoad = teachingLoads.find(l => l.teacherId === addForm.teacherId && l.subjectId === addForm.subjectId && l.classroomId === addForm.classroomId);
    if (!matchingLoad) return showToast("ไม่มีภาระสอนของครู/วิชา/ห้องรายการนี้", "error");

    const required = Math.max(0, Number(matchingLoad.periods) || 0);
    if (required < 1) return showToast("จำนวนคาบของภาระสอนไม่ถูกต้อง", "error");

    const sameLoadSchedules = schedules.filter(s => s.teacherId === addForm.teacherId && s.subjectId === addForm.subjectId && s.classroomId === addForm.classroomId);
    const sameLoadFixed = fixedSchedules.filter(f => f.teacherId === addForm.teacherId && f.subjectId === addForm.subjectId && f.classroomId === addForm.classroomId);
    const assignedItems = [...sameLoadFixed, ...sameLoadSchedules];
    const remaining = required - assignedItems.length;
    if (remaining <= 0) return showToast("ภาระสอนรายการนี้จัดครบแล้ว", "error");

    const maps = buildConstraintMaps(schedules, unavailabilities, classroomUnavailabilities, fixedSchedules, periods);
    const selectedValid = canPlaceSchedule(addForm.teacherId, addForm.classroomId, addForm.day, addForm.periodId, maps, false);
    if (!selectedValid.ok) {
      const reasons = { "BREAK_PERIOD": "จัดลงคาบพักเที่ยงไม่ได้", "ROOM_UNAVAILABLE": "ห้องไม่สะดวก", "TEACHER_UNAVAILABLE": "ครูไม่สะดวก", "FIXED_SCHEDULE_CONFLICT": "มีกิจกรรมล็อกแล้ว", "ROOM_CONFLICT": "ห้องมีวิชาอื่นแล้ว", "TEACHER_CONFLICT": "ครูมีสอนแล้ว", "TEACHER_DAILY_LIMIT": "ครูสอนเกิน 5 คาบต่อวันไม่ได้", "TEACHER_CONSECUTIVE_LIMIT": "ครูสอนเกิน 3 คาบติดต่อกันไม่ได้" };
      return showToast(reasons[selectedValid.reason] || "ติดเงื่อนไขจัดไม่ได้", "error");
    }

    const subject = getSubject(addForm.subjectId);
    const blockSize = Number(subject?.periodBlock) === 2 ? 2 : 1;
    let plannedEntries = [{ day: addForm.day, periodId: addForm.periodId }];

    if (blockSize === 2) {
      const realIndex = new Map(periods.map((p, index) => [p.id, index]));
      const selectedIndex = realIndex.get(addForm.periodId);
      const pairing = analyzeBlockPairing(assignedItems, periods);
      const adjacentUnpaired = pairing.singles.find(item => item.day === addForm.day && Math.abs((realIndex.get(item.periodId) ?? -999) - selectedIndex) === 1);

      if (!adjacentUnpaired) {
        const allowedSingles = required % 2;
        if (remaining === 1) {
          if (pairing.singles.length >= allowedSingles) return showToast("วิชานี้ต้องจัดเป็น 2 คาบติดกัน กรุณาเลือกคาบที่ติดกับคาบเดิมที่ยังไม่มีคู่", "error");
        } else {
          const neighborPeriods = [selectedIndex - 1, selectedIndex + 1]
            .map(index => periods[index])
            .filter(p => p && !p.isBreak);

          const candidatePairs = [];
          neighborPeriods.forEach(neighbor => {
            const neighborValid = canPlaceSchedule(addForm.teacherId, addForm.classroomId, addForm.day, neighbor.id, maps, false);
            if (neighborValid.ok) {
              const pair = [{ day: addForm.day, periodId: addForm.periodId }, { day: addForm.day, periodId: neighbor.id }];
              if (canPlaceTeacherEntries(addForm.teacherId, pair, maps).ok) candidatePairs.push(pair);
            }
          });

          if (candidatePairs.length === 0) return showToast("วิชานี้กำหนดเป็น Block 2 แต่ไม่มีคาบข้างเคียงที่ว่างสำหรับจัดติดกัน", "error");

          const placementState = buildPlacementState([...fixedSchedules, ...schedules]);
          candidatePairs.sort((a, b) => scorePlacementOption(b, matchingLoad, placementState, periods, getSubject) - scorePlacementOption(a, matchingLoad, placementState, periods, getSubject));
          plannedEntries = candidatePairs[0];
        }
      }

      if (plannedEntries.length > remaining) return showToast("จำนวนคาบที่เหลือไม่พอสำหรับ Block 2", "error");

      if (assignedItems.length + plannedEntries.length >= required) {
        const projected = [...assignedItems, ...plannedEntries.map((entry, index) => ({ id: `preview_${index}`, ...entry }))];
        const projectedSingles = analyzeBlockPairing(projected, periods).singles.length;
        if (projectedSingles > (required % 2)) return showToast("รูปแบบคาบต่อเนื่องยังไม่สมบูรณ์ กรุณาจัดคาบให้เป็นคู่ติดกัน", "error");
      }
    }

    const teacherLimit = canPlaceTeacherEntries(addForm.teacherId, plannedEntries, maps);
    if (!teacherLimit.ok) return showToast(teacherLimit.reason === 'TEACHER_DAILY_LIMIT'
      ? 'ครูสอนเกิน 5 คาบต่อวันไม่ได้' : 'ครูสอนเกิน 3 คาบติดต่อกันไม่ได้', 'error');

    try {
      if (saveScheduleVersion) await saveScheduleVersion('before_manual_add');
      const batch = writeBatch(db);
      const termPath = `school_data/${appId}/terms/${getTermKey(schoolSettings)}`;
      plannedEntries.forEach(entry => {
        const scheduleRef = doc(collection(db, `${termPath}/schedules`));
        batch.set(scheduleRef, {
          id: scheduleRef.id, teacherId: addForm.teacherId, subjectId: addForm.subjectId, classroomId: addForm.classroomId, day: entry.day, periodId: entry.periodId, source: 'manual'
        });
      });
      await batch.commit();
      setIsAddModalOpen(false);
      showToast(plannedEntries.length === 2 ? "เพิ่มคาบเรียนต่อเนื่อง 2 คาบสำเร็จ" : "เพิ่มคาบเรียนสำเร็จ");
    } catch (error) {
      console.error(error);
      showToast("บันทึกคาบเรียนไม่สำเร็จ", "error");
    }
  };

  const getCellContent = (day, periodId) => {
     if (!selectedId) return null;
     const cellSch = schedules.filter(s => viewMode === 'room' ? s.classroomId === selectedId && s.day === day && s.periodId === periodId : s.teacherId === selectedId && s.day === day && s.periodId === periodId);
     const cellFix = fixedSchedules.filter(f => viewMode === 'room' ? fixedAppliesToClassroom(f, selectedId) && f.day === day && f.periodId === periodId : f.teacherId === selectedId && f.day === day && f.periodId === periodId);
     if (cellSch.length === 0 && cellFix.length === 0) return null;
     const isConflict = cellSch.length + cellFix.length > 1;

     return (
       <div className={`p-1.5 h-full w-full rounded-xl flex flex-col items-center justify-center text-xs text-center border-l-4 shadow-sm relative group ${isConflict ? 'bg-rose-50 border-rose-500' : 'bg-white border-[#d4af37]'}`}>
         {isConflict && <div className="absolute -top-2 -right-2 text-rose-500 bg-rose-100 rounded-full p-0.5 shadow-sm"><AlertCircle size={14}/></div>}
         {cellFix.map((f, idx) => (
           <div key={f.id} className={idx > 0 || cellSch.length > 0 ? "mt-2 pt-2 border-t w-full relative" : "w-full relative"}>
             <div className="font-bold text-amber-700 flex items-center justify-center gap-1"><Pin size={10}/> {getSubject(f.subjectId).name}</div>
             <button onClick={(e) => { e.stopPropagation(); handleRequestDelete('fixed', f); }} className="absolute -top-1 -right-1 opacity-0 group-hover:opacity-100 bg-rose-500 text-white rounded-full p-0.5 z-10"><X size={12}/></button>
           </div>
         ))}
         {cellSch.map((sch, idx) => (
           <div key={sch.id} className={idx > 0 || cellFix.length > 0 ? "mt-2 pt-2 border-t w-full relative" : "w-full relative"}>
             <div className="font-bold text-[#081a39] truncate">{getSubject(sch.subjectId).name}</div>
             {viewMode === 'room' ? (
               <><div className="text-slate-600 truncate">{getShortTeacherName(getTeacher(sch.teacherId).name)}</div></>
             ) : (
               <><div className="text-slate-600 truncate">{getClassroom(sch.classroomId).name}</div></>
             )}
             <button onClick={(e) => { e.stopPropagation(); handleRequestDelete('schedule', sch); }} className="absolute -top-1 -right-1 opacity-0 group-hover:opacity-100 bg-rose-500 text-white rounded-full p-0.5 z-10"><X size={12}/></button>
           </div>
         ))}
       </div>
     );
  };

  return (
    <div className="space-y-6 animate-in fade-in">
      <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-100 flex flex-wrap items-center justify-between gap-4">
        <div className="flex items-center gap-4">
          <select value={viewMode} onChange={e => {setViewMode(e.target.value); setSelectedId('');}} className="p-3 border border-slate-200 rounded-xl bg-slate-50 font-bold text-[#081a39]"><option value="room">ตารางเรียน (รายห้อง)</option><option value="teacher">ตารางสอน (รายครู)</option></select>
          <select value={selectedId} onChange={e => setSelectedId(e.target.value)} className="p-3 border border-slate-200 rounded-xl min-w-[250px] bg-slate-50 font-medium"><option value="">-- เลือกเพื่อแสดงตาราง --</option>{viewMode === 'room' ? classrooms.map(c => <option key={c.id} value={c.id}>{c.name}</option>) : teachers.map(t => <option key={t.id} value={t.id}>{t.name}</option>)}</select>
        </div>
        <button onClick={runSmartSchedule} disabled={isScheduling} className={`flex items-center gap-2 px-6 py-3 text-white rounded-xl shadow-md font-bold transition-all ${isScheduling ? 'bg-slate-400 cursor-not-allowed' : 'bg-gradient-to-r from-[#d4af37] to-[#b8952b] hover:shadow-lg'}`}><RefreshCw size={18} className={isScheduling ? "animate-spin" : ""} /> {isScheduling ? "กำลังจัดตาราง (AI)..." : "จัดตารางอัตโนมัติ"}</button>
      </div>
      <div className="bg-white rounded-2xl shadow-sm border border-slate-100 overflow-x-auto p-4 min-h-[500px]">
        <table className="w-full text-left border-collapse min-w-[900px] table-fixed">
          <thead className="bg-[#081a39] text-white rounded-t-xl border-b-4 border-[#d4af37]"><tr><th className="p-4 w-28 text-center font-bold">วัน/เวลา</th>{periods.map(p => <th key={p.id} className="p-2 border-l border-[#122b5e] text-center text-sm"><div className="font-bold">{p.name}</div><div className="text-[#d4af37] text-xs font-normal opacity-90">{p.start}-{p.end}</div></th>)}</tr></thead>
          <tbody>
            {DAYS.map(day => (
              <tr key={day} className="border-b border-slate-100 group"><td className="p-4 font-bold text-center bg-slate-50 border-r border-slate-100 text-[#081a39]">{day}</td>{periods.map(p => {
                  const content = getCellContent(day, p.id);
                  return (
                    <td key={`${day}-${p.id}`} className={`p-1.5 border-r border-slate-100 h-28 align-top relative group/cell ${p.isBreak ? 'bg-slate-50/80' : (selectedId ? 'hover:bg-slate-50 cursor-pointer' : '')}`} onClick={() => { if (!p.isBreak && selectedId) handleCellClick(day, p.id); }}>
                      {p.isBreak ? <div className="h-full flex items-center justify-center text-slate-400 text-sm font-bold opacity-60">พักเที่ยง</div> : <>{content}{!content && selectedId && !isScheduling && <div className="absolute inset-0 flex items-center justify-center opacity-0 group-hover/cell:opacity-100 transition-opacity"><div className="bg-[#081a39]/5 p-2 rounded-full text-[#081a39]/40"><Plus size={24} /></div></div>}</>}
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <Modal isOpen={isAddModalOpen} onClose={() => setIsAddModalOpen(false)} title="เพิ่มตารางสอน (รายคาบ)">
        <div className="space-y-4">
          <div className="grid grid-cols-2 gap-4"><div><label className="block text-sm font-bold mb-1">วัน</label><input value={addForm.day} disabled className="w-full p-3 border rounded-xl bg-slate-100" /></div><div><label className="block text-sm font-bold mb-1">คาบที่</label><input value={periods.find(p=>p.id===addForm.periodId)?.name} disabled className="w-full p-3 border rounded-xl bg-slate-100" /></div></div>
          <div><label className="block text-sm font-bold mb-1">ครูผู้สอน</label><select value={addForm.teacherId} onChange={e => setAddForm(prev => ({ ...prev, teacherId: e.target.value, subjectId: '' }))} disabled={viewMode === 'teacher'} className="w-full p-3 border rounded-xl bg-slate-50 disabled:bg-slate-100"><option value="">-- เลือกครูผู้สอน --</option>{teachers.map(t => <option key={t.id} value={t.id}>{t.name}</option>)}</select></div>
          <div><label className="block text-sm font-bold mb-1">รายวิชา</label><select value={addForm.subjectId} onChange={e => setAddForm({...addForm, subjectId: e.target.value})} disabled={!addForm.teacherId || selectableManualSubjects.length === 0} className="w-full p-3 border rounded-xl bg-slate-50 disabled:bg-slate-100"><option value="">{!addForm.teacherId ? '-- กรุณาเลือกครูผู้สอนก่อน --' : availableManualSubjects.length === 0 ? '-- ไม่พบรายวิชาในภาระสอน --' : selectableManualSubjects.length === 0 ? '-- รายวิชาของครู/ห้องนี้จัดครบแล้ว --' : '-- เลือกรายวิชา --'}</option>{availableManualSubjects.map(s => <option key={s.id} value={s.id} disabled={!!s.isComplete}>{s.optionLabel}</option>)}</select>{addForm.teacherId && availableManualSubjects.length === 0 && <p role="status" className="text-sm text-amber-700 mt-2">{addForm.classroomId ? 'ไม่พบภาระสอนสำหรับครูและห้องเรียนนี้' : 'ไม่พบภาระสอนสำหรับครูคนนี้'} กรุณาตรวจสอบเมนูภาระสอน</p>}{addForm.teacherId && addForm.classroomId && availableManualSubjects.length > 0 && selectableManualSubjects.length === 0 && <p role="status" className="text-sm text-emerald-700 mt-2">รายวิชาในภาระสอนของครูและห้องนี้จัดครบทุกคาบแล้ว</p>}</div>
          <div><label className="block text-sm font-bold mb-1">ห้องเรียน</label><select value={addForm.classroomId} onChange={e => setAddForm(prev => ({ ...prev, classroomId: e.target.value, subjectId: '' }))} disabled={viewMode === 'room'} className="w-full p-3 border rounded-xl bg-slate-50 disabled:bg-slate-100"><option value="">-- เลือกห้องเรียน --</option>{classrooms.map(c => <option key={c.id} value={c.id}>{c.name}</option>)}</select></div>
          <button onClick={handleManualAdd} className="w-full mt-2 bg-[#081a39] text-white py-3 rounded-xl font-bold">บันทึกคาบเรียน</button>
        </div>
      </Modal>
    </div>
  );
};

const ReportsView = ({ health, teachers, subjects, classrooms, periods, schoolSettings, schedules, fixedSchedules, getSubject, getTeacher, getClassroom }) => {
  const [activeSubTab, setActiveSubTab] = useState('health');
  const [filterLvl, setFilterLvl] = useState('all');

  const filteredMissing = useMemo(() => {
    let list = [...health.missingLoads];
    if (filterLvl !== 'all') list = list.filter(m => getClassroom(m.classroomId).name.includes(filterLvl));
    return list.sort((a,b) => b.missing - a.missing);
  }, [health.missingLoads, filterLvl]);

  const handleExportCSV = (type) => {
    try {
      const termSuffix = `${schoolSettings.academicYear}-T${schoolSettings.semester}`;
      const downloadCSV = (csvContent, fileName) => {
        const blob = new Blob(["\uFEFF" + csvContent], { type: 'text/csv;charset=utf-8;' });
        const url = URL.createObjectURL(blob); const link = document.createElement("a"); link.href = url; link.download = fileName;
        document.body.appendChild(link); link.click(); document.body.removeChild(link); URL.revokeObjectURL(url);
      };
      const aoaToCSV = (aoa) => aoa.map(row => row.map(cell => {
        const safeCell = sanitizeExcelInput(cell);
        return `"${String(safeCell).replace(/"/g, '""')}"`;
      }).join(',')).join('\n');

      if (type === 'allRooms' || type === 'all') {
        const activePs = periods.filter(p=>!p.isBreak);
        const wsData = [['วัน', 'ห้อง', ...activePs.map(p=>p.name)]];
        DAYS.forEach(day => {
          classrooms.forEach(c => {
             const row = [day, c.name];
             activePs.forEach(p => {
               const sch = schedules.filter(s => s.classroomId === c.id && s.day === day && s.periodId === p.id);
               const fixed = fixedSchedules.filter(f => fixedAppliesToClassroom(f, c.id) && f.day === day && f.periodId === p.id);
               const combined = [...fixed, ...sch];
               row.push(combined.map(s => getSubject(s.subjectId).name).join(' + '));
             });
             wsData.push(row);
          }); wsData.push([]); 
        });
        downloadCSV(aoaToCSV(wsData), `BHS-All-Classrooms-${termSuffix}.csv`);
      }
    } catch (e) { console.error(e); }
  };

  return (
    <div className="animate-in fade-in space-y-6">
      <div className="bg-white p-4 rounded-2xl shadow-sm border border-slate-100 flex flex-wrap gap-2">
        <button onClick={() => setActiveSubTab('health')} className={`px-4 py-2 rounded-lg font-bold text-sm transition-all ${activeSubTab === 'health' ? 'bg-[#081a39] text-[#d4af37] shadow-md' : 'text-slate-600 hover:bg-slate-100'}`}>ตรวจสอบตาราง (Health Check)</button>
        <button onClick={() => setActiveSubTab('missing')} className={`px-4 py-2 rounded-lg font-bold text-sm transition-all ${activeSubTab === 'missing' ? 'bg-[#081a39] text-[#d4af37] shadow-md' : 'text-slate-600 hover:bg-slate-100'}`}>คาบที่จัดไม่ครบ</button>
        <button onClick={() => setActiveSubTab('workload')} className={`px-4 py-2 rounded-lg font-bold text-sm transition-all ${activeSubTab === 'workload' ? 'bg-[#081a39] text-[#d4af37] shadow-md' : 'text-slate-600 hover:bg-slate-100'}`}>ภาระงานครู (Workload)</button>
        <button onClick={() => setActiveSubTab('grid')} className={`px-4 py-2 rounded-lg font-bold text-sm transition-all ${activeSubTab === 'grid' ? 'bg-[#081a39] text-[#d4af37] shadow-md' : 'text-slate-600 hover:bg-slate-100'}`}>ตารางรวมโรงเรียน</button>
      </div>

      {activeSubTab === 'health' && (
        <div className="grid grid-cols-2 gap-6">
          <div className="bg-white rounded-2xl shadow-sm border border-rose-100 p-6">
            <h3 className="font-bold text-rose-700 mb-4 flex items-center gap-2"><AlertTriangle size={18}/> ข้อผิดพลาดร้ายแรง (Hard Conflicts)</h3>
            <ul className="space-y-2 text-sm text-slate-700">
              {health.teacherConflicts.map((c,i) => <li key={`tc-${i}`} className="p-2 bg-rose-50 rounded border border-rose-100">{c}</li>)}
              {health.roomConflicts.map((c,i) => <li key={`rc-${i}`} className="p-2 bg-rose-50 rounded border border-rose-100">{c}</li>)}
              {health.unavailConflicts.map((c,i) => <li key={`uc-${i}`} className="p-2 bg-rose-50 rounded border border-rose-100">{c}</li>)}
              {health.cUnavailConflicts.map((c,i) => <li key={`cuc-${i}`} className="p-2 bg-rose-50 rounded border border-rose-100">{c}</li>)}
              {health.breakConflicts.map((c,i) => <li key={`bc-${i}`} className="p-2 bg-rose-50 rounded border border-rose-100">{c}</li>)}
              {health.blockConflicts.map((c,i) => <li key={`block-${i}`} className="p-2 bg-rose-50 rounded border border-rose-100">{c}</li>)}
              {health.teacherLimitConflicts.map((c,i) => <li key={`limit-${i}`} className="p-2 bg-rose-50 rounded border border-rose-100">{c}</li>)}
              {health.orphans.map((c,i) => <li key={`oc-${i}`} className="p-2 bg-slate-100 rounded border border-slate-300">ข้อมูลสูญหาย: {c}</li>)}
              {health.status === 'empty' && <li className="text-slate-600 p-2 bg-slate-50 rounded border border-slate-200">ยังไม่มีภาระสอน กรุณานำเข้าหรือเพิ่มข้อมูลก่อนจัดตาราง</li>}
              {health.status === 'ready' && <li className="text-emerald-600 font-bold p-2 bg-emerald-50 rounded border border-emerald-100">ตรวจสอบผ่าน: ไม่พบข้อผิดพลาดร้ายแรง</li>}
            </ul>
          </div>
          <div className="bg-white rounded-2xl shadow-sm border border-amber-100 p-6">
            <h3 className="font-bold text-amber-700 mb-4 flex items-center gap-2"><AlertCircle size={18}/> ข้อควรระวัง (Warnings & Missing)</h3>
            <ul className="space-y-2 text-sm text-slate-700">
              {health.overloads.map((c,i) => <li key={`ov-${i}`} className="p-2 bg-amber-50 rounded border border-amber-100">{c}</li>)}
              {health.missing > 0 && <li className="p-2 bg-amber-50 rounded border border-amber-100">มีภาระสอนที่ยังจัดตารางไม่ครบรวม {health.missing} คาบ</li>}
              {(!health.overloads.length && health.missing === 0) && <li className="text-emerald-600 font-bold p-2 bg-emerald-50 rounded border border-emerald-100">ตรวจสอบผ่าน: จัดคาบครบและไม่มีสอนเกิน</li>}
            </ul>
          </div>
        </div>
      )}

      {activeSubTab === 'missing' && (
        <div className="bg-white rounded-2xl shadow-sm border border-slate-100 overflow-hidden">
          <div className="p-4 bg-slate-50 border-b flex gap-3">
             <select value={filterLvl} onChange={e=>setFilterLvl(e.target.value)} className="p-2 border rounded-lg text-sm font-bold"><option value="all">ทุกระดับชั้น</option><option value="ม.1">ม.1</option><option value="ม.2">ม.2</option><option value="ม.3">ม.3</option><option value="ม.4">ม.4</option><option value="ม.5">ม.5</option><option value="ม.6">ม.6</option></select>
          </div>
          <table className="w-full text-left">
            <thead className="bg-[#081a39] text-white"><tr><th className="p-3">วิชา</th><th className="p-3 text-center">ห้อง</th><th className="p-3">ครูผู้สอน</th><th className="p-3 text-center">ต้องการ</th><th className="p-3 text-center">จัดแล้ว</th><th className="p-3 text-center text-rose-300">ขาด</th></tr></thead>
            <tbody>
              {filteredMissing.map((m, i) => (
                <tr key={i} className="border-b hover:bg-slate-50">
                  <td className="p-3 font-bold">{getSubject(m.subjectId).name}</td><td className="p-3 text-center">{getClassroom(m.classroomId).name}</td><td className="p-3">{getTeacher(m.teacherId).name}</td><td className="p-3 text-center">{m.required}</td><td className="p-3 text-center text-emerald-600 font-bold">{m.scheduled}</td><td className="p-3 text-center text-rose-600 font-bold bg-rose-50/50">{m.missing}</td>
                </tr>
              ))}
              {!filteredMissing.length && <tr><td colSpan="6" className="p-10 text-center text-emerald-600 font-bold">จัดคาบสอนครบถ้วนแล้ว!</td></tr>}
            </tbody>
          </table>
        </div>
      )}

      {activeSubTab === 'workload' && (
        <div className="bg-white rounded-2xl shadow-sm border border-slate-100 overflow-hidden">
          <table className="w-full text-left text-sm">
            <thead className="bg-[#081a39] text-white"><tr><th className="p-3">ชื่อครู</th><th className="p-3 text-center">ภาระสอน</th><th className="p-3 text-center">จัดแล้ว</th><th className="p-3 text-center">สถานะ</th><th className="p-3 text-center">Gap</th><th className="p-3 text-center">ติดกันสูงสุด</th></tr></thead>
            <tbody>
              {teachers.map((t, i) => {
                const st = health.teacherStats[t.id] || { req: 0, sch: 0, gaps: 0, maxConsec: 0 };
                return (
                  <tr key={i} className="border-b hover:bg-slate-50">
                    <td className="p-3 font-bold">{t.name}</td><td className="p-3 text-center">{st.req}</td><td className="p-3 text-center font-bold">{st.sch}</td>
                    <td className="p-3 text-center">{st.sch < st.req ? <span className="text-amber-600 bg-amber-50 px-2 py-1 rounded text-xs">ขาด {st.req - st.sch}</span> : st.sch > st.req ? <span className="text-rose-600 bg-rose-50 px-2 py-1 rounded text-xs">เกิน {st.sch - st.req}</span> : <span className="text-emerald-600 text-xs">ครบ</span>}</td>
                    <td className="p-3 text-center">{st.gaps}</td><td className="p-3 text-center">{st.maxConsec}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}

      {activeSubTab === 'grid' && (
        <div className="bg-white rounded-2xl shadow-sm border border-slate-100 overflow-x-auto p-4">
           <div className="p-2 mb-4 bg-blue-50 text-blue-800 rounded-lg text-sm font-bold flex justify-between items-center">
             ตารางรวมโรงเรียน (โหมดห้องเรียน) <button onClick={() => handleExportCSV('all')} className="bg-emerald-600 text-white px-3 py-1 rounded">Export CSV</button>
           </div>
           {DAYS.map(day => (
             <div key={day} className="mb-8">
               <h4 className="font-bold text-lg mb-2 text-[#081a39]">{day}</h4>
               <table className="w-full text-left border-collapse table-fixed text-[11px]">
                 <thead className="bg-slate-100 text-slate-700"><tr><th className="p-2 border w-20 bg-slate-200 text-center">ห้อง</th>{periods.filter(p=>!p.isBreak).map(p => <th key={p.id} className="p-2 border text-center">{p.name}</th>)}</tr></thead>
                 <tbody>
                   {classrooms.map(c => (
                     <tr key={`${day}-${c.id}`} className="hover:bg-amber-50/30">
                       <td className="p-2 border font-bold text-center bg-slate-50">{c.name}</td>
                       {periods.filter(p=>!p.isBreak).map(p => {
                         const sch = schedules.filter(s => s.classroomId === c.id && s.day === day && s.periodId === p.id);
                         const fixed = fixedSchedules.filter(f => fixedAppliesToClassroom(f, c.id) && f.day === day && f.periodId === p.id);
                         const combined = [...fixed, ...sch];
                         return <td key={p.id} className="p-1 border text-center h-12">{combined.map(s => <div key={s.id} className="font-bold truncate">{getSubject(s.subjectId).name}</div>)}</td>
                       })}
                     </tr>
                   ))}
                 </tbody>
               </table>
             </div>
           ))}
        </div>
      )}
    </div>
  );
};

const PrintView = ({ schedules, classrooms, teachers, periods, schoolSettings, getSubject, getShortTeacherName, getClassroom, getTeacher, fixedSchedules }) => {
  const [viewMode, setViewMode] = useState('room');
  const [selectedId, setSelectedId] = useState('');

  const getCellContent = (targetId, day, periodId, vMode) => {
    const cellSchedules = schedules.filter(s => vMode === 'room' ? s.classroomId === targetId && s.day === day && s.periodId === periodId : s.teacherId === targetId && s.day === day && s.periodId === periodId);
    const cellFixed = fixedSchedules.filter(f => vMode === 'room' ? fixedAppliesToClassroom(f, targetId) && f.day === day && f.periodId === periodId : f.teacherId === targetId && f.day === day && f.periodId === periodId);
    const combined = [...cellFixed, ...cellSchedules];
    if (combined.length === 0) return null;
    return (
      <div className="flex flex-col items-center justify-center text-[11px] text-center w-full h-full">
        {combined.map((sch, idx) => (
          <div key={sch.id} className={`w-full ${idx > 0 ? 'mt-1 pt-1 border-t border-black/20' : ''}`}>
            <div className={`font-bold truncate text-[11px] ${cellFixed.includes(sch) ? 'text-amber-800' : ''}`}>{sch.subjectId}</div>
            {vMode === 'room' ? (
               <><div className="truncate text-[11px]">{sch.teacherId ? getShortTeacherName(getTeacher(sch.teacherId).name) : '-'}</div><div className="text-[10px] truncate">{getSubject(sch.subjectId).name}</div></>
             ) : (
               <><div className="truncate text-[11px]">{getClassroom(sch.classroomId).name}</div><div className="text-[10px] truncate">{getSubject(sch.subjectId).name}</div></>
             )}
          </div>
        ))}
      </div>
    );
  };

  const PrintPage = ({ id, mode }) => (
    <div className="bg-white p-10 rounded-xl border border-slate-200 text-black shadow-sm print-page mb-6">
      <div className="text-center mb-8">
        <h1 className="text-3xl font-bold mb-2">{String(schoolSettings?.schoolName || 'โรงเรียน')}</h1>
        <h2 className="text-xl mb-1">ตาราง{mode === 'room' ? 'เรียน' : 'สอน'} ภาคเรียนที่ {String(schoolSettings?.semester || '1')} ปีการศึกษา {String(schoolSettings?.academicYear || '2569')}</h2>
        <h3 className="text-lg font-bold">{mode === 'room' ? `ชั้น ${getClassroom(id).name}` : `ชื่อครูผู้สอน: ${getTeacher(id).name}`}</h3>
      </div>
      <table className="w-full text-left border-collapse border-2 border-black mb-12 print-table table-fixed">
        <thead><tr><th className="border-2 border-black p-3 text-center text-sm w-24 font-bold">วัน/เวลา</th>{periods.map(p => <th key={p.id} className="border-2 border-black p-2 text-center text-xs"><div className="font-bold">{p.name}</div><div className="font-normal text-[11px]">{p.start}-{p.end}</div></th>)}</tr></thead>
        <tbody>
          {DAYS.map(day => (
            <tr key={day}><td className="border-2 border-black p-3 font-bold text-center text-sm">{day}</td>{periods.map(p => <td key={`${day}-${p.id}`} className={`border-2 border-black p-1 h-20 align-middle ${p.isBreak ? 'bg-gray-200/50' : ''}`}>{p.isBreak ? <div className="text-center text-sm text-gray-500 font-bold">พักเที่ยง</div> : getCellContent(id, day, p.id, mode)}</td>)}</tr>
          ))}
        </tbody>
      </table>
      <div className="flex justify-between px-28 mt-24 text-sm font-medium">
        <div className="text-center"><p>ลงชื่อ........................................................</p><p className="mt-3">({String(schoolSettings?.deputyName || '')})</p><p className="mt-1">รองผู้อำนวยการ</p></div>
        <div className="text-center"><p>ลงชื่อ........................................................</p><p className="mt-3">({String(schoolSettings?.directorName || '')})</p><p className="mt-1">ผู้อำนวยการโรงเรียน</p></div>
      </div>
    </div>
  );

  return (
    <div className="animate-in fade-in">
      <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-100 mb-6 flex flex-wrap items-center gap-4 no-print">
        <div className="p-3 bg-[#081a39] rounded-xl shadow-md"><Printer className="text-[#d4af37]" size={28} /></div>
        <select value={viewMode} onChange={e => {setViewMode(e.target.value); setSelectedId('');}} className="p-3 border border-slate-200 rounded-xl bg-slate-50 font-bold text-[#081a39]"><option value="room">พิมพ์ตารางเรียน (รายห้อง)</option><option value="teacher">พิมพ์ตารางสอน (รายครู)</option></select>
        <select value={selectedId} onChange={e => setSelectedId(e.target.value)} className="p-3 border border-slate-200 rounded-xl min-w-[250px] bg-slate-50 font-medium">
          <option value="">-- เลือกเพื่อดูตัวอย่างพิมพ์ --</option><option value="ALL_PRINT" className="font-bold text-[#081a39]">-- พิมพ์ทั้งหมด --</option>
          {viewMode === 'room' ? classrooms.map(c => <option key={c.id} value={c.id}>{c.name}</option>) : teachers.map(t => <option key={t.id} value={t.id}>{t.name}</option>)}
        </select>
        <button onClick={() => window.print()} disabled={!selectedId} className="px-8 py-3 bg-[#081a39] text-white rounded-xl shadow-md font-bold disabled:opacity-50 ml-auto hover:bg-[#122b5e]">พิมพ์เอกสาร</button>
      </div>
      <div id="print-area">{selectedId === 'ALL_PRINT' ? (viewMode === 'room' ? classrooms : teachers).map(item => <PrintPage key={item.id} id={item.id} mode={viewMode} />) : (selectedId && <PrintPage id={selectedId} mode={viewMode} />)}</div>
    </div>
  );
};

export default function App() {
  const [user, setUser] = useState(null);
  const [activeTab, setActiveTab] = useState('dashboard');
  const [isSidebarOpen, setIsSidebarOpen] = useState(false);
  const [toast, setToast] = useState(null);
  const [confirmData, setConfirmData] = useState(null);
  const [isInitialLoading, setIsInitialLoading] = useState(true);

  const [teachers, setTeachers] = useState([]);
  const [subjects, setSubjects] = useState([]);
  const [classrooms, setClassrooms] = useState([]);
  const [periods, setPeriods] = useState(DEFAULT_PERIODS);
  const [schoolSettings, setSchoolSettings] = useState({ schoolName: 'โรงเรียนบ้านแฮดศึกษา', semester: '1', academicYear: '2569', directorName: 'นายสัญญา ประชากูล', deputyName: 'นายชยางกูร ข้อยุ่น' });

  const [teachingLoads, setTeachingLoads] = useState([]);
  const [unavailabilities, setUnavailabilities] = useState([]);
  const [classroomUnavailabilities, setClassroomUnavailabilities] = useState([]);
  const [fixedSchedules, setFixedSchedules] = useState([]);
  const [schedules, setSchedules] = useState([]);
  
  const [isScheduling, setIsScheduling] = useState(false);
  const [scheduleResult, setScheduleResult] = useState(null);
  const [autoScheduleIssues, setAutoScheduleIssues] = useState([]);

  const getTermKey = (settings) => `${settings?.academicYear || '2569'}-${settings?.semester || '1'}`;

  const health = useMemo(() => {
    if (!periods || periods.length === 0) return { status: teachingLoads.length === 0 ? 'empty' : 'warning', missing: 0, scheduledForLoads: 0, totalRequired: 0, completionPercent: 0, teacherConflicts: [], roomConflicts: [], unavailConflicts: [], cUnavailConflicts: [], breakConflicts: [], orphans: [], missingLoads: [], overloads: [], blockConflicts: [], teacherLimitConflicts: [], teacherStats: {} };
    return calculateScheduleHealth(schedules, teachingLoads, unavailabilities, classroomUnavailabilities, fixedSchedules, teachers, subjects, classrooms, periods);
  }, [schedules, teachingLoads, unavailabilities, classroomUnavailabilities, fixedSchedules, teachers, subjects, classrooms, periods]);

  useEffect(() => {
    signInAnonymously(auth).catch(err => console.error(err));
    const unsubscribe = onAuthStateChanged(auth, u => setUser(u || null));
    return () => unsubscribe();
  }, []);

  useEffect(() => {
    if (!user) return;
    const basePath = `school_data/${appId}`; 
    let unsubscribers = [];
    const createListener = (col, setter) => onSnapshot(collection(db, `${basePath}/${col}`), snap => setter(snap.docs.map(doc => ({ ...doc.data(), id: doc.id }))), err => console.error(err));

    unsubscribers.push(createListener('teachers', setTeachers));
    unsubscribers.push(createListener('subjects', setSubjects));
    unsubscribers.push(createListener('classrooms', setClassrooms));
    
    unsubscribers.push(onSnapshot(collection(db, `${basePath}/periods`), snap => {
       const data = snap.docs.map(doc => ({ ...doc.data(), id: doc.id })).sort((a, b) => Number(a.id) - Number(b.id));
       if(data.length > 0) setPeriods(data); else setPeriods(DEFAULT_PERIODS); // ไม่เขียนข้อมูลอัตโนมัติจากทุก client
    }));
    unsubscribers.push(onSnapshot(doc(db, `${basePath}/config/settings`), snap => { if (snap.exists()) setSchoolSettings(snap.data()); setIsInitialLoading(false); }));

    return () => unsubscribers.forEach(unsub => unsub());
  }, [user]);

  useEffect(() => {
    if (!user || !schoolSettings.academicYear || !schoolSettings.semester) return;
    const termKey = getTermKey(schoolSettings);
    const termPath = `school_data/${appId}/terms/${termKey}`;
    let unsubscribers = [];
    const createTermListener = (col, setter) => onSnapshot(collection(db, `${termPath}/${col}`), snap => setter(snap.docs.map(doc => ({ ...doc.data(), id: doc.id }))));

    unsubscribers.push(createTermListener('teachingLoads', setTeachingLoads));
    unsubscribers.push(createTermListener('unavailabilities', setUnavailabilities));
    unsubscribers.push(createTermListener('classroomUnavailabilities', setClassroomUnavailabilities));
    unsubscribers.push(createTermListener('fixedSchedules', setFixedSchedules));
    unsubscribers.push(createTermListener('schedules', setSchedules));

    return () => unsubscribers.forEach(unsub => unsub && unsub());
  }, [user, schoolSettings.academicYear, schoolSettings.semester]);

  const showToast = (msg, type = 'success') => { setToast({ msg: String(msg), type }); setTimeout(() => setToast(null), 3000); };

  const dbAction = async (collectionName, id, data, action = 'set') => {
    if (!user) return false;
    const isTermData = ['teachingLoads', 'schedules', 'unavailabilities', 'classroomUnavailabilities', 'fixedSchedules'].includes(collectionName);
    const basePath = isTermData ? `school_data/${appId}/terms/${getTermKey(schoolSettings)}` : `school_data/${appId}`;
    try {
      if (action === 'set') await setDoc(doc(db, `${basePath}/${collectionName}`, id), { ...(data || {}), id });
      else if (action === 'delete') await deleteDoc(doc(db, `${basePath}/${collectionName}`, id));
      return true;
    } catch (e) { showToast("บันทึกไม่สำเร็จ", 'error'); return false; }
  };

  // STEP 09: Safely migrate a subject document ID and every term's live references.
  // Fail CLOSED if cross-term collection-group queries or permissions are missing.
  // Historical scheduleVersions / downloaded JSON backups are intentionally immutable.
  const renameSubjectId = async (oldId, nextSubject, onSuccess) => {
    const nextId = String(nextSubject?.id || '').trim();
    if (!user) return showToast('กรุณาเข้าสู่ระบบก่อน', 'error');
    if (!oldId || !nextId || oldId === nextId || nextId.includes('/') || nextId === '.' || nextId === '..') {
      return showToast('รหัสวิชาใหม่ไม่ถูกต้อง', 'error');
    }
    const basePath = `school_data/${appId}`;
    const oldRef = doc(db, `${basePath}/subjects`, oldId);
    const nextRef = doc(db, `${basePath}/subjects`, nextId);
    const referencedCollections = ['teachingLoads', 'schedules', 'fixedSchedules'];
    const matchingTermPath = ref => {
      const path = String(ref?.path || '');
      const parts = path.split('/');
      return parts.length === 6 && parts[0] === 'school_data' && parts[1] === appId &&
        parts[2] === 'terms' && /^\d{4}-[12]$/.test(parts[3]) && referencedCollections.includes(parts[4]);
    };
    const readCrossTermReferences = async () => {
      const snapshots = await Promise.all(referencedCollections.map(name =>
        getDocsFromServer(query(collectionGroup(db, name), where('subjectId', '==', oldId)))
      ));
      const result = snapshots.flatMap(snap => snap.docs.filter(d => matchingTermPath(d.ref))
        .map(d => ({ ref: d.ref, data: d.data() })));
      return result.sort((a, b) => a.ref.path.localeCompare(b.ref.path));
    };
    const sameReferences = (a, b) => a.length === b.length && a.every((item, i) =>
      item.ref.path === b[i].ref.path && areFirestoreRecordsEqual(item.data, b[i].data)
    );
    const explainError = error => {
      if (error?.code === 'permission-denied' || error?.code === 'failed-precondition') {
        return 'ยังตรวจภาระสอนทุกภาคเรียนไม่ได้ (สิทธิ์หรือดัชนี Collection Group) จึงไม่เปลี่ยนรหัสวิชา';
      }
      if (error?.message === 'SUBJECT_ID_TAKEN') return 'รหัสวิชาใหม่นี้มีอยู่แล้ว ไม่สามารถเปลี่ยนได้';
      if (error?.message === 'SUBJECT_DATA_CHANGED' || error?.message === 'SUBJECT_REFS_CHANGED') return 'มีผู้แก้ข้อมูลรายวิชาหรือภาระสอนระหว่างดำเนินการ กรุณาตรวจสอบแล้วลองใหม่';
      if (error?.code === 'ATOMIC_LIMIT' || error?.code === 'ATOMIC_SIZE') return error.message;
      return 'เปลี่ยนรหัสวิชาไม่สำเร็จ ข้อมูลเดิมยังไม่ได้ถูกเปลี่ยน';
    };
    try {
      const [oldSnap, nextSnap, initialRefs] = await Promise.all([
        getDocFromServer(oldRef), getDocFromServer(nextRef), readCrossTermReferences(),
      ]);
      if (!oldSnap.exists()) throw new Error('SUBJECT_DATA_CHANGED');
      if (nextSnap.exists()) throw new Error('SUBJECT_ID_TAKEN');
      const originalData = oldSnap.data();
      const countByCollection = name => initialRefs.filter(x => x.ref.path.split('/')[4] === name).length;
      setConfirmData({
        type: 'restore', title: 'ยืนยันเปลี่ยนรหัสวิชา', confirmText: 'ยืนยันเปลี่ยนรหัส',
        message: `เปลี่ยนรหัส ${oldId} เป็น ${nextId} ใช่หรือไม่?\nภาระสอน ${countByCollection('teachingLoads')} รายการ / ตารางสอน ${countByCollection('schedules')} รายการ / กิจกรรมล็อก ${countByCollection('fixedSchedules')} รายการ (รวมทุกภาคเรียน)\nระบบจะเปลี่ยนข้อมูลเหล่านี้พร้อมกัน หากล้มเหลวจะไม่เปลี่ยนบางส่วน`,
        action: async () => {
          setConfirmData(prev => ({ ...prev, isLoading: true }));
          try {
            const liveRefs = await readCrossTermReferences();
            if (!sameReferences(initialRefs, liveRefs)) throw new Error('SUBJECT_REFS_CHANGED');
            const planned = [
              { type: 'set', ref: nextRef, data: { ...nextSubject, id: nextId } },
              { type: 'delete', ref: oldRef },
              ...liveRefs.map(item => ({ type: 'set', ref: item.ref, data: { ...item.data, subjectId: nextId } })),
            ];
            assertAtomicWriteCapacity(planned);
            await runTransaction(db, async transaction => {
              // Read all target documents before starting any writes.
              const [liveOld, liveNew, ...currentRefs] = await Promise.all([
                transaction.get(oldRef), transaction.get(nextRef),
                ...liveRefs.map(item => transaction.get(item.ref)),
              ]);
              if (!liveOld.exists() || !areFirestoreRecordsEqual(originalData, liveOld.data())) throw new Error('SUBJECT_DATA_CHANGED');
              if (liveNew.exists()) throw new Error('SUBJECT_ID_TAKEN');
              if (currentRefs.some((snap, i) => !snap.exists() || !areFirestoreRecordsEqual(snap.data(), liveRefs[i].data))) throw new Error('SUBJECT_REFS_CHANGED');
              transaction.set(nextRef, { ...nextSubject, id: nextId });
              liveRefs.forEach(item => transaction.update(item.ref, { subjectId: nextId }));
              transaction.delete(oldRef);
            });
            setConfirmData(null);
            onSuccess?.();
            showToast(`เปลี่ยนรหัสวิชา ${oldId} → ${nextId} สำเร็จ`);
          } catch (error) {
            console.error('Rename subject ID failed', error);
            showToast(explainError(error), 'error');
            setConfirmData(prev => ({ ...prev, isLoading: false }));
          }
        }
      });
    } catch (error) {
      console.error('Prepare subject ID rename failed', error);
      showToast(explainError(error), 'error');
    }
  };

  // STEP 08: Bulk subject delete uses the same reference guard as single delete,
  // re-checks current-term Firestore data from SERVER, then deletes in one atomic batch.
  const bulkDeleteSubjects = (ids) => {
    const chosenIds = [...new Set((Array.isArray(ids) ? ids : []).map(String))];
    if (!chosenIds.length) return showToast('กรุณาเลือกรายวิชาที่ต้องการลบ', 'error');
    const byId = new Map(subjects.map(item => [item.id, item]));
    const deletable = new Set(getBulkDeletableSubjectIds(subjects, teachingLoads, schedules, fixedSchedules));
    if (chosenIds.some(id => !byId.has(id) || !deletable.has(id))) {
      return showToast('มีวิชาที่ถูกใช้งานหรือไม่มีในระบบ กรุณาเลือกใหม่', 'error');
    }
    const selectedSnapshot = chosenIds.map(id => byId.get(id));
    const termKey = getTermKey(schoolSettings);
    const names = selectedSnapshot.slice(0, 5).map(item => `• ${item.id} ${item.name}`).join('\n');
    setConfirmData({
      type: 'delete', title: `ยืนยันลบ ${chosenIds.length} รายวิชา`, confirmText: `ลบ ${chosenIds.length} รายการ`,
      message: `คุณต้องการลบรายวิชาต่อไปนี้ถาวรใช่หรือไม่?\n${names}${chosenIds.length > 5 ? `\n...และอีก ${chosenIds.length - 5} รายการ` : ''}\n\nรายวิชาเป็นข้อมูลส่วนกลาง อาจถูกใช้งานในภาคเรียนอื่น กรุณาตรวจสอบก่อนลบ`,
      action: async () => {
        setConfirmData(prev => ({ ...prev, isLoading: true }));
        try {
          if (!user) throw new Error('NOT_SIGNED_IN');
          const termPath = `school_data/${appId}/terms/${termKey}`;
          const [masterSnapshot, loadSnapshot, scheduleSnapshot, fixedSnapshot] = await Promise.all([
            getDocsFromServer(collection(db, `school_data/${appId}/subjects`)),
            getDocsFromServer(collection(db, `${termPath}/teachingLoads`)),
            getDocsFromServer(collection(db, `${termPath}/schedules`)),
            getDocsFromServer(collection(db, `${termPath}/fixedSchedules`)),
          ]);
          const liveSubjects = new Map(masterSnapshot.docs.map(d => [d.id, { ...d.data(), id: d.id }]));
          // Protect against changes made by another user since the confirmation opened.
          if (selectedSnapshot.some(item => !liveSubjects.has(item.id) || !areFirestoreRecordsEqual(item, liveSubjects.get(item.id)))) {
            throw new Error('SUBJECT_DATA_CHANGED');
          }
          const liveDeletable = new Set(getBulkDeletableSubjectIds(
            selectedSnapshot,
            loadSnapshot.docs.map(d => d.data()),
            scheduleSnapshot.docs.map(d => d.data()),
            fixedSnapshot.docs.map(d => d.data()),
          ));
          if (chosenIds.some(id => !liveDeletable.has(id))) throw new Error('SUBJECT_IN_USE');
          const operations = chosenIds.map(id => ({ type: 'delete', ref: doc(db, `school_data/${appId}/subjects`, id) }));
          await commitAtomicOperations(operations);
          showToast(`ลบรายวิชาสำเร็จ ${chosenIds.length} รายการ`);
          setConfirmData(null);
        } catch (error) {
          console.error('Bulk delete subjects failed', error);
          if (error.message === 'SUBJECT_IN_USE') showToast('พบรายวิชาที่มีภาระสอน/ตารางสอน/กิจกรรมล็อกเพิ่มขึ้น กรุณาตรวจสอบใหม่ ยังไม่ได้ลบข้อมูล', 'error');
          else if (error.message === 'SUBJECT_DATA_CHANGED') showToast('มีผู้แก้ไขข้อมูลรายวิชาแล้ว กรุณาตรวจสอบรายการใหม่ ยังไม่ได้ลบข้อมูล', 'error');
          else showToast(['ATOMIC_LIMIT', 'ATOMIC_SIZE'].includes(error.code) ? error.message : 'ลบหลายรายวิชาไม่สำเร็จ ยังไม่ได้ลบเพียงบางส่วน', 'error');
          setConfirmData(prev => ({ ...prev, isLoading: false }));
        }
      }
    });
  };

  const handleRequestDelete = (type, item) => {
    let errors = [];
    let scheduleDeleteItems = [item];
    let deleteMessage = 'ข้อมูลนี้จะถูกลบถาวร ดำเนินการต่อหรือไม่?';

    if (type === 'teacher') {
      if (teachingLoads.some(l => l.teacherId === item.id)) errors.push('มีภาระสอน');
      if (schedules.some(s => s.teacherId === item.id)) errors.push('มีตารางสอน');
      if (fixedSchedules.some(f => f.teacherId === item.id)) errors.push('มีกิจกรรมล็อก');
      if (unavailabilities.some(u => u.teacherId === item.id)) errors.push('มีเวลาที่ไม่สะดวก');
    } else if (type === 'subject') {
      if (teachingLoads.some(l => l.subjectId === item.id) || schedules.some(s => s.subjectId === item.id) || fixedSchedules.some(f => f.subjectId === item.id)) errors.push('ถูกใช้งานอยู่');
    } else if (type === 'classroom') {
      if (teachingLoads.some(l => l.classroomId === item.id) || schedules.some(s => s.classroomId === item.id) || fixedSchedules.some(f => f.classroomId === item.id) || classroomUnavailabilities.some(u => u.classroomId === item.id)) errors.push('ถูกใช้งานอยู่');
    } else if (type === 'load') {
      const hasSchedule = schedules.some(s => s.teacherId === item.teacherId && s.subjectId === item.subjectId && s.classroomId === item.classroomId);
      const hasFixed = fixedSchedules.some(f => f.teacherId === item.teacherId && f.subjectId === item.subjectId && f.classroomId === item.classroomId);
      if (hasSchedule || hasFixed) errors.push('มีตารางสอนที่อ้างอิงภาระสอนนี้');
    } else if (type === 'schedule' && Number(getSubject(item.subjectId)?.periodBlock) === 2) {
      const relatedSchedules = schedules.filter(s => s.teacherId === item.teacherId && s.subjectId === item.subjectId && s.classroomId === item.classroomId);
      const relatedFixed = fixedSchedules.filter(f => f.teacherId === item.teacherId && f.subjectId === item.subjectId && f.classroomId === item.classroomId);
      const taggedItems = [
        ...relatedFixed.map(f => ({ ...f, __source: 'fixed' })),
        ...relatedSchedules.map(s => ({ ...s, __source: 'schedule' })),
      ];
      const target = taggedItems.find(x => x.__source === 'schedule' && x.id === item.id);
      const partner = findBlockPartner(target, taggedItems, periods);
      if (partner?.__source === 'fixed') {
        return showToast("คาบนี้จับคู่กับกิจกรรมล็อกอยู่ หากต้องการเปลี่ยนให้จัดการกิจกรรมล็อกก่อน", "error");
      }
      if (partner?.__source === 'schedule') {
        const partnerSchedule = relatedSchedules.find(s => s.id === partner.id);
        if (partnerSchedule) {
          scheduleDeleteItems = [item, partnerSchedule];
          deleteMessage = 'วิชานี้เป็น Block 2 ระบบจะลบคาบที่เป็นคู่ติดกันทั้ง 2 คาบ ดำเนินการต่อหรือไม่?';
        }
      }
    }

    if (errors.length > 0) return showToast(`ไม่สามารถลบได้: ${errors.join(', ')}`, "error");

    const cMap = { 'teacher': 'teachers', 'subject': 'subjects', 'classroom': 'classrooms', 'load': 'teachingLoads', 'unavail': 'unavailabilities', 'cUnavail': 'classroomUnavailabilities', 'fixed': 'fixedSchedules', 'schedule': 'schedules' };
    setConfirmData({
      type: 'delete', title: 'ยืนยันการลบ', message: deleteMessage, confirmText: 'ลบข้อมูล',
      action: async () => {
        setConfirmData(prev => ({ ...prev, isLoading: true }));
        try {
          if (type === 'schedule') await saveScheduleVersion('before_manual_delete');
          if (type === 'schedule' && scheduleDeleteItems.length > 1) {
            const termPath = `school_data/${appId}/terms/${getTermKey(schoolSettings)}`;
            const operations = scheduleDeleteItems.map(s => ({ type: 'delete', ref: doc(db, `${termPath}/schedules`, s.id) }));
            await commitOperationsInChunks(operations);
            showToast("ลบคาบต่อเนื่อง 2 คาบสำเร็จ");
            setConfirmData(null);
            return;
          }
          const ok = await dbAction(cMap[type], item.id, null, 'delete');
          if (ok) { showToast("ลบสำเร็จ"); setConfirmData(null); }
          else setConfirmData(prev => ({ ...prev, isLoading: false }));
        } catch (error) {
          console.error(error);
          showToast("ลบข้อมูลไม่สำเร็จ", "error");
          setConfirmData(prev => ({ ...prev, isLoading: false }));
        }
      }
    });
  };

  const commitOperationsInChunks = async (operations, chunkSize = 400) => {
    for (let i = 0; i < operations.length; i += chunkSize) {
      const batch = writeBatch(db);
      operations.slice(i, i + chunkSize).forEach(op => {
        if (op.type === 'delete') batch.delete(op.ref);
        else batch.set(op.ref, op.data);
      });
      await batch.commit();
    }
  };


  const bulkUpsertTeachingLoads = async (items) => {
    if (!user || !Array.isArray(items) || items.length === 0) return false;
    try {
      const termPath = `school_data/${appId}/terms/${getTermKey(schoolSettings)}`;
      const operations = items.map(item => {
        const id = String(item.id || '').trim();
        const periodsValue = Number(item.periods);
        if (!id || !item.teacherId || !item.subjectId || !item.classroomId || !Number.isInteger(periodsValue) || periodsValue < 1) {
          throw new Error('ข้อมูลภาระสอนไม่ถูกต้อง');
        }
        const cleanItem = {
          id,
          teacherId: String(item.teacherId),
          subjectId: String(item.subjectId),
          classroomId: String(item.classroomId),
          periods: periodsValue,
        };
        return { type: 'set', ref: doc(db, `${termPath}/teachingLoads`, id), data: cleanItem };
      });
      // STEP 03: one atomic batch, not several partially committed chunks.
      await commitAtomicOperations(operations);
      return true;
    } catch (error) {
      console.error(error);
      showToast(['ATOMIC_LIMIT', 'ATOMIC_SIZE'].includes(error.code) ? error.message : "นำเข้าภาระสอนไม่สำเร็จ (ไม่มีการนำเข้าเป็นบางส่วน)", "error");
      return false;
    }
  };

  // STEP 06: Keep each critical save atomic, including school-wide schedules.
  // Firestore removed its 500-write Commit limit in 2023; the 10 MiB request limit
  // and index costs still apply. Use a conservative estimated-size safety guard.
  // Do not fall back to sequential batches: that could leave half a timetable saved.
  const ATOMIC_OPERATION_LIMIT = 3000;
  const ATOMIC_ESTIMATED_BYTES_LIMIT = 6 * 1024 * 1024;
  const assertAtomicWriteCapacity = (operations) => {
    if (operations.length > ATOMIC_OPERATION_LIMIT) {
      const error = new Error(`ต้องเปลี่ยน ${operations.length} รายการ เกินเพดาน ${ATOMIC_OPERATION_LIMIT} รายการต่อครั้ง ระบบยังไม่ได้แก้ข้อมูลใด ๆ`);
      error.code = 'ATOMIC_LIMIT';
      throw error;
    }
    // Estimate payload before creating the batch, allowing headroom for protocol
    // and Firestore index overhead. This is NOT a guarantee of server acceptance.
    const encoder = new TextEncoder();
    let estimatedBytes = 0;
    for (const op of operations) {
      estimatedBytes += encoder.encode(JSON.stringify({ type: op.type, path: op.ref.path, data: op.type === 'delete' ? null : op.data })).length + 200;
      if (estimatedBytes > ATOMIC_ESTIMATED_BYTES_LIMIT) {
        const error = new Error(`ข้อมูลที่จะบันทึกมีขนาดใหญ่เกินเกณฑ์ปลอดภัย (ประมาณ ${(estimatedBytes / 1048576).toFixed(1)} MiB) ยังไม่ได้แก้ข้อมูลใด ๆ กรุณาแบ่งข้อมูลนำเข้าหรือติดต่อผู้ดูแลระบบ`);
        error.code = 'ATOMIC_SIZE';
        throw error;
      }
    }
    return estimatedBytes;
  };
  const commitAtomicOperations = async (operations) => {
    assertAtomicWriteCapacity(operations);
    if (operations.length === 0) return;
    const batch = writeBatch(db);
    operations.forEach(op => {
      if (op.type === 'delete') batch.delete(op.ref);
      else batch.set(op.ref, op.data);
    });
    await batch.commit();
  };

  // Read SERVER data and assemble the entire replacement without writing anything.
  const prepareCollectionReplacement = async (collectionPath, items) => {
    if (!Array.isArray(items)) throw new Error('INVALID_REPLACEMENT_ITEMS');
    const currentSnapshot = await getDocsFromServer(collection(db, collectionPath));
    const previousItems = currentSnapshot.docs.map(d => ({ ...d.data(), id: d.id }));
    const currentById = new Map(previousItems.map(d => [d.id, d]));
    const targetIds = new Set();
    const operations = [];

    items.forEach(item => {
      if (!item || typeof item !== 'object' || Array.isArray(item)) throw new Error('INVALID_REPLACEMENT_ITEM');
      const cleanItem = { ...item };
      delete cleanItem._col;
      const itemId = String(cleanItem.id || '').trim();
      if (!itemId) throw new Error(`Missing document id in ${collectionPath}`);
      if (targetIds.has(itemId)) throw new Error(`Duplicate document id ${itemId} in ${collectionPath}`);
      targetIds.add(itemId);
      cleanItem.id = itemId;
      const currentItem = currentById.get(itemId);
      if (!currentItem || !areFirestoreRecordsEqual(currentItem, cleanItem)) {
        operations.push({ type: 'set', ref: doc(db, collectionPath, itemId), data: cleanItem });
      }
    });
    // Deletions are planned before any write, in the SAME atomic batch as sets.
    currentSnapshot.docs.filter(d => !targetIds.has(d.id)).forEach(d => {
      operations.push({ type: 'delete', ref: d.ref });
    });
    return { operations, previousItems };
  };

  const replaceCollectionData = async (collectionPath, items) => {
    const plan = await prepareCollectionReplacement(collectionPath, items);
    await commitAtomicOperations(plan.operations);
  };

  const readCollectionData = async (collectionPath) => {
    const snapshot = await getDocs(collection(db, collectionPath));
    return snapshot.docs.map(d => ({ ...d.data(), id: d.id }));
  };

  const pruneScheduleVersions = async (termPath) => {
    try {
      const versionsSnapshot = await getDocs(collection(db, `${termPath}/scheduleVersions`));
      const versions = versionsSnapshot.docs
        .map(d => ({ ref: d.ref, createdAt: Number(d.data()?.createdAt) || 0 }))
        .sort((a, b) => b.createdAt - a.createdAt);
      const excess = versions.slice(SCHEDULE_VERSION_LIMIT);
      if (excess.length > 0) await commitOperationsInChunks(excess.map(v => ({ type: 'delete', ref: v.ref })));
    } catch (error) {
      console.error('Prune schedule versions failed', error);
    }
  };

  const saveScheduleVersion = async (reason = 'manual_snapshot', scheduleItems = schedules, fixedItems = fixedSchedules, explicitTermKey = null) => {
    if (!user) return null;
    try {
      const termKey = explicitTermKey || getTermKey(schoolSettings);
      if (!termKey) return null;
      const termPath = `school_data/${appId}/terms/${termKey}`;
      const versionRef = doc(collection(db, `${termPath}/scheduleVersions`));
      const safeSchedules = (Array.isArray(scheduleItems) ? scheduleItems : []).map(cleanScheduleSnapshotItem);
      const safeFixedSchedules = (Array.isArray(fixedItems) ? fixedItems : []).map(cleanScheduleSnapshotItem);
      await setDoc(versionRef, {
        id: versionRef.id,
        createdAt: Date.now(),
        reason: String(reason || 'manual_snapshot'),
        scheduleCount: safeSchedules.length,
        fixedScheduleCount: safeFixedSchedules.length,
        schedules: safeSchedules,
        fixedSchedules: safeFixedSchedules,
      });
      await pruneScheduleVersions(termPath);
      return versionRef.id;
    } catch (error) {
      console.error('Save schedule version failed', error);
      return null;
    }
  };

  const handleRestoreBackup = (jsonData) => {
    const backupTermKey = jsonData?.termKey || (jsonData?.schoolSettings ? getTermKey(jsonData.schoolSettings) : '');
    const backupCollections = ['teachers', 'subjects', 'classrooms', 'periods', 'teachingLoads', 'unavailabilities', 'classroomUnavailabilities', 'fixedSchedules', 'schedules'];
    if (!jsonData || jsonData.backupVersion !== 2 || jsonData.appId !== appId || !/^\d{4}-[12]$/.test(backupTermKey) || !backupCollections.every(key => Array.isArray(jsonData[key]))) {
      return showToast("ไฟล์สำรองข้อมูลไม่ครบ หรือไม่ตรงระบบและภาคเรียนที่รองรับ", "error");
    }

    setConfirmData({
      type: 'restore', title: 'นำเข้าข้อมูล', message: `เขียนทับตารางเดิมสำหรับเทอม ${backupTermKey} ใช่หรือไม่?`, confirmText: 'นำเข้า',
      action: async () => {
        setConfirmData(prev => ({ ...prev, isLoading: true }));
        try {
          const basePath = `school_data/${appId}`;
          const targetTermPath = `${basePath}/terms/${backupTermKey}`;
          const masterNames = ['teachers', 'subjects', 'classrooms', 'periods'];
          // STEP 03: Plan ALL nine collections first; one batch commits everything or nothing.
          const plans = {};
          const allOperations = [];
          for (const colName of backupCollections) {
            const collectionPath = masterNames.includes(colName) ? `${basePath}/${colName}` : `${targetTermPath}/${colName}`;
            const plan = await prepareCollectionReplacement(collectionPath, jsonData[colName]);
            plans[colName] = plan;
            allOperations.push(...plan.operations);
          }
          if (jsonData.schoolSettings) {
            allOperations.push({ type: 'set', ref: doc(db, `${basePath}/config/settings`), data: jsonData.schoolSettings });
          }
          // Preflight count AND estimated size before writing any backup/target data.
          assertAtomicWriteCapacity(allOperations);
          // Only the schedule and fixed schedule are stored in legacy version history.
          const restoreVersionId = await saveScheduleVersion('before_backup_restore', plans.schedules.previousItems, plans.fixedSchedules.previousItems, backupTermKey);
          if (!restoreVersionId) throw new Error('RESTORE_SNAPSHOT_FAILED');
          await commitAtomicOperations(allOperations);
          showToast("นำเข้าข้อมูลสำเร็จ!"); setConfirmData(null);
        } catch (error) {
          console.error(error);
          showToast(['ATOMIC_LIMIT', 'ATOMIC_SIZE'].includes(error.code) ? error.message : "นำเข้าข้อมูลไม่สำเร็จ ระบบไม่ได้บันทึกข้อมูลเป็นบางส่วน", "error");
          setConfirmData(null);
        }
      }
    });
  };

  // STEP 02: Snapshot the inputs used to produce the auto-scheduler preview.
  // Keep this local to the preview/commit flow; no changes to regular CRUD or scheduler algorithm.
  const schedulePreviewCollectionNames = [
    'schedules', 'fixedSchedules', 'teachingLoads', 'unavailabilities',
    'classroomUnavailabilities', 'teachers', 'subjects', 'classrooms', 'periods',
  ];
  const captureSchedulePreviewSource = (data, termKey) => {
    const fingerprints = {};
    schedulePreviewCollectionNames.forEach(name => {
      // Ignore Firestore result ordering; compare records deterministically by document ID.
      const records = (Array.isArray(data[name]) ? data[name] : [])
        .map(cleanScheduleSnapshotItem)
        .sort((a, b) => String(a.id || '').localeCompare(String(b.id || '')));
      fingerprints[name] = stableSerializeData(records);
    });
    return { termKey, fingerprints };
  };
  const currentSchedulePreviewSource = () => captureSchedulePreviewSource({
    schedules, fixedSchedules, teachingLoads, unavailabilities,
    classroomUnavailabilities, teachers, subjects, classrooms, periods,
  }, getTermKey(schoolSettings));

  const runSmartSchedule = async () => {
    if (isScheduling) return;
    // STEP 04: แจ้งรายการที่ต้องแก้ก่อนขอล็อกหรือเริ่มสร้างตาราง
    const issues = validateAutoScheduleInputs({
      teachingLoads, teachers, subjects, classrooms, periods, schedules, fixedSchedules,
      unavailabilities, classroomUnavailabilities, semester: schoolSettings.semester,
    });
    if (issues.length) { setAutoScheduleIssues(issues); return; }
    setAutoScheduleIssues([]);
    // A frozen copy of the data as it was when scheduling began.
    const previewSource = currentSchedulePreviewSource();
    const lockCollection = `school_data/${appId}/terms/${getTermKey(schoolSettings)}/config`;
    const lockRef = doc(db, `${lockCollection}/schedulerLock`);
    // A randomly generated document ID acts as a unique owner token; no extra document is written.
    const lockOwnerId = doc(collection(db, lockCollection)).id;
    try {
      const acquired = await acquireSchedulerLock(lockRef, lockOwnerId);
      if (!acquired) return showToast("มีผู้อื่นจัดตารางอยู่", "error");
    } catch (error) {
      console.error('Acquire scheduler lock failed', error);
      return showToast("ตรวจล็อกการจัดตารางไม่ได้ (อาจออฟไลน์หรือไม่มีสิทธิ์)", "error");
    }
    setIsScheduling(true);

    setTimeout(async () => {
      try {
        const periodContext = buildPeriodContext(periods);
        const activePeriods = periodContext.activePeriods;
        if (activePeriods.length === 0) throw new Error('NO_ACTIVE_PERIODS');

        // Preserve ตารางที่ผู้ใช้จัดเองไว้ทั้งหมด และสร้างใหม่เฉพาะรายการที่ Auto Scheduler เคยสร้าง
        const preservedSchedules = schedules
          .filter(s => !isAutoGeneratedSchedule(s))
          .map(s => ({ ...cleanScheduleSnapshotItem(s), source: s.source || 'manual' }));
        const baseMaps = buildConstraintMaps(preservedSchedules, unavailabilities, classroomUnavailabilities, fixedSchedules, periods);
        const loadKeys = new Set(teachingLoads.map(getTeachingLoadKey));
        const fixedByLoad = new Map();
        const preservedByLoad = new Map();
        const fixedTeachingSchedules = [];

        fixedSchedules.forEach(f => {
          if (!f.teacherId) return;
          const key = getTeachingLoadKey(f);
          if (!loadKeys.has(key)) return;
          if (!fixedByLoad.has(key)) fixedByLoad.set(key, []);
          fixedByLoad.get(key).push(f);
          fixedTeachingSchedules.push(f);
        });

        preservedSchedules.forEach(s => {
          if (!s.teacherId) return;
          const key = getTeachingLoadKey(s);
          if (!loadKeys.has(key)) return;
          if (!preservedByLoad.has(key)) preservedByLoad.set(key, []);
          preservedByLoad.get(key).push(s);
        });

        const teacherTotalLoad = new Map();
        teachingLoads.forEach(load => teacherTotalLoad.set(load.teacherId, (teacherTotalLoad.get(load.teacherId) || 0) + (Number(load.periods) || 0)));

        const enumerateSingleOptions = (load, maps, specificDay = null, specificPeriodIds = null) => {
          const options = [];
          const days = specificDay ? [specificDay] : DAYS;
          const allowedIds = specificPeriodIds ? new Set(specificPeriodIds) : null;
          days.forEach(day => {
            activePeriods.forEach(period => {
              if (allowedIds && !allowedIds.has(period.id)) return;
              const valid = canPlaceSchedule(load.teacherId, load.classroomId, day, period.id, maps, false);
              if (valid.ok) options.push([{ day, periodId: period.id }]);
            });
          });
          return options;
        };

        const enumerateDoubleOptions = (load, maps) => {
          const options = [];
          DAYS.forEach(day => {
            periodContext.doublePairs.forEach(([p1, p2]) => {
              const valid1 = canPlaceSchedule(load.teacherId, load.classroomId, day, p1.id, maps, false);
              const valid2 = canPlaceSchedule(load.teacherId, load.classroomId, day, p2.id, maps, false);
              if (valid1.ok && valid2.ok) {
                const pair = [{ day, periodId: p1.id }, { day, periodId: p2.id }];
                if (canPlaceTeacherEntries(load.teacherId, pair, maps).ok) options.push(pair);
              }
            });
          });
          return options;
        };

        const getCompanionOptions = (load, anchorItem, maps) => {
          const neighborIds = periodContext.neighborIdsByPeriod.get(anchorItem.periodId) || [];
          if (neighborIds.length === 0) return [];
          return enumerateSingleOptions(load, maps, anchorItem.day, neighborIds);
        };

        const loadPlans = teachingLoads.map(load => {
          const key = getTeachingLoadKey(load);
          const fixedItems = fixedByLoad.get(key) || [];
          const preservedItems = preservedByLoad.get(key) || [];
          const anchorItems = [...fixedItems, ...preservedItems];
          const totalRequired = Math.max(0, Number(load.periods) || 0);
          const remaining = Math.max(0, totalRequired - anchorItems.length);
          const block = Number(getSubject(load.subjectId).periodBlock) === 2 ? 2 : 1;
          const staticOptionCount = block === 2 && remaining >= 2 ? enumerateDoubleOptions(load, baseMaps).length : enumerateSingleOptions(load, baseMaps).length;
          return { load, key, fixedItems, preservedItems, anchorItems, totalRequired, remaining, block, staticOptionCount };
        });

        const runSeed = (Date.now() ^ (teachingLoads.length * 2654435761) ^ fixedSchedules.length ^ preservedSchedules.length) >>> 0;
        const candidatePool = [];
        const candidateSignatures = new Set();
        let lastLockRenewal = Date.now();

        for (let c = 0; c < SCHEDULER_CONFIG.candidateCount; c++) {
          const random = createSeededRandom((runSeed + Math.imul(c + 1, 2246822519)) >>> 0);
          const generatedSchedule = [];
          const currentUnassigned = [];
          const currentMaps = cloneConstraintMaps(baseMaps);
          const placementState = buildPlacementState([...fixedTeachingSchedules, ...preservedSchedules]);
          let idSequence = 0;

          const addEntries = (load, entries) => {
            entries.forEach(entry => {
              const id = `sch_auto_${runSeed}_${c}_${idSequence++}`;
              const item = {
                id,
                teacherId: load.teacherId,
                subjectId: load.subjectId,
                classroomId: load.classroomId,
                day: entry.day,
                periodId: entry.periodId,
                source: 'auto',
                generationId: String(runSeed),
              };
              generatedSchedule.push(item);
              if (load.teacherId) currentMaps.teacherPlacements.add(`${load.teacherId}_${entry.day}_${entry.periodId}`);
              registerTeacherPeriod(currentMaps, item);
              currentMaps.roomPlacements.add(`${load.classroomId}_${entry.day}_${entry.periodId}`);
              registerPlacementState(placementState, item);
            });
          };

          const sortedPlans = loadPlans
            .map(plan => ({ ...plan, tieOrder: random() }))
            .sort((a, b) => {
              const aRatio = a.staticOptionCount / Math.max(1, a.remaining);
              const bRatio = b.staticOptionCount / Math.max(1, b.remaining);
              if (a.block !== b.block) return b.block - a.block;
              if (aRatio !== bRatio) return aRatio - bRatio;
              const teacherLoadDiff = (teacherTotalLoad.get(b.load.teacherId) || 0) - (teacherTotalLoad.get(a.load.teacherId) || 0);
              if (teacherLoadDiff !== 0) return teacherLoadDiff;
              if (a.remaining !== b.remaining) return b.remaining - a.remaining;
              return a.tieOrder - b.tieOrder;
            });

          for (const plan of sortedPlans) {
            const { load, block, totalRequired, anchorItems } = plan;
            let remain = plan.remaining;
            let mandatoryBlockMissing = 0;

            if (block === 2 && remain > 0 && anchorItems.length > 0) {
              const unpairedAnchors = analyzeBlockPairing(anchorItems, periodContext).singles;
              const allowedSingles = totalRequired % 2;
              const rankedAnchors = unpairedAnchors
                .map(item => ({ item, optionCount: getCompanionOptions(load, item, currentMaps).length, tie: random() }))
                .sort((a, b) => a.optionCount - b.optionCount || a.tie - b.tie);
              const companionTargets = rankedAnchors.slice(Math.min(allowedSingles, rankedAnchors.length)).map(x => x.item);

              for (const anchorItem of companionTargets) {
                if (remain <= 0) break;
                const options = getCompanionOptions(load, anchorItem, currentMaps);
                const chosen = choosePlacementOption(options, load, placementState, periodContext, getSubject, random);
                if (chosen) addEntries(load, chosen);
                else mandatoryBlockMissing += 1;
                remain -= 1;
              }
            }

            while (remain > 0) {
              let options = [];
              let placementSize = 1;
              if (block === 2 && remain >= 2) {
                options = enumerateDoubleOptions(load, currentMaps);
                placementSize = 2;
              } else {
                options = enumerateSingleOptions(load, currentMaps);
              }
              if (options.length === 0) break;
              const chosen = choosePlacementOption(options, load, placementState, periodContext, getSubject, random);
              if (!chosen) break;
              addEntries(load, chosen);
              remain -= placementSize;
            }

            const missing = mandatoryBlockMissing + remain;
            if (missing > 0) {
              const reason = mandatoryBlockMissing > 0 ? 'BLOCK_COMPANION_NOT_AVAILABLE' : (block === 2 && remain >= 2 ? 'DOUBLE_PERIOD_NOT_AVAILABLE' : 'NO_AVAILABLE_SLOT');
              currentUnassigned.push({ ...load, missing, reason });
            }
          }

          const fullSchedule = [...preservedSchedules, ...generatedSchedule];
          const score = calculateScheduleScore([...fixedTeachingSchedules, ...fullSchedule], teachers, subjects, periodContext, getSubject);
          const missingCount = currentUnassigned.reduce((acc, curr) => acc + (Number(curr.missing) || 0), 0);
          const finalScore = score - (missingCount * 1000);
          const signature = generatedSchedule
            .map(item => `${item.teacherId}|${item.subjectId}|${item.classroomId}|${item.day}|${item.periodId}`)
            .sort()
            .join('~');

          if (!candidateSignatures.has(signature)) {
            candidateSignatures.add(signature);
            candidatePool.push({
              schedule: fullSchedule,
              generatedSchedule,
              preservedCount: preservedSchedules.length,
              unassigned: currentUnassigned,
              score: finalScore,
              missingCount,
              generationId: String(runSeed),
            });
          }

          // คืนเวลาให้ browser เป็นช่วง ๆ เพื่อให้ spinner/การโต้ตอบไม่ค้างระหว่างคำนวณหลาย candidate
          if ((c + 1) % SCHEDULER_CONFIG.yieldEveryCandidates === 0 && c + 1 < SCHEDULER_CONFIG.candidateCount) {
            // Renew only if needed, to avoid adding a network round trip on every candidate.
            if (Date.now() - lastLockRenewal >= SCHEDULER_LOCK_RENEW_MS) {
              if (!(await renewSchedulerLock(lockRef, lockOwnerId))) throw new Error('SCHEDULER_LOCK_LOST');
              lastLockRenewal = Date.now();
            }
            await new Promise(resolve => setTimeout(resolve, 0));
          }
        }

        // Do not publish a result if this run has lost ownership of the lock.
        if (!(await renewSchedulerLock(lockRef, lockOwnerId))) throw new Error('SCHEDULER_LOCK_LOST');
        candidatePool.sort((a, b) => {
          if (a.missingCount !== b.missingCount) return a.missingCount - b.missingCount;
          if (a.score !== b.score) return b.score - a.score;
          return a.generatedSchedule.length - b.generatedSchedule.length;
        });

        const candidates = candidatePool.slice(0, SCHEDULER_CANDIDATE_KEEP);
        const bestCandidate = candidates[0] || { schedule: preservedSchedules, generatedSchedule: [], preservedCount: preservedSchedules.length, unassigned: [], score: -Infinity, missingCount: 0, generationId: String(runSeed) };
        if (findTeacherTeachingLimitConflicts([...fixedSchedules, ...bestCandidate.schedule], periods, teachers).length) {
          throw new Error('TEACHER_HARD_LIMIT_FAILED');
        }
        const totalReq = teachingLoads.reduce((a, curr) => a + (Number(curr.periods) || 0), 0);
        const missing = bestCandidate.unassigned.reduce((sum, item) => sum + (Number(item.missing) || 0), 0);
        setScheduleResult({
          ...bestCandidate,
          candidates,
          previewSource,
          stats: {
            totalReq,
            scheduled: Math.max(0, totalReq - missing),
            missing,
            preserved: preservedSchedules.length,
            generated: bestCandidate.generatedSchedule.length,
          },
        });
      } catch (err) {
        console.error(err);
        showToast(err.message === 'SCHEDULER_LOCK_LOST'
          ? "สิทธิ์ล็อกการจัดตารางสิ้นสุดหรือถูกเปลี่ยน ระบบยกเลิกผลรอบนี้ กรุณาจัดใหม่"
          : "เกิดข้อผิดพลาดในการประมวลผล", "error");
      } finally {
        try { await releaseSchedulerLock(lockRef, lockOwnerId); }
        catch (error) {
          console.error('Release scheduler lock failed', error);
          showToast("ปลดล็อกการจัดตารางไม่สำเร็จ กรุณาตรวจการเชื่อมต่อและลองใหม่ภายหลัง", "error");
        }
        setIsScheduling(false);
      }
    }, 100);
  };

  const commitSchedule = async () => {
    if (!scheduleResult) return;
    setConfirmData({
      type: 'restore', title: 'ยืนยันการใช้ตารางใหม่', message: 'ตารางเดิมในเทอมนี้จะถูกแทนที่ ยืนยันหรือไม่?', confirmText: 'บันทึกตาราง',
      action: async () => {
        setConfirmData(prev => ({...prev, isLoading: true}));
        const termKey = getTermKey(schoolSettings);
        const termPath = `school_data/${appId}/terms/${termKey}`;
        let previousSchedules = [];
        try {
          const expected = scheduleResult.previewSource;
          if (!expected || expected.termKey !== termKey) {
            throw Object.assign(new Error('ภาคเรียนเปลี่ยนไปหลังจัดตาราง กรุณาจัดใหม่'), { code: 'STALE_PREVIEW' });
          }

          // First reject changes already received via live Firestore listeners.
          const localNow = currentSchedulePreviewSource();
          if (schedulePreviewCollectionNames.some(name => localNow.fingerprints[name] !== expected.fingerprints[name])) {
            throw Object.assign(new Error('มีข้อมูลเปลี่ยนไประหว่างเปิดตัวอย่าง กรุณาจัดตารางใหม่'), { code: 'STALE_PREVIEW' });
          }

          // Then verify against the SERVER, not the browser cache. If offline, refuse to overwrite.
          const basePath = `school_data/${appId}`;
          const remoteSettings = await getDocFromServer(doc(db, `${basePath}/config/settings`));
          if (remoteSettings.exists() && getTermKey(remoteSettings.data()) !== termKey) {
            throw Object.assign(new Error('ภาคเรียนในฐานข้อมูลเปลี่ยนไป กรุณาโหลดใหม่'), { code: 'STALE_PREVIEW' });
          }
          const serverData = {};
          await Promise.all(schedulePreviewCollectionNames.map(async name => {
            const isTermData = ['schedules', 'fixedSchedules', 'teachingLoads', 'unavailabilities', 'classroomUnavailabilities'].includes(name);
            const path = isTermData ? `${termPath}/${name}` : `${basePath}/${name}`;
            const snap = await getDocsFromServer(collection(db, path));
            const records = snap.docs.map(d => ({ ...d.data(), id: d.id }));
            // The app displays default periods when there are no period documents yet.
            serverData[name] = name === 'periods' && records.length === 0 ? DEFAULT_PERIODS : records;
          }));
          const serverNow = captureSchedulePreviewSource(serverData, termKey);
          if (schedulePreviewCollectionNames.some(name => serverNow.fingerprints[name] !== expected.fingerprints[name])) {
            throw Object.assign(new Error('ข้อมูลจากอีกเครื่องเปลี่ยนไปหลังสร้างตัวอย่าง กรุณาจัดตารางใหม่'), { code: 'STALE_PREVIEW' });
          }

          // Use a fresh server copy for the recovery snapshot, not potentially stale React state.
          previousSchedules = serverData.schedules.map(cleanScheduleSnapshotItem);
          const versionId = await saveScheduleVersion('before_auto_commit', previousSchedules, serverData.fixedSchedules);
          if (!versionId) throw new Error('บันทึกจุดย้อนกลับไม่สำเร็จ จึงยังไม่เปลี่ยนตาราง');
          await replaceCollectionData(`${termPath}/schedules`, scheduleResult.schedule.map(cleanScheduleSnapshotItem));
          showToast("ใช้งานตารางสอนใหม่สำเร็จ!"); setScheduleResult(null); setConfirmData(null);
        } catch (e) {
          console.error(e);
          // STEP 03: Failed writeBatch is atomic; do not run a second replacement
          // as "rollback" (it could overwrite another user's concurrent changes).
          if (e.code === 'STALE_PREVIEW') setScheduleResult(null);
          const message = ['ATOMIC_LIMIT', 'ATOMIC_SIZE'].includes(e.code) ? e.message
            : e.code === 'STALE_PREVIEW' ? e.message
            : 'บันทึกไม่สำเร็จ ไม่ได้เขียนตารางเป็นบางส่วน: ' + e.message;
          showToast(message, 'error');
          setConfirmData(null);
        }
      }
    });
  };

  // Lookup maps ลดการ .find() ซ้ำจำนวนมากใน Scheduler, Reports และตารางแสดงผล
  const teacherLookup = useMemo(() => new Map(teachers.map(t => [t.id, t])), [teachers]);
  const subjectLookup = useMemo(() => new Map(subjects.map(s => [s.id, s])), [subjects]);
  const classroomLookup = useMemo(() => new Map(classrooms.map(c => [c.id, c])), [classrooms]);
  const getTeacher = (id) => teacherLookup.get(id) || { name: 'ไม่ทราบ' };
  const getSubject = (id) => subjectLookup.get(id) || { name: 'ไม่ทราบ', abbr: '-' };
  const getClassroom = (id) => id === ALL_CLASSROOMS_ID ? { id, name: 'ทุกห้องเรียน' } : (classroomLookup.get(id) || { name: 'ไม่ทราบ' });
  const activeSubjects = useMemo(() => subjects.filter(s => !s.term || s.term === 'all' || s.term === schoolSettings.semester), [subjects, schoolSettings.semester]);
  const autoGenerateClassrooms = async () => {
    const existingIds = new Set(classrooms.map(c => c.id));
    const pendingRooms = [];
    for (let m = 1; m <= 5; m++) {
      for (let r = 1; r <= 3; r++) {
        const id = `M${m}${r}`;
        if (!existingIds.has(id)) pendingRooms.push({ id, name: `ม.${m}/${r}` });
      }
    }
    for (let r = 1; r <= 2; r++) {
      const id = `M6${r}`;
      if (!existingIds.has(id)) pendingRooms.push({ id, name: `ม.6/${r}` });
    }
    if (pendingRooms.length === 0) return showToast("มีห้องเรียนครบแล้ว");
    try {
      const operations = pendingRooms.map(room => ({ type: 'set', ref: doc(db, `school_data/${appId}/classrooms`, room.id), data: room }));
      await commitOperationsInChunks(operations);
      showToast(`สร้างห้องเรียนอัตโนมัติสำเร็จ ${pendingRooms.length} ห้อง`);
    } catch (error) {
      console.error(error);
      showToast("สร้างห้องเรียนอัตโนมัติไม่สำเร็จ", "error");
    }
  };

  const TABS = [
    { id: 'dashboard', name: 'หน้าแรก', icon: LayoutDashboard }, { id: 'teachers', name: 'ข้อมูลครู', icon: Users }, { id: 'subjects', name: 'รายวิชา', icon: BookOpen },
    { id: 'classrooms', name: 'ห้องเรียน', icon: Home }, { id: 'loads', name: 'ภาระสอน', icon: Briefcase }, { id: 'constraints', name: 'ข้อจำกัดเวลา', icon: Clock },
    { id: 'schedule', name: 'จัดตารางสอน', icon: CalendarDays }, { id: 'reports', name: 'รายงานและ Export', icon: Activity }, { id: 'print', name: 'พิมพ์ตาราง', icon: Printer },
    { id: 'settings', name: 'ตั้งค่าระบบ', icon: Settings }
  ];

  const commonProps = { user, appId, db, dbAction, bulkUpsertTeachingLoads, bulkDeleteSubjects, renameSubjectId, showToast, handleRequestDelete, saveScheduleVersion, getTermKey, getTeacher, getSubject, getClassroom, getShortTeacherName, teachers, subjects, activeSubjects, classrooms, periods, teachingLoads, unavailabilities, classroomUnavailabilities, fixedSchedules, schedules, schoolSettings, health, fullDataForBackup: { teachers, subjects, classrooms, periods, teachingLoads, unavailabilities, classroomUnavailabilities, fixedSchedules, schedules } };

  if (isInitialLoading) return <div className="flex h-screen w-full items-center justify-center bg-slate-50"><div className="animate-spin text-[#d4af37]"><RefreshCw size={48}/></div></div>;

  return (
    <div className="flex h-screen bg-[#f8fafc] font-sans text-slate-900 selection:bg-[#d4af37]/30">
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;500;600;700&display=swap');
        .font-sans { font-family: 'Sarabun', sans-serif; }
        .custom-scrollbar::-webkit-scrollbar { width: 6px; height: 6px; }
        .custom-scrollbar::-webkit-scrollbar-track { background: transparent; }
        .custom-scrollbar::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 10px; }
        .custom-scrollbar::-webkit-scrollbar-thumb:hover { background: #081a39; }
        @media print { @page { size: A4 landscape; margin: 10mm; } body * { visibility: hidden; } #print-area, #print-area * { visibility: visible; } #print-area { position: absolute; left: 0; top: 0; width: 100%; } .no-print { display: none !important; } .print-table { -webkit-print-color-adjust: exact; print-color-adjust: exact; } .print-page { break-after: page; page-break-after: always; border: none !important; box-shadow: none !important; margin: 0 !important; padding: 0 !important;} }
      `}</style>
      
      {isSidebarOpen && <div className="fixed inset-0 bg-slate-900/50 z-20 md:hidden backdrop-blur-sm" onClick={() => setIsSidebarOpen(false)} />}
      <aside className={`fixed md:static inset-y-0 left-0 w-64 bg-gradient-to-b from-[#081a39] to-[#153063] border-r border-[#d4af37]/20 text-white flex flex-col z-30 transform transition-transform duration-300 shadow-xl ${isSidebarOpen ? 'translate-x-0' : '-translate-x-full md:translate-x-0'}`}>
        <div className="p-6 border-b border-[#122b5e]">
          <h1 className="text-xl font-bold flex items-center gap-3 text-white"><div className="p-2 bg-gradient-to-br from-[#d4af37] to-[#b8952b] rounded-lg shadow-md"><CalendarDays className="text-[#081a39]" size={24}/></div> BHS Scheduler</h1>
        </div>
        <nav className="flex-1 overflow-y-auto py-6 custom-scrollbar space-y-1 px-3">
          {TABS.map(tab => {
            const IconComponent = tab.icon;
            return <button key={tab.id} onClick={() => { setActiveTab(tab.id); setIsSidebarOpen(false); }} className={`w-full flex items-center gap-3 px-4 py-3 rounded-xl transition-all font-medium ${activeTab === tab.id ? 'bg-gradient-to-r from-[#d4af37] to-[#e6c55c] text-[#081a39] shadow-md font-bold' : 'text-slate-300 hover:text-white hover:bg-white/10'}`}><IconComponent size={20} className={activeTab === tab.id ? "text-[#081a39]" : "opacity-80"} /> {tab.name}</button>;
          })}
        </nav>
      </aside>

      <main className="flex-1 flex flex-col min-w-0 overflow-hidden relative">
        <header className="md:hidden bg-white border-b border-[#081a39]/10 p-3 flex items-center shadow-sm z-10 no-print">
          <button onClick={() => setIsSidebarOpen(true)} className="p-2 text-[#081a39] hover:bg-slate-100 rounded-lg flex items-center gap-2 font-bold transition-colors"><Menu size={24} /> เมนู</button>
        </header>
        <div className="flex-1 overflow-y-auto p-4 sm:p-8 custom-scrollbar">
          {activeTab === 'dashboard' && <Dashboard {...commonProps} />}
          {activeTab === 'settings' && <SettingsView {...commonProps} handleRestoreBackup={handleRestoreBackup} />}
          {activeTab === 'teachers' && <TeachersView {...commonProps} />}
          {activeTab === 'subjects' && <SubjectsView {...commonProps} importBasicSubjects={()=>{}} />}
          {activeTab === 'classrooms' && <ClassroomsView {...commonProps} autoGenerateClassrooms={autoGenerateClassrooms} />}
          {activeTab === 'loads' && <LoadsView {...commonProps} />}
          {activeTab === 'constraints' && <ConstraintsView {...commonProps} />}
          {activeTab === 'schedule' && <ScheduleView {...commonProps} runSmartSchedule={runSmartSchedule} isScheduling={isScheduling} />}
          {activeTab === 'reports' && <ReportsView {...commonProps} />}
          {activeTab === 'print' && <PrintView {...commonProps} />}
        </div>
      </main>

      <Modal isOpen={autoScheduleIssues.length > 0} onClose={() => setAutoScheduleIssues([])} title="ตรวจข้อมูลก่อนจัดตารางอัตโนมัติ">
        <div className="space-y-4">
          <p className="text-sm text-slate-700">พบข้อมูลที่ต้องแก้ก่อนเริ่มจัดตาราง กรุณาตรวจสอบรายการด้านล่างแล้วลองกดจัดตารางอีกครั้ง ระบบยังไม่ได้เปลี่ยนตารางเดิม</p>
          <ul className="list-disc pl-5 space-y-2 text-sm text-rose-700" role="alert">
            {autoScheduleIssues.map((issue, index) => <li key={index}>{issue}</li>)}
          </ul>
          {autoScheduleIssues.length === 30 && <p className="text-xs text-amber-700">ระบบแสดงสูงสุด 30 รายการ กรุณาแก้ไขแล้วตรวจอีกครั้ง</p>}
          <button onClick={() => setAutoScheduleIssues([])} className="w-full py-3 rounded-xl bg-[#081a39] text-white font-bold">รับทราบ</button>
        </div>
      </Modal>

      <Modal isOpen={!!scheduleResult} onClose={() => setScheduleResult(null)} title="ผลการจัดตารางอัตโนมัติ (AI Preview)">
        {scheduleResult && (
          <div className="space-y-6">
            <div className="grid grid-cols-2 gap-4">
              <div className="bg-emerald-50 border border-emerald-200 p-4 rounded-xl text-center"><div className="text-emerald-700 font-bold text-sm mb-1">จัดลงสำเร็จ</div><div className="text-2xl font-bold text-emerald-800">{scheduleResult.stats.scheduled} <span className="text-sm">/ {scheduleResult.stats.totalReq} คาบ</span></div></div>
              <div className={`border p-4 rounded-xl text-center ${scheduleResult.stats.missing > 0 ? 'bg-rose-50 border-rose-200' : 'bg-slate-50 border-slate-200'}`}><div className={`font-bold text-sm mb-1 ${scheduleResult.stats.missing > 0 ? 'text-rose-700' : 'text-slate-600'}`}>จัดไม่ได้</div><div className={`text-2xl font-bold ${scheduleResult.stats.missing > 0 ? 'text-rose-800' : 'text-slate-800'}`}>{scheduleResult.stats.missing} <span className="text-sm">คาบ</span></div></div>
            </div>
            {scheduleResult.unassigned.length > 0 && (
              <div className="border border-rose-200 rounded-xl overflow-hidden">
                <div className="bg-rose-50 p-3 font-bold text-rose-700 flex items-center gap-2 border-b border-rose-200"><AlertTriangle size={18}/> รายการที่จัดไม่ได้ (รอแก้ไข/จัดด้วยมือ)</div>
                <div className="max-h-48 overflow-y-auto bg-white p-3 space-y-2">
                  {scheduleResult.unassigned.map((u, i) => (
                    <div key={i} className="text-sm flex justify-between border-b pb-2 last:border-0 last:pb-0"><div><span className="font-bold text-[#081a39]">{getSubject(u.subjectId).name}</span> ห้อง {getClassroom(u.classroomId).name}</div><div className="text-rose-600 text-right">ขาด {u.missing} คาบ<div className="text-[10px] text-slate-500">{u.reason}</div></div></div>
                  ))}
                </div>
              </div>
            )}
            <div className="flex gap-3 pt-4 border-t border-slate-100"><button onClick={() => setScheduleResult(null)} className="flex-1 py-3 bg-slate-100 hover:bg-slate-200 text-[#081a39] rounded-xl font-bold transition-all">ยกเลิก (คงตารางเดิม)</button><button onClick={commitSchedule} className="flex-1 py-3 bg-[#081a39] hover:bg-[#122b5e] text-white rounded-xl font-bold shadow-md flex justify-center items-center gap-2"><Check size={18}/> ใช้งานตารางนี้</button></div>
          </div>
        )}
      </Modal>

      {toast && (
        <div className={`fixed bottom-6 right-6 px-6 py-4 rounded-xl shadow-xl border flex items-center gap-3 animate-in slide-in-from-bottom-5 z-[70] font-medium ${toast.type === 'error' ? 'bg-rose-50 text-rose-700 border-rose-200' : 'bg-[#081a39] text-[#d4af37] border-[#122b5e]'}`}>
          {toast.type === 'error' ? <AlertCircle size={22} /> : <CheckCircle2 size={22} />} {String(toast.msg)}
        </div>
      )}
      <ConfirmModal confirmData={confirmData} setConfirmData={setConfirmData} />
    </div>
  );
}
