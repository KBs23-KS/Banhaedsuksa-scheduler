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

const getFirebaseConfig = () => ({
  apiKey: import.meta.env.VITE_FIREBASE_API_KEY,
  authDomain: import.meta.env.VITE_FIREBASE_AUTH_DOMAIN,
  projectId: import.meta.env.VITE_FIREBASE_PROJECT_ID,
  storageBucket: import.meta.env.VITE_FIREBASE_STORAGE_BUCKET,
  messagingSenderId: import.meta.env.VITE_FIREBASE_MESSAGING_SENDER_ID,
  appId: import.meta.env.VITE_FIREBASE_APP_ID,
  measurementId: import.meta.env.VITE_FIREBASE_MEASUREMENT_ID
});

const app = initializeApp(getFirebaseConfig());
const auth = getAuth(app);
const db = getFirestore(app);
const appId = 'bhs-school-scheduler-v3'; 

// STAGE 27: ภาษาไทยสำหรับข้อความแจ้งเตือนที่ผู้ใช้มองเห็นเท่านั้น
// คงรหัสเหตุผลเดิมไว้สำหรับการตรวจข้อบังคับและการบันทึกข้อมูล
// เพิ่มรหัส/คำแนะนำใหม่ที่นี่ แทนการแก้คำแปลแยกแต่ละหน้าจอ
const USER_ERROR_TH = Object.freeze({
  MISSING_SOURCE: 'ไม่พบคาบเรียนต้นทาง อาจมีการลบหรือเปลี่ยนตารางจากเครื่องอื่น กรุณาโหลดตารางใหม่แล้วลองอีกครั้ง',
  LOCKED_SOURCE: 'คาบเรียนต้นทางถูกล็อกไว้ กรุณาปลดล็อกก่อนย้ายคาบ',
  LOCKED_PARTNER: 'คาบคู่ที่ต้องย้ายพร้อมกันถูกล็อกไว้ กรุณาปลดล็อกคาบคู่ก่อน',
  INVALID_DAY: 'วันที่เลือกไม่ถูกต้อง กรุณาเลือกวันเรียนใหม่',
  INVALID_PERIOD: 'คาบปลายทางไม่ถูกต้องหรือไม่มีอยู่ในตาราง กรุณาเลือกคาบใหม่',
  INVALID_SOURCE_PERIOD: 'ไม่พบคาบต้นทางในช่วงเวลาเรียนปัจจุบัน กรุณาโหลดตารางใหม่',
  INVALID_BLOCK_PAIR: 'วิชาคาบคู่ต้องอยู่ในคาบเรียนที่ติดกัน และห้ามข้ามคาบพัก กรุณาเลือกตำแหน่งใหม่',
  BREAK_PERIOD: 'คาบนี้เป็นเวลาพักกลางวัน ไม่สามารถจัดวิชาเรียนได้',
  ROOM_UNAVAILABLE: 'ห้องเรียนนี้ถูกปิดการจัดตารางอัตโนมัติในช่วงเวลาที่เลือก หากจำเป็นสามารถลงด้วยมือและยืนยันข้อยกเว้นได้',
  TEACHER_UNAVAILABLE: 'ครูไม่ว่างในช่วงเวลานี้ หรือมี PLC กรุณาเลือกคาบอื่น',
  FIXED_SCHEDULE_CONFLICT: 'ช่วงเวลานี้มีกิจกรรมล็อกตายตัวอยู่แล้ว กรุณาตรวจรายการกิจกรรมก่อน',
  ROOM_CONFLICT: 'ห้องเรียนมีคาบเรียนหรือกิจกรรมอื่นในช่วงเวลาปลายทางแล้ว กรุณาเลือกคาบว่างหรือย้ายรายการเดิมออกก่อน',
  TEACHER_CONFLICT: 'ครูมีคาบสอนหรือกิจกรรมอื่นในช่วงเวลาเดียวกัน กรุณาเลือกเวลาที่ครูว่าง',
  TEACHER_DAILY_LIMIT: 'ครูมีภาระสอนและกิจกรรมครบเพดานของวันนั้นแล้ว กรุณากระจายคาบไปวันอื่น',
  TEACHER_CONSECUTIVE_LIMIT: 'คาบสอนรายวิชาติดต่อกันเกินเงื่อนไข กรุณาเว้นคาบหรือเลือกช่วงเวลาอื่น',
  TEACHER_FOUR_PERIOD_PATTERN: 'การกระจายคาบสอน 4 คาบของครูไม่ผ่านรูปแบบที่กำหนด กรุณาปรับตำแหน่งคาบ',
  SUBJECT_DAILY_LIMIT: 'รายวิชานี้เกินจำนวนคาบสูงสุดต่อวันของห้องเรียน กรุณาเลือกวันอื่น',
  CUSTOM_RULE_BLOCKED: 'ขัดกับเงื่อนไขจัดตารางของครู รายวิชา หรือห้องเรียน กรุณาตรวจเมนูข้อจำกัดและเงื่อนไข',
  DOUBLE_PERIOD_NOT_AVAILABLE: 'ไม่พบคาบว่างติดกัน 2 คาบสำหรับวิชาคาบคู่ กรุณาตรวจคาบล็อกหรือสลับคาบอื่น',
  BLOCK_COMPANION_NOT_AVAILABLE: 'ไม่พบคาบที่ติดกับคาบเดิมเพื่อจับคู่วิชา กรุณาตรวจคาบข้างเคียง',
  NO_AVAILABLE_SLOT: 'ไม่พบคาบว่างที่ผ่านเงื่อนไขทั้งหมด กรุณาตรวจข้อจำกัดหรือปรับคาบที่จัดไว้',
  SUBJECT_ID_TAKEN: 'รหัสรายวิชานี้มีอยู่ในระบบแล้ว กรุณาใช้รหัสอื่น',
  SUBJECT_DATA_CHANGED: 'ข้อมูลรายวิชาเปลี่ยนไปจากเครื่องอื่น กรุณาโหลดข้อมูลใหม่แล้วลองอีกครั้ง',
  SUBJECT_REFS_CHANGED: 'ข้อมูลที่เชื่อมโยงกับรายวิชาเปลี่ยนไป กรุณาโหลดข้อมูลใหม่ก่อนดำเนินการ',
  SUBJECT_IN_USE: 'รายวิชานี้ถูกใช้อยู่ในภาระสอน ตาราง หรือกิจกรรม จึงยังลบไม่ได้',
  NOT_SIGNED_IN: 'ยังไม่ได้เชื่อมต่อระบบ กรุณาตรวจการเข้าสู่ระบบแล้วลองอีกครั้ง',
  TIME_RULE_EXISTS: 'รายวิชานี้มีเงื่อนไขช่วงเวลาอยู่แล้ว กรุณาตรวจรายการเดิม',
  FIXED_SELECTION_EMPTY: 'ยังไม่ได้เลือกกิจกรรมที่ต้องการจัดการ กรุณาเลือกรายการก่อน',
  FIXED_SELECTION_CHANGED: 'รายการกิจกรรมที่เลือกเปลี่ยนไป กรุณาโหลดข้อมูลใหม่แล้วเลือกอีกครั้ง',
  FIXED_TERM_CHANGED: 'ภาคเรียนเปลี่ยนไประหว่างดำเนินการ กรุณาโหลดใหม่',
  FIXED_DATA_CHANGED: 'ข้อมูลกิจกรรมเปลี่ยนไปจากเครื่องอื่น กรุณาตรวจและเลือกใหม่',
  FIXED_HISTORY_FAILED: 'บันทึกประวัติก่อนลบกิจกรรมไม่สำเร็จ ระบบยังไม่ได้ลบข้อมูล',
  BACKUP_SETTINGS_MISSING: 'ไม่พบการตั้งค่าภาคเรียนที่ต้องใช้สำรองข้อมูล กรุณาตรวจการตั้งค่า',
  BACKUP_TERM_CHANGED: 'ภาคเรียนเปลี่ยนระหว่างสำรองข้อมูล กรุณาเริ่มสำรองใหม่',
  BACKUP_SETTINGS_CHANGED: 'การตั้งค่าโรงเรียนเปลี่ยนระหว่างสำรองข้อมูล กรุณาเริ่มใหม่',
  SNAPSHOT_ROW_TOO_LARGE: 'ข้อมูลประวัติบางรายการมีขนาดใหญ่เกินไป ยังบันทึกประวัติไม่สำเร็จ',
  UNSUPPORTED_HISTORY_FORMAT: 'รูปแบบประวัติตารางนี้ยังไม่รองรับ กรุณาใช้รายการประวัติอื่น',
  RESTORE_SNAPSHOT_FAILED: 'ไม่สามารถสร้างจุดย้อนกลับก่อนกู้คืนได้ ระบบจึงยังไม่เปลี่ยนข้อมูล',
  INVALID_REPLACEMENT_ITEMS: 'ข้อมูลชุดที่จะนำเข้าไม่ถูกต้อง ระบบยังไม่ได้เขียนทับ',
  INVALID_REPLACEMENT_ITEM: 'พบข้อมูลรายการที่นำเข้าไม่ถูกต้อง ระบบยังไม่ได้เขียนทับ',
  NO_ACTIVE_PERIODS: 'ไม่พบคาบเรียนที่เปิดใช้งาน กรุณาตรวจการตั้งค่าคาบเรียน',
  SCHEDULER_LOCK_LOST: 'สิทธิ์ล็อกการจัดตารางหมดอายุหรือเปลี่ยนไป ระบบหยุดทำงานเพื่อป้องกันตารางชน กรุณาลองใหม่',
  SUBJECT_DAILY_LIMIT_FAILED: 'ผลจัดตารางไม่ผ่านเงื่อนไขจำนวนคาบรายวิชาต่อวัน ระบบยังไม่บันทึกผล',
  TEACHER_HARD_LIMIT_FAILED: 'ผลจัดตารางไม่ผ่านข้อบังคับภาระสอนครู ระบบยังไม่บันทึกผล',
  STALE_PREVIEW: 'มีข้อมูลตารางเปลี่ยนไประหว่างดูตัวอย่าง กรุณาจัดตารางใหม่ก่อนบันทึก',
  ATOMIC_LIMIT: 'รายการที่จะบันทึกมีจำนวนมากเกินข้อจำกัดต่อครั้ง กรุณาแบ่งข้อมูลเป็นชุดเล็กลง',
  ATOMIC_SIZE: 'ข้อมูลที่จะบันทึกมีขนาดใหญ่เกินข้อจำกัดต่อครั้ง กรุณาแบ่งข้อมูลเป็นชุดเล็กลง',
});
const FIREBASE_ERROR_TH = Object.freeze({
  'auth/invalid-api-key': 'ค่า Firebase API Key ไม่ถูกต้อง กรุณาตรวจ Environment Variables บน Vercel แล้วเผยแพร่เว็บไซต์ใหม่',
  'auth/operation-not-allowed': 'ยังไม่ได้เปิดวิธีเข้าสู่ระบบแบบไม่ระบุตัวตน (Anonymous) ใน Firebase Authentication',
  'auth/network-request-failed': 'เชื่อมต่อระบบยืนยันตัวตนไม่ได้ กรุณาตรวจอินเทอร์เน็ต',
  'auth/unauthorized-domain': 'โดเมนเว็บไซต์นี้ยังไม่ได้รับอนุญาตใน Firebase Authentication',
  'auth/too-many-requests': 'มีคำขอเข้าสู่ระบบถี่เกินไป กรุณารอสักครู่แล้วลองใหม่',
  'permission-denied': 'ไม่มีสิทธิ์อ่านหรือเขียนข้อมูล กรุณาตรวจ Firebase Security Rules',
  'unauthenticated': 'ยังไม่ได้เข้าสู่ระบบ กรุณาโหลดหน้าใหม่แล้วลองอีกครั้ง',
  'unavailable': 'ไม่สามารถเชื่อมต่อฐานข้อมูลได้ในขณะนี้ กรุณาตรวจอินเทอร์เน็ตแล้วลองใหม่',
  'resource-exhausted': 'ระบบได้รับคำขอมากเกินไป กรุณารอสักครู่แล้วลองใหม่',
  'deadline-exceeded': 'การเชื่อมต่อใช้เวลานานเกินไป กรุณาลองอีกครั้ง',
  'failed-precondition': 'ข้อมูลหรือการตั้งค่าบางอย่างยังไม่พร้อม กรุณาตรวจสอบก่อนทำรายการ',
  'not-found': 'ไม่พบข้อมูลที่ต้องการ อาจถูกลบหรือเปลี่ยนแปลงแล้ว กรุณาโหลดข้อมูลใหม่',
});
const translateUserMessage = (message) => {
  const raw = String(message?.message ?? message ?? '').trim();
  if (!raw) return 'ไม่สามารถดำเนินการได้ กรุณาลองอีกครั้ง';
  if (/^Invalid Format$/i.test(raw)) return 'รูปแบบไฟล์ไม่ถูกต้อง กรุณาเลือกไฟล์สำรองจากระบบนี้';
  if (/^Missing document id in /i.test(raw)) return 'รายการข้อมูลนำเข้าไม่มีรหัสเอกสารครบถ้วน กรุณาตรวจไฟล์ก่อนนำเข้า';
  if (/^Duplicate document id /i.test(raw)) return 'พบรหัสเอกสารซ้ำในข้อมูลนำเข้า กรุณาตรวจไฟล์ก่อนนำเข้า';
  // แปลงข้อผิดพลาดของ Firebase ก่อนแปลงรหัสของระบบจัดตาราง
  const firebaseTranslated = raw.replace(/(?:FirebaseError:\s*)?Firebase:\s*Error\s*\(([a-z-]+\/[a-z-]+)\)\.?/gi,
    (_, code) => FIREBASE_ERROR_TH[code.toLowerCase()] || 'เชื่อมต่อ Firebase ไม่สำเร็จ กรุณาตรวจการตั้งค่าและลองใหม่');
  const withFirebaseCodes = firebaseTranslated.replace(/\b(auth\/[a-z-]+|permission-denied|unauthenticated|unavailable|resource-exhausted|deadline-exceeded|failed-precondition|not-found)\b/gi,
    (code) => FIREBASE_ERROR_TH[code.toLowerCase()] || 'ระบบเชื่อมต่อไม่สำเร็จ กรุณาลองใหม่');
  return withFirebaseCodes.replace(/\b[A-Z][A-Z0-9]*(?:_[A-Z0-9]+)+\b/g,
    code => USER_ERROR_TH[code] || 'เกิดข้อผิดพลาดในการดำเนินการ กรุณาลองใหม่หรือติดต่อผู้ดูแลระบบ')
    .replace(/^ย้ายไม่ได้\s*:\s*/, 'ไม่สามารถย้ายคาบเรียนได้: ')
    .replace(/^ย้ายคาบไม่สำเร็จ\s*:\s*/, 'ไม่สามารถย้ายคาบเรียนได้: ');
};


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
// One subject-specific rule can cover every teaching period of a weekday:
// block = forbid that day; prefer = bonus for lessons on that day (any period).
// Store a single marker instead of seven records in the existing Firebase collection.
const ALL_DAY_PERIOD_ID = '__ALL_DAY__';
// STEP 10: กิจกรรมล็อกทุกห้องเก็บเป็นเอกสารเดียว (ครูผู้ดูแลไม่นับซ้ำหลายห้อง)
const ALL_CLASSROOMS_ID = '__ALL_CLASSROOMS__';
// Stage 30: saved in existing fixedSchedules collection, so backups, PLC, history and Auto
// continue using their existing locking semantics. Old fixed rows remain 'general'.
const isDevelopmentActivity = item => item?.activityType === 'development';
const fixedActivityDisplayName = (item, getSubject, groups = [], roomId = '') => {
  const local = roomId && (groups || []).find(g => g.fixedId === item?.id && g.classroomId === roomId);
  return String(local?.name || item?.activityName || (item?.subjectId && getSubject(item.subjectId)?.name) || 'กิจกรรม').trim();
};
const plcDisplayName = subjectArea => `PLC กลุ่มสาระฯ ${subjectArea || '-'}`;

// New activity support: one locked activity can reserve many teachers, without
// making duplicate student lessons or counting the same teacher once per room.
const fixedActivityTeacherIds = item => [...new Set([
  ...(item?.teacherId ? [item.teacherId] : []),
  ...(Array.isArray(item?.teacherIds) ? item.teacherIds : [])
].filter(Boolean))];
const activityTeacherEntries = (activityGroups, fixedSchedules) => {
  const byId = new Map((fixedSchedules || []).map(item => [item.id, item]));
  const extraFixed = (fixedSchedules || []).flatMap(f => fixedActivityTeacherIds(f)
    .filter(id => id !== f.teacherId).map(id => ({
      id:`fixed_teacher_${f.id}_${id}`, teacherId:id, day:f.day, periodId:f.periodId,
      subjectId:f.subjectId, classroomId:f.classroomId, scopeLabel:f.scopeLabel || '',activityType:f.activityType,
      __activityTeacherOnly:true, __activityName:f.activityName || ''
    })));
  const grouped = (activityGroups || []).filter(item => byId.has(item.fixedId) &&
    fixedAppliesToClassroom(byId.get(item.fixedId), item.classroomId)).flatMap(item => {
    const f = byId.get(item.fixedId);
    return fixedActivityTeacherIds(item).filter(id => !fixedActivityTeacherIds(f).includes(id)).map(id => ({
      id:`activity_${item.id}_${id}`, teacherId:id, day:f.day, periodId:f.periodId,
      subjectId:f.subjectId, classroomId:item.classroomId, activityType:f.activityType, __activityTeacherOnly:true,
      __activityName:item.name
    }));
  });
  return [...extraFixed, ...grouped];
};
const classroomGrade = room => {
  const name = String(room?.name || '').trim();
  const m = name.match(/(?:ม\.?\s*|มัธยมศึกษาปีที่\s*)([1-6])(?:\D|$)/i);
  return m ? `ม.${m[1]}` : '';
};
// A manual exception is valid only for the same occupied slot in a closed room.
const isApprovedManualRoomException = (item, unavailable) =>
  item?.source === 'manual' && item?.manualRoomClosureOverride === true &&
  item.classroomId === unavailable.classroomId && item.day === unavailable.day &&
  String(item.periodId) === String(unavailable.periodId);

const fixedAppliesToClassroom = (fixedItem, classroomId) => {
  if (!fixedItem || !classroomId) return false;
  if (fixedItem.classroomId !== ALL_CLASSROOMS_ID) return fixedItem.classroomId === classroomId;
  if (Array.isArray(fixedItem.includedClassroomIds)) return fixedItem.includedClassroomIds.includes(classroomId);
  return !(fixedItem.excludedClassroomIds || []).includes(classroomId);
};
const fixedTargetRoomIds = (fixedItem, classrooms) => (classrooms || [])
  .filter(room => fixedAppliesToClassroom(fixedItem, room.id)).map(room => room.id);
// Pure planner for single-day closure. The caller commits all returned slots atomically.
const planClassroomDayClosure = ({classroomId,day,periods,schedules=[],fixedSchedules=[],classroomUnavailabilities=[]}) => {
  if (!classroomId || !DAYS.includes(day)) return {ok:false, reason:'กรุณาระบุห้องและวันให้ถูกต้อง', periodIds:[]};
  if (schedules.some(s=>s.classroomId===classroomId && s.day===day) ||
      fixedSchedules.some(f=>f.day===day && fixedAppliesToClassroom(f,classroomId)))
    return {ok:false,reason:'มีคาบเรียนหรือกิจกรรมที่ล็อกไว้ในวันนี้อยู่แล้ว กรุณาตรวจตารางก่อน',periodIds:[]};
  const existing=new Set(classroomUnavailabilities.filter(x=>x.classroomId===classroomId&&x.day===day).map(x=>String(x.periodId)));
  return {ok:true,reason:'',periodIds:(periods||[]).filter(p=>!p.isBreak&&!existing.has(String(p.id))).map(p=>p.id)};
};
// One teacher/day/period = one load slot regardless of report/table representation.
// Stage 26: กิจกรรมล็อกมีครูรับผิดชอบ = ภาระประจำวัน แต่ไม่ใช่คาบสอนรายวิชา
// ครูรับผิดชอบหลายห้องพร้อมกันให้นับเพียงหนึ่งคาบต่อช่วงเวลา
const isTeacherActivityDuty = item => item?.__isFixedActivity === true ||
  item?.__activityTeacherOnly === true || item?.__kind === 'กิจกรรมล็อก';
const markFixedTeacherActivities = items => (items || []).map(item => ({...item,__isFixedActivity:true}));
const buildTeacherOccupancy = (schedules = [], fixedSchedules = [], activityGroups = []) =>
  [...schedules, ...markFixedTeacherActivities(fixedSchedules), ...activityTeacherEntries(activityGroups, fixedSchedules)]
    .filter(item => Boolean(item.teacherId));
const summarizeTeacherWeek = (items, periods) => {
  const days = new Map();
  items.forEach(item => {
    if (!item.teacherId || !DAYS.includes(item.day)) return;
    const key = teacherDayKey(item.teacherId, item.day);
    if (!days.has(key)) days.set(key, { teacherId: item.teacherId, day: item.day, ids: new Set() });
    days.get(key).ids.add(item.periodId);
  });
  const byTeacher = new Map();
  days.forEach(row => {
    if (!byTeacher.has(row.teacherId)) byTeacher.set(row.teacherId,{total:0, fiveDays:[],heavyDays:0,dayCounts:{},longestConsecutive:0});
    const stat = byTeacher.get(row.teacherId);
    const count = row.ids.size;
    stat.total += count;
    stat.dayCounts[row.day] = count;
    stat.longestConsecutive = Math.max(stat.longestConsecutive,calculateDayTeachingShape([...row.ids],periods).maxConsec);
    if (count >= 4) stat.heavyDays++;
    if (count === 5) stat.fiveDays.push(row.day);
  });
  return byTeacher;
};
const SUBJECT_AREAS = [
  "ไม่มีกลุ่มสาระการเรียนรู้", "ภาษาไทย", "คณิตศาสตร์", "วิทยาศาสตร์และเทคโนโลยี", 
  "สังคมศึกษา ศาสนา และวัฒนธรรม", "สุขศึกษาและพลศึกษา", "ศิลปะ", "การงานอาชีพ", "ภาษาต่างประเทศ"
];

// STAGE 21 — PLC is a shared teacher-only reservation, NOT a teaching load.
// Membership is calculated from the teachers' current subject area; a published copy
// receives the subject-area catalog frozen at the time of publication.
const PLC_AREAS = SUBJECT_AREAS.filter(area => area !== 'ไม่มีกลุ่มสาระการเรียนรู้');
const plcIdForArea = area => `PLC_${PLC_AREAS.indexOf(area)}`;
const plcTeachersForArea = (area, teachers = []) => teachers.filter(t => t.subjectArea === area);
const buildPLCReservations = (plcSessions = [], teachers = []) => (plcSessions || []).flatMap(session =>
  plcTeachersForArea(session.subjectArea, teachers).map(t => ({
    id:`plc_${session.id}_${t.id}`, plcSessionId:session.id, teacherId:t.id,
    subjectArea:session.subjectArea, day:session.day, periodId:session.periodId,
    __isPLC:true
  })));
const plcSlotIssues = ({ subjectArea, day, periodId, teachers = [], schedules = [],
  fixedSchedules = [], activityGroups = [], unavailabilities = [], plcSessions = [], periods = [] }) => {
  const members = plcTeachersForArea(subjectArea,teachers);
  if (!members.length) return [{reason:'ไม่พบครูในกลุ่มสาระนี้'}];
  if (!DAYS.includes(day) || !periods.some(p=>String(p.id)===String(periodId)&&!p.isBreak))
    return [{reason:'วันหรือคาบไม่ถูกต้อง'}];
  const busy = [...schedules,...fixedSchedules,...activityTeacherEntries(activityGroups,fixedSchedules)];
  return members.flatMap(teacher => {
    const issues=[];
    if (busy.some(x => x.teacherId===teacher.id&&x.day===day&&String(x.periodId)===String(periodId)))
      issues.push({teacherId:teacher.id,teacherName:teacher.name,reason:'มีสอนหรือกิจกรรมล็อก'});
    if (unavailabilities.some(x=>x.teacherId===teacher.id&&x.day===day&&String(x.periodId)===String(periodId)))
      issues.push({teacherId:teacher.id,teacherName:teacher.name,reason:'ครูกำหนดว่าไม่สะดวก'});
    if (plcSessions.some(x=>x.subjectArea!==subjectArea&&x.day===day&&String(x.periodId)===String(periodId) &&
      plcTeachersForArea(x.subjectArea,teachers).some(t=>t.id===teacher.id)))
      issues.push({teacherId:teacher.id,teacherName:teacher.name,reason:'มี PLC ของอีกกลุ่ม'});
    return issues;
  });
};
const rankPLCSlots = (args) => DAYS.flatMap((day,dayIndex) =>
  (args.periods||[]).filter(p=>!p.isBreak).map((period,index)=>{
    const issues=plcSlotIssues({...args,day,periodId:period.id});
    // Fewer gaps for the shared group; don't privilege Friday last period.
    return {day,periodId:period.id,periodName:period.name,issues,
      score:100-issues.length*10000-Math.abs(index-3)*3-dayIndex};
  })).sort((a,b)=>a.issues.length-b.issues.length||b.score-a.score);

const SUBJECT_COLORS = {
  "ไม่มีกลุ่มสาระการเรียนรู้": "text-slate-500", "ภาษาไทย": "text-rose-600", "คณิตศาสตร์": "text-blue-600",
  "วิทยาศาสตร์และเทคโนโลยี": "text-emerald-600", "สังคมศึกษา ศาสนา และวัฒนธรรม": "text-amber-600", 
  "สุขศึกษาและพลศึกษา": "text-orange-500", "ศิลปะ": "text-purple-600", "การงานอาชีพ": "text-cyan-600",
  "ภาษาต่างประเทศ": "text-indigo-600"
};

// STAGE 2: ขีดจำกัดวิชาเดียวกันต่อหนึ่งวัน แยกแต่ละห้องเรียน
const subjectDailyKey = (classroomId, subjectId, day) => JSON.stringify([classroomId, subjectId, day]);
const registerSubjectDayPeriod = (maps, item) => {
  if (!item.subjectId || !item.classroomId) return;
  const key = subjectDailyKey(item.classroomId, item.subjectId, item.day);
  maps.subjectDayCounts.set(key, (maps.subjectDayCounts.get(key) || 0) + 1);
};
const canPlaceSubjectDailyEntries = (load, entries, maps) => {
  if (!load.subjectId || !load.classroomId) return { ok: true, reason: '' };
  const limits = (maps.customRules || []).filter(rule => rule.mode === 'dailyLimit' && rule.targetType === 'subject' && rule.targetId === load.subjectId);
  // STAGE 29: สำหรับ Auto รายวิชาคาบเดี่ยว (periodBlock !== 2) ใช้ได้เพียง
  // 1 คาบ/วัน/รหัส/ห้อง ส่วนวิชาคาบคู่ได้สูงสุด 2 คาบ/วัน/รหัส/ห้อง
  // นับคาบที่ผู้จัดตารางลงเองและคาบที่ล็อกไว้ก่อนด้วย แต่ไม่เปลี่ยนกฎลงมือเอง
  if (!limits.length && !maps.enforceAutoSubjectDailyCap) return { ok: true, reason: '' };
  const autoCap = maps.autoSubjectBlockSizes?.get(String(load.subjectId)) === 2 ? 2 : 1;
  const maxPerDay = Math.min(
    ...(limits.map(rule => Number(rule.maxPerDay))),
    ...(maps.enforceAutoSubjectDailyCap ? [autoCap] : [])
  );
  const perDay = new Map();
  entries.forEach(entry => perDay.set(entry.day, (perDay.get(entry.day) || 0) + 1));
  for (const [day, proposed] of perDay) {
    const regular = maps.subjectDayCounts.get(subjectDailyKey(load.classroomId, load.subjectId, day)) || 0;
    // กิจกรรมล็อกทุกห้องต้องนับเป็นหนึ่งคาบในทุกห้อง
    const common = load.classroomId === ALL_CLASSROOMS_ID ? 0 :
      (maps.subjectDayCounts.get(subjectDailyKey(ALL_CLASSROOMS_ID, load.subjectId, day)) || 0);
    if (regular + common + proposed > maxPerDay) return { ok: false, reason: 'SUBJECT_DAILY_LIMIT' };
  }
  return { ok: true, reason: '' };
};
const findSubjectDailyLimitConflicts = (items, classrooms, customRules) => {
  const rules = (customRules || []).filter(r => r.mode === 'dailyLimit');
  if (!rules.length) return [];
  const maps = { subjectDayCounts: new Map(), customRules: rules };
  items.forEach(item => {
    if (item.classroomId === ALL_CLASSROOMS_ID) {
      fixedTargetRoomIds(item,classrooms).forEach(id => registerSubjectDayPeriod(maps,{...item,classroomId:id}));
    } else registerSubjectDayPeriod(maps,item);
  });
  const result = [];
  rules.forEach(rule => {
    const seenDays = new Set(items.filter(x => x.subjectId === rule.targetId).map(x => x.day));
    seenDays.forEach(day => {
      classrooms.forEach(room => {
        const regular = maps.subjectDayCounts.get(subjectDailyKey(room.id, rule.targetId, day)) || 0;
        if (regular > Number(rule.maxPerDay)) result.push(`วิชา ${rule.targetId} ห้อง ${room.id} วัน${day}: ${regular} คาบ เกินกำหนด ${rule.maxPerDay} คาบ/วัน`);
      });
    });
  });
  return [...new Set(result)];
};

// STAGE 30: Count timetable entries per subject / room / day, including
// room-wide fixed lessons where appropriate (never count development activities).
const countSubjectDayDistribution = (items = [], subjects = [], classrooms = []) => {
  const knownSubjects = new Set(subjects.map(s => String(s.id)));
  const counts = new Map();
  items.forEach(item => {
    if (!item?.subjectId || !item?.day || !knownSubjects.has(String(item.subjectId)) || isDevelopmentActivity(item)) return;
    const targets = item.classroomId === ALL_CLASSROOMS_ID ? fixedTargetRoomIds(item, classrooms) : [item.classroomId];
    targets.forEach(roomId => {
      if (!roomId) return;
      const key = subjectDailyKey(roomId, item.subjectId, item.day);
      if (!counts.has(key)) counts.set(key, {roomId,subjectId:String(item.subjectId),day:item.day,count:0});
      counts.get(key).count++;
    });
  });
  return counts;
};

// Auto must NOT create excess same-subject/day lessons, but pre-existing manual
// lessons and pinned anchors are deliberately preserved. The protected baseline
// must be supplied explicitly by the caller, never inferred from generated rows.
const findAutoSubjectDistributionConflicts = (items = [], subjects = [], classrooms = [], protectedAnchors = []) => {
  const blockBySubject = new Map(subjects.map(s => [String(s.id), Number(s.periodBlock) === 2 ? 2 : 1]));
  const roomNames = new Map(classrooms.map(c => [String(c.id), c.name || c.id]));
  const counts = countSubjectDayDistribution(items, subjects, classrooms);
  const baseline = countSubjectDayDistribution(protectedAnchors, subjects, classrooms);
  return [...counts.entries()].filter(([key,x]) =>
    x.count > Math.max(blockBySubject.get(x.subjectId), baseline.get(key)?.count || 0)
  ).map(([key,x]) => {
    const max = blockBySubject.get(x.subjectId);
    const desc = max === 1 ? 'ไม่ได้ตั้งเป็นคาบคู่ ต้องกระจายไปคนละวัน' : 'ตั้งเป็นคาบคู่ แต่ห้ามเกิน 2 คาบในวันเดียว';
    return `รหัสวิชา ${x.subjectId} ห้อง ${roomNames.get(String(x.roomId)) || x.roomId} วัน${x.day}: มี ${x.count} คาบ (${desc}) Auto ห้ามเพิ่มคาบจนเกินเกณฑ์เดิมหรือเกินคาบที่คงไว้`;
  });
};

// STAGE 1: กฎรายคาบที่ผู้ดูแลตั้งเอง (แยกตามภาคเรียน)
// block = ข้อบังคับ, prefer = ข้อแนะนำในการเลือกรอบจัดตาราง
const isValidCustomRule = (rule, teachers, subjects, classrooms, periods) => {
  const source = rule?.targetType === 'teacher' ? teachers : rule?.targetType === 'subject' ? subjects : rule?.targetType === 'classroom' ? classrooms : null;
  if (!source || !source.some(item => String(item.id) === String(rule.targetId))) return false;
  if (rule.mode === 'dailyLimit') return rule.targetType === 'subject' &&
    Number.isInteger(Number(rule.maxPerDay)) && Number(rule.maxPerDay) >= 1 &&
    Number(rule.maxPerDay) <= periods.filter(p => !p.isBreak).length;
  // Stage 22: subject-wide morning/afternoon condition, across all weekdays.
  if (rule.mode === 'timeWindow') return rule.targetType === 'subject' &&
    ['morning','afternoon'].includes(rule.timeWindow) &&
    ['required','preferred'].includes(rule.strictness);
  if (!['block', 'prefer'].includes(rule.mode) || !DAYS.includes(rule.day)) return false;
  // Both subject rules may target a weekday without specifying a period.
  // Block means no lesson may occur on that day; prefer awards Auto points on that day.
  // Teacher and classroom conditions still require a specific period.
  if (String(rule.periodId) === ALL_DAY_PERIOD_ID)
    return rule.targetType === 'subject' && periods.some(p => !p.isBreak);
  return periods.some(p => !p.isBreak && String(p.id) === String(rule.periodId));
};
const customRuleMatches = (rule, item) => !!(rule && item && rule.day === item.day &&
  (String(rule.periodId) === ALL_DAY_PERIOD_ID
    ? rule.targetType === 'subject' && ['block', 'prefer'].includes(rule.mode)
    : String(rule.periodId) === String(item.periodId)) &&
  (rule.targetType === 'teacher' ? !!item.teacherId && rule.targetId === item.teacherId :
   rule.targetType === 'subject' ? !!item.subjectId && rule.targetId === item.subjectId :
   rule.targetType === 'classroom' ? (rule.targetId === item.classroomId || (item.classroomId === ALL_CLASSROOMS_ID && fixedAppliesToClassroom(item,rule.targetId))) : false));
// A time window is based on the actual break position, not the internal period ID.
// With default periods, morning = learning periods 1–4 and afternoon = 5–7.
const isSubjectInTimeWindow = (periodId, windowName, periodsOrContext = DEFAULT_PERIODS) => {
  const context = getPeriodContext(periodsOrContext);
  const found = context.activePeriods.find(p => String(p.id) === String(periodId));
  if (!found) return false;
  const morning = [...context.morningIds].some(id => String(id) === String(found.id));
  return windowName === 'morning' ? morning : !morning;
};
const blocksCustomRule = (rules, item, periodsOrContext = DEFAULT_PERIODS) => (rules || []).some(r =>
  (r.mode === 'block' && customRuleMatches(r, item)) ||
  (r.mode === 'timeWindow' && r.strictness === 'required' &&
    r.targetType === 'subject' && !!item?.subjectId && String(r.targetId) === String(item.subjectId) &&
    !isSubjectInTimeWindow(item.periodId, r.timeWindow, periodsOrContext)));
const customRulePreferenceScore = (rules, item, periodsOrContext = DEFAULT_PERIODS) => (rules || []).reduce((score, r) => {
  if (r.mode === 'prefer' && customRuleMatches(r, item)) return score + 18;
  if (r.mode === 'timeWindow' && r.strictness === 'preferred' &&
      r.targetType === 'subject' && !!item?.subjectId && String(r.targetId) === String(item.subjectId)) {
    return score + (isSubjectInTimeWindow(item.periodId, r.timeWindow, periodsOrContext) ? 24 : -24);
  }
  return score;
}, 0);

const SCHEDULER_CONFIG = { candidateCount: 50, maxConsecutiveTeacherPeriods: 2, maxTeacherPeriodsPerDay: 4, absoluteMaxTeacherPeriodsPerDay: 5, candidateChoicePool: 4, placementJitter: 2.5, yieldEveryCandidates: 5 };
// STAGE 09: รูปแบบครูที่มีสอน 4 คาบ/วัน ตามเกณฑ์ที่โรงเรียนยืนยัน (15 แนะนำ, 8 สำรอง, 12 ห้าม)
// ใช้เฉพาะกรณีมีคาบเรียนจริง 7 คาบและพักเที่ยงอยู่ระหว่างคาบเรียนที่ 4–5
// หมายเลขในรูปแบบเป็นลำดับคาบเรียน 1–7 ไม่ใช่ Firestore periodId (คาบ 5 ใช้ id '6')
const FOUR_PERIOD_PREFERRED_PATTERNS = new Set([
  '1,2,5,6', '1,2,5,7', '1,2,6,7', '1,3,5,6', '1,3,5,7', '1,3,6,7',
  '1,4,5,7', '1,4,6,7', '2,3,5,6', '2,3,5,7', '2,3,6,7', '2,4,5,7',
  '2,4,6,7', '3,4,5,7', '3,4,6,7',
]);
const FOUR_PERIOD_FALLBACK_PATTERNS = new Set([
  '1,2,4,5', '1,2,4,6', '1,2,4,7', '1,3,4,5',
  '1,3,4,6', '1,3,4,7', '1,4,5,6', '2,4,5,6',
]);
const FOUR_PERIOD_BLOCKED_PATTERNS = new Set([
  '1,2,3,4', '1,2,3,5', '1,2,3,6', '1,2,3,7',
  '1,5,6,7', '2,3,4,5', '2,3,4,6', '2,3,4,7',
  '2,5,6,7', '3,4,5,6', '3,5,6,7', '4,5,6,7',
]);
const getFourPeriodTeachingPattern = (dayPeriods, periodsOrContext) => {
  const ctx = getPeriodContext(periodsOrContext);
  // กรณีโรงเรียนแก้จำนวนคาบ/ตำแหน่งพักเที่ยง ไม่บังคับชุด 35 รูปแบบที่อิง 7 คาบ
  const morningCount = ctx.activePeriods.filter(p => ctx.morningIds.has(p.id)).length;
  if (ctx.activeOrder.length !== 7 || morningCount !== 4 || ctx.safePeriods.some(p => p.isBreak &&
    ctx.realIndex.get(p.id) < ctx.realIndex.get(ctx.activeOrder[3]))) return null;
  if (ctx.realIndex.get(ctx.activeOrder[4]) !== ctx.realIndex.get(ctx.activeOrder[3]) + 2) return null;
  const distinct = [...new Set(dayPeriods)].filter(id => ctx.activeIndex.has(id));
  if (distinct.length !== 4) return null;
  const key = distinct.map(id => ctx.activeIndex.get(id) + 1).sort((a,b)=>a-b).join(',');
  if (FOUR_PERIOD_PREFERRED_PATTERNS.has(key)) return { key, level: 'preferred' };
  if (FOUR_PERIOD_FALLBACK_PATTERNS.has(key)) return { key, level: 'fallback' };
  return { key, level: 'blocked' }; // 12 รูปแบบต้องห้าม รวม 3,4,5,6
};

// ชุดคาบที่กระจุกข้ามพักเที่ยง ใช้ตรวจทั้งวันที่สอน 4 และ 5 คาบ
const hasSchoolPeriodSequence = (dayPeriods, periodsOrContext, numbers) => {
  const ctx = getPeriodContext(periodsOrContext);
  if (ctx.activeOrder.length !== 7 || ctx.activePeriods.filter(p=>ctx.morningIds.has(p.id)).length !== 4) return false;
  const selected = new Set(dayPeriods);
  return numbers.every(n => selected.has(ctx.activeOrder[n-1]));
};

// STAGE 09.1: วันสอน 3 คาบควรกระจายคาบให้สมดุล (เงื่อนไขให้คะแนน ไม่ใช่ข้อห้ามเด็ดขาด)
// คาบ 4–5 ยังคงจัดคู่กันได้ แต่พยายามเลี่ยง 3–4–5 และ 4–5–6
// เพื่อไม่ให้กฎการให้คะแนนของ 3 คาบไปเปลี่ยนกติกา 35 รูปแบบของวันที่มี 4 คาบ
const getThreePeriodTeachingPattern = (dayPeriods, periodsOrContext) => {
  const ctx = getPeriodContext(periodsOrContext);
  const morningCount = ctx.activePeriods.filter(p => ctx.morningIds.has(p.id)).length;
  if (ctx.activeOrder.length !== 7 || morningCount !== 4) return null;
  if (ctx.realIndex.get(ctx.activeOrder[4]) !== ctx.realIndex.get(ctx.activeOrder[3]) + 2) return null;
  const distinct = [...new Set(dayPeriods)].filter(id => ctx.activeIndex.has(id));
  if (distinct.length !== 3) return null;
  const nums = distinct.map(id => ctx.activeIndex.get(id) + 1).sort((a,b) => a-b);
  const key = nums.join(',');
  if (key === '3,4,5' || key === '4,5,6') return { key, level: 'avoid' };
  // การเว้นคาบในช่วงเดียวกัน และมีสอนทั้งก่อนและหลังพักเที่ยง เป็นตัวเลือกแรก
  const sameSessionAdjacency = nums.filter((n,i) => i > 0 && n === nums[i-1] + 1 && n !== 5).length;
  const bothSessions = nums.some(n=>n<=4) && nums.some(n=>n>=5);
  if (sameSessionAdjacency === 0 && bothSessions) return { key, level: 'preferred' };
  return { key, level: 'fallback' };
};

// STAGE 10: คาบที่ 5 เป็นข้อยกเว้นที่ต้องแสดงให้ผู้ใช้ยืนยันก่อนบันทึก
const findTeacherFifthPeriodDays = (items, periods, teachers = []) => {
  const grouped = new Map();
  const names = new Map(teachers.map(t => [t.id, t.name]));
  (items || []).forEach(item => {
    if (!item.teacherId || !DAYS.includes(item.day)) return;
    const key = teacherDayKey(item.teacherId, item.day);
    if (!grouped.has(key)) grouped.set(key, { teacherId: item.teacherId, day: item.day, periodIds: new Set() });
    grouped.get(key).periodIds.add(String(item.periodId));
  });
  return [...grouped.values()].filter(x => x.periodIds.size === 5).map(x => ({
    teacherId: x.teacherId, teacherName: names.get(x.teacherId) || x.teacherId,
    day: x.day, count: x.periodIds.size,
    periodIds: [...x.periodIds],
  })).sort((a,b) => String(a.teacherName).localeCompare(String(b.teacherName), 'th') || DAYS.indexOf(a.day)-DAYS.indexOf(b.day));
};


// STAGE 11 - whole-school demand check: count room demand separately from teacher activity load.
// This is an arithmetic precheck, not a complete timetable solver.
const summarizeSchoolCapacity = ({classrooms, teachers, periods, classroomUnavailabilities=[], fixedSchedules=[], activityGroups=[], schedules=[]}) => {
  const active = periods.filter(p=>!p.isBreak);
  const daily = DAYS.map(day=>{
    let roomSlots=0, fixedRoomSlots=0, unavailableRoomSlots=0;
    classrooms.forEach(c=>active.forEach(p=>{
      if (classroomUnavailabilities.some(x=>x.classroomId===c.id && x.day===day && x.periodId===p.id)) unavailableRoomSlots++;
      else if (fixedSchedules.some(f=>f.day===day && f.periodId===p.id && fixedAppliesToClassroom(f,c.id))) fixedRoomSlots++;
      else roomSlots++;
    }));
    const activityTeacherSlots=buildTeacherOccupancy([],fixedSchedules,activityGroups)
      .filter(x=>x.day===day).length;
    return {day,roomSlots,fixedRoomSlots,unavailableRoomSlots,activityTeacherSlots,
      teacherFourCapacity:teachers.length*4,teacherFiveCapacity:teachers.length*5,
      potentialTotalTeacherDemand:roomSlots+activityTeacherSlots};
  });
  return {daily,totalRoomSlots:daily.reduce((a,d)=>a+d.roomSlots,0),
    totalActivityTeacherSlots:daily.reduce((a,d)=>a+d.activityTeacherSlots,0)};
};

// STAGE 15: Evaluate the actual teacher occupancy including regular lessons and
// activities. This metric ranks VALID candidate timetables; it never overrides hard rules.
const evaluateWeeklyTeachingBalance = (schedules, fixedSchedules, activityGroups, periods, teachers = []) => {
  const rows = summarizeTeacherWeek(buildTeacherOccupancy(schedules,fixedSchedules,activityGroups),periods);
  const teacherDays = [...rows.entries()];
  const occupancy=buildTeacherOccupancy(schedules,fixedSchedules,activityGroups);
  const details = teacherDays.map(([id,stat])=>({ teacherId:id,
    teacherName:teachers.find(t=>t.id===id)?.name || id,
    total:stat.total, heavyDays:stat.heavyDays, fiveDays:stat.fiveDays,
    dayCounts:stat.dayCounts,
    gaps:DAYS.reduce((sum,day)=>sum+calculateDayTeachingShape(
      occupancy.filter(x=>x.teacherId===id&&x.day===day).map(x=>x.periodId),periods).gaps,0),
    imbalance:(()=>{const counts=DAYS.map(d=>stat.dayCounts[d]||0);
      const avg=counts.reduce((a,b)=>a+b,0)/DAYS.length;
      return counts.reduce((total,c)=>total+(c-avg)**2,0);})(),
  }));
  // Five-period days should be very rare. The weekly distribution then decides
  // between equally complete, equally valid timetables.
  const penalty=details.reduce((n,r)=>n+r.fiveDays.length*180+r.heavyDays*12+r.imbalance*6+r.gaps*3,0);
  return {penalty,details,heavyDays:details.reduce((n,r)=>n+r.heavyDays,0),
    fiveDays:details.reduce((n,r)=>n+r.fiveDays.length,0),
    gaps:details.reduce((n,r)=>n+r.gaps,0)};
};

// Non-destructive repair hints: move one AUTO single-period slot, then test that
// the unassigned SINGLE-period load fits. Block-2 stays under the normal planner.
const suggestSingleSlotRepair = ({load,schedules,fixedSchedules,activityGroups,periods,plcSessions=[],
  unavailabilities,classroomUnavailabilities,customRules,subjects,classrooms,teachers,teachingLoads}) => {
  if (Number(subjects.find(x=>x.id===load.subjectId)?.periodBlock)===2) return [];
  const movable=(schedules||[]).filter(x=>isAutoGeneratedSchedule(x) && !x.keepOnRegenerate &&
    (x.classroomId===load.classroomId || x.teacherId===load.teacherId) &&
    Number(subjects.find(y=>y.id===x.subjectId)?.periodBlock)!==2).slice(0,24);
  const active=periods.filter(x=>!x.isBreak);
  const proposals=[];
  for(const candidate of movable) {
    const reduced=schedules.filter(x=>x.id!==candidate.id);
    const firstMaps=buildConstraintMaps(reduced,unavailabilities,classroomUnavailabilities,
      fixedSchedules,periods,customRules,activityGroups,classrooms,plcSessions,teachers);
    if (!canPlaceSchedule(load.teacherId,load.classroomId,candidate.day,candidate.periodId,
      firstMaps,false,load.subjectId).ok) continue;
    const inserted={...load,id:'preview_repair',day:candidate.day,periodId:candidate.periodId,source:'auto'};
    for(const day of DAYS) for(const p of active) {
      if(day===candidate.day && p.id===candidate.periodId) continue;
      const maps=buildConstraintMaps([...reduced,inserted],unavailabilities,classroomUnavailabilities,
        fixedSchedules,periods,customRules,activityGroups,classrooms,plcSessions,teachers);
      if (!canPlaceSchedule(candidate.teacherId,candidate.classroomId,day,p.id,maps,false,candidate.subjectId).ok) continue;
      const moved={...candidate,day,periodId:p.id};
      const trial=[...reduced,inserted,moved];
      const health=calculateScheduleHealth(trial,teachingLoads||[],unavailabilities,classroomUnavailabilities,
        fixedSchedules,teachers,subjects,classrooms,periods,customRules,activityGroups,plcSessions);
      if(health.status==='error') continue;
      proposals.push({from:candidate,to:{day,periodId:p.id},target:{day:candidate.day,periodId:candidate.periodId}});
      break;
    }
    if(proposals.length>=2) break;
  }
  return proposals;
};

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

// STAGE 7: locked automatic rows remain anchors for future regeneration.
const isAutoGeneratedSchedule = (scheduleItem) => {
  if (scheduleItem?.keepOnRegenerate === true || scheduleItem?.source === 'manual') return false;
  if (!scheduleItem) return false;
  if (scheduleItem.source === 'auto') return true;
  return /^sch_auto_/.test(String(scheduleItem.id || ''));
};

// STAGE 28: Plan a bulk reset of draft LESSONS only. Fixed activities, PLC and
// published versions live in separate collections and are never delete targets.
// Manual entries are preserved by default (as in Auto regeneration), but users
// may explicitly opt in to clearing unpinned manual lessons.
const planUnlockedScheduleClear = (scheduleRows = [], targetRoomId = '') => {
  if (!Array.isArray(scheduleRows)) throw new Error('INVALID_SCHEDULE_ROWS');
  const selected = [], kept = [];
  let autoCount = 0, manualCount = 0, pinnedCount = 0, outOfScopeCount = 0;
  const seenIds = new Set();
  scheduleRows.forEach(row => {
    const id = String(row?.id || '').trim();
    if (!id || id.includes('/') || seenIds.has(id)) throw new Error('INVALID_SCHEDULE_ROWS');
    seenIds.add(id);
    if (targetRoomId && String(row.classroomId) !== String(targetRoomId)) {
      outOfScopeCount++; kept.push(row); return;
    }
    if (row.keepOnRegenerate === true) { pinnedCount++; kept.push(row); return; }
    // Includes manual, auto and legacy rows: only a deliberate lock protects a lesson.
    selected.push(row);
    if (isAutoGeneratedSchedule(row)) autoCount++; else manualCount++;
  });
  return { selected, kept, autoCount, manualCount, pinnedCount, outOfScopeCount,
    deleteCount:selected.length, targetRoomId };
};

// STAGE 8: pure drag planner; never mutates input schedules or writes data.
const buildScheduleDragPlan = ({schedules, sourceId, targetDay, targetPeriodId, periods, subject}) => {
  const source=schedules.find(x=>x.id===sourceId);
  if(!source) throw new Error('MISSING_SOURCE');
  if(source.keepOnRegenerate) throw new Error('LOCKED_SOURCE');
  if(!DAYS.includes(targetDay)) throw new Error('INVALID_DAY');
  const targetIndex=periods.findIndex(p=>p.id===targetPeriodId&&!p.isBreak);
  if(targetIndex<0) throw new Error('INVALID_PERIOD');
  const currentIndex=periods.findIndex(p=>p.id===source.periodId);
  if(currentIndex<0) throw new Error('INVALID_SOURCE_PERIOD');
  const sameLoad=schedules.filter(x=>x.teacherId===source.teacherId&&x.subjectId===source.subjectId&&x.classroomId===source.classroomId);
  const partner=Number(subject?.periodBlock)===2 ? findBlockPartner(source,sameLoad,periods) : null;
  const moved=partner?[source,partner]:[source];
  if(moved.some(item=>item.keepOnRegenerate)) throw new Error('LOCKED_PARTNER');
  const planned=moved.map(item=>{
    const index=targetIndex+(periods.findIndex(p=>p.id===item.periodId)-currentIndex);
    const slot=periods[index];
    if(!slot || slot.isBreak) throw new Error('INVALID_BLOCK_PAIR');
    return {...item,day:targetDay,periodId:slot.id,source:'manual'};
  });
  const remaining=schedules.filter(x=>!moved.some(item=>item.id===x.id));
  return {source,moved,planned,remaining};
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

// STAGE 5A: เทียบประวัติตารางแบบอ่านอย่างเดียว
// "เพิ่มในปัจจุบัน" หมายถึงมีในตารางปัจจุบัน แต่ไม่มีในประวัติ
// STAGE 5B: deterministic integrity fingerprint for snapshot rows. This is a corruption
// detector (not a cryptographic signature or security authorization).
const fingerprintScheduleRows = rows => {
  const serialized = stableSerializeData((rows || []).map(cleanScheduleSnapshotItem)
    .sort((a,b) => String(a.id).localeCompare(String(b.id))));
  let hash = 2166136261;
  for (let i=0;i<serialized.length;i++) hash = Math.imul(hash ^ serialized.charCodeAt(i), 16777619);
  return `${serialized.length}:${(hash >>> 0).toString(16)}`;
};
// STAGE 19: a stable print reference without changing the draft or publication data.
// Published copies point to their stored Firestore version; drafts get a content-based
// verification code that changes whenever schedule/fixed/activity rows change.
const buildTimetablePrintCode = ({ academicYear, semester, versionId, schedules = [], fixedSchedules = [], activityGroups = [] }) => {
  const prefix = `BHS-${String(academicYear || '----')}-${String(semester || '-')}`;
  if (versionId) return `${prefix}-P-${String(versionId)}`;
  const summary = [
    { id: 'regular', fingerprint: fingerprintScheduleRows(schedules) },
    { id: 'fixed', fingerprint: fingerprintScheduleRows(fixedSchedules) },
    { id: 'groups', fingerprint: fingerprintScheduleRows(activityGroups) },
  ];
  const shortHash = fingerprintScheduleRows(summary).split(':')[1].toUpperCase().padStart(8, '0');
  return `${prefix}-D-${shortHash}`;
};

const makeScheduleSnapshotChunks = (schedules, fixedSchedules, activityGroups = [], plcSessions = []) => {
  const empty = () => ({ schedules: [], fixedSchedules: [], activityGroups: [], plcSessions: [] });
  const chunks = []; let current = empty(); let byteSize=0;
  const flush = () => { if (current.schedules.length || current.fixedSchedules.length || current.activityGroups.length || current.plcSessions.length) {
    chunks.push(current); current = empty(); byteSize=0;
  }};
  for (const [kind, items] of [['schedules', schedules], ['fixedSchedules', fixedSchedules], ['activityGroups',activityGroups],['plcSessions',plcSessions]]) {
    for (const item of items) {
      const bytes = new TextEncoder().encode(JSON.stringify(item)).length + 150;
      if (bytes > 180000) throw new Error('SNAPSHOT_ROW_TOO_LARGE');
      if (byteSize + bytes > 180000 || current.schedules.length + current.fixedSchedules.length + current.activityGroups.length + current.plcSessions.length >= 150) flush();
      current[kind].push(item);byteSize+=bytes;
    }
  }
  flush(); return chunks;
};

const compareScheduleSnapshotRows = (currentItems, oldItems) => {
  const current = new Map(currentItems.map(item => [String(item.id), cleanScheduleSnapshotItem(item)]));
  const old = new Map(oldItems.map(item => [String(item.id), cleanScheduleSnapshotItem(item)]));
  const result = { added: [], deleted: [], changed: [], identical: 0 };
  current.forEach((item, id) => {
    if (!old.has(id)) result.added.push(item);
    else if (!areFirestoreRecordsEqual(item, old.get(id))) result.changed.push({ before: old.get(id), after: item });
    else result.identical++;
  });
  old.forEach((item, id) => { if (!current.has(id)) result.deleted.push(item); });
  return result;
};

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

const calculateScheduleHealth = (schedules, loads, unavail, cUnavail, fixed, teachers, subjects, classrooms, periods, customRules = [], activityGroups = [], plcSessions = []) => {
  const health = {
    status: 'ready', totalRequired: 0, scheduledForLoads: 0, missing: 0, overScheduled: 0, completionPercent: 0,
    teacherConflicts: [], roomConflicts: [], unavailConflicts: [], cUnavailConflicts: [], breakConflicts: [], blockConflicts: [], teacherLimitConflicts: [], teacherFivePeriodWarnings: [], teacherActivityConsecutiveWarnings: [], teacherManualConsecutiveWarnings: [], customRuleConflicts: [], subjectDailyConflicts: [], orphans: [], fixedConflicts: [], missingLoads: [], overloads: [],
    teacherStats: {}, roomStats: {}, plcConflicts: [], 
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
    const subjectExists = subjectMap.has(sch.subjectId) || (isFixed && isDevelopmentActivity(sch) && Boolean(String(sch.activityName || '').trim()));
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
      ? fixedTargetRoomIds(sch, classrooms) : (sch.classroomId ? [sch.classroomId] : []);
    if (isFixed && sch.classroomId===ALL_CLASSROOMS_ID && !roomIds.length)
      health.orphans.push(`กิจกรรมล็อก ${sch.id || '-'} ไม่มีห้องเรียนที่เข้าร่วม`);
    roomIds.forEach(roomId => {
      const rKey = `${roomId}_${sch.day}_${sch.periodId}`;
      if (!rPlacement[rKey]) rPlacement[rKey] = [];
      rPlacement[rKey].push(sch);
    });

    const hasCompleteTeachingRef = Boolean(sch.teacherId && sch.subjectId && sch.classroomId !== ALL_CLASSROOMS_ID && sch.classroomId);
    const lKey = hasCompleteTeachingRef ? `${sch.teacherId}_${sch.subjectId}_${sch.classroomId}` : '';
    const matchedLoad = !isFixed || !isDevelopmentActivity(sch) ? (lKey ? loadTracker[lKey] : null) : null;

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
  // Activity group teachers count as busy, but the all-room activity already occupies
  // the classroom slot; group overlays must NOT duplicate room occupancy.
  const groupTeacherItems=activityTeacherEntries(activityGroups,fixed);
  groupTeacherItems.forEach(item=>{
    if (!teacherMap.has(item.teacherId)) health.orphans.push(`ครูกิจกรรมกลุ่ม ${item.teacherId} ไม่มีในระบบ`);
    if (unavail.some(u=>u.teacherId===item.teacherId&&u.day===item.day&&u.periodId===item.periodId))
      health.unavailConflicts.push(`ครูกิจกรรมกลุ่ม ${item.teacherId} ไม่ว่างวัน${item.day} คาบ ${item.periodId}`);
    const key=`${item.teacherId}_${item.day}_${item.periodId}`;
    (tPlacement[key] ||= []).push(item);
  });
  activityGroups.forEach(g=>{
    const root=fixed.find(f=>f.id===g.fixedId && f.classroomId===ALL_CLASSROOMS_ID);
    if(!root || !fixedAppliesToClassroom(root,g.classroomId) || !roomMap.has(g.classroomId) || !String(g.name||'').trim())
      health.orphans.push(`กิจกรรมกลุ่ม ${g.id||'-'} อ้างอิงกิจกรรมหลักหรือห้องเรียนไม่ถูกต้อง`);
  });

  // PLC reserves the teacher only; do not create classroom occupation or teaching-load counts.
  const plcAreaCount=new Map();
  plcSessions.forEach(plc => {
    plcAreaCount.set(plc.subjectArea,(plcAreaCount.get(plc.subjectArea)||0)+1);
    if ((plcAreaCount.get(plc.subjectArea)||0)>1)
      health.plcConflicts.push(`พบ PLC กลุ่ม ${plc.subjectArea||'-'} ซ้ำมากกว่า 1 รายการ`);
    if (!PLC_AREAS.includes(plc.subjectArea) || !DAYS.includes(plc.day) ||
        !periods.some(p=>p.id===plc.periodId&&!p.isBreak))
      health.plcConflicts.push(`PLC กลุ่ม ${plc.subjectArea||'-'} มีวัน/คาบไม่ถูกต้อง`);
    if (!plcTeachersForArea(plc.subjectArea,teachers).length)
      health.plcConflicts.push(`PLC กลุ่ม ${plc.subjectArea||'-'} ไม่มีครูในกลุ่มสาระ`);
    plcSlotIssues({subjectArea:plc.subjectArea,day:plc.day,periodId:plc.periodId,
      teachers,schedules,fixedSchedules:fixed,activityGroups,unavailabilities:unavail,
      plcSessions:plcSessions.filter(x=>x.id!==plc.id),periods}).forEach(issue=>
      health.plcConflicts.push(`PLC ${plc.subjectArea}: ${issue.teacherName||'-'} ${issue.reason}`));
  });
  Object.entries(tPlacement).filter(([_, arr]) => arr.length > 1).forEach(([_, arr]) => health.teacherConflicts.push(`ครู ${teacherMap.get(arr[0].teacherId)?.name || arr[0].teacherId} ซ้อนทับ ${arr.length} วิชา (วัน${arr[0].day} ${getPeriodDisplayName(periods, arr[0].periodId)})`));
  Object.entries(rPlacement).filter(([_, arr]) => arr.length > 1).forEach(([_, arr]) => health.roomConflicts.push(`ห้อง ${roomMap.get(arr[0].classroomId)?.name || arr[0].classroomId} ซ้อนทับ ${arr.length} วิชา (วัน${arr[0].day} ${getPeriodDisplayName(periods, arr[0].periodId)})`));

  unavail.forEach(u => { if (tPlacement[`${u.teacherId}_${u.day}_${u.periodId}`]) health.unavailConflicts.push(`ครู ${teacherMap.get(u.teacherId)?.name || u.teacherId} ถูกจัดในเวลาไม่สะดวก (วัน${u.day} ${getPeriodDisplayName(periods, u.periodId)})`); });
  cUnavail.forEach(u => {
    const occupied = rPlacement[`${u.classroomId}_${u.day}_${u.periodId}`] || [];
    if (occupied.some(item => !isApprovedManualRoomException(item,u)))
      health.cUnavailConflicts.push(`ห้อง ${roomMap.get(u.classroomId)?.name || u.classroomId} ถูกจัดในเวลาไม่ว่าง (วัน${u.day} ${getPeriodDisplayName(periods, u.periodId)})`);
  });

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

  const occupied = buildTeacherOccupancy(schedules, fixed, activityGroups);
  const weeklyStats = summarizeTeacherWeek(occupied, periods);
  const byTeacherTeaching = new Map(), byTeacherActivities = new Map();
  schedules.forEach(item => { if (item.teacherId) byTeacherTeaching.set(item.teacherId, (byTeacherTeaching.get(item.teacherId) || 0) + 1); });
  [...fixed, ...groupTeacherItems].forEach(item => { if (item.teacherId) byTeacherActivities.set(item.teacherId, (byTeacherActivities.get(item.teacherId) || 0) + 1); });
  teachers.forEach(t => {
    const stat = health.teacherStats[t.id] || (health.teacherStats[t.id] = {req:0,sch:0,days:{},gaps:0,maxConsec:0,subjects:new Set()});
    const ws = weeklyStats.get(t.id);
    stat.regularPeriods = byTeacherTeaching.get(t.id) || 0;
    stat.activityPeriods = byTeacherActivities.get(t.id) || 0;
    stat.totalWeeklyPeriods = ws?.total || 0;
    stat.heavyDays = ws?.heavyDays || 0;
    stat.fiveDays = ws?.fiveDays || [];
    stat.dailyCounts = ws?.dayCounts || {};
    // Report all teacher occupancy; don't hide groups/fixed activities from gaps.
    stat.gaps = DAYS.reduce((sum,d) => sum + calculateDayTeachingShape(
      occupied.filter(item=>item.teacherId===t.id && item.day===d).map(item=>item.periodId),periods).gaps,0);
    stat.maxConsec = ws?.longestConsecutive || 0;
  });
  health.subjectDailyConflicts = findSubjectDailyLimitConflicts([...schedules, ...fixed.filter(f => !isDevelopmentActivity(f))], classrooms, customRules);
  health.customRuleConflicts = [...schedules, ...fixed].filter(item => blocksCustomRule(customRules, item, periods)).map(item =>
    `ข้อบังคับที่ตั้งเอง: ${item.subjectId || '-'} / วัน${item.day} ${getPeriodDisplayName(periods, item.periodId)} ไม่สามารถจัดได้`);
  health.teacherLimitConflicts = findTeacherTeachingLimitConflicts(occupied, periods, teachers);
  health.teacherManualConsecutiveWarnings = findManualThreePeriodWarnings(occupied, periods, teachers);
  health.teacherFivePeriodWarnings = findTeacherFifthPeriodDays(occupied, periods, teachers)
    .map(x => `ครู ${x.teacherName} วัน${x.day} มีคาบสอนและกิจกรรมรวม 5 คาบ (ข้อยกเว้น ควรพิจารณากระจายใหม่)`);
  health.teacherActivityConsecutiveWarnings = findTeacherActivityConsecutiveWarnings(occupied, periods, teachers);
  const hasHardConflict = health.teacherConflicts.length > 0 || health.roomConflicts.length > 0 || health.unavailConflicts.length > 0 || health.cUnavailConflicts.length > 0 || health.breakConflicts.length > 0 || health.blockConflicts.length > 0 || health.teacherLimitConflicts.length > 0 || health.customRuleConflicts.length > 0 || health.subjectDailyConflicts.length > 0 || health.orphans.length > 0 || health.plcConflicts.length > 0;
  if (hasHardConflict) health.status = 'error';
  else if (loads.length === 0) health.status = 'empty';
  else if (health.missing > 0 || health.overScheduled > 0 || health.teacherFivePeriodWarnings.length || health.teacherActivityConsecutiveWarnings.length || health.teacherManualConsecutiveWarnings.length) health.status = 'warning';
  else health.status = 'ready';

  return health;
};

// ข้อบังคับครู: นับคาบที่ครูมีสอน/กิจกรรมล็อกจริงรวมกัน และไม่นับคาบพักเป็นคาบติดกัน
const teacherDayKey = (teacherId, day) => JSON.stringify([teacherId, day]);
const registerTeacherPeriod = (maps, item, isActivity = isTeacherActivityDuty(item)) => {
  if (!item.teacherId) return;
  const key = teacherDayKey(item.teacherId, item.day);
  if (!maps.teacherDayPeriods.has(key)) maps.teacherDayPeriods.set(key, new Set());
  maps.teacherDayPeriods.get(key).add(item.periodId);
  // เฉพาะคาบสอนรายวิชาที่ใช้ตรวจข้อห้ามสอนติดกัน / รูปแบบคาบสอน
  if (!isActivity) {
    if (!maps.teacherInstructionDayPeriods.has(key)) maps.teacherInstructionDayPeriods.set(key, new Set());
    maps.teacherInstructionDayPeriods.get(key).add(item.periodId);
  }
};

// Manual-only exception: permit up to THREE consecutive lessons after explicit confirmation.
// Auto calls this function without the option, so its two-lesson limit is unchanged.
const canPlaceTeacherEntries = (teacherId, entries, maps, isActivityPlacement = false, allowManualThree = false) => {
  if (!teacherId) return { ok: true, reason: '' };
  const byDay = new Map();
  entries.forEach(item => {
    if (!byDay.has(item.day)) byDay.set(item.day, []);
    byDay.get(item.day).push(item.periodId);
  });
  for (const [day, additions] of byDay) {
    const existing = maps.teacherDayPeriods.get(teacherDayKey(teacherId, day)) || new Set();
    const result = new Set([...existing, ...additions]);
    if (result.size > (maps.maxTeacherPeriodsPerDay ?? SCHEDULER_CONFIG.maxTeacherPeriodsPerDay)) {
      return { ok: false, reason: 'TEACHER_DAILY_LIMIT' };
    }
    // กิจกรรมอาจอยู่ติดกับวิชาสอน เช่น สอนคาบ 5–6 แล้วทำลูกเสือคาบ 7
    // ไม่ห้ามเพียงเพราะเป็นหน้าที่ต่อเนื่อง 3 คาบ แต่เตือนใน Health Check แทน
    const instruction = new Set([...(maps.teacherInstructionDayPeriods.get(teacherDayKey(teacherId, day)) || []), ...(isActivityPlacement ? [] : additions)]);
    // Exact 3-lesson runs confirmed in the MANUAL timetable are safe anchors.
    // An Auto candidate may coexist on a separate run, but cannot extend such
    // a streak to 4 or create a new unapproved 3-lesson run.
    const instructionRuns = getTeacherTeachingRuns([...instruction], maps.periodContext);
    const unapprovedRun = instructionRuns.some(streak =>
      streak.length > (allowManualThree ? 3 : SCHEDULER_CONFIG.maxConsecutiveTeacherPeriods) &&
      !(streak.length === 3 && maps.approvedManualTeachingStreaks?.has(
        JSON.stringify([String(teacherId),day,streak]))));
    if (unapprovedRun) return { ok: false, reason: 'TEACHER_CONSECUTIVE_LIMIT' };
    if (getFourPeriodTeachingPattern([...instruction], maps.periodContext)?.level === 'blocked' ||
        hasSchoolPeriodSequence([...instruction], maps.periodContext, [3,4,5,6])) {
      return { ok: false, reason: 'TEACHER_FOUR_PERIOD_PATTERN' };
    }
  }
  return { ok: true, reason: '' };
};

const buildConstraintMaps = (schedules, tUnavail, cUnavail, fixedSch, periods = DEFAULT_PERIODS, customRules = [], activityGroups = [], classrooms = [], plcSessions = [], teachers = []) => {
  const maps = { tUnavailMap: new Set(), cUnavailMap: new Set(), fixedMap: new Set(), teacherPlacements: new Set(), roomPlacements: new Set(), teacherDayPeriods: new Map(), teacherInstructionDayPeriods: new Map(), subjectDayCounts: new Map(), periodContext: buildPeriodContext(periods), customRules, fixedItems: fixedSch || [], maxTeacherPeriodsPerDay: SCHEDULER_CONFIG.maxTeacherPeriodsPerDay };
  tUnavail.forEach(u => maps.tUnavailMap.add(`${u.teacherId}_${u.day}_${u.periodId}`));
  buildPLCReservations(plcSessions,teachers).forEach(x=>
    maps.tUnavailMap.add(`${x.teacherId}_${x.day}_${x.periodId}`));
  cUnavail.forEach(u => maps.cUnavailMap.add(`${u.classroomId}_${u.day}_${u.periodId}`));
  fixedSch.forEach(f => {
    if (f.classroomId !== ALL_CLASSROOMS_ID) maps.fixedMap.add(`${f.day}_${f.periodId}_${f.classroomId}`);
    else if (!Array.isArray(f.includedClassroomIds) && !(f.excludedClassroomIds || []).length)
      maps.fixedMap.add(`${f.day}_${f.periodId}_${ALL_CLASSROOMS_ID}`);
    if (f.teacherId) maps.teacherPlacements.add(`${f.teacherId}_${f.day}_${f.periodId}`);
    registerTeacherPeriod(maps, f, true);
    if (f.classroomId === ALL_CLASSROOMS_ID &&
        (Array.isArray(f.includedClassroomIds) || (f.excludedClassroomIds || []).length)) {
      fixedTargetRoomIds(f,classrooms).forEach(id => {
        if (!isDevelopmentActivity(f)) registerSubjectDayPeriod(maps, {...f,classroomId:id});
        maps.fixedMap.add(`${f.day}_${f.periodId}_${id}`);
      });
    } else if (!isDevelopmentActivity(f)) registerSubjectDayPeriod(maps, f);
    maps.roomPlacements.add(`${f.classroomId}_${f.day}_${f.periodId}`);
  });
  activityTeacherEntries(activityGroups,fixedSch).forEach(item=>{
    maps.teacherPlacements.add(`${item.teacherId}_${item.day}_${item.periodId}`);
    registerTeacherPeriod(maps,item,true);
  });
  schedules.forEach(s => {
    if (s.teacherId) maps.teacherPlacements.add(`${s.teacherId}_${s.day}_${s.periodId}`);
    registerTeacherPeriod(maps, s);
    registerSubjectDayPeriod(maps, s);
    maps.roomPlacements.add(`${s.classroomId}_${s.day}_${s.periodId}`);
  });
  // Store only already-confirmed exact manual three-period streaks. This is a
  // read-only baseline and is copied to each Auto candidate constraint map.
  maps.approvedManualTeachingStreaks = new Set();
  groupTeacherDuties([...fixedSch, ...schedules]).forEach(({teacherId,day,lessons}) => {
    getTeacherTeachingRuns([...lessons], maps.periodContext).forEach(streak => {
      if (streak.length === 3 && getApprovedManualTeachingStreak(schedules,teacherId,day,streak))
        maps.approvedManualTeachingStreaks.add(JSON.stringify([String(teacherId),day,streak]));
    });
  });
  return maps;
};

const canPlaceSchedule = (teacherId, classroomId, day, periodId, maps, isBreak, subjectId = '', allowManualClosedRoom = false, isActivityPlacement = false, allowManualThree = false) => {
  if (isBreak) return { ok: false, reason: "BREAK_PERIOD" };
  if (blocksCustomRule(maps.customRules, { teacherId, classroomId, subjectId, day, periodId }, maps.periodContext)) return { ok: false, reason: 'CUSTOM_RULE_BLOCKED' };
  if (!allowManualClosedRoom && maps.cUnavailMap.has(`${classroomId}_${day}_${periodId}`)) return { ok: false, reason: "ROOM_UNAVAILABLE" };
  if (teacherId && maps.tUnavailMap.has(`${teacherId}_${day}_${periodId}`)) return { ok: false, reason: "TEACHER_UNAVAILABLE" };
  if (maps.fixedMap.has(`${day}_${periodId}_${classroomId}`) || maps.fixedMap.has(`${day}_${periodId}_${ALL_CLASSROOMS_ID}`) ||
      (maps.fixedItems || []).some(f => f.day===day && f.periodId===periodId && fixedAppliesToClassroom(f,classroomId)))
    return { ok: false, reason: "FIXED_SCHEDULE_CONFLICT" };
  if (maps.roomPlacements.has(`${classroomId}_${day}_${periodId}`)) return { ok: false, reason: "ROOM_CONFLICT" };
  if (teacherId && maps.teacherPlacements.has(`${teacherId}_${day}_${periodId}`)) return { ok: false, reason: "TEACHER_CONFLICT" };
  const subjectLimit = canPlaceSubjectDailyEntries({ classroomId, subjectId }, [{ day, periodId }], maps);
  if (!subjectLimit.ok) return subjectLimit;
  const teachingLimit = canPlaceTeacherEntries(teacherId, [{ day, periodId }], maps, isActivityPlacement, allowManualThree);
  if (!teachingLimit.ok) return teachingLimit;
  return { ok: true, reason: "" };
};

// STEP 04: ตรวจความพร้อมของข้อมูลก่อนสร้างตาราง (ไม่มีการเขียน Firebase)
// ตรวจเฉพาะสิ่งที่ทำให้ตารางไม่ถูกต้อง โดยไม่ปฏิเสธเพียงเพราะช่องเวลามีไม่พอ
const validateAutoScheduleInputs = ({ teachingLoads, teachers, subjects, classrooms, periods,
  schedules, fixedSchedules, unavailabilities, classroomUnavailabilities, semester, customRules = [], activityGroups = [], plcSessions = [] }) => {
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

  plcSessions.forEach(plc=>{
    const issues=plcSlotIssues({subjectArea:plc.subjectArea,day:plc.day,periodId:plc.periodId,
      teachers,schedules:schedules.filter(x=>!isAutoGeneratedSchedule(x)),fixedSchedules,
      activityGroups,unavailabilities,plcSessions:plcSessions.filter(x=>x.id!==plc.id),periods});
    issues.forEach(issue=>addProblem(`PLC ${plc.subjectArea}: ${issue.teacherName||'-'} ${issue.reason}`));
  });
  anchors.forEach((item, index) => {
    const label = `${item.__kind} รายการที่ ${index + 1}`;
    if (item.teacherId && !teacherIds.has(String(item.teacherId))) addProblem(`${label}: ไม่พบครู ${item.teacherId}`);
    if (!(item.__kind === 'กิจกรรมล็อก' && isDevelopmentActivity(item) && String(item.activityName || '').trim()) && !validSubject(item.subjectId)) addProblem(`${label}: ไม่พบรายวิชา ${item.subjectId || '-'} ในภาคเรียนนี้`);
    const appliesToAll = item.__kind === 'กิจกรรมล็อก' && item.classroomId === ALL_CLASSROOMS_ID;
    if (appliesToAll ? fixedTargetRoomIds(item,classrooms).length === 0 : !classroomIds.has(String(item.classroomId))) {
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
    const affectedRoomIds = appliesToAll ? fixedTargetRoomIds(item,classrooms) : [item.classroomId];
    affectedRoomIds.forEach(roomId => {
      const slotRoom = JSON.stringify([roomId, item.day, String(item.periodId)]);
      if (usedRooms.has(slotRoom)) addProblem(`${label}: ห้อง ${roomId} มีตารางซ้อนในวัน${item.day} คาบ ${item.periodId}`);
      usedRooms.add(slotRoom);
      if (unavailableRooms.has(slotRoom) && !(item.__kind === 'ตารางที่จัดด้วยมือ' && item.source === 'manual' && item.manualRoomClosureOverride === true)) addProblem(`${label}: ห้อง ${roomId} ถูกจัดในเวลาที่ไม่ว่าง`);
    });
    if (item.__kind === 'ตารางที่จัดด้วยมือ' && !item.teacherId) addProblem(`${label}: ไม่ระบุครูผู้สอน`);
    if (item.teacherId && !appliesToAll && !isDevelopmentActivity(item)) {
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
  customRules.forEach((rule, index) => {
    if (!isValidCustomRule(rule, teachers, subjects, classrooms, periods) ||
        (rule.targetType === 'subject' && !validSubject(rule.targetId))) {
      addProblem(`เงื่อนไขเพิ่มเติมรายการที่ ${index + 1}: อ้างอิงข้อมูลไม่ถูกต้องหรือใช้กับวิชาคนละภาคเรียน`);
    }
  });
  findSubjectDailyLimitConflicts(anchors.filter(item => !isDevelopmentActivity(item)), classrooms, customRules).forEach(addProblem);
  // Pre-existing manual/pinned lessons are protected exceptions, NOT an Auto
  // input error. Still reject invalid fixed-only subject distributions.
  findAutoSubjectDistributionConflicts(
    anchors.filter(item => item.__kind === 'กิจกรรมล็อก' && !isDevelopmentActivity(item)),
    subjects, classrooms
  ).forEach(addProblem);
  anchors.filter(item => blocksCustomRule(customRules, item, periods)).forEach(item =>
    addProblem(`ตารางเดิม/กิจกรรมล็อกฝ่าฝืนข้อบังคับเพิ่มเติม: วัน${item.day} คาบ ${item.periodId}`));
  const groupTeachers=activityTeacherEntries(activityGroups,fixedSchedules);
  activityGroups.forEach(g=>{
    if(!fixedSchedules.some(f=>f.id===g.fixedId&&f.classroomId===ALL_CLASSROOMS_ID && fixedAppliesToClassroom(f,g.classroomId))) addProblem(`กิจกรรมกลุ่ม ${g.id}: ไม่พบกิจกรรมล็อกของห้องที่อ้างอิง`);
    if(!classroomIds.has(g.classroomId)) addProblem(`กิจกรรมกลุ่ม ${g.id}: ไม่พบห้องเรียน`);
  });
  const busy=new Set(anchors.filter(x=>x.teacherId).map(x=>JSON.stringify([x.teacherId,x.day,String(x.periodId)])));
  groupTeachers.forEach(item=>{
    const key=JSON.stringify([item.teacherId,item.day,String(item.periodId)]);
    if(!teacherIds.has(item.teacherId)) addProblem(`ครูกิจกรรม ${item.teacherId} ไม่มีในระบบ`);
    if(busy.has(key)) addProblem(`ครูกิจกรรม ${item.teacherId} ซ้อนกับคาบอื่น วัน${item.day} คาบ ${item.periodId}`);
    busy.add(key);
    if(unavailableTeachers.has(key)) addProblem(`ครูกิจกรรม ${item.teacherId} ไม่ว่างในเวลาที่กำหนด`);
  });
  findTeacherTeachingLimitConflicts([...anchors,...groupTeachers], periods, teachers).forEach(addProblem);
  return problems;
};

// STAGE 3: วิเคราะห์ขอบเขตความเป็นไปได้แบบไม่เขียนข้อมูลและไม่เรียกตัวจัดตาราง
// เป็นเพียง necessary conditions: ผ่านการตรวจนี้ยังไม่ได้รับประกันว่าจัดครบได้
const analyzeScheduleFeasibility = ({ teachingLoads, teachers, subjects, classrooms, periods,
  schedules, fixedSchedules, unavailabilities, classroomUnavailabilities, customRules = [], activityGroups = [], plcSessions = [], semester }) => {
  const invalidInputs = validateAutoScheduleInputs({ teachingLoads, teachers, subjects, classrooms,
    periods, schedules, fixedSchedules, unavailabilities, classroomUnavailabilities, customRules,activityGroups,plcSessions,semester });
  const result = { invalidInputs, blocking: [], advisories: [], totalRequired: teachingLoads.reduce((n,l)=>n+(Number(l.periods)||0),0),
    activePeriods: periods.filter(p=>!p.isBreak).length, teacherCount: teachers.length, roomCount: classrooms.length };
  // ข้อมูลอ้างอิงไม่ถูกต้อง: ไม่คำนวณขีดจำกัดต่อ เพื่อหลีกเลี่ยงรายงานผิด
  if (invalidInputs.length) return result;
  const active = periods.filter(p => !p.isBreak);
  const allSlots = DAYS.length * active.length;
  const anchored = schedules.filter(s => !isAutoGeneratedSchedule(s));
  const anchors = [...anchored, ...fixedSchedules];
  const anchoredSubjectDays = countSubjectDayDistribution(anchors, subjects, classrooms);
  const protectedDailyExceptions = findAutoSubjectDistributionConflicts(anchors,subjects,classrooms);
  protectedDailyExceptions.forEach(message => result.advisories.push(
    message.replace(/ Auto ห้ามเพิ่มคาบจนเกินเกณฑ์เดิมหรือเกินคาบที่คงไว้$/, ' (คงคาบที่ลงเอง/ล็อกไว้ และ Auto จะไม่เพิ่มคาบในวันนั้น)')
  ));
  const baseMaps = buildConstraintMaps(anchored, unavailabilities, classroomUnavailabilities,
    fixedSchedules, periods, customRules,activityGroups,classrooms,plcSessions,teachers);
  baseMaps.enforceAutoSubjectDailyCap = true;
  baseMaps.autoSubjectBlockSizes = new Map(subjects.map(s=>[String(s.id), Number(s.periodBlock)===2 ? 2 : 1]));
  const teacherDemand = new Map(), roomDemand = new Map();
  teachingLoads.forEach(load => {
    const count = Number(load.periods);
    teacherDemand.set(load.teacherId, (teacherDemand.get(load.teacherId)||0)+count);
    roomDemand.set(load.classroomId, (roomDemand.get(load.classroomId)||0)+count);
  });
  const seen = new Set();
  const block = message => { if (!seen.has(message)) { seen.add(message); result.blocking.push(message); } };

  const fixedDutyRows = buildTeacherOccupancy([],fixedSchedules,activityGroups);
  const existingActivityDuty = groupTeacherDuties(fixedDutyRows);
  teacherDemand.forEach((required, teacherId) => {
    // 5 คาบรวมกิจกรรมต่อวันเป็นเพดานจริง; คำนวณความจุเหลือโดยหักกิจกรรมไว้ก่อน
    let upper=0,preferredCapacity=0;
    DAYS.forEach(day=>{
      const activities=existingActivityDuty.get(teacherDayKey(teacherId,day))?.all.size || 0;
      const slots = active.filter(p =>
        !unavailabilities.some(u => u.teacherId === teacherId && u.day === day && String(u.periodId) === String(p.id)) &&
        !fixedDutyRows.some(a => a.teacherId===teacherId && a.day===day && String(a.periodId)===String(p.id)) &&
        !blocksCustomRule(customRules, {teacherId, day, periodId:p.id, classroomId:'', subjectId:''}, periods)
      ).length;
      upper += Math.min(Math.max(0,SCHEDULER_CONFIG.absoluteMaxTeacherPeriodsPerDay-activities),slots);
      preferredCapacity += Math.min(Math.max(0,SCHEDULER_CONFIG.maxTeacherPeriodsPerDay-activities),slots);
    });
    if (required > upper) block(`ครู ${teachers.find(t=>t.id===teacherId)?.name || teacherId}: ภาระสอน ${required} คาบ แต่เมื่อนับกิจกรรมล็อกแล้ว เหลือความจุสูงสุดประมาณ ${upper} คาบสอน/สัปดาห์`);
    else if (required > preferredCapacity) result.advisories.push(`ครู ${teachers.find(t=>t.id===teacherId)?.name || teacherId}: มีภาระสอน ${required} คาบ และมีกิจกรรมล็อก ทำให้เกินเป้าหมายรวม 4 คาบบางวันได้ (ความจุที่เป้าหมายประมาณ ${preferredCapacity} คาบ/สัปดาห์)`);
    else if (upper > 0 && required / upper >= .85) result.advisories.push(`ครู ${teachers.find(t=>t.id===teacherId)?.name || teacherId}: ภาระสอนใช้ ${Math.round(required/upper*100)}% ของเพดานสูงสุดที่คำนวณแบบคร่าว ๆ`);
  });
  roomDemand.forEach((required, roomId) => {
    if (required > allSlots) block(`ห้อง ${classrooms.find(r=>r.id===roomId)?.name || roomId}: ต้องจัด ${required} คาบ แต่มีเพียง ${allSlots} คาบเรียนต่อสัปดาห์`);
  });
  // ขอบเขตสูงสุดตามจำนวนวัน: คาบเดี่ยวไม่เกิน 5 วัน วันละ 1 คาบ
  const roomSubjectDemand = new Map();
  teachingLoads.forEach(load => {
    const key = JSON.stringify([String(load.classroomId), String(load.subjectId)]);
    roomSubjectDemand.set(key, (roomSubjectDemand.get(key) || 0) + (Number(load.periods) || 0));
  });
  roomSubjectDemand.forEach((required, key) => {
    const [roomId, subjectId] = JSON.parse(key);
    const cap = baseMaps.autoSubjectBlockSizes.get(subjectId) === 2 ? 2 : 1;
    // An existing manual/pinned overflow on a particular day is already part
    // of the required weekly load; Auto can still use the other weekdays.
    const protectedOverflow = DAYS.reduce((sum,day) =>
      sum + Math.max(0,(anchoredSubjectDays.get(subjectDailyKey(roomId,subjectId,day))?.count || 0)-cap),0);
    const upperWeekly = DAYS.length * cap + protectedOverflow;
    if (required > upperWeekly) block(`วิชา ${subjectId} ห้อง ${classrooms.find(r=>String(r.id)===roomId)?.name || roomId}: ต้องการ ${required} คาบ/สัปดาห์ แต่จัดได้สูงสุด ${upperWeekly} คาบ/สัปดาห์ (นับข้อยกเว้นคาบเดิมแล้ว) กรุณาตรวจจำนวนคาบหรือ Block รายวิชา`);
  });

  teachingLoads.forEach(load => {
    const required = Number(load.periods);
    const matching = anchors.filter(a => a.teacherId === load.teacherId && a.subjectId === load.subjectId && a.classroomId === load.classroomId).length;
    let available = 0;
    DAYS.forEach(day => active.forEach(p => {
      if (canPlaceSchedule(load.teacherId, load.classroomId, day, p.id, baseMaps, false, load.subjectId).ok) available++;
    }));
    // Each free slot is considered independently: this is an optimistic upper bound.
    // A real timetable may place fewer slots due to conflicts, blocks and consecutive-period rules.
    if (required > matching + available) {
      block(`วิชา ${subjects.find(s=>s.id===load.subjectId)?.name || load.subjectId} ห้อง ${classrooms.find(r=>r.id===load.classroomId)?.name || load.classroomId} (${teachers.find(t=>t.id===load.teacherId)?.name || load.teacherId}): ต้องการ ${required} คาบ แต่มีคาบเดิม ${matching} + ช่องว่างที่ผ่านเงื่อนไขเบื้องต้นไม่เกิน ${available} คาบ`);
    }
    if (Number(subjects.find(s=>s.id===load.subjectId)?.periodBlock) === 2 && required >= 2) {
      const pairs = DAYS.reduce((acc, day) => acc + buildPeriodContext(periods).doublePairs.filter(([a,b]) =>
        canPlaceSchedule(load.teacherId, load.classroomId, day, a.id, baseMaps, false, load.subjectId).ok &&
        canPlaceSchedule(load.teacherId, load.classroomId, day, b.id, baseMaps, false, load.subjectId).ok &&
        canPlaceTeacherEntries(load.teacherId, [{day,periodId:a.id},{day,periodId:b.id}], baseMaps).ok &&
        canPlaceSubjectDailyEntries(load, [{day,periodId:a.id},{day,periodId:b.id}], baseMaps).ok
      ).length, 0);
      if (pairs === 0 && required > matching) result.advisories.push(`วิชาคาบคู่ ${load.subjectId} ห้อง ${load.classroomId}: ไม่พบคาบคู่ที่ว่างอยู่ในขณะนี้ อาจต้องย้ายคาบที่จัดไว้`);
    }
  });
  return result;
};

// STAGE 4: อธิบายสาเหตุเชิงสังเกตของช่องที่จัดไม่ได้หลังจัดตาราง
// ไม่เปลี่ยนอัลกอริทึม ไม่ย้ายคาบ และไม่กล่าวอ้างว่าเป็นสาเหตุเดียวที่แท้จริง
const describeUnassignedReasons = ({ load, schedules, fixedSchedules, periods,
  unavailabilities, classroomUnavailabilities, customRules = [], activityGroups = [], classrooms = [], plcSessions = [], teachers = [], subjects = [] }) => {
  const labels = {
    BREAK_PERIOD: 'คาบพักกลางวัน', ROOM_UNAVAILABLE: 'ห้องเรียนไม่ว่างตามข้อจำกัด',
    TEACHER_UNAVAILABLE: 'ครูไม่สะดวกหรือมี PLC ในคาบนี้', FIXED_SCHEDULE_CONFLICT: 'มีกิจกรรมล็อกคาบนี้',
    ROOM_CONFLICT: 'ห้องมีวิชาอื่นแล้ว', TEACHER_CONFLICT: 'ครูมีสอนอีกห้องในคาบนี้',
    TEACHER_DAILY_LIMIT: 'ครูเต็มเพดานต่อวันที่กำลังใช้ (รอบแรก 4 / รอบสำรอง 5 คาบ)', TEACHER_CONSECUTIVE_LIMIT: 'ครูติดกันเกิน 2 คาบ',
    TEACHER_FOUR_PERIOD_PATTERN: 'รูปแบบสอน 4 คาบไม่อยู่ใน 23 รูปแบบที่อนุญาต',
    SUBJECT_DAILY_LIMIT: 'รหัสวิชานี้ครบจำนวนคาบต่อวันที่อนุญาตแล้ว (คาบเดี่ยว 1 / คาบคู่ 2)', CUSTOM_RULE_BLOCKED: 'ขัดกับข้อบังคับเพิ่มเติม',
  };
  const hints = {
    ROOM_UNAVAILABLE: 'ตรวจช่วงเวลาที่ห้องไม่ว่าง', TEACHER_UNAVAILABLE: 'ตรวจช่วงเวลาที่ครูไม่สะดวก',
    FIXED_SCHEDULE_CONFLICT: 'ตรวจเวลาของกิจกรรมล็อก ก่อนปรับแก้รายการที่เกี่ยวข้อง',
    ROOM_CONFLICT: 'พิจารณาสลับคาบในห้องเรียนนี้ โดยคงข้อบังคับทั้งหมด',
    TEACHER_CONFLICT: 'พิจารณาสลับคาบครูในห้องอื่น โดยคงข้อบังคับทั้งหมด',
    TEACHER_DAILY_LIMIT: 'กระจายภาระสอนของครูไปวันอื่น',
    TEACHER_CONSECUTIVE_LIMIT: 'จัดให้มีคาบว่างคั่นก่อนสอนต่อ',
    TEACHER_FOUR_PERIOD_PATTERN: 'ลองกระจายคาบสอนใหม่ เช่น 1,3,5,6 หรือ 2,4,5,7 และหลีกเลี่ยง 3,4,5,6',
    SUBJECT_DAILY_LIMIT: 'วิชาคาบเดี่ยวให้แยกคนละวัน ส่วนวิชาคาบคู่ห้ามเกิน 2 คาบต่อวัน',
    CUSTOM_RULE_BLOCKED: 'ตรวจข้อบังคับเพิ่มเติมของครู วิชา และห้อง',
  };
  const context = buildPeriodContext(periods);
  const maps = buildConstraintMaps(schedules, unavailabilities, classroomUnavailabilities,
    fixedSchedules, periods, customRules,activityGroups,classrooms,plcSessions,teachers);
  maps.enforceAutoSubjectDailyCap = true;
  maps.autoSubjectBlockSizes = new Map(subjects.map(s=>[String(s.id), Number(s.periodBlock)===2 ? 2 : 1]));
  const counts = new Map();
  let singleOptions = 0;
  DAYS.forEach(day => context.activePeriods.forEach(p => {
    const attempt = canPlaceSchedule(load.teacherId, load.classroomId, day, p.id, maps, false, load.subjectId);
    if (attempt.ok) singleOptions++;
    else counts.set(attempt.reason, (counts.get(attempt.reason)||0)+1);
  }));
  const topReasons = [...counts].sort((a,b)=>b[1]-a[1]).slice(0,3)
    .map(([code,count])=>({ code, count, label:labels[code] || translateUserMessage(code) }));
  let pairOptions = 0;
  if (load.reason === 'DOUBLE_PERIOD_NOT_AVAILABLE' || load.reason === 'BLOCK_COMPANION_NOT_AVAILABLE') {
    DAYS.forEach(day => context.doublePairs.forEach(([a,b]) => {
      const entries = [{day,periodId:a.id},{day,periodId:b.id}];
      if (entries.every(e=>canPlaceSchedule(load.teacherId,load.classroomId,e.day,e.periodId,maps,false,load.subjectId).ok) &&
          canPlaceTeacherEntries(load.teacherId,entries,maps).ok && canPlaceSubjectDailyEntries(load,entries,maps).ok) pairOptions++;
    }));
  }
  const mainCode = topReasons[0]?.code || '';
  const suggestion = load.reason === 'BLOCK_COMPANION_NOT_AVAILABLE'
    ? 'ตรวจคาบที่ล็อกหรือจัดเองซึ่งต้องมีคาบข้างเคียงสำหรับจับคู่'
    : load.reason === 'DOUBLE_PERIOD_NOT_AVAILABLE' && pairOptions === 0
      ? 'หาคาบว่างติดกัน 2 คาบ โดยอาจสลับคาบอื่นก่อน แล้วตรวจข้อบังคับเดิมซ้ำ'
      : hints[mainCode] || 'ลองจัดตารางใหม่หรือปรับคาบที่จัดด้วยมือ โดยตรวจข้อบังคับทั้งหมด';
  return { singleOptions, pairOptions, topReasons, suggestion };
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
  fixedItems: maps.fixedItems,
  teacherPlacements: new Set(maps.teacherPlacements),
  roomPlacements: new Set(maps.roomPlacements),
  teacherDayPeriods: new Map([...maps.teacherDayPeriods].map(([day, values]) => [day, new Set(values)])),
  teacherInstructionDayPeriods: new Map([...maps.teacherInstructionDayPeriods].map(([day, values]) => [day, new Set(values)])),
  subjectDayCounts: new Map(maps.subjectDayCounts),
  periodContext: maps.periodContext,
  customRules: maps.customRules,
  enforceAutoSubjectDailyCap: maps.enforceAutoSubjectDailyCap,
  autoSubjectBlockSizes: maps.autoSubjectBlockSizes,
  approvedManualTeachingStreaks: maps.approvedManualTeachingStreaks,
  maxTeacherPeriodsPerDay: maps.maxTeacherPeriodsPerDay,
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

// Stage 26: เพดาน 4/5 คาบนับกิจกรรม+คาบสอนรวมกัน; ไม่เอากิจกรรมไปนับว่า "สอนวิชา" 3 คาบติด
const groupTeacherDuties = (items = []) => {
  const byDay = new Map();
  (items || []).forEach(item => {
    if (!item.teacherId || !DAYS.includes(item.day)) return;
    const key = teacherDayKey(item.teacherId, item.day);
    if (!byDay.has(key)) byDay.set(key,{teacherId:item.teacherId,day:item.day,all:new Set(),lessons:new Set(),activities:new Set()});
    const row = byDay.get(key);
    const periodId=String(item.periodId);
    row.all.add(periodId);
    if (isTeacherActivityDuty(item)) row.activities.add(periodId);
    else row.lessons.add(periodId);
  });
  return byDay;
};
// Identify exact three-in-a-row runs (a lunch break is NOT consecutive).
// Approved manual exceptions are bound to a teacher, day and these exact period IDs;
// reusing a moved row's stale approval cannot authorize a different sequence.
const getTeacherTeachingRuns = (periodIds, periodsOrContext) => {
  const ctx = getPeriodContext(periodsOrContext);
  const ordered = [...new Set(periodIds.map(String))].filter(id => ctx.realIndex.has(id))
    .sort((a,b) => ctx.realIndex.get(a) - ctx.realIndex.get(b));
  const runs = []; let run = [];
  ordered.forEach(id => {
    if (run.length && ctx.realIndex.get(id) !== ctx.realIndex.get(run[run.length-1]) + 1) {
      runs.push(run); run = [];
    }
    run.push(id);
  });
  if (run.length) runs.push(run);
  return runs;
};
const getApprovedManualTeachingStreak = (items, teacherId, day, streak) =>
  (items || []).some(item => {
    const approval = item.manualConsecutiveApproval;
    return item.source === 'manual' && item.teacherId === teacherId && item.day === day &&
      approval?.teacherId === teacherId && approval?.day === day &&
      Array.isArray(approval.periodIds) && approval.periodIds.length === 3 &&
      streak.length === 3 && streak.every((id,index) => String(approval.periodIds[index]) === id) &&
      streak.includes(String(item.periodId));
  });
const findManualThreePeriodWarnings = (items, periods, teachers = []) => {
  const names = new Map(teachers.map(t=>[t.id,t.name]));
  const notices=[];
  groupTeacherDuties(items).forEach(({teacherId,day,lessons})=>{
    getTeacherTeachingRuns([...lessons],periods).forEach(streak=>{
      if (streak.length===3 && getApprovedManualTeachingStreak(items,teacherId,day,streak))
        notices.push(`ครู ${names.get(teacherId)||teacherId} วัน${day} สอนรายวิชา 3 คาบติดกัน (${streak.map(id=>getPeriodDisplayName(periods,id)).join(', ')}) — อนุมัติให้จัดด้วยมือแล้ว`);
    });
  });
  return notices;
};
const findTeacherTeachingLimitConflicts = (items, periods, teachers = []) => {
  const names = new Map(teachers.map(t => [t.id, t.name]));
  const violations = [];
  groupTeacherDuties(items).forEach(({teacherId,day,all,lessons}) => {
    const teacherName = names.get(teacherId) || teacherId;
    if (all.size > SCHEDULER_CONFIG.absoluteMaxTeacherPeriodsPerDay) {
      violations.push(`ครู ${teacherName} วัน${day} มีคาบสอนและกิจกรรมรวม ${all.size} คาบ (ห้ามเกิน ${SCHEDULER_CONFIG.absoluteMaxTeacherPeriodsPerDay} คาบ/วัน)`);
    }
    const lessonPeriods=[...lessons];
    const maxConsec = calculateDayTeachingShape(lessonPeriods, periods).maxConsec;
    const unapproved = getTeacherTeachingRuns(lessonPeriods,periods).filter(streak =>
      streak.length > SCHEDULER_CONFIG.maxConsecutiveTeacherPeriods &&
      !getApprovedManualTeachingStreak(items,teacherId,day,streak));
    if (unapproved.length) {
      violations.push(`ครู ${teacherName} วัน${day} สอนรายวิชาติดกัน ${maxConsec} คาบ (สูงสุด ${SCHEDULER_CONFIG.maxConsecutiveTeacherPeriods} คาบติดกัน ยกเว้น 3 คาบที่จัดด้วยมือและยืนยันแล้ว; กิจกรรมไม่ใช่คาบสอนรายวิชา)`);
    }
    const pattern = getFourPeriodTeachingPattern(lessonPeriods, periods);
    if ((pattern?.level === 'blocked' || hasSchoolPeriodSequence(lessonPeriods, periods, [3,4,5,6])) &&
        maxConsec <= SCHEDULER_CONFIG.maxConsecutiveTeacherPeriods) {
      violations.push(`ครู ${teacherName} วัน${day} มีรูปแบบคาบสอนรายวิชาผิดเกณฑ์ (${pattern?.key || 'ช่วง 3–4–5–6'})`);
    }
  });
  return violations;
};
// เมื่อคาบสอนต่อเนื่องกับกิจกรรมตั้งแต่ 3 ช่วงขึ้นไป เป็นคำเตือน ไม่ใช่ Hard Conflict
const findTeacherActivityConsecutiveWarnings = (items, periods, teachers = []) => {
  const names = new Map(teachers.map(t=>[t.id,t.name]));
  const ctx=getPeriodContext(periods);
  const messages=[];
  groupTeacherDuties(items).forEach(({teacherId,day,all,activities})=>{
    if (!activities.size || all.size<3) return;
    const ordered=[...all].filter(id=>ctx.realIndex.has(id)).sort((a,b)=>ctx.realIndex.get(a)-ctx.realIndex.get(b));
    let streak=[];
    const addWarning=()=>{
      if(streak.length>=3 && streak.some(id=>activities.has(id)))
        messages.push(`ครู ${names.get(teacherId)||teacherId} วัน${day} มีสอน/กิจกรรมต่อเนื่อง ${streak.length} คาบ (รวมกิจกรรมล็อก) ควรตรวจความเหมาะสม`);
    };
    ordered.forEach(id=>{
      if(streak.length && ctx.realIndex.get(id)!==ctx.realIndex.get(streak[streak.length-1])+1){addWarning();streak=[];}
      streak.push(id);
    });
    addWarning();
  });
  return messages;
};

const getTeacherDayShapeScore = (dayPeriods, periodsOrContext) => {
  if (!dayPeriods || dayPeriods.length === 0) return 0;
  const shape = calculateDayTeachingShape(dayPeriods, periodsOrContext);
  let score = shape.gaps * WEIGHTS.TEACHER_GAP;
  if (shape.maxConsec > SCHEDULER_CONFIG.maxConsecutiveTeacherPeriods) {
    score += (shape.maxConsec - SCHEDULER_CONFIG.maxConsecutiveTeacherPeriods) * WEIGHTS.TEACHER_CONSECUTIVE;
  }
  if (shape.gaps === 0 && dayPeriods.length > 1) score += WEIGHTS.NO_GAP;
  // ช่วยให้อัลกอริทึมเลือกรูปแบบกระจายตัวดี 15 แบบก่อน 8 แบบสำรอง
  // คะแนนเป็นความชอบ ไม่ยกเลิกกฎห้าม 12 แบบที่ตรวจใน canPlaceTeacherEntries
  const pattern = getFourPeriodTeachingPattern(dayPeriods, periodsOrContext);
  if (pattern?.level === 'preferred') score += 70;
  else if (pattern?.level === 'fallback') score -= 70;
  else if (pattern?.level === 'blocked') score -= 200;
  // รูปแบบ 1,4,5,6 และ 2,4,5,6 ยังอนุญาตตาม 35 รูปแบบเดิม แต่หลีกเลี่ยงเป็นพิเศษ
  if (pattern?.key === '1,4,5,6' || pattern?.key === '2,4,5,6') score -= 100;
  // สำหรับวันที่ต้องสอน 5 คาบ หลีกเลี่ยง 4–5–6 อย่างมากเช่นกัน
  if (dayPeriods.length === 5 && hasSchoolPeriodSequence(dayPeriods, periodsOrContext, [4,5,6])) score -= 220;
  const threePattern = getThreePeriodTeachingPattern(dayPeriods, periodsOrContext);
  if (threePattern?.level === 'preferred') score += 55;
  else if (threePattern?.level === 'fallback') score -= 25;
  else if (threePattern?.level === 'avoid') score -= 140;
  return score;
};

const buildPlacementState = (seedSchedules) => {
  const state = { teacherDayPeriods: new Map(), teacherInstructionDayPeriods:new Map(), teacherDayCounts: new Map(), roomSubjectDayCounts: new Map() };
  seedSchedules.forEach(sch => {
    if (sch.teacherId) {
      const tdKey = `${sch.teacherId}_${sch.day}`;
      if (!state.teacherDayPeriods.has(tdKey)) state.teacherDayPeriods.set(tdKey, new Set());
      state.teacherDayPeriods.get(tdKey).add(sch.periodId);
      if(!isTeacherActivityDuty(sch)){
        if(!state.teacherInstructionDayPeriods.has(tdKey))state.teacherInstructionDayPeriods.set(tdKey,new Set());
        state.teacherInstructionDayPeriods.get(tdKey).add(sch.periodId);
      }
      state.teacherDayCounts.set(tdKey, state.teacherDayPeriods.get(tdKey).size);
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
    if(!isTeacherActivityDuty(scheduleItem)){
      if(!state.teacherInstructionDayPeriods.has(tdKey))state.teacherInstructionDayPeriods.set(tdKey,new Set());
      state.teacherInstructionDayPeriods.get(tdKey).add(scheduleItem.periodId);
    }
    state.teacherDayCounts.set(tdKey, state.teacherDayPeriods.get(tdKey).size);
  }
  const rsdKey = `${scheduleItem.classroomId}_${scheduleItem.subjectId}_${scheduleItem.day}`;
  state.roomSubjectDayCounts.set(rsdKey, (state.roomSubjectDayCounts.get(rsdKey) || 0) + 1);
};

const scorePlacementOption = (entries, load, state, periodsOrContext, getSubject, customRules = []) => {
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
      const oldInstructions=Array.from(state.teacherInstructionDayPeriods.get(tdKey) || []);
      score += getTeacherDayShapeScore([...oldInstructions,...periodIds], periodContext) - getTeacherDayShapeScore(oldInstructions, periodContext);
      // รวมเวลาเว้นว่างระหว่างคาบสอนกับกิจกรรมเพื่อจัดตารางให้ครูเดินคาบสะดวก
      const oldGaps=calculateDayTeachingShape(oldPeriods,periodContext).gaps;
      const newGaps=calculateDayTeachingShape(newPeriods,periodContext).gaps;
      score -= (newGaps-oldGaps)*2;

      const oldCount = state.teacherDayCounts.get(tdKey) || 0;
      const newCount = oldCount + periodIds.length;
      const oldOver = Math.max(0, oldCount - SCHEDULER_CONFIG.maxTeacherPeriodsPerDay);
      const newOver = Math.max(0, newCount - SCHEDULER_CONFIG.maxTeacherPeriodsPerDay);
      score -= (newOver - oldOver) * 240;

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

  entries.forEach(entry => { score += customRulePreferenceScore(customRules, { ...load, ...entry }, periodContext); });
  return score;
};

const choosePlacementOption = (options, load, state, periodsOrContext, getSubject, random, customRules = []) => {
  if (!options.length) return null;
  const ranked = options.map(entries => ({
    entries,
    score: scorePlacementOption(entries, load, state, periodsOrContext, getSubject, customRules) + (random() * SCHEDULER_CONFIG.placementJitter),
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

const calculateScheduleScore = (schedules, teachers, subjects, periodsOrContext, getSubject, customRules = [], activityTeacherItems = []) => {
  let score = WEIGHTS.BASE_SCORE;
  const teacherStats = {}; const roomSubjectStats = {};
  const periodContext = getPeriodContext(periodsOrContext);
  const morningPeriods = periodContext.morningIds;

  schedules.forEach(sch => {
    score += customRulePreferenceScore(customRules, sch, periodContext);
    if (sch.teacherId) {
      if (!teacherStats[sch.teacherId]) teacherStats[sch.teacherId] = { days: {}, lessonDays: {} };
      if (!teacherStats[sch.teacherId].days[sch.day]) teacherStats[sch.teacherId].days[sch.day] = [];
      teacherStats[sch.teacherId].days[sch.day].push(sch.periodId);
      if (!teacherStats[sch.teacherId].lessonDays[sch.day]) teacherStats[sch.teacherId].lessonDays[sch.day]=[];
      teacherStats[sch.teacherId].lessonDays[sch.day].push(sch.periodId);
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

  // Activity group teachers occupy slots and must influence fairness just like lessons.
  activityTeacherItems.forEach(item => {
    if (!item.teacherId) return;
    if (!teacherStats[item.teacherId]) teacherStats[item.teacherId]={days:{},lessonDays:{}};
    if (!teacherStats[item.teacherId].days[item.day]) teacherStats[item.teacherId].days[item.day]=[];
    teacherStats[item.teacherId].days[item.day].push(item.periodId);
  });
  Object.values(teacherStats).forEach(tStat => {
    let totalPeriods = 0;
    Object.entries(tStat.days).forEach(([day,dayPeriods]) => {
      totalPeriods += new Set(dayPeriods).size;
      const overload = Math.max(0, new Set(dayPeriods).size - SCHEDULER_CONFIG.maxTeacherPeriodsPerDay);
      score -= overload * 240;
      // รูปแบบคาบสอนต่อเนื่องให้นับเฉพาะวิชา ไม่ลงโทษเมื่อกิจกรรมล็อกอยู่ติดกับวิชา
      score += getTeacherDayShapeScore(tStat.lessonDays[day] || [], periodContext);
    });
    if (totalPeriods > 0) {
      // Repeated 4/5-period days are tiring even if each individual day is valid.
      const heavyDays = Object.values(tStat.days).filter(rows => new Set(rows).size >= 4).length;
      score -= Math.max(0,heavyDays - 2) * 18;
      score -= Object.values(tStat.days).filter(rows => new Set(rows).size === 5).length * 350;
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

// STAGE 14: Dashboard derives class progress from the actual teaching loads,
// not from a blanket 35 periods per classroom. Activity locks and college-day
// closures are displayed separately and never counted as completed subject loads.
const buildDashboardOverview = ({ teachers = [], activeSubjects = [], classrooms = [],
  teachingLoads = [], schedules = [], fixedSchedules = [], activityGroups = [],
  classroomUnavailabilities = [], periods = [], health = {} }) => {
  const byRoom = new Map(classrooms.map(room => [room.id, {
    id: room.id, name: room.name || room.id, required: 0, scheduled: 0,
    missing: 0, overScheduled: 0, activities: 0, blockedSlots: 0,
    availableSlots: 0, filledSlots: 0, unfilledSlots: 0, status: 'unset'
  }]));
  const keyOf = item => JSON.stringify([String(item.teacherId || ''), String(item.subjectId || ''), String(item.classroomId || '')]);
  const loadMap = new Map();
  teachingLoads.forEach(load => {
    const key = keyOf(load);
    const req = Number(load.periods);
    if (!Number.isFinite(req) || req <= 0) return;
    const prev = loadMap.get(key) || { required: 0, scheduled: 0, classroomId: load.classroomId };
    prev.required += req;
    loadMap.set(key, prev);
    if (byRoom.has(load.classroomId)) byRoom.get(load.classroomId).required += req;
  });
  // Mirrors the health-check treatment of explicitly scheduled lessons and
  // room-specific fixed lessons. All-room activities do not fulfill subject loads.
  [...schedules, ...fixedSchedules.filter(item => item.classroomId !== ALL_CLASSROOMS_ID)]
    .forEach(item => {
      const load = loadMap.get(keyOf(item));
      if (load) load.scheduled += 1;
    });
  loadMap.forEach(load => {
    const room = byRoom.get(load.classroomId);
    if (!room) return;
    room.scheduled += Math.min(load.required, load.scheduled);
    room.missing += Math.max(0, load.required - load.scheduled);
    room.overScheduled += Math.max(0, load.scheduled - load.required);
  });
  const activitiesByRoom = new Map();
  fixedSchedules.forEach(fixed => {
    if (!DAYS.includes(fixed.day)) return;
    const targets = fixed.classroomId === ALL_CLASSROOMS_ID
      ? fixedTargetRoomIds(fixed, classrooms) : [fixed.classroomId];
    targets.forEach(roomId => {
      if (!byRoom.has(roomId)) return;
      if (!activitiesByRoom.has(roomId)) activitiesByRoom.set(roomId, new Set());
      activitiesByRoom.get(roomId).add(JSON.stringify([fixed.day, String(fixed.periodId)]));
    });
  });
  const blockedByRoom = new Map();
  classroomUnavailabilities.forEach(entry => {
    if (!byRoom.has(entry.classroomId)) return;
    if (!blockedByRoom.has(entry.classroomId)) blockedByRoom.set(entry.classroomId, new Set());
    blockedByRoom.get(entry.classroomId).add(JSON.stringify([entry.day, String(entry.periodId)]));
  });
  const activeSlots=DAYS.flatMap(day=>periods.filter(p=>!p.isBreak).map(p=>JSON.stringify([day,String(p.id)])));
  const slotsByRoom=new Map();
  schedules.forEach(item=>{
    if(!slotsByRoom.has(item.classroomId))slotsByRoom.set(item.classroomId,new Set());
    slotsByRoom.get(item.classroomId).add(JSON.stringify([item.day,String(item.periodId)]));
  });
  const rooms = [...byRoom.values()].map(room => {
    room.activities = activitiesByRoom.get(room.id)?.size || 0;
    room.blockedSlots = blockedByRoom.get(room.id)?.size || 0;
    const blocked=blockedByRoom.get(room.id)||new Set();
    const lessons=slotsByRoom.get(room.id)||new Set();
    const activities=activitiesByRoom.get(room.id)||new Set();
    room.availableSlots=activeSlots.filter(key=>!blocked.has(key)).length;
    room.filledSlots=activeSlots.filter(key=>!blocked.has(key)&&(lessons.has(key)||activities.has(key))).length;
    room.unfilledSlots=room.availableSlots-room.filledSlots;
    room.status = room.required === 0 ? 'unset' : room.missing === 0 && room.overScheduled === 0 ? 'complete' : 'incomplete';
    return room;
  }).sort((a, b) => a.name.localeCompare(b.name, 'th', { numeric: true }));
  const occupied = buildTeacherOccupancy(schedules, fixedSchedules, activityGroups);
  const byTeacherDay = new Map();
  occupied.forEach(item => {
    if (!item.teacherId || !DAYS.includes(item.day)) return;
    const key = JSON.stringify([item.teacherId, item.day]);
    if (!byTeacherDay.has(key)) byTeacherDay.set(key, { teacherId: item.teacherId, day: item.day, periodIds: new Set() });
    byTeacherDay.get(key).periodIds.add(String(item.periodId));
  });
  const names = new Map(teachers.map(t => [t.id, t.name || t.id]));
  const fiveDays = [];
  const heavyPatterns = [];
  byTeacherDay.forEach(row => {
    const count = row.periodIds.size;
    const ids = [...row.periodIds];
    const teacherName = names.get(row.teacherId) || row.teacherId;
    if (count === 5) fiveDays.push({ ...row, teacherName, count });
    const three = getThreePeriodTeachingPattern(ids, periods);
    const four = getFourPeriodTeachingPattern(ids, periods);
    if (three?.level === 'avoid') heavyPatterns.push({ teacherId: row.teacherId, teacherName, day: row.day, pattern: three.key });
    if (four?.level === 'blocked') heavyPatterns.push({ teacherId: row.teacherId, teacherName, day: row.day, pattern: four.key });
  });
  const hardConflictKeys = ['teacherConflicts', 'roomConflicts', 'unavailConflicts',
    'cUnavailConflicts', 'breakConflicts', 'blockConflicts', 'teacherLimitConflicts',
    'customRuleConflicts', 'subjectDailyConflicts', 'orphans'];
  const conflictCount = hardConflictKeys.reduce((sum, key) => sum + (health[key]?.length || 0), 0);
  return {
    rooms, fiveDays, heavyPatterns, conflictCount,
    roomsWithOpenSlots:rooms.filter(room=>room.unfilledSlots>0).length,
    completeRooms: rooms.filter(room => room.status === 'complete').length,
    incompleteRooms: rooms.filter(room => room.status === 'incomplete').length,
    unsetRooms: rooms.filter(room => room.status === 'unset').length,
    teacherCount: teachers.length, subjectCount: activeSubjects.length,
    roomCount: classrooms.length, loadCount: teachingLoads.length,
  };
};

const Dashboard = ({ schoolSettings, teachers, activeSubjects, classrooms, teachingLoads,
  schedules, fixedSchedules, activityGroups, classroomUnavailabilities, periods, health,
  onNavigate, onOpenRoom }) => {
  const [roomFilter, setRoomFilter] = useState('all');
  const [roomSearch, setRoomSearch] = useState('');
  const [showAllAlerts, setShowAllAlerts] = useState(false);
  const summary = useMemo(() => buildDashboardOverview({ teachers, activeSubjects, classrooms,
    teachingLoads, schedules, fixedSchedules, activityGroups, classroomUnavailabilities,
    periods, health }), [teachers, activeSubjects, classrooms, teachingLoads, schedules,
    fixedSchedules, activityGroups, classroomUnavailabilities, periods, health]);
  const filteredRooms = summary.rooms.filter(room =>
    (roomFilter === 'all' || room.status === roomFilter) &&
    room.name.toLocaleLowerCase().includes(roomSearch.trim().toLocaleLowerCase()));
  const statusLabels = {
    ready: 'ภาระสอนครบ ไม่มีข้อขัดแย้ง', warning: 'มีรายการต้องตรวจสอบ',
    empty: 'ยังไม่มีข้อมูลภาระสอน', error: 'พบข้อขัดแย้งของตาราง'
  };
  const currentStatus = health.status || 'empty';
  const progress = Number(health.completionPercent) || 0;
  const allAlerts = [
    ...summary.fiveDays.map(item => ({ type: 'five', key: `five_${item.teacherId}_${item.day}`,
      title: `${item.teacherName} · วัน${item.day}`, message: '5 คาบ/วัน — ข้อยกเว้นที่ควรตรวจสอบ' })),
    ...summary.heavyPatterns.map(item => ({ type: 'heavy', key: `heavy_${item.teacherId}_${item.day}`,
      title: `${item.teacherName} · วัน${item.day}`, message: `คาบ ${item.pattern.replaceAll(',', '–')} กระจุกตัว ควรปรับการกระจาย` }))
  ];
  return (
    <div className="space-y-6 animate-in fade-in duration-500">
      <div className="bg-gradient-to-br from-[#081a39] via-[#122b5e] to-[#081a39] rounded-3xl p-6 sm:p-8 text-white shadow-xl border border-[#d4af37]/30">
        <div className="flex flex-col md:flex-row md:items-center justify-between gap-5">
          <div className="space-y-2">
            <h1 className="text-3xl sm:text-4xl font-bold">ภาพรวมการจัดตารางสอน</h1>
            <p className="text-[#e6c55c] font-medium">{String(schoolSettings?.schoolName || 'โรงเรียน')} · ภาคเรียนที่ {String(schoolSettings?.semester || '1')}/{String(schoolSettings?.academicYear || '2569')}</p>
            <p className="text-sm text-slate-200">ติดตามภาระสอนรายวิชา สถานะห้องเรียน และรายการที่ต้องตรวจสอบ</p>
          </div>
          <div className="bg-white/10 p-4 rounded-2xl border border-white/20 md:min-w-[215px]">
            <p className="text-sm font-bold flex items-center gap-2"><span className={`w-2.5 h-2.5 rounded-full ${currentStatus === 'ready' ? 'bg-emerald-400' : currentStatus === 'empty' ? 'bg-slate-300' : 'bg-amber-400'}`}></span>{statusLabels[currentStatus] || 'กำลังตรวจสอบ'}</p>
            <p className="text-4xl font-bold text-[#e6c55c] mt-2">{progress}%</p>
            <p className="text-xs text-slate-200 mt-1">จัดภาระสอนรายวิชา {health.scheduledForLoads || 0} / {health.totalRequired || 0} คาบ</p>
          </div>
        </div>
        <div className="mt-5 h-2.5 rounded-full bg-white/15 overflow-hidden" role="progressbar" aria-label="ความคืบหน้าการจัดภาระสอน" aria-valuenow={progress} aria-valuemin={0} aria-valuemax={100}>
          <div className="h-full bg-[#d4af37] rounded-full transition-all" style={{ width: `${Math.max(0, Math.min(100, progress))}%` }}></div>
        </div>
      </div>

      {activeSubjects.some(subject=>subject.needsScheduleReview && teachingLoads.some(load=>load.subjectId===subject.id)) && <div className="bg-amber-50 border border-amber-200 rounded-xl p-4 flex flex-wrap items-center gap-3">
        <AlertTriangle size={18} className="text-amber-700"/>
        <span className="text-sm text-amber-900 flex-1">มีรายวิชาใหม่จาก Excel ที่ยังไม่ได้ตรวจคาบเดี่ยว/คาบคู่และช่วงเวลาที่เหมาะสม ต้องตรวจในเมนูรายวิชาก่อนจัดอัตโนมัติ</span>
        <button onClick={()=>onNavigate('basicData')} className="px-3 py-2 rounded-lg bg-amber-700 text-white text-sm font-bold">เปิดข้อมูลพื้นฐาน</button>
      </div>}
      <div className="grid grid-cols-2 lg:grid-cols-4 gap-3 sm:gap-4">
        {[
          {label:'ครูทั้งหมด', count: summary.teacherCount, unit:'คน', icon: Users},
          {label:'ห้องเรียนทั้งหมด', count: summary.roomCount, unit:'ห้อง', icon: Home},
          {label:'รายวิชาในภาคเรียน', count: summary.subjectCount, unit:'วิชา', icon: BookOpen},
          {label:'คาบรายวิชาที่ยังไม่จัด', count: health.missing || 0, unit:'คาบ', icon: CalendarDays}
        ].map(item => {
          const Icon = item.icon;
          return <div key={item.label} className="bg-white p-4 sm:p-5 rounded-2xl shadow-sm border border-slate-100">
            <div className="flex items-center justify-between gap-2 text-xs sm:text-sm font-semibold text-slate-500"><span>{item.label}</span><Icon size={18}/></div>
            <div className="text-3xl font-bold text-[#081a39] mt-3">{item.count} <span className="text-sm font-normal text-slate-500">{item.unit}</span></div>
          </div>;
        })}
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-3 gap-4">
        <div className="lg:col-span-2 bg-white border border-slate-100 rounded-2xl p-5 sm:p-6 shadow-sm">
          <div className="flex flex-wrap justify-between items-center gap-3 mb-4">
            <div><h2 className="text-xl font-bold text-[#081a39]">สถานะการจัดตารางรายห้อง</h2><p className="text-xs text-slate-500 mt-1">นับเฉพาะภาระสอนรายวิชาที่กำหนดให้ห้อง ไม่ใช้เป้าหมาย 35 คาบเท่ากันทุกห้อง</p></div>
            <button onClick={() => onNavigate('schedule')} className="text-sm font-bold text-blue-700 hover:underline">ไปหน้าจัดตาราง →</button>
          </div>
          <p className="text-xs text-blue-800 bg-blue-50 border border-blue-100 rounded-xl px-3 py-2 mb-3">แยก 2 สถานะ: <strong>ภาระวิชาจัดครบ</strong> กับ <strong>ช่องในตารางเรียนยังว่าง</strong> · ช่องว่างอาจตั้งใจเว้นไว้ได้ ไม่ถือว่าผิดโดยอัตโนมัติ</p>
          <div className="grid grid-cols-3 gap-2 mb-4 text-center">
            <div className="bg-emerald-50 border border-emerald-100 rounded-xl p-3"><p className="font-bold text-2xl text-emerald-700">{summary.completeRooms}</p><p className="text-xs text-emerald-800">ภาระครบ</p></div>
            <div className="bg-amber-50 border border-amber-100 rounded-xl p-3"><p className="font-bold text-2xl text-amber-700">{summary.incompleteRooms}</p><p className="text-xs text-amber-800">ยังไม่ครบ</p></div>
            <div className="bg-slate-50 border border-slate-200 rounded-xl p-3"><p className="font-bold text-2xl text-slate-700">{summary.unsetRooms}</p><p className="text-xs text-slate-600">ยังไม่กำหนด</p></div>
          </div>
          <div className="flex flex-col sm:flex-row gap-2 mb-3">
            <select aria-label="กรองสถานะห้องเรียน" value={roomFilter} onChange={e => setRoomFilter(e.target.value)} className="border border-slate-200 rounded-xl px-3 py-2 text-sm bg-white">
              <option value="all">ทุกห้องเรียน</option><option value="incomplete">ยังไม่ครบ</option><option value="complete">ภาระครบ</option><option value="unset">ยังไม่กำหนด</option>
            </select>
            <label className="flex-1 flex items-center gap-2 px-3 border border-slate-200 rounded-xl"><Search size={17} className="text-slate-400"/><input aria-label="ค้นหาห้องเรียน" value={roomSearch} onChange={e => setRoomSearch(e.target.value)} className="w-full py-2 outline-none text-sm bg-transparent" placeholder="ค้นหา เช่น ม.4/3" /></label>
          </div>
          <div className="max-h-[410px] overflow-y-auto space-y-2 pr-1 custom-scrollbar">
            {filteredRooms.length === 0 && <p className="p-5 text-center text-slate-500 text-sm">ไม่พบห้องเรียนตามตัวกรอง</p>}
            {filteredRooms.map(room => {
              const percent = room.required ? Math.round(room.scheduled / room.required * 100) : 0;
              return <button key={room.id} type="button" onClick={() => onOpenRoom(room.id)} className="w-full text-left border border-slate-100 bg-slate-50/60 hover:bg-blue-50 hover:border-blue-200 transition rounded-xl p-3 focus-visible:outline-2 focus-visible:outline-blue-600" aria-label={`เปิดตารางห้อง ${room.name}`}>
                <div className="flex items-center justify-between gap-3">
                  <div className="font-bold text-[#081a39] text-sm">{room.name}</div>
                  <div className="flex items-center gap-2"><span className={`text-xs font-semibold px-2 py-1 rounded-full ${room.status==='complete'?'bg-emerald-100 text-emerald-700':room.status==='incomplete'?'bg-amber-100 text-amber-800':'bg-slate-200 text-slate-600'}`}>{room.status==='complete'?'ภาระครบ':room.status==='incomplete'?(room.missing > 0 ? `ขาด ${room.missing} คาบ` : 'จัดเกินภาระ'):'ยังไม่กำหนดภาระ'}</span><span className="text-sm text-blue-600">↗</span></div>
                </div>
                <div className="flex items-center gap-3 mt-2"><div className="flex-1 bg-slate-200 h-2 rounded-full overflow-hidden"><div className={`h-full rounded-full ${room.status==='complete'?'bg-emerald-500':'bg-amber-500'}`} style={{width:`${Math.min(100,percent)}%`}}></div></div><span className="text-xs text-slate-600 font-medium whitespace-nowrap">{room.scheduled}/{room.required} คาบ</span></div>
                <p className="text-xs text-slate-600 mt-1.5">ตารางเรียน: ใช้แล้ว {room.filledSlots}/{room.availableSlots} ช่อง (ไม่รวมเวลาปิดเรียน){room.unfilledSlots>0?` · เหลือ ${room.unfilledSlots} ช่องให้ตรวจ`: ' · ไม่มีช่องว่าง'}</p>
                {(room.activities > 0 || room.blockedSlots > 0 || room.overScheduled > 0) && <p className="text-xs text-slate-500 mt-1">{[room.activities > 0 ? `กิจกรรมล็อก ${room.activities} คาบ` : '',room.blockedSlots > 0 ? `ปิดเวลา ${room.blockedSlots} คาบ` : '',room.overScheduled > 0 ? `จัดเกินภาระ ${room.overScheduled} คาบ` : ''].filter(Boolean).join(' · ')}</p>}
              </button>;
            })}
          </div>
          <p className="text-xs text-slate-500 mt-3">กดชื่อห้องเพื่อเปิดตารางเรียนของห้องนั้น · กิจกรรมและช่วงเวลาปิดห้องแยกจากภาระรายวิชา</p>
        </div>

        <div className="space-y-4">
          <div className="bg-white border border-slate-100 rounded-2xl p-5 shadow-sm">
            <h2 className="font-bold text-lg text-[#081a39] flex gap-2 items-center"><AlertTriangle size={20} className={summary.conflictCount ? 'text-rose-500' : 'text-slate-400'}/>รายการที่ต้องตรวจสอบ</h2>
            <div className="grid grid-cols-2 gap-2 mt-4">
              <div className="bg-rose-50 rounded-xl p-3"><p className="text-2xl font-bold text-rose-700">{summary.conflictCount}</p><p className="text-xs text-rose-700">ข้อขัดแย้ง</p></div>
              <div className="bg-amber-50 rounded-xl p-3"><p className="text-2xl font-bold text-amber-700">{summary.fiveDays.length}</p><p className="text-xs text-amber-700">วันภาระรวม 5 คาบ</p></div>
            </div>
            {!!summary.heavyPatterns.length && <p className="text-sm text-amber-700 mt-3">พบรูปแบบคาบที่ควรกระจายใหม่ {summary.heavyPatterns.length} วัน-ครู</p>}
            {summary.conflictCount === 0 && allAlerts.length === 0 && health.missing === 0 && teachingLoads.length > 0 && <p className="text-sm text-emerald-700 mt-3 flex gap-2 items-center"><CheckCircle2 size={17}/> ไม่พบรายการเตือนจากข้อมูลปัจจุบัน</p>}
            {health.missing > 0 && <p className="text-sm text-amber-700 mt-3">ยังเหลือภาระสอน {health.missing} คาบที่ต้องจัด</p>}
            <div className="space-y-2 mt-3">
              {(showAllAlerts ? allAlerts : allAlerts.slice(0, 4)).map(alert => <div key={alert.key} className="border-l-2 border-amber-400 pl-3 py-1"><p className="text-sm font-semibold text-slate-800">{alert.title}</p><p className="text-xs text-slate-600">{alert.message}</p></div>)}
              {allAlerts.length > 4 && <button onClick={() => setShowAllAlerts(s => !s)} className="text-sm text-blue-700 font-semibold hover:underline">{showAllAlerts ? 'แสดงน้อยลง' : `ดูทั้งหมด ${allAlerts.length} รายการ`}</button>}
            </div>
            <button onClick={() => onNavigate('reports')} className="w-full mt-4 border border-slate-200 hover:bg-slate-50 text-[#081a39] text-sm font-bold py-2.5 rounded-xl">เปิดรายงานตรวจตาราง →</button>
          </div>
          <div className="bg-white border border-slate-100 rounded-2xl p-5 shadow-sm">
            <h2 className="font-bold text-lg text-[#081a39] mb-3">ทำงานต่ออย่างรวดเร็ว</h2>
            <div className="grid grid-cols-2 gap-2">
              {[
                {id:'loads',label:'เพิ่มภาระสอน',icon: Briefcase},
                {id:'schedule',label:'จัดตาราง',icon:CalendarDays},
                {id:'reports',label:'ตรวจตาราง',icon:Activity},
                {id:'print',label:'พิมพ์ตาราง',icon:Printer}
              ].map(link => { const Icon=link.icon;return <button key={link.id} onClick={() => onNavigate(link.id)} className="flex flex-col items-start gap-2 p-3 text-left bg-slate-50 hover:bg-blue-50 border border-slate-100 rounded-xl text-[#081a39] font-semibold text-sm"><Icon size={20} className="text-blue-700"/>{link.label}</button>; })}
            </div>
          </div>
        </div>
      </div>
      <p className="text-xs text-slate-500">สถานะ “ภาระครบ” หมายถึงจัดคาบรายวิชาตรงตามภาระสอนของห้อง ไม่ใช่การรับรองว่าไม่มีคาบชน กรุณาตรวจส่วนแจ้งเตือนประกอบ · ตัวเลขใช้ข้อมูลภาคเรียนที่เลือกในระบบ</p>
    </div>
  );
};

const SettingsView = ({ schoolSettings, db, user, appId, showToast, getTermKey, fullDataForBackup, handleRestoreBackup }) => {
  const [form, setForm] = useState(schoolSettings);
  const fileInputRef = useRef(null);
  // Step 0: indicate a server-verified backup is in progress.
  const [isBackingUp, setIsBackingUp] = useState(false);
  useEffect(() => { setForm(schoolSettings); }, [schoolSettings]);

  const handleSave = async () => {
    if (!user) return;
    try { await setDoc(doc(db, `school_data/${appId}/config/settings`), form); showToast("บันทึกการตั้งค่าแล้ว"); } 
    catch (e) { showToast("เกิดข้อผิดพลาดในการบันทึก", "error"); }
  };

  // STEP 00: Export all nine collections from the FIRESTORE SERVER, never from
  // partially initialized React listeners or stale term data. Preserve JSON v2.
  // This is a best-effort multi-collection export, NOT an atomic database snapshot.
  const handleBackup = async () => {
    if (isBackingUp) return;
    if (!user) return showToast("กรุณาเชื่อมต่อระบบก่อนสำรองข้อมูล", "error");
    setIsBackingUp(true);
    try {
      const basePath = `school_data/${appId}`;
      const settingsRef = doc(db, `${basePath}/config/settings`);
      const settingsSnap = await getDocFromServer(settingsRef);
      if (!settingsSnap.exists()) throw new Error('BACKUP_SETTINGS_MISSING');
      const serverSettings = settingsSnap.data();
      const termKey = getTermKey(serverSettings);
      if (termKey !== getTermKey(schoolSettings)) throw new Error('BACKUP_TERM_CHANGED');
      const names = ['teachers', 'subjects', 'classrooms', 'periods', 'teachingLoads', 'unavailabilities', 'classroomUnavailabilities', 'fixedSchedules', 'schedules', 'scheduleRules', 'activityGroups', 'plcSessions'];
      const masterNames = new Set(['teachers', 'subjects', 'classrooms', 'periods']);
      const entries = await Promise.all(names.map(async name => {
        const path = masterNames.has(name) ? `${basePath}/${name}` : `${basePath}/terms/${termKey}/${name}`;
        const snapshot = await getDocsFromServer(collection(db, path));
        return [name, snapshot.docs.map(item => ({ ...item.data(), id: item.id }))];
      }));
      // Reject if the active term/settings changed while downloading the data.
      const finalSettingsSnap = await getDocFromServer(settingsRef);
      if (!finalSettingsSnap.exists() || stableSerializeData(serverSettings) !== stableSerializeData(finalSettingsSnap.data())) {
        throw new Error('BACKUP_SETTINGS_CHANGED');
      }
      const backupData = {
        backupVersion: 2, appId, academicYear: serverSettings.academicYear,
        semester: serverSettings.semester, termKey, exportedAt: new Date().toISOString(),
        schoolSettings: serverSettings, ...Object.fromEntries(entries),
      };
      const blob = new Blob([JSON.stringify(backupData, null, 2)], { type: "application/json" });
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = `BHS-Backup-${termKey}-${new Date().toISOString().split('T')[0]}.json`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);
      showToast("ดาวน์โหลดไฟล์สำรองข้อมูลจากเซิร์ฟเวอร์สำเร็จ");
    } catch (error) {
      console.error('Server backup failed', error);
      showToast("สำรองข้อมูลไม่สำเร็จ ไม่ได้สร้างไฟล์สำรองที่อาจไม่ครบ กรุณาตรวจการเชื่อมต่อและภาคเรียน", "error");
    } finally {
      setIsBackingUp(false);
    }
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
          <button onClick={handleBackup} disabled={isBackingUp} className="flex items-center justify-center gap-2 py-3 bg-slate-50 border border-slate-200 text-[#081a39] rounded-xl hover:bg-slate-100 font-bold transition-all shadow-sm disabled:opacity-50 disabled:cursor-wait"><Download size={18} className="text-emerald-600" /> {isBackingUp ? "กำลังอ่านข้อมูลจากเซิร์ฟเวอร์..." : "สำรองข้อมูล (JSON)"}</button>
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

const SubjectsView = ({ activeSubjects, subjects, teachingLoads, schedules, fixedSchedules, bulkDeleteSubjects, renameSubjectId, dbAction, showToast, handleRequestDelete, schoolSettings, onOpenDevelopment }) => {
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [form, setForm] = useState({ id: '', name: '', abbr: '', periodsPerWeek: 1, term: '1', periodBlock: 1, preferredTime: 'any' });
  const [isEditing, setIsEditing] = useState(false);
  const [originalSubjectId, setOriginalSubjectId] = useState('');
  const [newItemKind, setNewItemKind] = useState('subject'); // UI only; activities use the existing fixed-activity editor
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
    if (!isEditing && newItemKind !== 'subject') return; // Never save activities into subjects collection
    const nextId = String(form.id || '').trim();
    const periodsPerWeek = Number(form.periodsPerWeek);
    if (!nextId || !String(form.name || '').trim()) return showToast("กรอกข้อมูลให้ครบ", "error");
    if (nextId.includes('/') || nextId === '.' || nextId === '..') return showToast("รหัสวิชาไม่สามารถมี / หรือเป็น . หรือ .. ได้", "error");
    if (!Number.isInteger(periodsPerWeek) || periodsPerWeek < 1) return showToast("จำนวนคาบ/สัปดาห์ต้องเป็นจำนวนเต็มบวก", "error");
    if (subjects.some(s => s.id === nextId && (!isEditing || s.id !== originalSubjectId))) return showToast("รหัสวิชานี้มีอยู่แล้ว", "error");
    const updated = { ...form, id: nextId, name: form.name.trim(), periodsPerWeek, periodBlock: Number(form.periodBlock),
      needsScheduleReview: false }; // saving the subject editor confirms Block/preferred time
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
        <button onClick={() => {setForm({id:'', name:'', abbr:'', periodsPerWeek: 1, term: schoolSettings.semester, periodBlock: 1, preferredTime: 'any'}); setNewItemKind('subject'); setIsEditing(false); setOriginalSubjectId(''); setIsModalOpen(true);}} className="flex items-center gap-2 px-4 py-2 bg-gradient-to-r from-[#081a39] to-[#122b5e] text-white rounded-xl shadow-md font-bold"><Plus size={18} className="text-[#d4af37]"/> เพิ่มรายวิชา / กิจกรรม</button>
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
                <td className="p-4 text-slate-500 font-mono text-sm">{d.id}</td><td className="p-4 font-medium text-[#081a39]">{d.name} <span className="text-xs text-slate-400">({d.abbr})</span>{d.needsScheduleReview&&<span className="ml-2 bg-amber-100 text-amber-800 rounded px-2 py-1 text-xs">ตรวจ Block และช่วงเวลา ก่อน Auto</span>}{!deletableSet.has(d.id) && <span className="ml-2 text-xs text-amber-700">(ถูกใช้งานอยู่)</span>}</td><td className="p-4 text-center font-bold text-[#d4af37]">{d.periodsPerWeek}</td>
                <td className="p-4 text-center"><button onClick={() => { setForm({periodBlock: 1, preferredTime: 'any', ...d}); setOriginalSubjectId(d.id); setIsEditing(true); setIsModalOpen(true); }} className="p-2 text-slate-400 hover:text-[#081a39]"><Edit size={18}/></button><button onClick={() => handleRequestDelete('subject', d)} className="p-2 text-slate-400 hover:text-rose-500"><Trash2 size={18}/></button></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <Modal isOpen={isModalOpen} onClose={() => setIsModalOpen(false)} title={isEditing ? 'แก้ไขรายวิชา' : 'เพิ่มรายวิชา / กิจกรรม'}>
        <div className="space-y-4">
          {!isEditing && <div><label className="block text-sm font-bold text-[#081a39] mb-1" htmlFor="bhs-subject-kind">ประเภทข้อมูล</label>
            <select id="bhs-subject-kind" value={newItemKind} onChange={e=>setNewItemKind(e.target.value)} className="w-full p-3 border border-slate-200 rounded-xl bg-slate-50">
              <option value="subject">รายวิชาพื้นฐาน/เพิ่มเติม</option>
              <option value="development">กิจกรรมพัฒนาผู้เรียน</option>
            </select>
          </div>}
          {isEditing || newItemKind === 'subject' ? <>
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
          </> : <div className="bg-blue-50 border border-blue-200 p-4 rounded-xl space-y-3">
            <p className="font-bold text-blue-900">กิจกรรมพัฒนาผู้เรียนไม่ต้องกรอกรหัสวิชา</p>
            <p className="text-sm text-blue-900">ใช้แบบฟอร์มกิจกรรมเดิมเพื่อระบุชื่อ วัน คาบ ห้องเรียน และครูผู้รับผิดชอบ โดยไม่สร้างรายวิชาซ้ำและไม่เพิ่มภาระสอนรายวิชา</p>
            <button type="button" onClick={() => { setIsModalOpen(false); onOpenDevelopment?.(); }} className="w-full bg-[#081a39] text-white py-3 rounded-xl font-bold">ไปเพิ่ม / จัดการกิจกรรมพัฒนาผู้เรียน</button>
          </div>}
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
const validSubjectCode = id => !!id && id !== '.' && id !== '..' && !id.includes('/');
const isSubjectActiveInTerm = (subject, semester) => !!subject && (!subject.term || subject.term === 'all' || String(subject.term) === String(semester));
const defaultImportedSubject = ({ id, name, periods, semester, periodBlock = 1, preferredTime = 'any', abbr = '', needsScheduleReview = false }) => ({
  id: normalizeExcelText(id),
  name: normalizeExcelText(name),
  abbr: normalizeExcelText(abbr) || normalizeExcelText(name),
  periodsPerWeek: Number(periods),
  term: String(semester),
  periodBlock: Number(periodBlock),
  preferredTime,
  needsScheduleReview, // Excel-created subjects must be reviewed before Auto
});
const classifySubjectRows = (rows, subjects, semester) => {
  const known = new Map(subjects.map(s => [normalizeExcelText(s.id), s]));
  const created = new Map();
  const errors = [];
  rows.forEach(({ subjectId, subjectName, periods, prefix }) => {
    const id = normalizeExcelText(subjectId);
    const name = normalizeExcelText(subjectName);
    const label = prefix || id;
    if (!validSubjectCode(id) || !name) { errors.push(`${label}: รหัสวิชาหรือชื่อวิชาไม่ถูกต้อง`); return; }
    const found = known.get(id) || created.get(id);
    if (found) {
      if (normalizeExcelText(found.name) !== name) errors.push(`${label}: รหัสวิชา ${id} มีชื่อไม่ตรงกัน (${found.name} / ${name})`);
      else if (!isSubjectActiveInTerm(found, semester)) errors.push(`${label}: รหัส ${id} มีอยู่แล้วแต่ไม่เปิดใช้ในภาคเรียนนี้ กรุณาจัดการในเมนูรายวิชา`);
    } else {
      const value = Number(periods);
      if (!Number.isInteger(value) || value < 1) errors.push(`${label}: จำนวนคาบ/สัปดาห์ไม่ถูกต้อง`);
      else created.set(id, defaultImportedSubject({ id, name, periods: value, semester, needsScheduleReview:true }));
    }
  });
  return { subjectsToCreate: [...created.values()], errors };
};

const LoadsView = ({ teachingLoads, teachers, subjects, activeSubjects, classrooms, saveTeachingLoadsWithSubjects, setConfirmData, showToast, handleRequestDelete, schoolSettings, getTeacher, getSubject, getClassroom }) => {
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [isEditing, setIsEditing] = useState(false);
  const [isImporting, setIsImporting] = useState(false);
  const [form, setForm] = useState({ id: '', teacherId: '', subjectId: '', classroomId: '', periods: 1 });
  const [createNewSubject, setCreateNewSubject] = useState(false);
  const [newSubject, setNewSubject] = useState({ id: '', name: '', abbr: '', periodBlock: 1, preferredTime: 'any' });
  const [isSaving, setIsSaving] = useState(false);
  const [originalLoad, setOriginalLoad] = useState(null);
  const resetNewSubject = () => { setCreateNewSubject(false); setNewSubject({ id: '', name: '', abbr: '', periodBlock: 1, preferredTime: 'any' }); };
  const excelInputRef = useRef(null);
  const [selectedTeacherId, setSelectedTeacherId] = useState('all');
  // Table-only view: retain original teachingLoad records and Firestore write paths.
  const loadDisplayRows = teachingLoads.map(load => ({
    load, id: load.id, teacherName: getTeacher(load.teacherId).name || '',
    subjectName: getSubject(load.subjectId).name || '',
    classroomName: getClassroom(load.classroomId).name || '',
    classroomSortKey: (getClassroom(load.classroomId).name || '').replace(/\d+/g, digits => digits.padStart(5, '0')),
    periods: Number(load.periods) || 0
  }));
  const { items: sortedLoadRows, requestSort: sortLoads, sortConfig: loadSortConfig } = useSortableData(loadDisplayRows, { key: 'teacherName', direction: 'asc' });
  const displayedLoadRows = selectedTeacherId === 'all'
    ? sortedLoadRows : sortedLoadRows.filter(row => row.load.teacherId === selectedTeacherId);
  const displayedPeriods = displayedLoadRows.reduce((sum, row) => sum + row.periods, 0);
  useEffect(() => {
    if (selectedTeacherId !== 'all' && !teachers.some(t => t.id === selectedTeacherId)) setSelectedTeacherId('all');
  }, [selectedTeacherId, teachers]);
  useEffect(() => { setSelectedTeacherId('all'); }, [schoolSettings.academicYear, schoolSettings.semester]);

  const handleSave = async () => {
    if (isSaving) return;
    const periodsValue = Number(form.periods);
    const subjectId = createNewSubject ? normalizeExcelText(newSubject.id) : form.subjectId;
    if (!form.teacherId || !subjectId || !form.classroomId || !Number.isInteger(periodsValue) || periodsValue < 1) return showToast('กรุณากรอกข้อมูลให้ครบ และกำหนดจำนวนคาบเป็นจำนวนเต็มบวก', 'error');
    if (!validSubjectCode(subjectId)) return showToast('รหัสวิชาไม่ถูกต้อง (ห้าม / และรหัส . หรือ ..)', 'error');
    let proposedSubjects = [];
    if (createNewSubject) {
      const subjectName = normalizeExcelText(newSubject.name);
      if (!subjectName) return showToast('กรุณากรอกชื่อรายวิชาใหม่', 'error');
      const result = classifySubjectRows([{ subjectId, subjectName, periods: periodsValue, prefix: 'วิชาใหม่' }], subjects, schoolSettings.semester);
      if (result.errors.length) return showToast(result.errors[0], 'error');
      // Explicitly choosing 'new' must not silently reuse an existing subject.
      if (!result.subjectsToCreate.length) return showToast('รหัสวิชานี้มีอยู่แล้ว กรุณาเลือกจากรายวิชาเดิม', 'error');
      proposedSubjects = [defaultImportedSubject({ id: subjectId, name: subjectName, abbr: newSubject.abbr,
        periods: periodsValue, semester: schoolSettings.semester, periodBlock: newSubject.periodBlock, preferredTime: newSubject.preferredTime })];
    }
    const isDuplicate = teachingLoads.some(l => l.teacherId === form.teacherId && l.subjectId === subjectId && l.classroomId === form.classroomId && l.id !== form.id);
    if (isDuplicate) return showToast('มีภาระสอนรายการนี้อยู่แล้ว กรุณาแก้ไขรายการเดิม', 'error');
    const id = isEditing && form.id ? form.id : `L_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`;
    setIsSaving(true);
    try {
      const item = { id, teacherId: form.teacherId, subjectId, classroomId: form.classroomId, periods: periodsValue,
        expectedExisting: isEditing, expectedPreviousTriple: isEditing && originalLoad ? getTeachingLoadKey(originalLoad) : '',
        expectedPreviousPeriods: isEditing && originalLoad ? Number(originalLoad.periods) : undefined };
      const nameAssertions = !createNewSubject ? { [subjectId]: activeSubjects.find(s=>s.id===subjectId)?.name || '' } : {};
      const ok = await saveTeachingLoadsWithSubjects([item], proposedSubjects, nameAssertions);
      if (ok) {
        setForm({ id: '', teacherId: '', subjectId: '', classroomId: '', periods: 1 });
        resetNewSubject(); setOriginalLoad(null); setIsModalOpen(false);
        showToast(proposedSubjects.length ? 'สร้างรายวิชาและบันทึกภาระสอนสำเร็จ' : (isEditing ? 'บันทึกการแก้ไขสำเร็จ' : 'เพิ่มภาระสอนสำเร็จ'));
      }
    } finally { setIsSaving(false); }
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
        ['4', 'หากรหัสวิชาใหม่ ระบบจะสร้างรายวิชาให้อัตโนมัติ (รูปแบบคาบเดี่ยว เวลาใดก็ได้); หากรหัสเดิม ชื่อต้องตรงกัน'],
        ['5', 'ชื่อห้องต้องตรงกับชีต "ห้องเรียน" (ชื่อห้องในระบบต้องไม่ซ้ำกัน)'],
        ['6', 'จำนวนคาบ/สัปดาห์ต้องเป็นจำนวนเต็มตั้งแต่ 1 ขึ้นไป'],
        ['7', 'ห้ามมีครู+วิชา+ห้องซ้ำกันมากกว่า 1 แถวในไฟล์เดียว'],
        ['8', 'ระบบแสดงตัวอย่างจำนวนวิชาใหม่และภาระสอนก่อนยืนยัน; สามารถปรับคาบคู่/เวลาเหมาะสมที่เมนูรายวิชาภายหลัง'],
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
      const subjectById = new Map(subjects.map(s => [normalizeExcelText(s.id), s]));
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
      const subjectRows = [];
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
        if (subject && normalizeExcelText(subject.name) !== subjectName) {
          errors.push(`${prefix}: ชื่อวิชาไม่ตรงกับรหัส ${subjectId} (${subject.name})`);
          return;
        }
        if (subject && !isSubjectActiveInTerm(subject, schoolSettings.semester)) {
          errors.push(`${prefix}: รหัส ${subjectId} มีอยู่แล้วแต่ไม่เปิดใช้ในภาคเรียนนี้`);
          return;
        }
        if (!validSubjectCode(subjectId)) {
          errors.push(`${prefix}: รหัสวิชาไม่ถูกต้อง`);
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
        importedItems.push({ id, teacherId, subjectId, classroomId, periods: periodsValue,
          expectedExisting: Boolean(existing), expectedPreviousTriple: existing ? getTeachingLoadKey(existing) : '',
          expectedPreviousPeriods: existing ? Number(existing.periods) : undefined });
        subjectRows.push({ subjectId, subjectName, periods: periodsValue, prefix });
        if (existing) updateCount += 1;
        else createCount += 1;
      });

      if (errors.length > 0) {
        const preview = errors.slice(0, 3).join(' | ');
        const more = errors.length > 3 ? ` และอีก ${errors.length - 3} จุด` : '';
        showToast(`นำเข้าไม่ได้ พบ ${errors.length} จุด: ${preview}${more}`, 'error');
        return;
      }

      const classification = classifySubjectRows(subjectRows, subjects, schoolSettings.semester);
      if (classification.errors.length) {
        showToast(`ตรวจรายวิชาไม่ผ่าน: ${classification.errors.slice(0, 3).join(' | ')}`, 'error');
        return;
      }
      const newSubjects = classification.subjectsToCreate;
      const sampleNames = newSubjects.slice(0, 8).map(s => `${s.id} ${s.name}`).join('\n');
      const moreNames = newSubjects.length > 8 ? `\n...และอีก ${newSubjects.length - 8} วิชา` : '';
      setConfirmData({
        type: 'restore', title: 'ตรวจสอบก่อนนำเข้า Excel', confirmText: 'ยืนยันนำเข้า',
        message: `Excel ${files.length} ไฟล์ รวม ${importedItems.length} ภาระสอน\nเพิ่มภาระสอน ${createCount} / อัปเดต ${updateCount}\nสร้างรายวิชาใหม่ ${newSubjects.length} วิชา\n${sampleNames}${moreNames}\n\nวิชาใหม่ตั้งต้นเป็นคาบเดี่ยวและเวลาใดก็ได้ และจะถูกทำเครื่องหมาย "รอตรวจ" กรุณาเปิดข้อมูลพื้นฐาน > รายวิชา ตรวจ Block และช่วงเวลา แล้วบันทึกวิชาใหม่แต่ละรายการก่อนจัด Auto`,
        action: async () => {
          setConfirmData(prev => ({ ...prev, isLoading: true }));
          const names = Object.fromEntries(subjectRows.map(row => [row.subjectId, row.subjectName]));
          const ok = await saveTeachingLoadsWithSubjects(importedItems, newSubjects, names);
          if (ok) { setConfirmData(null); showToast(`นำเข้าสำเร็จ ${importedItems.length} ภาระสอน และเพิ่มวิชา ${newSubjects.length} วิชา`); }
          else setConfirmData(prev => ({ ...prev, isLoading: false }));
        }
      });
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
          <button onClick={() => { setForm({ id: '', teacherId: '', subjectId: '', classroomId: '', periods: 1 }); resetNewSubject(); setOriginalLoad(null); setIsEditing(false); setIsModalOpen(true); }} className="flex items-center gap-2 px-4 py-2 bg-gradient-to-r from-[#081a39] to-[#122b5e] text-white rounded-xl shadow-md font-bold"><Plus size={18} className="text-[#d4af37]"/> เพิ่มภาระสอน</button>
        </div>
      </div>
      <div className="bg-white rounded-2xl shadow-sm border border-slate-100 overflow-hidden flex flex-col">
         <div className="px-4 py-4 border-b border-slate-100 flex flex-wrap gap-4 items-end justify-between">
           <div className="flex flex-wrap gap-3 items-end">
             <label className="text-sm font-bold text-[#081a39]">ดูภาระสอน
               <select value={selectedTeacherId === 'all' ? 'all' : 'one'} onChange={e=>setSelectedTeacherId(e.target.value === 'all' ? 'all' : (teachers[0]?.id || 'all'))} className="block mt-1 p-2.5 border border-slate-200 rounded-lg min-w-[150px] bg-white">
                 <option value="all">ครูทั้งหมด</option><option value="one">รายบุคคล</option>
               </select>
             </label>
             {selectedTeacherId !== 'all' && <label className="text-sm font-bold text-[#081a39]">เลือกครู
               <select value={selectedTeacherId} onChange={e=>setSelectedTeacherId(e.target.value)} className="block mt-1 p-2.5 border border-slate-200 rounded-lg min-w-[220px] bg-white">
                 {teachers.map(t=><option key={t.id} value={t.id}>{t.name}</option>)}
               </select>
             </label>}
           </div>
           <p className="text-sm text-slate-600">แสดง <strong className="text-[#081a39]">{displayedLoadRows.length}</strong> รายการ · รวม <strong className="text-[#081a39]">{displayedPeriods}</strong> คาบ/สัปดาห์ (เฉพาะรายวิชา ไม่รวมกิจกรรมและ PLC)</p>
         </div>
         <div className="overflow-x-auto custom-scrollbar">
           <table className="w-full text-left border-collapse">
              <thead className="bg-[#081a39] text-white border-b-4 border-[#d4af37]">
                <tr>
                  <th className="p-4 font-bold cursor-pointer select-none" onClick={() => sortLoads('teacherName')}>ครูผู้สอน <SortIcon column="teacherName" sortConfig={loadSortConfig}/></th>
                  <th className="p-4 font-bold cursor-pointer select-none" onClick={() => sortLoads('subjectName')}>รายวิชา <SortIcon column="subjectName" sortConfig={loadSortConfig}/></th>
                  <th className="p-4 font-bold text-center cursor-pointer select-none" onClick={() => sortLoads('classroomSortKey')}>ห้อง <SortIcon column="classroomSortKey" sortConfig={loadSortConfig}/></th>
                  <th className="p-4 font-bold text-center cursor-pointer select-none" onClick={() => sortLoads('periods')}>คาบ/สัปดาห์ <SortIcon column="periods" sortConfig={loadSortConfig}/></th>
                  <th className="p-4 font-bold text-center w-28">จัดการ</th>
                </tr>
              </thead>
              <tbody>
                {displayedLoadRows.map(({ load: l }) => (
                  <tr key={l.id} className="border-b border-slate-50 hover:bg-[#081a39]/5">
                    <td className="p-4 font-medium text-[#081a39]">{getTeacher(l.teacherId).name}</td><td className="p-4 text-slate-600"><span className="text-xs font-mono bg-slate-100 text-slate-500 px-2 py-1 rounded mr-2">{l.subjectId}</span>{getSubject(l.subjectId).name}</td><td className="p-4 text-center font-bold text-[#d4af37]">{getClassroom(l.classroomId).name}</td><td className="p-4 text-center font-medium">{l.periods}</td>
                    <td className="p-4 text-center"><div className="flex items-center justify-center gap-2"><button onClick={() => { setForm(l); resetNewSubject(); setOriginalLoad({...l}); setIsEditing(true); setIsModalOpen(true); }} className="p-2 text-slate-400 hover:text-[#081a39] rounded-lg"><Edit size={18}/></button><button onClick={() => handleRequestDelete('load', l)} className="p-2 text-slate-400 hover:text-rose-500 rounded-lg"><Trash2 size={18}/></button></div></td>
                  </tr>
                ))}
                {displayedLoadRows.length === 0 && <tr><td colSpan={5} className="p-8 text-center text-slate-500">ไม่พบภาระสอนของครูที่เลือกในภาคเรียนนี้</td></tr>}
              </tbody>
           </table>
         </div>
      </div>
      <Modal isOpen={isModalOpen} onClose={() => setIsModalOpen(false)} title={isEditing ? "แก้ไขภาระการสอน" : "เพิ่มภาระการสอน"}>
        <div className="space-y-4">
          <div><label className="block text-sm font-bold text-[#081a39] mb-1">ครูผู้สอน</label><select value={form.teacherId} onChange={e => setForm({...form, teacherId: e.target.value})} className="w-full p-3 border border-slate-200 rounded-xl bg-slate-50"><option value="">-- เลือกครูผู้สอน --</option>{teachers.map(t => <option key={t.id} value={t.id}>{t.name}</option>)}</select></div>
          <div><label className="block text-sm font-bold text-[#081a39] mb-1">รายวิชา</label>
            <select value={createNewSubject ? '__CREATE_NEW_SUBJECT__' : form.subjectId} onChange={e => {
              if (e.target.value === '__CREATE_NEW_SUBJECT__') { setCreateNewSubject(true); setForm(prev => ({ ...prev, subjectId: '' })); }
              else { setCreateNewSubject(false); setForm(prev => ({ ...prev, subjectId: e.target.value })); }
            }} className="w-full p-3 border border-slate-200 rounded-xl bg-slate-50">
              <option value="">-- เลือกรายวิชา --</option>
              {activeSubjects.map(s => <option key={s.id} value={s.id}>{s.name} ({s.id})</option>)}
              <option value="__CREATE_NEW_SUBJECT__">+ เพิ่มรายวิชาใหม่พร้อมภาระสอน</option>
            </select>
            {createNewSubject && <div className="mt-3 space-y-3 bg-amber-50 border border-amber-200 rounded-xl p-3">
              <p className="text-sm font-bold text-amber-900">สร้างรายวิชาอัตโนมัติพร้อมบันทึกภาระสอน</p>
              <div><label className="block text-xs font-bold mb-1">รหัสวิชาใหม่</label><input value={newSubject.id} onChange={e=>setNewSubject(prev=>({...prev,id:e.target.value}))} className="w-full p-2 border rounded-lg" placeholder="เช่น ค21101" /></div>
              <div><label className="block text-xs font-bold mb-1">ชื่อวิชา</label><input value={newSubject.name} onChange={e=>setNewSubject(prev=>({...prev,name:e.target.value}))} className="w-full p-2 border rounded-lg" placeholder="เช่น คณิตศาสตร์พื้นฐาน" /></div>
              <div><label className="block text-xs font-bold mb-1">ชื่อย่อ (ไม่กรอกจะใช้ชื่อวิชา)</label><input value={newSubject.abbr} onChange={e=>setNewSubject(prev=>({...prev,abbr:e.target.value}))} className="w-full p-2 border rounded-lg" /></div>
              <div className="grid grid-cols-2 gap-2">
                <div><label className="block text-xs font-bold mb-1">คาบต่อเนื่อง</label><select value={newSubject.periodBlock} onChange={e=>setNewSubject(prev=>({...prev,periodBlock:Number(e.target.value)}))} className="w-full p-2 border rounded-lg"><option value={1}>คาบเดี่ยว</option><option value={2}>คาบคู่ (Block 2)</option></select></div>
                <div><label className="block text-xs font-bold mb-1">ช่วงเวลาที่เหมาะสม</label><select value={newSubject.preferredTime} onChange={e=>setNewSubject(prev=>({...prev,preferredTime:e.target.value}))} className="w-full p-2 border rounded-lg"><option value="any">เวลาใดก็ได้</option><option value="morning">ช่วงเช้า</option><option value="afternoon">ช่วงบ่าย</option></select></div>
              </div>
              <p className="text-xs text-amber-800">จำนวนคาบรายวิชาเริ่มต้นจะเท่าภาระสอนรายการนี้ สามารถแก้ที่เมนูรายวิชาภายหลัง</p>
            </div>}
          </div>
          <div><label className="block text-sm font-bold text-[#081a39] mb-1">ห้องเรียน</label><select value={form.classroomId} onChange={e => setForm({...form, classroomId: e.target.value})} className="w-full p-3 border border-slate-200 rounded-xl bg-slate-50"><option value="">-- เลือกห้องเรียน --</option>{classrooms.map(c => <option key={c.id} value={c.id}>{c.name}</option>)}</select></div>
          <div><label className="block text-sm font-bold text-[#081a39] mb-1">จำนวนคาบ/สัปดาห์</label><input type="number" min="1" value={form.periods} onChange={e => setForm({...form, periods: parseInt(e.target.value)})} className="w-full p-3 border border-slate-200 rounded-xl bg-slate-50" /></div>
          <button onClick={handleSave} disabled={isSaving} className="w-full mt-2 bg-[#081a39] text-white py-3 rounded-xl font-bold disabled:opacity-50">{isSaving ? "กำลังบันทึก..." : (createNewSubject ? "สร้างวิชาและบันทึกภาระสอน" : "บันทึกภาระสอน")}</button>
        </div>
      </Modal>
    </div>
  );
};

// STAGE 25: Plan fixed-activity bulk deletion without changing unrelated records.
// Child activities cannot outlive their selected parent. Delete only children of
// selected parents, together with parents, in one atomic Firestore operation.
const planFixedActivityBulkDeletion = (selectedIds, fixedItems = [], groupItems = []) => {
  const ids = [...new Set((Array.isArray(selectedIds) ? selectedIds : []).map(String))];
  if (!ids.length || ids.some(id => !id || id.includes('/'))) throw new Error('FIXED_SELECTION_EMPTY');
  const byId = new Map(fixedItems.map(item => [String(item.id), item]));
  if (ids.some(id => !byId.has(id))) throw new Error('FIXED_SELECTION_CHANGED');
  const idSet = new Set(ids);
  return { ids, selected: ids.map(id => byId.get(id)),
    linkedGroups: groupItems.filter(group => idSet.has(String(group.fixedId))) };
};

const ConstraintsView = ({ activityGroups, unavailabilities, classroomUnavailabilities, fixedSchedules, schedules, teachers, classrooms, activeSubjects, periods, dbAction, blockClassroomDay, blockClassroomDays, showToast, handleRequestDelete, bulkDeleteFixedSchedules, schoolSettings, getTeacher, getClassroom, getSubject, customRules, rulesReady, plcSessions = [], mode = 'time' }) => {
  // Stage 13: use the existing editor either for time restrictions or fixed activities.
  const [activeSubTab, setActiveSubTab] = useState(mode === 'fixed' || mode === 'development' ? 'fixed' : 'teacher');
  const isDevEditor = mode === 'development';
  const [tForm, setTForm] = useState({ teacherId: '', day: 'จันทร์', periodId: '1' });
  const [cForm, setCForm] = useState({ classroomId: '', day: 'จันทร์', periodId: '1', reason: '' });
  const [fForm, setFForm] = useState({ subjectId: '', classroomId: '', teacherId: '', teacherIds: [], day: 'จันทร์', periodId: '1',includedClassroomIds:[], activityName:'' });
  const [fixedRoomSubset,setFixedRoomSubset] = useState(false);
  const [activityGrade,setActivityGrade] = useState('');
  const [editingFixedTeachers,setEditingFixedTeachers] = useState('');
  const [fixedTeacherDraft,setFixedTeacherDraft] = useState([]);
  const [bulkClosedRoomIds,setBulkClosedRoomIds] = useState([]);
  const [selectedFixedIds,setSelectedFixedIds] = useState([]);
  const selectedFixedSet = useMemo(() => new Set(selectedFixedIds), [selectedFixedIds]);
  const visibleFixedSchedules = useMemo(() => fixedSchedules.filter(f =>
    isDevEditor ? isDevelopmentActivity(f) : !isDevelopmentActivity(f)), [fixedSchedules,isDevEditor]);
  const allFixedSelected = visibleFixedSchedules.length > 0 && visibleFixedSchedules.every(item => selectedFixedSet.has(String(item.id)));
  const currentFixedTerm = `${schoolSettings?.academicYear || ''}-${schoolSettings?.semester || ''}`;
  useEffect(() => { setSelectedFixedIds([]); }, [currentFixedTerm]);
  useEffect(() => {
    const available = new Set(visibleFixedSchedules.map(item => String(item.id)));
    setSelectedFixedIds(previous => previous.filter(id => available.has(id)));
  }, [visibleFixedSchedules]);
  const toggleFixedSelected = (id, checked) => setSelectedFixedIds(previous => checked
    ? [...new Set([...previous, String(id)])]
    : previous.filter(itemId => itemId !== String(id)));
  const toggleAllFixedSelected = checked => setSelectedFixedIds(checked ? visibleFixedSchedules.map(item => String(item.id)) : []);


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
      if (!rulesReady) return showToast('ยังโหลดเงื่อนไขไม่สำเร็จ', 'error');
      if (!fForm.classroomId || (isDevEditor ? !String(fForm.activityName||'').trim() : !fForm.subjectId)) return showToast(isDevEditor ? 'กรุณากรอกชื่อกิจกรรมและเลือกห้องเรียน' : 'กรุณาเลือกวิชาและห้อง', 'error');
      const maps = buildConstraintMaps(schedules, unavailabilities, classroomUnavailabilities, fixedSchedules, periods, customRules, activityGroups, classrooms, plcSessions, teachers);
      const chosenTeachers = fixedActivityTeacherIds(fForm);
      const existingOccupancy = buildTeacherOccupancy(schedules, fixedSchedules, activityGroups);
      const fifthTeachers = [];
      for (const id of chosenTeachers) {
        if (!teachers.some(t => t.id === id)) return showToast('ไม่พบครูที่เลือกในระบบ', 'error');
        const existingToday=(maps.teacherDayPeriods.get(teacherDayKey(id,fForm.day)) || new Set()).size;
        if (existingToday >= SCHEDULER_CONFIG.absoluteMaxTeacherPeriodsPerDay)
          return showToast('ครูที่เลือกมีภาระครบ 5 คาบในวันนี้แล้ว', 'error');
        if (existingToday >= SCHEDULER_CONFIG.maxTeacherPeriodsPerDay) fifthTeachers.push(id);
      }
      if (fifthTeachers.length) {
        const names = fifthTeachers.map(id=>teachers.find(t=>t.id===id)?.name || id).join(', ');
        if (!window.confirm(`กิจกรรมนี้จะทำให้ครูต่อไปนี้มีภาระ 5 คาบในวัน${fForm.day}: ${names}\nยืนยันว่าเป็นกรณีจำเป็นจริงหรือไม่?`)) return;
        maps.maxTeacherPeriodsPerDay=SCHEDULER_CONFIG.absoluteMaxTeacherPeriodsPerDay;
      }
      const savedForm={...fForm,teacherId:'',teacherIds:chosenTeachers,
        activityType:isDevEditor ? 'development' : 'general',
        activityName:isDevEditor ? String(fForm.activityName).trim() : ''};
      if(isDevEditor) savedForm.subjectId='';
      if(fForm.classroomId !== ALL_CLASSROOMS_ID || !fixedRoomSubset) delete savedForm.includedClassroomIds;
      if (savedForm.classroomId === ALL_CLASSROOMS_ID && activityGrade && fixedRoomSubset) savedForm.scopeLabel = activityGrade;
      if(fixedRoomSubset && fForm.classroomId===ALL_CLASSROOMS_ID && !fForm.includedClassroomIds.length)
        return showToast('กรุณาเลือกห้องเรียนอย่างน้อยหนึ่งห้อง','error');
      const targetRooms = classrooms.filter(c => fixedAppliesToClassroom(savedForm,c.id));
      if (targetRooms.length === 0) return showToast("ยังไม่มีห้องเรียนให้ล็อกกิจกรรม", "error");
      const selectedPeriod = periods.find(p => p.id === fForm.periodId);
      if (!selectedPeriod || selectedPeriod.isBreak) return showToast("ไม่สามารถล็อกกิจกรรมในคาบพักหรือคาบที่ไม่มีอยู่", "error");
      // The fixed activity occupies each student room once; all chosen teachers
      // are checked independently without multiplying their load by room count.
      const reasons = { "ROOM_UNAVAILABLE": "ห้องไม่ว่างในเวลานี้", "TEACHER_UNAVAILABLE": "ครูไม่สะดวกในเวลานี้", "FIXED_SCHEDULE_CONFLICT": "ห้องนี้มีกิจกรรมล็อกแล้ว", "ROOM_CONFLICT": "ห้องมีตารางอยู่แล้ว", "TEACHER_CONFLICT": "ครูมีตารางอยู่แล้ว", "TEACHER_DAILY_LIMIT": "ครูสอนเกินเพดานต่อวันที่อนุมัติไม่ได้", "TEACHER_CONSECUTIVE_LIMIT": "การลงด้วยมือจะทำให้สอนรายวิชาติดกันเกิน 3 คาบ กรุณาเลือกคาบอื่น", "CUSTOM_RULE_BLOCKED": "ขัดกับเงื่อนไขห้ามจัดที่ตั้งไว้", "SUBJECT_DAILY_LIMIT": "วิชานี้เต็มจำนวนคาบสูงสุดต่อวันของห้องแล้ว" };
      for (const room of targetRooms) {
        const valid = canPlaceSchedule('', room.id, fForm.day, fForm.periodId, maps, false, fForm.subjectId);
        if (!valid.ok) return showToast(`${room.name}: ${reasons[valid.reason] || 'ล็อกไม่ได้'}`, 'error');
      }
      for (const id of chosenTeachers) {
        const valid = canPlaceSchedule(id, targetRooms[0].id, fForm.day, fForm.periodId, maps, false, fForm.subjectId, false, true);
        if (!valid.ok) return showToast(`${teachers.find(t=>t.id===id)?.name}: ${reasons[valid.reason] || 'ครูไม่ว่าง'}`, 'error');
        // Check the same rules applied to all teacher activities before saving.
        if (findTeacherTeachingLimitConflicts([...existingOccupancy,{
          teacherId:id,day:fForm.day,periodId:fForm.periodId,__isFixedActivity:true
        }],periods,teachers).length) return showToast('ครูที่เลือกขัดกับข้อบังคับการกระจายคาบ', 'error');
      }
      const id = `F_${Date.now()}`;
      const ok = await dbAction('fixedSchedules', id, { ...savedForm, id });
      if (ok) { showToast(isDevEditor ? 'บันทึกกิจกรรมพัฒนาผู้เรียนและล็อกวัน–คาบแล้ว' : 'บันทึกล็อกตารางสำเร็จ'); if(isDevEditor) setFForm(previous=>({...previous,activityName:''})); }
    }
  };

  const changeFixedActivityType = async (item, nextType) => {
    const name = fixedActivityDisplayName(item, getSubject);
    const originalSubjectId = item.originalSubjectId || item.subjectId || '';
    if (nextType === 'general' && !activeSubjects.some(subject => subject.id === originalSubjectId))
      return showToast('ไม่พบรหัสรายวิชาเดิม จึงยังเปลี่ยนเป็นกิจกรรมทั่วไปไม่ได้ กรุณาใช้ประเภทกิจกรรมพัฒนาผู้เรียนต่อ', 'error');
    if (!window.confirm(`เปลี่ยนประเภท “${name}” เป็น ${nextType === 'development' ? 'กิจกรรมพัฒนาผู้เรียน' : 'กิจกรรมล็อกตายตัวทั่วไป'} หรือไม่?\nวัน คาบ ครู และห้องเรียนที่ล็อกไว้จะไม่เปลี่ยน`)) return;
    // Keep original subject reference for reversibility, but never let a learning
    // activity consume a regular teaching subject code or daily subject quota.
    const next = {...item,activityType:nextType,originalSubjectId,
      subjectId:nextType === 'development' ? '' : originalSubjectId,
      activityName:nextType === 'development' ? name : (item.activityName || '')};
    const ok = await dbAction('fixedSchedules',item.id,next);
    if(ok) {setSelectedFixedIds([]); showToast('เปลี่ยนประเภทกิจกรรมสำเร็จ โดยไม่ย้ายคาบหรือแก้ข้อมูลครู');}
  };

  const renameDevelopmentActivity = async item => {
    const raw = window.prompt('แก้ชื่อกิจกรรมพัฒนาผู้เรียน (วันและคาบที่ล็อกไว้จะไม่เปลี่ยน)',fixedActivityDisplayName(item,getSubject));
    if (raw === null) return;
    const activityName = raw.trim();
    if (!activityName || activityName.length > 120) return showToast('ชื่อกิจกรรมต้องมีความยาว 1–120 ตัวอักษร', 'error');
    if (activityName === item.activityName) return;
    const ok = await dbAction('fixedSchedules',item.id,{...item,activityName});
    if(ok) showToast('แก้ชื่อกิจกรรมสำเร็จ โดยคงวัน คาบ และครูเดิมไว้');
  };

  const saveExistingFixedTeachers = async item => {
    if (!rulesReady) return showToast('รอโหลดข้อมูลเงื่อนไขก่อน','error');
    const chosen = [...new Set(fixedTeacherDraft)];
    if (chosen.some(id=>!teachers.some(t=>t.id===id))) return showToast('พบรหัสครูที่ไม่มีในระบบ','error');
    const rest = fixedSchedules.filter(f=>f.id!==item.id);
    const maps = buildConstraintMaps(schedules,unavailabilities,classroomUnavailabilities,
      rest,periods,customRules,activityGroups,classrooms,plcSessions,teachers);
    const targets = classrooms.filter(c=>fixedAppliesToClassroom(item,c.id));
    if (!targets.length) return showToast('กิจกรรมไม่มีห้องที่เข้าร่วม','error');
    const existing = buildTeacherOccupancy(schedules,rest,activityGroups);
    const fifth = [];
    for (const teacherId of chosen) {
      const existingToday = (maps.teacherDayPeriods.get(teacherDayKey(teacherId,item.day)) || new Set()).size;
      if (existingToday >= SCHEDULER_CONFIG.absoluteMaxTeacherPeriodsPerDay)
        return showToast('ครูบางคนมีภาระครบ 5 คาบแล้ว','error');
      if (existingToday >= SCHEDULER_CONFIG.maxTeacherPeriodsPerDay) fifth.push(teacherId);
    }
    if (fifth.length) {
      if (!window.confirm(`ครู ${fifth.map(id=>teachers.find(t=>t.id===id)?.name || id).join(', ')} จะมีกิจกรรมรวม 5 คาบในวัน${item.day}\nยืนยันข้อยกเว้นหรือไม่?`)) return;
      maps.maxTeacherPeriodsPerDay=SCHEDULER_CONFIG.absoluteMaxTeacherPeriodsPerDay;
    }
    for (const teacherId of chosen) {
      const valid = canPlaceSchedule(teacherId,targets[0].id,item.day,item.periodId,maps,false,item.subjectId,false,true);
      if (!valid.ok) return showToast(`${teachers.find(t=>t.id===teacherId)?.name}: มีคาบชนหรือไม่ผ่านข้อบังคับ (${valid.reason})`,'error');
      const conflicts = findTeacherTeachingLimitConflicts([...existing,{teacherId,day:item.day,periodId:item.periodId,__isFixedActivity:true}],periods,teachers);
      if (conflicts.length) return showToast(conflicts[0],'error');
    }
    const ok = await dbAction('fixedSchedules',item.id,{...item,teacherId:chosen.includes(item.teacherId) ? item.teacherId : '',teacherIds:chosen});
    if (ok) {setEditingFixedTeachers('');setFixedTeacherDraft([]);showToast('เพิ่มหรือปรับครูผู้รับผิดชอบสำเร็จ');}
  };

  const SubTabBtn = ({ id, label, icon: Icon }) => (
    <button onClick={() => setActiveSubTab(id)} className={`flex items-center gap-2 px-4 py-2 rounded-lg font-bold text-sm transition-all ${activeSubTab === id ? 'bg-[#081a39] text-[#d4af37] shadow-md' : 'bg-white text-slate-500 hover:bg-slate-100 border border-slate-200'}`}><Icon size={16} /> {label}</button>
  );

  return (
    <div className="animate-in fade-in space-y-6">
      {mode === 'time' && <div className="flex gap-3 flex-wrap bg-slate-50 p-2 rounded-xl w-fit border border-slate-200"><SubTabBtn id="teacher" label="เวลาที่ครูไม่สะดวก" icon={Users} /><SubTabBtn id="room" label="เวลาที่ห้องไม่ว่าง" icon={Home} /></div>}
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
              <button onClick={() => handleAdd('room')} className="w-full bg-[#081a39] text-white py-3 rounded-xl font-bold">บันทึกคาบเดียว</button>
              <button onClick={() => blockClassroomDay(cForm.classroomId,cForm.day,cForm.reason)} className="w-full bg-amber-600 text-white py-3 rounded-xl font-bold">ปิดทั้งวัน (ห้องที่เลือก)</button>
              <details className="border border-amber-200 p-3 rounded-xl space-y-2">
                <summary className="font-bold text-sm cursor-pointer">ปิดทั้งวันหลายห้องพร้อมกัน (เช่น ห้องทวิศึกษา)</summary>
                <div className="flex gap-3 text-xs py-2"><button type="button" className="underline" onClick={()=>setBulkClosedRoomIds(classrooms.map(x=>x.id))}>เลือกทั้งหมด</button>
                  <button type="button" className="underline" onClick={()=>setBulkClosedRoomIds([])}>ล้างการเลือก</button></div>
                <div className="grid grid-cols-2 gap-2 max-h-40 overflow-y-auto">{classrooms.map(c=><label key={c.id} className="text-xs flex gap-2 items-center">
                  <input type="checkbox" checked={bulkClosedRoomIds.includes(c.id)} onChange={e=>setBulkClosedRoomIds(old=>e.target.checked?[...old,c.id]:old.filter(x=>x!==c.id))}/>{c.name}</label>)}</div>
                <button type="button" disabled={!bulkClosedRoomIds.length} onClick={()=>blockClassroomDays(bulkClosedRoomIds,cForm.day,cForm.reason)}
                  className="w-full bg-indigo-700 text-white py-2 rounded-lg font-bold disabled:opacity-40">ปิดวัน{cForm.day} ทั้งหมด {bulkClosedRoomIds.length} ห้อง</button>
              </details>
              <p className="text-xs text-slate-500">ปิดทั้งวันแบบครั้งเดียว โดยข้ามรายการซ้ำ และไม่อนุญาตหากมีตารางเรียน/กิจกรรมล็อกในวันนั้นแล้ว</p>
            </div>
          )}
          {activeSubTab === 'fixed' && (
            <div className="space-y-4">
              {isDevEditor ? <label className="block text-sm font-bold">ชื่อกิจกรรมพัฒนาผู้เรียน
                <input type="text" maxLength={120} placeholder="เช่น ชุมนุม / ลูกเสือ–เนตรนารี / เศรษฐกิจพอเพียง" value={fForm.activityName} onChange={e=>setFForm(old=>({...old,activityName:e.target.value}))} className="w-full mt-2 p-3 border rounded-xl bg-slate-50 font-normal" />
                <span className="block text-xs text-slate-500 mt-1 font-normal">กิจกรรมนี้ไม่ต้องสร้างเป็นรายวิชาก่อน กำหนดวัน–คาบเองและล็อกทันที</span>
              </label> : <select value={fForm.subjectId} onChange={e => setFForm({...fForm, subjectId: e.target.value})} className="w-full p-3 border rounded-xl bg-slate-50"><option value="">-- เลือกวิชา / กิจกรรม --</option>{activeSubjects.map(s => <option key={s.id} value={s.id}>{s.name}</option>)}</select>}
              <select value={fForm.classroomId} onChange={e => {setActivityGrade('');setFixedRoomSubset(false);setFForm({...fForm, classroomId: e.target.value, includedClassroomIds:[]});}} className="w-full p-3 border rounded-xl bg-slate-50"><option value="">-- เลือกห้องเรียน --</option><option value={ALL_CLASSROOMS_ID}>ทุกห้องเรียน</option>{classrooms.map(c => <option key={c.id} value={c.id}>{c.name}</option>)}</select>
              {fForm.classroomId === ALL_CLASSROOMS_ID && <div className="border border-amber-200 rounded-xl p-3 space-y-2">
                <label className="text-xs font-bold block">เลือกกิจกรรมตามระดับชั้น (เช่น ลูกเสือ ม.1 ไม่จำกัดห้อง)</label>
                <select value={activityGrade} className="w-full p-2 rounded-lg border bg-white" onChange={e=>{
                  const grade=e.target.value;setActivityGrade(grade);
                  if(grade){setFixedRoomSubset(true);setFForm(old=>({...old,includedClassroomIds:classrooms.filter(c=>classroomGrade(c)===grade).map(c=>c.id)}));}
                  else {setFixedRoomSubset(false);setFForm(old=>({...old,includedClassroomIds:[]}));}
                }}><option value="">-- ทุกห้อง / เลือกห้องเอง --</option>
                  {[...new Set(classrooms.map(classroomGrade).filter(Boolean))].sort().map(grade=><option key={grade} value={grade}>{grade} ทุกห้อง</option>)}
                </select>
                <label className="text-sm flex items-center gap-2"><input type="checkbox" checked={fixedRoomSubset} onChange={e=>{setFixedRoomSubset(e.target.checked);setActivityGrade('');}}/> เลือกเฉพาะบางห้อง / ยกเว้นห้องที่ไม่ร่วมกิจกรรม</label>
                {fixedRoomSubset ? <>
                  <div className="flex gap-2 text-xs"><button type="button" className="underline" onClick={()=>setFForm({...fForm,includedClassroomIds:classrooms.map(c=>c.id)})}>เลือกทุกห้องก่อน</button><button type="button" className="underline" onClick={()=>setFForm({...fForm,includedClassroomIds:[]})}>ล้างการเลือก</button></div>
                  <div className="max-h-40 overflow-y-auto grid grid-cols-2 gap-1">{classrooms.map(c=><label key={c.id} className="text-xs flex gap-2 items-center"><input type="checkbox" checked={fForm.includedClassroomIds.includes(c.id)} onChange={e=>{setActivityGrade('');setFForm(old=>({...old,includedClassroomIds:e.target.checked?[...old.includedClassroomIds,c.id]:old.includedClassroomIds.filter(id=>id!==c.id)}));}}/>{c.name}</label>)}</div>
                  <p className="text-xs text-amber-700">กิจกรรมจะล็อกเฉพาะ {fForm.includedClassroomIds.length} ห้องที่เลือก</p>
                </> : <p className="text-xs text-amber-700">กิจกรรมนี้จะปรากฏในทุกห้อง รวมถึงห้องที่เพิ่มในภายหลัง</p>}
              </div>}
              <div className="border rounded-xl p-3 space-y-2">
                <div className="flex flex-wrap items-center justify-between gap-2"><strong className="text-sm">ครูผู้รับผิดชอบ ({fForm.teacherIds.length} คน)</strong>
                  <div className="flex gap-3 text-xs"><button type="button" className="underline text-indigo-700" onClick={()=>setFForm(old=>({...old,teacherIds:teachers.map(t=>t.id)}))}>เลือกครูทั้งหมด</button>
                    <button type="button" className="underline text-rose-700" onClick={()=>setFForm(old=>({...old,teacherIds:[]}))}>ล้างการเลือก</button></div></div>
                <div className="max-h-44 overflow-auto grid grid-cols-1 sm:grid-cols-2 gap-1">{teachers.map(t=><label key={t.id} className="text-xs flex items-center gap-2 p-1">
                  <input type="checkbox" checked={fForm.teacherIds.includes(t.id)} onChange={e=>setFForm(old=>({...old,teacherIds:e.target.checked?[...old.teacherIds,t.id]:old.teacherIds.filter(id=>id!==t.id)}))}/>{t.name}</label>)}</div>
                <p className="text-xs text-slate-500">กิจกรรมแสดงในตารางครูที่เลือกทุกคน โดยนับคาบกิจกรรมเพียงคนละหนึ่งครั้ง</p>
              </div>
              <div className="grid grid-cols-2 gap-3"><select value={fForm.day} onChange={e => setFForm({...fForm, day: e.target.value})} className="w-full p-3 border rounded-xl bg-slate-50">{DAYS.map(d => <option key={d} value={d}>{d}</option>)}</select><select value={fForm.periodId} onChange={e => setFForm({...fForm, periodId: e.target.value})} className="w-full p-3 border rounded-xl bg-slate-50">{periods.filter(p=>!p.isBreak).map(p => <option key={p.id} value={p.id}>{p.name}</option>)}</select></div>
              <button onClick={() => handleAdd('fixed')} className="w-full bg-[#081a39] text-white py-3 rounded-xl font-bold">{isDevEditor ? 'บันทึกและล็อกกิจกรรมพัฒนาผู้เรียน' : 'บันทึก'}</button>
            </div>
          )}
        </div>
        <div className="lg:col-span-2 bg-white rounded-2xl shadow-sm border border-slate-100 overflow-hidden flex flex-col min-h-[400px]">
           {activeSubTab === 'fixed' && <div className="no-print flex flex-wrap justify-between items-center gap-3 px-4 py-3 bg-slate-50 border-b border-slate-200">
             <div className="flex flex-wrap items-center gap-3 text-sm">
               <label className="flex items-center gap-2 font-semibold text-[#081a39] cursor-pointer">
                 <input type="checkbox" aria-label="เลือกกิจกรรมล็อกทั้งหมด" checked={allFixedSelected} onChange={e => toggleAllFixedSelected(e.target.checked)} disabled={!visibleFixedSchedules.length}/>
                 เลือกทั้งหมด ({visibleFixedSchedules.length})
               </label>
               <span className="text-slate-500">เลือกแล้ว {selectedFixedIds.length} รายการ</span>
               <button type="button" disabled={!selectedFixedIds.length} onClick={() => setSelectedFixedIds([])} className="text-indigo-700 underline disabled:text-slate-400 disabled:no-underline">ล้างการเลือก</button>
             </div>
             <button type="button" disabled={!selectedFixedIds.length || !rulesReady} onClick={() => bulkDeleteFixedSchedules(selectedFixedIds)} className="inline-flex items-center gap-2 px-3 py-2 bg-rose-600 hover:bg-rose-700 text-white text-sm font-bold rounded-lg disabled:opacity-40 disabled:cursor-not-allowed"><Trash2 size={16}/> ลบที่เลือก ({selectedFixedIds.length})</button>
           </div>}
           <div className="overflow-x-auto custom-scrollbar">
           <table className="w-full text-left border-collapse">
              <thead className="bg-[#081a39] text-white border-b-4 border-[#d4af37]"><tr>{activeSubTab === 'fixed' && <th className="p-4 text-center w-12"><input type="checkbox" aria-label="เลือกกิจกรรมล็อกทั้งหมดในตาราง" checked={allFixedSelected} disabled={!visibleFixedSchedules.length} onChange={e => toggleAllFixedSelected(e.target.checked)}/></th>}<th className="p-4 font-bold">ข้อมูลรายการ</th><th className="p-4 font-bold text-center">วัน</th><th className="p-4 font-bold text-center">คาบที่</th><th className="p-4 font-bold text-center w-20">ลบ</th></tr></thead>
              <tbody>
                {activeSubTab === 'teacher' && unavailabilities.map(u => (
                  <tr key={u.id} className="border-b border-slate-50 hover:bg-[#081a39]/5"><td className="p-4 font-medium text-[#081a39]">{getTeacher(u.teacherId).name}</td><td className="p-4 text-center">{u.day}</td><td className="p-4 text-center font-bold text-rose-500">{getPeriodDisplayName(periods, u.periodId)}</td><td className="p-4 text-center"><button onClick={() => handleRequestDelete('unavail', u)} className="p-2 text-slate-400 hover:text-rose-500"><Trash2 size={18}/></button></td></tr>
                ))}
                {activeSubTab === 'room' && classroomUnavailabilities.map(u => (
                  <tr key={u.id} className="border-b border-slate-50 hover:bg-[#081a39]/5"><td className="p-4 font-medium text-[#081a39]">{getClassroom(u.classroomId).name} <span className="text-xs text-slate-400 ml-2">{u.reason}</span></td><td className="p-4 text-center">{u.day}</td><td className="p-4 text-center font-bold text-rose-500">{getPeriodDisplayName(periods, u.periodId)}</td><td className="p-4 text-center"><button onClick={() => handleRequestDelete('cUnavail', u)} className="p-2 text-slate-400 hover:text-rose-500"><Trash2 size={18}/></button></td></tr>
                ))}
                {activeSubTab === 'fixed' && visibleFixedSchedules.map(f => (
                  <tr key={f.id} className="border-b border-slate-50 hover:bg-[#081a39]/5"><td className="p-4 text-center"><input type="checkbox" aria-label={`เลือกกิจกรรม ${fixedActivityDisplayName(f,getSubject)} วัน${f.day} คาบ ${getPeriodDisplayName(periods, f.periodId)}`} checked={selectedFixedSet.has(String(f.id))} onChange={e => toggleFixedSelected(f.id, e.target.checked)}/></td><td className="p-4 font-medium text-[#081a39]"><span className="text-amber-600 font-bold">{fixedActivityDisplayName(f,getSubject)}</span> <span className="text-xs text-slate-500 ml-2">ห้อง: {f.scopeLabel || getClassroom(f.classroomId).name} · ครู {fixedActivityTeacherIds(f).length} คน</span>
                    {isDevelopmentActivity(f) && <button type="button" className="ml-2 underline text-indigo-700 text-xs" onClick={()=>renameDevelopmentActivity(f)}>แก้ชื่อ</button>}
                    <select aria-label={`ประเภทกิจกรรม ${fixedActivityDisplayName(f,getSubject)}`} title="เปลี่ยนประเภทโดยไม่ลบหรือย้ายคาบ" className="ml-2 text-xs border rounded-md p-1 bg-white" value={isDevelopmentActivity(f)?'development':'general'} onChange={e=>changeFixedActivityType(f,e.target.value)}>
                      <option value="general">กิจกรรมล็อกตายตัวทั่วไป</option><option value="development">กิจกรรมพัฒนาผู้เรียน</option>
                    </select>
                    <button type="button" className="ml-2 underline text-indigo-700 text-xs" onClick={()=>{
                      setEditingFixedTeachers(editingFixedTeachers===f.id?'':f.id);
                      setFixedTeacherDraft(fixedActivityTeacherIds(f));
                    }}>กำหนดครู</button>
                    {editingFixedTeachers===f.id && <div className="mt-3 border border-indigo-200 bg-indigo-50 rounded-xl p-3 space-y-2">
                      <div className="flex gap-4 text-xs"><button type="button" onClick={()=>setFixedTeacherDraft(teachers.map(t=>t.id))} className="underline">เลือกทั้งหมด</button>
                        <button type="button" onClick={()=>setFixedTeacherDraft([])} className="underline">ล้างการเลือก</button></div>
                      <div className="grid grid-cols-1 md:grid-cols-2 max-h-44 overflow-y-auto gap-1">{teachers.map(t=><label key={t.id} className="text-xs flex items-center gap-2">
                        <input type="checkbox" checked={fixedTeacherDraft.includes(t.id)} onChange={e=>setFixedTeacherDraft(old=>e.target.checked?[...old,t.id]:old.filter(id=>id!==t.id))}/>{t.name}</label>)}</div>
                      <div className="flex gap-2"><button type="button" className="bg-indigo-800 text-white rounded-lg px-3 py-2 text-xs" onClick={()=>saveExistingFixedTeachers(f)}>บันทึกครู</button>
                      <button type="button" className="bg-white border rounded-lg px-3 py-2 text-xs" onClick={()=>setEditingFixedTeachers('')}>ยกเลิก</button></div>
                    </div>}</td><td className="p-4 text-center">{f.day}</td><td className="p-4 text-center font-bold text-rose-500">{getPeriodDisplayName(periods, f.periodId)}</td><td className="p-4 text-center"><button onClick={() => handleRequestDelete('fixed', f)} className="p-2 text-slate-400 hover:text-rose-500"><Trash2 size={18}/></button></td></tr>
                ))}
              </tbody>
           </table>
           </div>
        </div>
      </div>
    </div>
  );
};

// STAGE 1: ผู้ดูแลเพิ่มเงื่อนไขรายวัน/คาบ โดยไม่เปลี่ยนข้อบังคับครู 5/3
const CustomRulesView = ({ customRules, teachers, activeSubjects, classrooms, periods, schedules, fixedSchedules, dbAction, showToast, rulesReady, saveSubjectTimeWindows }) => {
  const [form, setForm] = useState({ targetType: 'teacher', targetId: '', day: 'จันทร์', periodId: '', mode: 'block' });
  const [saving, setSaving] = useState(false);
  const [dailySubjectId, setDailySubjectId] = useState('');
  const [dailyMax, setDailyMax] = useState('2');
  const [timeSubjectIds, setTimeSubjectIds] = useState([]);
  const [timeSearch, setTimeSearch] = useState('');
  const [timeWindow, setTimeWindow] = useState('morning');
  const [timeStrictness, setTimeStrictness] = useState('required');
  const targets = form.targetType === 'teacher' ? teachers : form.targetType === 'subject' ? activeSubjects : classrooms;
  const canSelectSubjectAllDay = form.targetType === 'subject' && ['block', 'prefer'].includes(form.mode);
  const listForRule = r => r.targetType === 'teacher' ? teachers : r.targetType === 'subject' ? activeSubjects : classrooms;
  const labelForRule = r => {
    const target = listForRule(r).find(x => String(x.id) === String(r.targetId));
    return r.targetType === 'subject' && target
      ? `${target.id} — ${target.name}`
      : target?.name || r.targetId;
  };
  const handleAdd = async () => {
    if (saving) return;
    if (!rulesReady) return showToast('ยังโหลดเงื่อนไขไม่สำเร็จ กรุณาตรวจการเชื่อมต่อ', 'error');
    if (!isValidCustomRule(form, teachers, activeSubjects, classrooms, periods)) return showToast('กรุณาระบุครู/วิชา/ห้อง วัน และคาบหรือทั้งวันที่ถูกต้อง', 'error');
    if (customRules.some(r => r.targetType === form.targetType && r.targetId === form.targetId && r.day === form.day && String(r.periodId) === String(form.periodId) && r.mode === form.mode))
      return showToast('มีเงื่อนไขนี้แล้ว', 'error');
    if (form.mode === 'block' && [...schedules, ...fixedSchedules].some(item => customRuleMatches(form, item)))
      return showToast(form.periodId === ALL_DAY_PERIOD_ID
        ? 'มีรายวิชานี้อยู่ในตารางวันดังกล่าวแล้ว กรุณาย้ายคาบเดิมออกก่อนกำหนดห้ามทั้งวัน'
        : 'มีตารางเดิมหรือกิจกรรมล็อกในช่วงนี้ กรุณาจัดการตารางเดิมก่อน', 'error');
    setSaving(true);
    try {
      const id = `RULE_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
      const ok = await dbAction('scheduleRules', id, { ...form, id });
      if (ok) { setForm(prev => ({ ...prev, targetId: '' })); showToast('เพิ่มเงื่อนไขสำเร็จ'); }
    } finally { setSaving(false); }
  };
  const handleAddTimeWindow = async () => {
    if (saving) return;
    if (!timeSubjectIds.length) return showToast('กรุณาเลือกอย่างน้อย 1 รายวิชา', 'error');
    setSaving(true);
    try {
      const ok = await saveSubjectTimeWindows(timeSubjectIds, timeWindow, timeStrictness);
      if (ok) setTimeSubjectIds([]);
    } finally { setSaving(false); }
  };
  const subjectsMatchingSearch = activeSubjects.filter(s =>
    `${s.id} ${s.name}`.toLowerCase().includes(timeSearch.trim().toLowerCase()));
  const subjectsWithTimeRules = new Set(customRules.filter(r => r.mode === 'timeWindow').map(r => String(r.targetId)));
  const selectableSubjects = subjectsMatchingSearch.filter(s => !subjectsWithTimeRules.has(String(s.id)));
  const toggleTimeSubject = (id, checked) => setTimeSubjectIds(prev =>
    checked ? [...new Set([...prev, id])] : prev.filter(x => x !== id));
  const handleAddDailyLimit = async () => {
    if (!rulesReady || saving) return showToast('ยังไม่พร้อมบันทึกเงื่อนไข', 'error');
    const rule = { targetType: 'subject', targetId: dailySubjectId, mode: 'dailyLimit', maxPerDay: Number(dailyMax) };
    if (!isValidCustomRule(rule, teachers, activeSubjects, classrooms, periods)) return showToast('เลือกวิชาและจำนวนคาบสูงสุดให้ถูกต้อง', 'error');
    if (customRules.some(x => x.mode === 'dailyLimit' && x.targetId === dailySubjectId)) return showToast('วิชานี้มีข้อจำกัดจำนวนคาบแล้ว กรุณาลบรายการเก่าก่อน', 'error');
    if (findSubjectDailyLimitConflicts([...schedules, ...fixedSchedules], classrooms, [rule]).length) return showToast('ตารางเดิมมีวิชานี้เกินค่าที่กำหนด กรุณาจัดการตารางเดิมก่อน', 'error');
    setSaving(true);
    try {
      const id = `DAILY_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
      const ok = await dbAction('scheduleRules', id, { ...rule, id });
      if (ok) { setDailySubjectId(''); showToast('เพิ่มเพดานคาบรายวิชาสำเร็จ'); }
    } finally { setSaving(false); }
  };
  return (
    <div className="space-y-6">
      <div className="bg-white rounded-2xl border border-slate-100 p-6 max-w-4xl">
        <h2 className="text-xl font-bold text-[#081a39] mb-2">เงื่อนไขการจัดตารางเพิ่มเติม</h2>
        {!rulesReady && <p role="alert" className="text-rose-700 mb-3">กำลังโหลดเงื่อนไขจาก Firebase หรือไม่มีสิทธิ์อ่านข้อมูล กรุณาตรวจการเชื่อมต่อก่อนแก้ไข</p>}
        <p className="text-sm text-slate-600 mb-4">กฎบังคับใช้กับการจัดอัตโนมัติและการเพิ่มคาบด้วยมือ ส่วนกฎแนะนำมีผลกับการให้คะแนนตอนจัดอัตโนมัติเท่านั้น ครูเป้าหมายรวมคาบสอน+กิจกรรม 4 คาบ/วัน (5 คาบเมื่อจำเป็นและยืนยัน) และไม่เกิน 2 คาบสอนรายวิชาติดกัน; กิจกรรมจัดติดกับวิชาได้โดยจะแจ้งเตือนเมื่อภาระต่อเนื่องหนัก</p>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
          <label className="text-sm font-bold">ประเภทข้อมูล<select className="w-full border rounded-xl p-3 mt-1" value={form.targetType} onChange={e => setForm(prev => ({ ...prev, targetType: e.target.value, targetId: '', periodId: prev.periodId === ALL_DAY_PERIOD_ID ? '' : prev.periodId }))}><option value="teacher">ครู</option><option value="subject">รายวิชา</option><option value="classroom">ห้องเรียน</option></select></label>
          <label className="text-sm font-bold">เลือกบุคคล/รายการ<select className="w-full border rounded-xl p-3 mt-1" value={form.targetId} onChange={e => setForm(prev => ({ ...prev, targetId: e.target.value }))}><option value="">-- เลือก --</option>{targets.map(x => <option key={x.id} value={x.id}>{x.name} ({x.id})</option>)}</select></label>
          <label className="text-sm font-bold">วัน<select className="w-full border rounded-xl p-3 mt-1" value={form.day} onChange={e => setForm(prev => ({ ...prev, day: e.target.value }))}>{DAYS.map(day => <option key={day} value={day}>{day}</option>)}</select></label>
          <label className="text-sm font-bold">คาบ<select className="w-full border rounded-xl p-3 mt-1" value={form.periodId} onChange={e => setForm(prev => ({ ...prev, periodId: e.target.value }))}><option value="">-- เลือกคาบ --</option>{canSelectSubjectAllDay && <option value={ALL_DAY_PERIOD_ID}>{form.mode === 'prefer' ? 'ทั้งวัน (เลือกคาบใดก็ได้)' : 'ทั้งวัน (ทุกคาบ)'}</option>}{periods.filter(p => !p.isBreak).map(p => <option key={p.id} value={p.id}>{p.name}</option>)}</select></label>
          <label className="text-sm font-bold sm:col-span-2">เงื่อนไข<select className="w-full border rounded-xl p-3 mt-1" value={form.mode} onChange={e => setForm(prev => ({ ...prev, mode: e.target.value, periodId: prev.targetType !== 'subject' && prev.periodId === ALL_DAY_PERIOD_ID ? '' : prev.periodId }))}><option value="block">ข้อบังคับ: ห้ามจัดในคาบ/วันที่เลือก</option><option value="prefer">ข้อแนะนำ: พยายามจัดในวัน/คาบที่เลือก</option></select></label>
        </div>
        {canSelectSubjectAllDay && <p className="text-sm text-slate-600 mt-3">{form.mode === 'block'
          ? 'ห้ามจัด + ทั้งวัน: ห้ามรายวิชานี้ทุกคาบในวันดังกล่าวสำหรับทุกห้องที่ใช้รหัสวิชานี้'
          : 'พยายามเลือก + ทั้งวัน: ให้ Auto ให้ความสำคัญกับวันดังกล่าว จะเป็นคาบใดก็ได้ หากจัดไม่ได้ยังเลือกวันอื่นได้ ไม่ใช่ข้อบังคับ'} ทั้งสองแบบยังเลือกเฉพาะคาบตามเดิมได้</p>}
        <button disabled={saving || !rulesReady || !form.targetId || !form.periodId} onClick={handleAdd} className="mt-4 px-6 py-3 rounded-xl bg-[#081a39] text-white font-bold disabled:opacity-50">{saving ? 'กำลังบันทึก...' : 'เพิ่มเงื่อนไข'}</button>
      </div>
      <div className="bg-white border rounded-2xl p-6 max-w-4xl">
        <h3 className="font-bold text-lg mb-2">กำหนดช่วงเวลาเรียนรายวิชา (เช้า / บ่าย)</h3>
        <p className="text-slate-600 text-sm mb-4">เลือกหลายรหัสวิชาพร้อมกันได้ เช่น คณิตศาสตร์พื้นฐานและคณิตศาสตร์เพิ่มเติม กฎจะใช้ทุกวันในภาคเรียนนี้ โดยอ้างอิงคาบพักเที่ยงที่ตั้งค่าไว้</p>
        <label className="block text-sm font-bold mb-2">ค้นหารายวิชา
          <input value={timeSearch} onChange={e => setTimeSearch(e.target.value)} placeholder="ค้นหารหัสวิชา หรือชื่อวิชา"
            className="w-full border rounded-xl p-3 mt-1 font-normal" />
        </label>
        <div className="border rounded-xl max-h-56 overflow-y-auto p-3 space-y-2 mb-3">
          {!subjectsMatchingSearch.length && <p className="text-slate-500 text-sm">ไม่พบรายวิชาในภาคเรียนนี้</p>}
          {subjectsMatchingSearch.map(s => <label key={s.id} className={`flex items-center gap-2 p-2 rounded-lg text-sm ${subjectsWithTimeRules.has(String(s.id)) ? 'text-slate-400 bg-slate-50' : 'hover:bg-slate-50'}`}>
            <input type="checkbox" disabled={saving || subjectsWithTimeRules.has(String(s.id))} checked={timeSubjectIds.includes(String(s.id))}
              onChange={e => toggleTimeSubject(String(s.id), e.target.checked)} />
            <span>{s.name} ({s.id}) {subjectsWithTimeRules.has(String(s.id)) ? '— มีกฎอยู่แล้ว' : ''}</span>
          </label>)}
        </div>
        <div className="flex flex-wrap justify-between items-center gap-2 mb-4">
          <span className="text-sm text-slate-600">เลือก {timeSubjectIds.length} วิชา</span>
          <div className="flex gap-3 text-sm">
            <button type="button" className="text-blue-700 underline" onClick={() => setTimeSubjectIds(prev => [...new Set([...prev, ...selectableSubjects.map(s => String(s.id))])])}>เลือกที่ค้นหาทั้งหมด</button>
            <button type="button" className="text-slate-600 underline" onClick={() => setTimeSubjectIds([])}>ล้างที่เลือก</button>
          </div>
        </div>
        <div className="grid sm:grid-cols-2 gap-3">
          <label className="text-sm font-bold">ช่วงเวลาที่ต้องการ
            <select value={timeWindow} onChange={e => setTimeWindow(e.target.value)} className="w-full border rounded-xl p-3 mt-1">
              <option value="morning">ช่วงเช้า (ก่อนพักเที่ยง)</option><option value="afternoon">ช่วงบ่าย (หลังพักเที่ยง)</option>
            </select>
          </label>
          <label className="text-sm font-bold">ระดับเงื่อนไข
            <select value={timeStrictness} onChange={e => setTimeStrictness(e.target.value)} className="w-full border rounded-xl p-3 mt-1">
              <option value="required">ข้อบังคับ — ห้ามจัดนอกช่วงนี้</option><option value="preferred">ข้อแนะนำ — พยายามจัดช่วงนี้</option>
            </select>
          </label>
        </div>
        <p className="text-xs text-slate-500 mt-3">ข้อบังคับจะตรวจทั้ง Auto เพิ่มด้วยมือ และย้ายคาบ หากคาบเดิมขัดกับกฎใหม่ ระบบจะไม่บันทึกกฎนี้ ส่วนข้อแนะนำใช้จัดอันดับตัวเลือก Auto เท่านั้น</p>
        <button disabled={saving || !rulesReady || !timeSubjectIds.length} onClick={handleAddTimeWindow}
          className="mt-4 px-6 py-3 rounded-xl bg-[#081a39] text-white font-bold disabled:opacity-50">
          {saving ? 'กำลังบันทึก...' : `บันทึกช่วงเวลา ${timeSubjectIds.length} รายวิชา`}
        </button>
      </div>
      <div className="bg-white border rounded-2xl p-6 max-w-4xl">
        <h3 className="font-bold text-lg mb-3">จำกัดจำนวนคาบวิชาเดียวกันต่อวัน (แต่ละห้องเรียน)</h3>
        <p className="text-slate-600 text-sm mb-3">เช่น ตั้งคณิตศาสตร์ไม่เกิน 2 คาบต่อวัน/ห้อง โดยวิชาแบบ Block 2 ยังสามารถจัดสองคาบติดกันได้</p>
        <div className="flex flex-wrap gap-3 items-end">
          <label className="text-sm font-bold flex-1 min-w-[200px]">รายวิชา<select className="w-full border rounded-xl p-3 mt-1" value={dailySubjectId} onChange={e=>setDailySubjectId(e.target.value)}><option value="">-- เลือกรายวิชา --</option>{activeSubjects.map(x=><option key={x.id} value={x.id}>{x.name} ({x.id})</option>)}</select></label>
          <label className="text-sm font-bold w-36">สูงสุด (คาบ/วัน)<input className="w-full border rounded-xl p-3 mt-1" type="number" min="1" max={periods.filter(p=>!p.isBreak).length} value={dailyMax} onChange={e=>setDailyMax(e.target.value)} /></label>
          <button disabled={saving || !rulesReady || !dailySubjectId} onClick={handleAddDailyLimit} className="px-5 py-3 bg-[#081a39] text-white rounded-xl font-bold disabled:opacity-50">เพิ่มเพดานคาบ</button>
        </div>
      </div>
      <div className="bg-white border rounded-2xl p-6 max-w-4xl">
        <h3 className="font-bold text-lg mb-3">รายการเงื่อนไข ({customRules.length})</h3>
        {!customRules.length && <p className="text-slate-500 text-sm">ยังไม่มีเงื่อนไขเพิ่มเติม</p>}
        <div className="space-y-2">{customRules.map(rule => <div key={rule.id} className="flex flex-wrap gap-2 items-center justify-between border rounded-xl p-3"><div className="text-sm"><span className="font-bold">{labelForRule(rule)}</span> · {rule.mode === 'dailyLimit' ? `ไม่เกิน ${rule.maxPerDay} คาบต่อวัน/ห้อง` : rule.mode === 'timeWindow' ? `ทุกวัน · ช่วง${rule.timeWindow === 'morning' ? 'เช้า' : 'บ่าย'}` : `วัน${rule.day} · ${rule.targetType === 'subject' && String(rule.periodId) === ALL_DAY_PERIOD_ID ? (rule.mode === 'prefer' ? 'ทั้งวัน (เลือกคาบใดก็ได้)' : 'ทั้งวัน (ทุกคาบ)') : getPeriodDisplayName(periods, rule.periodId)}`} <span className={rule.mode === 'block' || (rule.mode === 'timeWindow' && rule.strictness === 'required') ? 'text-rose-700' : 'text-emerald-700'}>({rule.mode === 'dailyLimit' ? 'ข้อบังคับจำนวนคาบ' : rule.mode === 'timeWindow' ? (rule.strictness === 'required' ? 'บังคับช่วงเวลา' : 'แนะนำช่วงเวลา') : rule.mode === 'block' ? 'ห้ามจัด' : 'ควรจัด'})</span></div><button className="text-rose-700 border border-rose-200 px-3 py-2 rounded-lg" onClick={async () => { if (window.confirm('ยืนยันลบเงื่อนไขนี้?')) await dbAction('scheduleRules', rule.id, null, 'delete'); }}>ลบ</button></div>)}</div>
      </div>
    </div>
  );
};

// STAGE 6: per-classroom activity labels under a single all-room fixed time slot.
const ActivityGroupsView = ({activityGroups, fixedSchedules, classrooms, teachers, schedules, rulesReady,
  unavailabilities, periods, customRules, dbAction, showToast, getSubject}) => {
  const roots=fixedSchedules.filter(f=>f.classroomId===ALL_CLASSROOMS_ID);
  const [form,setForm]=useState({fixedId:'',classroomId:'',name:'',teacherId:''});
  const [busy,setBusy]=useState(false);
  const save=async()=>{
    if(!rulesReady) return showToast('รอโหลดเงื่อนไขและกิจกรรมกลุ่มก่อน','error');
    const root=roots.find(f=>f.id===form.fixedId);
    if(!root || !fixedAppliesToClassroom(root,form.classroomId) || !classrooms.some(c=>c.id===form.classroomId) || !form.name.trim())
      return showToast('เลือกกิจกรรมหลัก ห้องเรียน และกรอกชื่อกิจกรรมก่อน','error');
    if(activityGroups.some(x=>x.fixedId===form.fixedId&&x.classroomId===form.classroomId))
      return showToast('ห้องนี้มีกิจกรรมย่อยของช่วงเวลาที่เลือกอยู่แล้ว','error');
    if(form.teacherId){
      const groups=activityTeacherEntries(activityGroups,fixedSchedules);
      const other=[...schedules,...fixedSchedules,...groups].filter(x=>x.teacherId===form.teacherId);
      if(other.some(x=>x.day===root.day&&x.periodId===root.periodId))
        return showToast('ครูคนนี้มีคาบหรือกิจกรรมอื่นในเวลาเดียวกัน','error');
      if(unavailabilities.some(x=>x.teacherId===form.teacherId&&x.day===root.day&&x.periodId===root.periodId))
        return showToast('ครูไม่สะดวกในช่วงเวลานี้','error');
      if(blocksCustomRule(customRules,{teacherId:form.teacherId,day:root.day,periodId:root.periodId}))
        return showToast('ครูติดเงื่อนไขห้ามจัดในช่วงเวลานี้','error');
      const extended=[...buildTeacherOccupancy(schedules,fixedSchedules,activityGroups).filter(x=>x.teacherId===form.teacherId),
        {teacherId:form.teacherId,day:root.day,periodId:root.periodId,__activityTeacherOnly:true}];
      if(findTeacherTeachingLimitConflicts(extended,periods,teachers).length)
        return showToast('เกินข้อบังคับสูงสุด 5 คาบต่อวัน หรือมีคาบติดกันที่ไม่อนุญาต','error');
      const newFifth=findTeacherFifthPeriodDays(extended,periods,teachers);
      if(newFifth.length && !findTeacherFifthPeriodDays(other,periods,teachers).length &&
         !window.confirm(`ครูคนนี้จะมีกิจกรรมรวม 5 คาบในวัน${root.day}\nยืนยันข้อยกเว้นที่จำเป็นจริงหรือไม่?`)) return;
    }
    setBusy(true);
    try {
      const id=`AG_${Date.now()}_${Math.random().toString(36).slice(2,8)}`;
      const ok=await dbAction('activityGroups',id,{...form,id,name:form.name.trim()});
      if(ok){setForm(old=>({...old,name:'',teacherId:''}));showToast('บันทึกกิจกรรมกลุ่มสำเร็จ');}
    } finally {setBusy(false);}
  };
  return <div className="space-y-5">
    <div className="bg-white rounded-2xl p-6 border space-y-4">
      <h2 className="text-xl font-bold">กิจกรรมพัฒนาผู้เรียนแยกตามห้อง</h2>
      <p className="text-sm text-slate-600">สร้างกิจกรรมล็อก “ทุกห้องเรียน” ก่อน แล้วกำหนดชื่อกิจกรรม/ครูให้แต่ละห้องได้โดยไม่เพิ่มคาบล็อกซ้ำ</p>
      <select className="w-full border rounded-xl p-3" value={form.fixedId} onChange={e=>setForm({...form,fixedId:e.target.value})}><option value="">-- เลือกกิจกรรมหลักที่ล็อกทุกห้อง --</option>{roots.map(f=><option key={f.id} value={f.id}>{f.day} คาบ {f.periodId} · {fixedActivityDisplayName(f,getSubject)}</option>)}</select>
      <select className="w-full border rounded-xl p-3" value={form.classroomId} onChange={e=>setForm({...form,classroomId:e.target.value})}><option value="">-- เลือกห้องเรียน --</option>{classrooms.map(c=><option key={c.id} value={c.id}>{c.name}</option>)}</select>
      <input className="w-full border rounded-xl p-3" placeholder="ชื่อกิจกรรม เช่น ชุมนุมดนตรี" value={form.name} onChange={e=>setForm({...form,name:e.target.value})}/>
      <select className="w-full border rounded-xl p-3" value={form.teacherId} onChange={e=>setForm({...form,teacherId:e.target.value})}><option value="">-- ครูผู้รับผิดชอบ (ไม่ระบุได้) --</option>{teachers.map(t=><option key={t.id} value={t.id}>{t.name}</option>)}</select>
      <button onClick={save} disabled={busy||!roots.length||!rulesReady} className="px-6 py-3 rounded-xl bg-[#081a39] text-white font-bold disabled:opacity-50">บันทึกกิจกรรมกลุ่ม</button>
    </div>
    <div className="bg-white rounded-2xl p-5 border space-y-3"><h3 className="font-bold">รายการกิจกรรมย่อย ({activityGroups.length})</h3>
      {activityGroups.map(g=><div key={g.id} className="flex justify-between gap-3 items-center border-b pb-2">
        <div><strong>{classrooms.find(c=>c.id===g.classroomId)?.name||g.classroomId}: {g.name}</strong><p className="text-xs text-slate-500">{teachers.find(t=>t.id===g.teacherId)?.name||'ไม่ได้ระบุครู'}</p></div>
        <button onClick={async()=>{if(window.confirm('ยืนยันลบกิจกรรมกลุ่มนี้?')) await dbAction('activityGroups',g.id,null,'delete');}} className="text-rose-600 border rounded-lg p-2"><Trash2 size={16}/></button>
      </div>)}
    </div>
  </div>;
};


// Stage 21: term-specific PLC editor. One reservation per subject area per week.
const PLCView = ({teachers,plcSessions,periods,schedules,fixedSchedules,activityGroups,
  unavailabilities,db,appId,getTermKey,schoolSettings,showToast,rulesReady}) => {
  const [area,setArea]=useState(PLC_AREAS[0]||'');
  const [day,setDay]=useState('จันทร์');
  const [periodId,setPeriodId]=useState('1');
  const [saving,setSaving]=useState(false);
  const existing=plcSessions.find(p=>p.subjectArea===area);
  useEffect(()=>{if(existing){setDay(existing.day);setPeriodId(existing.periodId);}},[area,existing?.day,existing?.periodId]);
  const members=plcTeachersForArea(area,teachers);
  const options=useMemo(()=>rankPLCSlots({subjectArea:area,teachers,periods,schedules,fixedSchedules,
    activityGroups,unavailabilities,plcSessions:plcSessions.filter(x=>x.subjectArea!==area)}),
    [area,teachers,periods,schedules,fixedSchedules,activityGroups,unavailabilities,plcSessions]);
  const selected=options.find(x=>x.day===day&&String(x.periodId)===String(periodId));
  const anchoredSchedules=schedules.filter(x=>!isAutoGeneratedSchedule(x));
  const movableSchedules=schedules.filter(x=>isAutoGeneratedSchedule(x)&&!x.keepOnRegenerate);
  const candidateIssues=plcSlotIssues({subjectArea:area,day,periodId,teachers,
    schedules:schedules.filter(x=>!isAutoGeneratedSchedule(x)||x.keepOnRegenerate),
    fixedSchedules,activityGroups,unavailabilities,
    plcSessions:plcSessions.filter(x=>x.subjectArea!==area),periods});
  const needsAutoMove=selected&&selected.issues.length>0 && candidateIssues.length===0 &&
    movableSchedules.some(x=>x.day===day&&String(x.periodId)===String(periodId)&&members.some(t=>t.id===x.teacherId));
  const save=async()=>{
    if(!rulesReady||saving||!members.length) return showToast('ข้อมูลยังไม่พร้อมหรือกลุ่มไม่มีครู','error');
    setSaving(true);
    try{
      const termKey=getTermKey(schoolSettings);
      const path=`school_data/${appId}/terms/${termKey}`;
      // Read all relevant inputs from SERVER immediately before writing the reservation.
      const [ts,ss,fs,gs,us,ps]=await Promise.all([
        getDocsFromServer(collection(db,`school_data/${appId}/teachers`)),
        getDocsFromServer(collection(db,`${path}/schedules`)),
        getDocsFromServer(collection(db,`${path}/fixedSchedules`)),
        getDocsFromServer(collection(db,`${path}/activityGroups`)),
        getDocsFromServer(collection(db,`${path}/unavailabilities`)),
        getDocsFromServer(collection(db,`${path}/plcSessions`))]);
      const rows=snap=>snap.docs.map(d=>({...d.data(),id:d.id}));
      if(getTermKey(schoolSettings)!==termKey)throw new Error('ภาคเรียนเปลี่ยนไป กรุณาทำใหม่');
      const remotePLC=rows(ps),id=plcIdForArea(area);
      if(remotePLC.some(x=>x.subjectArea===area&&x.id!==id))
        throw new Error('พบ PLC กลุ่มนี้มากกว่าหนึ่งรายการ กรุณาตรวจข้อมูลเดิม');
      const conflictsWithAuto=rows(ss).filter(x=>x.day===day&&String(x.periodId)===String(periodId)&&
        rows(ts).some(t=>t.id===x.teacherId&&t.subjectArea===area));
      const issues=plcSlotIssues({subjectArea:area,day,periodId,teachers:rows(ts),
        schedules:rows(ss).filter(x=>!isAutoGeneratedSchedule(x)||x.keepOnRegenerate),fixedSchedules:rows(fs),activityGroups:rows(gs),
        unavailabilities:rows(us),plcSessions:remotePLC.filter(x=>x.id!==id),periods});
      if(issues.length) throw new Error(issues.slice(0,3).map(x=>`${x.teacherName||'-'}: ${x.reason}`).join(' / '));
      if(conflictsWithAuto.length && !window.confirm(`คาบ PLC นี้ทับคาบ Auto ที่ย้ายได้ ${conflictsWithAuto.length} รายการ\nเมื่อบันทึก PLC ตารางจะมีข้อขัดแย้งชั่วคราวจนกว่าจะจัด Auto ใหม่\nต้องการจอง PLC และกลับไปกดจัด Auto ใหม่ใช่ไหม?`)) return;
      await setDoc(doc(db,`${path}/plcSessions`,id),{id,subjectArea:area,day,periodId});
      showToast(conflictsWithAuto.length
        ? 'จอง PLC แล้ว กรุณาไปจัดตาราง Auto ใหม่และยืนยันผลก่อนประกาศใช้'
        : 'บันทึก PLC สำเร็จ ครูทั้งกลุ่มจะถูกกันคาบนี้ออกจากการจัด Auto');
    }catch(e){console.error('PLC save failed',e);showToast('ยังบันทึก PLC ไม่ได้: '+e.message,'error');}
    finally{setSaving(false);}
  };
  const remove=async()=>{
    if(!existing||saving||!window.confirm(`ยืนยันลบเวลาประชุม PLC กลุ่ม${area} หรือไม่?`))return;
    setSaving(true);
    try{await deleteDoc(doc(db,`school_data/${appId}/terms/${getTermKey(schoolSettings)}/plcSessions`,existing.id));
      showToast('ลบเวลา PLC แล้ว');}
    catch(e){showToast('ลบ PLC ไม่สำเร็จ','error');}finally{setSaving(false);}
  };
  return <div className="space-y-5">
    <div className="bg-white border rounded-2xl p-5 space-y-4">
      <h3 className="font-bold text-lg">PLC กลุ่มสาระ — 1 คาบต่อสัปดาห์</h3>
      <p className="text-sm text-slate-600">ขั้นตอนที่แนะนำ: กำหนด PLC ของแต่ละกลุ่มสาระให้เรียบร้อยก่อนลงวิชาหรือกดจัด Auto แล้วระบบจะกันคาบนี้ให้ครูทุกคนในกลุ่มโดยอัตโนมัติ</p>
      <p className="text-sm text-slate-600">PLC กันเวลาครู ไม่เป็นวิชาสอนและไม่ใช้ช่องตารางเรียนของนักเรียน สมาชิกครูอ้างอิงจากข้อมูลครู &gt; กลุ่มสาระ</p>
      <label className="block text-sm font-bold">กลุ่มสาระ</label>
      <select value={area} onChange={e=>setArea(e.target.value)} className="border rounded-xl p-3 w-full">
        {PLC_AREAS.map(a=><option key={a} value={a}>{a}</option>)}
      </select>
      <p className="text-sm">ครูในกลุ่ม ({members.length} คน): {members.map(t=>t.name).join(', ')||'ยังไม่มีครูในกลุ่ม'}</p>
      <div className="grid grid-cols-2 gap-3">
        <label className="text-sm">วัน<select value={day} onChange={e=>setDay(e.target.value)} className="border rounded-xl p-3 w-full mt-1">{DAYS.map(d=><option key={d}>{d}</option>)}</select></label>
        <label className="text-sm">คาบ<select value={periodId} onChange={e=>setPeriodId(e.target.value)} className="border rounded-xl p-3 w-full mt-1">{periods.filter(p=>!p.isBreak).map(p=><option key={p.id} value={p.id}>{p.name}</option>)}</select></label>
      </div>
      <div className="flex flex-wrap gap-2 items-center">
        <button onClick={()=>{const first=options.find(o=>!o.issues.length);if(first){setDay(first.day);setPeriodId(first.periodId);}else showToast('ยังไม่พบเวลาที่ทุกคนว่างพร้อมกัน กรุณาปรับคาบ Auto หรือข้อจำกัดก่อน','error');}}
          className="px-4 py-3 rounded-xl border border-blue-200 bg-blue-50 font-bold">ค้นหาคาบที่ว่างร่วมกัน</button>
        <button disabled={!rulesReady||saving||!selected||!!candidateIssues.length||!members.length} onClick={save}
          className="bg-[#081a39] text-white px-5 py-3 rounded-xl font-bold disabled:opacity-50">{saving?'กำลังบันทึก...':needsAutoMove?'จอง PLC แล้วจัด Auto ใหม่':'ยืนยันเวลา PLC'}</button>
        {existing&&<button disabled={saving} onClick={remove} className="border border-rose-200 text-rose-700 px-4 py-3 rounded-xl">ลบ PLC</button>}
      </div>
      {selected?.issues.length ? <div className="bg-amber-50 p-3 rounded-xl text-sm text-amber-900">
        เวลาที่เลือกยังไม่ว่างพร้อมกัน: {selected.issues.slice(0,5).map(x=>`${x.teacherName||'-'} (${x.reason})`).join(', ')}
        <p className="mt-2">{needsAutoMove?'สามารถจองคาบนี้แล้วกดจัด Auto ใหม่ทันทีได้ เพราะไม่ทับคาบมือหรือ Auto ที่ล็อก':'คาบที่จัดด้วยมือและ Auto ที่ล็อกแล้วจะไม่ถูกย้าย ให้เลือกคาบอื่นหรือแก้เงื่อนไขก่อน'}</p>
      </div>:<p className="text-emerald-700 text-sm">{members.length?'ครูทั้งหมดในกลุ่มว่างตรงกันในเวลาที่เลือก':'กรุณากำหนดครูในกลุ่มสาระก่อน'}</p>}
    </div>
    <div className="bg-white rounded-2xl p-5 border">
      <h3 className="font-bold mb-3">ช่วงเวลาว่างร่วมที่แนะนำ</h3>
      <div className="flex flex-wrap gap-2">{options.filter(o=>!o.issues.length).slice(0,8).map(o=><button key={o.day+'_'+o.periodId}
        onClick={()=>{setDay(o.day);setPeriodId(o.periodId);}} className="border rounded-lg bg-emerald-50 px-3 py-2 text-sm">{o.day} {o.periodName}</button>)}</div>
      {!options.some(o=>!o.issues.length)&&<p className="text-amber-800 text-sm">ไม่พบคาบว่างร่วมกันทันที อาจจองคาบที่ทับเฉพาะ Auto ที่ไม่ล็อก แล้วจัด Auto ใหม่เพื่อเปิดเวลาร่วมกัน</p>}
    </div>
    <div className="bg-white rounded-2xl p-5 border"><h3 className="font-bold mb-2">PLC ที่กำหนดในภาคเรียนนี้</h3>
      {plcSessions.map(p=><p key={p.id} className="border-b py-2 text-sm">{p.subjectArea}: {p.day} {periods.find(x=>x.id===p.periodId)?.name||p.periodId} ({plcTeachersForArea(p.subjectArea,teachers).length} คน)</p>)}
      {!plcSessions.length&&<p className="text-slate-500 text-sm">ยังไม่มี PLC ที่กำหนด</p>}
    </div>
  </div>;
};

// STAGE 13: navigation-only grouping. Existing editors retain their original data and write paths.
const ManagementTabs = ({ title, tabs, value, onChange }) => (
  <div className="bg-white rounded-2xl border border-slate-200 shadow-sm p-4 sm:p-5 no-print">
    <h2 className="text-xl font-bold text-[#081a39] mb-4">{title}</h2>
    <div role="tablist" aria-label={title} className="flex flex-wrap gap-2">
      {tabs.map(item => {
        const Icon = item.icon;
        const selected = value === item.id;
        return <button
          key={item.id} type="button" role="tab" aria-selected={selected}
          onClick={() => onChange(item.id)}
          className={`inline-flex items-center gap-2 rounded-xl px-4 py-3 font-bold text-sm transition-colors ${selected ? 'bg-[#081a39] text-[#d4af37] shadow-md' : 'bg-slate-50 border border-slate-200 text-slate-700 hover:bg-slate-100'}`}>
          <Icon size={17} />{item.name}
        </button>;
      })}
    </div>
  </div>
);

const BasicDataHub = ({ commonProps, autoGenerateClassrooms, onOpenDevelopment }) => {
  const [section, setSection] = useState('teachers');
  const tabs = [
    { id: 'teachers', name: 'ข้อมูลครู', icon: Users },
    { id: 'subjects', name: 'รายวิชา', icon: BookOpen },
    { id: 'classrooms', name: 'ห้องเรียน', icon: Home },
  ];
  return <div className="space-y-6 animate-in fade-in">
    <ManagementTabs title="ข้อมูลพื้นฐาน" tabs={tabs} value={section} onChange={setSection} />
    {section === 'teachers' && <TeachersView {...commonProps} />}
    {section === 'subjects' && <SubjectsView {...commonProps} onOpenDevelopment={onOpenDevelopment} importBasicSubjects={() => {}} />}
    {section === 'classrooms' && <ClassroomsView {...commonProps} autoGenerateClassrooms={autoGenerateClassrooms} />}
  </div>;
};

const ConstraintsHub = ({ commonProps, initialSection = 'time' }) => {
  const [section, setSection] = useState(initialSection);
  const tabs = [
    { id: 'time', name: 'ข้อจำกัดเวลา', icon: Clock },
    { id: 'rules', name: 'เงื่อนไขการจัดตาราง', icon: Filter },
    { id: 'fixed', name: 'กิจกรรมล็อกตายตัว', icon: Pin },
    { id: 'development', name: 'กิจกรรมพัฒนาผู้เรียน', icon: Activity },
    { id: 'plc', name: 'PLC กลุ่มสาระ', icon: Users },
  ];
  return <div className="space-y-6 animate-in fade-in">
    <ManagementTabs title="ข้อจำกัดและเงื่อนไข" tabs={tabs} value={section} onChange={setSection} />
    {section === 'time' && <ConstraintsView key="time-constraints" mode="time" {...commonProps} />}
    {section === 'rules' && <CustomRulesView {...commonProps} />}
    {section === 'plc' && <PLCView {...commonProps} />}
    {section === 'fixed' && <ConstraintsView key="fixed-activities" mode="fixed" {...commonProps} />}
    {section === 'development' && <>
      <div className="bg-blue-50 text-blue-900 rounded-xl border border-blue-100 p-4 text-sm">กิจกรรมพัฒนาผู้เรียนกำหนดชื่อ วัน คาบ ห้อง และครูผู้รับผิดชอบด้วยตนเอง เมื่อบันทึกจะล็อกทันที Auto ไม่ย้ายหรือจัดวิชาทับ แสดงเฉพาะชื่อกิจกรรมในตารางครู/นักเรียนและเอกสารพิมพ์ กิจกรรมเดิมย้ายประเภทได้ที่เมนูกิจกรรมล็อกตายตัว</div>
      <ConstraintsView key="development-activities" mode="development" {...commonProps} />
      <details className="bg-white rounded-2xl border border-slate-200 p-4 sm:p-6 shadow-sm">
        <summary className="cursor-pointer font-bold text-[#081a39]">
          กำหนดกิจกรรมพัฒนาผู้เรียนแยกตามห้อง (เฉพาะกรณีที่ต้องการ)
        </summary>
        <p className="text-sm text-slate-600 mt-3 mb-4">สร้างกิจกรรมหลักแบบทุกห้องหรือเลือกหลายห้องไว้ก่อน แล้วเปิดส่วนนี้เพื่อกำหนดชื่อกิจกรรมและครูรับผิดชอบรายห้อง</p>
        <ActivityGroupsView {...commonProps} />
      </details>
    </>}
  </div>;
};

// ใช้ข้อมูลภาระสอนจริงเพื่อกำหนดห้องที่เลือกได้เมื่อเพิ่มคาบด้วยมือ
// หากยังไม่เลือกรายวิชาให้คงการเลือกห้องแบบเดิม เพื่อให้เลือกห้องก่อนรายวิชาได้
const buildManualRoomOptions = (classrooms, teachingLoads, schedules, fixedSchedules, teacherId, subjectId) => {
  if (!teacherId || !subjectId) return classrooms.map(room => ({ ...room, optionLabel: room.name }));
  const loadsByRoom = new Map(teachingLoads.filter(load =>
    load.teacherId === teacherId && load.subjectId === subjectId
  ).map(load => [load.classroomId, load]));
  return classrooms.filter(room => loadsByRoom.has(room.id)).map(room => {
    const load = loadsByRoom.get(room.id);
    const required = Math.max(0, Number(load.periods) || 0);
    const belongsToLoad = row => row.teacherId === teacherId &&
      row.subjectId === subjectId && row.classroomId === room.id;
    const assigned = schedules.filter(belongsToLoad).length + fixedSchedules.filter(belongsToLoad).length;
    const remaining = Math.max(0, required - assigned);
    return {
      ...room,
      remaining,
      isComplete: remaining === 0,
      optionLabel: remaining === 0
        ? `${room.name} — จัดครบแล้ว (${required}/${required} คาบ)`
        : `${room.name} — เหลือ ${remaining}/${required} คาบ`
    };
  });
};

const ALL_TIMETABLE_VIEW_ID = '__BHS_ALL_TIMETABLES__';

const ScheduleView = ({ activityGroups, schedules, teachingLoads, classrooms, teachers, activeSubjects, periods, unavailabilities, classroomUnavailabilities, fixedSchedules, customRules, rulesReady, getSubject, getShortTeacherName, getClassroom, getTeacher, runSmartSchedule, clearUnlockedSchedules, dbAction, showToast, handleRequestDelete, isScheduling, db, appId, getTermKey, schoolSettings, saveScheduleVersion, publishDraftSchedule, publishedSnapshot, publishedLoading, publishedError, plcSessions = [], initialRoomId = '' }) => {
  const [viewMode, setViewMode] = useState('room');
  const [selectedId, setSelectedId] = useState(initialRoomId);
  const [hoveredSlot, setHoveredSlot] = useState(null);
  const [draggedScheduleId, setDraggedScheduleId] = useState('');
  const [dropFeedback, setDropFeedback] = useState(null);
  const dragSourceIdRef = useRef('');
  const dropKeyRef = useRef('');
  const dropOutcomeRef = useRef(null);
  const teacherActivitySlots = useMemo(() => activityTeacherEntries(activityGroups,fixedSchedules), [activityGroups,fixedSchedules]);
  const teacherPLCSlots = useMemo(() => buildPLCReservations(plcSessions,teachers), [plcSessions,teachers]);
  const displayedEntities = useMemo(() => {
    const source = viewMode === 'room' ? classrooms : teachers;
    return selectedId === ALL_TIMETABLE_VIEW_ID ? source : source.filter(item => item.id === selectedId);
  }, [viewMode, selectedId, classrooms, teachers]);
  const clearDragVisuals = () => {
    dragSourceIdRef.current = '';
    dropKeyRef.current = '';
    dropOutcomeRef.current = null;
    setDraggedScheduleId('');
    setDropFeedback(null);
  };
  useEffect(() => { if (initialRoomId) { setViewMode('room'); setSelectedId(initialRoomId); } }, [initialRoomId]);
  const [isAddModalOpen, setIsAddModalOpen] = useState(false);
  const [feasibility, setFeasibility] = useState(null);
  const [editingRows,setEditingRows] = useState(false);
  const [clearModalOpen,setClearModalOpen] = useState(false);
  const [clearRoomId,setClearRoomId] = useState('');
  const resetPreview = useMemo(() => planUnlockedScheduleClear(schedules, clearRoomId), [schedules, clearRoomId]);
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
  const manualRoomOptions = useMemo(() => buildManualRoomOptions(
    classrooms, teachingLoads, schedules, fixedSchedules, addForm.teacherId, addForm.subjectId
  ), [classrooms, teachingLoads, schedules, fixedSchedules, addForm.teacherId, addForm.subjectId]);
  const selectedManualRoom = manualRoomOptions.find(room => room.id === addForm.classroomId);
  const handleManualRoomChange = classroomId => {
    // เปลี่ยนห้องแล้วคงรายวิชาเดิมถ้าครูมีภาระสอนวิชานี้ในห้องใหม่
    const currentSubjectId = addForm.subjectId;
    const subjectStillValid = !classroomId || !currentSubjectId || teachingLoads.some(load =>
      load.teacherId === addForm.teacherId && load.subjectId === currentSubjectId && load.classroomId === classroomId
    );
    setAddForm(prev => ({ ...prev, classroomId, subjectId: subjectStillValid ? prev.subjectId : '' }));
    if (!subjectStillValid) showToast('ห้องนี้ไม่มีภาระสอนในรายวิชาเดิม กรุณาเลือกรายวิชาใหม่', 'error');
  };
  // Only suggest HARD-VALID slots under the normal four-period/day target.
  const manualSlotSuggestions = useMemo(()=>{
    if(!isAddModalOpen || !rulesReady || !addForm.teacherId || !addForm.subjectId || !addForm.classroomId) return [];
    const subject=activeSubjects.find(x=>x.id===addForm.subjectId);
    const block=Number(subject?.periodBlock)===2?2:1;
    const load=teachingLoads.find(x=>x.teacherId===addForm.teacherId&&x.subjectId===addForm.subjectId&&x.classroomId===addForm.classroomId);
    if(!load) return [];
    const remaining=Number(load.periods)-(schedules.filter(x=>x.teacherId===load.teacherId&&x.subjectId===load.subjectId&&x.classroomId===load.classroomId).length+
      fixedSchedules.filter(x=>x.teacherId===load.teacherId&&x.subjectId===load.subjectId&&x.classroomId===load.classroomId).length);
    if(remaining<block) return [];
    const maps=buildConstraintMaps(schedules,unavailabilities,classroomUnavailabilities,fixedSchedules,periods,customRules,activityGroups,classrooms,plcSessions,teachers);
    const state=buildPlacementState([...markFixedTeacherActivities(fixedSchedules),...activityTeacherEntries(activityGroups,fixedSchedules),...schedules]);
    const slots=[];
    DAYS.forEach(day=>periods.forEach((p,index)=>{
      if(p.isBreak||!canPlaceSchedule(load.teacherId,load.classroomId,day,p.id,maps,false,load.subjectId).ok) return;
      const entries=[{day,periodId:p.id}];
      if(block===2){const next=periods[index+1];
        if(!next||next.isBreak||!canPlaceSchedule(load.teacherId,load.classroomId,day,next.id,maps,false,load.subjectId).ok)return;
        entries.push({day,periodId:next.id});
      }
      if(!canPlaceTeacherEntries(load.teacherId,entries,maps).ok||!canPlaceSubjectDailyEntries(load,entries,maps).ok) return;
      const score=scorePlacementOption(entries,load,state,periods,id=>activeSubjects.find(x=>x.id===id)||{},customRules);
      slots.push({day,periodId:p.id,periodName:p.name,score,block});
    }));
    return slots.sort((a,b)=>b.score-a.score).slice(0,5);
  },[isAddModalOpen,addForm.teacherId,addForm.subjectId,addForm.classroomId,
    schedules,teachingLoads,activeSubjects,fixedSchedules,activityGroups,periods,
    unavailabilities,classroomUnavailabilities,customRules,classrooms,rulesReady]);

  const handleCellClick = (day, periodId, entityId) => {
    if (!entityId || isScheduling || editingRows) return;
    setAddForm({ teacherId: viewMode === 'teacher' ? entityId : '', subjectId: '', classroomId: viewMode === 'room' ? entityId : '', day, periodId });
    setIsAddModalOpen(true);
  };

  const handleFeasibilityClick = () => {
    if (!rulesReady) return showToast('ยังโหลดเงื่อนไขเพิ่มเติมไม่สำเร็จ กรุณาลองใหม่', 'error');
    setFeasibility(analyzeScheduleFeasibility({ teachingLoads, teachers, subjects:activeSubjects,
      classrooms, periods, schedules, fixedSchedules, unavailabilities,
      classroomUnavailabilities, customRules, semester:schoolSettings.semester, activityGroups, plcSessions }));
  };

  const handleManualAdd = async () => {
    if (!rulesReady) return showToast('ยังโหลดเงื่อนไขไม่สำเร็จ', 'error');
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

    const maps = buildConstraintMaps(schedules, unavailabilities, classroomUnavailabilities, fixedSchedules, periods, customRules, activityGroups, classrooms, plcSessions, teachers);
    const subject = getSubject(addForm.subjectId);
    const blockSize = Number(subject?.periodBlock) === 2 ? 2 : 1;
    const bookedToday = (maps.teacherDayPeriods.get(teacherDayKey(addForm.teacherId, addForm.day)) || new Set()).size;
    // ผู้จัดตารางอนุมัติข้อยกเว้น 5 คาบด้วยตนเองเท่านั้น
    if (bookedToday + blockSize > SCHEDULER_CONFIG.maxTeacherPeriodsPerDay) {
      if (bookedToday + blockSize > SCHEDULER_CONFIG.absoluteMaxTeacherPeriodsPerDay)
        return showToast('ครูมีคาบสอนและกิจกรรมรวมเกิน 5 คาบต่อวันไม่ได้', 'error');
      if (!window.confirm(`ครูคนนี้มีคาบสอนและกิจกรรมรวม ${bookedToday} คาบในวัน${addForm.day} แล้ว\nเพิ่มอีก ${blockSize} คาบจะเป็น ${bookedToday + blockSize} คาบ ซึ่งเกินเป้าหมาย 4 คาบ\nยืนยันกรณีจำเป็นจริง ๆ หรือไม่?`)) return;
      maps.maxTeacherPeriodsPerDay = SCHEDULER_CONFIG.absoluteMaxTeacherPeriodsPerDay;
    }
    const selectedValid = canPlaceSchedule(addForm.teacherId, addForm.classroomId, addForm.day, addForm.periodId, maps, false, addForm.subjectId, true, false, true);
    if (!selectedValid.ok) {
      const reasons = { "BREAK_PERIOD": "จัดลงคาบพักเที่ยงไม่ได้", "ROOM_UNAVAILABLE": "ห้องไม่สะดวก", "TEACHER_UNAVAILABLE": "ครูไม่สะดวก", "FIXED_SCHEDULE_CONFLICT": "มีกิจกรรมล็อกแล้ว", "ROOM_CONFLICT": "ห้องมีวิชาอื่นแล้ว", "TEACHER_CONFLICT": "ครูมีสอนแล้ว", "TEACHER_DAILY_LIMIT": "ครูสอนเกินเพดานต่อวันที่อนุมัติไม่ได้", "TEACHER_CONSECUTIVE_LIMIT": "ครูสอนเกิน 2 คาบติดต่อกันไม่ได้", "CUSTOM_RULE_BLOCKED": "ขัดกับเงื่อนไขห้ามจัดที่ตั้งไว้", "SUBJECT_DAILY_LIMIT": "วิชานี้เต็มจำนวนคาบสูงสุดต่อวันของห้องแล้ว" };
      return showToast(reasons[selectedValid.reason] || "ติดเงื่อนไขจัดไม่ได้", "error");
    }

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
            const neighborValid = canPlaceSchedule(addForm.teacherId, addForm.classroomId, addForm.day, neighbor.id, maps, false, addForm.subjectId, true, false, true);
            if (neighborValid.ok) {
              const pair = [{ day: addForm.day, periodId: addForm.periodId }, { day: addForm.day, periodId: neighbor.id }];
              if (canPlaceTeacherEntries(addForm.teacherId, pair, maps, false, true).ok && canPlaceSubjectDailyEntries(addForm, pair, maps).ok) candidatePairs.push(pair);
            }
          });

          if (candidatePairs.length === 0) return showToast("วิชานี้กำหนดเป็น Block 2 แต่ไม่มีคาบข้างเคียงที่ว่างสำหรับจัดติดกัน", "error");

          const placementState = buildPlacementState([...markFixedTeacherActivities(fixedSchedules), ...schedules]);
          candidatePairs.sort((a, b) => scorePlacementOption(b, matchingLoad, placementState, periods, getSubject, customRules) - scorePlacementOption(a, matchingLoad, placementState, periods, getSubject, customRules));
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

    const subjectLimit = canPlaceSubjectDailyEntries(addForm, plannedEntries, maps);
    if (!subjectLimit.ok) return showToast('วิชาเดียวกันเกินจำนวนคาบสูงสุดต่อวันของห้องนี้', 'error');
    const teacherLimit = canPlaceTeacherEntries(addForm.teacherId, plannedEntries, maps, false, true);
    if (!teacherLimit.ok) return showToast(teacherLimit.reason === 'TEACHER_DAILY_LIMIT'
      ? 'ครูสอนเกินเพดานต่อวันที่อนุมัติไม่ได้' : 'ครูสอนต่อเนื่องผิดเงื่อนไขหรือรูปแบบ 4 คาบไม่อนุญาต', 'error');

    const afterInstruction = [...(maps.teacherInstructionDayPeriods.get(teacherDayKey(addForm.teacherId,addForm.day)) || []),
      ...plannedEntries.map(item=>item.periodId)];
    const newThreeStreak = getTeacherTeachingRuns(afterInstruction,periods).find(streak=>
      streak.length===3 && plannedEntries.some(item=>streak.includes(String(item.periodId))));
    const consecutiveApproval = newThreeStreak ? {teacherId:addForm.teacherId,day:addForm.day,periodIds:newThreeStreak} : null;
    if (consecutiveApproval && !window.confirm(`คำเตือน: ครู ${getTeacher(addForm.teacherId).name} จะสอนรายวิชา 3 คาบติดต่อกันในวัน${addForm.day}\nคาบ: ${newThreeStreak.map(id=>getPeriodDisplayName(periods,id)).join(', ')}\n\nการจัดด้วยมือครั้งนี้เป็นข้อยกเว้นเท่านั้น ระบบ Auto ยังคงห้ามสร้าง 3 คาบติดเอง\nยืนยันให้บันทึกคาบดังกล่าวหรือไม่?`)) return;

    const closedEntries=plannedEntries.filter(entry=>maps.cUnavailMap.has(`${addForm.classroomId}_${entry.day}_${entry.periodId}`));
    if (closedEntries.length && !window.confirm(`ห้อง ${getClassroom(addForm.classroomId).name} ถูกปิดการจัด Auto ในวัน${addForm.day} คาบ ${closedEntries.map(entry=>getPeriodDisplayName(periods,entry.periodId)).join(', ')}\n\nอนุญาตให้จัดด้วยมือเป็นกรณีพิเศษหรือไม่? คาบนี้จะคงไว้เมื่อจัด Auto ใหม่`)) return;
    try {
      if (saveScheduleVersion) await saveScheduleVersion('before_manual_add');
      const batch = writeBatch(db);
      const termPath = `school_data/${appId}/terms/${getTermKey(schoolSettings)}`;
      plannedEntries.forEach(entry => {
        const scheduleRef = doc(collection(db, `${termPath}/schedules`));
        batch.set(scheduleRef, {
          id: scheduleRef.id, teacherId: addForm.teacherId, subjectId: addForm.subjectId, classroomId: addForm.classroomId, day: entry.day, periodId: entry.periodId, source: 'manual', ...(consecutiveApproval ? { manualConsecutiveApproval: consecutiveApproval } : {}), manualRoomClosureOverride: maps.cUnavailMap.has(`${addForm.classroomId}_${entry.day}_${entry.periodId}`)
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

  // STAGE 7: preserve selected generated lessons when re-running the scheduler.
  const toggleKeepSchedule = async item => {
    if(isScheduling||editingRows) return;
    const target=schedules.find(x=>x.id===item.id);
    if(!target) return;
    if(!isAutoGeneratedSchedule(target) && target.source!=='auto')
      return showToast('คาบที่เพิ่มด้วยมือถูกเก็บไว้เมื่อจัดใหม่อยู่แล้ว');
    const next=target.keepOnRegenerate!==true;
    const related=schedules.filter(x=>x.teacherId===target.teacherId&&x.subjectId===target.subjectId&&x.classroomId===target.classroomId);
    const partner=Number(getSubject(target.subjectId)?.periodBlock)===2 ? findBlockPartner(target,related,periods) : null;
    const affected=partner?[target,partner]:[target];
    setEditingRows(true);
    try {
      const snapshotId=await saveScheduleVersion('before_toggle_keep');
      if(!snapshotId) throw new Error('สร้างจุดย้อนกลับไม่สำเร็จ');
      const batch=writeBatch(db);
      const termPath=`school_data/${appId}/terms/${getTermKey(schoolSettings)}`;
      affected.forEach(row=>batch.set(doc(db,`${termPath}/schedules`,row.id),{...row,keepOnRegenerate:next}));
      await batch.commit();
      showToast(next?'ล็อกคาบไว้สำหรับการจัดใหม่แล้ว':'ปลดล็อกคาบสำหรับจัดใหม่แล้ว');
    }catch(error){console.error(error);showToast('เปลี่ยนสถานะล็อกไม่สำเร็จ','error');}
    finally{setEditingRows(false);}
  };

  // STAGE 8: drop keeps source room/teacher/subject, and moves Block-2 pairs together.
  const moveScheduleByDrag = async (scheduleId, targetDay, targetPeriodId) => {
    if(isScheduling || editingRows || !rulesReady) return;
    let movePlan;
    try {
      movePlan=buildScheduleDragPlan({schedules,sourceId:scheduleId,targetDay,targetPeriodId,
        periods,subject:getSubject(schedules.find(s=>s.id===scheduleId)?.subjectId)});
    }catch(error){return showToast('ย้ายไม่ได้: '+error.message,'error');}
    const {source,moved,planned,remaining}=movePlan;
    if(planned.every((item,i)=>item.day===moved[i].day&&item.periodId===moved[i].periodId)) return;
    const maps=buildConstraintMaps(remaining,unavailabilities,classroomUnavailabilities,fixedSchedules,periods,customRules,activityGroups,classrooms,plcSessions,teachers);
    const existingGroupItems=activityTeacherEntries(activityGroups,fixedSchedules);
    const oldFifth=findTeacherFifthPeriodDays([...schedules,...fixedSchedules,...existingGroupItems],periods,teachers);
    const projectedFifth=findTeacherFifthPeriodDays([...remaining,...planned,...fixedSchedules,...existingGroupItems],periods,teachers);
    const newFifth=projectedFifth.filter(x=>!oldFifth.some(y=>x.teacherId===y.teacherId&&x.day===y.day));
    if(projectedFifth.length) maps.maxTeacherPeriodsPerDay=SCHEDULER_CONFIG.absoluteMaxTeacherPeriodsPerDay;
    const entries=planned.map(x=>({day:x.day,periodId:x.periodId}));
    for(const item of planned){
      const valid=canPlaceSchedule(item.teacherId,item.classroomId,item.day,item.periodId,maps,false,item.subjectId,false,false,true);
      if(!valid.ok) return showToast(`ย้ายไม่ได้: ${valid.reason}`,'error');
    }
    if(!canPlaceTeacherEntries(source.teacherId,entries,maps,false,true).ok || !canPlaceSubjectDailyEntries(source,entries,maps).ok)
      return showToast('ย้ายไม่ได้ เพราะเกินข้อบังคับครูหรือวิชาในวันเดียวกัน','error');
    const projectedInstruction=[...(maps.teacherInstructionDayPeriods.get(teacherDayKey(source.teacherId,targetDay)) || []),
      ...entries.map(x=>x.periodId)];
    const targetThreeStreak=getTeacherTeachingRuns(projectedInstruction,periods).find(streak=>
      streak.length===3 && entries.some(item=>streak.includes(String(item.periodId))));
    const moveApproval=targetThreeStreak ? {teacherId:source.teacherId,day:targetDay,periodIds:targetThreeStreak} : null;
    const approvedPlanned=planned.map(item=>{
      const next={...item};
      delete next.manualConsecutiveApproval; // approval from old location cannot travel to a new slot
      if(moveApproval) next.manualConsecutiveApproval=moveApproval;
      return next;
    });
    const outcome=findTeacherTeachingLimitConflicts(buildTeacherOccupancy([...remaining,...approvedPlanned],fixedSchedules,activityGroups),periods,teachers);
    if(outcome.length) return showToast('ย้ายไม่ได้: '+outcome[0],'error');
    const subjectIssues=findSubjectDailyLimitConflicts([...remaining,...planned,...fixedSchedules],classrooms,customRules);
    if(subjectIssues.length) return showToast('ย้ายไม่ได้: '+subjectIssues[0],'error');
    if(!window.confirm(`ย้าย ${moved.length} คาบ ไปวัน${targetDay} คาบ ${targetPeriodId} และตำแหน่งติดกัน (ถ้ามี) ใช่หรือไม่?${newFifth.length ? `\n\nคำเตือน: ${newFifth.map(x=>`ครู ${x.teacherName} วัน${x.day} จะมีภาระรวม 5 คาบ`).join(' / ')}\nยืนยันข้อยกเว้นเมื่อจำเป็นจริง ๆ` : ''}${moveApproval ? `\n\nคำเตือน: ครู ${getTeacher(source.teacherId).name} จะสอนรายวิชา 3 คาบติดกันวัน${targetDay} (${targetThreeStreak.map(id=>getPeriodDisplayName(periods,id)).join(', ')})\nการย้ายด้วยมือครั้งนี้ต้องได้รับอนุมัติ และ Auto ยังห้ามสร้าง 3 คาบติดเอง` : ''}`)) return;
    setEditingRows(true);
    try {
      const termPath=`school_data/${appId}/terms/${getTermKey(schoolSettings)}`;
      // Check the affected source documents from server to avoid a stale drag.
      for(const item of moved){
        const snap=await getDocFromServer(doc(db,`${termPath}/schedules`,item.id));
        if(!snap.exists() || !areFirestoreRecordsEqual({...snap.data(),id:snap.id},item))
          throw new Error('คาบต้นทางมีการเปลี่ยนแปลง กรุณาโหลดหน้าใหม่');
      }
      const latest=await getDocsFromServer(collection(db,`${termPath}/schedules`));
      const now=latest.docs.map(d=>({...d.data(),id:d.id}));
      if(fingerprintScheduleRows(now)!==fingerprintScheduleRows(schedules))
        throw new Error('ตารางถูกแก้จากเครื่องอื่น โปรดลองใหม่');
      const snapshotId=await saveScheduleVersion('before_drag_move');
      if(!snapshotId) throw new Error('สร้างจุดย้อนกลับไม่สำเร็จ');
      const batch=writeBatch(db);
      approvedPlanned.forEach(item=>batch.set(doc(db,`${termPath}/schedules`,item.id),item));
      await batch.commit();
      showToast(`ย้ายคาบสำเร็จ ${planned.length} คาบ`);
    } catch(error) { console.error(error);showToast('ย้ายคาบไม่สำเร็จ: '+error.message,'error'); }
    finally {setEditingRows(false);}
  };

  // Preview drag feasibility before displaying a green destination. The commit handler
  // still rechecks every hard constraint and verifies the latest server data.
  const previewDrop = (scheduleId, entityId, targetDay, targetPeriodId) => {
    if (!rulesReady || isScheduling || editingRows) return {status:'blocked', message:'กำลังโหลดกฎหรือระบบไม่พร้อมแก้ไข'};
    const source = schedules.find(item => item.id === scheduleId);
    if (!source) return {status:'blocked', message:'ไม่พบคาบต้นทาง'};
    if ((viewMode === 'room' && source.classroomId !== entityId) ||
        (viewMode === 'teacher' && source.teacherId !== entityId)) {
      return {status:'blocked',message:'ย้ายข้ามห้องหรือข้ามครูด้วยการลากไม่ได้ ต้องใช้ห้อง/ครูเดิม'};
    }
    let plan;
    try {
      plan = buildScheduleDragPlan({schedules,sourceId:scheduleId,targetDay,targetPeriodId,
        periods,subject:getSubject(source.subjectId)});
    } catch (error) {
      return {status:'blocked',message:USER_ERROR_TH[error.message] || error.message};
    }
    const {moved,planned,remaining} = plan;
    if (planned.every((item,i)=>item.day===moved[i].day && item.periodId===moved[i].periodId))
      return {status:'blocked',message:'เป็นช่องเดิม ไม่จำเป็นต้องย้าย'};
    const maps = buildConstraintMaps(remaining,unavailabilities,classroomUnavailabilities,
      fixedSchedules,periods,customRules,activityGroups,classrooms,plcSessions,teachers);
    const groupItems = activityTeacherEntries(activityGroups,fixedSchedules);
    const oldFifth = findTeacherFifthPeriodDays([...schedules,...fixedSchedules,...groupItems],periods,teachers);
    const projectedFifth = findTeacherFifthPeriodDays([...remaining,...planned,...fixedSchedules,...groupItems],periods,teachers);
    const newFifth = projectedFifth.filter(item=>!oldFifth.some(old=>old.teacherId===item.teacherId && old.day===item.day));
    if(projectedFifth.length) maps.maxTeacherPeriodsPerDay=SCHEDULER_CONFIG.absoluteMaxTeacherPeriodsPerDay;
    const entries=planned.map(item=>({day:item.day,periodId:item.periodId}));
    for(const item of planned){
      const result=canPlaceSchedule(item.teacherId,item.classroomId,item.day,item.periodId,maps,
        false,item.subjectId,false,false,true);
      if (!result.ok) return {status:'blocked',message:USER_ERROR_TH[result.reason] || result.reason};
    }
    const teacherCheck=canPlaceTeacherEntries(source.teacherId,entries,maps,false,true);
    if(!teacherCheck.ok) return {status:'blocked',message:USER_ERROR_TH[teacherCheck.reason] || teacherCheck.reason};
    const subjectCheck=canPlaceSubjectDailyEntries(source,entries,maps);
    if(!subjectCheck.ok) return {status:'blocked',message:USER_ERROR_TH[subjectCheck.reason] || subjectCheck.reason};
    const instruction=[...(maps.teacherInstructionDayPeriods.get(teacherDayKey(source.teacherId,targetDay))||[]),
      ...entries.map(item=>item.periodId)];
    const streak=getTeacherTeachingRuns(instruction,periods).find(run=>run.length===3 && entries.some(item=>run.includes(String(item.periodId))));
    const approval=streak ? {teacherId:source.teacherId,day:targetDay,periodIds:streak} : null;
    const approvedPlanned=planned.map(item=>{
      const copy={...item}; delete copy.manualConsecutiveApproval;
      if(approval) copy.manualConsecutiveApproval=approval;
      return copy;
    });
    const limitIssues=findTeacherTeachingLimitConflicts(buildTeacherOccupancy([...remaining,...approvedPlanned],fixedSchedules,activityGroups),periods,teachers);
    if(limitIssues.length) return {status:'blocked',message:limitIssues[0]};
    const subjectIssues=findSubjectDailyLimitConflicts([...remaining,...planned,...fixedSchedules],classrooms,customRules);
    if(subjectIssues.length) return {status:'blocked',message:subjectIssues[0]};
    if(approval || newFifth.length) return {status:'confirm',message:approval
      ? 'จะเกิดการสอน 3 คาบติดกัน ต้องยืนยันก่อนย้าย' : 'ครูจะมีภาระ 5 คาบในวันนี้ ต้องยืนยันก่อนย้าย'};
    return {status:'allowed',message:'ย้ายมาคาบนี้ได้ (ระบบจะตรวจซ้ำก่อนบันทึก)'};
  };

  const updateDropFeedback = (scheduleId, entityId, day, periodId) => {
    const key=JSON.stringify([scheduleId,entityId,day,String(periodId)]);
    if(dropKeyRef.current===key) return dropOutcomeRef.current?.status || 'blocked';
    dropKeyRef.current=key;
    // Browser has begun a native drag by the time the destination receives dragover.
    // Updating visual state here is safer than during onDragStart.
    if (draggedScheduleId !== scheduleId) setDraggedScheduleId(scheduleId);
    const outcome=previewDrop(scheduleId,entityId,day,periodId);
    dropOutcomeRef.current=outcome;
    setDropFeedback({key,entityId,day,periodId,...outcome});
    return outcome.status;
  };

  const getCellContent = (day, periodId, entityId) => {
     if (!entityId) return null;
     const cellSch = schedules.filter(s => viewMode === 'room' ? s.classroomId === entityId && s.day === day && s.periodId === periodId : s.teacherId === entityId && s.day === day && s.periodId === periodId);
     const cellFix = fixedSchedules.filter(f => viewMode === 'room' ? fixedAppliesToClassroom(f, entityId) && f.day === day && f.periodId === periodId : f.teacherId === entityId && f.day === day && f.periodId === periodId);
     const cellGroups = viewMode==='teacher' ? teacherActivitySlots.filter(x=>x.teacherId===entityId&&x.day===day&&x.periodId===periodId).map(x=>({...x,name:x.__activityName || fixedActivityDisplayName(x,getSubject)})) : [];
     const cellPLC = viewMode==='teacher' ? teacherPLCSlots
       .filter(x=>x.teacherId===entityId&&x.day===day&&x.periodId===periodId) : [];
     if (cellSch.length === 0 && cellFix.length === 0 && cellGroups.length===0 && cellPLC.length===0) return null;
     const isConflict = cellSch.length + cellFix.length + cellGroups.length + cellPLC.length > 1;

     return (
       <div className={`p-1.5 h-full w-full rounded-xl flex flex-col items-center justify-center text-xs text-center border-l-4 shadow-sm relative group ${isConflict ? 'bg-rose-50 border-rose-500' : 'bg-white border-[#d4af37]'}`}>
         {isConflict && <div className="absolute -top-2 -right-2 text-rose-500 bg-rose-100 rounded-full p-0.5 shadow-sm"><AlertCircle size={14}/></div>}
         {cellFix.map((f, idx) => (
           <div key={f.id} className={idx > 0 || cellSch.length > 0 ? "mt-2 pt-2 border-t w-full relative" : "w-full relative"}>
             <div className="font-bold text-amber-700 flex items-center justify-center gap-1">{viewMode==='room' && f.classroomId===ALL_CLASSROOMS_ID
               ? fixedActivityDisplayName(f,getSubject,activityGroups,entityId)
               : fixedActivityDisplayName(f,getSubject)}</div>
             <button onClick={(e) => { e.stopPropagation(); handleRequestDelete('fixed', f); }} className="absolute -top-1 -right-1 opacity-0 group-hover:opacity-100 bg-rose-500 text-white rounded-full p-0.5 z-10"><X size={12}/></button>
           </div>
         ))}
         {cellGroups.map(g=><div key={g.id} className="text-indigo-700 font-bold border-t pt-1">{g.name}</div>)}
         {cellPLC.map(item=><div key={item.id} className="text-blue-800 font-bold border-t pt-1">{plcDisplayName(item.subjectArea)}</div>)}
         {cellSch.map((sch, idx) => (
           <div key={sch.id} draggable={!isScheduling&&!editingRows&&rulesReady} onDragStart={e=>{
               // IMPORTANT: do not trigger React re-render on dragstart; a synchronous
               // grid/layout update may cancel the browser's native drag operation.
               if (isScheduling || editingRows || !rulesReady) { e.preventDefault(); return; }
               dragSourceIdRef.current=sch.id; dropKeyRef.current=''; dropOutcomeRef.current=null;
               e.dataTransfer.setData('text/plain', String(sch.id));
               e.dataTransfer.effectAllowed='move';
             }} onDragEnd={clearDragVisuals}
             title="ลากเพื่อย้ายคาบ หรือกดไอคอนเพื่อเก็บคาบนี้เมื่อจัดใหม่"
             className={`${idx > 0 || cellFix.length > 0 ? 'mt-2 pt-2 border-t w-full relative' : 'w-full relative'} cursor-grab active:cursor-grabbing rounded-md ${draggedScheduleId === sch.id ? 'bg-blue-100 ring-2 ring-blue-500 opacity-65' : 'hover:bg-blue-50'}`}>
             <div className="font-bold text-[#081a39] truncate">{getSubject(sch.subjectId).name}</div>
             {viewMode === 'room' ? (
               <><div className="text-slate-600 truncate">{getShortTeacherName(getTeacher(sch.teacherId).name)}</div></>
             ) : (
               <><div className="text-slate-600 truncate">{getClassroom(sch.classroomId).name}</div></>
             )}
             <button onClick={e=>{e.stopPropagation();toggleKeepSchedule(sch);}} title={sch.keepOnRegenerate?'ปลดล็อกคาบ':'ล็อกคาบไว้เมื่อจัดใหม่'} className="absolute top-0 left-0 text-xs text-indigo-700 bg-indigo-50 rounded p-0.5 z-10">{sch.keepOnRegenerate?'🔒':'🔓'}</button>
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
          <select value={viewMode} onChange={e => {setViewMode(e.target.value); setSelectedId(''); setHoveredSlot(null);clearDragVisuals();}} className="p-3 border border-slate-200 rounded-xl bg-slate-50 font-bold text-[#081a39]"><option value="room">ตารางเรียน (รายห้อง)</option><option value="teacher">ตารางสอน (รายครู)</option></select>
          <select value={selectedId} onChange={e => {setSelectedId(e.target.value);setHoveredSlot(null);clearDragVisuals();}} className="p-3 border border-slate-200 rounded-xl min-w-[250px] bg-slate-50 font-medium"><option value="">-- เลือกเพื่อแสดงตาราง --</option><option value={ALL_TIMETABLE_VIEW_ID}>{viewMode === 'room' ? `ทุกห้องเรียน (${classrooms.length} ห้อง)` : `ครูทั้งหมด (${teachers.length} คน)`}</option>{viewMode === 'room' ? classrooms.map(c => <option key={c.id} value={c.id}>{c.name}</option>) : teachers.map(t => <option key={t.id} value={t.id}>{t.name}</option>)}</select>
        </div>
        <div className="flex flex-wrap items-center gap-2">
          <button onClick={handleFeasibilityClick} disabled={isScheduling || !rulesReady} className="flex items-center gap-2 px-4 py-3 bg-slate-100 text-[#081a39] border border-slate-200 rounded-xl font-bold disabled:opacity-50"><Search size={17}/> ตรวจความเป็นไปได้</button>
          <button type="button" onClick={() => { setClearRoomId(''); setClearModalOpen(true); }} disabled={isScheduling || editingRows || !rulesReady || !schedules.length} className="flex items-center gap-2 px-4 py-3 border border-rose-200 bg-white text-rose-700 rounded-xl font-bold hover:bg-rose-50 disabled:opacity-50"><Trash2 size={17}/> ล้างคาบไม่ล็อก</button>
        <button onClick={publishDraftSchedule} disabled={isScheduling||!rulesReady} className="flex items-center gap-2 px-4 py-3 bg-emerald-700 text-white rounded-xl font-bold disabled:opacity-50">ประกาศใช้ฉบับร่าง</button>
        <button onClick={runSmartSchedule} disabled={isScheduling} className={`flex items-center gap-2 px-6 py-3 text-white rounded-xl shadow-md font-bold transition-all ${isScheduling ? 'bg-slate-400 cursor-not-allowed' : 'bg-gradient-to-r from-[#d4af37] to-[#b8952b] hover:shadow-lg'}`}><RefreshCw size={18} className={isScheduling ? "animate-spin" : ""} /> {isScheduling ? "กำลังจัดตาราง (AI)..." : "จัดตารางอัตโนมัติ"}</button>
        </div>
      </div>
      <div className="text-xs text-slate-600">กำลังแก้ไข: <strong>ฉบับร่าง</strong> · {publishedLoading?'กำลังตรวจฉบับประกาศ...':publishedError||(
        publishedSnapshot ? `มีฉบับประกาศใช้แล้ว (${new Date(publishedSnapshot.publishedAt).toLocaleString('th-TH')})` : 'ยังไม่เคยประกาศใช้')}
        {publishedSnapshot&&<span className="ml-2 text-amber-700">การแก้ร่างไม่เปลี่ยนฉบับประกาศ</span>}</div>
      {!selectedId && <div className="bg-white border border-dashed border-slate-300 text-center rounded-2xl p-12 text-slate-600">เลือกห้องเรียน ครู หรือเลือกทั้งหมดเพื่อดูตารางและจัดคาบด้วยมือ</div>}
      {draggedScheduleId && <div role="status" aria-live="polite" className={`fixed bottom-4 right-4 z-50 max-w-[min(92vw,520px)] pointer-events-none rounded-xl border-2 px-4 py-3 shadow-lg text-sm font-semibold ${
        dropFeedback?.status==='blocked' ? 'bg-rose-50 text-rose-900 border-rose-500' :
        dropFeedback?.status==='confirm' ? 'bg-amber-50 text-amber-900 border-amber-500' :
        dropFeedback?.status==='allowed' ? 'bg-emerald-50 text-emerald-900 border-emerald-500' :
        'bg-blue-50 text-blue-900 border-blue-500'
      }`}>
        กำลังลาก: {getSubject(schedules.find(s=>s.id===draggedScheduleId)?.subjectId).name || 'รายวิชา'}
        {dropFeedback ? ` → ${viewMode === 'room' ? getClassroom(dropFeedback.entityId).name : getTeacher(dropFeedback.entityId).name} วัน${dropFeedback.day} ${getPeriodDisplayName(periods,dropFeedback.periodId)} : ${dropFeedback.message}` : ' — เลื่อนไปยังช่องปลายทางเพื่อดูสถานะ'}
      </div>}
      <div className="space-y-6">
        {displayedEntities.map(entity => (
          <section key={`${viewMode}-${entity.id}`} className="bg-white rounded-2xl shadow-sm border border-slate-200 p-4" aria-label={`${viewMode==='room'?'ตารางเรียน':'ตารางสอน'} ${entity.name}`}>
            <div className="mb-3 flex flex-wrap items-center justify-between gap-2">
              <h3 className="font-bold text-lg text-[#081a39]">{viewMode==='room'?'ตารางเรียน':'ตารางสอน'} {entity.name}</h3>
              <span className="text-xs text-slate-500">คลิกช่องว่างเพื่อเพิ่มคาบด้วยมือ · ลากคาบเพื่อย้ายภายใน{viewMode==='room'?'ห้อง':'ตารางครู'}นี้</span>
            </div>
            <div className="overflow-x-auto" onMouseLeave={()=>{if(!dragSourceIdRef.current) setHoveredSlot(null);}}>
              <table className="w-full text-left border-collapse min-w-[900px] table-fixed">
                <thead className="bg-[#081a39] text-white border-b-4 border-[#d4af37]"><tr>
                  <th className="p-4 w-28 text-center font-bold">วัน/เวลา</th>
                  {periods.map(p => <th key={p.id} className="p-2 border-l border-[#122b5e] text-center text-sm">
                    <div className="font-bold">{p.name}</div><div className="text-[#d4af37] text-xs font-normal opacity-90">{p.start}-{p.end}</div>
                  </th>)}
                </tr></thead>
                <tbody>
                  {DAYS.map(day => <tr key={day} className="border-b border-slate-200 group">
                    <td className="p-4 font-bold text-center border-r border-slate-200 text-[#081a39] bg-slate-50">{day}</td>
                    {periods.map(p => {
                      const content=getCellContent(day,p.id,entity.id);
                      const isHovered=hoveredSlot?.entityId===entity.id && hoveredSlot.day===day && String(hoveredSlot.periodId)===String(p.id);
                      const isDrop=!!draggedScheduleId && dropFeedback?.entityId===entity.id && dropFeedback.day===day && String(dropFeedback.periodId)===String(p.id);
                      const state=isDrop ? dropFeedback.status : '';
                      const cellColor=state==='allowed' ? 'bg-emerald-100 ring-4 ring-inset ring-emerald-500 outline outline-2 outline-offset-[-3px] outline-emerald-500' :
                        state==='confirm' ? 'bg-amber-100 ring-4 ring-inset ring-amber-500 outline outline-2 outline-offset-[-3px] outline-amber-500' :
                        state==='blocked' ? 'bg-rose-100 ring-4 ring-inset ring-rose-500 outline outline-2 outline-offset-[-3px] outline-rose-500' :
                        p.isBreak ? 'bg-slate-100' : isHovered ? 'bg-blue-100 ring-2 ring-inset ring-blue-500' : 'bg-white hover:bg-blue-50';
                      return <td key={`${entity.id}-${day}-${p.id}`} title={`วัน${day} ${p.name} · ${entity.name}`}
                        className={`p-1.5 border-r border-b border-slate-200 h-28 align-top relative group/cell transition-colors duration-100 ${cellColor} ${!p.isBreak && !isScheduling && !editingRows ? 'cursor-pointer' : ''}`}
                        onMouseEnter={()=>{if(!dragSourceIdRef.current && !p.isBreak)setHoveredSlot({entityId:entity.id,day,periodId:p.id});}}
                        onDragOver={e=>{
                          const sourceId=dragSourceIdRef.current;
                          if(!sourceId) return;
                          // Native dragover must be cancelled for the table cell to receive onDrop.
                          e.preventDefault();
                          e.stopPropagation();
                          const status=updateDropFeedback(sourceId,entity.id,day,p.id);
                          e.dataTransfer.dropEffect=status==='blocked'?'none':'move';
                        }}
                        onDrop={e=>{
                          e.preventDefault();e.stopPropagation();
                          const sourceId=dragSourceIdRef.current || e.dataTransfer.getData('text/plain');
                          const feedback=sourceId ? previewDrop(sourceId,entity.id,day,p.id) : {status:'blocked',message:'ไม่พบคาบต้นทาง'};
                          clearDragVisuals();
                          if(feedback.status==='blocked') return showToast('ย้ายไม่ได้: '+feedback.message,'error');
                          moveScheduleByDrag(sourceId,day,p.id);
                        }}
                        onClick={()=>{if(!p.isBreak && !dragSourceIdRef.current) handleCellClick(day,p.id,entity.id);}}>
                        {p.isBreak ? <div className="h-full flex flex-col items-center justify-center text-slate-500 text-sm font-bold">พักเที่ยง{isDrop && <span className="text-xs text-rose-800 mt-1">✕ วางไม่ได้</span>}</div> : <>
                          {content}
                          {!content && !draggedScheduleId && !isScheduling && !editingRows && <div className="absolute inset-0 flex items-center justify-center opacity-0 group-hover/cell:opacity-100 pointer-events-none transition-opacity"><div className="bg-blue-200 border border-blue-400 px-2 py-1 rounded-lg text-blue-900 flex items-center gap-1 font-bold text-xs"><Plus size={16}/> เพิ่มคาบ</div></div>}
                          {isHovered && !draggedScheduleId && <span className="pointer-events-none absolute bottom-1 right-1 text-[10px] bg-blue-700 text-white rounded px-1.5 py-0.5 shadow">{day} · {p.name}</span>}
                          {isDrop && <span className={`pointer-events-none absolute bottom-1 right-1 z-20 text-[10px] font-bold rounded px-1.5 py-1 shadow ${state==='allowed'?'bg-emerald-700 text-white':state==='confirm'?'bg-amber-600 text-white':'bg-rose-700 text-white'}`}>
                            {state==='allowed'?'✓ วางได้':state==='confirm'?'! ต้องยืนยัน':'✕ วางไม่ได้'}
                          </span>}
                        </>}
                      </td>;
                    })}
                  </tr>)}
                </tbody>
              </table>
            </div>
          </section>
        ))}
      </div>
      <Modal isOpen={clearModalOpen} onClose={() => setClearModalOpen(false)} title="ล้างคาบที่ไม่ล็อก (ฉบับร่าง)">
        <div className="space-y-4 text-sm">
          <div className="bg-slate-50 border border-slate-200 rounded-xl p-4 space-y-3">
            <label className="block font-bold" htmlFor="clear-room-scope">เลือกขอบเขตการล้างคาบ</label>
            <select id="clear-room-scope" className="w-full border border-slate-300 rounded-xl p-3 bg-white"
              value={clearRoomId} onChange={e=>setClearRoomId(e.target.value)}>
              <option value="">ทุกห้องเรียน (ล้างตารางทั้งหมดที่ไม่ล็อก)</option>
              {classrooms.map(room=><option key={room.id} value={room.id}>{room.name}</option>)}
            </select>
            <p className="text-slate-600">ล้างคาบเรียนที่ยังไม่ได้กดล็อกทั้งหมดในขอบเขตที่เลือก ทั้งคาบ Auto และคาบลงด้วยมือ</p>
          </div>
          <div className="bg-rose-50 border border-rose-200 rounded-xl p-4 space-y-1 text-rose-900">
            <p className="font-bold">จะลบ {resetPreview.deleteCount} คาบ · {clearRoomId ? `เฉพาะ ${getClassroom(clearRoomId).name}` : 'ทุกห้องเรียน'}</p>
            <p>คาบ Auto {resetPreview.autoCount} คาบ · คาบลงด้วยมือ/รายการเดิม {resetPreview.manualCount} คาบ</p>
            <p className="text-slate-700">รักษาคาบที่กดล็อก 🔒 {resetPreview.pinnedCount} คาบ{clearRoomId ? ` · คาบห้องอื่น ${resetPreview.outOfScopeCount} คาบไม่ถูกลบ` : ''}</p>
          </div>
          <p className="text-slate-600">กิจกรรมล็อกตายตัว, PLC, เงื่อนไข/ข้อจำกัด, ภาระสอน และตารางฉบับประกาศใช้ <strong>ไม่ถูกลบ</strong> ระบบจะบันทึกประวัติก่อนลบ และให้ยืนยันอีกครั้ง</p>
          <div className="flex gap-3">
            <button type="button" onClick={()=>setClearModalOpen(false)} className="flex-1 py-3 bg-slate-100 text-slate-800 rounded-xl font-bold">ยกเลิก</button>
            <button type="button" disabled={!resetPreview.deleteCount || isScheduling || editingRows}
              onClick={()=>{setClearModalOpen(false);clearUnlockedSchedules(clearRoomId);}}
              className="flex-1 py-3 bg-rose-600 text-white rounded-xl font-bold disabled:opacity-50">ตรวจสอบและล้าง {resetPreview.deleteCount} คาบ</button>
          </div>
        </div>
      </Modal>
      <Modal isOpen={!!feasibility} onClose={() => setFeasibility(null)} title="วิเคราะห์ความเป็นไปได้ก่อนจัดตาราง">
        {feasibility && <div className="space-y-4 text-sm">
          <p className="text-slate-700">ภาระสอน {feasibility.totalRequired} คาบ · {feasibility.teacherCount} ครู · {feasibility.roomCount} ห้อง · {feasibility.activePeriods} คาบเรียนต่อวัน</p>
          {feasibility.invalidInputs.length > 0 ? <div className="p-3 bg-rose-50 border border-rose-200 rounded-lg"><strong>ข้อมูลต้องแก้ก่อนจัดตาราง</strong><ul className="list-disc pl-5 mt-2 space-y-1">{feasibility.invalidInputs.map((msg,i)=><li key={i}>{msg}</li>)}</ul></div> :
           feasibility.blocking.length > 0 ? <div className="p-3 bg-rose-50 border border-rose-200 rounded-lg"><strong>พบข้อจำกัดด้านความจุที่ทำให้จัดครบไม่ได้</strong><ul className="list-disc pl-5 mt-2 space-y-1">{feasibility.blocking.map((msg,i)=><li key={i}>{msg}</li>)}</ul></div> :
           <div className="p-3 bg-emerald-50 border border-emerald-200 rounded-lg font-medium text-emerald-800">ยังไม่พบข้อจำกัดด้านความจุที่ยืนยันว่าจัดไม่ครบ</div>}
          {feasibility.advisories.length > 0 && <div className="p-3 bg-amber-50 border border-amber-200 rounded-lg"><strong>ข้อสังเกตเพิ่มเติม</strong><ul className="list-disc pl-5 mt-2 space-y-1">{feasibility.advisories.map((msg,i)=><li key={i}>{msg}</li>)}</ul></div>}
          <p className="text-slate-500">การวิเคราะห์นี้เป็นการคัดกรองเบื้องต้น ไม่ใช่การทดลองแก้โจทย์ครบทุกความเป็นไปได้ และไม่รับประกันว่าการจัดตารางอัตโนมัติจะจัดได้ครบ</p>
          <button onClick={()=>setFeasibility(null)} className="w-full p-3 bg-[#081a39] text-white rounded-xl font-bold">ปิดรายงาน</button>
        </div>}
      </Modal>
      <Modal isOpen={isAddModalOpen} onClose={() => setIsAddModalOpen(false)} title="เพิ่มตารางสอน (รายคาบ)">
        <div className="space-y-4">
          <div className="grid grid-cols-2 gap-4"><div><label className="block text-sm font-bold mb-1">วัน</label><input value={addForm.day} disabled className="w-full p-3 border rounded-xl bg-slate-100" /></div><div><label className="block text-sm font-bold mb-1">คาบที่</label><input value={periods.find(p=>p.id===addForm.periodId)?.name} disabled className="w-full p-3 border rounded-xl bg-slate-100" /></div></div>
          <div><label className="block text-sm font-bold mb-1">ครูผู้สอน</label><select value={addForm.teacherId} onChange={e => setAddForm(prev => ({ ...prev, teacherId: e.target.value, subjectId: '' }))} disabled={viewMode === 'teacher'} className="w-full p-3 border rounded-xl bg-slate-50 disabled:bg-slate-100"><option value="">-- เลือกครูผู้สอน --</option>{teachers.map(t => <option key={t.id} value={t.id}>{t.name}</option>)}</select></div>
          <div><label className="block text-sm font-bold mb-1">รายวิชา</label><select value={addForm.subjectId} onChange={e => setAddForm({...addForm, subjectId: e.target.value})} disabled={!addForm.teacherId || selectableManualSubjects.length === 0} className="w-full p-3 border rounded-xl bg-slate-50 disabled:bg-slate-100"><option value="">{!addForm.teacherId ? '-- กรุณาเลือกครูผู้สอนก่อน --' : availableManualSubjects.length === 0 ? '-- ไม่พบรายวิชาในภาระสอน --' : selectableManualSubjects.length === 0 ? '-- รายวิชาของครู/ห้องนี้จัดครบแล้ว --' : '-- เลือกรายวิชา --'}</option>{availableManualSubjects.map(s => <option key={s.id} value={s.id} disabled={!!s.isComplete}>{s.optionLabel}</option>)}</select>{addForm.teacherId && availableManualSubjects.length === 0 && <p role="status" className="text-sm text-amber-700 mt-2">{addForm.classroomId ? 'ไม่พบภาระสอนสำหรับครูและห้องเรียนนี้' : 'ไม่พบภาระสอนสำหรับครูคนนี้'} กรุณาตรวจสอบเมนูภาระสอน</p>}{addForm.teacherId && addForm.classroomId && availableManualSubjects.length > 0 && selectableManualSubjects.length === 0 && <p role="status" className="text-sm text-emerald-700 mt-2">รายวิชาในภาระสอนของครูและห้องนี้จัดครบทุกคาบแล้ว</p>}</div>
          <div><label className="block text-sm font-bold mb-1">ห้องเรียน</label>
            <select value={addForm.classroomId} onChange={e => handleManualRoomChange(e.target.value)} disabled={viewMode === 'room'} className="w-full p-3 border rounded-xl bg-slate-50 disabled:bg-slate-100">
              <option value="">{addForm.teacherId && addForm.subjectId && manualRoomOptions.length === 0 ? '-- ไม่พบห้องที่มีภาระสอนรายวิชานี้ --' : '-- เลือกห้องเรียน --'}</option>
              {manualRoomOptions.map(room => <option key={room.id} value={room.id}>{room.optionLabel}</option>)}
            </select>
            {viewMode === 'teacher' && addForm.teacherId && addForm.subjectId && <p role="status" className="text-xs text-slate-600 mt-2">แสดงเฉพาะ {manualRoomOptions.length} ห้องที่มีภาระสอนของครูในรายวิชานี้ โดยคงรายวิชาเดิมเมื่อเปลี่ยนห้องที่สอนได้</p>}
            {addForm.teacherId && addForm.subjectId && manualRoomOptions.length === 0 && <p role="status" className="text-sm text-amber-700 mt-2">ไม่พบห้องเรียนที่มีภาระสอนรายวิชานี้ กรุณาตรวจสอบข้อมูลภาระสอน</p>}
            {addForm.teacherId && addForm.subjectId && addForm.classroomId && selectedManualRoom?.isComplete && <p role="status" className="text-sm text-amber-700 mt-2">ห้อง {selectedManualRoom.name} จัดรายวิชานี้ครบตามภาระสอนแล้ว ไม่สามารถเพิ่มคาบได้อีก</p>}
            {addForm.teacherId && addForm.subjectId && addForm.classroomId && !selectedManualRoom && <p role="status" className="text-sm text-rose-700 mt-2">ห้องนี้ไม่มีภาระสอนสำหรับครูและรายวิชาที่เลือก กรุณาเลือกห้องหรือรายวิชาใหม่</p>}
          </div>
          {addForm.classroomId && classroomUnavailabilities.some(u=>u.classroomId===addForm.classroomId&&u.day===addForm.day&&u.periodId===addForm.periodId) && <div className="bg-amber-50 border border-amber-300 text-amber-900 text-sm p-3 rounded-xl">
            ห้องนี้ถูกปิดไม่ให้ Auto จัดในคาบที่เลือก แต่สามารถลงด้วยมือเป็นกรณีพิเศษได้ ระบบจะขอให้ยืนยันก่อนบันทึก และยังตรวจครูชน / PLC / คาบซ้ำตามปกติ
          </div>}
          {addForm.teacherId&&addForm.subjectId&&addForm.classroomId&&<p className="bg-amber-50 text-amber-900 border border-amber-200 rounded-xl p-3 text-sm">หากการจัดด้วยมือทำให้ครูสอนรายวิชา 3 คาบติดกัน ระบบจะถามยืนยันก่อนบันทึก (Auto ยังหลีกเลี่ยงตามเดิม) ส่วนคาบครูชน PLC และข้อบังคับอื่นยังห้ามตามปกติ</p>}
          {addForm.teacherId&&addForm.subjectId&&addForm.classroomId&&<div className="border border-emerald-200 bg-emerald-50 rounded-xl p-3 space-y-2">
            <strong className="text-sm text-emerald-800">คาบที่ระบบแนะนำ (ไม่เกิน 4 คาบ/วัน)</strong>
            {manualSlotSuggestions.length?<div className="flex flex-wrap gap-2">{manualSlotSuggestions.map((slot,i)=><button
              key={`${slot.day}_${slot.periodId}`} type="button" onClick={()=>setAddForm(prev=>({...prev,day:slot.day,periodId:slot.periodId}))}
              className="px-3 py-2 bg-white rounded-lg border border-emerald-200 text-sm font-medium">{i+1}. {slot.day} {slot.periodName}{slot.block===2?' (คาบคู่)':''}</button>)}</div>:<p className="text-xs text-amber-800">ไม่พบช่องแนะนำตามข้อบังคับ 4 คาบ กรุณาตรวจเงื่อนไขหรือคาบที่ลงไว้</p>}
          </div>}
          <button onClick={handleManualAdd} disabled={!!(addForm.teacherId && addForm.subjectId && addForm.classroomId && (!selectedManualRoom || selectedManualRoom.isComplete))} className="w-full mt-2 bg-[#081a39] text-white py-3 rounded-xl font-bold disabled:opacity-50 disabled:cursor-not-allowed">บันทึกคาบเรียน</button>
        </div>
      </Modal>
    </div>
  );
};

// STAGE 5A: ประวัติตารางปัจจุบัน (อ่านอย่างเดียว ไม่กู้คืน/เขียนข้อมูล)
const ScheduleHistoryView = ({ db, appId, schoolSettings, getTermKey, getSubject, getTeacher, getClassroom, readScheduleVersion, restoreScheduleVersion }) => {
  const termKey = getTermKey(schoolSettings);
  const [versions, setVersions] = useState([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [comparison, setComparison] = useState(null);
  const [selectedId, setSelectedId] = useState('');
  const termPath = `school_data/${appId}/terms/${termKey}`;

  const refreshVersions = async () => {
    setLoading(true); setError(''); setComparison(null);setSelectedId('');
    try {
      const snap = await getDocsFromServer(collection(db, `${termPath}/scheduleVersions`));
      const records = snap.docs.map(d => ({ ...d.data(), id:d.id }))
        .sort((a,b) => (Number(b.createdAt)||0) - (Number(a.createdAt)||0));
      setVersions(records);
    } catch (e) {
      console.error('Read schedule history failed', e);
      setVersions([]);setError('อ่านประวัติจากเซิร์ฟเวอร์ไม่สำเร็จ กรุณาตรวจอินเทอร์เน็ตและสิทธิ์ Firebase');
    } finally { setLoading(false); }
  };
  useEffect(() => { refreshVersions(); }, [termKey]);

  const compareVersion = async record => {
    setLoading(true);setError('');setComparison(null);setSelectedId(record.id);
    try {
      const restored = await readScheduleVersion(termKey,record.id);
      const [currentSchedules,currentFixed,currentGroups,currentPLC] = await Promise.all([
        getDocsFromServer(collection(db, `${termPath}/schedules`)),
        getDocsFromServer(collection(db, `${termPath}/fixedSchedules`)),
        getDocsFromServer(collection(db, `${termPath}/activityGroups`)),
        getDocsFromServer(collection(db, `${termPath}/plcSessions`)),
      ]);
      const newSchedules = currentSchedules.docs.map(d => ({ ...d.data(), id:d.id }));
      const newFixed = currentFixed.docs.map(d => ({ ...d.data(), id:d.id }));
      const newGroups = currentGroups.docs.map(d => ({ ...d.data(), id:d.id }));
      const newPLC=currentPLC.docs.map(d=>({...d.data(),id:d.id}));
      setComparison({ record,
        fingerprints:{ schedules:fingerprintScheduleRows(newSchedules), fixedSchedules:fingerprintScheduleRows(newFixed),activityGroups:fingerprintScheduleRows(newGroups),plcSessions:fingerprintScheduleRows(newPLC) },
        regular:compareScheduleSnapshotRows(newSchedules,restored.schedules),
        fixed:compareScheduleSnapshotRows(newFixed,restored.fixedSchedules),
        groups:restored.includesActivityGroups ? compareScheduleSnapshotRows(newGroups,restored.activityGroups) : null,
        plc:restored.includesPLCSessions ? compareScheduleSnapshotRows(newPLC,restored.plcSessions) : null,
      });
    } catch (e) {
      console.error('Compare schedule history failed', e);
      setError(e.message === 'UNSUPPORTED_HISTORY_FORMAT'
        ? 'ไฟล์ประวัตินี้เป็นรูปแบบที่ยังไม่รองรับการเทียบ กรุณาอย่าใช้คืนค่า'
        : 'อ่านตารางล่าสุดจากเซิร์ฟเวอร์ไม่ได้ กรุณาลองใหม่');
    } finally { setLoading(false); }
  };
  const describeItem = item => item.subjectArea ? `${plcDisplayName(item.subjectArea)} / วัน${item.day || '-'} คาบ ${item.periodId || '-'}` : `${getSubject(item.subjectId)?.name || item.subjectId || '-'} / ${getClassroom(item.classroomId)?.name || item.classroomId || '-'} / ${getTeacher(item.teacherId)?.name || '-'} / วัน${item.day || '-'} คาบ ${item.periodId || '-'}`;
  const changeRows = diff => [
    ...diff.added.map(item=>({label:'เพิ่มในตารางปัจจุบัน',item})),
    ...diff.deleted.map(item=>({label:'มีในประวัติ แต่ไม่มีในปัจจุบัน',item})),
    ...diff.changed.map(pair=>({label:'เปลี่ยนรายละเอียด',item:pair.after,was:pair.before})),
  ];

  return <div className="space-y-5 animate-in fade-in">
    <div className="bg-white rounded-2xl shadow-sm border border-slate-100 p-6 flex flex-wrap items-center justify-between gap-3">
      <div><h2 className="text-xl font-bold text-[#081a39] flex gap-2 items-center"><Clock size={20}/> ประวัติและเปรียบเทียบตาราง</h2>
        <p className="text-sm text-slate-500 mt-1">ภาคเรียน {termKey} · ตรวจและเปรียบเทียบก่อนกู้คืน · รองรับประวัติรูปแบบเก่า</p>
      </div>
      <button disabled={loading} onClick={refreshVersions} className="px-4 py-2 rounded-xl bg-[#081a39] text-white font-bold disabled:opacity-50">โหลดประวัติใหม่</button>
    </div>
    {error && <p role="alert" className="bg-rose-50 text-rose-700 border border-rose-200 rounded-xl p-4">{error}</p>}
    <div className="bg-white border rounded-2xl p-5 space-y-3">
      <h3 className="font-bold">รายการเวอร์ชันที่บันทึกไว้ ({versions.length})</h3>
      {loading && <p className="text-slate-500">กำลังอ่านข้อมูลจาก Firebase...</p>}
      {!loading && versions.length===0 && !error && <p className="text-slate-500">ยังไม่พบประวัติที่บันทึกไว้ในภาคเรียนนี้</p>}
      {versions.map(v=><div key={v.id} className="flex flex-wrap items-center gap-2 justify-between border-b pb-3 last:border-0">
        <div><div className="font-bold">{Number.isFinite(Number(v.createdAt)) ? new Date(Number(v.createdAt)).toLocaleString('th-TH') : '-'} · {v.reason || 'ไม่มีเหตุผลกำกับ'}</div>
          <p className="text-xs text-slate-500">ตารางสอน {Number(v.scheduleCount)||0} คาบ / กิจกรรมล็อก {Number(v.fixedScheduleCount)||0} รายการ / กิจกรรมย่อย {v.formatVersion>=3 ? (Number(v.activityGroupCount)||0) : 'ไม่รวมในเวอร์ชันเก่า'}</p></div>
        <button disabled={loading} onClick={()=>compareVersion(v)} className="px-3 py-2 rounded-lg border border-slate-300 hover:bg-slate-50 disabled:opacity-50">เปรียบเทียบ</button>
      </div>)}
    </div>
    {comparison && selectedId===comparison.record.id && <div className="bg-white border rounded-2xl p-5 space-y-4">
      <h3 className="font-bold text-lg">เปรียบเทียบตารางปัจจุบันกับประวัติที่เลือก</h3>
      <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
        {[['ตารางสอน',comparison.regular],['กิจกรรมล็อก',comparison.fixed],...(comparison.groups?[['กิจกรรมกลุ่ม',comparison.groups]]:[]),...(comparison.plc?[['PLC กลุ่มสาระ',comparison.plc]]:[])].map(([name,diff])=><div key={name} className="border rounded-xl p-4">
          <div className="font-bold mb-1">{name}</div>
          <p className="text-sm text-slate-600">เพิ่มในปัจจุบัน {diff.added.length} · ไม่มีแล้ว {diff.deleted.length} · เปลี่ยนรายละเอียด {diff.changed.length} · เหมือนเดิม {diff.identical}</p>
          <ul className="list-disc pl-5 mt-2 max-h-56 overflow-y-auto text-xs text-slate-600 space-y-2">
            {changeRows(diff).slice(0,80).map((row,i)=><li key={i}><strong>{row.label}:</strong> {describeItem(row.item)}{row.was && <div className="text-slate-500">เดิม: {describeItem(row.was)}</div>}</li>)}
          </ul>
          {changeRows(diff).length>80 && <p className="text-amber-700 text-xs mt-2">แสดงรายละเอียดสูงสุด 80 รายการ แต่ยอดรวมแสดงครบ</p>}
        </div>)}
      </div>
      {!comparison.groups && <p className="text-sm text-amber-700">เวอร์ชันเก่าไม่เก็บกิจกรรมกลุ่ม ระบบจะคงกิจกรรมกลุ่มปัจจุบันไว้และตรวจความเข้ากันได้ก่อนกู้คืน</p>}
      {!comparison.plc && <p className="text-sm text-amber-700">เวอร์ชันเก่าไม่เก็บ PLC ระบบจะรักษา PLC ปัจจุบันไว้เมื่อกู้คืน</p>}
      <p className="text-sm text-slate-500">การกู้คืนจะตรวจความขัดแย้งกับข้อมูลปัจจุบันและสร้างจุดย้อนกลับก่อนเสมอ</p>
      <button disabled={loading} onClick={()=>restoreScheduleVersion(comparison.record.id,comparison.fingerprints)} className="w-full px-4 py-3 bg-amber-600 text-white rounded-xl font-bold disabled:opacity-50">กู้คืนตารางเวอร์ชันที่เลือก (ต้องยืนยัน)</button>
    </div>}
  </div>;
};

const ReportsView = ({ health:draftHealth, teachers:draftTeachers, subjects:draftSubjects,
  classrooms:draftClassrooms, periods:draftPeriods, schoolSettings:draftSettings,
  teachingLoads, unavailabilities, classroomUnavailabilities, customRules,
  schedules:draftSchedules, fixedSchedules:draftFixedSchedules, activityGroups:draftGroups, plcSessions,
  publishedSnapshot, getSubject:lookupSubject, getTeacher:lookupTeacher, getClassroom:lookupClassroom }) => {
  const [activeSubTab, setActiveSubTab] = useState('health');
  const [reportMode,setReportMode]=useState('draft');
  const showingPublished=reportMode==='published'&&!!publishedSnapshot;
  const schedules=showingPublished?publishedSnapshot.schedules:draftSchedules;
  const fixedSchedules=showingPublished?publishedSnapshot.fixedSchedules:draftFixedSchedules;
  const activityGroups=showingPublished?publishedSnapshot.activityGroups:draftGroups;
  const displayedPLC=showingPublished?(publishedSnapshot.referenceData?.plcSessions||[]):(plcSessions||[]);
  const referenceData=showingPublished?publishedSnapshot.referenceData:null;
  const teachers=referenceData?.teachers||draftTeachers;
  const subjects=referenceData?.subjects||draftSubjects;
  const classrooms=referenceData?.classrooms||draftClassrooms;
  const periods=referenceData?.periods||draftPeriods;
  const schoolSettings=referenceData?.schoolSettings||draftSettings;
  const publishedTeacherMap=useMemo(()=>new Map((referenceData?.teachers||[]).map(x=>[x.id,x])),[referenceData]);
  const publishedSubjectMap=useMemo(()=>new Map((referenceData?.subjects||[]).map(x=>[x.id,x])),[referenceData]);
  const publishedClassMap=useMemo(()=>new Map((referenceData?.classrooms||[]).map(x=>[x.id,x])),[referenceData]);
  const getTeacher=id=>publishedTeacherMap.get(id)||lookupTeacher(id);
  const getSubject=id=>publishedSubjectMap.get(id)||lookupSubject(id);
  const getClassroom=id=>publishedClassMap.get(id)||lookupClassroom(id);
  const health=useMemo(()=>{
    if(!showingPublished)return draftHealth;
    const context=referenceData;
    return calculateScheduleHealth(schedules, context?.teachingLoads||teachingLoads,
      context?.unavailabilities||unavailabilities,
      context?.classroomUnavailabilities||classroomUnavailabilities,
      fixedSchedules,teachers,subjects,classrooms,periods,context?.scheduleRules||customRules,activityGroups, showingPublished?(context?.plcSessions||[]):plcSessions);
  },[showingPublished,referenceData,schedules,fixedSchedules,activityGroups,teachers,subjects,classrooms,periods,
      teachingLoads,unavailabilities,classroomUnavailabilities,customRules,plcSessions,draftHealth]);
  const [filterLvl, setFilterLvl] = useState('all');

  const filteredMissing = useMemo(() => {
    let list = [...health.missingLoads];
    if (filterLvl !== 'all') list = list.filter(m => getClassroom(m.classroomId).name.includes(filterLvl));
    return list.sort((a,b) => b.missing - a.missing);
  }, [health.missingLoads, filterLvl]);

  const exportFormattedExcel = () => {
    try {
      const wb=XLSX.utils.book_new();const active=periods.filter(p=>!p.isBreak);
      const usedNames=new Set();
      const append=(name,rows)=>{
        let safe=String(name).replace(/[\\/*?:\[\]]/g,'-').slice(0,30)||'Sheet';
        const base=safe;let i=2;
        while(usedNames.has(safe))safe=base.slice(0,26)+'-'+i++;
        usedNames.add(safe);
        const ws=XLSX.utils.aoa_to_sheet(rows);
        ws['!cols']=[{wch:17},...active.map(()=>({wch:28}))];
        ws['!merges']=[{s:{r:0,c:0},e:{r:0,c:active.length}}];
        ws['!freeze']={xSplit:1,ySplit:3};
        ws['!autofilter']={ref:`A3:${XLSX.utils.encode_col(active.length)}${rows.length}`};
        XLSX.utils.book_append_sheet(wb,ws,safe);
      };
      const summary=[['BHS Scheduler — ตาราง'+(showingPublished?'ฉบับประกาศใช้':'ฉบับร่าง')],
        ['ภาคเรียน',schoolSettings.semester,'ปีการศึกษา',schoolSettings.academicYear],
        ['ประเภท','จำนวน'],['ห้องเรียน',classrooms.length],['ครู',teachers.length],
        ['คาบสอนในตาราง',schedules.length],['กิจกรรมล็อก',fixedSchedules.length]];
      const ss=XLSX.utils.aoa_to_sheet(summary);ss['!cols']=[{wch:35},{wch:22},{wch:22},{wch:18}];
      XLSX.utils.book_append_sheet(wb,ss,'สรุป');usedNames.add('สรุป');
      const title=(item,mode)=>mode==='room'?`ตารางเรียน ${item.name}`:`ตารางสอน ${item.name}`;
      const value=(item,day,p,mode)=>{
        const regular=schedules.filter(x=>x.day===day&&x.periodId===p.id&& (mode==='room'?x.classroomId===item.id:x.teacherId===item.id));
        const fixed=fixedSchedules.filter(x=>x.day===day&&x.periodId===p.id&& (mode==='room'?fixedAppliesToClassroom(x,item.id):x.teacherId===item.id));
        const groups=mode==='teacher'?activityTeacherEntries(activityGroups,fixedSchedules).filter(x=>x.teacherId===item.id&&x.day===day&&x.periodId===p.id):[];
        const plc=mode==='teacher'?buildPLCReservations(displayedPLC,teachers).filter(x=>x.teacherId===item.id&&x.day===day&&x.periodId===p.id):[];
        return [...regular,...fixed,...groups,...plc].map(x=>{
          if(x.__isPLC)return sanitizeExcelInput(plcDisplayName(x.subjectArea));
          const group=mode==='room'?activityGroups.find(g=>g.fixedId===x.id&&g.classroomId===item.id):
            activityGroups.find(g=>'activity_'+g.id===x.id);
          const isActivity=fixed.includes(x)||groups.includes(x);
          const subject=isActivity
            ? (x.__activityName||group?.name||fixedActivityDisplayName(x,getSubject,activityGroups,mode==='room'?item.id:''))
            : (isDevelopmentActivity(x)?fixedActivityDisplayName(x,getSubject):getSubject(x.subjectId).name);
          if(isActivity||isDevelopmentActivity(x)) return sanitizeExcelInput(subject);
          const other=mode==='room'?(x.teacherId?getTeacher(x.teacherId).name:(fixedActivityTeacherIds(x).length?`ครู ${fixedActivityTeacherIds(x).length} คน`:'ไม่ระบุครู')):
            (x.scopeLabel || getClassroom(x.classroomId).name);
          return sanitizeExcelInput(subject+' — '+other);
        }).join(' / ');
      };
      const build=(item,mode)=>{
        const rows=[[title(item,mode)+' · '+(showingPublished?'ฉบับประกาศใช้':'ฉบับร่าง')],
          ['ภาคเรียน '+schoolSettings.semester+' ปี '+schoolSettings.academicYear],
          ['วัน / คาบ',...active.map(p=>p.name)]];
        DAYS.forEach(day=>rows.push([day,...active.map(p=>value(item,day,p,mode))]));
        return rows;
      };
      classrooms.forEach(c=>append('ห้อง-'+c.name,build(c,'room')));
      teachers.forEach(t=>append('ครู-'+t.name,build(t,'teacher')));
      XLSX.writeFile(wb,`BHS-Timetable-${showingPublished?'Published':'Draft'}-${schoolSettings.academicYear}-T${schoolSettings.semester}.xlsx`);
    }catch(error){console.error(error);window.alert(translateUserMessage('สร้าง Excel ไม่สำเร็จ: '+error.message));}
  };

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
               row.push(combined.map(s =>
                 (s.classroomId === ALL_CLASSROOMS_ID
                   ? activityGroups.find(g=>g.fixedId===s.id&&g.classroomId===c.id)?.name : '') ||
                   fixedActivityDisplayName(s,getSubject,activityGroups,c.id)).join(' + '));
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

      <div className="bg-blue-50 p-3 rounded-lg text-xs text-blue-900">ผลตรวจสุขภาพ ภาระที่ขาด และรายงานทั้งหมดใช้ข้อมูลฉบับที่เลือกด้านล่างตรงกัน
        {showingPublished&&!referenceData&&<strong className="block text-amber-800 mt-1">ฉบับประกาศเก่าไม่มีข้อมูลอ้างอิง ณ วันที่ประกาศ ผลตรวจและชื่อครู/วิชาอาจอิงข้อมูลปัจจุบัน โปรดประกาศใหม่เมื่อตรวจร่างครบแล้ว</strong>}
      </div>
      {<div className="bg-white border rounded-xl p-3 flex gap-3 items-center flex-wrap">
        <label className="font-bold text-sm">แสดงตารางจาก</label>
        <select className="border rounded-lg p-2" value={reportMode} onChange={e=>setReportMode(e.target.value)}>
          <option value="draft">ฉบับร่าง (ล่าสุด)</option>
          {publishedSnapshot&&<option value="published">ฉบับประกาศใช้</option>}
        </select><span className={`text-xs font-semibold px-2 py-1 rounded-full ${showingPublished?'bg-emerald-100 text-emerald-800':'bg-amber-100 text-amber-800'}`}>{showingPublished?'ประกาศใช้แล้ว':'ฉบับร่าง'}</span><span className="text-xs text-slate-500">{showingPublished?`ประกาศเมื่อ ${new Date(publishedSnapshot.publishedAt).toLocaleString('th-TH')} · ตารางที่ประกาศแล้วไม่เปลี่ยนเมื่อแก้ฉบับร่าง`:'ตารางฉบับร่างที่กำลังแก้ไข'}</span>
      </div>}
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
              {health.customRuleConflicts.map((c,i) => <li key={`rule-${i}`} className="p-2 bg-rose-50 rounded border border-rose-100">{c}</li>)}
              {health.subjectDailyConflicts.map((c,i) => <li key={`daily-${i}`} className="p-2 bg-rose-50 rounded border border-rose-100">{c}</li>)}
              {(health.plcConflicts||[]).map((c,i)=><li key={`plc-${i}`} className="p-2 bg-rose-50 rounded border border-rose-200">PLC: {c}</li>)}
              {health.orphans.map((c,i) => <li key={`oc-${i}`} className="p-2 bg-slate-100 rounded border border-slate-300">ข้อมูลสูญหาย: {c}</li>)}
              {health.status === 'empty' && <li className="text-slate-600 p-2 bg-slate-50 rounded border border-slate-200">ยังไม่มีภาระสอน กรุณานำเข้าหรือเพิ่มข้อมูลก่อนจัดตาราง</li>}
              {health.status === 'ready' && <li className="text-emerald-600 font-bold p-2 bg-emerald-50 rounded border border-emerald-100">ตรวจสอบผ่าน: ไม่พบข้อผิดพลาดร้ายแรง</li>}
            </ul>
          </div>
          <div className="bg-white rounded-2xl shadow-sm border border-amber-100 p-6">
            <h3 className="font-bold text-amber-700 mb-4 flex items-center gap-2"><AlertCircle size={18}/> ข้อควรระวัง (Warnings & Missing)</h3>
            <ul className="space-y-2 text-sm text-slate-700">
              {health.overloads.map((c,i) => <li key={`ov-${i}`} className="p-2 bg-amber-50 rounded border border-amber-100">{c}</li>)}
              {(health.teacherFivePeriodWarnings || []).map((c,i) => <li key={`five-${i}`} className="p-2 bg-amber-50 rounded border border-amber-100">{c}</li>)}
              {(health.teacherActivityConsecutiveWarnings || []).map((c,i) => <li key={`activity-cont-${i}`} className="p-2 bg-amber-50 rounded border border-amber-100">{c}</li>)}
              {(health.teacherManualConsecutiveWarnings || []).map((c,i) => <li key={`manual-three-${i}`} className="p-2 bg-amber-50 rounded border border-amber-100">{c}</li>)}
              {health.missing > 0 && <li className="p-2 bg-amber-50 rounded border border-amber-100">มีภาระสอนที่ยังจัดตารางไม่ครบรวม {health.missing} คาบ</li>}
              {(!health.overloads.length && health.missing === 0 && !(health.teacherFivePeriodWarnings||[]).length && !(health.teacherActivityConsecutiveWarnings||[]).length && !(health.teacherManualConsecutiveWarnings||[]).length) && <li className="text-emerald-600 font-bold p-2 bg-emerald-50 rounded border border-emerald-100">ตรวจสอบผ่าน: จัดคาบครบและไม่พบคำเตือนเรื่องภาระครู</li>}
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
            <thead className="bg-[#081a39] text-white"><tr><th className="p-3">ชื่อครู</th><th className="p-3 text-center">ภาระสอน</th><th className="p-3 text-center">จัดสอน</th><th className="p-3 text-center">กิจกรรม/คาบล็อก</th><th className="p-3 text-center">PLC</th><th className="p-3 text-center">รวมภาระจริง</th><th className="p-3 text-center">วันที่ภาระรวม 4–5 คาบ</th><th className="p-3 text-center">สถานะ</th><th className="p-3 text-center">Gap</th><th className="p-3 text-center">ติดกันสูงสุด</th><th className="p-3 text-center">การกระจาย จ–ศ</th></tr></thead>
            <tbody>
              {teachers.map((t, i) => {
                const st = health.teacherStats[t.id] || { req: 0, sch: 0, gaps: 0, maxConsec: 0 };
                return (
                  <tr key={i} className="border-b hover:bg-slate-50">
                    <td className="p-3 font-bold">{t.name}</td><td className="p-3 text-center">{st.req}</td><td className="p-3 text-center font-bold">{st.regularPeriods||0}</td>
                    <td className="p-3 text-center">{st.activityPeriods||0}</td>
                    <td className="p-3 text-center">{buildPLCReservations(displayedPLC,teachers).filter(x=>x.teacherId===t.id).length}</td>
                    <td className="p-3 text-center font-bold">{st.totalWeeklyPeriods||0}</td>
                    <td className="p-3 text-center">{st.heavyDays||0}{st.fiveDays?.length ? ` (5 คาบ: ${st.fiveDays.join(', ')})` : ''}</td>
                    <td className="p-3 text-center">{st.sch < st.req ? <span className="text-amber-600 bg-amber-50 px-2 py-1 rounded text-xs">ขาด {st.req - st.sch}</span> : st.sch > st.req ? <span className="text-rose-600 bg-rose-50 px-2 py-1 rounded text-xs">เกิน {st.sch - st.req}</span> : <span className="text-emerald-600 text-xs">ครบ</span>}</td>
                    <td className="p-3 text-center">{st.gaps}</td><td className="p-3 text-center">{st.maxConsec}</td><td className="p-3 text-center text-xs">{DAYS.map(d=>st.dailyCounts?.[d]||0).join(' / ')}</td>
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
             ตารางรวมโรงเรียน (โหมดห้องเรียน) <div className="flex gap-2"><button onClick={exportFormattedExcel} className="bg-indigo-700 text-white px-3 py-1 rounded">Export Excel (รายห้อง+รายครู)</button><button onClick={() => handleExportCSV('all')} className="bg-emerald-600 text-white px-3 py-1 rounded">Export CSV</button></div>
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
                         return <td key={p.id} className="p-1 border text-center h-12">{combined.map(s => <div key={s.id} className="font-bold truncate">{(s.classroomId===ALL_CLASSROOMS_ID ? activityGroups.find(g=>g.fixedId===s.id&&g.classroomId===c.id)?.name : '') || (isDevelopmentActivity(s) ? fixedActivityDisplayName(s,getSubject,activityGroups,c.id) : getSubject(s.subjectId).name)}</div>)}</td>
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

const PrintView = ({ activityGroups:draftGroups, schedules:draftSchedules,
  classrooms:draftClassrooms, teachers:draftTeachers, periods:draftPeriods,
  schoolSettings:draftSettings, plcSessions:draftPLC, publishedSnapshot, publishedLoading, publishedError,
  getSubject:lookupSubject, getShortTeacherName, getClassroom:lookupClassroom,
  getTeacher:lookupTeacher, fixedSchedules:draftFixed }) => {
  const [viewMode, setViewMode] = useState('room');
  const [printMode,setPrintMode]=useState('draft');
  const showingPublished=printMode==='published'&&!!publishedSnapshot;
  const schedules=showingPublished?publishedSnapshot.schedules:draftSchedules;
  const fixedSchedules=showingPublished?publishedSnapshot.fixedSchedules:draftFixed;
  const activityGroups=showingPublished?publishedSnapshot.activityGroups:draftGroups;
  const references=showingPublished?publishedSnapshot.referenceData:null;
  const printPLC=showingPublished?(references?.plcSessions||[]):(draftPLC||[]);
  const classrooms=references?.classrooms||draftClassrooms;
  const teachers=references?.teachers||draftTeachers;
  const periods=references?.periods||draftPeriods;
  const schoolSettings=references?.schoolSettings||draftSettings;
  const printVersionCode=useMemo(()=>buildTimetablePrintCode({
    academicYear:schoolSettings?.academicYear, semester:schoolSettings?.semester,
    versionId:showingPublished?publishedSnapshot?.versionId:null,
    schedules,fixedSchedules,activityGroups
  }),[schoolSettings?.academicYear,schoolSettings?.semester,showingPublished,publishedSnapshot?.versionId,schedules,fixedSchedules,activityGroups]);
  const publishedAtText=showingPublished&&publishedSnapshot?.publishedAt
    ?new Date(publishedSnapshot.publishedAt).toLocaleString('th-TH'):'-';
  const getSubject=id=>references?.subjects?.find(x=>x.id===id)||lookupSubject(id);
  const getTeacher=id=>teachers.find(x=>x.id===id)||lookupTeacher(id);
  const getClassroom=id=>id===ALL_CLASSROOMS_ID?{id,name:'ทุกห้องเรียน'}:(classrooms.find(x=>x.id===id)||lookupClassroom(id));
  // If publication becomes available, print the official copy by default.
  useEffect(()=>{setPrintMode(publishedSnapshot?'published':'draft');},[publishedSnapshot?.versionId]);
  const [selectedId, setSelectedId] = useState('');
  const handlePrint=()=>{
    if (!selectedId || publishedLoading) return;
    // A confirmation is mandatory for the Print button when printing a draft.
    if (!showingPublished && !window.confirm('คุณกำลังพิมพ์ตารางฉบับร่าง ซึ่งยังไม่ใช่ตารางทางการของโรงเรียน\n\nต้องการพิมพ์ฉบับร่างต่อหรือไม่?')) return;
    window.print();
  };

  const getCellContent = (targetId, day, periodId, vMode) => {
    const cellSchedules = schedules.filter(s => vMode === 'room' ? s.classroomId === targetId && s.day === day && s.periodId === periodId : s.teacherId === targetId && s.day === day && s.periodId === periodId);
    const cellFixed = fixedSchedules.filter(f => vMode === 'room' ? fixedAppliesToClassroom(f, targetId) && f.day === day && f.periodId === periodId : f.teacherId === targetId && f.day === day && f.periodId === periodId);
    const groupTeacherRows = vMode==='teacher' ? activityTeacherEntries(activityGroups,fixedSchedules)
      .filter(item=>item.teacherId===targetId&&item.day===day&&item.periodId===periodId)
      .map(item=>({...item,groupLabel:item.__activityName || fixedActivityDisplayName(item,getSubject)})) : [];
    const plcRows=vMode==='teacher'?buildPLCReservations(printPLC,teachers)
      .filter(x=>x.teacherId===targetId&&x.day===day&&x.periodId===periodId):[];
    const combined = [...cellFixed, ...cellSchedules,...groupTeacherRows,...plcRows];
    if (combined.length === 0) return null;
    return (
      <div className="flex flex-col items-center justify-center text-[11px] text-center w-full h-full">
        {combined.map((sch, idx) => (
          sch.__isPLC ? <div key={sch.id} className="font-bold text-[10px] leading-tight break-words">{plcDisplayName(sch.subjectArea)}</div> : <div key={sch.id} className={`w-full ${idx > 0 ? 'mt-1 pt-1 border-t border-black/20' : ''}`}>
            {(cellFixed.includes(sch) || sch.__activityTeacherOnly || isDevelopmentActivity(sch)) ? <div className="font-bold text-[11px] leading-tight break-words">{sch.groupLabel || (vMode==='room' ? fixedActivityDisplayName(sch,getSubject,activityGroups,targetId) : fixedActivityDisplayName(sch,getSubject))}</div> : <>
            <div className={`font-bold truncate text-[11px] ${cellFixed.includes(sch) ? 'text-amber-800' : ''}`}>{sch.subjectId}</div>
            {vMode === 'room' ? (
               <><div className="truncate text-[11px]">{sch.teacherId ? getShortTeacherName(getTeacher(sch.teacherId).name) : (fixedActivityTeacherIds(sch).length ? `ครู ${fixedActivityTeacherIds(sch).length} คน` : '-')}</div><div className="text-[10px] truncate">{sch.groupLabel || (vMode==='room' && sch.classroomId===ALL_CLASSROOMS_ID ? activityGroups.find(g=>g.fixedId===sch.id&&g.classroomId===targetId)?.name : '') || getSubject(sch.subjectId).name}</div></>
             ) : (
               <><div className="truncate text-[11px]">{sch.scopeLabel || getClassroom(sch.classroomId).name}</div><div className="text-[10px] truncate">{sch.groupLabel || (vMode==='room' && sch.classroomId===ALL_CLASSROOMS_ID ? activityGroups.find(g=>g.fixedId===sch.id&&g.classroomId===targetId)?.name : '') || getSubject(sch.subjectId).name}</div></>
             )}
            </>}
          </div>
        ))}
      </div>
    );
  };

  const PrintPage = ({ id, mode }) => (
    <div className="bg-white p-10 rounded-xl border border-slate-200 text-black shadow-sm print-page mb-6">
      <div className="text-center mb-8 print-header">
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
      <div className="flex justify-between px-28 mt-24 text-sm font-medium print-signatures">
        <div className="text-center"><p>ลงชื่อ........................................................</p><p className="mt-3">({String(schoolSettings?.deputyName || '')})</p><p className="mt-1">รองผู้อำนวยการ</p></div>
        <div className="text-center"><p>ลงชื่อ........................................................</p><p className="mt-3">({String(schoolSettings?.directorName || '')})</p><p className="mt-1">ผู้อำนวยการโรงเรียน</p></div>
      </div>
    </div>
  );

  return (
    <div className="animate-in fade-in bhs-print-view">
      <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-100 mb-6 flex flex-wrap items-center gap-4 no-print">
        <div className="p-3 bg-[#081a39] rounded-xl shadow-md"><Printer className="text-[#d4af37]" size={28} /></div>
        <select value={viewMode} onChange={e => {setViewMode(e.target.value); setSelectedId('');}} className="p-3 border border-slate-200 rounded-xl bg-slate-50 font-bold text-[#081a39]"><option value="room">พิมพ์ตารางเรียน (รายห้อง)</option><option value="teacher">พิมพ์ตารางสอน (รายครู)</option></select>
        <select value={selectedId} onChange={e => setSelectedId(e.target.value)} className="p-3 border border-slate-200 rounded-xl min-w-[250px] bg-slate-50 font-medium">
          <option value="">-- เลือกเพื่อดูตัวอย่างพิมพ์ --</option><option value="ALL_PRINT" className="font-bold text-[#081a39]">-- พิมพ์ทั้งหมด --</option>
          {viewMode === 'room' ? classrooms.map(c => <option key={c.id} value={c.id}>{c.name}</option>) : teachers.map(t => <option key={t.id} value={t.id}>{t.name}</option>)}
        </select>
        <select value={printMode} onChange={e=>{setPrintMode(e.target.value);setSelectedId('');}} className="border rounded-xl p-3 font-bold">
          <option value="draft">พิมพ์ฉบับร่าง</option>{publishedSnapshot&&<option value="published">พิมพ์ฉบับประกาศใช้</option>}
        </select>
        <button onClick={handlePrint} disabled={!selectedId||publishedLoading} className="px-8 py-3 bg-[#081a39] text-white rounded-xl shadow-md font-bold disabled:opacity-50 ml-auto hover:bg-[#122b5e]">พิมพ์เอกสาร</button>
      </div>
      <div className="no-print flex flex-wrap items-center gap-3 mb-4 rounded-xl border border-slate-200 bg-white p-3 text-sm">
        <span className={`font-bold rounded-full px-3 py-1 ${showingPublished?'bg-emerald-100 text-emerald-800':'bg-amber-100 text-amber-800'}`}>
          {showingPublished?'ฉบับประกาศใช้แล้ว':'ฉบับร่าง (ยังไม่ประกาศ)'}
        </span>
        <span className="text-slate-700">{showingPublished?`วันที่ประกาศ: ${publishedAtText}`:'ยังแก้ไขได้ และต้องยืนยันก่อนพิมพ์'}</span>
      </div>
      {publishedLoading&&<p className="no-print mb-4 text-xs text-blue-800 bg-blue-50 p-3 rounded-lg">กำลังตรวจสอบฉบับประกาศใช้ กรุณารอสักครู่ก่อนพิมพ์</p>}
      {publishedError&&<p className="no-print mb-4 text-xs text-rose-800 bg-rose-50 p-3 rounded-lg">{publishedError} · โปรดตรวจสอบสถานะก่อนเลือกพิมพ์</p>}
      {showingPublished&&!references&&<p className="no-print mb-4 text-xs text-amber-800 bg-amber-50 p-3 rounded-lg">หมายเหตุ: ฉบับประกาศเก่าไม่มีข้อมูลอ้างอิง ณ วันที่ประกาศ ชื่อครู/วิชาอาจแสดงตามข้อมูลปัจจุบัน</p>}
      <div id="print-area">{selectedId === 'ALL_PRINT' ? (viewMode === 'room' ? classrooms : teachers).map(item => <PrintPage key={item.id} id={item.id} mode={viewMode} />) : (selectedId && <PrintPage id={selectedId} mode={viewMode} />)}</div>
    </div>
  );
};

export default function App() {
  const [user, setUser] = useState(null);
  const [activeTab, setActiveTab] = useState('dashboard');
  const [constraintsEntrySection, setConstraintsEntrySection] = useState('time');
  const [dashboardTargetRoom, setDashboardTargetRoom] = useState('');
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
  const [customRules, setCustomRules] = useState([]);
  const [activityGroups,setActivityGroups] = useState([]);
  const [plcSessions,setPLCSessions] = useState([]);
  const [plcLoadedTerm,setPLCLoadedTerm] = useState('');
  const [rulesLoadedTerm, setRulesLoadedTerm] = useState('');
  const [groupsLoadedTerm,setGroupsLoadedTerm] = useState('');
  
  const [isScheduling, setIsScheduling] = useState(false);
  const [scheduleResult, setScheduleResult] = useState(null);
  const [publishedSnapshot,setPublishedSnapshot] = useState(null);
  const [publishedLoading,setPublishedLoading] = useState(false);
  const [publishedError,setPublishedError] = useState('');
  const [autoScheduleIssues, setAutoScheduleIssues] = useState([]);

  const getTermKey = (settings) => `${settings?.academicYear || '2569'}-${settings?.semester || '1'}`;
  const rulesReady = !!user && rulesLoadedTerm === getTermKey(schoolSettings) && groupsLoadedTerm === getTermKey(schoolSettings) && plcLoadedTerm === getTermKey(schoolSettings);

  const health = useMemo(() => {
    if (!periods || periods.length === 0) return { status: teachingLoads.length === 0 ? 'empty' : 'warning', missing: 0, scheduledForLoads: 0, totalRequired: 0, completionPercent: 0, teacherConflicts: [], roomConflicts: [], unavailConflicts: [], cUnavailConflicts: [], breakConflicts: [], orphans: [], missingLoads: [], overloads: [], blockConflicts: [], teacherLimitConflicts: [], teacherFivePeriodWarnings: [], teacherActivityConsecutiveWarnings: [], teacherManualConsecutiveWarnings: [], customRuleConflicts: [], subjectDailyConflicts: [], teacherStats: {} };
    return calculateScheduleHealth(schedules, teachingLoads, unavailabilities, classroomUnavailabilities, fixedSchedules, teachers, subjects, classrooms, periods, customRules, activityGroups, plcSessions);
  }, [schedules, teachingLoads, unavailabilities, classroomUnavailabilities, fixedSchedules, teachers, subjects, classrooms, periods, customRules, activityGroups, plcSessions]);

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
    setRulesLoadedTerm('');
    setGroupsLoadedTerm('');
    setPLCLoadedTerm('');
    let unsubscribers = [];
    const createTermListener = (col, setter) => onSnapshot(collection(db, `${termPath}/${col}`), snap => setter(snap.docs.map(doc => ({ ...doc.data(), id: doc.id }))));

    unsubscribers.push(createTermListener('teachingLoads', setTeachingLoads));
    unsubscribers.push(createTermListener('unavailabilities', setUnavailabilities));
    unsubscribers.push(createTermListener('classroomUnavailabilities', setClassroomUnavailabilities));
    unsubscribers.push(createTermListener('fixedSchedules', setFixedSchedules));
    unsubscribers.push(createTermListener('schedules', setSchedules));
    unsubscribers.push(onSnapshot(collection(db,`${termPath}/plcSessions`), snap=>{
      setPLCSessions(snap.docs.map(d=>({...d.data(),id:d.id})));setPLCLoadedTerm(termKey);
    },error=>{console.error('Load PLC failed',error);setPLCLoadedTerm('');}));
    unsubscribers.push(onSnapshot(collection(db,`${termPath}/activityGroups`), snap=>{
      setActivityGroups(snap.docs.map(d=>({...d.data(),id:d.id})));
      setGroupsLoadedTerm(termKey);
    },error=>{console.error('Load activity groups failed',error);setGroupsLoadedTerm('');}));
    unsubscribers.push(onSnapshot(collection(db, `${termPath}/scheduleRules`), snap => {
      setCustomRules(snap.docs.map(item => ({ ...item.data(), id: item.id })));
      setRulesLoadedTerm(termKey);
    }, error => { console.error('Load schedule rules failed', error); setRulesLoadedTerm(''); }));

    return () => unsubscribers.forEach(unsub => unsub && unsub());
  }, [user, schoolSettings.academicYear, schoolSettings.semester]);

  const showToast = (msg, type = 'success') => { setToast({ msg: type === 'error' ? translateUserMessage(msg) : String(msg), type }); setTimeout(() => setToast(null), type === 'error' ? 7000 : 3500); };

  const dbAction = async (collectionName, id, data, action = 'set') => {
    if (!user) return false;
    const isTermData = ['teachingLoads', 'schedules', 'unavailabilities', 'classroomUnavailabilities', 'fixedSchedules', 'scheduleRules', 'activityGroups', 'plcSessions'].includes(collectionName);
    const basePath = isTermData ? `school_data/${appId}/terms/${getTermKey(schoolSettings)}` : `school_data/${appId}`;
    try {
      if (action === 'set') await setDoc(doc(db, `${basePath}/${collectionName}`, id), { ...(data || {}), id });
      else if (action === 'delete') await deleteDoc(doc(db, `${basePath}/${collectionName}`, id));
      return true;
    } catch (e) { showToast("บันทึกไม่สำเร็จ", 'error'); return false; }
  };

  // Stage 22: save multiple subject time-window rules as one atomic operation.
  // Recheck server-side rules and existing lessons before saving; existing subject settings stay untouched.
  const saveSubjectTimeWindows = async (subjectIds, timeWindow, strictness) => {
    if (!user || !rulesReady) { showToast('ยังไม่พร้อมบันทึกเงื่อนไข กรุณารอโหลดข้อมูลก่อน', 'error'); return false; }
    const ids = [...new Set((subjectIds || []).map(String))];
    if (!ids.length || !['morning','afternoon'].includes(timeWindow) ||
        !['required','preferred'].includes(strictness) ||
        ids.some(id => !activeSubjects.some(s => String(s.id) === id))) {
      showToast('กรุณาเลือกรายวิชาและเงื่อนไขให้ถูกต้อง', 'error'); return false;
    }
    const termPath = `school_data/${appId}/terms/${getTermKey(schoolSettings)}`;
    const rulesPath = `${termPath}/scheduleRules`;
    try {
      const [ruleSnapshot, scheduleSnapshot, fixedSnapshot] = await Promise.all([
        getDocsFromServer(collection(db, rulesPath)),
        getDocsFromServer(collection(db, `${termPath}/schedules`)),
        getDocsFromServer(collection(db, `${termPath}/fixedSchedules`)),
      ]);
      const liveRules = ruleSnapshot.docs.map(d => d.data());
      if (liveRules.some(r => r.mode === 'timeWindow' && ids.includes(String(r.targetId)))) {
        showToast('บางวิชามีเงื่อนไขช่วงเวลาแล้ว กรุณาลบรายการเดิมก่อน', 'error'); return false;
      }
      const newRules = ids.map(targetId => ({
        id: `TW_${encodeURIComponent(targetId)}`, mode: 'timeWindow',
        targetType: 'subject', targetId, timeWindow, strictness,
      }));
      if (strictness === 'required') {
        const liveItems = [...scheduleSnapshot.docs, ...fixedSnapshot.docs].map(d => d.data());
        const conflicts = liveItems.filter(item => blocksCustomRule(newRules, item, periods));
        if (conflicts.length) {
          showToast(`มีตารางเดิม ${conflicts.length} คาบอยู่นอกช่วงเวลาที่จะบังคับ กรุณาปรับตารางก่อน`, 'error');
          return false;
        }
      }
      const refs = newRules.map(rule => doc(db, rulesPath, rule.id));
      await runTransaction(db, async tx => {
        const existing = await Promise.all(refs.map(ref => tx.get(ref)));
        if (existing.some(snap => snap.exists())) throw new Error('TIME_RULE_EXISTS');
        newRules.forEach((rule, i) => tx.set(refs[i], rule));
      });
      showToast(`เพิ่มเงื่อนไขช่วงเวลา ${newRules.length} รายวิชาสำเร็จ`);
      return true;
    } catch (error) {
      console.error('Save subject time windows failed', error);
      showToast(error?.message === 'TIME_RULE_EXISTS' ? 'มีเงื่อนไขของวิชานี้แล้ว กรุณาตรวจสอบอีกครั้ง' :
        'บันทึกเงื่อนไขรายวิชาไม่สำเร็จ ไม่มีการบันทึกบางส่วน', 'error');
      return false;
    }
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

  // STAGE 28: Bulk-clear unpinned draft lessons, preserving fixed activities,
  // PLC and the separately published version. All deletes use ONE atomic batch.
  const clearUnlockedSchedules = (targetRoomId = '') => {
    if (!user || !rulesReady || isScheduling) return showToast('กรุณารอให้ระบบโหลดข้อมูลหรือจัดตารางเสร็จก่อน', 'error');
    const termKey = getTermKey(schoolSettings);
    if (targetRoomId && !classrooms.some(c=>String(c.id)===String(targetRoomId))) return showToast('ไม่พบห้องเรียนที่เลือก กรุณาโหลดหน้าจอใหม่', 'error');
    let preview;
    try { preview = planUnlockedScheduleClear(schedules, targetRoomId); }
    catch { return showToast('ข้อมูลตารางไม่ถูกต้อง กรุณาโหลดหน้าจอใหม่ก่อน', 'error'); }
    if (!preview.deleteCount) return showToast('ไม่มีคาบที่ไม่ได้ล็อกให้ล้าง');
    const expectedFingerprint = fingerprintScheduleRows(schedules);
    const mode = targetRoomId ? `เฉพาะห้อง ${getClassroom(targetRoomId).name}` : 'ทุกห้องเรียน';
    setConfirmData({
      type: 'delete', title: 'ยืนยันล้างคาบไม่ล็อก', confirmText: `ล้าง ${preview.deleteCount} คาบ`,
      message: `ภาคเรียน ${termKey} · ${mode}\n\nจะลบคาบที่ไม่ล็อกทั้งหมด ${preview.deleteCount} คาบ (Auto ${preview.autoCount} / ลงด้วยมือ ${preview.manualCount})\nรักษาคาบที่กดล็อก 🔒 ${preview.pinnedCount} คาบ${targetRoomId ? ` และคาบของห้องอื่น ${preview.outOfScopeCount} คาบ` : ''}\nรักษากิจกรรมล็อกตายตัว / PLC / ข้อจำกัดทั้งหมด\n\nระบบบันทึกประวัติไว้ก่อนลบ สามารถกู้คืนจากเมนูประวัติตารางได้\nตารางฉบับประกาศใช้เดิมไม่เปลี่ยน\n\nยืนยันดำเนินการหรือไม่?`,
      action: async () => {
        setConfirmData(prev=>({...prev,isLoading:true}));
        setIsScheduling(true);
        const termPath=`school_data/${appId}/terms/${termKey}`;
        const lockCollection=`${termPath}/config`;
        const lockRef=doc(db,`${lockCollection}/schedulerLock`);
        const lockOwnerId=doc(collection(db,lockCollection)).id;
        let ownsLock=false;
        try {
          if (getTermKey(schoolSettings) !== termKey) throw new Error('CLEAR_TERM_CHANGED');
          const currentSettings=await getDocFromServer(doc(db,`school_data/${appId}/config/settings`));
          if (currentSettings.exists() && getTermKey(currentSettings.data()) !== termKey) throw new Error('CLEAR_TERM_CHANGED');
          ownsLock = await acquireSchedulerLock(lockRef, lockOwnerId);
          if (!ownsLock) throw new Error('CLEAR_LOCK_BUSY');
          const [ss,fs,gs,ps]=await Promise.all(['schedules','fixedSchedules','activityGroups','plcSessions'].map(name=>getDocsFromServer(collection(db,`${termPath}/${name}`))));
          const rows=snapshot=>snapshot.docs.map(item=>({...item.data(),id:item.id}));
          const liveSchedules=rows(ss), liveFixed=rows(fs), liveGroups=rows(gs), livePLC=rows(ps);
          if (fingerprintScheduleRows(liveSchedules)!==expectedFingerprint) throw new Error('CLEAR_DATA_CHANGED');
          const plan=planUnlockedScheduleClear(liveSchedules, targetRoomId);
          if (plan.deleteCount!==preview.deleteCount || plan.autoCount!==preview.autoCount || plan.manualCount!==preview.manualCount || plan.pinnedCount!==preview.pinnedCount) throw new Error('CLEAR_DATA_CHANGED');
          if (!plan.deleteCount) throw new Error('CLEAR_DATA_CHANGED');
          const operations=plan.selected.map(item=>({type:'delete',ref:doc(db,`${termPath}/schedules`,item.id)}));
          assertAtomicWriteCapacity(operations);
          const historyId=await saveScheduleVersion(targetRoomId ? 'before_clear_room_unlocked' : 'before_clear_all_unlocked',liveSchedules,liveFixed,termKey,liveGroups,null,livePLC);
          if (!historyId) throw new Error('CLEAR_HISTORY_FAILED');
          if (getTermKey(schoolSettings)!==termKey) throw new Error('CLEAR_TERM_CHANGED');
          const latestSettings=await getDocFromServer(doc(db,`school_data/${appId}/config/settings`));
          if (latestSettings.exists() && getTermKey(latestSettings.data()) !== termKey) throw new Error('CLEAR_TERM_CHANGED');
          if (!await renewSchedulerLock(lockRef,lockOwnerId)) throw new Error('CLEAR_LOCK_LOST');
          // A second fresh read detects intervening manual edits while history was saved.
          const verified=await getDocsFromServer(collection(db,`${termPath}/schedules`));
          if (fingerprintScheduleRows(rows(verified))!==expectedFingerprint) throw new Error('CLEAR_DATA_CHANGED');
          await commitAtomicOperations(operations);
          setScheduleResult(null);
          setConfirmData(null);
          showToast(`ล้างคาบที่ไม่ล็อกสำเร็จ ${plan.deleteCount} คาบ (${mode}) · รักษาคาบล็อก ${plan.pinnedCount} คาบ (มีประวัติก่อนลบ)`);
        } catch(error) {
          console.error('Clear unlocked schedule failed',error);
          const message={
            CLEAR_TERM_CHANGED:'ภาคเรียนเปลี่ยนไป ระบบยังไม่ได้ล้างคาบ กรุณาตรวจสอบภาคเรียนใหม่',
            CLEAR_LOCK_BUSY:'มีผู้อื่นกำลังจัดตารางอยู่ ยังไม่ได้ล้างคาบ กรุณารอสักครู่',
            CLEAR_LOCK_LOST:'ล็อกจัดตารางหมดอายุ ระบบยังไม่ได้ล้างคาบ กรุณาลองใหม่',
            CLEAR_DATA_CHANGED:'ตารางมีการแก้ไขจากเครื่องอื่นหรือข้อมูลใหม่เข้ามา ยังไม่ได้ล้างคาบ กรุณาโหลดหน้าจอใหม่',
            CLEAR_HISTORY_FAILED:'บันทึกประวัติก่อนลบไม่สำเร็จ ระบบยังไม่ได้ล้างคาบ',
            INVALID_SCHEDULE_ROWS:'ข้อมูลตารางไม่ถูกต้อง ระบบยังไม่ได้ล้างคาบ',
          }[error.message];
          showToast(message||(['ATOMIC_LIMIT','ATOMIC_SIZE'].includes(error.code)?error.message:'ล้างคาบไม่สำเร็จ กรุณาตรวจสิทธิ์และการเชื่อมต่อ Firebase'), 'error');
          setConfirmData(prev=>({...prev,isLoading:false}));
        } finally {
          if (ownsLock) {
            try { if (!await releaseSchedulerLock(lockRef,lockOwnerId)) console.error('Could not release schedule clearing lock: ownership changed'); }
            catch(error) {console.error('Release clear lock failed',error);}
          }
          setIsScheduling(false);
        }
      },
    });
  };

  // STAGE 25: Confirm bulk fixed-activity deletion, including ONLY child
  // activity groups linked to the selected roots. Re-read all affected records
  // from Firestore before deleting and keep a history snapshot for recovery.
  const bulkDeleteFixedSchedules = (selectedIds) => {
    if (!user || !rulesReady) return showToast('กรุณารอโหลดกิจกรรมและเงื่อนไขให้ครบก่อน', 'error');
    let initial;
    try { initial = planFixedActivityBulkDeletion(selectedIds, fixedSchedules, activityGroups); }
    catch { return showToast('รายการกิจกรรมเปลี่ยนไป กรุณาเลือกใหม่', 'error'); }
    const termKey = getTermKey(schoolSettings);
    const selectedGroupSnapshot = initial.linkedGroups;
    const previews = initial.selected.slice(0, 6).map(item =>
      `• ${getSubject(item.subjectId).name} — ${item.day} ${getPeriodDisplayName(periods,item.periodId)}`).join('\n');
    const noteGroups = selectedGroupSnapshot.length
      ? `\nมีกิจกรรมย่อยรายห้องที่ผูกกับรายการเหล่านี้ ${selectedGroupSnapshot.length} รายการ ซึ่งจะถูกลบไปด้วย (ไม่ลบกิจกรรมย่อยของกิจกรรมอื่น)` : '';
    setConfirmData({
      type: 'delete', title: `ยืนยันลบ ${initial.ids.length} กิจกรรม`, confirmText: `ลบ ${initial.ids.length} รายการ`,
      message: `ต้องการลบกิจกรรมล็อกที่เลือกออกจากตารางฉบับร่างใช่หรือไม่?\n${previews}${initial.ids.length > 6 ? `\n...และอีก ${initial.ids.length-6} รายการ` : ''}${noteGroups}\n\nคาบกิจกรรมจะหายจากตารางครูและนักเรียนในฉบับร่าง ส่วนตารางที่ประกาศใช้แล้วจะยังคงเดิมจนกว่าจะประกาศใหม่\nระบบจะเก็บประวัติก่อนลบเพื่อกู้คืนได้`,
      action: async () => {
        setConfirmData(prev => ({ ...prev, isLoading: true }));
        try {
          if (!user || getTermKey(schoolSettings) !== termKey) throw new Error('FIXED_TERM_CHANGED');
          const termPath = `school_data/${appId}/terms/${termKey}`;
          const [fixedSnap, groupSnap, schedulesSnap, plcSnap] = await Promise.all([
            getDocsFromServer(collection(db, `${termPath}/fixedSchedules`)),
            getDocsFromServer(collection(db, `${termPath}/activityGroups`)),
            getDocsFromServer(collection(db, `${termPath}/schedules`)),
            getDocsFromServer(collection(db, `${termPath}/plcSessions`)),
          ]);
          const liveFixed = fixedSnap.docs.map(d => ({ ...d.data(), id:d.id }));
          const liveGroups = groupSnap.docs.map(d => ({ ...d.data(), id:d.id }));
          const liveSchedules = schedulesSnap.docs.map(d => ({ ...d.data(), id:d.id }));
          const livePLC = plcSnap.docs.map(d => ({ ...d.data(), id:d.id }));
          const livePlan = planFixedActivityBulkDeletion(initial.ids, liveFixed, liveGroups);
          // If any selected root or linked child was edited, added or removed
          // since the confirmation, cancel the entire deletion for safety.
          if (initial.selected.some(original => {
            const current = livePlan.selected.find(item => String(item.id) === String(original.id));
            return !current || !areFirestoreRecordsEqual(original,current);
          }) || selectedGroupSnapshot.length !== livePlan.linkedGroups.length ||
          selectedGroupSnapshot.some(original => {
            const current = livePlan.linkedGroups.find(g => String(g.id)===String(original.id));
            return !current || !areFirestoreRecordsEqual(original,current);
          })) throw new Error('FIXED_DATA_CHANGED');
          const operations = [
            ...livePlan.linkedGroups.map(g => ({ type:'delete', ref:doc(db,`${termPath}/activityGroups`,String(g.id)) })),
            ...livePlan.ids.map(id => ({ type:'delete', ref:doc(db,`${termPath}/fixedSchedules`,id) })),
          ];
          assertAtomicWriteCapacity(operations);
          const historyId = await saveScheduleVersion('before_bulk_fixed_delete',liveSchedules,liveFixed,termKey,liveGroups,null,livePLC);
          if (!historyId) throw new Error('FIXED_HISTORY_FAILED');
          if (getTermKey(schoolSettings) !== termKey) throw new Error('FIXED_TERM_CHANGED');
          await commitAtomicOperations(operations);
          setConfirmData(null);
          showToast(`ลบกิจกรรมล็อก ${livePlan.ids.length} รายการสำเร็จ${livePlan.linkedGroups.length ? ` (กิจกรรมย่อย ${livePlan.linkedGroups.length} รายการ)` : ''}`);
        } catch(error) {
          console.error('Bulk fixed activity deletion failed',error);
          const messages = {
            FIXED_TERM_CHANGED:'ภาคเรียนเปลี่ยนไป ยังไม่ได้ลบกิจกรรม กรุณาตรวจสอบใหม่',
            FIXED_DATA_CHANGED:'ข้อมูลกิจกรรมเปลี่ยนจากตอนเลือก ยังไม่ได้ลบ กรุณาตรวจสอบใหม่',
            FIXED_SELECTION_CHANGED:'มีรายการกิจกรรมถูกเปลี่ยนหรือลบไปแล้ว ยังไม่ได้ลบรายการอื่น',
            FIXED_HISTORY_FAILED:'บันทึกประวัติก่อนลบไม่สำเร็จ จึงยังไม่ได้ลบกิจกรรม',
          };
          const errorMessage = messages[error.message] || (['ATOMIC_LIMIT','ATOMIC_SIZE'].includes(error.code)
            ? error.message : 'ลบกิจกรรมหลายรายการไม่สำเร็จ ยังไม่ได้ลบเพียงบางส่วน');
          showToast(errorMessage, 'error');
          setConfirmData(prev => ({ ...prev, isLoading:false }));
        }
      },
    });
  };

  const handleRequestDelete = (type, item) => {
    if(type==='fixed' && activityGroups.some(g=>g.fixedId===item.id))
      return showToast('ต้องลบกิจกรรมย่อยของทุกห้องที่อ้างถึงกิจกรรมนี้ก่อน','error');
    let errors = [];
    let scheduleDeleteItems = [item];
    let deleteMessage = 'ข้อมูลนี้จะถูกลบถาวร ดำเนินการต่อหรือไม่?';

    if (type === 'teacher') {
      if (teachingLoads.some(l => l.teacherId === item.id)) errors.push('มีภาระสอน');
      if (schedules.some(s => s.teacherId === item.id)) errors.push('มีตารางสอน');
      if (fixedSchedules.some(f => fixedActivityTeacherIds(f).includes(item.id)) || activityGroups.some(g=>fixedActivityTeacherIds(g).includes(item.id))) errors.push('มีกิจกรรมล็อก');
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


  // STAGE 12: Save loads and any newly discovered subjects in one transaction.
  // Read ALL referenced subject IDs from the server before any write. An existing
  // subject with the same code but a different name/term is never overwritten.
  const saveTeachingLoadsWithSubjects = async (items, proposedSubjects = [], subjectNameAssertions = {}) => {
    if (!user || !Array.isArray(items) || !items.length || !Array.isArray(proposedSubjects)) return false;
    try {
      const semester = String(schoolSettings.semester);
      const basePath = `school_data/${appId}`;
      const termPath = `${basePath}/terms/${getTermKey(schoolSettings)}`;
      const proposedById = new Map();
      for (const proposed of proposedSubjects) {
        const id = normalizeExcelText(proposed.id), name = normalizeExcelText(proposed.name);
        const weekly = Number(proposed.periodsPerWeek);
        if (!validSubjectCode(id) || !name || proposedById.has(id) || !Number.isInteger(weekly) || weekly < 1 ||
          ![1,2].includes(Number(proposed.periodBlock)) || !['any','morning','afternoon'].includes(proposed.preferredTime)) throw new Error('ข้อมูลรายวิชาใหม่ไม่ถูกต้อง');
        proposedById.set(id, { ...proposed, id, name, term: semester, periodsPerWeek: weekly });
      }
      const cleanItems = items.map(item => {
        const id = String(item.id || '').trim();
        const teacherId = String(item.teacherId || '').trim();
        const subjectId = String(item.subjectId || '').trim();
        const classroomId = String(item.classroomId || '').trim();
        const periods = Number(item.periods);
        if (!id || id.includes('/') || !teacherId || !validSubjectCode(subjectId) || !classroomId || !Number.isInteger(periods) || periods < 1) throw new Error('ข้อมูลภาระสอนไม่ถูกต้อง');
        return { id, teacherId, subjectId, classroomId, periods,
          expectedExisting: item.expectedExisting === true, expectedPreviousTriple: item.expectedPreviousTriple || '',
          expectedPreviousPeriods: item.expectedPreviousPeriods }; 
      });
      const uniqueLoadIds = new Set(cleanItems.map(item => item.id));
      const uniqueTriples = new Set(cleanItems.map(getTeachingLoadKey));
      if (uniqueLoadIds.size !== cleanItems.length || uniqueTriples.size !== cleanItems.length) throw new Error('พบภาระสอนซ้ำในข้อมูลที่นำเข้า');
      // The Firebase WEB SDK currently caps runTransaction at 500 writes.
      // Preserve the existing bulk-import path for large files with NO new subjects.
      // Mixed files above this threshold must be split to keep each subject/load pair atomic.
      if (proposedById.size + cleanItems.length > 450) {
        if (proposedById.size === 0) {
          // Existing large-import behavior, but validate the live subject catalog first.
          const serverSubjects = await getDocsFromServer(collection(db, `${basePath}/subjects`));
          const available = new Map(serverSubjects.docs.map(snap => [snap.id, snap.data()]));
          for (const item of cleanItems) {
            const live = available.get(item.subjectId);
            if (!live || !isSubjectActiveInTerm(live, semester) ||
                (subjectNameAssertions[item.subjectId] && normalizeExcelText(live.name) !== normalizeExcelText(subjectNameAssertions[item.subjectId])))
              throw new Error(`รายวิชา ${item.subjectId} มีข้อมูลไม่ตรงกับฐานข้อมูล กรุณาโหลดข้อมูลใหม่`);
          }
          return await bulkUpsertTeachingLoads(cleanItems);
        }
        throw new Error('มีรายวิชาใหม่และรายการรวมเกิน 450 รายการ กรุณาแบ่งไฟล์ Excel เป็นชุดย่อย');
      }
      const subjectIds = [...new Set(cleanItems.map(item => item.subjectId))];
      for (const id of proposedById.keys()) if (!subjectIds.includes(id)) throw new Error(`รายวิชาใหม่ ${id} ไม่มีภาระสอนรองรับ`);
      const subjectRefs = subjectIds.map(id => doc(db, `${basePath}/subjects`, id));
      const loadRefs = cleanItems.map(item => doc(db, `${termPath}/teachingLoads`, item.id));
      const currentLoadSnapshot = await getDocsFromServer(collection(db, `${termPath}/teachingLoads`));
      const importedIds = new Set(cleanItems.map(i => i.id));
      const existingOtherTriples = new Set(currentLoadSnapshot.docs
        .filter(snap => !importedIds.has(snap.id)).map(snap => getTeachingLoadKey(snap.data())));
      if (cleanItems.some(i => existingOtherTriples.has(getTeachingLoadKey(i)))) throw new Error('พบภาระสอนซ้ำกับรายการอื่นในระบบ กรุณาโหลดข้อมูลใหม่');
      await runTransaction(db, async tx => {
        // All reads precede writes, including current load versions.
        const subjectSnaps = await Promise.all(subjectRefs.map(ref => tx.get(ref)));
        const loadSnaps = await Promise.all(loadRefs.map(ref => tx.get(ref)));
        for (let idx = 0; idx < subjectIds.length; idx++) {
          const id = subjectIds[idx];
          const live = subjectSnaps[idx].exists() ? subjectSnaps[idx].data() : null;
          const proposed = proposedById.get(id);
          if (!live && !proposed) throw new Error(`ไม่พบรายวิชา ${id} ในฐานข้อมูล กรุณาเพิ่มวิชาหรือโหลดข้อมูลใหม่`);
          if (live && proposed) throw new Error(`รหัสวิชา ${id} มีอยู่แล้วในฐานข้อมูล กรุณาตรวจสอบชื่อวิชาและลองใหม่`);
          if (live && !isSubjectActiveInTerm(live, semester)) throw new Error(`รายวิชา ${id} ไม่เปิดใช้ภาคเรียน ${semester}`);
          if (live && subjectNameAssertions[id] && normalizeExcelText(live.name) !== normalizeExcelText(subjectNameAssertions[id]))
            throw new Error(`ชื่อวิชา ${id} เปลี่ยนไปจากตอนตรวจสอบ กรุณาโหลดข้อมูลใหม่`);
          if (live && !validSubjectCode(id)) throw new Error(`รหัสวิชา ${id} ไม่ถูกต้อง`);
        }
        for (let i = 0; i < cleanItems.length; i++) {
          const live = loadSnaps[i].exists() ? loadSnaps[i].data() : null;
          if (cleanItems[i].expectedExisting) {
            if (!live || getTeachingLoadKey(live) !== cleanItems[i].expectedPreviousTriple ||
                Number(live.periods) !== Number(cleanItems[i].expectedPreviousPeriods))
              throw new Error(`รายการภาระสอน ${cleanItems[i].id} เปลี่ยนไปแล้ว กรุณาโหลดข้อมูลใหม่`);
          } else if (live) {
            throw new Error(`มีรหัสรายการภาระสอน ${cleanItems[i].id} อยู่แล้ว กรุณาลองอีกครั้ง`);
          }
        }
        subjectIds.forEach((id, i) => {
          const proposed = proposedById.get(id);
          if (proposed) tx.set(subjectRefs[i], proposed);
        });
        cleanItems.forEach((item,i) => tx.set(loadRefs[i], { id:item.id, teacherId:item.teacherId, subjectId:item.subjectId, classroomId:item.classroomId, periods:item.periods }));
      });
      return true;
    } catch (error) {
      console.error('Save loads and new subjects failed', error);
      showToast(`บันทึกไม่สำเร็จ: ${error?.message || 'กรุณาลองอีกครั้ง'} (ไม่มีการบันทึกเป็นบางส่วน)`, 'error');
      return false;
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

  // STAGE 17.1: a published timetable freezes the *display and validation* references.
  // Use small separate documents; never place a potentially large school-wide map
  // in the manifest or a single Firestore document.
  const PUBLISHED_REF_CHARS = 30000;
  const snapshotTextHash = value => {
    let h=2166136261;
    for(let i=0;i<value.length;i++) h=Math.imul(h ^ value.charCodeAt(i),16777619)>>>0;
    return `${value.length}:${h.toString(16)}`;
  };
  const readPublishedReferences = async (versionPath,manifest) => {
    if(!manifest.referencePartCount) return null; // legacy publications are still readable
    if(!Number.isInteger(manifest.referencePartCount)||manifest.referencePartCount<1||manifest.referencePartCount>1000)
      throw new Error('จำนวนส่วนข้อมูลอ้างอิงฉบับประกาศไม่ถูกต้อง');
    const snapshot=await getDocsFromServer(collection(db,`${versionPath}/references`));
    if(snapshot.docs.length!==manifest.referencePartCount)throw new Error('ข้อมูลชื่อครู/วิชาของฉบับประกาศไม่ครบ');
    const parts=new Map(snapshot.docs.map(d=>[d.id,d.data()]));
    let serialized='';
    for(let i=0;i<manifest.referencePartCount;i++){
      const part=parts.get(`ref_${String(i).padStart(5,'0')}`);
      if(typeof part?.text!=='string')throw new Error('ข้อมูลอ้างอิงฉบับประกาศเสียหาย');
      serialized+=part.text;
    }
    if(snapshotTextHash(serialized)!==manifest.referenceFingerprint)throw new Error('ข้อมูลอ้างอิงฉบับประกาศไม่ตรงฉบับที่บันทึก');
    const data=JSON.parse(serialized);
    const required=['teachers','subjects','classrooms','periods','teachingLoads','unavailabilities','classroomUnavailabilities','scheduleRules'];
    if(required.some(name=>!Array.isArray(data[name]))||!data.schoolSettings)throw new Error('ข้อมูลอ้างอิงฉบับประกาศไม่ครบ');
    // Old Stage 20 published copies predate PLC: keep them readable as PLC-free copies.
    if(!Array.isArray(data.plcSessions))data.plcSessions=[];
    return data;
  };

  // STAGE 5B: history manifests point to bounded-size documents. The manifest is
  // committed LAST, so incomplete chunk uploads are not advertised as restorable.
  const readScheduleVersion = async (termKey, versionId) => {
    const versionPath = `school_data/${appId}/terms/${termKey}/scheduleVersions/${versionId}`;
    const header = await getDocFromServer(doc(db, versionPath));
    if (!header.exists()) throw new Error('ไม่พบประวัติตารางนี้');
    const data = header.data();
    if (Array.isArray(data.schedules) && Array.isArray(data.fixedSchedules)) return {
      ...data, id:versionId, schedules:data.schedules.map(cleanScheduleSnapshotItem),
      fixedSchedules:data.fixedSchedules.map(cleanScheduleSnapshotItem),
      activityGroups:Array.isArray(data.activityGroups) ? data.activityGroups.map(cleanScheduleSnapshotItem) : null,
      plcSessions:Array.isArray(data.plcSessions) ? data.plcSessions.map(cleanScheduleSnapshotItem) : null,
      includesActivityGroups:Array.isArray(data.activityGroups),
      includesPLCSessions:Array.isArray(data.plcSessions),
      referenceData:await readPublishedReferences(versionPath,data),
    }; // preserve legacy format; do not erase live groups if absent
    if (![2,3,4].includes(data.formatVersion) || data.state !== 'ready' || !Number.isInteger(data.chunkCount) ||
      data.chunkCount < 0 || data.chunkCount > 1000) throw new Error('รูปแบบประวัติไม่ถูกต้องหรือยังบันทึกไม่สำเร็จ');
    const chunkSnap = await getDocsFromServer(collection(db, `${versionPath}/chunks`));
    if (chunkSnap.docs.length !== data.chunkCount) throw new Error('ประวัติตารางมีชิ้นส่วนไม่ครบ');
    const byId = new Map(chunkSnap.docs.map(d => [d.id,d.data()]));
    const result={ schedules:[], fixedSchedules:[], activityGroups:[], plcSessions:[] };
    for(let i=0;i<data.chunkCount;i++) {
      const part=byId.get(`part_${String(i).padStart(5,'0')}`);
      if(!part || !Array.isArray(part.schedules) || !Array.isArray(part.fixedSchedules)) throw new Error('ประวัติตารางมีชิ้นส่วนเสียหาย');
      result.schedules.push(...part.schedules);
      result.fixedSchedules.push(...part.fixedSchedules);
      if (data.formatVersion >= 3) {
        if (!Array.isArray(part.activityGroups)) throw new Error('ประวัติกิจกรรมกลุ่มมีชิ้นส่วนเสียหาย');
        result.activityGroups.push(...part.activityGroups);
      }
      if(data.formatVersion>=4){
        if(!Array.isArray(part.plcSessions))throw new Error('ประวัติ PLC มีชิ้นส่วนเสียหาย');
        result.plcSessions.push(...part.plcSessions);
      }
    }
    if(result.schedules.length!==data.scheduleCount || result.fixedSchedules.length!==data.fixedScheduleCount ||
      fingerprintScheduleRows(result.schedules)!==data.scheduleFingerprint ||
      fingerprintScheduleRows(result.fixedSchedules)!==data.fixedFingerprint ||
      (data.formatVersion>=3 && (result.activityGroups.length!==data.activityGroupCount ||
      fingerprintScheduleRows(result.activityGroups)!==data.activityGroupFingerprint)) ||
      (data.formatVersion>=4 && (result.plcSessions.length!==data.plcCount ||
      fingerprintScheduleRows(result.plcSessions)!==data.plcFingerprint)))
      throw new Error('ข้อมูลประวัติไม่ตรงกับรายการที่บันทึกไว้');
    return { ...data, id:versionId, ...result,
      includesActivityGroups:data.formatVersion>=3,
      includesPLCSessions:data.formatVersion>=4,
      referenceData:await readPublishedReferences(versionPath,data),
      activityGroups:data.formatVersion>=3 ? result.activityGroups : null,
      plcSessions:data.formatVersion>=4 ? result.plcSessions : null };
  };
  // STAGE 15: published data is an IMMUTABLE history snapshot, never the live
  // schedules collection. Editors may continue improving the draft afterwards.
  useEffect(()=>{
    if(!user) {setPublishedSnapshot(null);return;}
    const termKey=getTermKey(schoolSettings);
    let alive=true;
    let requestSerial=0;
    const ref=doc(db,`school_data/${appId}/terms/${termKey}/config/publishedTimetable`);
    setPublishedSnapshot(null);setPublishedError('');setPublishedLoading(true);
    const unsubscribe=onSnapshot(ref,async snap=>{
      const requestId=++requestSerial;
      if(!alive)return;
      if(!snap.exists()||!snap.data().versionId){setPublishedSnapshot(null);setPublishedLoading(false);return;}
      const meta=snap.data();
      try{
        const version=await readScheduleVersion(termKey,meta.versionId);
        if(!version.includesActivityGroups) throw new Error('ฉบับประกาศไม่มีข้อมูลกิจกรรมกลุ่มครบ');
        if(alive&&requestId===requestSerial)setPublishedSnapshot({termKey,versionId:meta.versionId,publishedAt:meta.publishedAt,
          schedules:version.schedules,fixedSchedules:version.fixedSchedules,activityGroups:version.activityGroups,
          referenceData:version.referenceData});
        if(alive&&requestId===requestSerial)setPublishedError('');
      }catch(error){console.error('Cannot load published timetable',error);
        if(alive&&requestId===requestSerial){setPublishedSnapshot(null);setPublishedError('โหลดตารางฉบับประกาศไม่ได้ กรุณาตรวจประวัติและการเชื่อมต่อ');}}
      finally{if(alive&&requestId===requestSerial)setPublishedLoading(false);}
    },error=>{console.error(error);if(alive){setPublishedError('อ่านสถานะประกาศใช้ไม่ได้');setPublishedLoading(false);}});
    return()=>{alive=false;unsubscribe();};
  },[user,schoolSettings.academicYear,schoolSettings.semester]);

  const pruneScheduleVersions = async termPath => {
    try {
      const snapshot = await getDocsFromServer(collection(db, `${termPath}/scheduleVersions`));
      const versions=snapshot.docs.map(d=>({ref:d.ref, data:d.data()}))
        .sort((a,b)=>(Number(b.data.createdAt)||0)-(Number(a.data.createdAt)||0));
      // Do not delete a version referenced by the published timetable pointer.
      const publishedRef=await getDocFromServer(doc(db,`${termPath}/config/publishedTimetable`));
      const pinned=publishedRef.exists()?publishedRef.data().versionId:null;
      for(const item of versions.filter(x=>x.ref.id!==pinned).slice(SCHEDULE_VERSION_LIMIT)) {
        // Delete child documents first; deleting a Firestore parent does not delete children.
        const children=await getDocsFromServer(collection(db, `${item.ref.path}/chunks`));
        for(const part of children.docs) await deleteDoc(part.ref);
        const references=await getDocsFromServer(collection(db,`${item.ref.path}/references`));
        for(const part of references.docs) await deleteDoc(part.ref);
        await deleteDoc(item.ref);
      }
    } catch(error) { console.error('Prune schedule versions failed; old history retained',error); }
  };
  const saveScheduleVersion = async (reason = 'manual_snapshot', scheduleItems = schedules, fixedItems = fixedSchedules, explicitTermKey = null, groupItems = activityGroups, referenceData = null, plcItems = plcSessions) => {
    if (!user) return null;
    const termKey=explicitTermKey || getTermKey(schoolSettings);
    if (!/^\d{4}-[12]$/.test(termKey)) return null;
    const termPath=`school_data/${appId}/terms/${termKey}`;
    const versionRef=doc(collection(db, `${termPath}/scheduleVersions`));
    const regular=(Array.isArray(scheduleItems)?scheduleItems:[]).map(cleanScheduleSnapshotItem);
    const fixed=(Array.isArray(fixedItems)?fixedItems:[]).map(cleanScheduleSnapshotItem);
    const groups=(Array.isArray(groupItems)?groupItems:[]).map(cleanScheduleSnapshotItem);
    const plcs=(Array.isArray(plcItems)?plcItems:[]).map(cleanScheduleSnapshotItem);
    const chunks=makeScheduleSnapshotChunks(regular,fixed,groups,plcs);
    const referenceText=referenceData ? JSON.stringify(referenceData) : '';
    const referenceParts=referenceData ? (referenceText.match(/[\s\S]{1,30000}/g)||[]) : [];
    const saved=[];
    try {
      for(let i=0;i<chunks.length;i++) {
        const partRef=doc(db, `${versionRef.path}/chunks/part_${String(i).padStart(5,'0')}`);
        await setDoc(partRef,chunks[i]);saved.push(partRef);
      }
      for(let i=0;i<referenceParts.length;i++) {
        const partRef=doc(db,`${versionRef.path}/references/ref_${String(i).padStart(5,'0')}`);
        await setDoc(partRef,{text:referenceParts[i]});saved.push(partRef);
      }
      await setDoc(versionRef, { id:versionRef.id, formatVersion:4,state:'ready',createdAt:Date.now(),
        reason:String(reason || 'manual_snapshot'),scheduleCount:regular.length,fixedScheduleCount:fixed.length,
        activityGroupCount:groups.length,plcCount:plcs.length,chunkCount:chunks.length,scheduleFingerprint:fingerprintScheduleRows(regular),
        fixedFingerprint:fingerprintScheduleRows(fixed),activityGroupFingerprint:fingerprintScheduleRows(groups),plcFingerprint:fingerprintScheduleRows(plcs),
        referencePartCount:referenceParts.length,referenceFingerprint:referenceData?snapshotTextHash(referenceText):null });
    } catch(error) {
      console.error('Save schedule version failed; no ready manifest created',error);
      for(const part of saved) try { await deleteDoc(part); } catch(cleanupError) { console.error('Orphan history chunk',cleanupError); }
      return null;
    }
    await pruneScheduleVersions(termPath);
    return versionRef.id;
  };

  // STAGE 5C: restores both the regular and fixed collection as ONE atomic write.
  // Reject obsolete previews and invalid historical references before any write.
  const restoreScheduleVersion = (versionId, expectedFingerprints) => {
    const termKey=getTermKey(schoolSettings);
    setConfirmData({ type:'restore',title:'ยืนยันกู้คืนตารางสอน',confirmText:'กู้คืนเวอร์ชันนี้',
      message:`ระบบจะกู้คืนตารางสอน กิจกรรมล็อก กิจกรรมกลุ่ม และ PLC (หากเวอร์ชันนั้นบันทึกไว้) ของภาคเรียน ${termKey} กลับไปตามประวัติที่เลือก\nจะบันทึกจุดย้อนกลับของตารางปัจจุบันก่อน และไม่เปลี่ยนข้อมูลครู/ภาระสอน ยืนยันหรือไม่?`,
      action:async()=>{
        setConfirmData(prev=>({...prev,isLoading:true}));
        try {
          if (getTermKey(schoolSettings)!==termKey) throw new Error('ภาคเรียนเปลี่ยนไป');
          const restored=await readScheduleVersion(termKey,versionId);
          const path=`school_data/${appId}/terms/${termKey}`;
          const [liveSch,liveFix,liveGroups,livePLC] = await Promise.all([
            getDocsFromServer(collection(db,`${path}/schedules`)),
            getDocsFromServer(collection(db,`${path}/fixedSchedules`)),
            getDocsFromServer(collection(db,`${path}/activityGroups`)),
            getDocsFromServer(collection(db,`${path}/plcSessions`)),
          ]);
          const previousSchedules=liveSch.docs.map(x=>({...x.data(),id:x.id}));
          const previousFixed=liveFix.docs.map(x=>({...x.data(),id:x.id}));
          const previousGroups=liveGroups.docs.map(x=>({...x.data(),id:x.id}));
          const previousPLC=livePLC.docs.map(x=>({...x.data(),id:x.id}));
          const targetGroups=restored.includesActivityGroups ? restored.activityGroups : previousGroups;
          const targetPLC=restored.includesPLCSessions ? restored.plcSessions : previousPLC;
          if (fingerprintScheduleRows(previousSchedules)!==expectedFingerprints.schedules ||
              fingerprintScheduleRows(previousFixed)!==expectedFingerprints.fixedSchedules ||
              fingerprintScheduleRows(previousGroups)!==expectedFingerprints.activityGroups ||
              fingerprintScheduleRows(previousPLC)!==expectedFingerprints.plcSessions)
            throw new Error('ตารางเปลี่ยนหลังเปิดหน้าประวัติ กรุณาโหลดและเปรียบเทียบใหม่');
          if (targetGroups.some(g => !restored.fixedSchedules.some(f => f.id===g.fixedId &&
              f.classroomId===ALL_CLASSROOMS_ID && fixedAppliesToClassroom(f,g.classroomId))))
            throw new Error('กิจกรรมกลุ่มไม่ตรงกับกิจกรรมหลักของประวัติที่เลือก');
          const healthCheck=calculateScheduleHealth(restored.schedules,teachingLoads,unavailabilities,
            classroomUnavailabilities,restored.fixedSchedules,teachers,subjects,classrooms,periods,customRules,targetGroups,targetPLC);
          if (healthCheck.status==='error') throw new Error('ประวัตินี้ขัดกับข้อมูลหรือข้อบังคับปัจจุบัน กรุณาตรวจภาระสอนและเงื่อนไขก่อน');
          const [regularPlan,fixedPlan,groupsPlan,plcPlan] = await Promise.all([
            prepareCollectionReplacement(`${path}/schedules`,restored.schedules),
            prepareCollectionReplacement(`${path}/fixedSchedules`,restored.fixedSchedules),
            ...(restored.includesActivityGroups ? [prepareCollectionReplacement(`${path}/activityGroups`,targetGroups)] : [Promise.resolve({operations:[]})]),
            ...(restored.includesPLCSessions ? [prepareCollectionReplacement(`${path}/plcSessions`,targetPLC)] : [Promise.resolve({operations:[]})]),
          ]);
          const operations=[...regularPlan.operations,...fixedPlan.operations,...(groupsPlan?.operations || []),...(plcPlan?.operations||[])];
          assertAtomicWriteCapacity(operations);
          const rollbackId=await saveScheduleVersion('before_history_restore',previousSchedules,previousFixed,termKey,previousGroups,null,previousPLC);
          if (!rollbackId) throw new Error('สร้างจุดย้อนกลับไม่สำเร็จ ระบบจึงไม่เปลี่ยนตาราง');
          // Re-check immediately before commit. Multi-client edits racing the commit
          // still require a shared version fence enforced by Security Rules/backend.
          const [checkSch,checkFixed,checkGroups,checkPLC]=await Promise.all([
            getDocsFromServer(collection(db,`${path}/schedules`)),
            getDocsFromServer(collection(db,`${path}/fixedSchedules`)),
            getDocsFromServer(collection(db,`${path}/activityGroups`)),
            getDocsFromServer(collection(db,`${path}/plcSessions`)),
          ]);
          if(fingerprintScheduleRows(checkSch.docs.map(d=>({...d.data(),id:d.id})))!==expectedFingerprints.schedules ||
             fingerprintScheduleRows(checkFixed.docs.map(d=>({...d.data(),id:d.id})))!==expectedFingerprints.fixedSchedules ||
             fingerprintScheduleRows(checkGroups.docs.map(d=>({...d.data(),id:d.id})))!==expectedFingerprints.activityGroups ||
             fingerprintScheduleRows(checkPLC.docs.map(d=>({...d.data(),id:d.id})))!==expectedFingerprints.plcSessions)
            throw new Error('ข้อมูลตารางเปลี่ยนขณะกู้คืน จึงยกเลิกโดยไม่เขียนทับ');
          await commitAtomicOperations(operations);
          setConfirmData(null);showToast('กู้คืนตารางสำเร็จ มีจุดย้อนกลับก่อนกู้คืนแล้ว');
        } catch(error) { console.error('Restore history failed',error);setConfirmData(null);
          showToast(`กู้คืนไม่สำเร็จ: ${error.message}`,'error'); }
      }
    });
  };

  const handleRestoreBackup = (jsonData) => {
    const backupTermKey = jsonData?.termKey || (jsonData?.schoolSettings ? getTermKey(jsonData.schoolSettings) : '');
    const backupCollections = ['teachers', 'subjects', 'classrooms', 'periods', 'teachingLoads', 'unavailabilities', 'classroomUnavailabilities', 'fixedSchedules', 'schedules'];
    // Older v2 backups do not contain scheduleRules. Preserve live rules for those backups.
    const restoreCollections = [...backupCollections, ...(Array.isArray(jsonData?.scheduleRules)?['scheduleRules']:[]), ...(Array.isArray(jsonData?.activityGroups)?['activityGroups']:[]), ...(Array.isArray(jsonData?.plcSessions)?['plcSessions']:[])];
    if (!jsonData || jsonData.backupVersion !== 2 || jsonData.appId !== appId || !/^\d{4}-[12]$/.test(backupTermKey) || !backupCollections.every(key => Array.isArray(jsonData[key]))) {
      return showToast("ไฟล์สำรองข้อมูลไม่ครบ หรือไม่ตรงระบบและภาคเรียนที่รองรับ", "error");
    }

    setConfirmData({
      type: 'restore', title: 'นำเข้าข้อมูล', message: `เขียนทับตารางเดิมสำหรับเทอม ${backupTermKey} ใช่หรือไม่?${!Array.isArray(jsonData.scheduleRules) ? '\nไฟล์นี้ไม่มีเงื่อนไขเพิ่มเติม ระบบจะเก็บเงื่อนไขปัจจุบันไว้' : ''}`, confirmText: 'นำเข้า',
      action: async () => {
        setConfirmData(prev => ({ ...prev, isLoading: true }));
        try {
          const basePath = `school_data/${appId}`;
          const targetTermPath = `${basePath}/terms/${backupTermKey}`;
          const masterNames = ['teachers', 'subjects', 'classrooms', 'periods'];
          // STEP 03: Plan ALL nine collections first; one batch commits everything or nothing.
          const plans = {};
          const allOperations = [];
          for (const colName of restoreCollections) {
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
          const restoreVersionId = await saveScheduleVersion('before_backup_restore', plans.schedules.previousItems,
            plans.fixedSchedules.previousItems, backupTermKey,
            plans.activityGroups?.previousItems || activityGroups,null,plans.plcSessions?.previousItems || plcSessions);
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
    'classroomUnavailabilities', 'teachers', 'subjects', 'classrooms', 'periods', 'scheduleRules', 'activityGroups', 'plcSessions',
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
    classroomUnavailabilities, teachers, subjects, classrooms, periods, scheduleRules: customRules, activityGroups, plcSessions,
  }, getTermKey(schoolSettings));

  const runSmartSchedule = async () => {
    if (isScheduling) return;
    if (!rulesReady) return showToast('ยังโหลดเงื่อนไขเพิ่มเติมไม่สำเร็จ จึงไม่เริ่มจัดตาราง', 'error');
    // Recommended workflow: reserve PLC before teaching. Still allow the editor to
    // proceed deliberately when some groups do not need PLC in this term.
    const activeAreas=[...new Set(teachers.map(t=>t.subjectArea).filter(a=>PLC_AREAS.includes(a)))];
    const missingPLC=activeAreas.filter(a=>!plcSessions.some(p=>p.subjectArea===a));
    if(missingPLC.length && !window.confirm(`ยังไม่ได้กำหนด PLC ให้กลุ่มสาระ: ${missingPLC.join(', ')}\n\nแนะนำให้เข้า “ข้อจำกัดและเงื่อนไข > PLC กลุ่มสาระ” แล้วจองคาบก่อนจัดวิชา เพื่อให้ Auto ไม่จัดวิชาทับ\n\nต้องการจัด Auto ต่อโดยยังไม่กำหนด PLC ให้กลุ่มเหล่านี้หรือไม่?`)) return;
    // STAGE 17.4: never assume an imported subject is single period without review.
    const pendingSubjects=subjects.filter(s=>s.needsScheduleReview &&
      teachingLoads.some(load=>load.subjectId===s.id));
    if(pendingSubjects.length)return showToast(`มีรายวิชาที่ Excel สร้างไว้รอตรวจ Block/ช่วงเวลา ${pendingSubjects.length} วิชา: ${pendingSubjects.slice(0,3).map(s=>s.id).join(', ')} กรุณาแก้ในข้อมูลพื้นฐาน > รายวิชา ก่อนจัด Auto`, 'error');
    // STEP 04: แจ้งรายการที่ต้องแก้ก่อนขอล็อกหรือเริ่มสร้างตาราง
    const issues = validateAutoScheduleInputs({
      teachingLoads, teachers, subjects, classrooms, periods, schedules, fixedSchedules,
      unavailabilities, classroomUnavailabilities, semester: schoolSettings.semester, customRules, activityGroups, plcSessions,
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
        const baseMaps = buildConstraintMaps(preservedSchedules, unavailabilities, classroomUnavailabilities, fixedSchedules, periods, customRules, activityGroups, classrooms, plcSessions, teachers);
        baseMaps.enforceAutoSubjectDailyCap = true;
        baseMaps.autoSubjectBlockSizes = new Map(subjects.map(s=>[String(s.id), Number(s.periodBlock)===2 ? 2 : 1]));
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
              const valid = canPlaceSchedule(load.teacherId, load.classroomId, day, period.id, maps, false, load.subjectId);
              if (valid.ok) options.push([{ day, periodId: period.id }]);
            });
          });
          return options;
        };

        const enumerateDoubleOptions = (load, maps) => {
          const options = [];
          DAYS.forEach(day => {
            periodContext.doublePairs.forEach(([p1, p2]) => {
              const valid1 = canPlaceSchedule(load.teacherId, load.classroomId, day, p1.id, maps, false, load.subjectId);
              const valid2 = canPlaceSchedule(load.teacherId, load.classroomId, day, p2.id, maps, false, load.subjectId);
              if (valid1.ok && valid2.ok) {
                const pair = [{ day, periodId: p1.id }, { day, periodId: p2.id }];
                if (canPlaceTeacherEntries(load.teacherId, pair, maps).ok && canPlaceSubjectDailyEntries(load, pair, maps).ok) options.push(pair);
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

        // จัดรอบแรกด้วยเพดาน 4 คาบเสมอ; ถ้าครบแล้วจะไม่เปิดโหมด 5 คาบ
        // หากยังขาด จึงลองรอบสำรอง 5 คาบ โดยให้คะแนนรูปแบบ 5 คาบต่ำมาก
        for (const dayLimit of [SCHEDULER_CONFIG.maxTeacherPeriodsPerDay, SCHEDULER_CONFIG.absoluteMaxTeacherPeriodsPerDay]) {
        if (dayLimit > SCHEDULER_CONFIG.maxTeacherPeriodsPerDay && candidatePool.some(x => x.missingCount === 0)) break;
        // Stage 17.5: search the four-period space more thoroughly before trying fifth periods.
        // This is heuristic optimization, not a mathematical proof that four periods is impossible.
        const attempts = dayLimit === SCHEDULER_CONFIG.maxTeacherPeriodsPerDay
          ? SCHEDULER_CONFIG.candidateCount * 4 : SCHEDULER_CONFIG.candidateCount;
        for (let c = 0; c < attempts; c++) {
          const random = createSeededRandom((runSeed + Math.imul(c + 1, 2246822519) + Math.imul(dayLimit, 3266489917)) >>> 0);
          const generatedSchedule = [];
          const currentUnassigned = [];
          const currentMaps = cloneConstraintMaps(baseMaps);
          currentMaps.maxTeacherPeriodsPerDay = dayLimit;
          const placementState = buildPlacementState([
            ...markFixedTeacherActivities(fixedSchedules), ...activityTeacherEntries(activityGroups,fixedSchedules), ...preservedSchedules]);
          let idSequence = 0;

          const addEntries = (load, entries) => {
            entries.forEach(entry => {
              const id = `sch_auto_${runSeed}_${dayLimit}_${c}_${idSequence++}`;
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
              registerSubjectDayPeriod(currentMaps, item);
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
                const chosen = choosePlacementOption(options, load, placementState, periodContext, getSubject, random, customRules);
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
              const chosen = choosePlacementOption(options, load, placementState, periodContext, getSubject, random, customRules);
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
          const score = calculateScheduleScore(fullSchedule, teachers, subjects, periodContext,
            getSubject, customRules, [...markFixedTeacherActivities(fixedSchedules),...activityTeacherEntries(activityGroups,fixedSchedules)]);
          const missingCount = currentUnassigned.reduce((acc, curr) => acc + (Number(curr.missing) || 0), 0);
          const fivePeriodCount = findTeacherFifthPeriodDays([
            ...fixedSchedules, ...activityTeacherEntries(activityGroups, fixedSchedules), ...fullSchedule], periods).length;
          const weeklyBalance=evaluateWeeklyTeachingBalance(fullSchedule,fixedSchedules,activityGroups,periods,teachers);
          const finalScore = score - (missingCount * 1000) - weeklyBalance.penalty;
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
              fivePeriodCount,
              weeklyBalance,
              generationId: String(runSeed),
              maxDailyLimitUsed: dayLimit,
            });
          }

          // Stop early only after gathering three different complete 4-period choices;
          // otherwise search the extended 4-period space before allowing five periods.
          if(dayLimit===SCHEDULER_CONFIG.maxTeacherPeriodsPerDay && c>=69 &&
            candidatePool.filter(x=>x.maxDailyLimitUsed===dayLimit && x.missingCount===0).length>=3) break;
          // คืนเวลาให้ browser เป็นช่วง ๆ เพื่อให้ spinner/การโต้ตอบไม่ค้างระหว่างคำนวณหลาย candidate
          if ((c + 1) % SCHEDULER_CONFIG.yieldEveryCandidates === 0 && c + 1 < attempts) {
            // Renew only if needed, to avoid adding a network round trip on every candidate.
            if (Date.now() - lastLockRenewal >= SCHEDULER_LOCK_RENEW_MS) {
              if (!(await renewSchedulerLock(lockRef, lockOwnerId))) throw new Error('SCHEDULER_LOCK_LOST');
              lastLockRenewal = Date.now();
            }
            await new Promise(resolve => setTimeout(resolve, 0));
          }
        }
        } // สิ้นสุดรอบจัด 4 คาบและรอบสำรอง 5 คาบ

        // Do not publish a result if this run has lost ownership of the lock.
        if (!(await renewSchedulerLock(lockRef, lockOwnerId))) throw new Error('SCHEDULER_LOCK_LOST');
        const fourPeriodAttempts = candidatePool.filter(x=>x.maxDailyLimitUsed===SCHEDULER_CONFIG.maxTeacherPeriodsPerDay);
        const bestFourPeriodMissing = fourPeriodAttempts.length ? Math.min(...fourPeriodAttempts.map(x=>x.missingCount)) : null;
        const fourPeriodCandidatesTried=fourPeriodAttempts.length;
        candidatePool.sort((a, b) => {
          if (a.missingCount !== b.missingCount) return a.missingCount - b.missingCount;
          if (a.fivePeriodCount !== b.fivePeriodCount) return a.fivePeriodCount - b.fivePeriodCount;
          if (a.score !== b.score) return b.score - a.score;
          return a.generatedSchedule.length - b.generatedSchedule.length;
        });

        const candidates = candidatePool.slice(0, SCHEDULER_CANDIDATE_KEEP);
        const bestCandidate = candidates[0] || { schedule: preservedSchedules, generatedSchedule: [], preservedCount: preservedSchedules.length, unassigned: [], score: -Infinity, missingCount: 0, generationId: String(runSeed) };
        if (findSubjectDailyLimitConflicts([...fixedSchedules, ...bestCandidate.schedule], classrooms, customRules).length ||
            findAutoSubjectDistributionConflicts(
              [...fixedSchedules.filter(f => !isDevelopmentActivity(f)), ...bestCandidate.schedule],
              subjects, classrooms,
              [...fixedSchedules.filter(f => !isDevelopmentActivity(f)), ...preservedSchedules]
            ).length)
          throw new Error('SUBJECT_DAILY_LIMIT_FAILED');
        if (findTeacherTeachingLimitConflicts(buildTeacherOccupancy(bestCandidate.schedule,fixedSchedules,activityGroups), periods, teachers).length) {
          throw new Error('TEACHER_HARD_LIMIT_FAILED');
        }
        const totalReq = teachingLoads.reduce((a, curr) => a + (Number(curr.periods) || 0), 0);
        const missing = bestCandidate.unassigned.reduce((sum, item) => sum + (Number(item.missing) || 0), 0);
        const fivePeriodDays = findTeacherFifthPeriodDays(
          [...fixedSchedules, ...activityTeacherEntries(activityGroups, fixedSchedules), ...bestCandidate.schedule],
          periods, teachers);
        setScheduleResult({
          ...bestCandidate,
          fivePeriodDays,
          bestFourPeriodMissing,
          fourPeriodCandidatesTried,
          weeklyLoadSummary:[...summarizeTeacherWeek(buildTeacherOccupancy(bestCandidate.schedule,fixedSchedules,activityGroups),periods)]
            .map(([teacherId,stat])=>({teacherId,teacherName:teachers.find(t=>t.id===teacherId)?.name||teacherId,...stat}))
            .sort((a,b)=>b.total-a.total),
          candidates,
          candidateIndex:0,
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

  // Switch among up to three candidate previews without writing Firebase.
  const chooseScheduleCandidate = index => {
    setScheduleResult(previous => {
      if (!previous || !previous.candidates?.[index]) return previous;
      const candidate=previous.candidates[index];
      const allOccupied=buildTeacherOccupancy(candidate.schedule,fixedSchedules,activityGroups);
      const missing=candidate.unassigned.reduce((a,x)=>a+(Number(x.missing)||0),0);
      const totalReq=teachingLoads.reduce((a,x)=>a+(Number(x.periods)||0),0);
      return {...previous,...candidate,candidates:previous.candidates,candidateIndex:index,
        previewSource:previous.previewSource,bestFourPeriodMissing:previous.bestFourPeriodMissing,
        fourPeriodCandidatesTried:previous.fourPeriodCandidatesTried,
        fivePeriodDays:findTeacherFifthPeriodDays(allOccupied,periods,teachers),
        weeklyLoadSummary:[...summarizeTeacherWeek(allOccupied,periods)].map(([teacherId,stat])=>
          ({teacherId,teacherName:teachers.find(t=>t.id===teacherId)?.name||teacherId,...stat})),
        stats:{totalReq,scheduled:Math.max(0,totalReq-missing),missing,
          preserved:candidate.preservedCount,generated:candidate.generatedSchedule.length}};
    });
  };

  const commitSchedule = async () => {
    if (!scheduleResult) return;
    setConfirmData({
      type: 'restore', title: 'ยืนยันการใช้ตารางใหม่',
      message: `ตารางเดิมในเทอมนี้จะถูกแทนที่ ยืนยันหรือไม่?${scheduleResult.fivePeriodDays?.length ? `\n\nอนุมัติข้อยกเว้น 5 คาบ/วัน: ${scheduleResult.fivePeriodDays.map(x=>`ครู ${x.teacherName} วัน${x.day}`).join(', ')}\nระบบลองค้นหาแบบ 4 คาบ ${scheduleResult.fourPeriodCandidatesTried||0} ตัวเลือกแล้ว แต่ยังไม่พบผลครบ การค้นหาไม่ใช่หลักฐานว่าไม่มีตารางแบบ 4 คาบ` : ''}`,
      confirmText: 'บันทึกตาราง',
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
            const isTermData = ['schedules', 'fixedSchedules', 'teachingLoads', 'unavailabilities', 'classroomUnavailabilities', 'scheduleRules', 'activityGroups', 'plcSessions'].includes(name);
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
  // Block an entire school day in one atomic batch; never partially lock a room.
  const publishDraftSchedule = () => {
    if(!rulesReady)return showToast('กรุณารอโหลดข้อมูลภาคเรียนให้ครบก่อนประกาศตาราง','error');
    const termKey=getTermKey(schoolSettings);
    const todayStatus=calculateScheduleHealth(schedules,teachingLoads,unavailabilities,classroomUnavailabilities,
      fixedSchedules,teachers,subjects,classrooms,periods,customRules,activityGroups,plcSessions);
    if(todayStatus.status==='error'||todayStatus.status==='empty'||todayStatus.missing>0||todayStatus.overScheduled>0)
      return showToast('ยังมีข้อขัดแย้งหรือภาระสอนที่จัดไม่ครบ ไม่สามารถประกาศใช้ได้','error');
    const days5=findTeacherFifthPeriodDays(buildTeacherOccupancy(schedules,fixedSchedules,activityGroups),periods,teachers);
    setConfirmData({type:'restore',title:'ประกาศใช้ตารางจากฉบับร่าง',confirmText:'ยืนยันประกาศใช้',
      message:`เผยแพร่ตารางฉบับร่างของภาคเรียน ${termKey} เป็นฉบับประกาศใช้แยกต่างหาก
หลังจากนี้แก้ร่างต่อได้โดยไม่กระทบฉบับประกาศ${days5.length?`
ข้อยกเว้น 5 คาบ: ${days5.map(x=>x.teacherName+' วัน'+x.day).join(', ')}`:''}`,
      action:async()=>{
        setConfirmData(prev=>({...prev,isLoading:true}));
        try{
          const path=`school_data/${appId}/terms/${termKey}`;
          // Stage 23: PLC is part of the exact published timetable, not only a client-side check.
          const [ss,fs,gs,ps]=await Promise.all(['schedules','fixedSchedules','activityGroups','plcSessions'].map(col=>
            getDocsFromServer(collection(db,`${path}/${col}`))));
          const fromServer=[ss,fs,gs,ps].map(snapshot=>snapshot.docs.map(x=>({...x.data(),id:x.id})));
          if(getTermKey(schoolSettings)!==termKey||
            fingerprintScheduleRows(fromServer[0])!==fingerprintScheduleRows(schedules)||
            fingerprintScheduleRows(fromServer[1])!==fingerprintScheduleRows(fixedSchedules)||
            fingerprintScheduleRows(fromServer[2])!==fingerprintScheduleRows(activityGroups)||
            fingerprintScheduleRows(fromServer[3])!==fingerprintScheduleRows(plcSessions))
            throw new Error('ข้อมูลเปลี่ยนจากเครื่องอื่น กรุณาตรวจตารางและประกาศใหม่');
          const masterBase=`school_data/${appId}`;
          const refNames=['teachers','subjects','classrooms','periods','teachingLoads','unavailabilities','classroomUnavailabilities','scheduleRules'];
          const refs=await Promise.all(refNames.map(name=>getDocsFromServer(collection(db,
            ['teachers','subjects','classrooms','periods'].includes(name)?`${masterBase}/${name}`:`${path}/${name}`))));
          const frozen=Object.fromEntries(refNames.map((name,i)=>[name,refs[i].docs.map(d=>({...d.data(),id:d.id}))]));
          if(!frozen.periods.length)frozen.periods=DEFAULT_PERIODS;
          // The publication reference includes the same PLC reservations as its history snapshot.
          frozen.plcSessions=fromServer[3];
          frozen.schoolSettings={...schoolSettings};
          // The displayed catalog must match the catalog used to approve this publication.
          for(const name of ['teachers','subjects','classrooms','periods','teachingLoads','unavailabilities','classroomUnavailabilities','scheduleRules','plcSessions']) {
            const local={teachers,subjects,classrooms,periods,teachingLoads,unavailabilities,
              classroomUnavailabilities,scheduleRules:customRules,plcSessions}[name];
            if(fingerprintScheduleRows(frozen[name])!==fingerprintScheduleRows(local))
              throw new Error('ข้อมูลพื้นฐานหรือเงื่อนไขเปลี่ยนขณะประกาศ กรุณาตรวจข้อมูลใหม่');
          }
          // Validate the frozen publication with the frozen PLC and all its teaching references.
          const publishHealth=calculateScheduleHealth(fromServer[0],frozen.teachingLoads,
            frozen.unavailabilities,frozen.classroomUnavailabilities,fromServer[1],
            frozen.teachers,frozen.subjects,frozen.classrooms,frozen.periods,
            frozen.scheduleRules,fromServer[2],frozen.plcSessions);
          if(publishHealth.status==='error'||publishHealth.status==='empty'||publishHealth.missing>0||publishHealth.overScheduled>0)
            throw new Error('ตารางหรือ PLC มีข้อขัดแย้ง กรุณาตรวจและจัดใหม่ก่อนประกาศ');
          const versionId=await saveScheduleVersion('published',fromServer[0],fromServer[1],termKey,fromServer[2],frozen,frozen.plcSessions);
          if(!versionId)throw new Error('สร้างฉบับประกาศไม่สำเร็จ');
          await setDoc(doc(db,`${path}/config/publishedTimetable`),{
            versionId,publishedAt:Date.now(),termKey});
          showToast('ประกาศใช้ตารางสำเร็จ (ตารางฉบับร่างยังแก้ได้)');
          setConfirmData(null);
        }catch(error){console.error(error);showToast('ประกาศไม่สำเร็จ: '+error.message,'error');setConfirmData(null);}
      }});
  };

  const blockClassroomDays = async (roomIds,day,reason='') => {
    const unique=[...new Set((roomIds||[]).filter(id=>classrooms.some(c=>c.id===id)))];
    if(!unique.length||!DAYS.includes(day)) return showToast('กรุณาเลือกห้องและวันที่ถูกต้อง','error');
    if(!window.confirm(`ยืนยันปิดวัน${day} ทุกคาบเรียนสำหรับ ${unique.length} ห้องหรือไม่?\nระบบจะไม่ปิดบางห้องหากห้องใดขัดกับตารางที่มีอยู่`))return;
    const termKey=getTermKey(schoolSettings);
    try{
      const termPath=`school_data/${appId}/terms/${termKey}`;
      const [ss,fs,us]=await Promise.all(['schedules','fixedSchedules','classroomUnavailabilities'].map(col=>
        getDocsFromServer(collection(db,`${termPath}/${col}`))));
      if(getTermKey(schoolSettings)!==termKey)throw new Error('ภาคเรียนเปลี่ยน กรุณาลองใหม่');
      const snapshotSchedules=ss.docs.map(d=>({...d.data(),id:d.id}));
      const snapshotFixed=fs.docs.map(d=>({...d.data(),id:d.id}));
      const snapshotUnavailable=us.docs.map(d=>({...d.data(),id:d.id}));
      const plans=unique.map(classroomId=>({classroomId,
        ...planClassroomDayClosure({classroomId,day,periods,schedules:snapshotSchedules,
          fixedSchedules:snapshotFixed,classroomUnavailabilities:snapshotUnavailable})}));
      const failed=plans.find(x=>!x.ok);
      if(failed)throw new Error(`${classrooms.find(c=>c.id===failed.classroomId)?.name||failed.classroomId}: ${failed.reason}`);
      const operations=plans.flatMap(p=>p.periodIds.map(periodId=>{
        const ref=doc(collection(db,`${termPath}/classroomUnavailabilities`));
        return {type:'set',ref,data:{id:ref.id,classroomId:p.classroomId,day,periodId,
          reason:String(reason||'ปิดทั้งวัน')}};
      }));
      if(!operations.length)return showToast('ห้องที่เลือกถูกปิดครบแล้ว');
      await commitAtomicOperations(operations);
      showToast(`ปิดวัน${day} สำเร็จ ${unique.length} ห้อง รวม ${operations.length} คาบ`);
    }catch(e){console.error('Multi-room closure failed',e);showToast('ปิดหลายห้องไม่สำเร็จ: '+e.message,'error');}
  };
  const blockClassroomDay = (roomId,day,reason='')=>blockClassroomDays([roomId],day,reason);
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
    { id: 'dashboard', name: 'หน้าแรก', icon: LayoutDashboard },
    { id: 'basicData', name: 'ข้อมูลพื้นฐาน', icon: Users },
    { id: 'loads', name: 'ภาระสอน', icon: Briefcase },
    { id: 'constraintsHub', name: 'ข้อจำกัดและเงื่อนไข', icon: Clock },
    { id: 'schedule', name: 'จัดตารางสอน', icon: CalendarDays },
    { id: 'history', name: 'ประวัติตาราง', icon: Clock },
    { id: 'reports', name: 'รายงานและ Export', icon: Activity },
    { id: 'print', name: 'พิมพ์ตาราง', icon: Printer },
    { id: 'settings', name: 'ตั้งค่าระบบ', icon: Settings }
  ];

  const commonProps = { user, appId, db, dbAction, bulkUpsertTeachingLoads, saveTeachingLoadsWithSubjects, setConfirmData, bulkDeleteSubjects, bulkDeleteFixedSchedules, clearUnlockedSchedules, renameSubjectId, blockClassroomDay, blockClassroomDays, showToast, handleRequestDelete, saveScheduleVersion, readScheduleVersion, restoreScheduleVersion, getTermKey, getTeacher, getSubject, getClassroom, getShortTeacherName, publishDraftSchedule, publishedSnapshot, publishedLoading, publishedError, teachers, subjects, activeSubjects, classrooms, periods, teachingLoads, unavailabilities, classroomUnavailabilities, fixedSchedules, schedules, customRules, rulesReady, saveSubjectTimeWindows, activityGroups, plcSessions, schoolSettings, health, fullDataForBackup: { teachers, subjects, classrooms, periods, teachingLoads, unavailabilities, classroomUnavailabilities, fixedSchedules, schedules, scheduleRules: customRules, activityGroups, plcSessions } };

  if (isInitialLoading) return <div className="flex h-screen w-full items-center justify-center bg-slate-50"><div className="animate-spin text-[#d4af37]"><RefreshCw size={48}/></div></div>;

  return (
    <div className="flex h-screen bg-[#f8fafc] font-sans text-slate-900 selection:bg-[#d4af37]/30 bhs-app-shell">
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;500;600;700&display=swap');
        .font-sans { font-family: 'Sarabun', sans-serif; }
        .custom-scrollbar::-webkit-scrollbar { width: 6px; height: 6px; }
        .custom-scrollbar::-webkit-scrollbar-track { background: transparent; }
        .custom-scrollbar::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 10px; }
        .custom-scrollbar::-webkit-scrollbar-thumb:hover { background: #081a39; }
        /* Print only the timetable: let the page use the full A4 landscape width. */
        @page { size: A4 landscape; margin: 8mm; }
        @media print {
          html, body, #root { width: 100% !important; height: auto !important; overflow: visible !important; margin: 0 !important; padding: 0 !important; background: white !important; }
          .bhs-app-shell, .bhs-main, .bhs-scroll-content, .bhs-print-view {
            display: block !important; position: static !important; width: 100% !important;
            height: auto !important; max-width: none !important; min-height: 0 !important;
            overflow: visible !important; margin: 0 !important; padding: 0 !important;
            background: white !important; box-shadow: none !important;
          }
          .bhs-app-shell > aside, .bhs-main > header, .no-print { display: none !important; }
          #print-area { display: block !important; position: static !important; left: auto !important; top: auto !important; width: 100% !important; max-width: none !important; margin: 0 !important; padding: 0 !important; overflow: visible !important; }
          #print-area .print-page {
            display: block !important; box-sizing: border-box !important; position: static !important;
            width: 100% !important; height: auto !important; min-height: 0 !important;
            padding: 0 !important; margin: 0 !important; border: 0 !important;
            border-radius: 0 !important; box-shadow: none !important;
            break-inside: avoid-page; page-break-inside: avoid;
            break-after: page; page-break-after: always;
          }
          #print-area .print-page:last-child { break-after: auto !important; page-break-after: auto !important; }
          #print-area .print-header { margin: 0 0 4mm !important; line-height: 1.2 !important; }
          #print-area .print-header h1 { font-size: 17pt !important; margin: 0 0 1.5mm !important; }
          #print-area .print-header h2 { font-size: 11.5pt !important; margin: 0 0 1mm !important; }
          #print-area .print-header h3 { font-size: 11pt !important; margin: 0 !important; }
          #print-area .print-header div { font-size: 9pt !important; margin: 0 0 1mm !important; }
          #print-area .print-table {
            width: 100% !important; max-width: 100% !important; table-layout: fixed !important;
            border-collapse: collapse !important; margin: 0 !important;
            -webkit-print-color-adjust: exact; print-color-adjust: exact;
          }
          #print-area .print-table th, #print-area .print-table td {
            box-sizing: border-box !important; border: 1px solid #111 !important;
            padding: 1.5mm 1mm !important; vertical-align: middle !important;
            text-align: center !important; font-size: 9pt !important; line-height: 1.2 !important;
            overflow-wrap: anywhere !important; word-break: break-word !important;
            white-space: normal !important;
          }
          #print-area .print-table th:first-child { width: 19mm !important; }
          #print-area .print-table thead th { height: 13mm !important; font-size: 9pt !important; }
          #print-area .print-table thead th div { font-size: 8.5pt !important; line-height: 1.15 !important; }
          #print-area .print-table tbody tr { height: 22mm !important; break-inside: avoid !important; page-break-inside: avoid !important; }
          #print-area .print-table tbody td { height: 22mm !important; }
          #print-area .print-table .truncate {
            overflow: visible !important; text-overflow: clip !important; white-space: normal !important;
          }
          #print-area .print-table td div { max-width: 100% !important; overflow-wrap: anywhere !important; }
          #print-area .print-signatures {
            display: flex !important; justify-content: space-between !important;
            margin: 8mm 0 0 !important; padding: 0 10mm !important;
            font-size: 10pt !important; line-height: 1.25 !important;
            break-inside: avoid !important;
          }
          #print-area .print-signatures > div { width: 46%; }
          #print-area .print-signatures p { margin: 0 0 1.5mm !important; }
          #print-area .print-version { margin-top: 3mm !important; text-align: right !important;
            font-size: 7.5pt !important; line-height: 1.1 !important; color: #555 !important;
            overflow-wrap: anywhere !important; }
        }
      `}</style>
      
      {isSidebarOpen && <div className="fixed inset-0 bg-slate-900/50 z-20 md:hidden backdrop-blur-sm" onClick={() => setIsSidebarOpen(false)} />}
      <aside className={`fixed md:static inset-y-0 left-0 w-64 bg-gradient-to-b from-[#081a39] to-[#153063] border-r border-[#d4af37]/20 text-white flex flex-col z-30 transform transition-transform duration-300 shadow-xl ${isSidebarOpen ? 'translate-x-0' : '-translate-x-full md:translate-x-0'}`}>
        <div className="p-6 border-b border-[#122b5e]">
          <h1 className="text-xl font-bold flex items-center gap-3 text-white"><div className="p-2 bg-gradient-to-br from-[#d4af37] to-[#b8952b] rounded-lg shadow-md"><CalendarDays className="text-[#081a39]" size={24}/></div> BHS Scheduler</h1>
        </div>
        <nav className="flex-1 overflow-y-auto py-6 custom-scrollbar space-y-1 px-3">
          {TABS.map(tab => {
            const IconComponent = tab.icon;
            return <button key={tab.id} onClick={() => { setDashboardTargetRoom(''); if (tab.id === 'constraintsHub') setConstraintsEntrySection('time'); setActiveTab(tab.id); setIsSidebarOpen(false); }} className={`w-full flex items-center gap-3 px-4 py-3 rounded-xl transition-all font-medium ${activeTab === tab.id ? 'bg-gradient-to-r from-[#d4af37] to-[#e6c55c] text-[#081a39] shadow-md font-bold' : 'text-slate-300 hover:text-white hover:bg-white/10'}`}><IconComponent size={20} className={activeTab === tab.id ? "text-[#081a39]" : "opacity-80"} /> {tab.name}</button>;
          })}
        </nav>
      </aside>

      <main className="flex-1 flex flex-col min-w-0 overflow-hidden relative bhs-main">
        <header className="md:hidden bg-white border-b border-[#081a39]/10 p-3 flex items-center shadow-sm z-10 no-print">
          <button onClick={() => setIsSidebarOpen(true)} className="p-2 text-[#081a39] hover:bg-slate-100 rounded-lg flex items-center gap-2 font-bold transition-colors"><Menu size={24} /> เมนู</button>
        </header>
        <div className="flex-1 overflow-y-auto p-4 sm:p-8 custom-scrollbar bhs-scroll-content">
          {activeTab === 'dashboard' && <Dashboard {...commonProps} onNavigate={(tab) => { setDashboardTargetRoom(''); setActiveTab(tab); }} onOpenRoom={(roomId) => { setDashboardTargetRoom(roomId); setActiveTab('schedule'); }} />}
          {activeTab === 'settings' && <SettingsView {...commonProps} handleRestoreBackup={handleRestoreBackup} />}
          {activeTab === 'basicData' && <BasicDataHub commonProps={commonProps} autoGenerateClassrooms={autoGenerateClassrooms} onOpenDevelopment={() => { setConstraintsEntrySection('development'); setDashboardTargetRoom(''); setActiveTab('constraintsHub'); setIsSidebarOpen(false); }} />}
          {activeTab === 'loads' && <LoadsView {...commonProps} />}
          {activeTab === 'constraintsHub' && <ConstraintsHub commonProps={commonProps} initialSection={constraintsEntrySection} />}
          {activeTab === 'schedule' && <ScheduleView {...commonProps} initialRoomId={dashboardTargetRoom} runSmartSchedule={runSmartSchedule} isScheduling={isScheduling} />}
          {activeTab === 'history' && <ScheduleHistoryView key={getTermKey(schoolSettings)} {...commonProps} />}
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
            <div className="border border-indigo-200 rounded-xl p-4 space-y-3">
              <h3 className="font-bold text-indigo-900">ระบบเลือกตารางที่เหมาะสมที่สุดที่ค้นพบให้แล้ว</h3>
              <p className="text-sm text-slate-700">{scheduleResult.candidateIndex===0
                ? 'กำลังแสดงแบบที่ระบบแนะนำ (แบบที่ 1) ตรวจผลด้านล่างแล้วกด “ใช้งานตารางนี้” ได้เลย'
                : `คุณเลือกแบบที่ ${scheduleResult.candidateIndex+1} แทนแบบที่ระบบแนะนำ`}</p>
              <p className="text-xs text-slate-500">ระบบจัดอันดับจากความครบของคาบ จำนวนวันที่มีคาบสอนรวมกิจกรรม 5 คาบ และความสมดุลของภาระครู โดยยังให้ตรวจสอบก่อนบันทึกจริง</p>
              {(scheduleResult.candidates||[]).length>1&&<details className="border border-indigo-100 rounded-lg p-3 bg-white">
                <summary className="cursor-pointer font-bold text-sm text-indigo-800">ไม่พอใจแบบที่เลือก? ดูและเปลี่ยนตัวเลือกอื่น ({scheduleResult.candidates.length} แบบ)</summary>
                <div className="grid grid-cols-1 sm:grid-cols-3 gap-2 mt-3">
                  {(scheduleResult.candidates||[]).map((option,index)=><button key={index} type="button"
                    onClick={()=>chooseScheduleCandidate(index)} className={`rounded-xl border p-3 text-left text-sm ${scheduleResult.candidateIndex===index?'border-indigo-600 bg-indigo-50':'border-slate-200 hover:bg-slate-50'}`}>
                    <strong>แบบ {index+1}{index===0?' (ระบบแนะนำ)':''}</strong><br/>ขาด {option.missingCount} คาบ<br/>
                    วันภาระรวม 5 คาบ {option.fivePeriodCount} วัน<br/>
                    วันหนัก 4–5 คาบ {option.weeklyBalance?.heavyDays??'-'} วัน<br/>
                    คาบว่างคั่น {option.weeklyBalance?.gaps??'-'} ช่อง
                  </button>)}
                </div>
              </details>}
            </div>
            <div className="grid grid-cols-2 gap-4">
              <div className="bg-emerald-50 border border-emerald-200 p-4 rounded-xl text-center"><div className="text-emerald-700 font-bold text-sm mb-1">จัดลงสำเร็จ</div><div className="text-2xl font-bold text-emerald-800">{scheduleResult.stats.scheduled} <span className="text-sm">/ {scheduleResult.stats.totalReq} คาบ</span></div></div>
              <div className={`border p-4 rounded-xl text-center ${scheduleResult.stats.missing > 0 ? 'bg-rose-50 border-rose-200' : 'bg-slate-50 border-slate-200'}`}><div className={`font-bold text-sm mb-1 ${scheduleResult.stats.missing > 0 ? 'text-rose-700' : 'text-slate-600'}`}>จัดไม่ได้</div><div className={`text-2xl font-bold ${scheduleResult.stats.missing > 0 ? 'text-rose-800' : 'text-slate-800'}`}>{scheduleResult.stats.missing} <span className="text-sm">คาบ</span></div></div>
            </div>
            {!!scheduleResult.fivePeriodDays?.length && (
              <div className="bg-amber-50 border-2 border-amber-300 rounded-xl p-4 space-y-2 text-amber-900">
                <div className="font-bold flex items-center gap-2"><AlertTriangle size={18}/> ต้องยืนยันข้อยกเว้น: มีครูสอนรวมกิจกรรม 5 คาบต่อวัน</div>
                <p className="text-xs">ระบบให้ความสำคัญกับ 4 คาบก่อนเสมอ คาบที่ 5 อาจเป็นคาบล็อกเดิมหรือมาจากรอบสำรองเมื่อรอบ 4 คาบยังจัดไม่ครบ กรุณาตรวจสอบและอนุมัติก่อนบันทึก (การค้นหาไม่รับประกันว่ารูปแบบ 4 คาบไม่มีทางทำได้)</p>
                <p className="text-xs mb-2">รอบทดลองจำกัด 4 คาบยังขาดอย่างน้อย {scheduleResult.bestFourPeriodMissing ?? '-'} คาบจากตัวเลือกที่ค้นพบ (ไม่ใช่หลักฐานว่าไม่มีวิธีจัดครบโดยใช้ 4 คาบ)</p>
                <ul className="list-disc pl-5 text-sm">{scheduleResult.fivePeriodDays.map(x=><li key={`${x.teacherId}_${x.day}`}>{x.teacherName} — วัน{x.day} ({x.count} คาบ) · ต้องตรวจและยืนยันก่อนบันทึก</li>)}</ul>
              </div>
            )}
            {scheduleResult.unassigned.length > 0 && (
              <div className="border border-rose-200 rounded-xl overflow-hidden">
                <div className="bg-rose-50 p-3 font-bold text-rose-700 flex items-center gap-2 border-b border-rose-200"><AlertTriangle size={18}/> รายการที่จัดไม่ได้ (รอแก้ไข/จัดด้วยมือ)</div>
                <div className="max-h-48 overflow-y-auto bg-white p-3 space-y-2">
                  {scheduleResult.unassigned.map((u, i) => (
                    <div key={i} className="text-sm border-b pb-3 last:border-0 last:pb-0 space-y-2">
                      <div className="flex justify-between gap-3"><div><span className="font-bold text-[#081a39]">{getSubject(u.subjectId).name}</span> ห้อง {getClassroom(u.classroomId).name} · {getTeacher(u.teacherId).name}</div><div className="text-rose-600 font-bold whitespace-nowrap">ขาด {u.missing} คาบ</div></div>
                      <details className="bg-slate-50 border border-slate-200 rounded-lg p-3">
                        <summary className="font-bold cursor-pointer text-[#081a39]">ดูสาเหตุที่เป็นไปได้และแนวทางแก้</summary>
                        {(() => {
                          const diag = describeUnassignedReasons({ load:u, schedules:scheduleResult.schedule,
                            fixedSchedules, periods, unavailabilities, classroomUnavailabilities, customRules, activityGroups, classrooms, plcSessions, teachers, subjects });
                          return <div className="mt-2 space-y-2 text-slate-700">
                            <p>ช่องเดี่ยวที่วางเพิ่มได้ในตารางผลลัพธ์ปัจจุบัน: {diag.singleOptions} ช่อง</p>
                            {u.reason !== 'NO_AVAILABLE_SLOT' && <p>ผลการวางคาบคู่ที่ว่างพร้อมกัน: {diag.pairOptions} คู่</p>}
                            <p className="font-semibold">สาเหตุที่พบจากช่องซึ่งถูกปฏิเสธ (มากสุด 3 กลุ่ม):</p>
                            {diag.topReasons.length ? <ul className="list-disc pl-5">{diag.topReasons.map(r=><li key={r.code}>{r.label} — {r.count} ช่อง</li>)}</ul> : <p>ยังมีช่องเดี่ยวที่ว่างบางส่วน อาจต้องจัดลำดับหรือจับคู่คาบใหม่</p>}
                            <p className="text-blue-900 font-medium">ข้อเสนอ: {diag.suggestion}</p>
                            {(()=>{const tips=suggestSingleSlotRepair({load:u,schedules:scheduleResult.schedule,
                              fixedSchedules,activityGroups,periods,unavailabilities,classroomUnavailabilities,
                              customRules,subjects,classrooms,teachers,teachingLoads,plcSessions});
                              return tips.length?<div className="bg-blue-50 rounded-lg p-2 space-y-2">
                                <strong>ตัวอย่างย้ายคาบ Auto เดี่ยวเพื่อเปิดช่อง (ยังไม่แก้จริง)</strong>
                                {tips.map((t,i)=><p key={i}>ลองย้าย {getSubject(t.from.subjectId).name} ห้อง {getClassroom(t.from.classroomId).name} จากวัน{t.from.day} {getPeriodDisplayName(periods,t.from.periodId)} ไปวัน{t.to.day} {getPeriodDisplayName(periods,t.to.periodId)} แล้วใช้ช่องเดิมกับวิชาที่ขาด</p>)}
                                <button className="underline font-bold" onClick={()=>{setScheduleResult(null);setDashboardTargetRoom(u.classroomId);setActiveTab('schedule');}}>เปิดตารางห้องนี้เพื่อตรวจด้วยมือ</button>
                              </div>:null;
                            })()}
                            <p className="text-xs text-slate-500">เป็นการอธิบายจากสภาพตารางผลลัพธ์ ไม่ใช่ข้อสรุปว่าสาเหตุนั้นทำให้จัดไม่ครบแน่นอน</p>
                          </div>;
                        })()}
                      </details>
                    </div>
                  ))}
                </div>
              </div>
            )}
            <div className="flex gap-3 pt-4 border-t border-slate-100"><button onClick={() => setScheduleResult(null)} className="flex-1 py-3 bg-slate-100 hover:bg-slate-200 text-[#081a39] rounded-xl font-bold transition-all">ยกเลิก (คงตารางเดิม)</button><button onClick={commitSchedule} className="flex-1 py-3 bg-[#081a39] hover:bg-[#122b5e] text-white rounded-xl font-bold shadow-md flex justify-center items-center gap-2"><Check size={18}/> ใช้งานตารางนี้</button></div>
          </div>
        )}
      </Modal>

      {toast && (
        <div role={toast.type === 'error' ? 'alert' : 'status'} className={`fixed bottom-6 right-6 max-w-[min(92vw,620px)] whitespace-pre-line break-words px-6 py-4 rounded-xl shadow-xl border flex items-start gap-3 animate-in slide-in-from-bottom-5 z-[70] font-medium ${toast.type === 'error' ? 'bg-rose-50 text-rose-700 border-rose-200' : 'bg-[#081a39] text-[#d4af37] border-[#122b5e]'}`}>
          {toast.type === 'error' ? <AlertCircle size={22} /> : <CheckCircle2 size={22} />} {String(toast.msg)}
        </div>
      )}
      <ConfirmModal confirmData={confirmData} setConfirmData={setConfirmData} />
    </div>
  );
}
