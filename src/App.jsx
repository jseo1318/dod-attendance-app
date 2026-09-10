import React, { useState, useEffect, useMemo, useRef, useCallback } from "react";
import * as XLSX from "xlsx";
import { supabase } from "./supabaseClient";

/* ───────────────────────── constants ───────────────────────── */

const FONT =
  '-apple-system, BlinkMacSystemFont, "Malgun Gothic", "Apple SD Gothic Neo", "Noto Sans KR", sans-serif';

const COLORS = {
  bg: "#F5F7F6",
  card: "#FFFFFF",
  border: "#DCE3E0",
  text: "#1F2A27",
  sub: "#5E6C68",
  teal: "#0F5C55",
  tealDark: "#0A3E3A",
  tealSoft: "#E4F0EE",
  amber: "#B8860B",
  amberSoft: "#FBF1DC",
  red: "#B4432F",
  redSoft: "#FBEAE6",
};

const SCHEDULE = {
  0: null,
  1: { start: "09:30", end: "19:00" },
  2: { start: "09:30", end: "19:00" },
  3: { start: "09:30", end: "19:00" },
  4: { start: "09:30", end: "19:00" },
  5: { start: "09:30", end: "21:00" },
  6: { start: "09:30", end: "16:00" },
};
const OT_MULTIPLIER = 1.5;
const DAY_MINUTES = 480;

const TEAM_ORDER = ["상담팀", "코디팀", "간호팀", "피부팀", "씨&마", "진료팀", "미지정"];
const TEAM_COLORS = {
  상담팀: "#5B4B63",
  코디팀: "#0F5C55",
  간호팀: "#3E5C76",
  피부팀: "#55624A",
  "씨&마": "#8B5E45",
  진료팀: "#2E3A59",
  미지정: "#5E6C68",
};
const POSITION_LIST = ["원장", "실장", "팀장", "부팀장", "사원"];
const POSITION_RANK = { 원장: 0, 실장: 1, 팀장: 2, 부팀장: 3, 사원: 4 };
const WEEKDAY_ORDER = [1, 2, 3, 4, 5, 6];
const WEEKDAY_LABELS = { 1: "월", 2: "화", 3: "수", 4: "목", 5: "금", 6: "토" };

/* ───────────────────────── helpers ───────────────────────── */

function pad(n) {
  return String(n).padStart(2, "0");
}
function toISO(d) {
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
}
function fmtDate(iso) {
  if (!iso) return "-";
  const [y, m, d] = iso.split("-");
  return `${y}.${m}.${d}`;
}
function timeToMinutes(hhmm) {
  if (!hhmm) return null;
  const m = String(hhmm).match(/(\d{1,2}):(\d{2})/);
  if (!m) return null;
  return parseInt(m[1], 10) * 60 + parseInt(m[2], 10);
}
function minutesToHM(min) {
  const sign = min < 0 ? "-" : "";
  const a = Math.abs(Math.round(min));
  return `${sign}${Math.floor(a / 60)}시간 ${a % 60}분`;
}
function minutesToDaysLabel(min) {
  const days = min / DAY_MINUTES;
  return `${days.toFixed(2).replace(/\.00$/, "")}일`;
}

function normalizeTimeStr(v) {
  if (v === undefined || v === null || v === "") return "";
  if (typeof v === "number") {
    const totalMin = Math.round(v * 24 * 60);
    if (totalMin <= 0 || totalMin > 1440) return "";
    return `${pad(Math.floor(totalMin / 60))}:${pad(totalMin % 60)}`;
  }
  const s = String(v).trim();
  const m = s.match(/(\d{1,2}):(\d{2})/);
  if (m) return `${pad(parseInt(m[1], 10))}:${pad(parseInt(m[2], 10))}`;
  return "";
}
function excelSerialToDate(serial) {
  const utcDays = Math.floor(serial - 25569);
  return new Date(utcDays * 86400 * 1000);
}
function normalizeDateStr(v) {
  if (v instanceof Date && !isNaN(v)) return toISO(v);
  if (typeof v === "number") {
    const d = excelSerialToDate(v);
    if (!isNaN(d)) return toISO(d);
  }
  const s = String(v).trim().replace(/\./g, "-").replace(/\//g, "-");
  const m = s.match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return `${m[1]}-${pad(parseInt(m[2], 10))}-${pad(parseInt(m[3], 10))}`;
  return "";
}
function normalizeEmpId(v, name) {
  const s = String(v || "").trim();
  if (/^\d+$/.test(s) && s.length < 4) return s.padStart(4, "0");
  return s || name;
}

function cumulativeLeaveDays(hireISO, asOfISO) {
  if (!hireISO) return { days: 0, serviceYears: 0, note: "입사일 미입력" };
  const hire = new Date(hireISO + "T00:00:00");
  const asOf = new Date(asOfISO + "T00:00:00");
  if (asOf < hire) return { days: 0, serviceYears: 0, note: "입사 전" };

  let years = asOf.getFullYear() - hire.getFullYear();
  const anniversaryThisYear = new Date(asOf.getFullYear(), hire.getMonth(), hire.getDate());
  if (asOf < anniversaryThisYear) years -= 1;
  if (years < 0) years = 0;

  if (years === 0) {
    let months =
      (asOf.getFullYear() - hire.getFullYear()) * 12 + (asOf.getMonth() - hire.getMonth());
    if (asOf.getDate() < hire.getDate()) months -= 1;
    months = Math.max(0, Math.min(months, 11));
    return { days: months, serviceYears: 0, note: `1년차 · 개근월차 ${months}일 발생` };
  }
  let total = 11;
  for (let k = 1; k <= years; k++) {
    total += Math.min(15 + Math.floor((k - 1) / 2), 25);
  }
  return { days: total, serviceYears: years, note: `근속 ${years}년차` };
}

function monthEndExclusive(month) {
  const [y, m] = month.split("-").map(Number);
  return toISO(new Date(y, m, 1));
}

function computeMetrics(record, employee) {
  const sched = SCHEDULE[record.dow];
  if (!sched) return { late: 0, otRaw: 0, otCredited: 0 };
  const custom = (employee && employee.customSchedule && employee.customSchedule[record.dow]) || {};
  const startStr = custom.start || sched.start;
  const endStr = custom.end || sched.end;
  let late = 0;
  if (record.checkin) {
    const inMin = timeToMinutes(record.checkin);
    const startMin = timeToMinutes(startStr);
    if (inMin !== null && inMin > startMin) late = inMin - startMin;
  }
  let otRaw = 0;
  let otCredited = 0;
  if (record.checkout) {
    const outMin = timeToMinutes(record.checkout);
    const endMin = timeToMinutes(endStr);
    if (outMin !== null && outMin > endMin) {
      otRaw = outMin - endMin;
      otCredited = Math.floor(otRaw * OT_MULTIPLIER);
    }
  }
  return { late, otRaw, otCredited };
}

const STATUS_LABELS = { 지각: "지각", 늦출: "늦출", 일퇴: "일퇴", 반차: "반차", OFF: "OFF", 출근: "출근" };
const STATUS_COLORS = {
  지각: { bg: "#FBEAE6", fg: "#B4432F" },
  늦출: { bg: "#E4F0EE", fg: "#0F5C55" },
  일퇴: { bg: "#FBF1DC", fg: "#B8860B" },
  반차: { bg: "#EDE7F6", fg: "#6B4C93" },
  OFF: { bg: "#EDEFEE", fg: "#5E6C68" },
  출근: { bg: "#E4F0EE", fg: "#0A3E3A" },
};

/* 요일 기본 스케줄 기준 "하루치" 시간(분) = 근무시간 - 휴게 90분, 토요일은 고정 6시간(360분) */
function offDayMinutes(dow, employee) {
  if (dow === 6) return 360;
  const sched = SCHEDULE[dow];
  if (!sched) return 0;
  const custom = (employee && employee.customSchedule && employee.customSchedule[dow]) || {};
  const start = custom.start || sched.start;
  const end = custom.end || sched.end;
  const raw = timeToMinutes(end) - timeToMinutes(start);
  return Math.max(0, raw - 90);
}

/* 특정 직원의 특정 날짜 상태를 계산: 수동 기록(override)이 있으면 그걸 우선 사용,
   없으면 원본 근태 데이터로부터 자동 산출 (일퇴/OFF/출근은 자동 감지하지 않음 — 순수 수동 카테고리) */
function computeDayStatus(employee, dateISO, record, override) {
  if (override) {
    return { status: override.status, source: "manual", minutes: override.minutes, note: override.note };
  }
  const dow = new Date(dateISO + "T00:00:00").getDay();
  const sched = SCHEDULE[dow];
  if (!sched) return { status: "OFF", source: "auto", minutes: null };
  if (!record || !record.checkin) return { status: "OFF", source: "auto", minutes: null };

  const m = computeMetrics(record, employee);
  if (m.late > 0) {
    return { status: record.excused ? "늦출" : "지각", source: "auto", minutes: m.late };
  }
  return { status: "출근", source: "auto", minutes: null };
}

/* ───────────────────────── xlsx parsing ───────────────────────── */

function parseCapsWorkbook(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: "array", cellDates: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: "" });

  let headerIdx = -1;
  let col = {};
  for (let i = 0; i < Math.min(rows.length, 10); i++) {
    const r = rows[i].map((c) => String(c).trim());
    if (r.includes("이름") && r.includes("출근")) {
      headerIdx = i;
      ["번호", "사용자ID", "사원번호", "이름", "근무일자", "근무일명칭", "출근", "퇴근"].forEach((h) => {
        const idx = r.indexOf(h);
        if (idx >= 0) col[h] = idx;
      });
      break;
    }
  }
  if (headerIdx === -1) {
    throw new Error("헤더(이름, 출근 등)를 찾을 수 없습니다. 캡스 원본 형식인지 확인해주세요.");
  }

  const records = [];
  const employeesFound = new Map();

  for (let i = headerIdx + 1; i < rows.length; i++) {
    const r = rows[i];
    const name = col["이름"] !== undefined ? String(r[col["이름"]] || "").trim() : "";
    if (!name) continue;
    const rawId = col["사원번호"] !== undefined ? r[col["사원번호"]] : "";
    const empId = normalizeEmpId(rawId, name);
    const dateISO = col["근무일자"] !== undefined ? normalizeDateStr(r[col["근무일자"]]) : "";
    if (!dateISO) continue;
    const checkin = col["출근"] !== undefined ? normalizeTimeStr(r[col["출근"]]) : "";
    const checkout = col["퇴근"] !== undefined ? normalizeTimeStr(r[col["퇴근"]]) : "";

    employeesFound.set(empId, name);
    const dow = new Date(dateISO + "T00:00:00").getDay();

    records.push({ employee_id: empId, employee_name: name, date: dateISO, dow, checkin, checkout });
  }
  return { records, employeesFound };
}

/* ───────────────────────── db <-> app mapping ───────────────────────── */

function empFromRow(row) {
  return {
    id: row.id,
    name: row.name,
    team: row.team || "미지정",
    position: row.position || "",
    hireDate: row.hire_date || "",
    openingLeaveMinutes: row.opening_leave_minutes || 0,
    openingOTMinutes: row.opening_ot_minutes || 0,
    customSchedule: row.custom_schedule || {},
    active: row.active !== false,
  };
}
function empToRow(e) {
  return {
    id: e.id,
    name: e.name,
    team: e.team || "미지정",
    position: e.position || "",
    hire_date: e.hireDate || null,
    opening_leave_minutes: e.openingLeaveMinutes || 0,
    opening_ot_minutes: e.openingOTMinutes || 0,
    custom_schedule: e.customSchedule || {},
    active: e.active !== false,
  };
}

/* ───────────────────────── main component ───────────────────────── */

export default function App() {
  const [employees, setEmployees] = useState(null);
  const [attendance, setAttendance] = useState(null); // flat array of raw records
  const [ledger, setLedger] = useState(null);
  const [uploadLog, setUploadLog] = useState(null);
  const [dayStatusOverrides, setDayStatusOverrides] = useState(null); // flat array, manual calendar overrides

  const [tab, setTab] = useState("dashboard");
  const [asOf, setAsOf] = useState(toISO(new Date()));
  const [error, setError] = useState("");
  const [uploadMsg, setUploadMsg] = useState("");
  const [uploading, setUploading] = useState(false);
  const fileRef = useRef(null);
  const [collapsed, setCollapsed] = useState({});
  const [selectedId, setSelectedId] = useState(null);
  const [session, setSession] = useState(undefined); // undefined = loading, null = logged out
  const [showLogin, setShowLogin] = useState(false);

  const [ledgerEmp, setLedgerEmp] = useState("");
  const [ledgerType, setLedgerType] = useState("leave");
  const [ledgerDir, setLedgerDir] = useState("use");
  const [ledgerMinutes, setLedgerMinutes] = useState("480");
  const [ledgerDate, setLedgerDate] = useState(toISO(new Date()));
  const [ledgerNote, setLedgerNote] = useState("");

  const [newName, setNewName] = useState("");
  const [newTeam, setNewTeam] = useState(TEAM_ORDER[0]);
  const [newPosition, setNewPosition] = useState(POSITION_LIST[3]);
  const [newHire, setNewHire] = useState("");

  const configOk = !!(import.meta.env.VITE_SUPABASE_URL && import.meta.env.VITE_SUPABASE_ANON_KEY);

  /* ---- fetchers ---- */
  const fetchEmployees = useCallback(async () => {
    const { data, error: err } = await supabase.from("employees").select("*").order("name");
    if (err) {
      console.error("employees fetch error:", err);
      setError(`직원 목록을 불러오지 못했습니다: ${err.message}`);
      setEmployees((prev) => prev || []);
      return;
    }
    setEmployees((data || []).map(empFromRow));
  }, []);

  const fetchAttendance = useCallback(async () => {
    const { data, error: err } = await supabase.from("attendance").select("*");
    if (err) {
      console.error("attendance fetch error:", err);
      setError(`근태 데이터를 불러오지 못했습니다: ${err.message}`);
      setAttendance((prev) => prev || []);
      return;
    }
    setAttendance(
      (data || []).map((r) => ({
        employeeId: r.employee_id,
        employeeName: r.employee_name,
        date: r.date,
        dow: r.dow,
        checkin: r.checkin,
        checkout: r.checkout,
        excused: !!r.excused,
      }))
    );
  }, []);

  const fetchLedger = useCallback(async () => {
    const { data, error: err } = await supabase
      .from("ledger")
      .select("*")
      .order("created_at", { ascending: false });
    if (err) {
      console.error("ledger fetch error:", err);
      setError(`원장 데이터를 불러오지 못했습니다: ${err.message}`);
      setLedger((prev) => prev || []);
      return;
    }
    setLedger(
      (data || []).map((r) => ({
        id: r.id,
        employeeId: r.employee_id,
        type: r.type,
        direction: r.direction,
        minutes: r.minutes,
        date: r.date,
        note: r.note || "",
      }))
    );
  }, []);

  const fetchUploadLog = useCallback(async () => {
    const { data, error: err } = await supabase
      .from("upload_log")
      .select("*")
      .order("uploaded_at", { ascending: false })
      .limit(50);
    if (err) {
      console.error("upload_log fetch error:", err);
      setError(`업로드 이력을 불러오지 못했습니다: ${err.message}`);
      setUploadLog((prev) => prev || []);
      return;
    }
    setUploadLog(
      (data || []).map((r) => ({
        id: r.id,
        fileName: r.file_name,
        uploadedAt: r.uploaded_at,
        months: r.months,
        rowCount: r.row_count,
        removed: r.removed,
      }))
    );
  }, []);

  const fetchDayStatus = useCallback(async () => {
    const { data, error: err } = await supabase.from("day_status").select("*");
    if (err) {
      console.error("day_status fetch error:", err);
      setError(`캘린더 수동 기록을 불러오지 못했습니다: ${err.message}`);
      setDayStatusOverrides((prev) => prev || []);
      return;
    }
    setDayStatusOverrides(
      (data || []).map((r) => ({
        employeeId: r.employee_id,
        date: r.date,
        status: r.status,
        minutes: r.minutes,
        note: r.note || "",
      }))
    );
  }, []);

  function retryAll() {
    setError("");
    fetchEmployees();
    fetchAttendance();
    fetchLedger();
    fetchUploadLog();
    fetchDayStatus();
  }

  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => setSession(data.session));
    const { data: listener } = supabase.auth.onAuthStateChange((_event, newSession) => {
      setSession(newSession);
    });
    return () => listener.subscription.unsubscribe();
  }, []);

  function requireAuth() {
    if (!session) {
      setError("로그인이 필요한 작업입니다. 우측 상단에서 관리자 로그인 해주세요.");
      setShowLogin(true);
      return false;
    }
    return true;
  }

  useEffect(() => {
    if (!configOk) return;
    fetchEmployees();
    fetchAttendance();
    fetchLedger();
    fetchUploadLog();
    fetchDayStatus();

    const channel = supabase
      .channel("dod-attendance-sync")
      .on("postgres_changes", { event: "*", schema: "public", table: "employees" }, fetchEmployees)
      .on("postgres_changes", { event: "*", schema: "public", table: "attendance" }, fetchAttendance)
      .on("postgres_changes", { event: "*", schema: "public", table: "ledger" }, fetchLedger)
      .on("postgres_changes", { event: "*", schema: "public", table: "upload_log" }, fetchUploadLog)
      .on("postgres_changes", { event: "*", schema: "public", table: "day_status" }, fetchDayStatus)
      .subscribe();

    return () => {
      supabase.removeChannel(channel);
    };
  }, [configOk, fetchEmployees, fetchAttendance, fetchLedger, fetchUploadLog, fetchDayStatus]);

  /* ---- employees ---- */
  async function addEmployee() {
    if (!requireAuth()) return;
    if (!newName.trim()) return;
    const row = empToRow({
      id: `${newName.trim()}_${Date.now()}`,
      name: newName.trim(),
      team: newTeam,
      position: newPosition,
      hireDate: newHire,
      openingLeaveMinutes: 0,
      openingOTMinutes: 0,
      customSchedule: {},
      active: true,
    });
    const { error: err } = await supabase.from("employees").insert(row);
    if (err) setError(`직원 추가 실패: ${err.message}`);
    else {
      setNewName("");
      setNewHire("");
      fetchEmployees();
    }
  }

  async function updateEmployee(id, patch) {
    if (!requireAuth()) return;
    const emp = employees.find((e) => e.id === id);
    const merged = { ...emp, ...patch };
    const dbPatch = empToRow(merged);
    delete dbPatch.id;
    setEmployees(employees.map((e) => (e.id === id ? merged : e))); // optimistic
    const { error: err } = await supabase.from("employees").update(dbPatch).eq("id", id);
    if (err) {
      setError(`직원 정보 저장 실패: ${err.message}`);
      fetchEmployees();
    }
  }

  async function removeEmployee(id) {
    if (!requireAuth()) return;
    const { error: err } = await supabase.from("employees").delete().eq("id", id);
    if (err) setError(`직원 삭제 실패: ${err.message}`);
    else fetchEmployees();
  }

  /* ---- 지각을 시간차감(연차/OT 사용)으로 처리 ---- */
  /* 지각 카운트 포함 여부와 차감(연차/OT) 처리는 서로 독립적으로 관리 */
  async function setLateExcused(employeeId, date, excused) {
    if (!requireAuth()) return;
    const { error: err } = await supabase
      .from("attendance")
      .update({ excused })
      .eq("employee_id", employeeId)
      .eq("date", date);
    if (err) setError(`처리 실패: ${err.message}`);
    else fetchAttendance();
  }
  async function setLateDeduction(employeeId, date, ledgerType, minutes, note) {
    if (!requireAuth()) return;
    const ledgerRow = {
      id: `LATE_${employeeId}_${date}`,
      employee_id: employeeId,
      type: ledgerType,
      direction: "use",
      minutes: Math.abs(parseInt(minutes, 10) || 0),
      date,
      note: note || "지각 시간차감",
    };
    const { error: err } = await supabase.from("ledger").upsert(ledgerRow, { onConflict: "id" });
    if (err) setError(`차감 기록 저장 실패: ${err.message}`);
    else fetchLedger();
  }
  async function removeLateDeduction(employeeId, date) {
    if (!requireAuth()) return;
    const { error: err } = await supabase.from("ledger").delete().eq("id", `LATE_${employeeId}_${date}`);
    if (err) setError(`차감 기록 삭제 실패: ${err.message}`);
    else fetchLedger();
  }

  /* ---- 캘린더 수동 상태 설정: 일퇴/OFF/출근/(원본 없는) 지각·늦출을 분 단위로 통합 관리 ----
     day_status에는 표시용 상태+분을, ledger에는 실제 연차/OT 증감을 DAYADJ_ 키로 함께 기록 */
  async function setDayAdjustment(employeeId, date, status, dedType, minutes, note) {
    if (!requireAuth()) return;
    const m = Math.round(Math.abs(parseFloat(minutes)) || 0);
    const { error: err1 } = await supabase
      .from("day_status")
      .upsert({ employee_id: employeeId, date, status, minutes: m, note: note || null }, { onConflict: "employee_id,date" });
    if (err1) {
      setError(`캘린더 상태 저장 실패: ${err1.message}`);
      return;
    }
    if (m > 0) {
      const direction = status === "출근" ? "adjust" : "use";
      const ledgerRow = {
        id: `DAYADJ_${employeeId}_${date}`,
        employee_id: employeeId,
        type: dedType,
        direction,
        minutes: m,
        date,
        note: note || `${status} 시간조정`,
      };
      const { error: err2 } = await supabase.from("ledger").upsert(ledgerRow, { onConflict: "id" });
      if (err2) setError(`차감 기록 저장 실패: ${err2.message}`);
    } else {
      await supabase.from("ledger").delete().eq("id", `DAYADJ_${employeeId}_${date}`);
    }
    await Promise.all([fetchDayStatus(), fetchLedger()]);
  }
  async function removeDayAdjustment(employeeId, date) {
    if (!requireAuth()) return;
    const { error: err } = await supabase.from("day_status").delete().eq("employee_id", employeeId).eq("date", date);
    if (err) {
      setError(`캘린더 상태 삭제 실패: ${err.message}`);
      return;
    }
    await supabase.from("ledger").delete().eq("id", `DAYADJ_${employeeId}_${date}`);
    await Promise.all([fetchDayStatus(), fetchLedger()]);
  }
  async function deleteMonthData(month) {
    if (!requireAuth()) return;
    if (!window.confirm(`${month} 근태 데이터를 전부 삭제할까요? 되돌릴 수 없습니다.`)) return;
    const start = `${month}-01`;
    const endExclusive = monthEndExclusive(month);
    const { error: err } = await supabase
      .from("attendance")
      .delete()
      .gte("date", start)
      .lt("date", endExclusive);
    if (err) setError(`삭제 실패: ${err.message}`);
    else fetchAttendance();
  }

  /* ---- upload: 파일에 포함된 월을 통째로 교체 ---- */
  async function handleFile(file) {
    if (!requireAuth()) return;
    setUploadMsg("");
    setError("");
    setUploading(true);
    try {
      const buf = await file.arrayBuffer();
      const { records, employeesFound } = parseCapsWorkbook(buf);
      if (records.length === 0) {
        setUploadMsg("읽을 수 있는 근태 데이터가 없습니다.");
        setUploading(false);
        return;
      }

      const monthsSeen = Array.from(new Set(records.map((r) => r.date.slice(0, 7)))).sort();
      let removedCount = 0;
      for (const month of monthsSeen) {
        const start = `${month}-01`;
        const endExclusive = monthEndExclusive(month);
        const { data: existing } = await supabase
          .from("attendance")
          .select("employee_id,date")
          .gte("date", start)
          .lt("date", endExclusive);
        removedCount += existing ? existing.length : 0;
        const { error: delErr } = await supabase
          .from("attendance")
          .delete()
          .gte("date", start)
          .lt("date", endExclusive);
        if (delErr) throw delErr;
      }

      const chunkSize = 500;
      for (let i = 0; i < records.length; i += chunkSize) {
        const chunk = records.slice(i, i + chunkSize);
        const { error: insErr } = await supabase.from("attendance").insert(chunk);
        if (insErr) throw insErr;
      }

      const logRow = {
        id: `U${Date.now()}`,
        file_name: file.name,
        uploaded_at: new Date().toISOString(),
        months: monthsSeen.join(", "),
        row_count: records.length,
        removed: removedCount,
      };
      const { error: logErr } = await supabase.from("upload_log").insert(logRow);
      if (logErr) throw logErr;

      await Promise.all([fetchAttendance(), fetchUploadLog()]);

      const knownNames = new Set(employees.map((e) => e.name));
      const unknown = Array.from(employeesFound.values()).filter((name) => !knownNames.has(name));

      setUploadMsg(
        `${monthsSeen.join(", ")} 데이터를 통째로 교체했습니다 · 이번 업로드 ${records.length}건 반영 (기존 ${removedCount}건 삭제 후 재입력)` +
          (unknown.length > 0
            ? ` · 직원 명단에 없는 이름 ${unknown.length}명 발견(${unknown.join(", ")}) — 대시보드에는 표시되지 않으며, 필요하면 직원 관리 탭에서 직접 추가해주세요`
            : "")
      );
    } catch (e) {
      setError(`업로드 처리 중 오류: ${e.message || e}`);
    }
    setUploading(false);
  }

  function onFileInput(e) {
    const f = e.target.files && e.target.files[0];
    if (f) handleFile(f);
    e.target.value = "";
  }
  function onDrop(e) {
    e.preventDefault();
    const f = e.dataTransfer.files && e.dataTransfer.files[0];
    if (f) handleFile(f);
  }

  /* ---- ledger ---- */
  async function insertLedgerEntry({ employeeId, type, direction, minutes, date, note }) {
    if (!requireAuth()) return;
    const row = {
      id: `L${Date.now()}${Math.floor(Math.random() * 1000)}`,
      employee_id: employeeId,
      type,
      direction,
      minutes: Math.abs(parseInt(minutes, 10) || 0),
      date,
      note: note || "",
    };
    const { error: err } = await supabase.from("ledger").insert(row);
    if (err) setError(`기록 추가 실패: ${err.message}`);
    else fetchLedger();
  }

  /* 대휴(대체휴무): 공휴일 근무로 발생. OT/연차와 동일하게 전부 수동으로 차감/조정 */
  async function grantDaehyu(employeeId, date, note) {
    await insertLedgerEntry({ employeeId, type: "daehyu", direction: "adjust", minutes: DAY_MINUTES, date, note: note || "대휴 발생(공휴일 근무)" });
  }

  async function addLedgerEntry() {
    if (!ledgerEmp || !ledgerMinutes) return;
    await insertLedgerEntry({
      employeeId: ledgerEmp,
      type: ledgerType,
      direction: ledgerDir,
      minutes: ledgerMinutes,
      date: ledgerDate,
      note: ledgerNote.trim(),
    });
    setLedgerNote("");
  }
  async function removeLedgerEntry(id) {
    if (!requireAuth()) return;
    const { error: err } = await supabase.from("ledger").delete().eq("id", id);
    if (err) setError(`기록 삭제 실패: ${err.message}`);
    else fetchLedger();
  }

  const employeeMap = useMemo(() => {
    const m = {};
    (employees || []).forEach((e) => (m[e.id] = e.name));
    return m;
  }, [employees]);

  const monthsInData = useMemo(() => {
    const map = {};
    (attendance || []).forEach((r) => {
      const m = r.date.slice(0, 7);
      map[m] = (map[m] || 0) + 1;
    });
    return Object.entries(map).sort((a, b) => b[0].localeCompare(a[0]));
  }, [attendance]);

  const summaryRows = useMemo(() => {
    if (!employees || !ledger || !attendance) return [];
    return employees.map((emp) => {
      const empRecords = attendance.filter(
        (r) => (r.employeeId === emp.id || r.employeeName === emp.name) && r.date <= asOf
      );
      const empLedger = ledger.filter((l) => l.employeeId === emp.id);

      let totalLateCount = 0;
      let totalLateMinutes = 0;
      let totalOtRaw = 0;
      let workedDays = 0;
      empRecords.forEach((r) => {
        if (r.checkin) workedDays++;
        const m = computeMetrics(r, emp);
        if (m.late > 0 && !r.excused) {
          totalLateCount++;
          totalLateMinutes += m.late;
        }
        totalOtRaw += m.otRaw;
      });
      // 매일 개별적으로 1.5배 후 버림 처리하면 소수점이 매일 손실되므로,
      // 원본 OT 분(raw)을 전체 기간 합산한 뒤 마지막에 한 번만 1.5배 적용
      const totalOtEarned = Math.floor(totalOtRaw * OT_MULTIPLIER);

      const otUsed = empLedger
        .filter((l) => l.type === "ot" && l.direction === "use")
        .reduce((s, l) => s + l.minutes, 0);
      const otAdjust = empLedger
        .filter((l) => l.type === "ot" && l.direction === "adjust")
        .reduce((s, l) => s + l.minutes, 0);
      const otRemaining = emp.openingOTMinutes + totalOtEarned - otUsed + otAdjust;

      const leaveCalc = cumulativeLeaveDays(emp.hireDate, asOf);
      const leaveEarnedMinutes = leaveCalc.days * DAY_MINUTES;
      const leaveUsed = empLedger
        .filter((l) => l.type === "leave" && l.direction === "use")
        .reduce((s, l) => s + l.minutes, 0);
      const leaveAdjust = empLedger
        .filter((l) => l.type === "leave" && l.direction === "adjust")
        .reduce((s, l) => s + l.minutes, 0);
      const leaveRemaining = emp.openingLeaveMinutes + leaveEarnedMinutes - leaveUsed + leaveAdjust;

      const daehyuGranted = empLedger
        .filter((l) => l.type === "daehyu" && (l.direction === "adjust" || l.direction === "grant"))
        .reduce((s, l) => s + l.minutes, 0);
      const daehyuUsed = empLedger
        .filter((l) => l.type === "daehyu" && l.direction === "use")
        .reduce((s, l) => s + l.minutes, 0);
      const daehyuRemaining = daehyuGranted - daehyuUsed;

      return {
        ...emp,
        workedDays,
        totalLateCount,
        totalLateMinutes,
        totalOtEarned,
        otUsed,
        otRemaining,
        leaveCalc,
        leaveEarnedMinutes,
        leaveUsed,
        leaveRemaining,
        daehyuGranted,
        daehyuUsed,
        daehyuRemaining,
      };
    });
  }, [employees, ledger, attendance, asOf]);

  const groupedByTeam = useMemo(() => {
    const groups = {};
    summaryRows.forEach((r) => {
      const t = r.team || "미지정";
      if (!groups[t]) groups[t] = [];
      groups[t].push(r);
    });
    Object.values(groups).forEach((rows) => {
      rows.sort((a, b) => {
        const ra = POSITION_RANK[a.position] ?? 99;
        const rb = POSITION_RANK[b.position] ?? 99;
        if (ra !== rb) return ra - rb;
        return a.name.localeCompare(b.name);
      });
    });
    return TEAM_ORDER.filter((t) => groups[t] && groups[t].length > 0).map((t) => ({
      team: t,
      rows: groups[t],
    }));
  }, [summaryRows]);

  function toggleTeam(team) {
    setCollapsed((c) => ({ ...c, [team]: !c[team] }));
  }

  if (!configOk) {
    return (
      <div style={{ padding: 40, fontFamily: FONT, color: COLORS.red }}>
        VITE_SUPABASE_URL / VITE_SUPABASE_ANON_KEY 환경변수가 설정되지 않았습니다.
        <br />
        Netlify 사이트 설정의 Environment variables에 두 값을 추가한 뒤 다시 배포(Deploy)해주세요.
      </div>
    );
  }

  if (!employees || !attendance || !ledger || !uploadLog || !dayStatusOverrides) {
    return (
      <div style={{ padding: 40, fontFamily: FONT }}>
        <div style={{ color: COLORS.sub, marginBottom: 12 }}>불러오는 중...</div>
        {error && (
          <div style={{ background: COLORS.redSoft, color: COLORS.red, padding: "10px 14px", borderRadius: 8, fontSize: 13.5, maxWidth: 560 }}>
            {error}
            <div style={{ marginTop: 8 }}>
              <button
                onClick={retryAll}
                style={{
                  marginTop: 4,
                  cursor: "pointer",
                  border: "none",
                  borderRadius: 6,
                  fontSize: 13,
                  padding: "7px 14px",
                  fontWeight: 600,
                  background: COLORS.teal,
                  color: "#fff",
                }}
              >
                다시 시도
              </button>
            </div>
          </div>
        )}
      </div>
    );
  }

  return (
    <div style={{ fontFamily: FONT, background: COLORS.bg, minHeight: "100vh", color: COLORS.text, padding: "28px 24px 60px" }}>
      <style>{`
        * { box-sizing: border-box; }
        body { margin: 0; }
        input, select { font-family: ${FONT}; }
        table { border-collapse: collapse; width: 100%; }
        th, td { text-align: left; padding: 9px 11px; font-size: 13px; white-space: nowrap; }
        tbody tr:not(:last-child) { border-bottom: 1px solid ${COLORS.border}; }
        .btn { cursor: pointer; border: none; border-radius: 6px; font-size: 13px; padding: 7px 14px; font-weight: 600; }
        .btn:active { transform: translateY(1px); }
        .btn:disabled { opacity: 0.5; cursor: not-allowed; }
        .input { border: 1px solid ${COLORS.border}; border-radius: 6px; padding: 7px 10px; font-size: 13.5px; background: #fff; color: ${COLORS.text}; }
        .input:focus { outline: 2px solid ${COLORS.teal}; outline-offset: 1px; }
        .tab { cursor: pointer; padding: 9px 4px; font-size: 14px; font-weight: 600; border-bottom: 2px solid transparent; color: ${COLORS.sub}; }
        .tab.active { color: ${COLORS.tealDark}; border-bottom-color: ${COLORS.teal}; }
        .dropzone { border: 2px dashed ${COLORS.border}; border-radius: 10px; padding: 36px 20px; text-align: center; color: ${COLORS.sub}; cursor: pointer; background: #fff; }
        .dropzone:hover { border-color: ${COLORS.teal}; }
        .card { background: ${COLORS.card}; border: 1px solid ${COLORS.border}; border-radius: 10px; padding: 16px; }
        .team-header { display: flex; align-items: center; gap: 8px; cursor: pointer; background: ${COLORS.tealDark}; color: #fff; padding: 10px 14px; border-radius: 8px 8px 0 0; font-weight: 700; font-size: 13.5px; user-select: none; }
        .chevron { transition: transform 0.15s ease; display: inline-block; }
        .chevron.closed { transform: rotate(-90deg); }
      `}</style>

      <header style={{ marginBottom: 22 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 10 }}>
          <div>
            <div style={{ fontSize: 12, letterSpacing: 0.3, color: COLORS.teal, fontWeight: 700, marginBottom: 4 }}>
              DOD Dermatology Cheongdam
            </div>
            <h1 style={{ fontSize: 24, fontWeight: 800, margin: 0, color: COLORS.tealDark }}>근태 · 연차 · OT 통합관리</h1>
            <p style={{ fontSize: 13, color: COLORS.sub, marginTop: 6 }}>
              Supabase 데이터베이스에 저장됩니다 · 캡스 근태 엑셀을 업로드하면 지각·연장근무가 자동 반영되고 매달 계속 누적됩니다.
            </p>
          </div>
          <div>
            {session ? (
              <div style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12.5, color: COLORS.sub }}>
                <span>관리자: {session.user.email}</span>
                <button className="btn" style={{ background: COLORS.tealSoft, color: COLORS.tealDark }} onClick={() => supabase.auth.signOut()}>
                  로그아웃
                </button>
              </div>
            ) : (
              <button className="btn" style={{ background: COLORS.teal, color: "#fff" }} onClick={() => setShowLogin(true)}>
                관리자 로그인
              </button>
            )}
          </div>
        </div>
      </header>

      {showLogin && !session && <LoginModal onClose={() => setShowLogin(false)} />}

      <div style={{ display: "flex", gap: 20, alignItems: "center", marginBottom: 18, flexWrap: "wrap" }}>
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
          {[
            ["dashboard", "대시보드"],
            ["calendar", "캘린더"],
            ["upload", "데이터 업로드"],
            ["ledger", "연차·OT 사용내역"],
            ["employees", "직원 관리"],
          ].map(([key, label]) => (
            <div key={key} className={`tab ${tab === key ? "active" : ""}`} onClick={() => setTab(key)}>
              {label}
            </div>
          ))}
        </div>
        <div style={{ marginLeft: "auto", display: "flex", alignItems: "center", gap: 8 }}>
          <label style={{ fontSize: 12.5, color: COLORS.sub }}>기준일</label>
          <input type="date" className="input" value={asOf} onChange={(e) => setAsOf(e.target.value)} />
          <button className="btn" style={{ background: COLORS.tealSoft, color: COLORS.tealDark }} onClick={retryAll}>
            새로고침
          </button>
        </div>
      </div>

      {error && (
        <div style={{ background: COLORS.redSoft, color: COLORS.red, padding: "8px 12px", borderRadius: 6, fontSize: 13, marginBottom: 14 }}>
          {error}
        </div>
      )}

      {tab === "dashboard" && (
        <DashboardTab groupedByTeam={groupedByTeam} collapsed={collapsed} toggleTeam={toggleTeam} onSelect={setSelectedId} />
      )}
      {tab === "calendar" && (
        <CalendarTab
          employees={employees}
          attendance={attendance}
          dayStatusOverrides={dayStatusOverrides}
          ledger={ledger}
          setLateExcused={setLateExcused}
          setLateDeduction={setLateDeduction}
          removeLateDeduction={removeLateDeduction}
          setDayAdjustment={setDayAdjustment}
          removeDayAdjustment={removeDayAdjustment}
        />
      )}
      {tab === "upload" && (
        <UploadTab
          fileRef={fileRef}
          onFileInput={onFileInput}
          onDrop={onDrop}
          uploadMsg={uploadMsg}
          uploadLog={uploadLog}
          uploading={uploading}
          monthsInData={monthsInData}
          deleteMonthData={deleteMonthData}
        />
      )}
      {tab === "ledger" && (
        <LedgerTab
          employees={employees}
          ledger={ledger}
          employeeMap={employeeMap}
          ledgerEmp={ledgerEmp}
          setLedgerEmp={setLedgerEmp}
          ledgerType={ledgerType}
          setLedgerType={setLedgerType}
          ledgerDir={ledgerDir}
          setLedgerDir={setLedgerDir}
          ledgerMinutes={ledgerMinutes}
          setLedgerMinutes={setLedgerMinutes}
          ledgerDate={ledgerDate}
          setLedgerDate={setLedgerDate}
          ledgerNote={ledgerNote}
          setLedgerNote={setLedgerNote}
          addLedgerEntry={addLedgerEntry}
          removeLedgerEntry={removeLedgerEntry}
        />
      )}
      {tab === "employees" && (
        <EmployeesTab
          employees={employees}
          updateEmployee={updateEmployee}
          removeEmployee={removeEmployee}
          newName={newName}
          setNewName={setNewName}
          newTeam={newTeam}
          setNewTeam={setNewTeam}
          newPosition={newPosition}
          setNewPosition={setNewPosition}
          newHire={newHire}
          setNewHire={setNewHire}
          addEmployee={addEmployee}
        />
      )}

      {selectedId && (
        <EmployeeDetailModal
          row={summaryRows.find((r) => r.id === selectedId)}
          ledger={ledger}
          removeLedgerEntry={removeLedgerEntry}
          insertLedgerEntry={insertLedgerEntry}
          grantDaehyu={grantDaehyu}
          updateEmployee={updateEmployee}
          removeEmployee={removeEmployee}
          onClose={() => setSelectedId(null)}
        />
      )}
    </div>
  );
}

function LoginModal({ onClose }) {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);

  async function handleLogin(e) {
    e.preventDefault();
    setError("");
    setLoading(true);
    const { error: err } = await supabase.auth.signInWithPassword({ email, password });
    setLoading(false);
    if (err) setError("로그인 실패: 이메일 또는 비밀번호를 확인해주세요.");
  }

  return (
    <div
      onClick={onClose}
      style={{
        position: "fixed", inset: 0, background: "rgba(10,20,18,0.45)",
        display: "flex", alignItems: "center", justifyContent: "center", zIndex: 1100, padding: 20,
      }}
    >
      <form
        onSubmit={handleLogin}
        onClick={(e) => e.stopPropagation()}
        style={{ background: "#fff", borderRadius: 14, width: "100%", maxWidth: 360, padding: 26, boxShadow: "0 20px 60px rgba(0,0,0,0.25)" }}
      >
        <div style={{ fontSize: 17, fontWeight: 800, color: COLORS.tealDark, marginBottom: 4 }}>관리자 로그인</div>
        <div style={{ fontSize: 12.5, color: COLORS.sub, marginBottom: 16 }}>
          수정 권한이 있는 계정으로 로그인하세요. 조회는 로그인 없이도 가능합니다.
        </div>
        <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          <input
            className="input"
            type="email"
            placeholder="이메일"
            value={email}
            onChange={(e) => setEmail(e.target.value)}
            required
            autoFocus
          />
          <input
            className="input"
            type="password"
            placeholder="비밀번호"
            value={password}
            onChange={(e) => setPassword(e.target.value)}
            required
          />
        </div>
        {error && <div style={{ color: COLORS.red, fontSize: 12.5, marginTop: 10 }}>{error}</div>}
        <div style={{ display: "flex", gap: 8, marginTop: 18 }}>
          <button type="submit" className="btn" style={{ background: COLORS.teal, color: "#fff", flex: 1 }} disabled={loading}>
            {loading ? "확인 중..." : "로그인"}
          </button>
          <button type="button" className="btn" style={{ background: "transparent", color: COLORS.sub }} onClick={onClose}>
            취소
          </button>
        </div>
      </form>
    </div>
  );
}

/* ───────────────────────── tabs ───────────────────────── */

function DashboardTab({ groupedByTeam, collapsed, toggleTeam, onSelect }) {
  return (
    <>
      {groupedByTeam.map(({ team, rows }) => {
        const isClosed = !!collapsed[team];
        const teamColor = TEAM_COLORS[team] || COLORS.tealDark;
        return (
          <div key={team} style={{ marginBottom: 16 }}>
            <div className="team-header" style={{ background: teamColor }} onClick={() => toggleTeam(team)}>
              <span className={`chevron ${isClosed ? "closed" : ""}`}>▾</span>
              {team} <span style={{ opacity: 0.75, fontWeight: 500 }}>· {rows.length}명</span>
            </div>
            {!isClosed && (
              <div className="card" style={{ borderRadius: "0 0 8px 8px", overflow: "auto" }}>
                <table>
                  <thead style={{ background: COLORS.tealSoft }}>
                    <tr>
                      <th>이름</th>
                      <th>직급</th>
                      <th style={{ paddingRight: 4 }}>누적 지각횟수</th>
                      <th style={{ paddingLeft: 4 }}>누적 지각시간</th>
                      <th>OT 잔여</th>
                      <th>연차 잔여</th>
                      <th>대휴 잔여</th>
                      <th></th>
                    </tr>
                  </thead>
                  <tbody>
                    {rows.map((r) => {
                      const otLow = r.otRemaining < 0;
                      const leaveLow = r.leaveRemaining <= DAY_MINUTES * 2 && r.leaveRemaining >= 0;
                      const leaveNeg = r.leaveRemaining < 0;
                      return (
                        <tr key={r.id}>
                          <td>
                            <button
                              onClick={() => onSelect(r.id)}
                              style={{
                                background: "none",
                                border: "none",
                                cursor: "pointer",
                                padding: 0,
                                fontSize: 13,
                                fontWeight: 700,
                                color: COLORS.tealDark,
                                textDecoration: "underline",
                                textDecorationColor: COLORS.border,
                                textUnderlineOffset: 3,
                              }}
                            >
                              {r.name}
                            </button>
                          </td>
                          <td style={{ color: COLORS.sub }}>{r.position || "-"}</td>
                          <td style={{ paddingRight: 4 }}>{r.totalLateCount}</td>
                          <td style={{ paddingLeft: 4 }}>{minutesToHM(r.totalLateMinutes)}</td>
                          <td>
                            <Badge text={minutesToHM(r.otRemaining)} tone={otLow ? "red" : "teal"} />
                          </td>
                          <td>
                            {r.hireDate ? (
                              <Badge
                                text={minutesToDaysLabel(r.leaveRemaining)}
                                tone={leaveNeg ? "red" : leaveLow ? "amber" : "teal"}
                              />
                            ) : (
                              <span style={{ color: COLORS.sub, fontSize: 12 }}>입사일 미입력</span>
                            )}
                          </td>
                          <td>
                            <Badge text={minutesToDaysLabel(r.daehyuRemaining)} tone="teal" />
                          </td>
                          <td>
                            <button
                              className="btn"
                              style={{ background: COLORS.tealSoft, color: COLORS.tealDark, padding: "4px 10px" }}
                              onClick={() => onSelect(r.id)}
                            >
                              상세 →
                            </button>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        );
      })}
    </>
  );
}

function WeeklyScheduleEditor({ employee, updateEmployee }) {
  function setDaySchedule(dow, field, value) {
    const current = { ...(employee.customSchedule || {}) };
    const dayEntry = { ...(current[dow] || {}) };
    if (value) dayEntry[field] = value;
    else delete dayEntry[field];
    if (Object.keys(dayEntry).length === 0) delete current[dow];
    else current[dow] = dayEntry;
    updateEmployee(employee.id, { customSchedule: current });
  }
  function clearDay(dow) {
    const current = { ...(employee.customSchedule || {}) };
    delete current[dow];
    updateEmployee(employee.id, { customSchedule: current });
  }

  return (
    <>
      <div style={{ fontSize: 12.5, color: COLORS.sub, marginBottom: 8 }}>
        요일별로 이 직원만 다르게 적용할 출근/퇴근 시각을 입력하세요. 비워두면 회사 기본 시간표를 그대로 사용합니다.
      </div>
      <table>
        <thead>
          <tr>
            <th style={{ padding: "4px 8px" }}>요일</th>
            <th style={{ padding: "4px 8px" }}>기본값</th>
            <th style={{ padding: "4px 8px" }}>개인 출근</th>
            <th style={{ padding: "4px 8px" }}>개인 퇴근</th>
            <th style={{ padding: "4px 8px" }}></th>
          </tr>
        </thead>
        <tbody>
          {WEEKDAY_ORDER.map((dow) => {
            const def = SCHEDULE[dow];
            const custom = (employee.customSchedule && employee.customSchedule[dow]) || {};
            return (
              <tr key={dow}>
                <td style={{ padding: "4px 8px", fontWeight: 600 }}>{WEEKDAY_LABELS[dow]}</td>
                <td style={{ padding: "4px 8px", color: COLORS.sub }}>{def.start}~{def.end}</td>
                <td style={{ padding: "4px 8px" }}>
                  <input
                    type="time"
                    className="input"
                    style={{ padding: "4px 6px", fontSize: 12.5 }}
                    value={custom.start || ""}
                    onChange={(ev) => setDaySchedule(dow, "start", ev.target.value)}
                  />
                </td>
                <td style={{ padding: "4px 8px" }}>
                  <input
                    type="time"
                    className="input"
                    style={{ padding: "4px 6px", fontSize: 12.5 }}
                    value={custom.end || ""}
                    onChange={(ev) => setDaySchedule(dow, "end", ev.target.value)}
                  />
                </td>
                <td style={{ padding: "4px 8px" }}>
                  {(custom.start || custom.end) && (
                    <button className="btn" style={{ background: "transparent", color: COLORS.sub, padding: "2px 8px", fontSize: 12 }} onClick={() => clearDay(dow)}>
                      기본값으로
                    </button>
                  )}
                </td>
              </tr>
            );
          })}
        </tbody>
      </table>
    </>
  );
}

function EmployeeDetailModal({
  row, ledger, removeLedgerEntry, insertLedgerEntry, grantDaehyu,
  updateEmployee, removeEmployee, onClose,
}) {
  const [expandedStat, setExpandedStat] = useState(null); // null | 'ot' | 'leave' | 'daehyu'
  const [addValue, setAddValue] = useState("");
  const [addDate, setAddDate] = useState(toISO(new Date()));
  const [addNote, setAddNote] = useState("");
  if (!row) return null;
  const otLow = row.otRemaining < 0;
  const leaveLow = row.leaveRemaining <= DAY_MINUTES * 2 && row.leaveRemaining >= 0;
  const leaveNeg = row.leaveRemaining < 0;
  const teamColor = TEAM_COLORS[row.team] || COLORS.tealDark;
  const statHistory = (ledger || [])
    .filter((l) => l.employeeId === row.id && l.type === expandedStat)
    .sort((a, b) => (a.date < b.date ? 1 : -1));

  function toggleStat(key) {
    setExpandedStat(expandedStat === key ? null : key);
    setAddValue("");
    setAddNote("");
  }
  function submitUse() {
    if (!addValue) return;
    if (expandedStat === "leave") {
      insertLedgerEntry({ employeeId: row.id, type: "leave", direction: "use", minutes: Math.round(parseFloat(addValue) * DAY_MINUTES), date: addDate, note: addNote });
    } else if (expandedStat === "ot") {
      insertLedgerEntry({ employeeId: row.id, type: "ot", direction: "use", minutes: addValue, date: addDate, note: addNote });
    }
    setAddValue("");
  }
  function submitDaehyuGrant() {
    grantDaehyu(row.id, addDate, addNote || undefined);
    setAddNote("");
  }
  function submitDaehyuUse() {
    if (!addValue) return;
    insertLedgerEntry({ employeeId: row.id, type: "daehyu", direction: "use", minutes: Math.round(parseFloat(addValue) * DAY_MINUTES), date: addDate, note: addNote || "대휴 사용" });
    setAddValue("");
  }

  const stat = (label, value, onClick) => (
    <div
      onClick={onClick}
      style={{
        background: COLORS.bg, borderRadius: 8, padding: "10px 12px",
        cursor: onClick ? "pointer" : "default",
        border: onClick ? `1px solid ${COLORS.border}` : "1px solid transparent",
      }}
    >
      <div style={{ fontSize: 11.5, color: COLORS.sub, marginBottom: 3, display: "flex", justifyContent: "space-between" }}>
        <span>{label}</span>
        {onClick && <span style={{ color: COLORS.teal }}>내역 ▸</span>}
      </div>
      <div style={{ fontSize: 14.5, fontWeight: 700 }}>{value}</div>
    </div>
  );

  return (
    <div
      onClick={onClose}
      style={{
        position: "fixed", inset: 0, background: "rgba(10,20,18,0.45)",
        display: "flex", alignItems: "center", justifyContent: "center", zIndex: 1000, padding: 20,
      }}
    >
      <div
        onClick={(e) => e.stopPropagation()}
        style={{
          background: "#fff", borderRadius: 14, width: "100%", maxWidth: 640,
          maxHeight: "88vh", overflow: "auto", boxShadow: "0 20px 60px rgba(0,0,0,0.25)",
        }}
      >
        <div style={{ background: teamColor, color: "#fff", padding: "18px 22px", borderRadius: "14px 14px 0 0", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
          <div>
            <div style={{ fontSize: 12, opacity: 0.85, marginBottom: 2 }}>{row.team} · {row.position || "직급 미지정"}</div>
            <div style={{ fontSize: 20, fontWeight: 800 }}>{row.name}</div>
          </div>
          <button onClick={onClose} style={{ background: "rgba(255,255,255,0.15)", border: "none", color: "#fff", borderRadius: 6, padding: "6px 10px", cursor: "pointer", fontSize: 13 }}>
            닫기 ✕
          </button>
        </div>

        <div style={{ padding: 22 }}>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 10, marginBottom: 20 }}>
            {stat("근무일수", `${row.workedDays}일`)}
            {stat("누적 지각횟수", `${row.totalLateCount}회`)}
            {stat("누적 지각시간", minutesToHM(row.totalLateMinutes))}
            {stat("누적 OT 적립", minutesToHM(row.totalOtEarned))}
            {stat("OT 사용", minutesToHM(row.otUsed), () => setExpandedStat(expandedStat === "ot" ? null : "ot"))}
            <div style={{ background: otLow ? COLORS.redSoft : COLORS.tealSoft, borderRadius: 8, padding: "10px 12px" }}>
              <div style={{ fontSize: 11.5, color: otLow ? COLORS.red : COLORS.tealDark, marginBottom: 3 }}>OT 잔여</div>
              <div style={{ fontSize: 14.5, fontWeight: 700, color: otLow ? COLORS.red : COLORS.tealDark }}>{minutesToHM(row.otRemaining)}</div>
            </div>
            {stat("연차 발생", row.hireDate ? `${row.leaveCalc.days}일` : "입사일 미입력")}
            {stat("연차 사용", minutesToDaysLabel(row.leaveUsed), () => toggleStat("leave"))}
            <div style={{ background: leaveNeg ? COLORS.redSoft : leaveLow ? COLORS.amberSoft : COLORS.tealSoft, borderRadius: 8, padding: "10px 12px" }}>
              <div style={{ fontSize: 11.5, color: leaveNeg ? COLORS.red : leaveLow ? COLORS.amber : COLORS.tealDark, marginBottom: 3 }}>연차 잔여</div>
              <div style={{ fontSize: 14.5, fontWeight: 700, color: leaveNeg ? COLORS.red : leaveLow ? COLORS.amber : COLORS.tealDark }}>
                {row.hireDate ? minutesToDaysLabel(row.leaveRemaining) : "-"}
              </div>
            </div>
            {stat("대휴 발생", minutesToDaysLabel(row.daehyuGranted))}
            {stat("대휴 사용", minutesToDaysLabel(row.daehyuUsed), () => toggleStat("daehyu"))}
            <div style={{ background: COLORS.tealSoft, borderRadius: 8, padding: "10px 12px" }}>
              <div style={{ fontSize: 11.5, color: COLORS.tealDark, marginBottom: 3 }}>대휴 잔여</div>
              <div style={{ fontSize: 14.5, fontWeight: 700, color: COLORS.tealDark }}>{minutesToDaysLabel(row.daehyuRemaining)}</div>
            </div>
          </div>

          {expandedStat && (
            <div style={{ marginBottom: 20, border: `1px solid ${COLORS.border}`, borderRadius: 8, overflow: "hidden" }}>
              <div style={{ background: COLORS.tealSoft, color: COLORS.tealDark, padding: "8px 12px", fontSize: 12.5, fontWeight: 700 }}>
                {expandedStat === "ot" ? "OT 사용/조정 내역" : expandedStat === "leave" ? "연차 사용/조정 내역" : "대휴 발생/사용 내역"}
              </div>

              <div style={{ padding: 10, borderBottom: `1px solid ${COLORS.border}`, display: "flex", gap: 6, flexWrap: "wrap", alignItems: "center", background: "#fafcfb" }}>
                {expandedStat === "leave" && (
                  <>
                    <input type="number" step="0.5" className="input" placeholder="일수" style={{ width: 80 }} value={addValue} onChange={(e) => setAddValue(e.target.value)} />
                    <span style={{ fontSize: 12, color: COLORS.sub }}>일 사용</span>
                    <input type="date" className="input" value={addDate} onChange={(e) => setAddDate(e.target.value)} />
                    <input className="input" placeholder="메모(선택)" style={{ width: 130 }} value={addNote} onChange={(e) => setAddNote(e.target.value)} />
                    <button className="btn" style={{ background: COLORS.teal, color: "#fff" }} onClick={submitUse}>추가</button>
                  </>
                )}
                {expandedStat === "ot" && (
                  <>
                    <input type="number" className="input" placeholder="분" style={{ width: 80 }} value={addValue} onChange={(e) => setAddValue(e.target.value)} />
                    <span style={{ fontSize: 12, color: COLORS.sub }}>분 사용</span>
                    <input type="date" className="input" value={addDate} onChange={(e) => setAddDate(e.target.value)} />
                    <input className="input" placeholder="메모(선택)" style={{ width: 130 }} value={addNote} onChange={(e) => setAddNote(e.target.value)} />
                    <button className="btn" style={{ background: COLORS.teal, color: "#fff" }} onClick={submitUse}>추가</button>
                  </>
                )}
                {expandedStat === "daehyu" && (
                  <>
                    <input type="date" className="input" value={addDate} onChange={(e) => setAddDate(e.target.value)} />
                    <input className="input" placeholder="메모(선택, 예: 8/15 근무)" style={{ width: 150 }} value={addNote} onChange={(e) => setAddNote(e.target.value)} />
                    <button className="btn" style={{ background: COLORS.teal, color: "#fff" }} onClick={submitDaehyuGrant}>대휴 발생(+1일) 추가</button>
                    <span style={{ borderLeft: `1px solid ${COLORS.border}`, height: 20, margin: "0 4px" }} />
                    <input type="number" step="0.5" className="input" placeholder="일수" style={{ width: 70 }} value={addValue} onChange={(e) => setAddValue(e.target.value)} />
                    <button className="btn" style={{ background: COLORS.tealSoft, color: COLORS.tealDark }} onClick={submitDaehyuUse}>대휴 사용 추가</button>
                  </>
                )}
              </div>

              <table>
                <thead>
                  <tr>
                    <th style={{ padding: "6px 10px" }}>날짜</th>
                    <th style={{ padding: "6px 10px" }}>구분</th>
                    <th style={{ padding: "6px 10px" }}>분</th>
                    <th style={{ padding: "6px 10px" }}>메모</th>
                    <th style={{ padding: "6px 10px" }}></th>
                  </tr>
                </thead>
                <tbody>
                  {statHistory.length === 0 && (
                    <tr>
                      <td colSpan={5} style={{ padding: 16, textAlign: "center", color: COLORS.sub }}>
                        내역이 없습니다.
                      </td>
                    </tr>
                  )}
                  {statHistory.map((l) => (
                    <tr key={l.id}>
                      <td style={{ padding: "6px 10px" }}>{fmtDate(l.date)}</td>
                      <td style={{ padding: "6px 10px" }}>{l.direction === "use" ? "사용" : l.direction === "grant" ? "발생" : "조정(+)"}</td>
                      <td style={{ padding: "6px 10px" }}>{l.minutes}분</td>
                      <td style={{ padding: "6px 10px", color: COLORS.sub }}>{l.note || "-"}</td>
                      <td style={{ padding: "6px 10px" }}>
                        <button className="btn" style={{ background: "transparent", color: COLORS.red, padding: "3px 8px", fontSize: 12 }} onClick={() => removeLedgerEntry(l.id)}>
                          삭제
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
          <div style={{ fontSize: 13, fontWeight: 700, color: COLORS.tealDark, marginBottom: 10 }}>기본 정보 수정</div>
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 18 }}>
            <select className="input" value={row.team || "미지정"} onChange={(ev) => updateEmployee(row.id, { team: ev.target.value })}>
              {TEAM_ORDER.map((t) => (<option key={t} value={t}>{t}</option>))}
            </select>
            <select className="input" value={row.position || ""} onChange={(ev) => updateEmployee(row.id, { position: ev.target.value })}>
              <option value="">직급 -</option>
              {POSITION_LIST.map((p) => (<option key={p} value={p}>{p}</option>))}
            </select>
            <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12.5, color: COLORS.sub }}>
              입사일
              <input type="date" className="input" value={row.hireDate} onChange={(ev) => updateEmployee(row.id, { hireDate: ev.target.value })} />
            </label>
            <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12.5, color: COLORS.sub }}>
              연차 초기값(분)
              <input type="number" className="input" style={{ width: 90 }} value={row.openingLeaveMinutes} onChange={(ev) => updateEmployee(row.id, { openingLeaveMinutes: parseInt(ev.target.value, 10) || 0 })} />
            </label>
            <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12.5, color: COLORS.sub }}>
              OT 초기값(분)
              <input type="number" className="input" style={{ width: 90 }} value={row.openingOTMinutes} onChange={(ev) => updateEmployee(row.id, { openingOTMinutes: parseInt(ev.target.value, 10) || 0 })} />
            </label>
          </div>

          <div style={{ fontSize: 12, color: COLORS.sub, marginBottom: 8 }}>
            지각/늦출/일퇴 등 일자별 기록은 캘린더 탭에서 확인하고 수정할 수 있어요.
          </div>

          <div style={{ marginTop: 6, borderTop: `1px solid ${COLORS.border}`, paddingTop: 14 }}>
            <button
              className="btn"
              style={{ background: COLORS.redSoft, color: COLORS.red }}
              onClick={() => {
                if (window.confirm(`${row.name} 님을 직원 목록에서 삭제할까요?`)) {
                  removeEmployee(row.id);
                  onClose();
                }
              }}
            >
              직원 삭제
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}

function Badge({ text, tone }) {
  const map = {
    teal: { bg: COLORS.tealSoft, fg: COLORS.tealDark },
    amber: { bg: COLORS.amberSoft, fg: COLORS.amber },
    red: { bg: COLORS.redSoft, fg: COLORS.red },
  };
  const c = map[tone] || map.teal;
  return (
    <span style={{ fontWeight: 700, padding: "2px 8px", borderRadius: 5, background: c.bg, color: c.fg }}>
      {text}
    </span>
  );
}

function CalendarTab({ employees, attendance, dayStatusOverrides, ledger, setLateExcused, setLateDeduction, removeLateDeduction, setDayAdjustment, removeDayAdjustment }) {
  const [cursor, setCursor] = useState(() => {
    const d = new Date();
    return new Date(d.getFullYear(), d.getMonth(), 1);
  });
  const [selectedDate, setSelectedDate] = useState(null);

  const attendanceIndex = useMemo(() => {
    const idx = {};
    attendance.forEach((r) => {
      const emp = employees.find((e) => e.id === r.employeeId || e.name === r.employeeName);
      if (!emp) return;
      idx[`${emp.id}_${r.date}`] = r;
    });
    return idx;
  }, [attendance, employees]);

  const overrideIndex = useMemo(() => {
    const idx = {};
    dayStatusOverrides.forEach((o) => {
      idx[`${o.employeeId}_${o.date}`] = o;
    });
    return idx;
  }, [dayStatusOverrides]);

  const cells = useMemo(() => {
    const year = cursor.getFullYear();
    const month = cursor.getMonth();
    const daysInMonth = new Date(year, month + 1, 0).getDate();
    const startDow = new Date(year, month, 1).getDay();
    const out = [];
    for (let i = 0; i < startDow; i++) out.push(null);
    for (let d = 1; d <= daysInMonth; d++) {
      const dateISO = `${year}-${pad(month + 1)}-${pad(d)}`;
      const people = [];
      employees.forEach((emp) => {
        const key = `${emp.id}_${dateISO}`;
        const st = computeDayStatus(emp, dateISO, attendanceIndex[key], overrideIndex[key]);
        if (st.source === "manual" || (st.status !== "출근" && st.status !== "OFF")) {
          people.push({ employeeId: emp.id, name: emp.name, ...st });
        }
      });
      out.push({ date: dateISO, day: d, people });
    }
    return out;
  }, [cursor, employees, attendanceIndex, overrideIndex]);

  function goMonth(delta) {
    setCursor(new Date(cursor.getFullYear(), cursor.getMonth() + delta, 1));
  }

  const monthlyLateCounts = useMemo(() => {
    const counts = {};
    cells.forEach((cell) => {
      if (!cell) return;
      cell.people.forEach((p) => {
        if (p.status === "지각") counts[p.name] = (counts[p.name] || 0) + 1;
      });
    });
    return Object.entries(counts).sort((a, b) => b[1] - a[1]);
  }, [cells]);

  return (
    <>
      <div className="card" style={{ marginBottom: 14 }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
          <button className="btn" style={{ background: COLORS.tealSoft, color: COLORS.tealDark }} onClick={() => goMonth(-1)}>‹ 이전달</button>
          <div style={{ fontSize: 16, fontWeight: 800, color: COLORS.tealDark }}>
            {cursor.getFullYear()}년 {cursor.getMonth() + 1}월
          </div>
          <button className="btn" style={{ background: COLORS.tealSoft, color: COLORS.tealDark }} onClick={() => goMonth(1)}>다음달 ›</button>
        </div>
      </div>

      <div className="card" style={{ marginBottom: 14 }}>
        <div style={{ fontSize: 13, fontWeight: 700, color: COLORS.tealDark, marginBottom: 8 }}>
          {cursor.getFullYear()}년 {cursor.getMonth() + 1}월 지각자
        </div>
        {monthlyLateCounts.length === 0 ? (
          <div style={{ fontSize: 12.5, color: COLORS.sub }}>이번 달 지각 기록이 없습니다.</div>
        ) : (
          <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
            {monthlyLateCounts.map(([name, count]) => (
              <span
                key={name}
                style={{
                  fontSize: 12.5, fontWeight: 600, padding: "4px 10px", borderRadius: 6,
                  background: STATUS_COLORS.지각.bg, color: STATUS_COLORS.지각.fg,
                }}
              >
                {name} 누적 {count}회
              </span>
            ))}
          </div>
        )}
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(7, 1fr)", gap: 6 }}>
        {["일", "월", "화", "수", "목", "금", "토"].map((d) => (
          <div key={d} style={{ textAlign: "center", fontSize: 12, fontWeight: 700, color: COLORS.sub, padding: "4px 0" }}>
            {d}
          </div>
        ))}
        {cells.map((cell, i) =>
          cell === null ? (
            <div key={`b${i}`} />
          ) : (
            <div
              key={cell.date}
              onClick={() => setSelectedDate(cell.date)}
              style={{
                background: COLORS.card, border: `1px solid ${COLORS.border}`, borderRadius: 8,
                minHeight: 150, padding: 7, cursor: "pointer", display: "flex", flexDirection: "column", gap: 4,
              }}
            >
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                <span style={{ fontSize: 12.5, fontWeight: 700, color: COLORS.text }}>{cell.day}</span>
                {cell.people.length > 0 && (
                  <span style={{ fontSize: 10, color: COLORS.sub }}>{cell.people.length}</span>
                )}
              </div>
              <div style={{ display: "flex", flexWrap: "wrap", gap: 3 }}>
                {cell.people.slice(0, 8).map((p) => {
                  const c = STATUS_COLORS[p.status] || STATUS_COLORS.OFF;
                  const showMin = p.minutes && ["지각", "늦출", "일퇴", "반차"].includes(p.status);
                  return (
                    <div
                      key={p.employeeId}
                      style={{
                        fontSize: 10.5, padding: "1px 5px", borderRadius: 4, background: c.bg, color: c.fg,
                        whiteSpace: "nowrap", fontWeight: 600,
                      }}
                    >
                      {p.name} {showMin ? `${p.minutes}분 ` : ""}{p.status}
                    </div>
                  );
                })}
                {cell.people.length > 8 && (
                  <div style={{ fontSize: 10, color: COLORS.sub }}>+{cell.people.length - 8}명</div>
                )}
              </div>
            </div>
          )
        )}
      </div>

      <div style={{ display: "flex", gap: 12, marginTop: 14, flexWrap: "wrap" }}>
        {Object.keys(STATUS_COLORS).map((s) => (
          <div key={s} style={{ display: "flex", alignItems: "center", gap: 5, fontSize: 12, color: COLORS.sub }}>
            <span style={{ width: 10, height: 10, borderRadius: 3, background: STATUS_COLORS[s].bg, border: `1px solid ${STATUS_COLORS[s].fg}` }} />
            {s}
          </div>
        ))}
      </div>

      {selectedDate && (
        <DayDetailModal
          date={selectedDate}
          employees={employees}
          attendanceIndex={attendanceIndex}
          overrideIndex={overrideIndex}
          ledger={ledger}
          setLateExcused={setLateExcused}
          setLateDeduction={setLateDeduction}
          removeLateDeduction={removeLateDeduction}
          setDayAdjustment={setDayAdjustment}
          removeDayAdjustment={removeDayAdjustment}
          onClose={() => setSelectedDate(null)}
        />
      )}
    </>
  );
}

/* 일퇴/OFF/출근/원본없는 지각·늦출의 분 단위 설정 인라인 폼
   OFF = 그날 기본 근무시간(-90분)만큼 연차 차감 / 출근(쉬는날 나옴) = 대휴 1일(480분) 고정 발생 */
function DayAdjustEditor({ status, initialMinutes, initialType, autoMinutes, onSave, onCancel }) {
  const isOff = status === "OFF";
  const isHalfOff = status === "반차";
  const isWork = status === "출근";
  const isAuto = isOff || isHalfOff || isWork;
  const fixedMinutes = isWork ? DAY_MINUTES : isHalfOff ? Math.round(autoMinutes / 2) : autoMinutes;
  const [minutes, setMinutes] = useState(String(isAuto ? fixedMinutes : initialMinutes || ""));
  const [dedType, setDedType] = useState(initialType || "ot");

  return (
    <div style={{ display: "flex", gap: 6, alignItems: "center", flexWrap: "wrap", background: COLORS.bg, padding: 8, borderRadius: 6 }}>
      {isOff && (
        <span style={{ fontSize: 12.5, color: COLORS.sub }}>
          요일 기본 근무시간 기준 자동 계산: <strong style={{ color: COLORS.text }}>{fixedMinutes}분</strong> 연차 차감
        </span>
      )}
      {isHalfOff && (
        <span style={{ fontSize: 12.5, color: COLORS.sub }}>
          반차 규정 기준 자동 계산(하루치 절반): <strong style={{ color: COLORS.text }}>{fixedMinutes}분</strong> 연차 차감
        </span>
      )}
      {isWork && (
        <span style={{ fontSize: 12.5, color: COLORS.sub }}>
          대휴(대체휴무) <strong style={{ color: COLORS.text }}>1일(480분)</strong> 발생
        </span>
      )}
      {!isAuto && (
        <>
          <input
            type="number"
            className="input"
            style={{ padding: "4px 6px", fontSize: 12.5, width: 70 }}
            placeholder="분"
            value={minutes}
            onChange={(e) => setMinutes(e.target.value)}
            autoFocus
          />
          <span style={{ fontSize: 12, color: COLORS.sub }}>분</span>
          <select className="input" style={{ padding: "4px 6px", fontSize: 12.5 }} value={dedType} onChange={(e) => setDedType(e.target.value)}>
            <option value="ot">OT에서 차감</option>
            <option value="leave">연차에서 차감</option>
          </select>
        </>
      )}
      <button
        className="btn"
        style={{ background: COLORS.teal, color: "#fff", padding: "4px 10px" }}
        onClick={() => {
          if (isWork) onSave(DAY_MINUTES, "daehyu");
          else if (isOff || isHalfOff) onSave(fixedMinutes, "leave");
          else onSave(minutes, dedType);
        }}
      >
        저장
      </button>
      <button className="btn" style={{ background: "transparent", color: COLORS.sub, padding: "4px 8px" }} onClick={onCancel}>
        취소
      </button>
    </div>
  );
}

function DayDetailModal({
  date, employees, attendanceIndex, overrideIndex, ledger,
  setLateExcused, setLateDeduction, removeLateDeduction,
  setDayAdjustment, removeDayAdjustment, onClose,
}) {
  const [showAdd, setShowAdd] = useState(false);
  const [addStep, setAddStep] = useState("pick"); // 'pick' | 'adjust'
  const [addEmp, setAddEmp] = useState("");
  const [addStatus, setAddStatus] = useState("일퇴");

  const [editingKey, setEditingKey] = useState(null); // employeeId_date currently showing the adjust editor
  const [editingStatus, setEditingStatus] = useState(null);

  const dow = new Date(date + "T00:00:00").getDay();
  const weekdayLabel = ["일", "월", "화", "수", "목", "금", "토"][dow];

  const rows = useMemo(() => {
    return employees
      .map((emp) => {
        const key = `${emp.id}_${date}`;
        const rec = attendanceIndex[key];
        const override = overrideIndex[key];
        const st = computeDayStatus(emp, date, rec, override);
        return { emp, rec, override, ...st };
      })
      .filter((r) => r.source === "manual" || (r.status !== "출근" && r.status !== "OFF"))
      .sort((a, b) => a.emp.name.localeCompare(b.emp.name));
  }, [employees, attendanceIndex, overrideIndex, date]);

  function changeStatus(row, newStatus) {
    if ((newStatus === "지각" || newStatus === "늦출") && row.rec) {
      removeDayAdjustment(row.emp.id, date);
      setLateExcused(row.rec.employeeId, date, newStatus === "늦출");
      setEditingKey(null);
      return;
    }
    // 일퇴/OFF/출근, 또는 원본 데이터 없는 지각·늦출은 분 단위 설정이 필요하므로 편집기를 연다
    setEditingKey(`${row.emp.id}_${date}`);
    setEditingStatus(newStatus);
  }

  function saveAdjustment(row, minutes, dedType) {
    setDayAdjustment(row.emp.id, date, editingStatus, dedType, minutes, "");
    setEditingKey(null);
    setEditingStatus(null);
  }

  function submitAdd() {
    if (!addEmp) return;
    const emp = employees.find((e) => e.id === addEmp);
    const key = `${addEmp}_${date}`;
    const rec = attendanceIndex[key];
    if ((addStatus === "지각" || addStatus === "늦출") && rec) {
      setLateExcused(rec.employeeId, date, addStatus === "늦출");
      setShowAdd(false);
      setAddEmp("");
      return;
    }
    setAddStep("adjust");
  }

  const existingLedgerFor = (row) => (ledger || []).find((l) => l.id === `DAYADJ_${row.emp.id}_${date}`);
  const existingLateLedgerFor = (row) => row.rec && (ledger || []).find((l) => l.id === `LATE_${row.rec.employeeId}_${date}`);

  return (
    <div
      onClick={onClose}
      style={{ position: "fixed", inset: 0, background: "rgba(10,20,18,0.45)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 1000, padding: 20 }}
    >
      <div
        onClick={(e) => e.stopPropagation()}
        style={{ background: "#fff", borderRadius: 14, width: "100%", maxWidth: 640, maxHeight: "88vh", overflow: "auto", boxShadow: "0 20px 60px rgba(0,0,0,0.25)" }}
      >
        <div style={{ background: COLORS.tealDark, color: "#fff", padding: "16px 22px", borderRadius: "14px 14px 0 0", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
          <div style={{ fontSize: 18, fontWeight: 800 }}>{fmtDate(date)} ({weekdayLabel})</div>
          <button onClick={onClose} style={{ background: "rgba(255,255,255,0.15)", border: "none", color: "#fff", borderRadius: 6, padding: "6px 10px", cursor: "pointer", fontSize: 13 }}>
            닫기 ✕
          </button>
        </div>

        <div style={{ padding: 20 }}>
          <button
            className="btn"
            style={{ background: COLORS.teal, color: "#fff", marginBottom: 14 }}
            onClick={() => {
              setShowAdd((v) => !v);
              setAddStep("pick");
              setAddEmp("");
            }}
          >
            {showAdd ? "취소" : "+ 추가"}
          </button>

          {showAdd && addStep === "pick" && (
            <div style={{ background: COLORS.bg, borderRadius: 8, padding: 12, marginBottom: 16, display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
              <select className="input" value={addEmp} onChange={(e) => setAddEmp(e.target.value)} style={{ width: 130 }}>
                <option value="">직원 선택</option>
                {employees.map((e) => (<option key={e.id} value={e.id}>{e.name}</option>))}
              </select>
              <select className="input" value={addStatus} onChange={(e) => setAddStatus(e.target.value)}>
                {Object.keys(STATUS_LABELS).map((s) => (<option key={s} value={s}>{s}</option>))}
              </select>
              <button className="btn" style={{ background: COLORS.teal, color: "#fff" }} onClick={submitAdd}>다음</button>
            </div>
          )}

          {showAdd && addStep === "adjust" && (
            <div style={{ background: COLORS.bg, borderRadius: 8, padding: 12, marginBottom: 16 }}>
              <div style={{ fontSize: 12.5, fontWeight: 700, marginBottom: 8 }}>
                {employees.find((e) => e.id === addEmp)?.name} · {addStatus}
              </div>
              <DayAdjustEditor
                status={addStatus}
                initialMinutes=""
                initialType="ot"
                autoMinutes={offDayMinutes(dow, employees.find((e) => e.id === addEmp))}
                onSave={(minutes, dedType) => {
                  setDayAdjustment(addEmp, date, addStatus, dedType, minutes, "");
                  setShowAdd(false);
                  setAddEmp("");
                  setAddStep("pick");
                }}
                onCancel={() => {
                  setShowAdd(false);
                  setAddEmp("");
                  setAddStep("pick");
                }}
              />
            </div>
          )}

          <table>
            <thead style={{ background: COLORS.tealSoft }}>
              <tr>
                <th>이름</th>
                <th>상태</th>
                <th>분 / 차감</th>
                <th></th>
              </tr>
            </thead>
            <tbody>
              {rows.length === 0 && (
                <tr>
                  <td colSpan={4} style={{ textAlign: "center", color: COLORS.sub, padding: 20 }}>
                    이 날짜에 표시할 기록이 없습니다. "+ 추가"로 기록을 남겨보세요.
                  </td>
                </tr>
              )}
              {rows.map((row) => {
                const key = `${row.emp.id}_${date}`;
                const isEditing = editingKey === key;
                const dayLedger = existingLedgerFor(row);
                const lateLedger = existingLateLedgerFor(row);
                return (
                  <React.Fragment key={row.emp.id}>
                    <tr>
                      <td style={{ fontWeight: 600 }}>{row.emp.name}</td>
                      <td>
                        <select
                          className="input"
                          style={{ padding: "4px 6px", fontSize: 12.5, fontWeight: 700, color: (STATUS_COLORS[row.status] || {}).fg }}
                          value={row.status}
                          onChange={(e) => changeStatus(row, e.target.value)}
                        >
                          {Object.keys(STATUS_LABELS).map((s) => (<option key={s} value={s}>{s}</option>))}
                        </select>
                      </td>
                      <td style={{ fontSize: 12.5 }}>
                        {row.status === "지각" || row.status === "늦출" ? (
                          lateLedger ? `${lateLedger.type === "ot" ? "OT" : "연차"} ${lateLedger.minutes}분 차감` : row.minutes ? `${row.minutes}분 (미차감)` : "-"
                        ) : dayLedger ? (
                          `${dayLedger.type === "ot" ? "OT" : dayLedger.type === "daehyu" ? "대휴" : "연차"} ${dayLedger.minutes}분 ${dayLedger.direction === "adjust" ? "가산" : "차감"}`
                        ) : (
                          "-"
                        )}
                      </td>
                      <td style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                        {(row.status === "지각" || row.status === "늦출") && row.rec && (
                          <button
                            className="btn"
                            style={{ background: COLORS.tealSoft, color: COLORS.tealDark, padding: "4px 10px" }}
                            onClick={() => { setEditingKey(key); setEditingStatus(row.status); }}
                          >
                            {lateLedger ? "차감 수정" : "차감 설정"}
                          </button>
                        )}
                        {(row.status === "일퇴" || row.status === "반차" || row.status === "OFF" || row.status === "출근" || ((row.status === "지각" || row.status === "늦출") && !row.rec)) && (
                          <button
                            className="btn"
                            style={{ background: COLORS.tealSoft, color: COLORS.tealDark, padding: "4px 10px" }}
                            onClick={() => { setEditingKey(key); setEditingStatus(row.status); }}
                          >
                            {dayLedger ? "수정" : "설정"}
                          </button>
                        )}
                        {row.source === "manual" && (
                          <button className="btn" style={{ background: "transparent", color: COLORS.sub, padding: "4px 8px", fontSize: 12 }} onClick={() => removeDayAdjustment(row.emp.id, date)}>
                            자동으로 되돌리기
                          </button>
                        )}
                      </td>
                    </tr>
                    {isEditing && (row.status === "지각" || row.status === "늦출") && row.rec && (
                      <tr>
                        <td colSpan={4}>
                          <div style={{ display: "flex", gap: 6, alignItems: "center", flexWrap: "wrap", background: COLORS.bg, padding: 8, borderRadius: 6 }}>
                            <input
                              type="number"
                              className="input"
                              style={{ padding: "4px 6px", fontSize: 12.5, width: 70 }}
                              placeholder="분"
                              defaultValue={lateLedger ? lateLedger.minutes : row.minutes || ""}
                              id={`ded-min-${key}`}
                            />
                            <span style={{ fontSize: 12, color: COLORS.sub }}>분</span>
                            <select className="input" style={{ padding: "4px 6px", fontSize: 12.5 }} defaultValue={lateLedger ? lateLedger.type : "ot"} id={`ded-type-${key}`}>
                              <option value="ot">OT에서 차감</option>
                              <option value="leave">연차에서 차감</option>
                            </select>
                            <button
                              className="btn"
                              style={{ background: COLORS.teal, color: "#fff", padding: "4px 10px" }}
                              onClick={() => {
                                const mEl = document.getElementById(`ded-min-${key}`);
                                const tEl = document.getElementById(`ded-type-${key}`);
                                setLateDeduction(row.rec.employeeId, date, tEl.value, mEl.value, "지각 시간차감");
                                setEditingKey(null);
                              }}
                            >
                              저장
                            </button>
                            <button className="btn" style={{ background: "transparent", color: COLORS.sub, padding: "4px 8px" }} onClick={() => setEditingKey(null)}>취소</button>
                          </div>
                        </td>
                      </tr>
                    )}
                    {isEditing && (row.status === "일퇴" || row.status === "반차" || row.status === "OFF" || row.status === "출근" || ((row.status === "지각" || row.status === "늦출") && !row.rec)) && (
                      <tr>
                        <td colSpan={4}>
                          <DayAdjustEditor
                            status={editingStatus}
                            initialMinutes={dayLedger ? dayLedger.minutes : row.minutes}
                            initialType={dayLedger ? dayLedger.type : "ot"}
                            autoMinutes={offDayMinutes(dow, row.emp)}
                            onSave={(minutes, dedType) => saveAdjustment(row, minutes, dedType)}
                            onCancel={() => setEditingKey(null)}
                          />
                        </td>
                      </tr>
                    )}
                  </React.Fragment>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}


function UploadTab({ fileRef, onFileInput, onDrop, uploadMsg, uploadLog, uploading, monthsInData, deleteMonthData }) {
  return (
    <>
      <div
        className="dropzone"
        onClick={() => !uploading && fileRef.current && fileRef.current.click()}
        onDrop={onDrop}
        onDragOver={(e) => e.preventDefault()}
        style={{ marginBottom: 16, opacity: uploading ? 0.6 : 1 }}
      >
        <div style={{ fontWeight: 700, color: COLORS.tealDark, marginBottom: 4 }}>
          {uploading ? "처리 중..." : "캡스 근태 엑셀 파일을 여기로 드래그하거나 클릭해서 업로드"}
        </div>
        <div style={{ fontSize: 12.5 }}>
          .xlsx 권장 (.xls도 지원) · 업로드하면 파일에 포함된 월(달)의 기존 데이터는 전부 삭제되고
          새 파일 내용으로 통째로 교체됩니다
        </div>
        <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }} onChange={onFileInput} disabled={uploading} />
      </div>

      {uploadMsg && (
        <div style={{ background: COLORS.tealSoft, color: COLORS.tealDark, padding: "10px 14px", borderRadius: 8, fontSize: 13.5, marginBottom: 16, fontWeight: 600 }}>
          {uploadMsg}
        </div>
      )}

      <div className="card" style={{ marginBottom: 16 }}>
        <div style={{ fontSize: 13, fontWeight: 700, marginBottom: 10, color: COLORS.tealDark }}>저장된 월별 데이터</div>
        <div style={{ fontSize: 12, color: COLORS.sub, marginBottom: 10 }}>
          잘못 올렸거나 더 이상 필요 없는 달의 근태 데이터를 통째로 삭제할 수 있습니다. 삭제 후에는 되돌릴 수 없습니다.
        </div>
        <table>
          <thead style={{ background: COLORS.tealSoft }}>
            <tr>
              <th>월</th>
              <th>저장된 기록 수</th>
              <th></th>
            </tr>
          </thead>
          <tbody>
            {monthsInData.length === 0 && (
              <tr>
                <td colSpan={3} style={{ textAlign: "center", color: COLORS.sub, padding: 20 }}>
                  저장된 근태 데이터가 없습니다.
                </td>
              </tr>
            )}
            {monthsInData.map(([month, count]) => (
              <tr key={month}>
                <td style={{ fontWeight: 600 }}>{month}</td>
                <td>{count}건</td>
                <td>
                  <button className="btn" style={{ background: "transparent", color: COLORS.red, padding: "4px 8px" }} onClick={() => deleteMonthData(month)}>
                    이 달 데이터 삭제
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <div className="card">
        <div style={{ fontSize: 13, fontWeight: 700, marginBottom: 10, color: COLORS.tealDark }}>업로드 이력</div>
        <table>
          <thead style={{ background: COLORS.tealSoft }}>
            <tr>
              <th>파일명</th>
              <th>업로드 시각</th>
              <th>대상 월</th>
              <th>처리 결과</th>
            </tr>
          </thead>
          <tbody>
            {uploadLog.length === 0 && (
              <tr>
                <td colSpan={4} style={{ textAlign: "center", color: COLORS.sub, padding: 20 }}>
                  아직 업로드된 파일이 없습니다.
                </td>
              </tr>
            )}
            {uploadLog.map((u) => (
              <tr key={u.id}>
                <td>{u.fileName}</td>
                <td>{new Date(u.uploadedAt).toLocaleString("ko-KR")}</td>
                <td>{u.months}</td>
                <td>{u.rowCount}건 반영 (기존 {u.removed}건 삭제 후 교체)</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </>
  );
}

function LedgerTab(props) {
  const {
    employees, ledger, employeeMap,
    ledgerEmp, setLedgerEmp, ledgerType, setLedgerType, ledgerDir, setLedgerDir,
    ledgerMinutes, setLedgerMinutes, ledgerDate, setLedgerDate, ledgerNote, setLedgerNote,
    addLedgerEntry, removeLedgerEntry,
  } = props;

  return (
    <>
      <div className="card" style={{ marginBottom: 16 }}>
        <div style={{ fontSize: 13, fontWeight: 700, marginBottom: 10, color: COLORS.tealDark }}>
          연차 / OT 사용·조정 기록 추가
        </div>
        <div style={{ fontSize: 12, color: COLORS.sub, marginBottom: 10 }}>
          자동 적립분(연차 발생, OT 적립)은 여기 입력할 필요 없습니다. 사용 내역이나 이전 잔여량(초기값
          조정)만 입력하세요. 조기퇴근은 규정상 차감하지 않습니다.
        </div>
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
          <select className="input" value={ledgerEmp} onChange={(e) => setLedgerEmp(e.target.value)} style={{ width: 120 }}>
            <option value="">직원 선택</option>
            {employees.map((e) => (
              <option key={e.id} value={e.id}>{e.name}</option>
            ))}
          </select>
          <select className="input" value={ledgerType} onChange={(e) => setLedgerType(e.target.value)}>
            <option value="leave">연차</option>
            <option value="ot">OT</option>
            <option value="daehyu">대휴</option>
          </select>
          <select className="input" value={ledgerDir} onChange={(e) => setLedgerDir(e.target.value)}>
            <option value="use">사용</option>
            <option value="adjust">조정(+)</option>
          </select>
          <input className="input" type="number" value={ledgerMinutes} onChange={(e) => setLedgerMinutes(e.target.value)} placeholder="분" style={{ width: 90 }} />
          <span style={{ fontSize: 12, color: COLORS.sub }}>분 (1일=480분, 반차=240분)</span>
          <input className="input" type="date" value={ledgerDate} onChange={(e) => setLedgerDate(e.target.value)} />
          <input className="input" placeholder="메모(선택)" value={ledgerNote} onChange={(e) => setLedgerNote(e.target.value)} style={{ width: 160 }} />
          <button className="btn" style={{ background: COLORS.teal, color: "#fff" }} onClick={addLedgerEntry}>추가</button>
        </div>
      </div>

      <div className="card">
        <table>
          <thead style={{ background: COLORS.tealSoft }}>
            <tr>
              <th>이름</th>
              <th>구분</th>
              <th>유형</th>
              <th>분</th>
              <th>날짜</th>
              <th>메모</th>
              <th></th>
            </tr>
          </thead>
          <tbody>
            {ledger.length === 0 && (
              <tr>
                <td colSpan={7} style={{ textAlign: "center", color: COLORS.sub, padding: 20 }}>등록된 기록이 없습니다.</td>
              </tr>
            )}
            {ledger.map((l) => (
              <tr key={l.id}>
                <td style={{ fontWeight: 600 }}>{employeeMap[l.employeeId] || "(삭제된 직원)"}</td>
                <td>{l.type === "leave" ? "연차" : l.type === "ot" ? "OT" : "대휴"}</td>
                <td>{l.direction === "use" ? "사용" : "조정(+)"}</td>
                <td>{l.minutes}분</td>
                <td>{fmtDate(l.date)}</td>
                <td style={{ color: COLORS.sub }}>{l.note || "-"}</td>
                <td>
                  <button className="btn" style={{ background: "transparent", color: COLORS.red, padding: "4px 8px" }} onClick={() => removeLedgerEntry(l.id)}>삭제</button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </>
  );
}

function EmployeesTab({
  employees, updateEmployee, removeEmployee,
  newName, setNewName, newTeam, setNewTeam, newPosition, setNewPosition, newHire, setNewHire, addEmployee,
}) {
  const [expandedId, setExpandedId] = useState(null);

  return (
    <>
      <div className="card" style={{ marginBottom: 16 }}>
        <div style={{ fontSize: 13, fontWeight: 700, marginBottom: 10, color: COLORS.tealDark }}>직원 추가</div>
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
          <input className="input" placeholder="이름" value={newName} onChange={(e) => setNewName(e.target.value)} style={{ width: 120 }} />
          <select className="input" value={newTeam} onChange={(e) => setNewTeam(e.target.value)}>
            {TEAM_ORDER.map((t) => (<option key={t} value={t}>{t}</option>))}
          </select>
          <select className="input" value={newPosition} onChange={(e) => setNewPosition(e.target.value)}>
            {POSITION_LIST.map((p) => (<option key={p} value={p}>{p}</option>))}
          </select>
          <input className="input" type="date" value={newHire} onChange={(e) => setNewHire(e.target.value)} />
          <button className="btn" style={{ background: COLORS.teal, color: "#fff" }} onClick={addEmployee}>추가</button>
        </div>
      </div>

      <div className="card" style={{ overflow: "auto" }}>
        <table>
          <thead style={{ background: COLORS.tealSoft }}>
            <tr>
              <th>이름</th><th>팀</th><th>직급</th><th>입사일</th>
              <th>연차 초기값(분)</th><th>OT 초기값(분)</th><th>근무시간</th><th></th>
            </tr>
          </thead>
          <tbody>
            {employees.map((e) => {
              const hasCustom = e.customSchedule && Object.keys(e.customSchedule).length > 0;
              const isOpen = expandedId === e.id;
              return (
                <React.Fragment key={e.id}>
                  <tr>
                    <td style={{ fontWeight: 600 }}>{e.name}</td>
                    <td>
                      <select className="input" style={{ padding: "4px 6px", fontSize: 12.5 }} value={e.team || "미지정"} onChange={(ev) => updateEmployee(e.id, { team: ev.target.value })}>
                        {TEAM_ORDER.map((t) => (<option key={t} value={t}>{t}</option>))}
                      </select>
                    </td>
                    <td>
                      <select className="input" style={{ padding: "4px 6px", fontSize: 12.5 }} value={e.position || ""} onChange={(ev) => updateEmployee(e.id, { position: ev.target.value })}>
                        <option value="">-</option>
                        {POSITION_LIST.map((p) => (<option key={p} value={p}>{p}</option>))}
                      </select>
                    </td>
                    <td>
                      <input type="date" className="input" style={{ padding: "4px 6px", fontSize: 12.5 }} value={e.hireDate} onChange={(ev) => updateEmployee(e.id, { hireDate: ev.target.value })} />
                    </td>
                    <td>
                      <input type="number" className="input" style={{ padding: "4px 6px", fontSize: 12.5, width: 90 }} value={e.openingLeaveMinutes} onChange={(ev) => updateEmployee(e.id, { openingLeaveMinutes: parseInt(ev.target.value, 10) || 0 })} />
                    </td>
                    <td>
                      <input type="number" className="input" style={{ padding: "4px 6px", fontSize: 12.5, width: 90 }} value={e.openingOTMinutes} onChange={(ev) => updateEmployee(e.id, { openingOTMinutes: parseInt(ev.target.value, 10) || 0 })} />
                    </td>
                    <td>
                      <button
                        className="btn"
                        style={{ background: hasCustom ? COLORS.amberSoft : COLORS.tealSoft, color: hasCustom ? COLORS.amber : COLORS.tealDark, padding: "4px 10px" }}
                        onClick={() => setExpandedId(isOpen ? null : e.id)}
                      >
                        {hasCustom ? "예외 설정됨" : "기본값 사용"} {isOpen ? "▲" : "▼"}
                      </button>
                    </td>
                    <td>
                      <button className="btn" style={{ background: "transparent", color: COLORS.red, padding: "4px 8px" }} onClick={() => removeEmployee(e.id)}>삭제</button>
                    </td>
                  </tr>
                  {isOpen && (
                    <tr>
                      <td colSpan={8} style={{ background: COLORS.bg, padding: 14 }}>
                        <WeeklyScheduleEditor employee={e} updateEmployee={updateEmployee} />
                      </td>
                    </tr>
                  )}
                </React.Fragment>
              );
            })}
          </tbody>
        </table>
      </div>
    </>
  );
}


