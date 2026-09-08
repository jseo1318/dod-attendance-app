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
const POSITION_LIST = ["원장", "실장", "팀장", "사원"];

const SPECIAL_SCHEDULE = {
  이보은: { customEnd: "18:00" },
  홍보미: { customStart: "10:00" },
};

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

function computeMetrics(record, employee) {
  const sched = SCHEDULE[record.dow];
  if (!sched) return { late: 0, otRaw: 0, otCredited: 0 };
  const startStr = (employee && employee.customStart) || sched.start;
  const endStr = (employee && employee.customEnd) || sched.end;
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
    customStart: row.custom_start || "",
    customEnd: row.custom_end || "",
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
    custom_start: e.customStart || "",
    custom_end: e.customEnd || "",
    active: e.active !== false,
  };
}

/* ───────────────────────── main component ───────────────────────── */

export default function App() {
  const [employees, setEmployees] = useState(null);
  const [attendance, setAttendance] = useState(null); // flat array of raw records
  const [ledger, setLedger] = useState(null);
  const [uploadLog, setUploadLog] = useState(null);

  const [tab, setTab] = useState("dashboard");
  const [asOf, setAsOf] = useState(toISO(new Date()));
  const [error, setError] = useState("");
  const [uploadMsg, setUploadMsg] = useState("");
  const [uploading, setUploading] = useState(false);
  const fileRef = useRef(null);
  const [collapsed, setCollapsed] = useState({});

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
      setError(`직원 목록을 불러오지 못했습니다: ${err.message}`);
      return;
    }
    setEmployees((data || []).map(empFromRow));
  }, []);

  const fetchAttendance = useCallback(async () => {
    const { data, error: err } = await supabase.from("attendance").select("*");
    if (err) {
      setError(`근태 데이터를 불러오지 못했습니다: ${err.message}`);
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
      }))
    );
  }, []);

  const fetchLedger = useCallback(async () => {
    const { data, error: err } = await supabase
      .from("ledger")
      .select("*")
      .order("created_at", { ascending: false });
    if (err) {
      setError(`원장 데이터를 불러오지 못했습니다: ${err.message}`);
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
      setError(`업로드 이력을 불러오지 못했습니다: ${err.message}`);
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

  useEffect(() => {
    if (!configOk) return;
    fetchEmployees();
    fetchAttendance();
    fetchLedger();
    fetchUploadLog();

    const channel = supabase
      .channel("dod-attendance-sync")
      .on("postgres_changes", { event: "*", schema: "public", table: "employees" }, fetchEmployees)
      .on("postgres_changes", { event: "*", schema: "public", table: "attendance" }, fetchAttendance)
      .on("postgres_changes", { event: "*", schema: "public", table: "ledger" }, fetchLedger)
      .on("postgres_changes", { event: "*", schema: "public", table: "upload_log" }, fetchUploadLog)
      .subscribe();

    return () => {
      supabase.removeChannel(channel);
    };
  }, [configOk, fetchEmployees, fetchAttendance, fetchLedger, fetchUploadLog]);

  /* ---- employees ---- */
  async function addEmployee() {
    if (!newName.trim()) return;
    const row = empToRow({
      id: `${newName.trim()}_${Date.now()}`,
      name: newName.trim(),
      team: newTeam,
      position: newPosition,
      hireDate: newHire,
      openingLeaveMinutes: 0,
      openingOTMinutes: 0,
      customStart: "",
      customEnd: "",
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
    const { error: err } = await supabase.from("employees").delete().eq("id", id);
    if (err) setError(`직원 삭제 실패: ${err.message}`);
    else fetchEmployees();
  }

  /* ---- upload: 파일에 포함된 월을 통째로 교체 ---- */
  async function handleFile(file) {
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
        const end = `${month}-31`;
        const { data: existing } = await supabase
          .from("attendance")
          .select("employee_id,date")
          .gte("date", start)
          .lte("date", end);
        removedCount += existing ? existing.length : 0;
        const { error: delErr } = await supabase
          .from("attendance")
          .delete()
          .gte("date", start)
          .lte("date", end);
        if (delErr) throw delErr;
      }

      const chunkSize = 500;
      for (let i = 0; i < records.length; i += chunkSize) {
        const chunk = records.slice(i, i + chunkSize);
        const { error: insErr } = await supabase.from("attendance").insert(chunk);
        if (insErr) throw insErr;
      }

      const existingIds = new Set(employees.map((e) => e.id));
      const existingNames = new Set(employees.map((e) => e.name));
      const toAdd = [];
      employeesFound.forEach((name, id) => {
        if (!existingIds.has(id) && !existingNames.has(name)) {
          toAdd.push(
            empToRow({
              id,
              name,
              team: "미지정",
              position: "",
              hireDate: "",
              openingLeaveMinutes: 0,
              openingOTMinutes: 0,
              customStart: "",
              customEnd: "",
              active: true,
            })
          );
        }
      });
      if (toAdd.length > 0) {
        const { error: empErr } = await supabase.from("employees").insert(toAdd);
        if (empErr) throw empErr;
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

      await Promise.all([fetchAttendance(), fetchEmployees(), fetchUploadLog()]);

      setUploadMsg(
        `${monthsSeen.join(", ")} 데이터를 통째로 교체했습니다 · 이번 업로드 ${records.length}건 반영 (기존 ${removedCount}건 삭제 후 재입력)` +
          (toAdd.length > 0 ? ` · 신규 인원 ${toAdd.length}명 자동 추가됨(팀/입사일 입력 필요)` : "")
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
  async function addLedgerEntry() {
    if (!ledgerEmp || !ledgerMinutes) return;
    const row = {
      id: `L${Date.now()}`,
      employee_id: ledgerEmp,
      type: ledgerType,
      direction: ledgerDir,
      minutes: Math.abs(parseInt(ledgerMinutes, 10) || 0),
      date: ledgerDate,
      note: ledgerNote.trim(),
    };
    const { error: err } = await supabase.from("ledger").insert(row);
    if (err) setError(`기록 추가 실패: ${err.message}`);
    else {
      setLedgerNote("");
      fetchLedger();
    }
  }
  async function removeLedgerEntry(id) {
    const { error: err } = await supabase.from("ledger").delete().eq("id", id);
    if (err) setError(`기록 삭제 실패: ${err.message}`);
    else fetchLedger();
  }

  const employeeMap = useMemo(() => {
    const m = {};
    (employees || []).forEach((e) => (m[e.id] = e.name));
    return m;
  }, [employees]);

  const summaryRows = useMemo(() => {
    if (!employees || !ledger || !attendance) return [];
    return employees.map((emp) => {
      const empRecords = attendance.filter(
        (r) => (r.employeeId === emp.id || r.employeeName === emp.name) && r.date <= asOf
      );
      const empLedger = ledger.filter((l) => l.employeeId === emp.id);

      let totalLateCount = 0;
      let totalLateMinutes = 0;
      let totalOtEarned = 0;
      let workedDays = 0;
      empRecords.forEach((r) => {
        if (r.checkin) workedDays++;
        const m = computeMetrics(r, emp);
        if (m.late > 0) {
          totalLateCount++;
          totalLateMinutes += m.late;
        }
        totalOtEarned += m.otCredited;
      });

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

  if (!employees || !attendance || !ledger || !uploadLog) {
    return <div style={{ padding: 40, fontFamily: FONT, color: COLORS.sub }}>불러오는 중...</div>;
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
        <div style={{ fontSize: 12, letterSpacing: 0.3, color: COLORS.teal, fontWeight: 700, marginBottom: 4 }}>
          DOD Dermatology Cheongdam
        </div>
        <h1 style={{ fontSize: 24, fontWeight: 800, margin: 0, color: COLORS.tealDark }}>근태 · 연차 · OT 통합관리</h1>
        <p style={{ fontSize: 13, color: COLORS.sub, marginTop: 6 }}>
          Supabase 데이터베이스에 저장됩니다 · 캡스 근태 엑셀을 업로드하면 지각·연장근무가 자동 반영되고 매달 계속 누적됩니다.
        </p>
      </header>

      <div style={{ display: "flex", gap: 20, alignItems: "center", marginBottom: 18, flexWrap: "wrap" }}>
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
          {[
            ["dashboard", "대시보드"],
            ["upload", "데이터 업로드"],
            ["ledger", "연차·OT 원장"],
            ["employees", "직원 관리"],
            ["settings", "규정 안내"],
          ].map(([key, label]) => (
            <div key={key} className={`tab ${tab === key ? "active" : ""}`} onClick={() => setTab(key)}>
              {label}
            </div>
          ))}
        </div>
        <div style={{ marginLeft: "auto", display: "flex", alignItems: "center", gap: 8 }}>
          <label style={{ fontSize: 12.5, color: COLORS.sub }}>기준일</label>
          <input type="date" className="input" value={asOf} onChange={(e) => setAsOf(e.target.value)} />
        </div>
      </div>

      {error && (
        <div style={{ background: COLORS.redSoft, color: COLORS.red, padding: "8px 12px", borderRadius: 6, fontSize: 13, marginBottom: 14 }}>
          {error}
        </div>
      )}

      {tab === "dashboard" && (
        <DashboardTab groupedByTeam={groupedByTeam} collapsed={collapsed} toggleTeam={toggleTeam} />
      )}
      {tab === "upload" && (
        <UploadTab
          fileRef={fileRef}
          onFileInput={onFileInput}
          onDrop={onDrop}
          uploadMsg={uploadMsg}
          uploadLog={uploadLog}
          uploading={uploading}
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
      {tab === "settings" && <SettingsTab />}
    </div>
  );
}

/* ───────────────────────── tabs ───────────────────────── */

function DashboardTab({ groupedByTeam, collapsed, toggleTeam }) {
  return (
    <>
      {groupedByTeam.map(({ team, rows }) => {
        const isClosed = !!collapsed[team];
        return (
          <div key={team} style={{ marginBottom: 16 }}>
            <div className="team-header" onClick={() => toggleTeam(team)}>
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
                      <th>근무일수</th>
                      <th>누적 지각횟수</th>
                      <th>누적 지각시간</th>
                      <th>누적 OT 적립</th>
                      <th>OT 사용</th>
                      <th>OT 잔여</th>
                      <th>연차 발생</th>
                      <th>연차 사용</th>
                      <th>연차 잔여</th>
                    </tr>
                  </thead>
                  <tbody>
                    {rows.map((r) => {
                      const otLow = r.otRemaining < 0;
                      const leaveLow = r.leaveRemaining <= DAY_MINUTES * 2 && r.leaveRemaining >= 0;
                      const leaveNeg = r.leaveRemaining < 0;
                      return (
                        <tr key={r.id}>
                          <td style={{ fontWeight: 600 }}>{r.name}</td>
                          <td style={{ color: COLORS.sub }}>{r.position || "-"}</td>
                          <td>{r.workedDays}</td>
                          <td>{r.totalLateCount}</td>
                          <td>{minutesToHM(r.totalLateMinutes)}</td>
                          <td>{minutesToHM(r.totalOtEarned)}</td>
                          <td>{minutesToHM(r.otUsed)}</td>
                          <td>
                            <Badge text={minutesToHM(r.otRemaining)} tone={otLow ? "red" : "teal"} />
                          </td>
                          <td>
                            {r.hireDate ? `${r.leaveCalc.days}일 (${minutesToDaysLabel(r.leaveEarnedMinutes)})` : "입사일 미입력"}
                          </td>
                          <td>{minutesToDaysLabel(r.leaveUsed)}</td>
                          <td>
                            {r.hireDate ? (
                              <Badge
                                text={minutesToDaysLabel(r.leaveRemaining)}
                                tone={leaveNeg ? "red" : leaveLow ? "amber" : "teal"}
                              />
                            ) : (
                              "-"
                            )}
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

function UploadTab({ fileRef, onFileInput, onDrop, uploadMsg, uploadLog, uploading }) {
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
                <td>{l.type === "leave" ? "연차" : "OT"}</td>
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
              <th>연차 초기값(분)</th><th>OT 초기값(분)</th><th>개인 출근기준</th><th>개인 퇴근기준</th><th></th>
            </tr>
          </thead>
          <tbody>
            {employees.map((e) => (
              <tr key={e.id}>
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
                  <input type="time" className="input" style={{ padding: "4px 6px", fontSize: 12.5 }} value={e.customStart || ""} onChange={(ev) => updateEmployee(e.id, { customStart: ev.target.value })} />
                </td>
                <td>
                  <input type="time" className="input" style={{ padding: "4px 6px", fontSize: 12.5 }} value={e.customEnd || ""} onChange={(ev) => updateEmployee(e.id, { customEnd: ev.target.value })} />
                </td>
                <td>
                  <button className="btn" style={{ background: "transparent", color: COLORS.red, padding: "4px 8px" }} onClick={() => removeEmployee(e.id)}>삭제</button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </>
  );
}

function SettingsTab() {
  const row = (label, value) => (
    <div style={{ display: "flex", padding: "10px 0", borderBottom: `1px solid ${COLORS.border}` }}>
      <div style={{ width: 180, color: COLORS.sub, fontSize: 13 }}>{label}</div>
      <div style={{ fontSize: 13.5 }}>{value}</div>
    </div>
  );
  return (
    <div className="card">
      <div style={{ fontSize: 14, fontWeight: 700, color: COLORS.tealDark, marginBottom: 10 }}>적용 중인 규정</div>
      {row("기본 근무시간", "월~목 09:30~19:00 · 금 09:30~21:00 · 토 09:30~16:00")}
      {row("개인별 예외", "이보은 퇴근기준 18:00 · 홍보미 출근기준 10:00 (직원 관리 탭에서 추가/수정 가능)")}
      {row("지각 기준", "각자 적용되는 출근기준시각 초과 시 지각, 유예시간 없음")}
      {row("조기퇴근", "규정 종료시각보다 일찍 퇴근해도 OT·연차 차감 없음")}
      {row("OT(연장근무) 계산", "종료시각 1분 초과부터 발생, 실제 초과분 × 1.5배(소수점 버림)로 적립 (예: 2분→3분, 3분→4분, 4분→6분)")}
      {row("연차 발생", "근로기준법 표준 — 입사 1년 미만: 개근 월 1일(최대 11일) · 1년 이상: 15일 + 매 2년마다 1일 가산(최대 25일)")}
      {row("연차 1일 환산", "8시간 = 480분")}
      {row("데이터 저장", "Supabase 데이터베이스 (Netlify에서 배포)")}
      {row("업로드 방식", "업로드한 파일에 포함된 월(달)의 기존 데이터를 전부 지우고 새로 반영 (완전 교체)")}
      <div style={{ marginTop: 14, fontSize: 12, color: COLORS.sub }}>
        ※ 규정이 바뀌면 src/App.jsx 상단의 SCHEDULE, OT_MULTIPLIER, DAY_MINUTES, SPECIAL_SCHEDULE 값만 수정하고
        다시 배포(git push)하면 반영됩니다.
      </div>
    </div>
  );
}
