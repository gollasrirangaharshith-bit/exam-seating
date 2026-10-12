import React, { useState, useMemo } from "react";
import * as XLSX from "xlsx";
import mammoth from "mammoth";
import * as pdfjsLib from "pdfjs-dist";
import pdfjsWorker from "pdfjs-dist/build/pdf.worker.min.mjs?url";
import JSZip from "jszip"; // new: needed for .pptx  ->  npm install jszip

pdfjsLib.GlobalWorkerOptions.workerSrc = pdfjsWorker;

// Starts at 2 because the two default ranges below already use ids rng-1 and rng-2.
// Starting at 0 caused the third range added to reuse id "rng-1", which made React
// confuse that new row with the first one — editing one row's Start/End appeared to
// change the other row's values too.
let rangeCounter = 2;
function newRange() {
  rangeCounter += 1;
  return { id: `rng-${rangeCounter}`, prefix: "", start: "", end: "" };
}

// One colour per branch, so each branch is easy to spot on the seating plan.
const BRANCH_COLORS = ["#0f766e", "#c2410c", "#4338ca", "#a16207", "#be185d", "#0369a1", "#4d7c0f", "#7e22ce", "#b91c1c", "#475569"];
function branchColor(order, branch) {
  const i = order ? order.indexOf(branch) : -1;
  return i < 0 ? "#94a3b8" : BRANCH_COLORS[i % BRANCH_COLORS.length];
}

function shuffle(arr) {
  const a = [...arr];
  for (let i = a.length - 1; i > 0; i -= 1) {
    const j = Math.floor(Math.random() * (i + 1));
    [a[i], a[j]] = [a[j], a[i]];
  }
  return a;
}

// Reads a PDF and returns text with one line per visual row, so parseTextLines works.
async function pdfToText(buf) {
  const pdf = await pdfjsLib.getDocument({ data: buf }).promise;
  let text = "";
  for (let p = 1; p <= pdf.numPages; p += 1) {
    const content = await (await pdf.getPage(p)).getTextContent();
    let lastY = null;
    for (const it of content.items) {
      const y = it.transform[5];
      if (lastY !== null) text += Math.abs(y - lastY) > 2 ? "\n" : "  ";
      text += it.str;
      lastY = y;
    }
    text += "\n";
  }
  return text;
}

// "CSE_Students_2.pdf" -> "CSE". Generic names like "More_Students" stay "Unknown".
function branchFromFileName(fileName) {
  const token = fileName.replace(/\.[^.]+$/, "").split(/[_\-\s.]/)[0].toUpperCase();
  return !token || /^(STUDENTS?|LIST|MORE|DATA|SAMPLE|ALL)$/.test(token) ? "Unknown" : token;
}

// Parse a 2D array of rows (from Excel/CSV) into student records.
// Expects a header row containing recognizable column names.
function parseSheetRows(rows) {
  const warnings = [];
  if (!rows || rows.length === 0) {
    return { students: [], warnings: ["The file appears to be empty."] };
  }
  const header = rows[0].map((h) => String(h ?? "").trim().toLowerCase());
  const findCol = (pattern) => header.findIndex((h) => pattern.test(h));
  const idx = {
    id: findCol(/^(roll|student)?[\s._-]*(no|number|id)\b|^roll/),
    name: findCol(/name/),
    branch: findCol(/branch|dept|department/),
    subject: findCol(/subject|paper|course/),
  };
  if (idx.id === -1) {
    warnings.push("Couldn't find a Roll No / ID column by name — using the first column.");
  }
  const students = [];
  for (let r = 1; r < rows.length; r += 1) {
    const row = rows[r] || [];
    const isBlank = row.every((c) => c === undefined || c === null || String(c).trim() === "");
    if (isBlank) continue;
    const id = String(row[idx.id !== -1 ? idx.id : 0] ?? "").trim();
    if (!id) continue;
    students.push({
      id,
      name: idx.name !== -1 ? String(row[idx.name] ?? "").trim() : "",
      branch: idx.branch !== -1 ? String(row[idx.branch] ?? "").trim() || "Unknown" : "Unknown",
      subject: idx.subject !== -1 ? String(row[idx.subject] ?? "").trim() : "",
    });
  }
  return { students, warnings };
}

// Best-effort parser for plain text extracted from Word/PDF/PPTX files.
// Expects each student on its own line, fields separated by commas, tabs, or multiple spaces.
function parseTextLines(text) {
  const lines = text
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter(Boolean);
  const warnings = [];
  const students = [];
  let skipped = 0;
  for (const line of lines) {
    if (/^(roll|id)\b/i.test(line)) continue; // likely a header line
    const parts = line
      .split(/\t|,|\s{2,}/)
      .map((p) => p.trim())
      .filter(Boolean);
    if (parts.length < 1 || !parts[0]) {
      skipped += 1;
      continue;
    }
    const [id, name = "", branch = "Unknown", subject = ""] = parts;
    students.push({ id, name, branch: branch || "Unknown", subject });
  }
  if (skipped > 0) {
    warnings.push(`${skipped} line(s) in the document couldn't be read and were skipped.`);
  }
  if (students.length === 0) {
    warnings.push(
      "No student rows were recognized. For Word/PDF/PPTX files, each student should be on its own line/slide, e.g. 'CSE001, Priya Sharma, CSE, Maths'."
    );
  }
  return { students, warnings };
}

// Parse PPTX student data by extracting text from each slide.
// One student per slide; the text boxes on a slide are joined with tabs as fields.
async function parsePptxStudents(file) {
  const zip = await JSZip.loadAsync(await file.arrayBuffer());
  const slidePaths = Object.keys(zip.files)
    .filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
    .sort((a, b) => {
      const na = Number(a.match(/slide(\d+)\.xml$/)?.[1] || 0);
      const nb = Number(b.match(/slide(\d+)\.xml$/)?.[1] || 0);
      return na - nb;
    });

  if (slidePaths.length === 0) {
    return { students: [], warnings: ["No slides were found in this PPTX file."] };
  }

  const parser = new DOMParser();
  const lines = [];
  for (const path of slidePaths) {
    const xml = await zip.files[path].async("text");
    const doc = parser.parseFromString(xml, "application/xml");
    const texts = Array.from(doc.getElementsByTagName("a:t"))
      .map((node) => (node.textContent || "").trim())
      .filter(Boolean);
    if (texts.length > 0) lines.push(texts.join("\t"));
  }
  return parseTextLines(lines.join("\n"));
}

// Parse a 2D array of rows (Excel/CSV) into a flat list of invigilator names.
// Uses a "name" column if there's a recognizable header, otherwise treats every
// non-empty cell in the sheet as a name.
function parseTeacherSheet(rows) {
  if (!rows || rows.length === 0) {
    return { names: [], warnings: ["The file appears to be empty."] };
  }
  const header = rows[0].map((h) => String(h ?? "").trim().toLowerCase());
  const nameIdx = header.findIndex((h) => /name/.test(h));
  const names = [];

  if (nameIdx !== -1) {
    for (let r = 1; r < rows.length; r += 1) {
      const v = String(rows[r]?.[nameIdx] ?? "").trim();
      if (v) names.push(v);
    }
  } else {
    // No recognizable header — treat every non-empty cell as a name.
    for (const row of rows) {
      for (const cell of row || []) {
        const v = String(cell ?? "").trim();
        if (v) names.push(v);
      }
    }
  }
  return { names, warnings: [] };
}

// Parse plain text (from Word/PDF) into a flat list of invigilator names,
// one per line or comma-separated — same convention as the manual text box.
function parseTeacherText(text) {
  const names = text
    .split(/[\n,]/)
    .map((t) => t.trim())
    .filter(Boolean);
  return { names, warnings: [] };
}

// ---------- room plan: rooms + layout + branch rules (read from a file) ----------

// "CSE, ECE" / "CSE and ECE" / "CSE+ECE" / "CSE / ECE" -> ["CSE", "ECE"]
function splitBranchList(text) {
  return String(text)
    .split(/[,;|/&+]|\band\b/i)
    .map((s) => s.trim())
    .filter(Boolean);
}

// Understands the ways people naturally write a room's branch rule:
//   (blank), ALL, any, mixed, everyone      -> any branch
//   CSE / CSE only / only CSE               -> just that branch
//   CSE, ECE / CSE and ECE / CSE + ECE      -> those branches mixed together
//   ALL except ECE / not ECE / except ECE   -> every branch but those
//   CSE:30, ECE:20 / CSE (30) / CSE=30      -> allowed, with a maximum per branch
function parseBranchCondition(raw) {
  const text = String(raw ?? "").trim();
  const cond = { any: false, only: [], except: [], quotas: {} };

  if (
    !text ||
    /^(all|any|every(one| branch(es)?)?|mixed|open|\*|no restrictions?)( branch(es)?| students?)?$/i.test(text)
  ) {
    cond.any = true;
    return cond;
  }

  const exceptMatch = text.match(
    /^(?:(?:all|any|every(?:one)?|mixed)(?:\s+branch(?:es)?)?\s*)?(?:except|excluding|exclude|but not|but|not|without|other than)\s+(.+)$/i
  );
  if (exceptMatch) {
    cond.any = true;
    cond.except = splitBranchList(exceptMatch[1]);
    return cond;
  }

  const cleaned = text.replace(
    /\b(only|just|exclusively|branches|branch|students|student|departments|department|depts|dept)\b\.?/gi,
    " "
  );
  for (const token of splitBranchList(cleaned)) {
    const q = token.match(/^(.+?)\s*(?:[:=]|\()\s*(\d+)\s*\)?\s*$/);
    if (q) {
      const name = q[1].trim();
      cond.only.push(name);
      cond.quotas[name.toLowerCase()] = parseInt(q[2], 10);
    } else {
      cond.only.push(token);
    }
  }
  if (cond.only.length === 0) cond.any = true;
  return cond;
}

function describeCondition(cond) {
  if (cond.any && cond.except.length) return `All except ${cond.except.join(", ")}`;
  if (cond.any) return "Any branch";
  const parts = cond.only.map((n) => {
    const q = cond.quotas[n.toLowerCase()];
    return q !== undefined ? `${n} (max ${q})` : n;
  });
  return parts.length === 1 ? `${parts[0]} only` : parts.join(" + ");
}

// Matches the names in a rule against the branches that really exist in the student
// list (case-insensitive) and returns what the room may seat.
function resolveCondition(cond, branchOrder) {
  const lookup = {};
  branchOrder.forEach((b) => {
    lookup[b.toLowerCase()] = b;
  });
  const unknown = [];
  let allowed = [];
  if (cond.any) {
    const ex = new Set(cond.except.map((n) => n.toLowerCase()));
    cond.except.forEach((n) => {
      if (!lookup[n.toLowerCase()]) unknown.push(n);
    });
    allowed = branchOrder.filter((b) => !ex.has(b.toLowerCase()));
  } else {
    for (const n of cond.only) {
      const m = lookup[n.toLowerCase()];
      if (!m) unknown.push(n);
      else if (!allowed.includes(m)) allowed.push(m);
    }
  }
  const quotas = {};
  Object.entries(cond.quotas).forEach(([k, v]) => {
    if (lookup[k]) quotas[lookup[k]] = v;
  });
  return { allowed, quotas, unknown };
}

// Works out which column holds what, from one heading row. Returns null if the row
// doesn't look like a heading for a room plan (needs rows, benches and seats columns).
function detectRoomPlanColumns(headingRow) {
  const header = (headingRow || []).map((h) => String(h ?? "").trim().toLowerCase());
  const used = new Set();
  const pick = (test) => {
    const found = [];
    header.forEach((h, i) => {
      if (!used.has(i) && h && test(h)) found.push(i);
    });
    found.forEach((i) => used.add(i));
    return found;
  };
  const condCols = pick((h) => /branch|dept|department|condition|allow|rule|restrict|stream/.test(h));
  const seatCols = pick((h) => /seat|student|person|people|occup/.test(h));
  const rowCols = pick((h) => /row/.test(h) && !/bench/.test(h));
  const benchCols = pick((h) => /bench/.test(h));
  const nameCols = pick((h) => /room|hall|class|lab|venue/.test(h));
  if (!rowCols.length || !benchCols.length || !seatCols.length) return null;
  return {
    name: nameCols.length ? nameCols[0] : -1,
    rows: rowCols[0],
    benches: benchCols[0],
    seats: seatCols[0],
    cond: condCols,
  };
}

// Reads a 2D array of rows (Excel/CSV) describing the rooms. Finds the heading row
// (even if there's a title above it) and the columns by their headings, in any order;
// if there are no usable headings, falls back to Room, Rows, Benches, Seats, Branches.
function parseRoomPlanSheet(rows) {
  const warnings = [];
  if (!rows || rows.length === 0) {
    return { rooms: [], warnings: ["The file appears to be empty."] };
  }

  let idx = null;
  let start = 1;
  for (let h = 0; h < Math.min(rows.length, 6) && !idx; h += 1) {
    idx = detectRoomPlanColumns(rows[h]);
    if (idx) start = h + 1;
  }
  if (!idx) {
    const widest = Math.max(...rows.map((r) => (r ? r.length : 0)));
    idx = {
      name: 0,
      rows: 1,
      benches: 2,
      seats: 3,
      cond: Array.from({ length: Math.max(0, widest - 4) }, (_, i) => i + 4),
    };
    const firstIsData =
      Number.isFinite(parseInt(rows[0]?.[1], 10)) && Number.isFinite(parseInt(rows[0]?.[2], 10));
    start = firstIsData ? 0 : 1;
    warnings.push(
      "Couldn't match the column headings, so columns were read in this order: Room, Rows, Benches per row, Seats per bench, Branches allowed."
    );
  }

  const out = [];
  for (let r = start; r < rows.length; r += 1) {
    const row = rows[r] || [];
    if (row.every((c) => c === undefined || c === null || String(c).trim() === "")) continue;
    const label = idx.name >= 0 ? String(row[idx.name] ?? "").trim() : "";
    const nRows = parseInt(row[idx.rows], 10);
    const nBenches = parseInt(row[idx.benches], 10);
    const nSeats = parseInt(row[idx.seats], 10);
    if (![nRows, nBenches, nSeats].every((n) => Number.isFinite(n) && n >= 1)) {
      warnings.push(
        `Line ${r + 1}${label ? ` (${label})` : ""}: rows, benches and seats must be whole numbers — skipped.`
      );
      continue;
    }
    const condText = idx.cond
      .map((i) => String(row[i] ?? "").trim())
      .filter(Boolean)
      .join(", ");
    const cond = parseBranchCondition(condText);
    out.push({
      name: label,
      rows: nRows,
      benches: nBenches,
      seats: nSeats,
      condition: cond,
      conditionText: describeCondition(cond),
    });
  }
  if (out.length === 0) warnings.push("No valid rooms were found.");
  return { rooms: out, warnings };
}

// Reads rooms from plain text (Word/PDF). One room per line:
//   Room 101, 3, 10, 2, CSE only
//   Room 102 <tab> 3 <tab> 10 <tab> 2 <tab> CSE, ECE
function parseRoomPlanText(text) {
  const warnings = [];
  const out = [];
  let skipped = 0;
  const lines = String(text).split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
  for (const line of lines) {
    const parts = line.split(/\t|\s{2,}|,|\|/).map((p) => p.trim()).filter(Boolean);
    const nums = [parts[1], parts[2], parts[3]].map((p) => parseInt(p, 10));
    if (parts.length < 4 || !nums.every((n) => Number.isFinite(n) && n >= 1)) {
      // A heading line such as "Room, Rows, Benches, Seats, Branches" is expected - don't warn.
      if (!/row|bench|seat/i.test(line)) skipped += 1;
      continue;
    }
    const cond = parseBranchCondition(parts.slice(4).join(", "));
    out.push({
      name: parts[0],
      rows: nums[0],
      benches: nums[1],
      seats: nums[2],
      condition: cond,
      conditionText: describeCondition(cond),
    });
  }
  if (skipped > 0) warnings.push(`${skipped} line(s) couldn't be read as a room and were skipped.`);
  if (out.length === 0) {
    warnings.push(
      "No rooms were recognized. Put one room per line, e.g. 'Room 101, 3, 10, 2, CSE only' (room, rows, benches per row, seats per bench, branches allowed)."
    );
  }
  return { rooms: out, warnings };
}

// Seats students room by room, following each room's own layout and branch rule.
// Rooms are filled in the order they appear in the plan. A room never takes a branch
// its rule doesn't allow, and never exceeds a branch's per-room maximum.
function seatByRoomPlan({ plan, queues, branchOrder, avoidSameBranch }) {
  let conflicts = 0;
  const warnings = [];
  const mentioned = new Set();
  const rooms = [];

  for (const room of plan) {
    const { allowed, quotas, unknown } = resolveCondition(room.condition, branchOrder);
    allowed.forEach((b) => mentioned.add(b));
    if (unknown.length) {
      warnings.push(
        `${room.name}: "${unknown.join('", "')}" isn't a branch in the student list, so it was ignored.`
      );
    }
    if (allowed.length === 0) {
      warnings.push(`${room.name}: its rule doesn't match any branch, so this room stays empty.`);
    }

    const seatedHere = {};
    allowed.forEach((b) => {
      seatedHere[b] = 0;
    });
    const eligible = () =>
      allowed.filter(
        (b) => queues[b].length > 0 && (quotas[b] === undefined || seatedHere[b] < quotas[b])
      );
    // A room that may only hold one branch can't keep branches apart on a bench, and
    // that's the person's own rule - so it isn't counted as a clash.
    const canMix = allowed.length > 1;

    const usedBranches = new Set();
    const rows = [];
    for (let r = 0; r < room.rows; r += 1) {
      const benches = [];
      for (let b = 0; b < room.benches; b += 1) {
        const usedThisBench = new Set();
        const seats = [];
        for (let s = 0; s < room.seats; s += 1) {
          const candidates = eligible();
          if (candidates.length === 0) {
            seats.push(null);
            continue;
          }
          let pick;
          if (avoidSameBranch && canMix) {
            let fresh = candidates.filter((x) => !usedThisBench.has(x));
            if (fresh.length === 0) {
              fresh = candidates;
              conflicts += 1;
            }
            fresh = [...fresh].sort((a, c) => queues[c].length - queues[a].length);
            pick = fresh[0];
          } else {
            pick = candidates[0];
          }
          seats.push(queues[pick].shift());
          seatedHere[pick] += 1;
          usedThisBench.add(pick);
          usedBranches.add(pick);
        }
        benches.push(seats);
      }
      rows.push(benches);
    }

    rooms.push({
      name: room.name,
      conditionText: room.conditionText,
      allowedBranches: allowed.filter((b) => usedBranches.has(b)),
      layout: { rows: room.rows, benches: room.benches, seats: room.seats },
      rows,
    });
  }

  branchOrder.forEach((b) => {
    const left = queues[b].length;
    if (left === 0) return;
    warnings.push(
      mentioned.has(b)
        ? `${left} ${b} student(s) couldn't be seated — the rooms whose rules allow ${b} are full.`
        : `${left} ${b} student(s) couldn't be seated — no room in the plan allows ${b}.`
    );
  });

  return { rooms, conflicts, warnings };
}

export default function ExamSeatingRegister() {
  const [inputMode, setInputMode] = useState("manual"); // "manual" | "upload"

  const [ranges, setRanges] = useState([
    { id: "rng-1", prefix: "CSE", start: "1", end: "60" },
    { id: "rng-2", prefix: "ECE", start: "1", end: "40" },
  ]);

  const [uploadedStudents, setUploadedStudents] = useState([]);
  const [duplicatesRemoved, setDuplicatesRemoved] = useState(0);
  const [uploadWarnings, setUploadWarnings] = useState([]);
  const [uploadError, setUploadError] = useState("");
  const [uploadFileNames, setUploadFileNames] = useState([]);
  const [parsing, setParsing] = useState(false);

  const [classrooms, setClassrooms] = useState(3);
  const [rowsPerRoom, setRowsPerRoom] = useState(3);
  const [benchesPerRow, setBenchesPerRow] = useState(10);
  const [seatsPerBench, setSeatsPerBench] = useState(2);
  const [roomMode, setRoomMode] = useState("same"); // "same" | "plan"
  const [roomPlan, setRoomPlan] = useState([]);
  const [roomPlanFileNames, setRoomPlanFileNames] = useState([]);
  const [roomPlanDuplicatesRemoved, setRoomPlanDuplicatesRemoved] = useState(0);
  const [roomPlanWarnings, setRoomPlanWarnings] = useState([]);
  const [roomPlanError, setRoomPlanError] = useState("");
  const [roomPlanParsing, setRoomPlanParsing] = useState(false);
  const [branchesPerRoom, setBranchesPerRoom] = useState(2);
  const [branchGroupMode, setBranchGroupMode] = useState("auto"); // "auto" | "custom"
  const [customGroupsText, setCustomGroupsText] = useState("");
  const [avoidSameBranch, setAvoidSameBranch] = useState(true);
  const [teacherInputMode, setTeacherInputMode] = useState("manual"); // "manual" | "upload"
  const [teacherText, setTeacherText] = useState(
    "Dr. Rao\nProf. Iyer\nMs. Fernandes\nMr. Sharma\nDr. Khan\nProf. Nair"
  );
  const [uploadedTeachers, setUploadedTeachers] = useState([]);
  const [teacherUploadFileNames, setTeacherUploadFileNames] = useState([]);
  const [teacherDuplicatesRemoved, setTeacherDuplicatesRemoved] = useState(0);
  const [teacherWarnings, setTeacherWarnings] = useState([]);
  const [teacherUploadError, setTeacherUploadError] = useState("");
  const [teacherParsing, setTeacherParsing] = useState(false);
  const [teachersPerRoom, setTeachersPerRoom] = useState(1);
  const [result, setResult] = useState(null);
  const [activeRoom, setActiveRoom] = useState(0);
  const [viewMode, setViewMode] = useState("sheet"); // "sheet" | "cards"
  const [noticeTitle, setNoticeTitle] = useState(
    "Exam: AY 2026-27 (ODD) | MSE-1 | SEMESTER-5/7"
  );
  const [noticeDate, setNoticeDate] = useState("");
  const [noticeTime, setNoticeTime] = useState("");
  const [roomNumbersText, setRoomNumbersText] = useState("");
  const [error, setError] = useState("");

  function updateRange(id, field, value) {
    setRanges((prev) =>
      prev.map((r) => (r.id === id ? { ...r, [field]: value } : r))
    );
  }

  function addRange() {
    setRanges((prev) => [...prev, newRange()]);
  }

  function removeRange(id) {
    setRanges((prev) => prev.filter((r) => r.id !== id));
  }

  async function handleFileUpload(e) {
    const files = Array.from(e.target.files || []);
    e.target.value = ""; // lets the same file be picked again later
    if (!files.length) return;

    setUploadError("");
    setParsing(true);

    const collected = [];
    const newWarnings = [];
    const loadedNames = [];

    // Each file has its own try/catch, so one bad file never blocks the rest.
    for (const file of files) {
      try {
        const ext = file.name.split(".").pop().toLowerCase();
        let parsed;

        if (["xlsx", "xls", "csv"].includes(ext)) {
          const wb = XLSX.read(await file.arrayBuffer(), { type: "array" });
          const sheet = wb.Sheets[wb.SheetNames[0]];
          parsed = parseSheetRows(XLSX.utils.sheet_to_json(sheet, { header: 1 }));
        } else if (ext === "docx") {
          const { value } = await mammoth.extractRawText({ arrayBuffer: await file.arrayBuffer() });
          parsed = parseTextLines(value);
        } else if (ext === "pdf") {
          parsed = parseTextLines(await pdfToText(await file.arrayBuffer()));
        } else if (ext === "pptx") {
          parsed = await parsePptxStudents(file);
        } else {
          newWarnings.push(`${file.name}: unsupported file type ".${ext}", skipped.`);
          continue;
        }

        // No Branch column? Use the branch in the file name (e.g. ECE_Students.xlsx).
        const fallback = branchFromFileName(file.name);
        parsed.students.forEach((s) => {
          if (s.branch === "Unknown") s.branch = fallback;
        });

        collected.push(...parsed.students);
        newWarnings.push(...parsed.warnings.map((w) => `${file.name}: ${w}`));
        if (parsed.students.length > 0) loadedNames.push(file.name);
      } catch (err) {
        newWarnings.push(`${file.name}: couldn't be read (${err.message}), skipped.`);
      }
    }

    // Merge into the existing list; drop duplicate IDs (case-insensitive).
    const seen = new Set(uploadedStudents.map((s) => s.id.trim().toLowerCase()));
    const fresh = [];
    let dupes = 0;
    for (const s of collected) {
      const key = s.id.trim().toLowerCase();
      if (!key) continue;
      if (seen.has(key)) {
        dupes += 1;
        continue;
      }
      seen.add(key);
      fresh.push(s);
    }

    setUploadedStudents((prev) => [...prev, ...fresh]);
    setDuplicatesRemoved((c) => c + dupes);
    setUploadFileNames((prev) => [...prev, ...loadedNames.filter((n) => !prev.includes(n))]);
    setUploadWarnings((prev) => [...prev, ...newWarnings]); // append, don't replace
    if (collected.length === 0) {
      setUploadError("No valid student rows were found in the selected file(s).");
    }
    setParsing(false);
  }

  function clearUploadedStudents() {
    setUploadedStudents([]);
    setDuplicatesRemoved(0);
    setUploadWarnings([]);
    setUploadError("");
    setUploadFileNames([]);
  }

  async function handleTeacherFileUpload(e) {
    const files = Array.from(e.target.files || []);
    e.target.value = ""; // lets the same file be picked again later
    if (!files.length) return;

    setTeacherUploadError("");
    setTeacherParsing(true);

    const collected = [];
    const newWarnings = [];
    const loadedNames = [];

    // Each file has its own try/catch, so one bad file never blocks the rest.
    for (const file of files) {
      try {
        const ext = file.name.split(".").pop().toLowerCase();
        let parsed;

        if (["xlsx", "xls", "csv"].includes(ext)) {
          const wb = XLSX.read(await file.arrayBuffer(), { type: "array" });
          const sheet = wb.Sheets[wb.SheetNames[0]];
          parsed = parseTeacherSheet(XLSX.utils.sheet_to_json(sheet, { header: 1 }));
        } else if (ext === "docx") {
          const { value } = await mammoth.extractRawText({ arrayBuffer: await file.arrayBuffer() });
          parsed = parseTeacherText(value);
        } else if (ext === "pdf") {
          parsed = parseTeacherText(await pdfToText(await file.arrayBuffer()));
        } else {
          newWarnings.push(`${file.name}: unsupported file type ".${ext}", skipped.`);
          continue;
        }

        collected.push(...parsed.names);
        newWarnings.push(...parsed.warnings.map((w) => `${file.name}: ${w}`));
        if (parsed.names.length > 0) loadedNames.push(file.name);
        else newWarnings.push(`${file.name}: no names were found.`);
      } catch (err) {
        newWarnings.push(`${file.name}: couldn't be read (${err.message}), skipped.`);
      }
    }

    // Merge into the existing list; drop duplicate names (case-insensitive).
    const seen = new Set(uploadedTeachers.map((n) => n.trim().toLowerCase()));
    const fresh = [];
    let dupes = 0;
    for (const n of collected) {
      const key = n.trim().toLowerCase();
      if (!key) continue;
      if (seen.has(key)) {
        dupes += 1;
        continue;
      }
      seen.add(key);
      fresh.push(n.trim());
    }

    setUploadedTeachers((prev) => [...prev, ...fresh]);
    setTeacherDuplicatesRemoved((c) => c + dupes);
    setTeacherUploadFileNames((prev) => [...prev, ...loadedNames.filter((n) => !prev.includes(n))]);
    setTeacherWarnings((prev) => [...prev, ...newWarnings]);
    if (collected.length === 0) {
      setTeacherUploadError("No names were found in the selected file(s).");
    }
    setTeacherParsing(false);
  }

  function clearUploadedTeachers() {
    setUploadedTeachers([]);
    setTeacherDuplicatesRemoved(0);
    setTeacherWarnings([]);
    setTeacherUploadError("");
    setTeacherUploadFileNames([]);
  }

  async function handleRoomPlanUpload(e) {
    const files = Array.from(e.target.files || []);
    e.target.value = ""; // lets the same file be picked again later
    if (!files.length) return;

    setRoomPlanError("");
    setRoomPlanParsing(true);

    const collected = [];
    const newWarnings = [];
    const loadedNames = [];

    // Each file has its own try/catch, so one bad file never blocks the rest.
    for (const file of files) {
      try {
        const ext = file.name.split(".").pop().toLowerCase();
        let parsed;

        if (["xlsx", "xls", "csv"].includes(ext)) {
          const wb = XLSX.read(await file.arrayBuffer(), { type: "array" });
          // Read every sheet, so a plan split across tabs (e.g. one per block) still works.
          const found = [];
          let firstWarnings = null;
          wb.SheetNames.forEach((sheetName) => {
            const one = parseRoomPlanSheet(
              XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1 })
            );
            if (firstWarnings === null) firstWarnings = one.warnings;
            if (one.rooms.length > 0) {
              found.push(...one.rooms);
              newWarnings.push(
                ...one.warnings.map((w) => `${file.name} (${sheetName}): ${w}`)
              );
            }
          });
          parsed = { rooms: found, warnings: found.length ? [] : firstWarnings || [] };
        } else if (ext === "docx") {
          const { value } = await mammoth.extractRawText({ arrayBuffer: await file.arrayBuffer() });
          parsed = parseRoomPlanText(value);
        } else if (ext === "pdf") {
          parsed = parseRoomPlanText(await pdfToText(await file.arrayBuffer()));
        } else {
          newWarnings.push(`${file.name}: unsupported file type ".${ext}", skipped.`);
          continue;
        }

        collected.push(...parsed.rooms);
        newWarnings.push(...parsed.warnings.map((w) => `${file.name}: ${w}`));
        if (parsed.rooms.length > 0) loadedNames.push(file.name);
      } catch (err) {
        newWarnings.push(`${file.name}: couldn't be read (${err.message}), skipped.`);
      }
    }

    // Merge into the existing plan; a room name that's already there is skipped.
    const seen = new Set(roomPlan.map((r) => r.name.trim().toLowerCase()));
    const fresh = [];
    let dupes = 0;
    for (const room of collected) {
      const name = room.name || String(roomPlan.length + fresh.length + 1);
      const key = name.trim().toLowerCase();
      if (seen.has(key)) {
        dupes += 1;
        continue;
      }
      seen.add(key);
      fresh.push({ ...room, name });
    }

    setRoomPlan((prev) => [...prev, ...fresh]);
    setRoomPlanDuplicatesRemoved((c) => c + dupes);
    setRoomPlanFileNames((prev) => [...prev, ...loadedNames.filter((n) => !prev.includes(n))]);
    setRoomPlanWarnings((prev) => [...prev, ...newWarnings]);
    if (collected.length === 0) {
      setRoomPlanError("No rooms were found in the selected file(s).");
    }
    setRoomPlanParsing(false);
  }

  function clearRoomPlan() {
    setRoomPlan([]);
    setRoomPlanFileNames([]);
    setRoomPlanDuplicatesRemoved(0);
    setRoomPlanWarnings([]);
    setRoomPlanError("");
  }

  // Builds a small example sheet in the browser so the expected format is obvious.
  function downloadSampleRoomPlan() {
    const sheetRows = [
      ["Room", "Rows", "Benches per row", "Seats per bench", "Branches allowed"],
      ["101", 3, 5, 2, "CSE only"],
      ["102", 3, 5, 2, "CSE, ECE"],
      ["103", 3, 5, 2, "ECE, ISE"],
      ["104", 3, 5, 2, "ALL except CSE"],
      ["105", 3, 5, 2, "AIDS:15, AIML:15"],
      ["106", 3, 5, 2, "ALL"],
    ];
    const ws = XLSX.utils.aoa_to_sheet(sheetRows);
    ws["!cols"] = [{ wch: 10 }, { wch: 8 }, { wch: 16 }, { wch: 16 }, { wch: 28 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Room plan");
    XLSX.writeFile(wb, "room_plan_sample.xlsx");
  }

  function generate() {
    setError("");

    const usingPlan = roomMode === "plan";
    const branchOrder = [];
    const queues = {};
    const studentDetails = {}; // id -> { name, subject }

    if (inputMode === "upload") {
      if (uploadedStudents.length === 0) {
        setError("Upload a student list first.");
        return;
      }
      for (const s of uploadedStudents) {
        const branch = s.branch || "Unknown";
        if (!queues[branch]) {
          queues[branch] = [];
          branchOrder.push(branch);
        }
        queues[branch].push(s.id);
        studentDetails[s.id] = { name: s.name, subject: s.subject };
      }
    } else {
      const cleanRanges = ranges
        .map((r, idx) => ({
          ...r,
          start: parseInt(r.start, 10),
          end: parseInt(r.end, 10),
          branch: r.prefix.trim() || `Branch ${idx + 1}`,
        }))
        .filter((r) => Number.isFinite(r.start) && Number.isFinite(r.end));

      if (cleanRanges.length === 0) {
        setError("Add at least one valid ID range before generating.");
        return;
      }
      for (const r of cleanRanges) {
        if (r.start > r.end) {
          setError(
            `In the "${r.branch}" range, start (${r.start}) is after end (${r.end}).`
          );
          return;
        }
      }

      for (const r of cleanRanges) {
        if (!queues[r.branch]) {
          queues[r.branch] = [];
          branchOrder.push(r.branch);
        }
        const pad = String(r.end).length;
        for (let n = r.start; n <= r.end; n += 1) {
          const num = String(n).padStart(pad, "0");
          queues[r.branch].push(r.prefix.trim() ? `${r.prefix.trim()}${num}` : num);
        }
      }
    }

    if (usingPlan) {
      if (roomPlan.length === 0) {
        setError("Upload a room plan first (or switch back to the same layout for every room).");
        return;
      }
    } else if (
      classrooms < 1 ||
      rowsPerRoom < 1 ||
      benchesPerRow < 1 ||
      seatsPerBench < 1 ||
      branchesPerRoom < 1
    ) {
      setError(
        "Classrooms, rows, benches per row, seats per bench, and branches per classroom must all be at least 1."
      );
      return;
    }

    const branchOf = {};
    branchOrder.forEach((b) => queues[b].forEach((id) => { branchOf[id] = b; }));

    const totalStudents = branchOrder.reduce((sum, b) => sum + queues[b].length, 0);
    const capacityPerRoom = rowsPerRoom * benchesPerRow * seatsPerBench;
    const roomCount = usingPlan ? roomPlan.length : classrooms;
    const capacity = usingPlan
      ? roomPlan.reduce((sum, r) => sum + r.rows * r.benches * r.seats, 0)
      : classrooms * capacityPerRoom;
    let planWarnings = [];

    let conflicts = 0;

    // Picks the next student for one seat, preferring the branch (from the allowed
    // set for this room) with the most students left that hasn't already been seated
    // on this particular bench. Falls back to repeating a branch only if forced to.
    function pickNext(usedThisBench, allowedBranches) {
      let candidates = allowedBranches.filter(
        (b) => queues[b].length > 0 && !usedThisBench.has(b)
      );
      let forced = false;
      if (candidates.length === 0) {
        candidates = allowedBranches.filter((b) => queues[b].length > 0);
        forced = candidates.length > 0;
      }
      if (candidates.length === 0) return null;

      candidates.sort((a, b) => queues[b].length - queues[a].length);
      const branch = candidates[0];
      const id = queues[branch].shift();
      if (forced) conflicts += 1;
      usedThisBench.add(branch);
      return { id, branch };
    }

    function emptyRows() {
      const rows = [];
      for (let r = 0; r < rowsPerRoom; r += 1) {
        const benches = [];
        for (let b = 0; b < benchesPerRow; b += 1) {
          benches.push(Array.from({ length: seatsPerBench }, () => null));
        }
        rows.push(benches);
      }
      return rows;
    }

    // Fills one room starting from `initialActive` (the branches meant to open this
    // room). If those branches run dry partway through, it grows the active set by
    // pulling in the next branch (in branchOrder) that still has students, so a
    // room never leaves benches empty while ANY branch still has students waiting —
    // it just keeps cascading through branches instead of stopping at its original set.
    // `usedBranchesSet` accumulates every branch that actually ends up seated in this
    // room, so the room's legend/label reflects what's really inside it.
    function fillRoom(initialActive, usedBranchesSet) {
      const active = [...initialActive];

      function ensureActive() {
        while (!active.some((b) => queues[b].length > 0)) {
          const next = branchOrder.find(
            (b) => queues[b].length > 0 && !active.includes(b)
          );
          if (!next) return false; // no students left anywhere
          active.push(next);
          usedBranchesSet.add(next);
        }
        return true;
      }

      const rows = [];
      for (let r = 0; r < rowsPerRoom; r += 1) {
        const benches = [];
        for (let b = 0; b < benchesPerRow; b += 1) {
          const usedThisBench = new Set();
          const seats = [];
          for (let s = 0; s < seatsPerBench; s += 1) {
            if (!ensureActive()) {
              seats.push(null);
              continue;
            }
            if (avoidSameBranch) {
              const picked = pickNext(usedThisBench, active);
              seats.push(picked ? picked.id : null);
            } else {
              const branch = active.find((b2) => queues[b2].length > 0);
              seats.push(branch ? queues[branch].shift() : null);
            }
          }
          benches.push(seats);
        }
        rows.push(benches);
      }
      return rows;
    }

    // Decides which branches "open" each room, in order. Automatic mode groups
    // branches into chunks of `branchesPerRoom`; custom mode uses the groups the
    // person typed in (any branch they didn't mention gets its own trailing group,
    // so nothing is ever left out). Either way, fillRoom's cascading above still
    // guarantees no room is left with empty seats while students remain elsewhere.
    function computeStartingGroups() {
      if (branchGroupMode === "custom" && customGroupsText.trim()) {
        const lookup = {};
        branchOrder.forEach((b) => {
          lookup[b.toLowerCase()] = b;
        });
        const covered = new Set();
        const groups = [];
        const rawGroups = customGroupsText.split(";").map((g) => g.trim()).filter(Boolean);
        for (const rawGroup of rawGroups) {
          const names = rawGroup.split(",").map((n) => n.trim()).filter(Boolean);
          const resolved = [];
          for (const n of names) {
            const match = lookup[n.toLowerCase()];
            if (match && !covered.has(match)) {
              resolved.push(match);
              covered.add(match);
            }
          }
          if (resolved.length > 0) groups.push(resolved);
        }
        for (const b of branchOrder) {
          if (!covered.has(b)) groups.push([b]);
        }
        return groups;
      }

      const groups = [];
      for (let i = 0; i < branchOrder.length; i += branchesPerRoom) {
        groups.push(branchOrder.slice(i, i + branchesPerRoom));
      }
      return groups;
    }

    let rooms = [];
    if (usingPlan) {
      // Room plan uploaded: every room has its own layout and its own branch rule.
      const planned = seatByRoomPlan({ plan: roomPlan, queues, branchOrder, avoidSameBranch });
      rooms = planned.rooms;
      conflicts = planned.conflicts;
      planWarnings = planned.warnings;
    } else {
      const sameLayout = { rows: rowsPerRoom, benches: benchesPerRow, seats: seatsPerBench };
      const startingGroups = computeStartingGroups();
      for (const group of startingGroups) {
        while (rooms.length < classrooms && group.some((b) => queues[b].length > 0)) {
          const usedBranchesSet = new Set(group.filter((b) => queues[b].length > 0));
          const rows = fillRoom(group, usedBranchesSet);
          rooms.push({ allowedBranches: [...usedBranchesSet], layout: sameLayout, rows });
        }
        if (rooms.length >= classrooms) break;
      }
      while (rooms.length < classrooms) {
        rooms.push({ allowedBranches: [], layout: sameLayout, rows: emptyRows() });
      }
    }

    const unassigned = branchOrder.flatMap((b) => queues[b]);

    const names =
      teacherInputMode === "upload"
        ? uploadedTeachers
        : teacherText
            .split(/[\n,]/)
            .map((t) => t.trim())
            .filter(Boolean);
    const shuffled = shuffle(names);
    let teacherShort = false;
    const roomTeachers = [];
    if (names.length > 0) {
      let cursor = 0;
      for (let c = 0; c < roomCount; c += 1) {
        const group = [];
        for (let t = 0; t < teachersPerRoom; t += 1) {
          if (cursor >= shuffled.length) {
            teacherShort = true;
            cursor = 0;
          }
          group.push(shuffled[cursor]);
          cursor += 1;
        }
        roomTeachers.push(group);
      }
    } else {
      for (let c = 0; c < roomCount; c += 1) roomTeachers.push([]);
    }

    setResult({
      rooms,
      totalStudents,
      capacity,
      unassigned,
      conflicts: avoidSameBranch ? conflicts : 0,
      roomTeachers,
      teacherShort,
      studentDetails,
      branchOf,
      branchOrder: [...branchOrder],
      usingPlan,
      planWarnings,
    });
    setActiveRoom(0);
  }

  const occupiedCount = useMemo(() => {
    if (!result) return 0;
    return result.totalStudents - result.unassigned.length;
  }, [result]);

  const customRoomNumbers = roomNumbersText
    .split(",")
    .map((n) => n.trim())
    .filter(Boolean);
  // Heading for a room's tab / title: "Room 101" for a room named in the plan,
  // otherwise "Room 01", "Room 02"...
  function roomTitle(room, i) {
    const n = room && room.name;
    if (!n) return `Room ${String(i + 1).padStart(2, "0")}`;
    return /^(room|hall|lab|class)/i.test(n) ? n : `Room ${n}`;
  }
  // The bare number shown after "Room No :" on the printed sheet. Numbers typed in the
  // "Room numbers" box win; otherwise the name from the room plan; otherwise 1, 2, 3...
  function roomLabel(i) {
    if (customRoomNumbers[i]) return customRoomNumbers[i];
    const n = result && result.rooms[i] && result.rooms[i].name;
    return n ? n.replace(/^room\s*[-:#]?\s*/i, "") : String(i + 1);
  }

  return (
    <div className="esr-wrap">
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Bricolage+Grotesque:wght@600;700;800&family=Figtree:wght@400;500;600;700&display=swap');

        .esr-wrap {
          --ink: #10233a;
          --ink-2: #1c3a5a;
          --paper: #eef1f4;
          --sheet: #ffffff;
          --line: #d5dbe3;
          --muted: #5b6b7d;
          --chalk: #f4c430;
          --teal: #0f766e;
          --teal-soft: #e3f3f1;
          --warn: #b45309;
          --warn-soft: #fef3e2;
          --danger: #b42318;
          --danger-soft: #fdecea;
          font-family: 'Figtree', system-ui, sans-serif;
          background: var(--paper);
          color: var(--ink);
          min-height: 100%;
          padding: 0 0 64px;
          line-height: 1.45;
        }
        .esr-wrap * { box-sizing: border-box; }

        /* ---------- header band ---------- */
        .esr-plate {
          background:
            radial-gradient(circle, rgba(16,35,58,.09) 1.4px, transparent 1.6px) 0 0 / 22px 22px,
            linear-gradient(120deg, #e6f0fa 0%, #f4f8fc 55%, #fff7dc 100%);
          padding: 36px 24px 66px;
          border-bottom: 4px solid var(--chalk);
          text-align: left;
        }
        .esr-plate-inner {
          max-width: 1180px;
          margin: 0 auto;
          display: flex;
          align-items: center;
          justify-content: space-between;
          gap: 28px;
        }
        .esr-plate-text { flex: 1; min-width: 0; text-align: left; }
        .esr-plate .eyebrow {
          display: inline-block;
          font-size: 13px;
          font-weight: 700;
          color: var(--ink);
          background: var(--chalk);
          padding: 4px 12px;
          border-radius: 999px;
          margin-bottom: 12px;
        }
        .esr-plate h1 {
          font-family: 'Bricolage Grotesque', sans-serif;
          font-weight: 800;
          font-size: clamp(32px, 5vw, 52px);
          line-height: 1.05;
          letter-spacing: -.025em;
          color: var(--ink);
          margin: 0 0 10px;
        }
        .esr-plate p { margin: 0; color: var(--muted); font-size: 15px; text-align: left; }
        .esr-plate-art { width: 270px; height: auto; flex-shrink: 0; }
        @media (max-width: 720px) {
          .esr-plate-art { display: none; }
          .esr-plate { padding: 28px 16px 56px; }
        }

        /* ---------- layout ---------- */
        .esr-grid {
          max-width: 1180px;
          margin: -36px auto 0;
          padding: 0 20px;
          display: grid;
          grid-template-columns: 360px 1fr;
          gap: 24px;
          align-items: start;
        }
        @media (max-width: 880px) { .esr-grid { grid-template-columns: 1fr; } }
        .esr-left { display: flex; flex-direction: column; gap: 18px; }

        .esr-panel, .esr-output {
          background: var(--sheet);
          border: 1px solid var(--line);
          border-radius: 14px;
          box-shadow: 0 1px 0 rgba(16,35,58,.04), 0 14px 30px -18px rgba(16,35,58,.35);
        }
        .esr-panel { padding: 22px; }
        .esr-panel h2 {
          font-family: 'Bricolage Grotesque', sans-serif;
          font-size: 20px;
          font-weight: 700;
          margin: 0 0 16px;
          padding-bottom: 12px;
          border-bottom: 2px solid var(--ink);
        }

        /* ---------- controls ---------- */
        .esr-mode-toggle {
          display: grid;
          grid-template-columns: 1fr 1fr;
          border: 1.5px solid var(--ink);
          border-radius: 10px;
          overflow: hidden;
          margin-bottom: 18px;
        }
        .esr-mode-btn {
          background: #fff;
          border: none;
          padding: 10px 8px;
          font: 600 13px 'Figtree', sans-serif;
          color: var(--ink);
          cursor: pointer;
        }
        .esr-mode-btn + .esr-mode-btn { border-left: 1.5px solid var(--ink); }
        .esr-mode-btn:hover { background: var(--paper); }
        .esr-mode-btn.active { background: var(--ink); color: #fff; }

        .esr-range-row {
          display: grid;
          grid-template-columns: 1fr 68px 68px 30px;
          gap: 7px;
          margin-bottom: 9px;
          align-items: center;
        }
        .esr-range-row input, .esr-num-field input, .esr-teacher-area {
          width: 100%;
          font: 500 14px 'Figtree', sans-serif;
          color: var(--ink);
          background: #fff;
          border: 1.5px solid var(--line);
          border-radius: 8px;
        }
        .esr-range-row input { padding: 9px 8px; }
        .esr-num-field input { padding: 9px; }
        .esr-range-row input:focus, .esr-num-field input:focus, .esr-teacher-area:focus,
        .esr-mode-btn:focus-visible, .esr-generate:focus-visible, .esr-tab:focus-visible,
        .esr-add-link:focus-visible, .esr-print-btn:focus-visible {
          outline: 3px solid var(--chalk);
          outline-offset: 1px;
          border-color: var(--ink);
        }
        .esr-remove-btn {
          background: var(--danger-soft);
          border: none;
          color: var(--danger);
          border-radius: 8px;
          height: 34px;
          cursor: pointer;
          font-size: 18px;
          line-height: 1;
        }
        .esr-remove-btn:hover { background: #f9d4cf; }

        .esr-add-link, .esr-print-btn {
          background: #fff;
          border: 1.5px solid var(--ink);
          color: var(--ink);
          font: 600 13px 'Figtree', sans-serif;
          cursor: pointer;
          padding: 8px 12px;
          margin-top: 4px;
          border-radius: 8px;
        }
        .esr-add-link:hover, .esr-print-btn:hover { background: var(--ink); color: #fff; }
        .esr-print-btn { margin-top: 0; }

        .esr-field-label { display: block; font-size: 14px; font-weight: 700; margin: 22px 0 8px; }
        .esr-num-row { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 8px; }
        .esr-num-field label { display: block; font-size: 12px; font-weight: 600; color: var(--muted); margin-bottom: 4px; }

        .esr-checkbox-row {
          display: flex; align-items: flex-start; gap: 9px;
          margin-top: 16px; font-size: 13px; color: var(--ink); cursor: pointer;
        }
        .esr-checkbox-row input { accent-color: var(--teal); width: 16px; height: 16px; margin-top: 2px; }

        .esr-teacher-area { padding: 10px; resize: vertical; min-height: 110px; }
        .esr-teacher-hint { font-size: 12px; color: var(--muted); margin: 8px 0 0; line-height: 1.5; }

        .esr-file-drop {
          border: 2px dashed var(--ink-2);
          border-radius: 12px;
          padding: 26px 14px;
          text-align: center;
          font-size: 13px;
          font-weight: 600;
          color: var(--ink);
          background: repeating-linear-gradient(45deg, #fff, #fff 10px, #f6f8fa 10px, #f6f8fa 20px);
          position: relative;
        }
        .esr-file-drop:hover { border-color: var(--teal); background: var(--teal-soft); }
        .esr-file-drop input[type="file"] { position: absolute; inset: 0; opacity: 0; cursor: pointer; }
        .esr-file-name { margin-top: 8px; font-size: 12px; font-weight: 500; color: var(--teal); }

        .esr-upload-summary {
          margin-top: 10px;
          padding: 10px 12px;
          border-radius: 8px;
          border-left: 4px solid var(--teal);
          background: var(--teal-soft);
          color: #0b4f4a;
          font-size: 12px;
        }
        .esr-upload-summary.is-warn { border-left-color: var(--warn); background: var(--warn-soft); color: #7a3a05; }
        .esr-file-list { margin-top: 6px; max-height: 84px; overflow-y: auto; word-break: break-word; opacity: .85; }

        .esr-generate {
          width: 100%;
          background: var(--chalk);
          color: var(--ink);
          border: 2px solid var(--ink);
          border-radius: 12px;
          padding: 15px;
          font: 800 16px 'Bricolage Grotesque', sans-serif;
          cursor: pointer;
          box-shadow: 4px 4px 0 var(--ink);
          transition: transform .12s, box-shadow .12s;
        }
        .esr-generate:hover { transform: translate(2px, 2px); box-shadow: 2px 2px 0 var(--ink); }
        .esr-generate:active { transform: translate(4px, 4px); box-shadow: 0 0 0 var(--ink); }

        .esr-error {
          margin-top: 12px; padding: 10px 12px; border-radius: 8px;
          font-size: 13px; color: var(--danger); background: var(--danger-soft);
          border-left: 4px solid var(--danger);
        }

        /* ---------- output ---------- */
        .esr-output { min-height: 520px; overflow: hidden; }
        .esr-empty-state { text-align: center; color: var(--muted); font-size: 14px; line-height: 1.7; padding: 90px 32px; }
        .esr-empty-state::before {
          content: "";
          display: block;
          width: 132px; height: 68px;
          margin: 0 auto 22px;
          background:
            linear-gradient(var(--ink) 0 0) 0 0 / 100% 3px no-repeat,
            radial-gradient(circle, var(--ink-2) 5px, transparent 5.5px) 0 12px / 22px 22px;
          opacity: .8;
        }

        .esr-summary {
          display: grid;
          grid-template-columns: repeat(5, minmax(90px, 1fr));
          border-bottom: 1px solid var(--line);
        }
        @media (max-width: 700px) { .esr-summary { grid-template-columns: repeat(2, 1fr); } }
        .esr-stat { padding: 16px 18px; border-right: 1px solid var(--line); }
        .esr-stat:last-child { border-right: none; }
        .esr-stat .num { font: 800 30px 'Bricolage Grotesque', sans-serif; color: var(--ink); line-height: 1; }
        .esr-stat .label { font-size: 12px; color: var(--muted); margin-top: 5px; font-weight: 500; }
        .esr-stat.warn { background: var(--warn-soft); }
        .esr-stat.warn .num { color: var(--warn); }

        .esr-tabs { display: flex; gap: 6px; overflow-x: auto; padding: 12px 18px 0; background: var(--paper); border-bottom: 1px solid var(--line); }
        .esr-tab {
          font: 600 13px 'Figtree', sans-serif;
          padding: 9px 16px;
          background: #e2e7ed;
          border: 1px solid var(--line);
          border-bottom: none;
          border-radius: 8px 8px 0 0;
          cursor: pointer;
          color: var(--muted);
          white-space: nowrap;
        }
        .esr-tab:hover { color: var(--ink); background: #edf0f4; }
        .esr-tab.active { background: #fff; color: var(--ink); box-shadow: inset 0 3px 0 var(--chalk); margin-bottom: -1px; }

        .esr-room-body { padding: 22px; }
        .esr-room-head { display: flex; align-items: center; flex-wrap: wrap; gap: 10px 14px; margin-bottom: 14px; }
        .esr-room-head h3 { font: 800 26px 'Bricolage Grotesque', sans-serif; margin: 0; margin-right: auto; }
        .esr-invigilator { padding: 8px 12px; border-radius: 8px; background: var(--paper); color: var(--muted); font-size: 13px; }
        .esr-invigilator .name { color: var(--ink); font-weight: 700; }

        .esr-legend { display: flex; flex-wrap: wrap; gap: 8px; margin-bottom: 14px; }
        .esr-chip {
          display: inline-flex; align-items: center; gap: 7px;
          padding: 5px 11px; border: 1.5px solid var(--line); border-radius: 999px;
          font-size: 13px; font-weight: 600;
        }
        .esr-chip i { width: 11px; height: 11px; border-radius: 50%; display: inline-block; }

        .esr-layout-note {
          margin: 0 0 16px; padding: 10px 14px;
          border-left: 4px solid var(--chalk); background: #fffbe8;
          font-size: 13px; color: #4a4320;
        }

        .esr-row-block { margin-bottom: 16px; border: 1px solid var(--line); border-radius: 10px; overflow: hidden; background: #fff; }
        .esr-row-title { padding: 9px 14px; background: var(--ink); color: #fff; font: 700 14px 'Bricolage Grotesque', sans-serif; }
        .esr-bench-grid {
          display: grid;
          grid-template-columns: repeat(var(--bench-count), minmax(96px, 1fr));
          gap: 10px; padding: 12px; overflow-x: auto;
        }
        .esr-bench-card {
          min-width: 96px; padding: 8px;
          border: 1px solid var(--line); border-bottom: 4px solid #c3ccd6;
          border-radius: 8px; background: #f6f8fa; text-align: center;
        }
        .esr-bench-title { font-size: 11px; font-weight: 700; color: var(--muted); margin-bottom: 6px; }
        .esr-seat {
          padding: 6px 4px; margin-top: 5px; border-radius: 6px;
          background: #fff; border: 1px solid var(--line);
          border-left: 5px solid var(--seat-accent, var(--line));
          font-size: 12px; font-weight: 700; color: var(--ink); word-break: break-word;
        }
        .esr-seat.empty { color: #9aa7b6; background: transparent; border-style: dashed; font-weight: 500; }
        .esr-seat-name { display: block; font-size: 10px; font-weight: 500; color: var(--muted); margin-top: 2px; }

        .esr-note { margin: 0 22px 20px; padding: 13px 15px; border-radius: 8px; border-left: 4px solid var(--danger); background: var(--danger-soft); font-size: 13px; }
        .esr-note .heading { font-weight: 700; margin-bottom: 4px; }
        .esr-note .ids { color: #5a3a37; line-height: 1.55; word-break: break-word; }

        @media (max-width: 600px) {
          .esr-grid { padding: 0 12px; }
          .esr-panel { padding: 16px; }
          .esr-range-row { grid-template-columns: 1fr 58px 58px 30px; }
          .esr-room-body { padding: 14px; }
        }

        /* ---------- room plan upload ---------- */
        .esr-plan-table { width: 100%; margin-top: 10px; border-collapse: collapse; font-size: 12px; }
        .esr-plan-table th, .esr-plan-table td {
          text-align: left; padding: 4px 6px; border-bottom: 1px solid rgba(15,118,110,.25);
        }
        .esr-plan-table th { font-weight: 700; }
        .esr-rule-list { margin: 4px 0 0; padding-left: 18px; }
        .esr-rule-list li { margin-bottom: 3px; }
        .esr-rule-list code {
          background: var(--paper); border: 1px solid var(--line); border-radius: 4px;
          padding: 0 4px; font-size: 11.5px; color: var(--ink);
        }

        /* ---------- printed exam-notice sheet ---------- */
        .esr-sheet {
          background: #fff; color: #000;
          border: 1px solid var(--line); border-radius: 6px;
          padding: 18px 18px 22px; overflow-x: auto;
        }
        .esr-sheet-head {
          display: flex; justify-content: space-between; align-items: baseline; gap: 12px;
          font: 700 14px 'Figtree', sans-serif; margin-bottom: 6px;
        }
        .esr-sheet-subhead {
          display: grid; grid-template-columns: 1fr auto 1fr; gap: 12px;
          font: 700 14px 'Figtree', sans-serif; margin-bottom: 12px;
        }
        .esr-sheet-subhead span:last-child { text-align: right; }
        .esr-sheet-blocks { display: flex; gap: 18px; align-items: flex-start; }
        table.esr-sheet-table { border-collapse: collapse; flex: 1 1 0; min-width: 0; }
        .esr-sheet-table th {
          border: 1.5px solid #000; padding: 8px 6px; text-align: center;
          font: 700 14px 'Figtree', sans-serif; white-space: nowrap;
        }
        .esr-sheet-table th.esr-sheet-corner { border: none; width: 22px; }
        .esr-sheet-table td {
          border-left: 1.5px solid #000; border-right: 1.5px solid #000;
          padding: 9px 6px; text-align: center; white-space: nowrap;
          font: 400 12.5px Cambria, Georgia, 'Times New Roman', serif;
        }
        .esr-sheet-table tbody tr:last-child td { border-bottom: 1.5px solid #000; }
        .esr-sheet-table td.esr-sheet-idx {
          border: none; width: 22px; padding: 9px 4px 9px 0; text-align: right;
          font: 400 11px 'Figtree', sans-serif; color: #000;
        }

        @media print {
          .esr-wrap { background: #fff; padding: 0; }
          .esr-plate, .esr-left, .esr-tabs, .esr-print-btn, .esr-summary,
          .esr-legend, .esr-layout-note, .esr-note,
          .esr-room-head .esr-mode-toggle { display: none !important; }
          .esr-grid { display: block; margin: 0; padding: 0; }
          .esr-output { border: none; box-shadow: none; }
          .esr-room-body { padding: 0; }
          .esr-sheet { border: none; padding: 0; overflow: visible; }
          .esr-sheet-blocks { gap: 14px; }
          .esr-sheet-table { break-inside: avoid; }
          .esr-row-block { break-inside: avoid; }
          .esr-row-title { background: #e2e7ed; color: #000; }
        }
        
/* ===== PROFESSIONAL INTERFACE UPGRADE ===== */

.esr-wrap {
  --ink: #142b4a;
  --ink-2: #24476d;
  --paper: #f3f6fb;
  --sheet: #ffffff;
  --line: #dce4ef;
  --muted: #64748b;
  --chalk: #f4c430;
  --teal: #087f8c;
  --teal-soft: #e5f6f6;

  background: #f3f6fb;
  color: #142b4a;
  font-family: "Figtree", Arial, sans-serif;
  line-height: 1.6;
}

/* Modern professional header */
.esr-plate {
  background:
    radial-gradient(
      circle,
      rgba(255, 255, 255, 0.12) 1px,
      transparent 1.6px
    ) 0 0 / 24px 24px,
    linear-gradient(125deg, #102746 0%, #1c456e 65%, #087f8c 100%);
  border-bottom: 4px solid #f4c430;
  padding: 54px 28px 76px;
  text-align: left;
}

.esr-plate .eyebrow {
  background: #f4c430;
  color: #142b4a;
  padding: 6px 14px;
  border-radius: 6px;
  font-size: 12px;
  font-weight: 800;
  letter-spacing: 0.7px;
  text-transform: uppercase;
  margin-bottom: 18px;
}

.esr-plate h1 {
  color: #ffffff;
  font-size: clamp(32px, 4vw, 48px);
  font-weight: 800;
  line-height: 1.12;
  letter-spacing: -1px;
  margin: 0 0 14px;
}

.esr-plate p {
  color: #dbeafe;
  font-size: 15px;
  line-height: 1.8;
}

.esr-plate-art {
  width: 235px;
  filter: drop-shadow(0 12px 22px rgba(0, 0, 0, 0.12));
}

/* Main page layout */
.esr-grid {
  max-width: 1380px;
  grid-template-columns: minmax(300px, 360px) minmax(0, 1fr);
  gap: 24px;
  padding: 0 24px;
  margin: -36px auto 0;
  align-items: start;
}

.esr-left {
  gap: 20px;
}

/* Clean cards */
.esr-panel,
.esr-output {
  background: #ffffff;
  border: 1px solid #dce4ef;
  border-radius: 16px;
  box-shadow: 0 8px 28px rgba(20, 43, 74, 0.06);
}

.esr-panel {
  padding: 24px;
}

.esr-panel h2 {
  color: #142b4a;
  font-size: 19px;
  border-bottom: 1px solid #e2e8f0;
  padding-bottom: 14px;
  margin-bottom: 18px;
}

/* Inputs and text areas */
.esr-range-row input,
.esr-num-field input,
.esr-teacher-area {
  border: 1px solid #cbd5e1;
  border-radius: 9px;
  transition: border-color 0.2s, box-shadow 0.2s;
}

.esr-range-row input:hover,
.esr-num-field input:hover,
.esr-teacher-area:hover {
  border-color: #94a3b8;
}

.esr-range-row input:focus,
.esr-num-field input:focus,
.esr-teacher-area:focus {
  border-color: #087f8c;
  outline: none;
  box-shadow: 0 0 0 3px rgba(8, 127, 140, 0.12);
}

/* Manual / upload selection */
.esr-mode-toggle {
  background: #f1f5f9;
  border: 1px solid #dce4ef;
  border-radius: 10px;
  padding: 4px;
  gap: 4px;
}

.esr-mode-btn {
  border-radius: 7px;
  padding: 11px 8px;
  transition: background 0.2s, color 0.2s;
}

.esr-mode-btn + .esr-mode-btn {
  border-left: none;
}

.esr-mode-btn:hover {
  background: #e2e8f0;
}

.esr-mode-btn.active {
  background: #142b4a;
  color: #ffffff;
}

/* Primary generate button */
.esr-generate {
  background: #142b4a;
  color: #ffffff;
  border: 1px solid #142b4a;
  border-radius: 10px;
  padding: 14px 18px;
  box-shadow: 0 5px 12px rgba(20, 43, 74, 0.16);
  transition: background 0.2s, transform 0.2s, box-shadow 0.2s;
}

.esr-generate:hover {
  background: #087f8c;
  border-color: #087f8c;
  box-shadow: 0 7px 16px rgba(8, 127, 140, 0.2);
  transform: translateY(-1px);
}

.esr-generate:active {
  transform: translateY(0);
}

/* Secondary buttons */
.esr-add-link,
.esr-print-btn {
  border: 1px solid #cbd5e1;
  border-radius: 9px;
  color: #24476d;
  background: #ffffff;
  transition: all 0.2s;
}

.esr-add-link:hover,
.esr-print-btn:hover {
  background: #142b4a;
  border-color: #142b4a;
  color: #ffffff;
}

/* Upload areas */
.esr-file-drop {
  background: #f8fafc;
  border: 1.5px dashed #94a3b8;
  border-radius: 12px;
  padding: 24px 14px;
  transition: background 0.2s, border-color 0.2s;
}

.esr-file-drop:hover {
  background: #eaf7f7;
  border-color: #087f8c;
}

/* Summary statistics */
.esr-summary {
  background: #ffffff;
  border-bottom: 1px solid #dce4ef;
}

.esr-stat {
  padding: 20px 18px;
  border-right: 1px solid #e2e8f0;
}

.esr-stat .num {
  color: #142b4a;
  font-size: 29px;
}

.esr-stat .label {
  color: #64748b;
  font-size: 12px;
}

/* Room navigation tabs */
.esr-tabs {
  background: #f1f5f9;
  gap: 8px;
  padding: 14px 18px 0;
}

.esr-tab {
  background: #e8eef6;
  color: #475569;
  border: 1px solid #dce4ef;
  border-radius: 8px 8px 0 0;
  padding: 10px 16px;
}

.esr-tab:hover {
  background: #ffffff;
  color: #142b4a;
}

.esr-tab.active {
  background: #ffffff;
  color: #142b4a;
  box-shadow: inset 0 3px 0 #087f8c;
}

/* Room and seating layout */
.esr-room-body {
  padding: 26px;
}

.esr-room-head h3 {
  color: #142b4a;
}

.esr-invigilator {
  background: #eff6ff;
  border: 1px solid #dbeafe;
  border-radius: 9px;
}

.esr-layout-note {
  background: #f8fafc;
  border: 1px solid #e2e8f0;
  border-left: 4px solid #087f8c;
  border-radius: 8px;
  padding: 13px 16px;
  color: #334155;
}

.esr-row-block {
  border: 1px solid #dce4ef;
  border-radius: 12px;
  box-shadow: 0 3px 10px rgba(20, 43, 74, 0.035);
}

.esr-row-title {
  background: #eaf0f7;
  color: #142b4a;
  padding: 11px 15px;
  border-bottom: 1px solid #dce4ef;
}

.esr-bench-card {
  background: #f8fafc;
  border: 1px solid #dce4ef;
  border-bottom: 3px solid #cbd5e1;
  border-radius: 9px;
  transition: border-color 0.2s, box-shadow 0.2s;
}

.esr-bench-card:hover {
  border-bottom-color: #087f8c;
  box-shadow: 0 3px 10px rgba(20, 43, 74, 0.06);
}

.esr-seat {
  background: #ffffff;
  border: 1px solid #e2e8f0;
  border-left: 4px solid var(--seat-accent, #cbd5e1);
  border-radius: 7px;
}

.esr-seat.empty {
  background: #f8fafc;
  border: 1px dashed #cbd5e1;
}

/* Better spacing on phones and tablets */
@media (max-width: 880px) {
  .esr-grid {
    grid-template-columns: minmax(0, 1fr);
    margin-top: -28px;
  }

  .esr-plate {
    padding: 38px 22px 56px;
  }
}

@media (max-width: 600px) {
  .esr-grid {
    padding: 0 12px;
    gap: 16px;
  }

  .esr-panel {
    padding: 18px;
  }

  .esr-room-body {
    padding: 16px;
  }

  .esr-plate h1 {
    font-size: 34px;
  }

  .esr-plate p {
    font-size: 14px;
  }

  .esr-bench-grid {
    grid-template-columns: repeat(
      var(--bench-count),
      minmax(96px, 1fr)
    );
    overflow-x: auto;
  }
}

`}</style>    

      <div className="esr-plate">
        <div className="esr-plate-inner">
          <div className="esr-plate-text">
            <div className="eyebrow">Examination cell</div>
            <h1>Exam seating register</h1>
            <p>Assigns students to classrooms → rows → benches → seats, with invigilators</p>
          </div>
          <svg className="esr-plate-art" viewBox="0 0 270 180" role="img" aria-label="Classroom plan with rows of benches">
            <rect x="2" y="2" width="266" height="176" rx="12" fill="#fff" stroke="#10233a" strokeWidth="2.5" />
            <rect x="40" y="12" width="190" height="13" rx="3" fill="#10233a" />
            <rect x="105" y="32" width="60" height="15" rx="3" fill="#f4c430" stroke="#10233a" strokeWidth="2" />
            <circle cx="135" cy="58" r="6" fill="#10233a" />
            {[0, 1, 2].map((r) =>
              [0, 1, 2, 3].map((c) => {
                const x = 22 + c * 60;
                const y = 78 + r * 32;
                return (
                  <g key={`${r}-${c}`}>
                    <rect x={x} y={y} width="46" height="13" rx="3" fill="#f6f8fa" stroke="#10233a" strokeWidth="2" />
                    <circle cx={x + 12} cy={y + 21} r="5" fill={(r + c) % 2 ? "#0f766e" : "#c2410c"} />
                    <circle cx={x + 34} cy={y + 21} r="5" fill={(r + c) % 2 ? "#c2410c" : "#4338ca"} />
                  </g>
                );
              })
            )}
            <rect x="238" y="140" width="24" height="34" rx="3" fill="none" stroke="#10233a" strokeWidth="2" strokeDasharray="4 3" />
          </svg>
        </div>
      </div>

      <div className="esr-grid">
        <div className="esr-left">
          <div className="esr-panel">
            <h2>Student list</h2>
            <div className="esr-mode-toggle">
              <button
                className={`esr-mode-btn ${inputMode === "manual" ? "active" : ""}`}
                onClick={() => setInputMode("manual")}
              >
                Enter ID ranges
              </button>
              <button
                className={`esr-mode-btn ${inputMode === "upload" ? "active" : ""}`}
                onClick={() => setInputMode("upload")}
              >
                Upload files
              </button>
            </div>

            {inputMode === "manual" && (
              <>
                {ranges.map((r) => (
                  <div className="esr-range-row" key={r.id}>
                    <input
                      type="text"
                      placeholder="Branch / prefix"
                      value={r.prefix}
                      onChange={(e) => updateRange(r.id, "prefix", e.target.value)}
                    />
                    <input
                      type="number"
                      placeholder="Start"
                      value={r.start}
                      onChange={(e) => updateRange(r.id, "start", e.target.value)}
                    />
                    <input
                      type="number"
                      placeholder="End"
                      value={r.end}
                      onChange={(e) => updateRange(r.id, "end", e.target.value)}
                    />
                    <button
                      className="esr-remove-btn"
                      onClick={() => removeRange(r.id)}
                      aria-label="Remove range"
                      title="Remove range"
                    >
                      ×
                    </button>
                  </div>
                ))}
                <button className="esr-add-link" onClick={addRange}>
                  + Add another range
                </button>
              </>
            )}

            {inputMode === "upload" && (
              <>
                <div className="esr-file-drop">
                  <input
                    type="file"
                    accept=".xlsx,.xls,.csv,.docx,.pdf,.pptx"
                    onChange={handleFileUpload}
                    multiple
                  />
                  {parsing
                    ? "Reading selected files…"
                    : "Click or drop multiple .xlsx, .csv, .docx, .pdf, or .pptx files"}
                  {uploadFileNames.length > 0 && !parsing && (
                    <div className="esr-file-name">
                      {uploadFileNames.length} file(s) loaded — click here to add more
                    </div>
                  )}
                </div>

                {uploadError && <div className="esr-error">{uploadError}</div>}

                {uploadedStudents.length > 0 && (
                  <div className="esr-upload-summary">
                    <strong>{uploadFileNames.length}</strong> file(s) loaded ·{" "}
                    <strong>{uploadedStudents.length}</strong> unique student(s)
                    {duplicatesRemoved > 0 && (
                      <>
                        {" "}
                        · <strong>{duplicatesRemoved}</strong> duplicate ID(s) removed
                      </>
                    )}
                    <div className="esr-file-list">{uploadFileNames.join(", ")}</div>
                  </div>
                )}

                {uploadWarnings.map((w, i) => (
                  <div className="esr-upload-summary is-warn" key={i}>
                    ⚠ {w}
                  </div>
                ))}

                {(uploadedStudents.length > 0 || uploadFileNames.length > 0) && (
                  <button
                    type="button"
                    className="esr-add-link"
                    onClick={clearUploadedStudents}
                    style={{ marginTop: "8px" }}
                  >
                    Clear uploaded student files
                  </button>
                )}

                
              </>
            )}

            <label className="esr-field-label">Room configuration</label>
            <div className="esr-mode-toggle">
              <button
                type="button"
                className={`esr-mode-btn ${roomMode === "same" ? "active" : ""}`}
                onClick={() => setRoomMode("same")}
              >
                Same layout for all rooms
              </button>
              <button
                type="button"
                className={`esr-mode-btn ${roomMode === "plan" ? "active" : ""}`}
                onClick={() => setRoomMode("plan")}
              >
                Upload room plan
              </button>
            </div>

            {roomMode === "same" && (
              <>
            <div className="esr-num-row">
              <div className="esr-num-field">
                <label>Classrooms</label>
                <input
                  type="number"
                  min="1"
                  value={classrooms}
                  onChange={(e) => setClassrooms(parseInt(e.target.value, 10) || 0)}
                />
              </div>
              <div className="esr-num-field">
                <label>Rows / classroom</label>
                <input
                  type="number"
                  min="1"
                  value={rowsPerRoom}
                  onChange={(e) => setRowsPerRoom(parseInt(e.target.value, 10) || 0)}
                />
              </div>
              <div className="esr-num-field">
                <label>Benches / row</label>
                <input
                  type="number"
                  min="1"
                  value={benchesPerRow}
                  onChange={(e) => setBenchesPerRow(parseInt(e.target.value, 10) || 0)}
                />
              </div>
            </div>
            <div className="esr-num-row" style={{ marginTop: "8px" }}>
              <div className="esr-num-field">
                <label>Students / bench</label>
                <input
                  type="number"
                  min="1"
                  value={seatsPerBench}
                  onChange={(e) => setSeatsPerBench(parseInt(e.target.value, 10) || 0)}
                />
              </div>
              <div className="esr-num-field">
                <label>Branches / bench</label>
                <input
                  type="number"
                  min="1"
                  value={branchesPerRoom}
                  disabled={branchGroupMode === "custom"}
                  onChange={(e) => setBranchesPerRoom(parseInt(e.target.value, 10) || 0)}
                  style={branchGroupMode === "custom" ? { opacity: 0.5 } : undefined}
                />
              </div>
            </div>

            <label className="esr-field-label">Which branches share a room?</label>
            <div className="esr-mode-toggle" style={{ marginBottom: "10px" }}>
              <button
                type="button"
                className={`esr-mode-btn ${branchGroupMode === "auto" ? "active" : ""}`}
                onClick={() => setBranchGroupMode("auto")}
              >
                Automatic (by count)
              </button>
              <button
                type="button"
                className={`esr-mode-btn ${branchGroupMode === "custom" ? "active" : ""}`}
                onClick={() => setBranchGroupMode("custom")}
              >
                I'll choose
              </button>
            </div>
            {branchGroupMode === "custom" && (
              <>
                <textarea
                  className="esr-teacher-area"
                  style={{ minHeight: "70px" }}
                  value={customGroupsText}
                  onChange={(e) => setCustomGroupsText(e.target.value)}
                  placeholder="CSE, ECE; ISE, AIDS; AIML, EEE"
                />
                <p className="esr-teacher-hint">
                  Semicolons separate groups, commas separate branches within a group —
                  each group fills a room (or several, if it has more students than one
                  room holds). Any branch you don't mention still gets seated on its
                  own afterward, and a room will never sit empty if a later branch still
                  has students left.
                </p>
              </>
            )}
              </>
            )}

            {roomMode === "plan" && (
              <>
                <div className="esr-file-drop">
                  <input
                    type="file"
                    accept=".xlsx,.xls,.csv,.docx,.pdf"
                    onChange={handleRoomPlanUpload}
                    multiple
                  />
                  {roomPlanParsing
                    ? "Reading room plan…"
                    : "Click or drop room plan files (.xlsx, .csv, .docx, .pdf)"}
                  {roomPlanFileNames.length > 0 && !roomPlanParsing && (
                    <div className="esr-file-name">
                      {roomPlanFileNames.length} file(s) loaded — click here to add more
                    </div>
                  )}
                </div>

                {roomPlanError && <div className="esr-error">{roomPlanError}</div>}

                {roomPlan.length > 0 && (
                  <div className="esr-upload-summary">
                    <strong>{roomPlan.length}</strong> room(s) loaded ·{" "}
                    <strong>
                      {roomPlan.reduce((sum, r) => sum + r.rows * r.benches * r.seats, 0)}
                    </strong>{" "}
                    seats in total
                    {roomPlanDuplicatesRemoved > 0 && (
                      <>
                        {" "}
                        · <strong>{roomPlanDuplicatesRemoved}</strong> repeated room(s) skipped
                      </>
                    )}
                    <table className="esr-plan-table">
                      <thead>
                        <tr>
                          <th>Room</th>
                          <th>Layout</th>
                          <th>Who sits here</th>
                        </tr>
                      </thead>
                      <tbody>
                        {roomPlan.map((r) => (
                          <tr key={r.name}>
                            <td>{r.name}</td>
                            <td>
                              {r.rows}×{r.benches}×{r.seats}
                            </td>
                            <td>{r.conditionText}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}

                {roomPlanWarnings.map((w, i) => (
                  <div className="esr-upload-summary is-warn" key={i}>
                    ⚠ {w}
                  </div>
                ))}

                <div style={{ display: "flex", gap: "8px", flexWrap: "wrap", marginTop: "8px" }}>
                  {(roomPlan.length > 0 || roomPlanFileNames.length > 0) && (
                    <button type="button" className="esr-add-link" onClick={clearRoomPlan}>
                      Clear room plan
                    </button>
                  )}
                  <button type="button" className="esr-add-link" onClick={downloadSampleRoomPlan}>
                    Download a sample sheet
                  </button>
                </div>

               
              </>
            )}

            <label className="esr-checkbox-row">
              <input
                type="checkbox"
                checked={avoidSameBranch}
                onChange={(e) => setAvoidSameBranch(e.target.checked)}
              />
              Keep students from the same branch off the same bench
            </label>
          </div>

          <div className="esr-panel">
            <h2>Invigilators</h2>
            <div className="esr-mode-toggle">
              <button
                className={`esr-mode-btn ${teacherInputMode === "manual" ? "active" : ""}`}
                onClick={() => setTeacherInputMode("manual")}
              >
                Type names
              </button>
              <button
                className={`esr-mode-btn ${teacherInputMode === "upload" ? "active" : ""}`}
                onClick={() => setTeacherInputMode("upload")}
              >
                Upload files
              </button>
            </div>

            {teacherInputMode === "manual" && (
              <>
                <textarea
                  className="esr-teacher-area"
                  value={teacherText}
                  onChange={(e) => setTeacherText(e.target.value)}
                  placeholder="One name per line, or comma-separated"
                />
                <p className="esr-teacher-hint">One name per line or comma-separated</p>
              </>
            )}

            {teacherInputMode === "upload" && (
              <>
                <div className="esr-file-drop">
                  <input
                    type="file"
                    accept=".xlsx,.xls,.csv,.docx,.pdf"
                    onChange={handleTeacherFileUpload}
                    multiple
                  />
                  {teacherParsing
                    ? "Reading selected files…"
                    : "Click or drop multiple .xlsx, .csv, .docx, or .pdf files"}
                  {teacherUploadFileNames.length > 0 && !teacherParsing && (
                    <div className="esr-file-name">
                      {teacherUploadFileNames.length} file(s) loaded — click here to add more
                    </div>
                  )}
                </div>

                {teacherUploadError && <div className="esr-error">{teacherUploadError}</div>}

                {uploadedTeachers.length > 0 && (
                  <div className="esr-upload-summary">
                    <strong>{teacherUploadFileNames.length}</strong> file(s) loaded ·{" "}
                    <strong>{uploadedTeachers.length}</strong> unique invigilator name(s)
                    {teacherDuplicatesRemoved > 0 && (
                      <>
                        {" "}
                        · <strong>{teacherDuplicatesRemoved}</strong> duplicate name(s) removed
                      </>
                    )}
                    <div className="esr-file-list">{teacherUploadFileNames.join(", ")}</div>
                  </div>
                )}

                {teacherWarnings.map((w, i) => (
                  <div className="esr-upload-summary is-warn" key={i}>
                    ⚠ {w}
                  </div>
                ))}

                {(uploadedTeachers.length > 0 || teacherUploadFileNames.length > 0) && (
                  <button
                    type="button"
                    className="esr-add-link"
                    onClick={clearUploadedTeachers}
                    style={{ marginTop: "8px" }}
                  >
                    Clear uploaded invigilator files
                  </button>
                )}

               
              </>
            )}

            <label className="esr-field-label">Invigilators per room</label>
            <div className="esr-num-row" style={{ gridTemplateColumns: "1fr" }}>
              <div className="esr-num-field">
                <input
                  type="number"
                  min="1"
                  value={teachersPerRoom}
                  onChange={(e) => setTeachersPerRoom(parseInt(e.target.value, 10) || 0)}
                />
              </div>
            </div>
          </div>

          <div className="esr-panel">
            <h2>Printed sheet details</h2>
            <label className="esr-field-label" style={{ marginTop: 0 }}>
              Title line
            </label>
            <input
              type="text"
              value={noticeTitle}
              onChange={(e) => setNoticeTitle(e.target.value)}
              placeholder="Exam: AY 2026-27 (ODD) | MSE-1 | SEMESTER-5/7"
              style={{
                width: "100%",
                font: "500 13px 'Figtree', sans-serif",
                padding: "9px",
                border: "1.5px solid var(--line)",
                borderRadius: "8px",
              }}
            />
            <div className="esr-num-row" style={{ marginTop: "10px" }}>
              <div className="esr-num-field">
                <label>Date</label>
                <input
                  type="text"
                  value={noticeDate}
                  onChange={(e) => setNoticeDate(e.target.value)}
                  placeholder="23.09.2026"
                />
              </div>
              <div className="esr-num-field">
                <label>Time</label>
                <input
                  type="text"
                  value={noticeTime}
                  onChange={(e) => setNoticeTime(e.target.value)}
                  placeholder="01:30 to 03:00"
                />
              </div>
            </div>
            <label className="esr-field-label">Room numbers (optional)</label>
            <input
              type="text"
              value={roomNumbersText}
              onChange={(e) => setRoomNumbersText(e.target.value)}
              placeholder="166, 167, 168"
              style={{
                width: "100%",
                font: "500 13px 'Figtree', sans-serif",
                padding: "9px",
                border: "1.5px solid var(--line)",
                borderRadius: "8px",
              }}
            />
            <p className="esr-teacher-hint">
              One per room, comma-separated, in tab order. Leave blank to just show
              "Room 1", "Room 2", etc.
            </p>
          </div>

          <button className="esr-generate" onClick={generate}>
            Generate seating chart
          </button>

          {error && <div className="esr-error">{error}</div>}
        </div>

        <div className="esr-output">
          {!result && (
            <div className="esr-empty-state">
              Fill in your student list, room configuration, and invigilators, then
              generate.
              <br />
              Seats fill room by room, bench by bench, keeping branches apart on each
              bench.
            </div>
          )}

          {result && (
            <>
              <div className="esr-summary">
                <div className="esr-stat">
                  <div className="num">{result.totalStudents}</div>
                  <div className="label">Total students</div>
                </div>
                <div className="esr-stat">
                  <div className="num">{result.capacity}</div>
                  <div className="label">Total seats</div>
                </div>
                <div className="esr-stat">
                  <div className="num">{occupiedCount}</div>
                  <div className="label">Seats filled</div>
                </div>
                <div className={`esr-stat ${result.unassigned.length ? "warn" : ""}`}>
                  <div className="num">{result.unassigned.length}</div>
                  <div className="label">Unassigned</div>
                </div>
                {avoidSameBranch && (
                  <div className={`esr-stat ${result.conflicts ? "warn" : ""}`}>
                    <div className="num">{result.conflicts}</div>
                    <div className="label">Branch clashes</div>
                  </div>
                )}
              </div>

              <div className="esr-tabs">
                {result.rooms.map((_, i) => (
                  <button
                    key={i}
                    className={`esr-tab ${activeRoom === i ? "active" : ""}`}
                    onClick={() => setActiveRoom(i)}
                  >
                    {roomTitle(result.rooms[i], i)}
                  </button>
                ))}
              </div>

              <div className="esr-room-body">
                <div className="esr-room-head">
                  <h3>{roomTitle(result.rooms[activeRoom], activeRoom)}</h3>
                  <div className="esr-mode-toggle" style={{ width: "auto", marginBottom: 0 }}>
                    <button
                      type="button"
                      className={`esr-mode-btn ${viewMode === "sheet" ? "active" : ""}`}
                      onClick={() => setViewMode("sheet")}
                    >
                      Printed sheet
                    </button>
                    <button
                      type="button"
                      className={`esr-mode-btn ${viewMode === "cards" ? "active" : ""}`}
                      onClick={() => setViewMode("cards")}
                    >
                      Cards
                    </button>
                  </div>
                  <button type="button" className="esr-print-btn" onClick={() => window.print()}>
                    Print this room
                  </button>
                  <div className="esr-invigilator">
                    Invigilator{result.roomTeachers[activeRoom]?.length > 1 ? "s" : ""}:{" "}
                    <span className="name">
                      {result.roomTeachers[activeRoom]?.length
                        ? result.roomTeachers[activeRoom].join(", ")
                        : "not assigned"}
                    </span>
                  </div>
                </div>
                {result.rooms[activeRoom].allowedBranches.length > 0 && (
                  <div className="esr-legend">
                    {result.rooms[activeRoom].allowedBranches.map((b) => (
                      <span key={b} className="esr-chip">
                        <i style={{ background: branchColor(result.branchOrder, b) }} />
                        {b}
                      </span>
                    ))}
                  </div>
                )}
                {(() => {
                  const lay = result.rooms[activeRoom].layout || {
                    rows: rowsPerRoom,
                    benches: benchesPerRow,
                    seats: seatsPerBench,
                  };
                  return (
                    <div className="esr-layout-note">
                      <strong>Seating layout:</strong> {lay.rows} row(s) × {lay.benches} bench(es) per row × {lay.seats} student(s) per bench
                      = <strong>{lay.rows * lay.benches * lay.seats} seats</strong> in this classroom.
                      {result.usingPlan && result.rooms[activeRoom].conditionText && (
                        <>
                          {" "}
                          <strong>Rule:</strong> {result.rooms[activeRoom].conditionText}.
                        </>
                      )}
                    </div>
                  );
                })()}

                {viewMode === "sheet" ? (
                  <div className="esr-sheet">
                    <div className="esr-sheet-head">
                      <div className="esr-sheet-title">{noticeTitle}</div>
                      <div className="esr-sheet-room">
                        Room No : {roomLabel(activeRoom)}
                      </div>
                    </div>
                    <div className="esr-sheet-subhead">
                      <span />
                      <span>Date: {noticeDate || "—"}</span>
                      <span>Time: {noticeTime || "—"}</span>
                    </div>
                    <div className="esr-sheet-blocks">
                      {result.rooms[activeRoom].rows.map((benches, rIdx) => {
                        // One column header per seat position: whichever branch first
                        // occupies that seat within this block of benches.
                        const seatHeaders = Array.from({ length: benches[0] ? benches[0].length : 0 }).map(
                          (_, sIdx) => {
                            for (const seats of benches) {
                              const id = seats[sIdx];
                              if (id) return result.branchOf?.[id] || "";
                            }
                            return "";
                          }
                        );
                        return (
                          <table className="esr-sheet-table" key={`block-${rIdx}`}>
                            <thead>
                              <tr>
                                <th className="esr-sheet-corner"></th>
                                {seatHeaders.map((h, i) => (
                                  <th key={i}>{h || "—"}</th>
                                ))}
                              </tr>
                            </thead>
                            <tbody>
                              {benches.map((seats, bIdx) => (
                                <tr key={bIdx}>
                                  <td className="esr-sheet-idx">{bIdx + 1}</td>
                                  {seats.map((id, sIdx) => (
                                    <td key={sIdx}>{id || ""}</td>
                                  ))}
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        );
                      })}
                    </div>
                  </div>
                ) : (
                  result.rooms[activeRoom].rows.map((benches, rIdx) => (
                    <div className="esr-row-block" key={`row-${rIdx}`}>
                      <div className="esr-row-title">
                        Row {rIdx + 1} — {benches.length} benches
                      </div>
                      <div
                        className="esr-bench-grid"
                        style={{ "--bench-count": benches.length }}
                      >
                        {benches.map((seats, bIdx) => (
                          <div className="esr-bench-card" key={`row-${rIdx}-bench-${bIdx}`}>
                            <div className="esr-bench-title">Bench {bIdx + 1}</div>
                            {seats.map((id, sIdx) => {
                              const details = id ? result.studentDetails?.[id] : null;
                              const tooltip = details
                                ? [details.name, details.subject].filter(Boolean).join(" • ")
                                : undefined;
                              return (
                                <div
                                  key={sIdx}
                                  className={`esr-seat ${id ? "" : "empty"}`}
                                  style={id ? { "--seat-accent": branchColor(result.branchOrder, result.branchOf?.[id]) } : undefined}
                                  title={tooltip}
                                >
                                  {id || "Empty"}
                                  {details?.name && (
                                    <span className="esr-seat-name">{details.name}</span>
                                  )}
                                </div>
                              );
                            })}
                          </div>
                        ))}
                      </div>
                    </div>
                  ))
                )}
              </div>

              {result.conflicts > 0 && (
                <div className="esr-note">
                  <div className="heading">
                    {result.conflicts} bench seat(s) couldn't avoid a same-branch neighbor
                  </div>
                  <div className="ids">
                    This happens when one branch has more students left than there are
                    other branches to pair them with. Add more branches or spread students
                    across more benches to reduce this.
                  </div>
                </div>
              )}

              {result.teacherShort && (
                <div className="esr-note">
                  <div className="heading">Invigilator list was reused</div>
                  <div className="ids">
                    There weren't enough names for one unique invigilator per room slot,
                    so the list was cycled back through.
                  </div>
                </div>
              )}

              {result.usingPlan && result.planWarnings.length > 0 && (
                <div className="esr-note">
                  <div className="heading">Room plan notes</div>
                  <div className="ids">
                    {result.planWarnings.map((w, i) => (
                      <div key={i}>• {w}</div>
                    ))}
                  </div>
                </div>
              )}

              {result.unassigned.length > 0 && (
                <div className="esr-note">
                  <div className="heading">
                    {result.unassigned.length} student(s) could not be seated —{" "}
                    {result.usingPlan
                      ? "no room with a matching rule had space left"
                      : "capacity was exceeded"}
                  </div>
                  <div className="ids">{result.unassigned.join(", ")}</div>
                </div>
              )}
            </>
          )}
        </div>
      </div>
    </div>
  );
}