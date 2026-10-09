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

  function generate() {
    setError("");

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

    if (
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
    const capacity = classrooms * capacityPerRoom;

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

    const startingGroups = computeStartingGroups();
    const rooms = [];
    for (const group of startingGroups) {
      while (rooms.length < classrooms && group.some((b) => queues[b].length > 0)) {
        const usedBranchesSet = new Set(group.filter((b) => queues[b].length > 0));
        const rows = fillRoom(group, usedBranchesSet);
        rooms.push({ allowedBranches: [...usedBranchesSet], rows });
      }
      if (rooms.length >= classrooms) break;
    }
    while (rooms.length < classrooms) {
      rooms.push({ allowedBranches: [], rows: emptyRows() });
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
      for (let c = 0; c < classrooms; c += 1) {
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
      for (let c = 0; c < classrooms; c += 1) roomTeachers.push([]);
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
    });
    setActiveRoom(0);
  }

  const occupiedCount = useMemo(() => {
    if (!result) return 0;
    return result.totalStudents - result.unassigned.length;
  }, [result]);

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

        @media print {
          .esr-wrap { background: #fff; padding: 0; }
          .esr-plate, .esr-left, .esr-tabs, .esr-print-btn, .esr-summary { display: none !important; }
          .esr-grid { display: block; margin: 0; padding: 0; }
          .esr-output { border: none; box-shadow: none; }
          .esr-row-block { break-inside: avoid; }
          .esr-row-title { background: #e2e7ed; color: #000; }
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
                    Room {String(i + 1).padStart(2, "0")}
                  </button>
                ))}
              </div>

              <div className="esr-room-body">
                <div className="esr-room-head">
                  <h3>Room {String(activeRoom + 1).padStart(2, "0")}</h3>
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
                <div className="esr-layout-note">
                  <strong>Seating layout:</strong> {rowsPerRoom} row(s) × {benchesPerRow} bench(es) per row × {seatsPerBench} student(s) per bench
                  = <strong>{rowsPerRoom * benchesPerRow * seatsPerBench} seats</strong> in this classroom.
                </div>

                {result.rooms[activeRoom].rows.map((benches, rIdx) => (
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
                ))}
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

              {result.unassigned.length > 0 && (
                <div className="esr-note">
                  <div className="heading">
                    {result.unassigned.length} student(s) could not be seated — capacity
                    was exceeded
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