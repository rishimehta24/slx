/**
 * Convert incident analysis PDF into an Excel (.xls) workbook.
 * Ported from Python convert_pdf_to_xls.py
 */
import ExcelJS from "exceljs";
import fs from "fs";
import path from "path";
import * as XLSX from "xlsx";

// pdf-parse is CommonJS
const pdfParse = require("pdf-parse") as (buffer: Buffer) => Promise<{ text: string }>;

const DATE_RE = /\d{1,2}\/\d{1,2}\/\d{2,4}/g;
const REPORT_TITLE = "Incident By Incident Type";
const SHEET_NAME = "incident_analysis_report";

/** Use when conversion must fail with a precise, actionable message (file, sheet, row, missing fields). */
export class ConversionError extends Error {
  constructor(
    message: string,
    public readonly fileName: string,
    public readonly sheetName: string = SHEET_NAME,
    public readonly rowNumbers?: number[],
    public readonly missingFields?: string[],
    public readonly looksLikeMergedColumns?: boolean
  ) {
    super(message);
    this.name = "ConversionError";
  }
}

const DEFAULT_COL_WIDTH = 18;
const COLUMN_WIDTHS: Record<number, number> = {
  0: 22,
  1: 20,
  2: 52,
  3: 36,
  4: 32,
  5: 42,
  6: 32,
  7: 32,
  8: 24,
  10: 26,
  11: 36,
  12: 90,
  13: 90,
};

const MERGED_VALUE_COLUMNS: Record<number, number> = {
  8: 9,
  15: 16,
  19: 20,
  30: 32,
  40: 41,
};

const COLUMN_HEADERS: string[] = [
  "Incident #",
  "Incident Type",
  "Resident Name",
  "Resident ID",
  "Admission",
  "Incident Date/Time",
  "Incident Location",
  "Incident Status",
  "Witnessed",
  "",
  "Sent to Hospital",
  "Resident Room Number",
  "Immediate Action Taken",
  "Incident Nursing Description",
  "Abrasion",
  "Bruise",
  "",
  "Burn",
  "Fracture",
  "HIR initiated",
  "",
  "Hematoma",
  "Laceration",
  "None noted at time of incident",
  "Other",
  "Red area only",
  "Reddened Area",
  "Skin Tear",
  "Sprain",
  "Suspected Fracture",
  "Unable to determine",
  "",
  "",
  "Abrasion",
  "Bruise",
  "Burn",
  "Fracture",
  "HIR initiated",
  "Hematoma",
  "Laceration",
  "None noted at time of incident",
  "",
  "Other",
  "Red area only",
  "Reddened Area",
  "Skin Tear",
  "Sprain",
  "Suspected Fracture",
  "Unable to determine",
  "Clutter",
  "Crowding",
  "Furniture",
  "Noise",
  "Other",
  "Pets",
  "Poor Lighting",
  "Rugs/Carpeting",
  "Wet Floor",
  "Anticoagulant Therapy",
  "Antihypertensive medication",
  "Confused",
  "Current UTI",
  "Drowsy",
  "Gait Imbalance",
  "Hypotensive",
  "Impaired Memory",
  "Incontinent",
  "Other",
  "Recent Illness",
  "Recent change in Cognition",
  "Recent change in Medications/New Medications",
  "Sedated",
  "Weakness/Fainted",
  "Active Exit Seeker",
  "Admitted within Last 72h",
  "Ambulating with Assist",
  "Ambulating without Assist",
  "Bed/chair alarm ringing ",
  "Call bell within reach",
  "Chair tilted",
  "Dislikes Roommate",
  "During Transfer",
  "Floor mat in place",
  "Hip protectors in place",
  "Improper Footwear",
  "Incorrect diet texture",
  "Incorrect fluid consistency",
  "Large Groups",
  "Lax/suppository in previous 24 hours",
  "Other",
  "Recent Room Change",
  "Restraint-Seat belt",
  "Restraint-chair prevents rising",
  "Restraint-table top",
  "Scheduled toileting plan",
  "Side rail(s) down",
  "Side rail(s) up",
  "Using Cane",
  "Using Walker",
  "Using Wheeled Walker",
  "Using wheelchair",
  "Wanderer",
];

function buildLookup(start: number, end: number): Map<string, [number, string]> {
  const lookup = new Map<string, [number, string]>();
  for (let idx = start; idx <= Math.min(end, COLUMN_HEADERS.length - 1); idx++) {
    const name = COLUMN_HEADERS[idx].trim();
    if (!name) continue;
    lookup.set(normalizeKey(name), [idx, name]);
  }
  return lookup;
}

const INJURY_DURING_LOOKUP = buildLookup(14, 31);
const ENV_LOOKUP = buildLookup(49, 57);
const PHYS_LOOKUP = buildLookup(58, 74);
const SIT_LOOKUP = buildLookup(75, 101);

export interface IncidentEntry {
  resident_name: string;
  resident_id: string;
  admission: string;
  incident_datetime: string;
  location: string;
  room_number: string;
  incident_datetime_sort: Date | null;
  immediate_action_text: string;
  incident_nursing_description: string;
  incident_number: string;
  incident_status: string;
  incident_type: string;
  witnessed: string;
  sent_to_hospital: string;
  injuries_during: Set<string>;
  injuries_post: Set<string>;
  factors: { environmental: Set<string>; physiological: Set<string>; situational: Set<string> };
}

export interface ReportMetadata {
  date_value: string;
  time_value: string;
  user: string;
  facility: string;
  resident_status_line: string;
  incident_status_line: string;
  reporting_period_line: string;
}

function normalizeKey(value: string): string {
  return value.replace(/\s+/g, " ").trim().toLowerCase();
}

function normalizeBool(value: string): string | null {
  if (!value) return null;
  const key = value.trim().toLowerCase();
  if (key === "y" || key === "yes") return "Y";
  if (key === "n" || key === "no") return "N";
  return null;
}

async function readLines(pdfPath: string): Promise<string[]> {
  const dataBuffer = fs.readFileSync(pdfPath);
  const data = await pdfParse(dataBuffer);
  const text = data.text;
  return text.split(/\r?\n/).map((line: string) => line.replace(/\r$/, ""));
}

function parseMetadata(lines: string[]): ReportMetadata {
  const meta: ReportMetadata = {
    date_value: "",
    time_value: "",
    user: "",
    facility: "",
    resident_status_line: "",
    incident_status_line: "",
    reporting_period_line: "",
  };
  let facilityFound = false;
  for (let i = 0; i < lines.length; i++) {
    const stripped = lines[i].trim();
    if (!stripped) continue;
    if (!meta.date_value && stripped.startsWith("Date:")) {
      meta.date_value = stripped.split("Date:", 1)[1]?.trim() ?? "";
    } else if (!meta.time_value && stripped.startsWith("Time:")) {
      meta.time_value = stripped.split("Time:", 1)[1]?.trim() ?? "";
    } else if (!meta.user && stripped.startsWith("User:")) {
      meta.user = stripped.split("User:", 1)[1]?.trim() ?? "";
    } else if (!facilityFound && meta.user) {
      if (stripped.includes(":") || stripped.includes("Incident") || stripped.startsWith("Page #")) continue;
      meta.facility = stripped;
      facilityFound = true;
    } else if (stripped.startsWith("Resident Status") && !meta.resident_status_line) {
      meta.resident_status_line = stripped
        .replace("Resident Status :", "Resident Status:")
        .replace("Unit :", "Unit:")
        .replace("Floor :", "Floor:");
    } else if (stripped.startsWith("Incident Status") && !meta.incident_status_line) {
      meta.incident_status_line = stripped.replace("Incident Status :", "Incident Status:");
    } else if (stripped.startsWith("Reporting Period") && !meta.reporting_period_line) {
      meta.reporting_period_line = stripped;
    }
    if (meta.reporting_period_line && facilityFound) break;
  }
  return meta;
}

function isEntryStart(text: string): boolean {
  if (!text || !text.includes("(")) return false;
  if (text.startsWith("Total ") || text.startsWith("Page #")) return false;
  return /\d/.test(text);
}

function splitNameId(line: string): [string, string] {
  if (!line.includes("(")) return [line.trim(), ""];
  const name = line.split("(", 1)[0].trim();
  const match = line.match(/\d+/);
  const residentId = match ? match[0] : "";
  return [name, residentId];
}

/** Rest of line after ") " or ")" (e.g. "12/22/20232/3/26 7:15AMResident's RoomWest 216-3"). */
function getRestAfterParen(line: string): string {
  const paren = line.indexOf(")");
  if (paren === -1) return line.trim();
  return line.slice(paren + 1).trim();
}

/**
 * If location contains a trailing room token (e.g. "Common BathroomWest 229-1"), split into
 * location and room so XLS columns align with WORKING format. No space between location and room in PDF is common.
 */
function splitLocationAndRoom(location: string, room: string): { location: string; room: string } {
  if (room || !location) return { location, room };
  const m = location.match(/^(.+?)((?:East|West)\s+[\d\-]+)\s*$/i);
  if (m) {
    return { location: m[1].trim(), room: m[2].trim() };
  }
  return { location, room };
}

/**
 * Parse when admission, incident date/time, location and room are on the same line as name (id).
 * Format: Name (ID)AdmissionDateIncidentDateTime LocationRoom (e.g. West 216-3). No space between dates.
 */
function parseInlineResidentLine(rest: string): { admission: string; incident: string; location: string; room: string } | null {
  const firstDateMatch = rest.match(/^(\d{1,2}\/\d{1,2}\/\d{2,4})/);
  if (!firstDateMatch) return null;
  const admission = firstDateMatch[1].trim();
  const afterFirst = rest.slice(firstDateMatch[0].length);
  const secondMatch = afterFirst.match(/^(\d{1,2}\/\d{1,2}\/\d{2,4}\s*\d{1,2}:\d{2}\s*(?:AM|PM))/i);
  if (!secondMatch) return null;
  const incident = secondMatch[1].trim();
  const afterIncident = afterFirst.slice(secondMatch[0].length).trim();
  const roomMatch = afterIncident.match(/\s+((?:East|West)\s+[\d\-]+)\s*$/i);
  const room = roomMatch ? roomMatch[1].trim() : "";
  const location = roomMatch ? afterIncident.slice(0, roomMatch.index).trim() : afterIncident.trim();
  return { admission, incident, location, room };
}

function nextNonempty(lines: string[], idx: number): [string, number] {
  while (idx < lines.length) {
    const candidate = lines[idx].trim();
    idx++;
    if (candidate) return [candidate, idx];
  }
  return ["", idx];
}

function splitAdmissionIncident(firstLine: string, lines: string[], idx: number): [string, string, number] {
  const matches = firstLine.match(DATE_RE);
  if (!matches || matches.length === 0) return [firstLine.trim(), "", idx];
  const admission = matches[0];
  let incident = "";
  if (matches.length >= 2) {
    const start = firstLine.indexOf(matches[1]);
    incident = firstLine.slice(start).trim();
  }
  if (!incident) {
    [incident, idx] = nextNonempty(lines, idx);
  }
  return [admission, incident, idx];
}

function collectUntilToken(lines: string[], idx: number, token: string): [string[], number] {
  const collected: string[] = [];
  while (idx < lines.length) {
    const stripped = lines[idx].trim();
    idx++;
    if (stripped === token) break;
    if (stripped) collected.push(stripped);
  }
  return [collected, idx];
}

function normalizeDateValue(value: string): string {
  value = value.trim();
  if (!value) return "";
  const formats = ["M/d/yyyy", "M/d/yy", "MM/dd/yyyy", "MM/dd/yy"];
  for (const fmt of formats) {
    const d = parseDate(value, fmt);
    if (d) return formatDate(d);
  }
  return value;
}

function parseDate(value: string, fmt: string): Date | null {
  const parts = value.split(/[\s/]+/);
  if (parts.length < 2) return null;
  const m = parseInt(parts[0], 10);
  const d = parseInt(parts[1], 10);
  const y = parseInt(parts[2] ?? "0", 10);
  const year = y < 100 ? 2000 + y : y;
  if (isNaN(m) || isNaN(d)) return null;
  const date = new Date(year, m - 1, d);
  if (isNaN(date.getTime())) return null;
  return date;
}

function formatDate(d: Date): string {
  const m = d.getMonth() + 1;
  const day = d.getDate();
  const y = d.getFullYear();
  return `${String(m).padStart(2, "0")}/${String(day).padStart(2, "0")}/${y}`;
}

/** Canonical 24h format for XLS output: MM/DD/YYYY HH:MM (e.g. 02/15/2026 15:15). */
function formatDatetime24h(d: Date): string {
  const m = d.getMonth() + 1;
  const day = d.getDate();
  const y = d.getFullYear();
  const h = d.getHours();
  const min = d.getMinutes();
  return `${String(m).padStart(2, "0")}/${String(day).padStart(2, "0")}/${y} ${String(h).padStart(2, "0")}:${String(min).padStart(2, "0")}`;
}

function normalizeDatetimeValue(value: string): [string, Date | null] {
  value = value.trim().replace(/ET/g, "").trim();
  if (!value) return ["", null];
  const patterns = [
    /(\d{1,2})\/(\d{1,2})\/(\d{2,4})\s+(\d{1,2}):(\d{2})\s*(AM|PM)/i,
    /(\d{1,2})\/(\d{1,2})\/(\d{2,4})\s+(\d{1,2}):(\d{2})/,
  ];
  for (const re of patterns) {
    const m = value.match(re);
    if (m) {
      const month = parseInt(m[1], 10);
      const day = parseInt(m[2], 10);
      let year = parseInt(m[3], 10);
      if (year < 100) year += 2000;
      let hour = parseInt(m[4], 10);
      const min = parseInt(m[5], 10);
      if (m[6]?.toUpperCase() === "PM" && hour < 12) hour += 12;
      if (m[6]?.toUpperCase() === "AM" && hour === 12) hour = 0;
      const dt = new Date(year, month - 1, day, hour, min);
      if (!isNaN(dt.getTime())) {
        return [formatDatetime24h(dt), dt];
      }
    }
  }
  return [value, null];
}

function parsePreSection(lines: string[]): [string, string, string, string, string, string[]] {
  const remaining = [...lines];
  const injuries: string[] = [];
  while (remaining.length > 0) {
    const key = normalizeKey(remaining[remaining.length - 1]);
    const entry = INJURY_DURING_LOOKUP.get(key);
    if (entry) {
      injuries.unshift(entry[1]);
      remaining.pop();
    } else break;
  }
  let sentToHospital = "";
  if (remaining.length > 0) {
    const maybe = normalizeBool(remaining[remaining.length - 1]);
    if (maybe) {
      sentToHospital = maybe;
      remaining.pop();
    }
  }
  const room = remaining.length > 0 ? remaining.pop()! : "";
  let witnessed = "";
  if (remaining.length > 0) {
    const maybe = normalizeBool(remaining[remaining.length - 1]);
    if (maybe) {
      witnessed = maybe;
      remaining.pop();
    }
  }
  const location = remaining.length > 0 ? remaining.pop()! : "";
  const immediateText = remaining.join(" ").trim();
  return [immediateText, location, witnessed, room, sentToHospital, injuries];
}

type FactorKey = "environmental" | "physiological" | "situational";

function parseFactors(
  lines: string[],
  idx: number
): [{ environmental: Set<string>; physiological: Set<string>; situational: Set<string> }, number] {
  const factors = {
    environmental: new Set<string>(),
    physiological: new Set<string>(),
    situational: new Set<string>(),
  };
  let current: FactorKey | null = null;
  const mapping: Record<FactorKey, Map<string, [number, string]>> = {
    environmental: ENV_LOOKUP,
    physiological: PHYS_LOOKUP,
    situational: SIT_LOOKUP,
  };
  while (idx < lines.length) {
    const stripped = lines[idx].trim();
    const lowered = stripped.toLowerCase();
    idx++;
    if (!stripped) continue;
    if (stripped.startsWith("Total ") || stripped.startsWith("Privileged and Confidential")) continue;
    if (stripped.startsWith("Date:") || stripped.startsWith("Time:") || stripped.startsWith("User:")) break;
    if (stripped.startsWith("Page #") || stripped.startsWith("Fall Incidents")) continue;
    if (stripped.startsWith("Resident Name") || stripped.startsWith("Admission Date")) continue;
    if (isEntryStart(stripped)) {
      idx--; // next entry line: leave idx so caller will reprocess this line
      break;
    }
    if (lowered.includes("predisposing environmental")) {
      current = "environmental";
      addFactorsFromLine(stripped, current, mapping[current], factors);
      continue;
    }
    if (lowered.includes("predisposing physiological")) {
      current = "physiological";
      addFactorsFromLine(stripped, current, mapping[current], factors);
      continue;
    }
    if (lowered.includes("predisposing situation")) {
      current = "situational";
      addFactorsFromLine(stripped, current, mapping[current], factors);
      continue;
    }
    if (stripped === "Notes") continue;
    if (current) {
      addFactorsFromLine(stripped, current, mapping[current], factors);
    }
  }
  return [factors, idx];
}

function addFactorsFromLine(
  line: string,
  current: FactorKey,
  map: Map<string, [number, string]>,
  factors: { environmental: Set<string>; physiological: Set<string>; situational: Set<string> }
): void {
  const colonIdx = line.indexOf(":");
  const valuePart = colonIdx >= 0 ? line.slice(colonIdx + 1).trim() : line.trim();
  if (!valuePart) return;
  for (const chunk of valuePart.split(",")) {
    const cleaned = chunk.trim();
    if (!cleaned) continue;
    const key = normalizeKey(cleaned);
    const entry = map.get(key);
    if (entry) factors[current].add(entry[1]);
  }
}

function parseEntries(lines: string[]): IncidentEntry[] {
  const entries: IncidentEntry[] = [];
  let idx = 0;
  while (idx < lines.length) {
    const stripped = lines[idx].trim();
    if (!stripped || !isEntryStart(stripped)) {
      idx++;
      continue;
    }
    const [name, residentId] = splitNameId(stripped);
    const rest = getRestAfterParen(stripped);
    const inline = parseInlineResidentLine(rest);

    let admission: string;
    let incident: string;
    let location: string;
    let room: string;
    let immediateText: string;
    let witnessed: string;
    let sentToHospital: string;
    let injuries: string[];

    if (inline) {
      admission = inline.admission;
      incident = inline.incident;
      location = inline.location;
      room = inline.room;
      idx++;
      const [preSection, idxAfterPre] = collectUntilToken(lines, idx, "Nursing Description");
      idx = idxAfterPre;
      [immediateText, , witnessed, , sentToHospital, injuries] = parsePreSection(preSection);
    } else {
      idx++;
      let admissionLine: string;
      [admissionLine, idx] = nextNonempty(lines, idx);
      [admission, incident, idx] = splitAdmissionIncident(admissionLine, lines, idx);
      let preSection: string[];
      [preSection, idx] = collectUntilToken(lines, idx, "Nursing Description");
      [immediateText, location, witnessed, room, sentToHospital, injuries] = parsePreSection(preSection);
    }

    const split = splitLocationAndRoom(location, room);
    location = split.location;
    room = split.room;

    let nursingLines: string[];
    [nursingLines, idx] = collectUntilToken(lines, idx, "Notes");
    const nursingDescription = nursingLines.join(" ").trim();
    const [factors, newIdx] = parseFactors(lines, idx);
    idx = newIdx;
    const [normalizedIncident, incidentDt] = normalizeDatetimeValue(incident);
    entries.push({
      resident_name: name,
      resident_id: residentId,
      admission: normalizeDateValue(admission),
      incident_datetime: normalizedIncident,
      incident_datetime_sort: incidentDt,
      location,
      room_number: room,
      immediate_action_text: nursingDescription,
      incident_nursing_description: immediateText,
      incident_number: "",
      incident_status: "",
      incident_type: "Fall",
      witnessed,
      sent_to_hospital: sentToHospital,
      injuries_during: new Set(injuries),
      injuries_post: new Set(),
      factors,
    });
  }
  entries.sort((a, b) => {
    const da = a.incident_datetime_sort?.getTime() ?? 0;
    const db = b.incident_datetime_sort?.getTime() ?? 0;
    if (db !== da) return db - da;
    return b.resident_name.localeCompare(a.resident_name);
  });
  entries.forEach((e, i) => {
    e.incident_number = String(i + 1);
  });
  return entries;
}

const headerStyle: Partial<ExcelJS.Style> = {
  font: { name: "Arial", size: 11, bold: true },
  alignment: { vertical: "middle", horizontal: "center", wrapText: true },
  fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FF808080" } },
  border: {
    top: { style: "thin" },
    left: { style: "thin" },
    bottom: { style: "thin" },
    right: { style: "thin" },
  },
};

const dataStyle: Partial<ExcelJS.Style> = {
  font: { name: "Arial", size: 10 },
  alignment: { vertical: "top", wrapText: true },
  border: {
    top: { style: "thin" },
    left: { style: "thin" },
    bottom: { style: "thin" },
    right: { style: "thin" },
  },
};

const groupStyle: Partial<ExcelJS.Style> = {
  font: { name: "Arial", size: 10, bold: true },
  alignment: { vertical: "middle", horizontal: "center" },
  fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FFD6E3F0" } },
  border: {
    top: { style: "thin" },
    left: { style: "thin" },
    bottom: { style: "thin" },
    right: { style: "thin" },
  },
};

const metadataStyle: Partial<ExcelJS.Style> = {
  font: { name: "Arial", size: 11, bold: true },
  alignment: { vertical: "middle", horizontal: "left" },
};

const footerStyle: Partial<ExcelJS.Style> = {
  font: { name: "Arial", size: 10, italic: true },
  alignment: { vertical: "middle", horizontal: "left" },
};

function writeWithMerge(
  sheet: ExcelJS.Worksheet,
  row: number,
  col: number,
  value: string,
  style: Partial<ExcelJS.Style>
): void {
  const endCol = MERGED_VALUE_COLUMNS[col] ?? col;
  const cell = sheet.getCell(row + 1, col + 1);
  cell.value = value;
  Object.assign(cell.style, style);
  if (endCol > col) {
    sheet.mergeCells(row + 1, col + 1, row + 1, endCol + 1);
  }
}

function applyColumnWidths(sheet: ExcelJS.Worksheet): void {
  for (let c = 0; c < COLUMN_HEADERS.length; c++) {
    const w = COLUMN_WIDTHS[c] ?? DEFAULT_COL_WIDTH;
    sheet.getColumn(c + 1).width = w;
  }
}

function writeMetadataRows(sheet: ExcelJS.Worksheet, meta: ReportMetadata): void {
  sheet.mergeCells(1, 1, 1, 9);
  sheet.getCell(1, 1).value = `Date:  ${meta.date_value}`;
  Object.assign(sheet.getCell(1, 1).style, metadataStyle);
  sheet.mergeCells(1, 10, 1, 20);
  sheet.getCell(1, 10).value = meta.facility || "Unknown";
  Object.assign(sheet.getCell(1, 10).style, metadataStyle);
  sheet.mergeCells(1, 21, 1, 32);
  sheet.getCell(1, 21).value = "Facility #: ";
  Object.assign(sheet.getCell(1, 21).style, metadataStyle);

  sheet.mergeCells(2, 1, 2, 9);
  sheet.getCell(2, 1).value = `Time: ${meta.time_value}`;
  Object.assign(sheet.getCell(2, 1).style, metadataStyle);
  sheet.mergeCells(2, 10, 2, 20);
  sheet.getCell(2, 10).value = REPORT_TITLE;
  Object.assign(sheet.getCell(2, 10).style, metadataStyle);

  sheet.mergeCells(3, 1, 3, 32);
  sheet.getCell(3, 1).value = `User:  ${meta.user}`;
  Object.assign(sheet.getCell(3, 1).style, metadataStyle);
  sheet.mergeCells(4, 1, 4, 32);
  sheet.getCell(4, 1).value = meta.resident_status_line || "Resident Status:";
  Object.assign(sheet.getCell(4, 1).style, metadataStyle);
  sheet.mergeCells(5, 1, 5, 32);
  sheet.getCell(5, 1).value = meta.incident_status_line || "Incident Status:";
  Object.assign(sheet.getCell(5, 1).style, metadataStyle);
  sheet.mergeCells(6, 1, 6, 32);
  sheet.getCell(6, 1).value = meta.reporting_period_line || "Reporting Period :";
  Object.assign(sheet.getCell(6, 1).style, metadataStyle);

  sheet.mergeCells(7, 1, 7, 14);
  sheet.getCell(7, 1).value = "";
  Object.assign(sheet.getCell(7, 1).style, groupStyle);
  sheet.mergeCells(7, 15, 7, 33);
  sheet.getCell(7, 15).value = "Injury Noted - During";
  Object.assign(sheet.getCell(7, 15).style, groupStyle);
  sheet.mergeCells(7, 34, 7, 49);
  sheet.getCell(7, 34).value = "Injury Noted - Post";
  Object.assign(sheet.getCell(7, 34).style, groupStyle);
  sheet.mergeCells(7, 50, 7, 58);
  sheet.getCell(7, 50).value = "Predisposing Factors (Environmental)";
  Object.assign(sheet.getCell(7, 50).style, groupStyle);
  sheet.mergeCells(7, 59, 7, 73);
  sheet.getCell(7, 59).value = "Predisposing Factors (Physiological)";
  Object.assign(sheet.getCell(7, 59).style, groupStyle);
  sheet.mergeCells(7, 74, 7, 102);
  sheet.getCell(7, 74).value = "Predisposing Factors (Situational)";
  Object.assign(sheet.getCell(7, 74).style, groupStyle);
  sheet.getRow(7).height = 27;
  sheet.getRow(8).height = 32;
  for (let col = 0; col < COLUMN_HEADERS.length; col++) {
    const value = COLUMN_HEADERS[col];
    if (!value) continue;
    const cell = sheet.getCell(8, col + 1);
    cell.value = value;
    Object.assign(cell.style, headerStyle);
    const endCol = MERGED_VALUE_COLUMNS[col] ?? col;
    if (endCol > col) sheet.mergeCells(8, col + 1, 8, endCol + 1);
  }
}

function writeEntries(sheet: ExcelJS.Worksheet, entries: IncidentEntry[]): void {
  const startRow = 8;
  for (let i = 0; i < entries.length; i++) {
    const entry = entries[i];
    const row = startRow + i;
    sheet.getRow(row + 1).height = 32;
    writeWithMerge(sheet, row, 0, entry.incident_number, dataStyle);
    writeWithMerge(sheet, row, 1, entry.incident_type, dataStyle);
    writeWithMerge(sheet, row, 2, entry.resident_name, dataStyle);
    writeWithMerge(sheet, row, 3, entry.resident_id, dataStyle);
    writeWithMerge(sheet, row, 4, entry.admission, dataStyle);
    const incidentDtStr = entry.incident_datetime_sort
      ? formatDatetime24h(entry.incident_datetime_sort)
      : entry.incident_datetime;
    writeWithMerge(sheet, row, 5, incidentDtStr, dataStyle);
    writeWithMerge(sheet, row, 6, entry.location, dataStyle);
    writeWithMerge(sheet, row, 7, entry.incident_status, dataStyle);
    writeWithMerge(sheet, row, 8, entry.witnessed, dataStyle);
    writeWithMerge(sheet, row, 10, entry.sent_to_hospital, dataStyle);
    writeWithMerge(sheet, row, 11, entry.room_number, dataStyle);
    writeWithMerge(sheet, row, 12, entry.immediate_action_text, dataStyle);
    writeWithMerge(sheet, row, 13, entry.incident_nursing_description, dataStyle);
    for (const injury of entry.injuries_during) {
      const key = normalizeKey(injury);
      const entryLookup = INJURY_DURING_LOOKUP.get(key);
      if (entryLookup) writeWithMerge(sheet, row, entryLookup[0], "Y", dataStyle);
    }
    for (const [group, map] of [
      ["environmental", ENV_LOOKUP],
      ["physiological", PHYS_LOOKUP],
      ["situational", SIT_LOOKUP],
    ] as const) {
      for (const factor of entry.factors[group]) {
        const key = normalizeKey(factor);
        const entryLookup = map.get(key);
        if (entryLookup) writeWithMerge(sheet, row, entryLookup[0], "Y", dataStyle);
      }
    }
  }
  const footerRow = startRow + entries.length;
  sheet.mergeCells(footerRow + 1, 1, footerRow + 1, COLUMN_HEADERS.length);
  sheet.getCell(footerRow + 1, 1).value =
    "Privileged and Confidential - Not part of the Medical Record - Do not Copy";
  Object.assign(sheet.getCell(footerRow + 1, 1).style, footerStyle);
}

export async function convert(pdfPath: string, outputPath: string): Promise<void> {
  const fileName = path.basename(pdfPath);
  const lines = await readLines(pdfPath);
  const meta = parseMetadata(lines);
  const entries = parseEntries(lines);

  if (entries.length > 0) {
    const firstRows = entries.slice(0, 5);
    const rowsMissingDatetime: number[] = [];
    let looksLikeMerged = false;
    for (let i = 0; i < firstRows.length; i++) {
      const e = firstRows[i];
      if (!e.incident_datetime && !e.incident_datetime_sort) rowsMissingDatetime.push(9 + i);
      if (!e.incident_status && e.location?.match(/(?:East|West)\s+[\d\-]+/i)) looksLikeMerged = true;
    }
    if (rowsMissingDatetime.length === firstRows.length) {
      throw new ConversionError(
        `Missing incident date/time in data rows (sheet row numbers: ${rowsMissingDatetime.join(", ")}). ` +
          (looksLikeMerged ? "Data may match HTML export / merged columns pattern." : "Check PDF has parseable dates."),
        fileName,
        SHEET_NAME,
        rowsMissingDatetime,
        ["Incident Date/Time"],
        looksLikeMerged
      );
    }
  }

  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet(SHEET_NAME, { views: [{ state: "frozen", ySplit: 8, xSplit: 0 }] });
  applyColumnWidths(sheet);
  writeMetadataRows(sheet, meta);
  writeEntries(sheet, entries);
  const xlsxBuffer = await workbook.xlsx.writeBuffer();
  const wb = XLSX.read(xlsxBuffer, { type: "buffer" });
  const xlsBuffer = XLSX.write(wb, { bookType: "xls", type: "buffer" });
  fs.writeFileSync(outputPath, xlsBuffer);
}
