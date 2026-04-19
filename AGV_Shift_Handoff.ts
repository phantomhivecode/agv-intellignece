/**
 * AGV Shift Handoff — Office Script for SharePoint Excel
 *
 * Automatically finds the most recent shift tab, analyzes its incidents,
 * and writes a concise problem-statement summary to a "Shift Handoff" tab.
 * Designed to be run at the end of a shift so the incoming engineer has
 * an instant briefing and talking points for the outgoing team.
 *
 * HOW TO USE:
 *   1. In Excel (web), go to Automate > New Script
 *   2. Paste this entire file, replacing the default main() stub
 *   3. Save as "AGV Shift Handoff"
 *   4. Click Run at end of shift — rebuilds the Shift Handoff tab fresh each time
 *
 * WHAT IT PRODUCES:
 *   - Most recent shift auto-detected (highest date + highest shift letter)
 *   - Repeat offenders: same vehicle + same fault ≥2 times this shift
 *   - Hot locations: same location ≥2 incidents this shift
 *   - Worth Asking: auto-generated talking points for the shift discussion
 *   - Full incident log for the shift at the bottom for reference
 */

function main(workbook: ExcelScript.Workbook) {

  // ── 1. Find most recent shift tab ────────────────────────────────────────
  const sheetNameRegex = /^(\d{2})[.,](\d{2})[.,](\d{4})\s*([A-D])$/;
  const sheetNames = workbook.getWorksheets().map(s => s.getName());

  // Sort by year → month → day → shift letter, all descending
  const shiftTabs = sheetNames
    .filter(n => sheetNameRegex.test(n))
    .sort((a, b) => {
      const ma = a.match(sheetNameRegex)!;
      const mb = b.match(sheetNameRegex)!;
      const keyA = `${ma[3]}${ma[1]}${ma[2]}${ma[4]}`;
      const keyB = `${mb[3]}${mb[1]}${mb[2]}${mb[4]}`;
      return keyB.localeCompare(keyA);
    });

  if (shiftTabs.length === 0) {
    console.log("No shift tabs found.");
    return;
  }

  const shiftName = shiftTabs[0];
  const shiftSheet = workbook.getWorksheet(shiftName);
  if (!shiftSheet) { console.log(`Could not open sheet: ${shiftName}`); return; }

  // ── 2. Parse incidents from this shift ───────────────────────────────────
  type Incident = {
    time: string;
    unit: string;
    category: string;
    location: string;
    description: string;
  };

  const unitRegex = /(AGV|AGF|Crane)\s*#?\s*(\d+)/i;
  const incidents: Incident[] = [];

  const usedRange = shiftSheet.getUsedRange();
  if (!usedRange) { console.log("Shift tab is empty."); return; }
  const values = usedRange.getValues();

  for (const row of values) {
    if (!row || row.length < 4) continue;
    const timeStr = normalizeTime(row[1]);
    const area    = row[2] ? String(row[2]).trim() : "";
    const desc    = row[3] ? String(row[3]).trim() : "";
    if (!timeStr || !desc || desc.toLowerCase() === "none") continue;

    const um   = desc.match(unitRegex);
    const unit = um ? `${um[1].toUpperCase()} ${um[2]}` : "Unknown";

    incidents.push({
      time:        timeStr,
      unit,
      category:    categorize(desc),
      location:    resolveLocation(area, desc),
      description: desc
    });
  }

  const total = incidents.length;

  // ── 3. Aggregate ─────────────────────────────────────────────────────────

  // Repeat offenders: same unit + same fault ≥2 times
  const unitFaultMap: Record<string, number> = {};
  for (const i of incidents) {
    if (i.unit === "Unknown" || i.category === "No Description") continue;
    const key = `${i.unit}||${i.category}`;
    unitFaultMap[key] = (unitFaultMap[key] || 0) + 1;
  }
  const repeatOffenders = Object.entries(unitFaultMap)
    .filter(([, c]) => c >= 2)
    .sort(([, a], [, b]) => b - a)
    .slice(0, 8);

  // Hot locations: same location ≥2 incidents
  const locMap: Record<string, number> = {};
  const locFaultMap: Record<string, Record<string, number>> = {};
  const locUnitMap: Record<string, Set<string>> = {};
  for (const i of incidents) {
    if (!i.location || i.location === "Unknown") continue;
    locMap[i.location] = (locMap[i.location] || 0) + 1;
    if (!locFaultMap[i.location]) locFaultMap[i.location] = {};
    locFaultMap[i.location][i.category] = (locFaultMap[i.location][i.category] || 0) + 1;
    if (!locUnitMap[i.location]) locUnitMap[i.location] = new Set();
    locUnitMap[i.location].add(i.unit);
  }
  const hotLocations = Object.entries(locMap)
    .filter(([, c]) => c >= 2)
    .sort(([, a], [, b]) => b - a)
    .slice(0, 8);

  // ── 4. Build talking points ───────────────────────────────────────────────
  const talkingPoints: string[] = [];

  for (const [key, count] of repeatOffenders.slice(0, 4)) {
    const [unit, fault] = key.split("||");
    const locs = incidents
      .filter(i => i.unit === unit && i.category === fault)
      .map(i => i.location);
    const locCounts: Record<string, number> = {};
    for (const l of locs) locCounts[l] = (locCounts[l] || 0) + 1;
    const topLoc = Object.entries(locCounts).sort(([,a],[,b]) => b-a)[0];
    const locNote = topLoc ? ` — mostly at ${topLoc[0]}` : "";
    if (count >= 4) {
      talkingPoints.push(`${unit}: ${count}x "${fault}" this shift${locNote}. Strong case for maintenance inspection.`);
    } else if (count === 3) {
      talkingPoints.push(`${unit}: ${count}x "${fault}" this shift${locNote}. Worth flagging to maintenance.`);
    } else {
      talkingPoints.push(`${unit}: ${count}x "${fault}" this shift${locNote}. Monitor next shift.`);
    }
  }

  for (const [loc, count] of hotLocations.slice(0, 3)) {
    const topFaultEntry = Object.entries(locFaultMap[loc]).sort(([,a],[,b]) => b-a)[0];
    const topFault = topFaultEntry[0];
    const units = Array.from(locUnitMap[loc]).filter(u => u !== "Unknown");
    if (units.length > 1) {
      talkingPoints.push(`${loc}: ${count} incidents (${units.length} different AGVs — likely a location issue). Top fault: "${topFault}".`);
    } else if (units.length === 1) {
      talkingPoints.push(`${loc}: ${count} incidents (all ${units[0]} — likely vehicle issue, not location).`);
    }
  }

  if (talkingPoints.length === 0) {
    talkingPoints.push("Quiet shift — no repeat patterns detected. No major action items.");
  }

  // ── 5. Write Shift Handoff tab ───────────────────────────────────────────
  let out = workbook.getWorksheet("Shift Handoff");
  if (out) out.delete();
  out = workbook.addWorksheet("Shift Handoff");
  out.activate();

  let r = 0;

  // Header
  writeCell(out, r, 0, "AGV Shift Handoff", { bold: true, size: 16 }); r++;
  writeCell(out, r, 0, `Shift: ${shiftName}`, { bold: true, size: 13 }); r++;
  writeCell(out, r, 0, `Generated: ${new Date().toLocaleString()}`, { italic: true, color: "#888888" }); r++;
  writeCell(out, r, 0, `Total incidents this shift: ${total}`, {}); r += 2;

  // Repeat offenders
  writeCell(out, r, 0, "Repeat Offenders", { bold: true, size: 13 }); r++;
  writeCell(out, r, 0, "Same vehicle, same fault, 2+ times this shift", { italic: true, color: "#888888" }); r++;
  if (repeatOffenders.length > 0) {
    writeHeaderRow(out, r, ["Unit", "Fault", "Count", "Locations"]); r++;
    for (const [key, count] of repeatOffenders) {
      const [unit, fault] = key.split("||");
      const locs = incidents
        .filter(i => i.unit === unit && i.category === fault)
        .map(i => i.location);
      const locCounts: Record<string, number> = {};
      for (const l of locs) locCounts[l] = (locCounts[l] || 0) + 1;
      const locStr = Object.entries(locCounts)
        .sort(([,a],[,b]) => b-a)
        .map(([l, c]) => c > 1 ? `${l} x${c}` : l)
        .join(", ");
      out.getRangeByIndexes(r, 0, 1, 4).setValues([[unit, fault, count, locStr]]);
      if (count >= 3) out.getRangeByIndexes(r, 0, 1, 4).getFormat().getFill().setColor("#FAEEDA");
      r++;
    }
  } else {
    writeCell(out, r, 0, "None this shift", { color: "#888888" }); r++;
  }
  r++;

  // Hot locations
  writeCell(out, r, 0, "Hot Locations", { bold: true, size: 13 }); r++;
  writeCell(out, r, 0, "Same location, 2+ incidents this shift", { italic: true, color: "#888888" }); r++;
  if (hotLocations.length > 0) {
    writeHeaderRow(out, r, ["Location", "Incidents", "Top Fault", "AGVs Affected"]); r++;
    for (const [loc, count] of hotLocations) {
      const topFaultEntry = Object.entries(locFaultMap[loc]).sort(([,a],[,b]) => b-a)[0];
      const units = Array.from(locUnitMap[loc]).filter(u => u !== "Unknown").join(", ");
      out.getRangeByIndexes(r, 0, 1, 4).setValues([[loc, count, topFaultEntry[0], units]]);
      if (count >= 3) out.getRangeByIndexes(r, 0, 1, 4).getFormat().getFill().setColor("#FAEEDA");
      r++;
    }
  } else {
    writeCell(out, r, 0, "None this shift", { color: "#888888" }); r++;
  }
  r++;

  // Worth asking
  writeCell(out, r, 0, "Worth Asking — Talking Points", { bold: true, size: 13 }); r++;
  writeCell(out, r, 0, "Auto-generated from shift patterns", { italic: true, color: "#888888" }); r++;
  for (const pt of talkingPoints) {
    writeCell(out, r, 0, `• ${pt}`, {});
    out.getRangeByIndexes(r, 0, 1, 6).merge();
    r++;
  }
  r++;

  // Full incident log
  writeCell(out, r, 0, "Full Incident Log — This Shift", { bold: true, size: 13 }); r++;
  writeHeaderRow(out, r, ["Time", "Unit", "Category", "Location", "Description"]); r++;
  for (const i of incidents) {
    out.getRangeByIndexes(r, 0, 1, 5).setValues([[
      i.time, i.unit, i.category, i.location, i.description
    ]]);
    r++;
  }

  // Column widths
  out.getRange("A:A").getFormat().setColumnWidth(180);
  out.getRange("B:B").getFormat().setColumnWidth(90);
  out.getRange("C:C").getFormat().setColumnWidth(220);
  out.getRange("D:D").getFormat().setColumnWidth(200);
  out.getRange("E:E").getFormat().setColumnWidth(400);

  console.log(`Shift Handoff written for ${shiftName} — ${total} incidents, ${repeatOffenders.length} repeat offenders, ${hotLocations.length} hot locations.`);
}

// ═══════════════════════════════════════════════════════════════════════════
// Helpers — copied from AGV Intelligence so this script runs standalone
// ═══════════════════════════════════════════════════════════════════════════

function normalizeTime(raw: string | number | boolean): string | null {
  if (raw === null || raw === undefined || raw === "") return null;
  if (typeof raw === "number") {
    if (raw >= 0 && raw < 2) {
      const totalMinutes = Math.round(raw * 24 * 60);
      const h = Math.floor(totalMinutes / 60) % 24;
      const m = totalMinutes % 60;
      return `${String(h).padStart(2, "0")}:${String(m).padStart(2, "0")}`;
    }
    return null;
  }
  if (typeof raw === "string") {
    const m = raw.trim().match(/^(\d{1,2}):(\d{2})/);
    if (m) return `${m[1].padStart(2, "0")}:${m[2]}`;
  }
  return null;
}

function categorize(desc: string): string {
  const d = desc.toLowerCase().trim();
  if (/^(agv|agf|crane)\s*#?\s*\d+\s*$/.test(d)) return "No Description";
  if (/conveyor is still transmitting|tec conveyor is still transmitting|tec.*still transmit|conv.*still transmit|conveyor.*still transmit|not sensing conv/.test(d)) return "Conveyor Still Transmitting";
  if (/right handshake timeout|\brhst\b|\brhsto\b|\brht\b|right handshake|rh timeout/.test(d)) return "Handshake - Right Timeout";
  if (/left handshake timeout|\blhst\b|\blht\b|left handshake|lh timeout|left hanshake/.test(d)) return "Handshake - Left Timeout";
  if (/handshake|handhake|hasndshake|hnadshake|handhsake|hansdhake|hanshake/.test(d)) {
    if (/\bleft\b/.test(d)) return "Handshake - Left Timeout";
    if (/\bright\b/.test(d)) return "Handshake - Right Timeout";
    return "Handshake - Unspecified";
  }
  if (/nav.*can.t match reflections|cant match reflections and targets/.test(d)) return "Nav - Can't Match Reflections";
  if (/nav.*no scanner data|nav - no scanner/.test(d)) return "Nav - No Scanner Data";
  if (/guid.*traveled too far|traveled too far without correction/.test(d)) return "Guid - Traveled Too Far";
  if (/nodelock|node lock|lost nav|not on node|node-lock|\bno nav\b|loss of nav/.test(d)) return "Navigation/Nodelock";
  if (/alignment reflector not detected|alignment reflector|alignment sensor did not turn on/.test(d)) return "Alignment Reflector Not Detected";
  if (/no auto-mode guidesafe|no auto mode guidesafe|guidesafe|guidsafe/.test(d)) return "No Auto-Mode Guidesafe";
  if (/completpicktransfer|completepicktransfer|complet.*pick.*transfer.*t/.test(d)) return "CompletePickTransfer Timeout";
  if (/competedrop|completedrop.*transfer|complet.*drop.*transfer.*t/.test(d)) return "CompleteDropTransfer Timeout";
  if (/startpicktransfer|start.*pick.*transfer/.test(d)) return "StartPickTransfer Timeout";
  if (/midpick transfer|transfer timeout|transfer fault|conveyor timeout/.test(d)) return "Transfer Timeout";
  if (/unachievable distance|unahievable distance|unacheivable distance|distance cmd/.test(d)) return "Unachievable Distance CMD";
  if (/aplus requested vehicle shutdown/.test(d)) return "APlus - Requested Shutdown";
  if (/move without setmove/.test(d)) return "APlus - Move Without SetMove";
  if (/faulting vehicle during manual load transfer/.test(d)) return "Manual Load Transfer Fault";
  if (/wamas|failed to receive order/.test(d)) return "WAMAS Order Fault";
  if (/parameters for positioning at png/.test(d)) return "PNG Positioning Invalid";
  if (/nonzero demands when pendant/.test(d)) return "Pendant Fault";
  if (/passedge tripped|pass edge tripped|passedge|overshoot causing pass edge/.test(d)) return "PassEdge Tripped";
  if (/momentary power loss|recovered from.*power loss/.test(d)) return "Momentary Power Loss";
  if (/invalid host parameters/.test(d)) return "Invalid Host Parameters";
  if (/lost host|host connection|nvc.*crash/.test(d)) return "Lost Host Connection";
  if (/recovery failed|recovery/.test(d)) return "Recovery Failed";
  if (/no movement range matched|movement range matches/.test(d)) return "No Movement Range Match";
  if (/loss of traction feedback/.test(d)) return "Loss of Traction Feedback";
  if (/unexpected traction feedback|traction/.test(d)) return "Unexpected Traction Feedback";
  if (/scanner data|no scanner|scanner could not read/.test(d)) return "Nav - No Scanner Data";
  if (/low battery|low batt level|battery shutdown|voltage below.*nominal|battery not.*comm.*charger/.test(d)) return "Battery/Low Power";
  if (/ads error|tc ads/.test(d)) return "TC ADS Error";
  if (/amplifier fault|em brake|braking distance/.test(d)) return "Brake/Amplifier Fault";
  if (/actu?ator/.test(d)) return "Actuator Fault";
  if (/bu?mper/.test(d)) return "Bumper Trip";
  if (/reflector|not picking|failed to pick|didn.t line up|not aligned|misaligned/.test(d)) return "Alignment Reflector Not Detected";
  if (/not dropping|will not drop|dropping stack crooked|unable to complete drop/.test(d)) return "Drop Failure";
  if (/collided|collision|collide/.test(d)) return "Collision";
  if (/protection case/.test(d)) return "No Protection Case Match";
  if (/\bestop\b|e-stop/.test(d)) return "E-Stop";
  if (/stuck idle|deadlock|wait condition|full infeed|went idle|idle with active/.test(d)) return "Stuck/Blocked";
  if (/drive inverter|lag distance/.test(d)) return "HU Drive Inverter Lag";
  if (/load did not release|load defect|getextents|getrack|pantograph|tile not in position/.test(d)) return "AGF Load Fault";
  if (/\bfrbt\b|\bflbt\b|\brbt\b/.test(d)) return "FRBT/FLBT - Vehicle Blocked";
  if (/roller bridge/.test(d)) return "Roller Bridge";
  if (/prelift|flapper/.test(d)) return "Prelift Never Executed";
  if (/encoder/.test(d)) return "Encoder";
  if (/mfs forced error|mlu faulted|lhd movement/.test(d)) return "Crane Fault";
  return "Other";
}

function resolveLocation(area: string, desc: string): string {
  const ZONES = new Set([
    "HBW","LGP","LBW","TMK","HDW","LFP","BEAUTY","MAIN AISLE","MAIN","TUNNEL",
    "SOUTH LOOP","PG","DISH","STAGING","RAMP","A/B","B/C","A/B AISLE","B/C AISLE"
  ]);
  const areaClean = area.trim();
  if (!ZONES.has(areaClean.toUpperCase())) {
    const canon = canonicalizeLocation(areaClean);
    return canon || areaClean || "Unknown";
  }
  const LOC_RE = /(?:\bat\b|\bby\b|@|\bin\b|\bon\b|\bnext to\b|\bleaving\b)\s+(conv(?:eyor)?\s?\d{1,3}|bu\s?\d{1,3}|[a-z]{1,3}\s?\d{1,3})/gi;
  let match: RegExpExecArray | null;
  while ((match = LOC_RE.exec(desc)) !== null) {
    const raw = match[1].trim();
    if (/^(agv|agf|crane)/i.test(raw)) continue;
    const canon = canonicalizeLocation(raw);
    if (canon) return canon;
  }
  return areaClean || "Unknown";
}

function canonicalizeLocation(raw: string): string {
  const s = raw.trim().replace(/\s+/g, "").toUpperCase();
  let m = s.match(/^CONV(?:EYOR)?(\d{1,3})$/);
  if (m) {
    const n = parseInt(m[1]);
    if (n >= 1 && n <= 80) {
      const crane    = n % 10 === 0 ? n / 10 : Math.ceil(n / 10);
      const dir      = n % 10 === 1 ? "Inbound" : n % 10 === 0 ? "Outbound" : "internal";
      if (dir !== "internal") return `Conv ${n} — Crane ${crane} ${dir}`;
      return `Conv ${n} (internal)`;
    }
    return `Conv ${n}`;
  }
  m = s.match(/^BU(\d+)$/);
  if (m) return `BU${parseInt(m[1]).toString().padStart(2, "0")}`;
  m = s.match(/^AB(\d+)$/);
  if (m) return `AB${m[1]}`;
  m = s.match(/^([A-Z])(\d+)$/);
  if (m) return `${m[1]}${parseInt(m[2]).toString().padStart(2, "0")}`;
  return "";
}

function writeCell(
  sheet: ExcelScript.Worksheet, row: number, col: number, value: string,
  opts: { bold?: boolean; italic?: boolean; size?: number; color?: string }
) {
  const cell = sheet.getRangeByIndexes(row, col, 1, 1);
  cell.setValue(value);
  const fmt = cell.getFormat().getFont();
  if (opts.bold)   fmt.setBold(true);
  if (opts.italic) fmt.setItalic(true);
  if (opts.size)   fmt.setSize(opts.size);
  if (opts.color)  fmt.setColor(opts.color);
}

function writeHeaderRow(sheet: ExcelScript.Worksheet, row: number, headers: string[]) {
  const range = sheet.getRangeByIndexes(row, 0, 1, headers.length);
  range.setValues([headers]);
  range.getFormat().getFont().setBold(true);
  range.getFormat().getFill().setColor("#F1EFE8");
}
