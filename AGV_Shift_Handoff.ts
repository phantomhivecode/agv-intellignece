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

    // These fault types are always vehicle-specific regardless of how many AGVs appear
    const vehicleFaults = ["Actuator Fault", "Conveyor Still Transmitting", "Brake/Amplifier Fault", "Encoder", "Guard Rail Bumper Trip", "APlus - Requested Shutdown", "Battery/Low Power"];
    const isVehicleFault = vehicleFaults.includes(topFault);

    if (isVehicleFault) {
      talkingPoints.push(`${loc}: ${count} incidents — top fault "${topFault}" is a vehicle issue, not a location issue. Check each AGV individually: ${units.join(", ")}.`);
    } else if (units.length > 1) {
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
  if (/conveyor is still transmitting|tec conveyor is still transmitting|tec.*still transmit|conv.*still transmit|conveyor.*still transmit|coneyor.*still transmit|not sensing conv|still trnsmitting|still trasmitting/.test(d)) return "Conveyor Still Transmitting";
  if (/right handshake timeout|\brhst\b|\brhsto\b|\brht\b|right handshake|right hnadshake|right handhake|rh timeout/.test(d)) return "Handshake - Right Timeout";
  if (/left handshake timeout|\blhst\b|\blht\b|left handshake|left hansdhake|left handshakle|left hasndshake|lh timeout|left hanshake/.test(d)) return "Handshake - Left Timeout";
  if (/handshake|handhake|hasndshake|hnadshake|handhsake|handshakle|hansdhake|hanshake/.test(d)) {
    if (/\bleft\b/.test(d)) return "Handshake - Left Timeout";
    if (/\bright\b/.test(d)) return "Handshake - Right Timeout";
    return "Handshake - Unspecified";
  }
  if (/nav.*can.t match reflections|nav.*cant match reflections|can.t match reflections and targets|cant match reflections and targets/.test(d)) return "Nav - Can't Match Reflections";
  if (/nav.*no scanner data|nav - no scanner|nav no scanner/.test(d)) return "Nav - No Scanner Data";
  if (/guid.*travel|traveled too far without correction|travelled too far|travel.*too far|drift fault/.test(d)) return "Guid - Traveled Too Far";
  if (/nodelock|node lock|lost nav|not on node|node-lock|\bno nav\b|loss of nav|\bnav\b/.test(d)) return "Navigation/Nodelock";
  if (/alignment reflector not detected|alignment reflector|alignment sensor did not turn on|reflector not detected|alignment rail not in position/.test(d)) return "Alignment Reflector Not Detected";
  if (/no auto-mode guidesafe|no auto mode guidesafe|guidesafe|guidsafe/.test(d)) return "No Auto-Mode Guidesafe";
  if (/completpicktransfer|completepicktransfer|complet.*pick.*transfer|pick transfer t\.?o|pick.*transfer.*timeout|completepicktansfer|completepicktanser/.test(d)) return "CompletePickTransfer Timeout";
  if (/competedrop|completedrop.*transfer|complet.*drop.*transfer|drop transfer t\.?o|complete drop transfer timeout|complete drp transfer/.test(d)) return "CompleteDropTransfer Timeout";
  if (/startpicktransfer|start.*pick.*transfer|startpick.*t\.?o/.test(d)) return "StartPickTransfer Timeout";
  if (/conv.*failed to receive transfer|failed to receive transfer from agv/.test(d)) return "Transfer Timeout";
  if (/midpick transfer|transfer timeout|transfer fault|complete transfer t\.?o|conveyor timeout/.test(d)) return "Transfer Timeout";
  if (/unachievable distance|unahievable distance|unacheivable distance|unecheivable distance|unaheivable distance|unechievable distance|distance cmd/.test(d)) return "Unachievable Distance CMD";
  if (/aplus requested vehicle shutdown|aplus.*shutdown|aplus requested|auto shutdown pending.*aplus/.test(d)) return "APlus - Requested Shutdown";
  if (/aplus.*move without setmove|move without setmove|aplus.*setmove/.test(d)) return "APlus - Move Without SetMove";
  if (/faulting vehicle during manual load transfer|manual load transfer/.test(d)) return "Manual Load Transfer Fault";
  if (/wamas|failed to receive order/.test(d)) return "WAMAS Order Fault";
  if (/parameters for positioning at png|positioning at png conveyors/.test(d)) return "PNG Positioning Invalid";
  if (/inventory does not exist on conveyor|vehicle picking.*inventory/.test(d)) return "Inventory Not On Conveyor";
  if (/nonzero demands when pendant|pendant activated/.test(d)) return "Pendant Fault";
  if (/too far from start node|start node of commanded move/.test(d)) return "Start Node Position Fault";
  if (/passedge tripped|pass edge tripped|conveyor passedge|passedge|pass edge|overshoot causing pass edge/.test(d)) return "PassEdge Tripped";
  if (/momentary power loss|recovered from.*power loss/.test(d)) return "Momentary Power Loss";
  if (/invalid host parameters/.test(d)) return "Invalid Host Parameters";
  if (/lost host|host connection|nvc.*crash|nvc.*closed with error|lost network connection/.test(d)) return "Lost Host Connection";
  if (/recovery failed|recovery/.test(d)) return "Recovery Failed";
  if (/no movement range matched|movement range matches|movement range matched|no movement range/.test(d)) return "No Movement Range Match";
  if (/loss of traction feedback/.test(d)) return "Loss of Traction Feedback";
  if (/unexpected traction feedback|traction/.test(d)) return "Unexpected Traction Feedback";
  if (/scanner data|no scanner|scanner could not read/.test(d)) return "Nav - No Scanner Data";
  if (/low battery|low batt level|battery shutdown|batt\.?\s*monitor|voltage below.*nominal|battery not.*comm.*charger|battery not.*reporting.*charger|battery.*offline|died on charger/.test(d)) return "Battery/Low Power";
  if (/flexi|no comms|can message/.test(d)) return "Comms/Flexisoft";
  if (/ads error|tc ads/.test(d)) return "TC ADS Error";
  if (/amplifier fault|em brake|braking distance/.test(d)) return "Brake/Amplifier Fault";
  if (/drive inverter|lag distance|drift inverter/.test(d)) return "HU Drive Inverter Lag";
  if (/actu?ator/.test(d)) return "Actuator Fault";
  if (/bu?mper|numper/.test(d)) return "Bumper Trip";
  if (/conveyor guard blocking|guard blocking stack|can.t load due to.*guard/.test(d)) return "Stack Hitting Guard Rail";
  if (/bumper.*guard rail|bumper.*guardrail|guard rail.*bumper|guardrail.*bumper|too close to guard|drove too close/.test(d)) return "Guard Rail Bumper Trip";
  if (/went crooked|crooked.*pick|hit the agv|realigned.*agv|wasn.t lined up|binstack/.test(d)) return "Load/Stack Alignment";
  if (/reflector|not picking|failed to pick|fail to pick|didn.t line up|did not lin up|did not line up|not aligned|misaligned|not lined up|stack too far off|alignment sensor/.test(d)) return "Alignment Reflector Not Detected";
  if (/not dropping|will not drop|dropping stack crooked|too far forward dropping|couldn.t finish drop|couldnt finsih drop|unable to complete drop|does not drop|doesn.t drop/.test(d)) return "Drop Failure";
  if (/collided|collision|collide|pushed.*crooked/.test(d)) return "Collision";
  if (/protection case/.test(d)) return "No Protection Case Match";
  if (/\bestop\b|e-stop|e stop/.test(d)) return "E-Stop";
  if (/stuck idle|deadlock|wait condition|full infeed|stuck in idle|went idle|idle with active|back.?up|causing a back|\bin idle\b|died in|died while|shut down in middle|just turned off|stuck evaluating|waiting for pick|waiting for drop|waiting to drop|waiting to pick/.test(d)) return "Stuck/Blocked";
  if (/load did not release|load defect fault|load push.pull|load beyond reach|getextents|getrack|pantograph|tile not in position/.test(d)) return "AGF Load Fault";
  if (/\bfrbt\b|\bflbt\b|\brbt\b/.test(d)) return "FRBT/FLBT - Vehicle Blocked";
  if (/roller bridge/.test(d)) return "Roller Bridge";
  if (/bin stack|load bar|retract loadbar|vehicle not loaded|pallet|sideshift|lateral offset/.test(d)) return "Load Handling";
  if (/ghost|phantom lu|ulid/.test(d)) return "Phantom/LU Status";
  if (/inventory/.test(d)) return "Inventory";
  if (/hold by/.test(d)) return "Hold by CS";
  if (/safety module|clearance|door to open/.test(d)) return "Safety";
  if (/prelift|flapper/.test(d)) return "Prelift Never Executed";
  if (/encoder/.test(d)) return "Encoder";
  if (/mfs forced error|mlu faulted|lhd movement/.test(d)) return "Crane Fault";
  if (/jbt testing|being worked on|autefa|coned off|contractors/.test(d)) return "Maintenance/Testing";
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
