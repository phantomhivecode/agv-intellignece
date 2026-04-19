/**
 * AGV Intelligence — Office Script for SharePoint Excel
 *
 * Reads all shift tabs in the workbook, extracts incidents, categorizes faults,
 * and writes a refreshed "Intelligence" tab with rankings, hot zones, signature
 * faults, and a daily trend.
 *
 * HOW TO USE:
 *   1. In Excel (web), go to Automate > New Script
 *   2. Paste this entire file, replacing the default main() stub
 *   3. Save as "AGV Intelligence"
 *   4. Click Run. Re-run any time — the Intelligence tab rebuilds fresh.
 *
 * ASSUMPTIONS (matches your current report format):
 *   - Each shift is on its own tab named like "04.17.2026 A"
 *   - A "Master" tab exists as the template (skipped)
 *   - Incident rows: column B = time (e.g. 07:02), column C = area, column D = description
 *   - Sheets with any other naming convention are skipped safely
 */

function main(workbook: ExcelScript.Workbook) {
  const DEBUG_LOG = true;

  // ---------- 1. Collect incidents from every shift tab ----------
  type Incident = {
    date: string;      // "2026-04-17"
    shift: string;     // "A" | "B" | "C" | "D" | "?"
    time: string;      // "HH:MM"
    area: string;      // raw Area / Line column value
    location: string;  // resolved specific location (e.g. "C03", "Conv 50") or zone fallback
    zone: string;      // broad zone (e.g. "HBW", "LGP") when resolvable
    description: string;
    unit: string;      // "AGV 42" | "CRANE 1" | "Unknown"
    category: string;
  };

  const incidents: Incident[] = [];
  const sheetNameRegex = /^(\d{2})[.,](\d{2})[.,](\d{4})\s*([A-D])?/;
  const timeRegex = /^(\d{1,2}):(\d{2})/;
  const unitRegex = /(AGV|AGF|Crane)\s*#?\s*(\d+)/i;

  const sheets = workbook.getWorksheets();
  // Pre-read all sheet names in one pass (avoids calling getName() inside the main loop).
  // We don't store Worksheet references — Office Scripts disallows aliasing APIs.
  const sheetNames: string[] = sheets.map(s => s.getName());

  for (let i = 0; i < sheets.length; i++) {
    const name = sheetNames[i];
    if (name === "Master" || name === "Intelligence") continue;

    const dateMatch = name.match(sheetNameRegex);
    if (!dateMatch) continue; // skip any non-shift tab

    const date = `${dateMatch[3]}-${dateMatch[1]}-${dateMatch[2]}`;
    const shift = dateMatch[4] || "?";

    // Reading each sheet's range is unavoidable — we have to look at every tab.
    // Office Scripts may flag this as a perf warning; it's acceptable here.
    const usedRange = sheets[i].getUsedRange();
    if (!usedRange) continue;
    const values = usedRange.getValues();

    for (const row of values) {
      if (!row || row.length < 4) continue;
      const rawTime = row[1];
      const area = row[2] ? String(row[2]).trim() : "";
      const desc = row[3] ? String(row[3]).trim() : "";
      if (!desc || desc.toLowerCase() === "none") continue;

      const timeStr = normalizeTime(rawTime);
      if (!timeStr) continue;

      const unitMatch = desc.match(unitRegex);
      const unit = unitMatch ? `${unitMatch[1].toUpperCase()} ${unitMatch[2]}` : "Unknown";

      const resolved = resolveLocation(area, desc);

      incidents.push({
        date, shift, time: timeStr, area,
        location: resolved.location, zone: resolved.zone,
        description: desc, unit, category: categorize(desc)
      });
    }
  }

  if (DEBUG_LOG) console.log(`Parsed ${incidents.length} incidents from ${sheets.length} tabs`);

  // ---------- 2. Aggregate ----------
  const unitCounts = countBy(incidents, i => i.unit);
  const locationCounts = countBy(incidents, i => i.location);
  const categoryCounts = countBy(incidents, i => i.category);
  const dayCounts = countBy(incidents, i => i.date);

  // Signature faults: for each vehicle, find its dominant fault category
  const byUnit: Record<string, string[]> = {};
  for (const i of incidents) {
    if (i.unit === "Unknown") continue;
    if (!byUnit[i.unit]) byUnit[i.unit] = [];
    byUnit[i.unit].push(i.category);
  }
  type Signature = { unit: string; total: number; topCat: string; topCount: number; pct: number };
  const signatures: Signature[] = [];
  for (const unit of Object.keys(byUnit)) {
    const cats = byUnit[unit];
    const top = topEntry(countBy(cats, c => c));
    signatures.push({
      unit, total: cats.length,
      topCat: top.key, topCount: top.count,
      pct: Math.round((top.count / cats.length) * 100)
    });
  }
  signatures.sort((a, b) => b.total - a.total);

  // Hot locations: specific location (or zone fallback) with dominant fault
  const byLocation: Record<string, string[]> = {};
  for (const i of incidents) {
    if (!i.location) continue;
    if (!byLocation[i.location]) byLocation[i.location] = [];
    byLocation[i.location].push(i.category);
  }
  const hotLocations: Signature[] = [];
  for (const loc of Object.keys(byLocation)) {
    const cats = byLocation[loc];
    const top = topEntry(countBy(cats, c => c));
    hotLocations.push({
      unit: loc, total: cats.length,
      topCat: top.key, topCount: top.count,
      pct: Math.round((top.count / cats.length) * 100)
    });
  }
  hotLocations.sort((a, b) => b.total - a.total);

  // ---------- 3. Rebuild Intelligence sheet ----------
  let sheet = workbook.getWorksheet("Intelligence");
  if (sheet) sheet.delete();
  sheet = workbook.addWorksheet("Intelligence");
  sheet.activate();

  let r = 0; // current row (0-indexed)

  // --- Header block ---
  writeCell(sheet, r, 0, "AGV Intelligence Report", { bold: true, size: 16 });
  r++;
  writeCell(sheet, r, 0, `Generated ${new Date().toLocaleString()}`, { italic: true, color: "#666666" });
  r += 2;

  // --- Summary tiles (as a small grid) ---
  const totalIncidents = incidents.length;
  const uniqueDays = Object.keys(dayCounts).length;
  const avgPerDay = uniqueDays > 0 ? Math.round(totalIncidents / uniqueDays) : 0;
  const topUnit = Object.keys(unitCounts).filter(u => u !== "Unknown")
    .sort((a, b) => unitCounts[b] - unitCounts[a])[0];
  const topLocation = Object.keys(locationCounts).sort((a, b) => locationCounts[b] - locationCounts[a])[0];

  writeCell(sheet, r, 0, "At a glance", { bold: true, size: 13 });
  r++;
  const tiles: [string, string][] = [
    ["Total incidents", String(totalIncidents)],
    ["Days covered", String(uniqueDays)],
    ["Avg per day", String(avgPerDay)],
    ["Top offender", `${topUnit} (${unitCounts[topUnit]})`],
    ["Hot location", `${topLocation} (${locationCounts[topLocation]})`]
  ];
  for (const [label, value] of tiles) {
    writeCell(sheet, r, 0, label, { color: "#666666" });
    writeCell(sheet, r, 1, value, { bold: true });
    r++;
  }
  r += 2;

  // --- Top 15 offending units ---
  writeCell(sheet, r, 0, "Top 15 offending units", { bold: true, size: 13 });
  r++;
  writeHeaderRow(sheet, r, ["Unit", "Incidents", "Signature fault", "%"]);
  r++;
  const topUnits = signatures.slice(0, 15);
  for (const s of topUnits) {
    const vals: (string | number)[] = [s.unit, s.total, s.topCat, s.pct + "%"];
    sheet.getRangeByIndexes(r, 0, 1, vals.length).setValues([vals]);
    // Flag signature patterns (>=40% concentration with >=5 incidents)
    if (s.pct >= 40 && s.total >= 5) {
      sheet.getRangeByIndexes(r, 2, 1, 1).getFormat().getFill().setColor("#FAEEDA");
    }
    r++;
  }
  r += 2;

  // --- Top 15 hot locations ---
  writeCell(sheet, r, 0, "Top 15 hot locations", { bold: true, size: 13 });
  r++;
  writeHeaderRow(sheet, r, ["Location", "Incidents", "Dominant fault", "%"]);
  r++;
  const topZones = hotLocations.slice(0, 15);
  for (const s of topZones) {
    const vals: (string | number)[] = [s.unit, s.total, s.topCat, s.pct + "%"];
    sheet.getRangeByIndexes(r, 0, 1, vals.length).setValues([vals]);
    if (s.pct >= 50 && s.total >= 5) {
      sheet.getRangeByIndexes(r, 2, 1, 1).getFormat().getFill().setColor("#FAEEDA");
    }
    r++;
  }
  r += 2;

  // --- Failure categories ---
  writeCell(sheet, r, 0, "Failure categories", { bold: true, size: 13 });
  r++;
  writeHeaderRow(sheet, r, ["Category", "Count", "%"]);
  r++;
  const catList = Object.keys(categoryCounts).sort((a, b) => categoryCounts[b] - categoryCounts[a]);
  for (const c of catList) {
    const n = categoryCounts[c];
    const pct = Math.round((n / totalIncidents) * 100);
    sheet.getRangeByIndexes(r, 0, 1, 3).setValues([[c, n, pct + "%"]]);
    r++;
  }
  r += 2;

  // --- Daily trend ---
  writeCell(sheet, r, 0, "Daily trend", { bold: true, size: 13 });
  r++;
  writeHeaderRow(sheet, r, ["Date", "Incidents"]);
  r++;
  const sortedDays = Object.keys(dayCounts).sort();
  const trendStartRow = r;
  for (const d of sortedDays) {
    sheet.getRangeByIndexes(r, 0, 1, 2).setValues([[d, dayCounts[d]]]);
    r++;
  }
  const trendEndRow = r - 1;
  r += 2;

  // --- Findings block (the "so what") ---
  writeCell(sheet, r, 0, "Key findings", { bold: true, size: 13 });
  r++;
  const findings = buildFindings(signatures, hotLocations, totalIncidents, avgPerDay);
  for (const line of findings) {
    writeCell(sheet, r, 0, line, {});
    // Merge across a few columns so the text has room to display
    sheet.getRangeByIndexes(r, 0, 1, 6).merge();
    r++;
  }

  // --- Chart: daily trend ---
  if (trendEndRow >= trendStartRow) {
    const trendRange = sheet.getRangeByIndexes(trendStartRow, 0, trendEndRow - trendStartRow + 1, 2);
    const chart = sheet.addChart(ExcelScript.ChartType.line, trendRange);
    chart.getTitle().setText("Daily incident trend");
    chart.setPosition(sheet.getCell(3, 6), sheet.getCell(20, 14));
    chart.getLegend().setVisible(false);
  }

  // --- Chart: top 5 offending units (bar chart) ---
  // Write a small data block to columns F:G (cols 5:6) starting at row 22 as chart source
  const top5StartRow = 22;
  sheet.getRangeByIndexes(top5StartRow, 5, 1, 2).setValues([["Unit", "Incidents"]]);
  const top5Units = signatures.filter(s => s.unit !== "Unknown").slice(0, 3);
  for (let t = 0; t < top5Units.length; t++) {
    sheet.getRangeByIndexes(top5StartRow + 1 + t, 5, 1, 2)
      .setValues([[top5Units[t].unit, top5Units[t].total]]);
  }
  const top5Range = sheet.getRangeByIndexes(top5StartRow, 5, top5Units.length + 1, 2);
  const top5Chart = sheet.addChart(ExcelScript.ChartType.barClustered, top5Range);
  top5Chart.getTitle().setText("Top Three Reported AGVs");
  top5Chart.setPosition(sheet.getCell(21, 6), sheet.getCell(35, 14));
  top5Chart.getLegend().setVisible(false);

  // Column widths
  sheet.getRange("A:A").getFormat().setColumnWidth(220);
  sheet.getRange("B:B").getFormat().setColumnWidth(90);
  sheet.getRange("C:C").getFormat().setColumnWidth(180);
  sheet.getRange("D:D").getFormat().setColumnWidth(70);

  console.log(`Intelligence tab rebuilt with ${totalIncidents} incidents across ${uniqueDays} days.`);
}

// ========== Helpers ==========

function normalizeTime(raw: string | number | boolean): string | null {
  if (raw === null || raw === undefined || raw === "") return null;
  if (typeof raw === "number") {
    // Excel stores times as fraction of a day
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

  // Bare unit name only — no actionable fault text
  if (/^(agv|agf|crane)\s*#?\s*\d+\s*$/.test(d)) return "No Description";

  // ── OFFICIAL ALARM TEXTS — anchored to exact system phrases ──────────────

  // Conveyor Is Still Transmitting (two official variants, same root cause)
  // Also catches typos: "tramitting", "transmittuing", "coneyor", "witout"
  if (/conveyor is still transmitting|tec conveyor is still transmitting|tec.*still transmit|conv.*still transmit|conveyor.*still transmit|coneyor.*still transmit|not sensing conv/.test(d)) return "Conveyor Still Transmitting";

  // Handshakes — official: "Left/Right Handshake Timeout 5min"
  // Catches all abbreviations and typos: LHST, RHST, LHT, RHT, RHSTO, hanshake, handhsake, etc.
  if (/right handshake timeout|\brhst\b|\brhsto\b|\brht\b|right handshake|right hnadshake|right handhake|rh timeout/.test(d)) return "Handshake - Right Timeout";
  if (/left handshake timeout|\blhst\b|\blht\b|left handshake|left hansdhake|left handshakle|left hasndshake|lh timeout|left hanshake/.test(d)) return "Handshake - Left Timeout";
  if (/handshake|handhake|hasndshake|hnadshake|handhsake|handshakle|hansdhake|hanshake/.test(d)) {
    if (/\bleft\b/.test(d)) return "Handshake - Left Timeout";
    if (/\bright\b/.test(d)) return "Handshake - Right Timeout";
    return "Handshake - Unspecified";
  }

  // Navigation — three distinct official alarms, different root causes
  if (/nav.*can.t match reflections|nav.*cant match reflections|can.t match reflections and targets|cant match reflections and targets/.test(d)) return "Nav - Can't Match Reflections";
  if (/nav.*no scanner data|nav - no scanner|nav no scanner/.test(d)) return "Nav - No Scanner Data";
  if (/guid.*traveled too far|traveled too far without correction/.test(d)) return "Guid - Traveled Too Far";
  if (/nodelock|node lock|lost nav|not on node|node-lock|\bno nav\b|loss of nav/.test(d)) return "Navigation/Nodelock";

  // Alignment Reflector Not Detected — official alarm + all paraphrase variants
  if (/alignment reflector not detected|alignment reflector|alignment sensor did not turn on|reflector not detected/.test(d)) return "Alignment Reflector Not Detected";

  // No Auto-Mode Guidesafe — official alarm
  if (/no auto-mode guidesafe|no auto mode guidesafe|guidesafe|guidsafe/.test(d)) return "No Auto-Mode Guidesafe";

  // Transfer Timeouts — three distinct official alarms (note: system itself spells "Complet" and "Compete")
  // "T O" is how engineers abbreviate "Timeout" in notes
  if (/completpicktransfer|completepicktransfer|complet.*pick.*transfer.*t|pick transfer t\.?o|pick.*transfer.*timeout/.test(d)) return "CompletePickTransfer Timeout";
  if (/competedrop|completedrop.*transfer|complet.*drop.*transfer.*t|drop transfer t\.?o|complete drop transfer timeout|complete drp transfer/.test(d)) return "CompleteDropTransfer Timeout";
  if (/startpicktransfer|start.*pick.*transfer|startpick.*t\.?o/.test(d)) return "StartPickTransfer Timeout";
  if (/midpick transfer|transfer timeout|transfer fault|complete transfer t\.?o|conveyor timeout/.test(d)) return "Transfer Timeout";

  // Unachievable Distance CMD — catches all misspellings seen in the data
  if (/unachievable distance|unahievable distance|unacheivable distance|unecheivable distance|unaheivable distance|unechievable distance|distance cmd/.test(d)) return "Unachievable Distance CMD";

  // APlus — two distinct official alarms
  if (/aplus requested vehicle shutdown|aplus.*shutdown|aplus requested/.test(d)) return "APlus - Requested Shutdown";
  if (/aplus.*move without setmove|move without setmove|aplus.*setmove/.test(d)) return "APlus - Move Without SetMove";

  // Faulting Vehicle During Manual Load Transfer — official alarm
  if (/faulting vehicle during manual load transfer|manual load transfer/.test(d)) return "Manual Load Transfer Fault";

  // WAMAS Order Fault — official alarm
  if (/wamas|failed to receive order/.test(d)) return "WAMAS Order Fault";

  // PNG Positioning Invalid — official alarm
  if (/parameters for positioning at png|positioning at png conveyors/.test(d)) return "PNG Positioning Invalid";

  // Inventory Not On Conveyor — official alarm
  if (/inventory does not exist on conveyor|vehicle picking.*inventory/.test(d)) return "Inventory Not On Conveyor";

  // Pendant Fault — official alarm: "Nonzero demands when pendant activated"
  if (/nonzero demands when pendant|pendant activated/.test(d)) return "Pendant Fault";

  // Start Node Position Fault — official alarm
  if (/too far from start node|start node of commanded move/.test(d)) return "Start Node Position Fault";

  // PassEdge Tripped — official alarm + paraphrase variants
  if (/passedge tripped|pass edge tripped|conveyor passedge|passedge|pass edge|overshoot causing pass edge/.test(d)) return "PassEdge Tripped";

  // Momentary Power Loss — official alarm
  if (/momentary power loss|recovered from.*power loss/.test(d)) return "Momentary Power Loss";

  // Host faults — two distinct types
  if (/invalid host parameters/.test(d)) return "Invalid Host Parameters";
  if (/lost host|host connection|nvc.*crash|nvc.*closed with error/.test(d)) return "Lost Host Connection";

  // Recovery Failed — official alarm
  if (/recovery failed|recovery/.test(d)) return "Recovery Failed";

  // No Movement Range Match — official alarm
  if (/no movement range matched|movement range matches|movement range matched/.test(d)) return "No Movement Range Match";

  // Traction — two distinct official alarms
  if (/loss of traction feedback/.test(d)) return "Loss of Traction Feedback";
  if (/unexpected traction feedback|traction/.test(d)) return "Unexpected Traction Feedback";

  // ── NON-OFFICIAL BUT WELL-UNDERSTOOD FAULTS ──────────────────────────────

  // Scanner — catches free-text variations
  if (/scanner data|no scanner|scanner could not read/.test(d)) return "Nav - No Scanner Data";

  // Battery / power
  if (/low battery|low batt level|battery shutdown|batt\.?\s*monitor|voltage below.*nominal|battery not.*comm.*charger|battery not.*reporting.*charger/.test(d)) return "Battery/Low Power";

  // Comms / Flexisoft
  if (/flexi|no comms|can message/.test(d)) return "Comms/Flexisoft";

  // TC ADS Error
  if (/ads error|tc ads/.test(d)) return "TC ADS Error";

  // Brake / Amplifier
  if (/amplifier fault|em brake|braking distance/.test(d)) return "Brake/Amplifier Fault";

  // Actuator — catches "Acuator" typo
  if (/actu?ator/.test(d)) return "Actuator Fault";

  // Bumper — catches "Bymper" typo
  if (/bu?mper/.test(d)) return "Bumper Trip";

  // Alignment variants not using exact official phrase
  if (/reflector|not picking|failed to pick|fail to pick|didn.t line up|did not lin up|did not line up|not aligned|misaligned|not lined up|stack too far off|alignment sensor/.test(d)) return "Alignment Reflector Not Detected";

  // Drop failure
  if (/not dropping|will not drop|dropping stack crooked|too far forward dropping|couldn.t finish drop|couldnt finsih drop|unable to complete drop|does not drop|doesn.t drop/.test(d)) return "Drop Failure";

  // Collision
  if (/collided|collision|collide|pushed.*crooked/.test(d)) return "Collision";

  // Protection Case
  if (/protection case/.test(d)) return "No Protection Case Match";

  // E-Stop
  if (/\bestop\b|e-stop|e stop/.test(d)) return "E-Stop";

  // Stuck / Blocked — includes "waiting for pick/drop operation to complete"
  if (/stuck idle|deadlock|wait condition|full infeed|stuck in idle|went idle|idle with active|\bblocking\b|back.?up|causing a back|\bin idle\b|died in|died while|shut down in middle|just turned off|stuck evaluating|waiting for pick|waiting for drop|waiting to drop|waiting to pick/.test(d)) return "Stuck/Blocked";

  // HU Drive Inverter Lag (crane-specific)
  if (/drive inverter|lag distance/.test(d)) return "HU Drive Inverter Lag";

  // AGF-specific load faults (AGF 501/502/503 load arm faults)
  if (/load did not release|load defect fault|load push.pull|load beyond reach|getextents|getrack|pantograph|tile not in position/.test(d)) return "AGF Load Fault";

  // FRBT / FLBT / RBT — vehicle-follows/blocks another (shop codes)
  if (/\bfrbt\b|\bflbt\b|\brbt\b/.test(d)) return "FRBT/FLBT - Vehicle Blocked";

  // Roller Bridge
  if (/roller bridge/.test(d)) return "Roller Bridge";

  // Load Handling
  if (/bin stack|load bar|retract loadbar|vehicle not loaded|pallet|sideshift|lateral offset/.test(d)) return "Load Handling";

  // Phantom / LU Status
  if (/ghost|phantom lu|ulid/.test(d)) return "Phantom/LU Status";

  // Inventory
  if (/inventory/.test(d)) return "Inventory";

  // Hold by CS
  if (/hold by/.test(d)) return "Hold by CS";

  // Safety / crane safety
  if (/safety module|clearance|door to open/.test(d)) return "Safety";

  // Prelift / Flapper (AGF-specific)
  if (/prelift|flapper/.test(d)) return "Prelift Never Executed";

  // Encoder
  if (/encoder/.test(d)) return "Encoder";

  // Crane faults
  if (/mfs forced error|mlu faulted|lhd movement/.test(d)) return "Crane Fault";

  // Maintenance / testing events (not faults)
  if (/jbt testing|being worked on|autefa|coned off|contractors/.test(d)) return "Maintenance/Testing";

  return "Other";
}

// resolveLocation — returns specific location from Area col or description,
// falling back to zone name only when no specific location is findable.
// Uses the official location list to validate extractions and derive zones.
//
// Examples:
//   area="LGP"  desc="AGV 14 LHST at H05"        → { location:"H05 (PG)",        zone:"LGP" }
//   area="HBW"  desc="AGV 13 bumper at Conv 80"   → { location:"Conv 80 — Crane 8 Outbound", zone:"HBW" }
//   area="C03"  desc="AGV 53 alignment fault"     → { location:"C03",              zone:"LGP" }
//   area="TMK"  desc="AGV 4 nav fault"            → { location:"TMK",              zone:"TMK" }
function resolveLocation(area: string, desc: string): { location: string; zone: string } {

  // ── Known zone names that should trigger a description lookup ──────────────
  const ZONES = new Set([
    "HBW","LGP","LBW","TMK","HDW","LFP","BEAUTY","MAIN AISLE","MAIN","TUNNEL",
    "SOUTH LOOP","PG","DISH","STAGING","RAMP","A/B","B/C","A/B AISLE","B/C AISLE"
  ]);

  const areaClean = area.trim();
  const areaUpper = areaClean.toUpperCase();

  // If area column already holds a specific location, canonicalize and use it
  if (!ZONES.has(areaUpper)) {
    const canon = canonicalizeLocation(areaClean);
    if (canon) {
      const z = deriveZone(canon);
      return { location: canon, zone: z };
    }
    if (areaClean) return { location: areaClean, zone: "" };
  }

  // Area is a zone — scan description for a specific location after trigger words.
  // Trigger words: at, by, @, in, on, next to, outside of, leaving, picking at, dropping at
  const LOC_RE = /(?:\bat\b|\bby\b|@|\bin\b|\bon\b|\bnext to\b|\boutside of\b|\bleaving\b)\s+(conv(?:eyor)?\s?\d{1,3}|bu\s?\d{1,3}|[a-z]{1,3}\s?\d{1,3})/gi;
  let match: RegExpExecArray | null;
  while ((match = LOC_RE.exec(desc)) !== null) {
    const raw = match[1].trim();
    // Skip AGV/AGF/Crane unit numbers (e.g. "by AGV 24")
    if (/^(agv|agf|crane)/i.test(raw)) continue;
    const canon = canonicalizeLocation(raw);
    if (canon) {
      const z = deriveZone(canon) || areaClean;
      return { location: canon, zone: z };
    }
  }

  // No specific location found — return zone as both
  return { location: areaClean || "Unknown", zone: areaClean };
}

// deriveZone — given a canonicalized location code, return its parent zone.
// Built from the official location list and crane conveyor system.
function deriveZone(loc: string): string {
  const l = loc.toUpperCase().replace(/\s+/g, "");

  // LGP: A-line, AB, B-line, C-line
  if (/^A\d{2}$/.test(l)) return "LGP";
  if (/^AB\d+$/.test(l)) return "LGP";
  if (/^B\d{2}$/.test(l)) return "LGP";
  if (/^C\d{2}$/.test(l)) return "LGP";

  // TMK: D-line, E-line (Technimark)
  if (/^D\d{2}$/.test(l)) return "TMK";
  if (/^E\d{2}$/.test(l)) return "TMK";

  // PG (HDW): H-line and BU lines (P&G lines)
  if (/^H\d{2}$/.test(l)) return "PG";
  if (/^BU\d{2}$/.test(l)) return "PG";

  // HBW: crane inbound/outbound conveyors (Conv 1–80, ending in 1 or 0)
  const convMatch = loc.match(/^Conv\s*(\d+)$/i);
  if (convMatch) {
    const n = parseInt(convMatch[1]);
    if (n >= 1 && n <= 80) return "HBW";
  }

  return "";
}

// canonicalizeLocation — normalize raw location text to standard form.
// C3→C03, BU4→BU04, Conv 1→"Conv 1 — Crane 1 Inbound", H5→H05, etc.
// Returns "" if the input doesn't match any known location pattern.
function canonicalizeLocation(raw: string): string {
  const s = raw.trim().replace(/\s+/g, "").toUpperCase();

  // Conv / Conveyor + number → enrich with crane and direction
  let m = s.match(/^CONV(?:EYOR)?(\d{1,3})$/);
  if (m) {
    const n = parseInt(m[1]);
    if (n >= 1 && n <= 80) {
      // Only inbound (ends in 1) and outbound (ends in 0) are AGV pick/drop points
      const isInbound  = n % 10 === 1;
      const isOutbound = n % 10 === 0;
      const crane = isOutbound ? n / 10 : Math.ceil(n / 10);
      if (isInbound)  return `Conv ${n} — Crane ${crane} Inbound`;
      if (isOutbound) return `Conv ${n} — Crane ${crane} Outbound`;
      // Internal buffer position — AGVs don't pick/drop here, use as-is
      return `Conv ${n} (internal)`;
    }
    return `Conv ${n}`;
  }

  // BU + number (P&G lines BU01–BU06)
  m = s.match(/^BU(\d+)$/);
  if (m) {
    const n = parseInt(m[1]);
    return `BU${n.toString().padStart(2, "0")}`;
  }

  // AB + number (AB1–AB5, LGP zone)
  m = s.match(/^AB(\d+)$/);
  if (m) return `AB${m[1]}`;

  // Single letter + digits: A01–A14, B01–B10, C02–C14, D01–D12, E01–E10, H01–H06
  // Zero-pads single digit: C3→C03, H5→H05
  m = s.match(/^([A-Z])(\d+)$/);
  if (m) {
    const num = parseInt(m[2]);
    return `${m[1]}${num.toString().padStart(2, "0")}`;
  }

  return "";
}

function countBy<T>(arr: T[], keyFn: (item: T) => string): Record<string, number> {
  const out: Record<string, number> = {};
  for (const item of arr) {
    const k = keyFn(item);
    out[k] = (out[k] || 0) + 1;
  }
  return out;
}

function topEntry(counts: Record<string, number>): { key: string; count: number } {
  let bestKey = "";
  let best = -1;
  for (const k of Object.keys(counts)) {
    if (counts[k] > best) { best = counts[k]; bestKey = k; }
  }
  return { key: bestKey, count: best };
}

function writeCell(
  sheet: ExcelScript.Worksheet, row: number, col: number, value: string,
  opts: { bold?: boolean; italic?: boolean; size?: number; color?: string }
) {
  const cell = sheet.getRangeByIndexes(row, col, 1, 1);
  cell.setValue(value);
  const fmt = cell.getFormat().getFont();
  if (opts.bold) fmt.setBold(true);
  if (opts.italic) fmt.setItalic(true);
  if (opts.size) fmt.setSize(opts.size);
  if (opts.color) fmt.setColor(opts.color);
}

function writeHeaderRow(sheet: ExcelScript.Worksheet, row: number, headers: string[]) {
  const range = sheet.getRangeByIndexes(row, 0, 1, headers.length);
  range.setValues([headers]);
  range.getFormat().getFont().setBold(true);
  range.getFormat().getFill().setColor("#F1EFE8");
}

function buildFindings(
  signatures: { unit: string; total: number; topCat: string; topCount: number; pct: number }[],
  hotLocations: { unit: string; total: number; topCat: string; topCount: number; pct: number }[],
  totalIncidents: number, avgPerDay: number
): string[] {
  const lines: string[] = [];
  lines.push(`• ${totalIncidents} total incidents averaging ${avgPerDay} per day.`);

  // Top 3 signature faults (where concentration matters)
  const signatureAlerts = signatures
    .filter(s => s.total >= 10 && s.pct >= 40)
    .slice(0, 3);
  for (const s of signatureAlerts) {
    lines.push(`• ${s.unit} — ${s.pct}% of ${s.total} incidents are ${s.topCat}. Likely vehicle-specific issue.`);
  }

  // Top 2 hot locations with concentration
  const locAlerts = hotLocations
    .filter(z => z.total >= 10 && z.pct >= 40)
    .slice(0, 2);
  for (const z of locAlerts) {
    lines.push(`• ${z.unit} — ${z.total} incidents, ${z.pct}% are ${z.topCat}. Worth investigating this location.`);
  }

  return lines;
}
