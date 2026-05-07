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
    // Wrap in try-catch so a single bad/empty sheet doesn't stop the whole script.
    let values: (string | number | boolean)[][] = [];
    try {
      const usedRange = sheets[i].getUsedRange();
      if (!usedRange) continue;
      values = usedRange.getValues();
    } catch (e) {
      console.log(`Skipping sheet "${sheetNames[i]}" — could not read range: ${e}`);
      continue;
    }

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

  // --- New Vehicle Performance (AGV 53–58) ---
  // Commissioned January 2026 — tracked separately from main fleet (AGV 1–52)
  const NEW_FLEET = ["AGV 53", "AGV 54", "AGV 55", "AGV 56", "AGV 57", "AGV 58"];
  const COMMISSION_DATE = new Date("2026-01-01");
  const today = new Date();
  const monthsInService = Math.floor(
    (today.getTime() - COMMISSION_DATE.getTime()) / (1000 * 60 * 60 * 24 * 30)
  );

  writeCell(sheet, r, 0, "New Vehicle Performance — AGV 53–58", { bold: true, size: 13 });
  r++;
  writeCell(sheet, r, 0, `Commissioned January 2026 · ${monthsInService} months in service`, { italic: true, color: "#666666" });
  r++;
  writeHeaderRow(sheet, r, ["Unit", "Incidents (MTD)", "Signature Fault", "%", "Top Location"]);
  r++;

  let newFleetTotal = 0;
  for (const unit of NEW_FLEET) {
    const unitIncidents = incidents.filter(i => i.unit === unit);
    const count = unitIncidents.length;
    newFleetTotal += count;

    if (count === 0) {
      sheet.getRangeByIndexes(r, 0, 1, 5).setValues([[unit, 0, "No incidents", "—", "—"]]);
      sheet.getRangeByIndexes(r, 0, 1, 5).getFormat().getFill().setColor("#EFF6FF");
    } else {
      const faultTally: Record<string, number> = {};
      const locTally: Record<string, number> = {};
      for (const inc of unitIncidents) {
        faultTally[inc.category] = (faultTally[inc.category] || 0) + 1;
        if (inc.location) locTally[inc.location] = (locTally[inc.location] || 0) + 1;
      }
      const sigFault = Object.keys(faultTally)
        .filter(f => f !== "Other" && f !== "No Description")
        .sort((a, b) => faultTally[b] - faultTally[a])[0] || "Other";
      const sigPct = Math.round(((faultTally[sigFault] || 0) / count) * 100);
      const topLoc = Object.keys(locTally).sort((a, b) => locTally[b] - locTally[a])[0] || "—";

      sheet.getRangeByIndexes(r, 0, 1, 5).setValues([[unit, count, sigFault, sigPct + "%", topLoc]]);
      sheet.getRangeByIndexes(r, 0, 1, 5).getFormat().getFill().setColor("#EFF6FF");

      // Amber highlight if 40%+ same fault with 10+ incidents — draws attention without labeling it
      if (sigPct >= 40 && count >= 10) {
        sheet.getRangeByIndexes(r, 2, 1, 2).getFormat().getFill().setColor("#FEF3C7");
      }
    }
    r++;
  }

  // Summary row
  const newFleetPct = totalIncidents > 0 ? Math.round((newFleetTotal / totalIncidents) * 100) : 0;
  sheet.getRangeByIndexes(r, 0, 1, 5).setValues([[
    "New Fleet Total", newFleetTotal, `${newFleetPct}% of all facility incidents`, "", ""
  ]]);
  sheet.getRangeByIndexes(r, 0, 1, 5).getFormat().getFont().setBold(true);
  sheet.getRangeByIndexes(r, 0, 1, 5).getFormat().getFill().setColor("#DBEAFE");
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

  // --- Chart: Top Three Reported AGVs ---
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
  top5Chart.setPosition(sheet.getCell(3, 6), sheet.getCell(18, 14));
  top5Chart.getLegend().setVisible(false);

  // Column widths
  sheet.getRange("A:A").getFormat().setColumnWidth(220);
  sheet.getRange("B:B").getFormat().setColumnWidth(90);
  sheet.getRange("C:C").getFormat().setColumnWidth(180);
  sheet.getRange("D:D").getFormat().setColumnWidth(70);

  // ── 4. Trend Log — append one block per day, never overwrite ────────────────
  // Each run writes a 3-row block (one per top AGV) sharing one background color.
  // A separator row divides blocks visually.
  // If the script is run more than once today, the log write is skipped.

  const RUN_COLORS = [
    "#E6F1FB", // soft blue
    "#EAF3DE", // soft green
    "#FAEEDA", // soft amber
    "#FAECE7", // soft coral
    "#EEEDFE", // soft purple
    "#E1F5EE", // soft teal
    "#FBEAF0", // soft pink
    "#F1EFE8", // soft gray
  ];

  let logSheet = workbook.getWorksheet("Monthly Trend Log");
  if (!logSheet) {
    logSheet = workbook.addWorksheet("Monthly Trend Log");

    // Header row — dark background, white text
    const headers = [
      "Date Run", "Summary",
      "Rank", "Unit", "Incidents (MTD)", "Signature Fault (MTD)",
      "Top Location", "Top Fault Category", "Other %"
    ];
    const headerRange = logSheet.getRangeByIndexes(0, 0, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.getFormat().getFont().setBold(true);
    headerRange.getFormat().getFont().setColor("#FFFFFF");
    headerRange.getFormat().getFill().setColor("#2C2C2A");

    // Column widths
    logSheet.getRange("A:A").getFormat().setColumnWidth(130);
    logSheet.getRange("B:B").getFormat().setColumnWidth(160);
    logSheet.getRange("C:C").getFormat().setColumnWidth(55);
    logSheet.getRange("D:D").getFormat().setColumnWidth(90);
    logSheet.getRange("E:E").getFormat().setColumnWidth(120);
    logSheet.getRange("F:F").getFormat().setColumnWidth(230);
    logSheet.getRange("G:G").getFormat().setColumnWidth(200);
    logSheet.getRange("H:H").getFormat().setColumnWidth(220);
    logSheet.getRange("I:I").getFormat().setColumnWidth(80);
  }

  // ── Once-per-day gate ──────────────────────────────────────────────────────
  // Read the last written date from column A. If it matches today, skip.
  const todayStr = new Date().toLocaleDateString();
  const logUsed = logSheet.getUsedRange();
  const existingRows = logUsed ? logUsed.getRowCount() : 1;

  let alreadyRanToday = false;
  if (existingRows > 1) {
    // Scan last few rows for today's date (blocks are 3 rows + 1 separator = 4 rows each)
    for (let check = existingRows - 1; check >= 1 && check >= existingRows - 5; check--) {
      const cellVal = String(logSheet.getRangeByIndexes(check, 0, 1, 1).getValue()).trim();
      if (cellVal && new Date(cellVal).toLocaleDateString() === todayStr) {
        alreadyRanToday = true;
        break;
      }
    }
  }

  if (alreadyRanToday) {
    console.log(`Trend Log — already ran today (${todayStr}), skipping write. Intelligence tab is up to date.`);
  } else {

    // ── Build top 3 with signature faults ─────────────────────────────────────
    const rankedUnits = Object.keys(unitCounts)
      .filter(u => u !== "Unknown")
      .sort((a, b) => unitCounts[b] - unitCounts[a])
      .slice(0, 3);

    // For each top unit, find their dominant fault category (signature fault)
    type UnitSig = { unit: string; count: number; sigFault: string; topLoc: string };
    const top3: UnitSig[] = rankedUnits.map(unit => {
      // Tally fault categories for this unit
      const faultTally: Record<string, number> = {};
      const locTally: Record<string, number> = {};
      for (const inc of incidents) {
        if (inc.unit !== unit) continue;
        faultTally[inc.category] = (faultTally[inc.category] || 0) + 1;
        if (inc.location) locTally[inc.location] = (locTally[inc.location] || 0) + 1;
      }
      const sigFault = Object.keys(faultTally)
        .filter(f => f !== "Other" && f !== "No Description")
        .sort((a, b) => faultTally[b] - faultTally[a])[0] || "Other";
      const topLoc = Object.keys(locTally)
        .sort((a, b) => locTally[b] - locTally[a])[0] || "—";
      return { unit, count: unitCounts[unit], sigFault, topLoc };
    });

    // Other % and top category
    const otherCount = categoryCounts["Other"] || 0;
    const otherPct = totalIncidents > 0 ? Math.round((otherCount / totalIncidents) * 100) : 0;
    const topCatForLog = Object.keys(categoryCounts)
      .filter(c => c !== "Other" && c !== "No Description")
      .sort((a, b) => categoryCounts[b] - categoryCounts[a])[0] || "N/A";

    // Run color — based on how many blocks exist already
    const blockCount = existingRows <= 1 ? 0 : Math.floor((existingRows - 1) / 4);
    const runColor = RUN_COLORS[blockCount % RUN_COLORS.length];

    // Summary lines for column B (only first row shows, others blank)
    const summaryLine = `${totalIncidents.toLocaleString()} incidents · ${avgPerDay}/day avg · ${uniqueDays} days · ${sheets.length - 2} shifts`;

    // Write 3 data rows — one per top AGV
    const startRow = existingRows;
    const ranks = ["#1", "#2", "#3"];
    for (let i = 0; i < 3; i++) {
      const u = top3[i];
      const rowData: (string | number)[] = [
        i === 0 ? new Date().toLocaleString() : "",   // date only on first row
        i === 0 ? summaryLine : "",                    // summary only on first row
        ranks[i],
        u ? u.unit : "—",
        u ? u.count : "—",
        u ? u.sigFault : "—",
        u ? u.topLoc : "—",                           // that AGV's own top location
        i === 0 ? topCatForLog : "",                   // top category on first row only
        i === 0 ? otherPct + "%" : "",                 // Other % on first row only
      ];
      const rowRange = logSheet.getRangeByIndexes(startRow + i, 0, 1, rowData.length);
      rowRange.setValues([rowData]);
      rowRange.getFormat().getFill().setColor(runColor);

      // Bold the date and rank cells
      if (i === 0) logSheet.getRangeByIndexes(startRow, 0, 1, 1).getFormat().getFont().setBold(true);
      logSheet.getRangeByIndexes(startRow + i, 2, 1, 1).getFormat().getFont().setBold(true);

      // Amber warning if Other % > 15
      if (i === 0 && otherPct > 15) {
        logSheet.getRangeByIndexes(startRow, 8, 1, 1).getFormat().getFill().setColor("#EF9F27");
        logSheet.getRangeByIndexes(startRow, 8, 1, 1).getFormat().getFont().setBold(true);
      }
    }

    // Separator row — thin gray line between blocks
    const sepRange = logSheet.getRangeByIndexes(startRow + 3, 0, 1, 9);
    sepRange.getFormat().getFill().setColor("#B4B2A9");

    console.log(`Trend Log — block written for ${todayStr}. Color: ${runColor}. Top 3: ${top3.map(u => u.unit).join(", ")}. Other: ${otherPct}%.`);
  }

  console.log(`Intelligence tab rebuilt — ${totalIncidents} incidents across ${uniqueDays} days.`);
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

  // ── AGF (LBW) ALARM TEXTS — must come before generic rules ───────────────
  // AGF 501-503 are dedicated to LBW and have a completely different fault vocabulary
  // from HBW AGVs. These rules anchor to the exact alarm texts from the AGF system.

  // Normal operations — completion/status events, not faults
  if (/^op complete$|^prelift completed$|autocharge attempt started|chargerstopinposition|autocharge operation complete without manual|getrack feedback validated after straylight|camera data being logged.*getrack|vehicle initialization started|stretchwrapdetection successfully cleared|driver detect/.test(d)) return "AGF Normal Op";

  // Prelift faults
  if (/prelift never executed|prelift did not complete.*unknown|prelift paused because vehicle is outside|prelift cancelled because vehicle is outside/.test(d)) return "AGF Prelift Fault";

  // Lift faults
  if (/lift not in position.*retry|lift pos retry due to lift pressure|lift height check failed/.test(d)) return "AGF Lift Fault";

  // Sideshift
  if (/sideshift limit reached to find center of rack/.test(d)) return "AGF Sideshift Limit";
  if (/retracting pallet.*pick between.*sideshift|retracting pallet.*pick beyond sideshift/.test(d)) return "AGF Sideshift Retract";

  // Pallet offset — critical (causes switch to manual)
  if (/pallet lateral offset exceeds sideshift tolerance|palletlatoffset.*outside of area.*switch to manual/.test(d)) return "AGF Pallet Offset - Critical";

  // Pallet offset — drop side
  if (/unable to achieve pallet lateral position on rack drop|pallet is \d+mm.*\d+mm offset from center of drop location|pallet is >75mm offset from center of drop/.test(d)) return "AGF Pallet Offset - Drop";

  // Pallet offset — pick side (forks)
  if (/pallet is \d+mm.*\d+mm offset from center of forks/.test(d)) return "AGF Pallet Offset - Pick";

  // Load faults
  if (/load did not release.*switch to manual|load did not release from flapper/.test(d)) return "AGF Load Did Not Release";
  if (/load push.*pull detected/.test(d)) return "AGF Load Push/Pull";

  // Camera / vision system
  if (/camera locateextents rack data exists with no upright|camera could not associate rack.*upright coverage|camera processing|camera is saving data for pick location blocked/.test(d)) return "AGF Camera Fault";

  // Pantograph
  if (/pantograph current overload detected|pantograph push detected/.test(d)) return "AGF Pantograph Overload";

  // Inventory / access blocked
  if (/inventory.*load beyond reach.*switch to manual|inventory.*access to rear pick blocked|getextents.*inventory.*front drop location occupied|can.t resume.*you need to clear all error/.test(d)) return "AGF Inventory Blocked";

  // AGF Bumper trips (different naming from HBW bumpers)
  if (/^front bumper$|^rear right bumper$|^rear left bumper$|^rear bumper$/.test(d)) return "AGF Bumper Trip";

  // Actuator faults (AGF-specific naming with Over/Under Position)
  if (/actuator #\d+ (over|under) position/.test(d)) return "AGF Actuator Fault";

  // Flapper
  if (/flapper tripped prematurely/.test(d)) return "AGF Flapper Trip";

  // EDM faults
  if (/instant off edm fault|delayed off edm fault/.test(d)) return "AGF EDM Fault";

  // Manual mode / interruption
  if (/agv operation interrupted by manual mode|auto.*manual state mismatch/.test(d)) return "AGF Manual Mode";

  // Drive overload
  if (/drive overload detected.*switch to manual/.test(d)) return "AGF Drive Overload";

  // Path deviation
  if (/slowing vehicle.*attempting to return vehicle to path|vehicle is \d+mm.*\d+mm offset from guidepath|unexpected vehicle direction/.test(d)) return "AGF Path Deviation";

  // Vehicle timeout
  if (/vehicle timed out.*operation cannot be completed/.test(d)) return "AGF Vehicle Timeout";

  // Charger faults
  if (/charger current below expectation|no current detected while charging.*switch to manual|no current detected.*positive charge not detected/.test(d)) return "AGF Charger Fault";
  if (/emergency-stop button pressed/.test(d)) return "E-Stop";
  if (/voltage below 75% of nominal/.test(d)) return "Battery/Low Power";
  if (/aplus move distance exceeds route move length/.test(d)) return "APlus - Move Without SetMove";

  // ── Engineer paraphrase variants found in May shift reports ──────────────
  // Engineers paraphrase AGF alarms rather than copying exact system text
  if (/prelift lateral offset exceeds sideshift|lateral offset exceeds sideshift/.test(d)) return "AGF Pallet Offset - Critical";
  if (/getextents.*inventory.*fron.*drop.*occupied|getextents.*inventory.*front.*drop.*occupied/.test(d)) return "AGF Inventory Blocked";
  if (/lift over position|lift under position|tilt not in position/.test(d)) return "AGF Lift Fault";
  if (/sat idle.*charger.*active order|wait condition.*charger.*order|in wait condition.*charger/.test(d)) return "Stuck/Blocked";
  if (/battery not reporting.*comm.*charger|battery not reporting.*charger/.test(d)) return "Battery/Low Power";

  // ── OFFICIAL ALARM TEXTS — anchored to exact system phrases ──────────────

  // Conveyor Is Still Transmitting — catches all typos across 9 months:
  // trnsmitting, trasmitting, transmittuing, coneyor, witout, tramitting
  if (/conveyor is still transmitting|tec conveyor is still transmitting|tec.*still transmit|conv.*still transmit|conveyor.*still transmit|coneyor.*still transmit|not sensing conv|still trnsmitting|still trasmitting/.test(d)) return "Conveyor Still Transmitting";

  // Handshakes — official: "Left/Right Handshake Timeout 5min"
  // Catches: LHST, RHST, LHT, RHT, RHSTO + all typo variants
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
  // Guid — catches all spellings: Travelled, Traveked, wihtout, and drift fault
  if (/guid.*travel|traveled too far without correction|travelled too far|travel.*too far|drift fault/.test(d)) return "Guid - Traveled Too Far";
  // NAV — standalone "NAV" or "AGV 28 NAV" — bare nav faults with no further detail
  if (/nodelock|node lock|lost nav|not on node|node-lock|\bno nav\b|loss of nav|\bnav\b/.test(d)) return "Navigation/Nodelock";

  // Alignment Reflector Not Detected — catches: "alignment rail not in position"
  if (/alignment reflector not detected|alignment reflector|alignment sensor did not turn on|reflector not detected|alignment rail not in position/.test(d)) return "Alignment Reflector Not Detected";

  // No Auto-Mode Guidesafe
  if (/no auto-mode guidesafe|no auto mode guidesafe|guidesafe|guidsafe/.test(d)) return "No Auto-Mode Guidesafe";

  // Transfer Timeouts — three distinct official alarms
  // Catches system typos: "Complet", "Compete", and engineer typos: Tansfer, Tanser
  if (/completpicktransfer|completepicktransfer|complet.*pick.*transfer|pick transfer t\.?o|pick.*transfer.*timeout|completepicktansfer|completepicktanser/.test(d)) return "CompletePickTransfer Timeout";
  if (/competedrop|completedrop.*transfer|complet.*drop.*transfer|drop transfer t\.?o|complete drop transfer timeout|complete drp transfer/.test(d)) return "CompleteDropTransfer Timeout";
  if (/startpicktransfer|start.*pick.*transfer|startpick.*t\.?o/.test(d)) return "StartPickTransfer Timeout";
  // Conv-level transfer failures (e.g. "Conv 51 failed to receive transfer from AGV 14")
  if (/conv.*failed to receive transfer|failed to receive transfer from agv/.test(d)) return "Transfer Timeout";
  if (/midpick transfer|transfer timeout|transfer fault|complete transfer t\.?o|conveyor timeout/.test(d)) return "Transfer Timeout";

  // Unachievable Distance CMD — catches all misspellings across 9 months
  if (/unachievable distance|unahievable distance|unacheivable distance|unecheivable distance|unaheivable distance|unechievable distance|distance cmd/.test(d)) return "Unachievable Distance CMD";

  // APlus — catches: "auto shutdown pending / aplus"
  if (/aplus requested vehicle shutdown|aplus.*shutdown|aplus requested|auto shutdown pending.*aplus/.test(d)) return "APlus - Requested Shutdown";
  if (/aplus.*move without setmove|move without setmove|aplus.*setmove/.test(d)) return "APlus - Move Without SetMove";

  // Manual Load Transfer
  if (/faulting vehicle during manual load transfer|manual load transfer/.test(d)) return "Manual Load Transfer Fault";

  // WAMAS Order Fault
  if (/wamas|failed to receive order/.test(d)) return "WAMAS Order Fault";

  // PNG Positioning Invalid
  if (/parameters for positioning at png|positioning at png conveyors/.test(d)) return "PNG Positioning Invalid";

  // Inventory Not On Conveyor
  if (/inventory does not exist on conveyor|vehicle picking.*inventory/.test(d)) return "Inventory Not On Conveyor";

  // Pendant Fault
  if (/nonzero demands when pendant|pendant activated/.test(d)) return "Pendant Fault";

  // Start Node Position Fault
  if (/too far from start node|start node of commanded move/.test(d)) return "Start Node Position Fault";

  // PassEdge Tripped
  if (/passedge tripped|pass edge tripped|conveyor passedge|passedge|pass edge|overshoot causing pass edge/.test(d)) return "PassEdge Tripped";

  // Momentary Power Loss
  if (/momentary power loss|recovered from.*power loss/.test(d)) return "Momentary Power Loss";

  // Host faults — catches: "Lost Network connection"
  if (/invalid host parameters/.test(d)) return "Invalid Host Parameters";
  if (/lost host|host connection|nvc.*crash|nvc.*closed with error|lost network connection/.test(d)) return "Lost Host Connection";

  // Recovery Failed
  if (/recovery failed|recovery/.test(d)) return "Recovery Failed";

  // No Movement Range Match — catches standalone "no movement range"
  if (/no movement range matched|movement range matches|movement range matched|no movement range/.test(d)) return "No Movement Range Match";

  // Traction — two distinct official alarms
  if (/loss of traction feedback/.test(d)) return "Loss of Traction Feedback";
  if (/unexpected traction feedback|traction/.test(d)) return "Unexpected Traction Feedback";

  // ── NON-OFFICIAL BUT WELL-UNDERSTOOD FAULTS ──────────────────────────────

  // Scanner
  if (/scanner data|no scanner|scanner could not read/.test(d)) return "Nav - No Scanner Data";

  // Battery / power — catches: "died on charger battery offline"
  if (/low battery|low batt level|battery shutdown|batt\.?\s*monitor|voltage below.*nominal|battery not.*comm.*charger|battery not.*reporting.*charger|battery.*offline|died on charger/.test(d)) return "Battery/Low Power";

  // Comms / Flexisoft
  if (/flexi|no comms|can message/.test(d)) return "Comms/Flexisoft";

  // TC ADS Error
  if (/ads error|tc ads/.test(d)) return "TC ADS Error";

  // Brake / Amplifier
  if (/amplifier fault|em brake|braking distance/.test(d)) return "Brake/Amplifier Fault";

  // HU Drive Inverter Lag — catches: "drift inverter"
  if (/drive inverter|lag distance|drift inverter/.test(d)) return "HU Drive Inverter Lag";

  // Actuator — catches "Acuator" typo
  if (/actu?ator/.test(d)) return "Actuator Fault";

  // Bumper — catches "Bymper" and "Numper" typos
  if (/bu?mper|numper/.test(d)) return "Bumper Trip";

  // Stack Hitting Guard Rail — location issue (conveyor guard physically blocking the load)
  // Triggers on: "can't load due to conveyor guard blocking stack", "guard blocking stack"
  if (/conveyor guard blocking|guard blocking stack|can.t load due to.*guard/.test(d)) return "Stack Hitting Guard Rail";

  // Guard Rail Bumper Trip — vehicle issue (AGV drove too close to guard rail)
  // Triggers on: "bumper trip on yellow guardrail", "drove too close to guard rail"
  if (/bumper.*guard rail|bumper.*guardrail|guard rail.*bumper|guardrail.*bumper|too close to guard|drove too close/.test(d)) return "Guard Rail Bumper Trip";

  // Load/Stack physically crooked or misloaded — new from Aug
  if (/went crooked|crooked.*pick|hit the agv|realigned.*agv|wasn.t lined up|binstack|wasn.t lined/.test(d)) return "Load/Stack Alignment";

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

  // Stuck / Blocked
  if (/stuck idle|deadlock|wait condition|full infeed|stuck in idle|went idle|idle with active|back.?up|causing a back|\bin idle\b|died in|died while|shut down in middle|just turned off|stuck evaluating|waiting for pick|waiting for drop|waiting to drop|waiting to pick/.test(d)) return "Stuck/Blocked";

  // AGF-specific load faults
  if (/load did not release|load defect fault|load push.pull|load beyond reach|getextents|getrack|pantograph|tile not in position/.test(d)) return "AGF Load Fault";

  // FRBT / FLBT / RBT
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

  // Safety
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
