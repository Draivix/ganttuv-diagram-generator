#!/usr/bin/env -S npx tsx
/**
 * Generates the AutoERP "Ganttův diagram" Excel template.
 *
 * Sheets:
 *   1. "Ganttův diagram" — main view: data table + 26-week timeline with formula-driven bars
 *   2. "Vizualizace"     — stacked bar chart simulating Gantt visualization
 *   3. "Návod"           — usage instructions
 *   4. "O šabloně"       — branding, license, AutoERP CTA
 *
 * Output: public/templates/ganttuv-diagram-vzor.xlsx
 */
import ExcelJS from 'exceljs';
import { mkdirSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);
const ROOT = join(__dirname, '..');
const OUT_DIR = join(ROOT, 'public', 'templates');
const OUT_FILE = join(OUT_DIR, 'ganttuv-diagram-vzor.xlsx');

mkdirSync(OUT_DIR, { recursive: true });

/* -------------------------------------------------------------------------- */
/*  Brand palette (ARGB)                                                      */
/* -------------------------------------------------------------------------- */

const C = {
  NAVY:        'FF1E3A5F',
  NAVY_DARK:   'FF14283F',
  NAVY_SOFT:   'FFE2E8F0',
  WHITE:       'FFFFFFFF',
  TEXT_DARK:   'FF1E293B',
  TEXT_BODY:   'FF334155',
  TEXT_MUTED:  'FF64748B',
  BG_PAGE:     'FFF8FAFC',
  BG_ALT:      'FFF1F5F9',
  BORDER:      'FFE2E8F0',
  BORDER_HARD: 'FFCBD5E1',
  ACCENT:      'FFF59E0B',
  ACCENT_SOFT: 'FFFEF3C7',
  ACCENT_HOVER:'FFD97706',
  SUCCESS:     'FF10B981',
  SUCCESS_SOFT:'FFD1FAE5',
  AMBER:       'FFFBBF24',
  AMBER_SOFT:  'FFFEF9C3',
  RED:         'FFEF4444',
  RED_SOFT:    'FFFEE2E2',
  PROGRESS:    'FFF59E0B',
  BAR_FILL:    'FF2563EB',
  ROW_ALT:     'FFFAFAFA',
};

const FONT = 'Segoe UI';

/* -------------------------------------------------------------------------- */
/*  Sample data — "Implementace ERP" project                                   */
/* -------------------------------------------------------------------------- */

interface TaskRow {
  id: string;
  task: string;
  owner: string;
  start: string;     // dd.mm.yyyy
  end: string;       // dd.mm.yyyy
  progress: number;  // 0..1
  status: 'Nezahájeno' | 'Probíhá' | 'Hotovo' | 'Zpožděno';
  dependency: string;
}

const SAMPLE: TaskRow[] = [
  // Příprava
  { id: '1.1', task: 'Kick-off meeting + zadání projektu',         owner: 'David Strejc',      start: '04.05.2026', end: '05.05.2026', progress: 1.00, status: 'Hotovo',     dependency: '—'    },
  { id: '1.2', task: 'Sběr požadavků od klíčových uživatelů',       owner: 'Anna Nováková',     start: '06.05.2026', end: '15.05.2026', progress: 1.00, status: 'Hotovo',     dependency: '1.1'  },
  { id: '1.3', task: 'Mapování stávajících procesů',                 owner: 'Petr Svoboda',      start: '11.05.2026', end: '22.05.2026', progress: 0.85, status: 'Probíhá',    dependency: '1.2'  },
  // Analýza
  { id: '2.1', task: 'Analýza dat z Pohody (export, čištění)',       owner: 'Tomáš Dvořák',      start: '18.05.2026', end: '29.05.2026', progress: 0.60, status: 'Probíhá',    dependency: '1.2'  },
  { id: '2.2', task: 'Návrh datového modelu a integrací',            owner: 'Jana Procházková',  start: '25.05.2026', end: '05.06.2026', progress: 0.30, status: 'Probíhá',    dependency: '2.1'  },
  { id: '2.3', task: 'Schválení specifikace zákazníkem',             owner: 'Klient',            start: '08.06.2026', end: '12.06.2026', progress: 0.00, status: 'Nezahájeno', dependency: '2.2'  },
  // Implementace
  { id: '3.1', task: 'Konfigurace AutoERP — moduly a uživatelé',     owner: 'Tomáš Dvořák',      start: '15.06.2026', end: '26.06.2026', progress: 0.00, status: 'Nezahájeno', dependency: '2.3'  },
  { id: '3.2', task: 'Migrace dat (zákazníci, sklad, faktury)',      owner: 'Anna Nováková',     start: '22.06.2026', end: '03.07.2026', progress: 0.00, status: 'Nezahájeno', dependency: '3.1'  },
  { id: '3.3', task: 'Integrace s e-shopem (Shoptet)',                owner: 'Petr Svoboda',      start: '29.06.2026', end: '10.07.2026', progress: 0.00, status: 'Nezahájeno', dependency: '3.1'  },
  { id: '3.4', task: 'Integrace s bankou (FIO API)',                  owner: 'Jana Procházková',  start: '06.07.2026', end: '17.07.2026', progress: 0.00, status: 'Nezahájeno', dependency: '3.1'  },
  { id: '3.5', task: 'Tisková sestava faktur a dodacích listů',       owner: 'Tomáš Dvořák',      start: '13.07.2026', end: '24.07.2026', progress: 0.00, status: 'Nezahájeno', dependency: '3.2'  },
  // Testování
  { id: '4.1', task: 'UAT — testování klíčových uživatelů',           owner: 'Klient + Anna',     start: '27.07.2026', end: '07.08.2026', progress: 0.00, status: 'Nezahájeno', dependency: '3.5'  },
  { id: '4.2', task: 'Oprava nálezů z UAT',                            owner: 'Tým AutoERP',       start: '03.08.2026', end: '14.08.2026', progress: 0.00, status: 'Nezahájeno', dependency: '4.1'  },
  { id: '4.3', task: 'Školení uživatelů (5 sezení)',                   owner: 'Anna Nováková',     start: '10.08.2026', end: '21.08.2026', progress: 0.00, status: 'Nezahájeno', dependency: '4.2'  },
  // Spuštění
  { id: '5.1', task: 'Ostrý start — go-live',                          owner: 'David Strejc',      start: '24.08.2026', end: '28.08.2026', progress: 0.00, status: 'Nezahájeno', dependency: '4.3'  },
  { id: '5.2', task: 'Hyper-care support (první 2 týdny)',             owner: 'Tým AutoERP',       start: '31.08.2026', end: '11.09.2026', progress: 0.00, status: 'Nezahájeno', dependency: '5.1'  },
  { id: '5.3', task: 'Vyhodnocení projektu + retrospektiva',           owner: 'David Strejc',      start: '14.09.2026', end: '18.09.2026', progress: 0.00, status: 'Nezahájeno', dependency: '5.2'  },
];

/* -------------------------------------------------------------------------- */
/*  Helpers                                                                    */
/* -------------------------------------------------------------------------- */

function parseDate(czDate: string): Date {
  // dd.mm.yyyy → Date (UTC noon to avoid timezone drift)
  const [d, m, y] = czDate.split('.').map((p) => parseInt(p.trim(), 10));
  return new Date(Date.UTC(y, m - 1, d, 12, 0, 0));
}

function addDays(d: Date, days: number): Date {
  const r = new Date(d);
  r.setUTCDate(r.getUTCDate() + days);
  return r;
}

function fmtCz(d: Date): string {
  const dd = String(d.getUTCDate()).padStart(2, '0');
  const mm = String(d.getUTCMonth() + 1).padStart(2, '0');
  return `${dd}.${mm}.${d.getUTCFullYear()}`;
}

const FONTS = {
  title:    { name: FONT, size: 18, bold: true, color: { argb: C.WHITE } },
  subtitle: { name: FONT, size: 10, italic: true, color: { argb: C.WHITE } },
  h2:       { name: FONT, size: 13, bold: true, color: { argb: C.NAVY } },
  h3:       { name: FONT, size: 11, bold: true, color: { argb: C.TEXT_DARK } },
  body:     { name: FONT, size: 10, color: { argb: C.TEXT_BODY } },
  muted:    { name: FONT, size: 9, italic: true, color: { argb: C.TEXT_MUTED } },
  tableHead:{ name: FONT, size: 10, bold: true, color: { argb: C.WHITE } },
  cell:     { name: FONT, size: 10, color: { argb: C.TEXT_DARK } },
  accent:   { name: FONT, size: 11, bold: true, color: { argb: C.WHITE } },
} as const;

const fillSolid = (argb: string) => ({ type: 'pattern' as const, pattern: 'solid' as const, fgColor: { argb } });
const thinBorder = { style: 'thin' as const, color: { argb: C.BORDER } };
const allThinBorders = { top: thinBorder, left: thinBorder, bottom: thinBorder, right: thinBorder };

/* -------------------------------------------------------------------------- */
/*  Workbook generation                                                        */
/* -------------------------------------------------------------------------- */

async function build() {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'AutoERP — Apertia Tech s.r.o.';
  wb.lastModifiedBy = 'AutoERP';
  wb.created = new Date();
  wb.modified = new Date();
  wb.company = 'AutoERP';
  wb.title = 'Ganttův diagram — projektový plán';
  wb.description = 'Šablona Ganttova diagramu — vytvořeno v AutoERP (autoerp.cz)';
  wb.keywords = 'Gantt, projektový plán, harmonogram, Excel, AutoERP';

  /* ──────────────────────────────────────────────────────────────────────── */
  /*  SHEET 1 — Ganttův diagram                                                */
  /* ──────────────────────────────────────────────────────────────────────── */

  const sh = wb.addWorksheet('Ganttův diagram', {
    views: [{ state: 'frozen', xSplit: 2, ySplit: 5, showGridLines: false }],
    pageSetup: {
      orientation: 'landscape',
      paperSize: 9,           // A4
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0,
      margins: { left: 0.3, right: 0.3, top: 0.5, bottom: 0.5, header: 0.2, footer: 0.2 },
      printTitlesRow: '5:5',
    },
    headerFooter: {
      oddFooter: '&L&"Segoe UI,Italic"&8&K64748B Vytvořeno v AutoERP — autoerp.cz/projektove-rizeni&R&8 Strana &P z &N',
    },
  });

  // Determine column count: 9 data cols + 26 timeline weeks
  const NUM_WEEKS = 26;
  const dataColumns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'];
  const timelineFirstColIdx = 10;                          // J
  const timelineLastColIdx = timelineFirstColIdx + NUM_WEEKS - 1;
  const lastDataRow = 5 + SAMPLE.length;                   // header at row 5, data starts row 6

  // Row 1-3 — branded header band
  sh.mergeCells('A1:AI3');
  const hdrCell = sh.getCell('A1');
  hdrCell.value = {
    richText: [
      { text: 'Ganttův diagram projektu\n', font: FONTS.title },
      { text: 'Vytvořeno v AutoERP — autoerp.cz/projektove-rizeni', font: FONTS.subtitle },
    ],
  };
  hdrCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  hdrCell.fill = fillSolid(C.NAVY);
  sh.getRow(1).height = 26;
  sh.getRow(2).height = 26;
  sh.getRow(3).height = 22;

  // Row 4 — spacer
  sh.getRow(4).height = 8;

  // Row 5 — Header row
  const headers = [
    { col: 'A', label: 'ID',         width: 6  },
    { col: 'B', label: 'Úkol',       width: 38 },
    { col: 'C', label: 'Odpovědný',   width: 18 },
    { col: 'D', label: 'Začátek',     width: 12 },
    { col: 'E', label: 'Konec',       width: 12 },
    { col: 'F', label: 'Dnů',         width: 7  },
    { col: 'G', label: 'Pokrok %',    width: 10 },
    { col: 'H', label: 'Status',      width: 13 },
    { col: 'I', label: 'Závislost',   width: 10 },
  ];
  headers.forEach((h, i) => {
    const cell = sh.getCell(`${h.col}5`);
    cell.value = h.label;
    cell.fill = fillSolid(C.NAVY);
    cell.font = FONTS.tableHead;
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    cell.border = {
      top:    { style: 'thin', color: { argb: C.NAVY_DARK } },
      bottom: { style: 'medium', color: { argb: C.NAVY_DARK } },
      left:   i === 0 ? { style: 'thin', color: { argb: C.NAVY_DARK } } : thinBorder,
      right:  thinBorder,
    };
    sh.getColumn(h.col).width = h.width;
  });

  // Determine project span and weekly grid origin (Monday on/before earliest start)
  const allStarts = SAMPLE.map((t) => parseDate(t.start));
  const earliest = new Date(Math.min(...allStarts.map((d) => d.getTime())));
  // shift to nearest Monday on/before
  const dayOfWeek = (earliest.getUTCDay() + 6) % 7; // 0=Mon..6=Sun
  const gridStart = addDays(earliest, -dayOfWeek);

  // Timeline header row — fill J5..AI5 with weekly date labels
  for (let w = 0; w < NUM_WEEKS; w++) {
    const colIdx = timelineFirstColIdx + w;
    const weekStart = addDays(gridStart, w * 7);
    const cell = sh.getCell(5, colIdx);
    cell.value = `T${w + 1}\n${fmtCz(weekStart).slice(0, 5)}`;
    cell.fill = fillSolid(C.NAVY);
    cell.font = { name: FONT, size: 8, bold: true, color: { argb: C.WHITE } };
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    cell.border = {
      top:    { style: 'thin', color: { argb: C.NAVY_DARK } },
      bottom: { style: 'medium', color: { argb: C.NAVY_DARK } },
      left:   thinBorder,
      right:  thinBorder,
    };
    sh.getColumn(colIdx).width = 4.5;
  }
  sh.getRow(5).height = 32;

  // Hidden helper row — store week-start serial dates in row 4 (cols J..AI)
  // We'll use them in the IF formulas. Hide this row visually but keep it usable.
  for (let w = 0; w < NUM_WEEKS; w++) {
    const colIdx = timelineFirstColIdx + w;
    const weekStart = addDays(gridStart, w * 7);
    const cell = sh.getCell(4, colIdx);
    cell.value = weekStart;
    cell.numFmt = 'dd.mm.yyyy';
    cell.font = { name: FONT, size: 1, color: { argb: C.NAVY } };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
  }
  sh.getRow(4).hidden = true;

  // Data rows
  SAMPLE.forEach((t, i) => {
    const r = 6 + i;
    const row = sh.getRow(r);
    row.height = 22;
    // ID
    const cId = sh.getCell(r, 1);
    cId.value = t.id;
    cId.font = { name: FONT, size: 10, bold: true, color: { argb: C.NAVY } };
    cId.alignment = { horizontal: 'center', vertical: 'middle' };
    // Task
    const cTask = sh.getCell(r, 2);
    cTask.value = t.task;
    cTask.font = FONTS.cell;
    cTask.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true, indent: 1 };
    // Owner
    const cOwner = sh.getCell(r, 3);
    cOwner.value = t.owner;
    cOwner.font = FONTS.cell;
    cOwner.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
    // Start / End — real dates
    const cStart = sh.getCell(r, 4);
    cStart.value = parseDate(t.start);
    cStart.numFmt = 'dd.mm.yyyy';
    cStart.font = FONTS.cell;
    cStart.alignment = { horizontal: 'center', vertical: 'middle' };
    const cEnd = sh.getCell(r, 5);
    cEnd.value = parseDate(t.end);
    cEnd.numFmt = 'dd.mm.yyyy';
    cEnd.font = FONTS.cell;
    cEnd.alignment = { horizontal: 'center', vertical: 'middle' };
    // Days — formula
    const cDays = sh.getCell(r, 6);
    cDays.value = { formula: `E${r}-D${r}+1` };
    cDays.font = FONTS.cell;
    cDays.alignment = { horizontal: 'center', vertical: 'middle' };
    cDays.numFmt = '0';
    // Progress
    const cProg = sh.getCell(r, 7);
    cProg.value = t.progress;
    cProg.numFmt = '0%';
    cProg.font = FONTS.cell;
    cProg.alignment = { horizontal: 'center', vertical: 'middle' };
    // Status
    const cStatus = sh.getCell(r, 8);
    cStatus.value = t.status;
    cStatus.font = FONTS.cell;
    cStatus.alignment = { horizontal: 'center', vertical: 'middle' };
    // Dependency
    const cDep = sh.getCell(r, 9);
    cDep.value = t.dependency;
    cDep.font = { ...FONTS.cell, color: { argb: C.TEXT_MUTED } };
    cDep.alignment = { horizontal: 'center', vertical: 'middle' };

    // Apply borders + alternating row fill on data cols
    const altFill = i % 2 === 0 ? null : C.ROW_ALT;
    for (let c = 1; c <= 9; c++) {
      const cell = sh.getCell(r, c);
      cell.border = allThinBorders;
      if (altFill) cell.fill = fillSolid(altFill);
    }

    // Timeline bars — formula: if week intersects task date range, fill cell
    for (let w = 0; w < NUM_WEEKS; w++) {
      const colIdx = timelineFirstColIdx + w;
      const cell = sh.getCell(r, colIdx);
      const colLetter = sh.getColumn(colIdx).letter;
      // Bar visible if (weekStart <= taskEnd) AND (weekEnd >= taskStart)
      cell.value = {
        formula: `IF(AND(${colLetter}$4<=$E${r},${colLetter}$4+6>=$D${r}),"",""))`.replace(',""))', '),"","")'),
      };
      // Set placeholder; conditional formatting handles actual fill
      cell.value = { formula: `IF(AND(${colLetter}$4<=$E${r},${colLetter}$4+6>=$D${r}),1,0)` };
      cell.numFmt = ';;;';   // hide the 1/0 number visually
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.border = thinBorder
        ? { top: thinBorder, bottom: thinBorder, left: thinBorder, right: thinBorder }
        : undefined;
      if (altFill) cell.fill = fillSolid(altFill);
    }
  });

  // Conditional formatting on timeline area — fill cell where formula =1
  const timelineFirstColLetter = sh.getColumn(timelineFirstColIdx).letter;
  const timelineLastColLetter = sh.getColumn(timelineLastColIdx).letter;
  const timelineRange = `${timelineFirstColLetter}6:${timelineLastColLetter}${lastDataRow}`;

  // Bar fill — when value = 1, paint navy
  sh.addConditionalFormatting({
    ref: timelineRange,
    rules: [
      {
        type: 'cellIs',
        operator: 'equal',
        formulae: ['1'],
        priority: 2,
        style: {
          fill: fillSolid(C.NAVY),
          font: { color: { argb: C.NAVY } },
          border: {
            top:    { style: 'thin', color: { argb: C.NAVY } },
            bottom: { style: 'thin', color: { argb: C.NAVY } },
            left:   { style: 'thin', color: { argb: C.NAVY } },
            right:  { style: 'thin', color: { argb: C.NAVY } },
          },
        },
      },
    ],
  });

  // Progress overlay — paint accent (amber) where the week is also covered by progress fraction
  // Progress covers from start..start + (end-start+1)*progress days. Use a separate CF.
  sh.addConditionalFormatting({
    ref: timelineRange,
    rules: [
      {
        type: 'expression',
        // formulae are absolute-ref-aware in conditional formatting in ExcelJS
        formulae: [
          // J6 area: weekStart >= taskStart AND weekStart <= taskStart + (taskEnd-taskStart+1)*progress - 1
          `AND(${timelineFirstColLetter}$4<=$E6,${timelineFirstColLetter}$4+6>=$D6,${timelineFirstColLetter}$4<=$D6+($E6-$D6+1)*$G6-1)`,
        ],
        priority: 1,
        style: {
          fill: fillSolid(C.ACCENT),
          font: { color: { argb: C.ACCENT } },
        },
      },
    ],
  });

  // Conditional formatting — Pokrok % column G (color scale red→amber→green)
  sh.addConditionalFormatting({
    ref: `G6:G${lastDataRow}`,
    rules: [
      {
        type: 'colorScale',
        priority: 5,
        cfvo: [
          { type: 'num', value: 0 },
          { type: 'num', value: 0.5 },
          { type: 'num', value: 1 },
        ],
        color: [
          { argb: C.RED_SOFT },
          { argb: C.AMBER_SOFT },
          { argb: C.SUCCESS_SOFT },
        ],
      },
    ],
  });

  // Status column data validation + conditional formatting
  for (let r = 6; r <= lastDataRow; r++) {
    sh.getCell(`H${r}`).dataValidation = {
      type: 'list',
      allowBlank: false,
      formulae: ['"Nezahájeno,Probíhá,Hotovo,Zpožděno"'],
      showErrorMessage: true,
      errorStyle: 'warning',
      errorTitle: 'Neplatný status',
      error: 'Vyberte z: Nezahájeno / Probíhá / Hotovo / Zpožděno.',
    };
  }

  const statusCFRules = [
    { value: 'Nezahájeno', fill: C.NAVY_SOFT,    font: C.TEXT_BODY  },
    { value: 'Probíhá',    fill: C.AMBER_SOFT,   font: C.ACCENT_HOVER },
    { value: 'Hotovo',     fill: C.SUCCESS_SOFT, font: C.SUCCESS    },
    { value: 'Zpožděno',   fill: C.RED_SOFT,     font: C.RED        },
  ];
  for (const rule of statusCFRules) {
    sh.addConditionalFormatting({
      ref: `H6:H${lastDataRow}`,
      rules: [
        {
          type: 'cellIs',
          operator: 'equal',
          formulae: [`"${rule.value}"`],
          priority: 3,
          style: {
            fill: fillSolid(rule.fill),
            font: { name: FONT, size: 10, bold: true, color: { argb: rule.font } },
          },
        },
      ],
    });
  }

  // Summary row — bold totals at bottom
  const sumRow = lastDataRow + 2;
  const cSumLabel = sh.getCell(`A${sumRow}`);
  sh.mergeCells(`A${sumRow}:E${sumRow}`);
  cSumLabel.value = 'Souhrn projektu';
  cSumLabel.font = { name: FONT, size: 11, bold: true, color: { argb: C.NAVY } };
  cSumLabel.alignment = { horizontal: 'right', vertical: 'middle', indent: 1 };
  cSumLabel.fill = fillSolid(C.BG_ALT);

  const cSumDays = sh.getCell(`F${sumRow}`);
  cSumDays.value = { formula: `SUM(F6:F${lastDataRow})` };
  cSumDays.font = { name: FONT, size: 11, bold: true, color: { argb: C.NAVY } };
  cSumDays.alignment = { horizontal: 'center', vertical: 'middle' };
  cSumDays.fill = fillSolid(C.BG_ALT);
  cSumDays.numFmt = '0';

  const cSumProg = sh.getCell(`G${sumRow}`);
  cSumProg.value = { formula: `SUMPRODUCT(F6:F${lastDataRow},G6:G${lastDataRow})/SUM(F6:F${lastDataRow})` };
  cSumProg.font = { name: FONT, size: 11, bold: true, color: { argb: C.ACCENT_HOVER } };
  cSumProg.alignment = { horizontal: 'center', vertical: 'middle' };
  cSumProg.fill = fillSolid(C.BG_ALT);
  cSumProg.numFmt = '0%';

  const cSumNote = sh.getCell(`H${sumRow}`);
  sh.mergeCells(`H${sumRow}:I${sumRow}`);
  cSumNote.value = 'celkem dnů / vážený pokrok';
  cSumNote.font = { name: FONT, size: 9, italic: true, color: { argb: C.TEXT_MUTED } };
  cSumNote.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
  cSumNote.fill = fillSolid(C.BG_ALT);

  // Footer note row
  const footRow = sumRow + 2;
  sh.mergeCells(`A${footRow}:I${footRow}`);
  const cFoot = sh.getCell(`A${footRow}`);
  cFoot.value = 'Tip: Pro automatický výpočet harmonogramu, sledování zakázek a kapacit zdrojů vyzkoušejte AutoERP — autoerp.cz/projektove-rizeni';
  cFoot.font = { name: FONT, size: 9, italic: true, color: { argb: C.NAVY } };
  cFoot.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
  cFoot.fill = fillSolid(C.ACCENT_SOFT);
  sh.getRow(footRow).height = 24;

  /* ──────────────────────────────────────────────────────────────────────── */
  /*  SHEET 2 — Vizualizace (chart)                                            */
  /* ──────────────────────────────────────────────────────────────────────── */
  // ExcelJS chart support is limited — we add a "rasterized" Gantt visualization
  // using cell-fill bars so the visualization is robust across Excel/LibreOffice/Sheets.

  const sh2 = wb.addWorksheet('Vizualizace', {
    views: [{ state: 'frozen', xSplit: 1, ySplit: 4, showGridLines: false }],
    pageSetup: { orientation: 'landscape', paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
  });

  // Header band
  sh2.mergeCells('A1:BA3');
  const h2 = sh2.getCell('A1');
  h2.value = {
    richText: [
      { text: 'Ganttova vizualizace — denní rozlišení\n', font: FONTS.title },
      { text: 'Každý sloupec = 1 den. Modrá = úkol. Žlutá = pokrok.', font: FONTS.subtitle },
    ],
  };
  h2.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  h2.fill = fillSolid(C.NAVY);
  sh2.getRow(1).height = 22;
  sh2.getRow(2).height = 22;
  sh2.getRow(3).height = 18;

  // Daily timeline — calculate total span
  const allEnds = SAMPLE.map((t) => parseDate(t.end));
  const latest = new Date(Math.max(...allEnds.map((d) => d.getTime())));
  const totalDays = Math.ceil((latest.getTime() - gridStart.getTime()) / 86400000) + 1;
  // Cap at 140 days (~20 weeks) for visual sanity
  const dayCount = Math.min(totalDays, 140);

  // Row 4 — task name col + day labels
  sh2.getCell('A4').value = 'Úkol';
  sh2.getCell('A4').fill = fillSolid(C.NAVY);
  sh2.getCell('A4').font = FONTS.tableHead;
  sh2.getCell('A4').alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
  sh2.getColumn('A').width = 38;

  for (let d = 0; d < dayCount; d++) {
    const colIdx = 2 + d;
    const day = addDays(gridStart, d);
    const cell = sh2.getCell(4, colIdx);
    // Show only first day of week
    const dow = (day.getUTCDay() + 6) % 7;
    cell.value = dow === 0 ? fmtCz(day).slice(0, 5) : '';
    cell.fill = fillSolid(C.NAVY);
    cell.font = { name: FONT, size: 7, bold: true, color: { argb: C.WHITE } };
    cell.alignment = { horizontal: 'left', vertical: 'middle' };
    sh2.getColumn(colIdx).width = 1.6;
  }
  sh2.getRow(4).height = 26;

  // Bars
  SAMPLE.forEach((t, i) => {
    const r = 5 + i;
    const taskCell = sh2.getCell(r, 1);
    taskCell.value = `${t.id}  ${t.task}`;
    taskCell.font = { name: FONT, size: 9, color: { argb: C.TEXT_DARK } };
    taskCell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: false, indent: 1 };
    taskCell.border = { bottom: thinBorder };
    sh2.getRow(r).height = 16;

    const ts = parseDate(t.start);
    const te = parseDate(t.end);
    const startDay = Math.round((ts.getTime() - gridStart.getTime()) / 86400000);
    const endDay = Math.round((te.getTime() - gridStart.getTime()) / 86400000);
    const taskLen = endDay - startDay + 1;
    const progressLen = Math.round(taskLen * t.progress);

    for (let d = 0; d < dayCount; d++) {
      const colIdx = 2 + d;
      const cell = sh2.getCell(r, colIdx);
      cell.value = '';
      const inTask = d >= startDay && d <= endDay;
      const inProgress = d >= startDay && d < startDay + progressLen;
      if (inProgress) {
        cell.fill = fillSolid(C.ACCENT);
      } else if (inTask) {
        cell.fill = fillSolid(C.NAVY);
      }
      cell.border = { right: { style: 'hair', color: { argb: C.BORDER } } };
    }
  });

  // Legend below
  const legRow = 5 + SAMPLE.length + 2;
  sh2.getCell(`A${legRow}`).value = 'Legenda:';
  sh2.getCell(`A${legRow}`).font = FONTS.h3;

  sh2.getCell(`B${legRow}`).fill = fillSolid(C.NAVY);
  sh2.mergeCells(`B${legRow}:D${legRow}`);
  sh2.getCell(`B${legRow}`).value = 'Úkol (plánováno)';
  sh2.getCell(`B${legRow}`).font = { name: FONT, size: 9, bold: true, color: { argb: C.WHITE } };
  sh2.getCell(`B${legRow}`).alignment = { horizontal: 'center', vertical: 'middle' };

  sh2.getCell(`F${legRow}`).fill = fillSolid(C.ACCENT);
  sh2.mergeCells(`F${legRow}:H${legRow}`);
  sh2.getCell(`F${legRow}`).value = 'Pokrok (hotová část)';
  sh2.getCell(`F${legRow}`).font = { name: FONT, size: 9, bold: true, color: { argb: C.NAVY_DARK } };
  sh2.getCell(`F${legRow}`).alignment = { horizontal: 'center', vertical: 'middle' };

  /* ──────────────────────────────────────────────────────────────────────── */
  /*  SHEET 3 — Návod                                                          */
  /* ──────────────────────────────────────────────────────────────────────── */
  const sh3 = wb.addWorksheet('Návod', {
    views: [{ showGridLines: false }],
    pageSetup: { orientation: 'portrait', paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0, margins: { left: 0.5, right: 0.5, top: 0.5, bottom: 0.5, header: 0.2, footer: 0.2 } },
  });
  sh3.getColumn(1).width = 4;
  sh3.getColumn(2).width = 90;

  // Title
  sh3.mergeCells('A1:B3');
  sh3.getCell('A1').value = {
    richText: [
      { text: 'Jak používat tuto šablonu\n', font: FONTS.title },
      { text: 'Krok za krokem — od prázdného souboru po hotový harmonogram.', font: FONTS.subtitle },
    ],
  };
  sh3.getCell('A1').alignment = { horizontal: 'left', vertical: 'middle', wrapText: true, indent: 1 };
  sh3.getCell('A1').fill = fillSolid(C.NAVY);
  sh3.getRow(1).height = 22;
  sh3.getRow(2).height = 22;
  sh3.getRow(3).height = 22;

  const steps: { title: string; body: string }[] = [
    {
      title: '1. Otevřete list „Ganttův diagram"',
      body: 'Hlavní list obsahuje tabulku úkolů a vpravo časovou osu rozdělenou po týdnech. Šablona je předvyplněná ukázkovým projektem „Implementace ERP". Vlastní data zadáte přepsáním řádků 6 a níže.',
    },
    {
      title: '2. Vyplňte sloupce A–E (úkol, vlastník, termíny)',
      body: 'Do sloupce A napište ID (např. 1.1), do sloupce B název úkolu, do sloupce C odpovědnou osobu, do sloupce D datum zahájení (formát dd.mm.rrrr) a do sloupce E datum dokončení.',
    },
    {
      title: '3. Sloupec F (Dnů) se spočítá automaticky',
      body: 'Vzorec =E-D+1 vrátí počet pracovních dnů úkolu. Pokud potřebujete přidat řádek, zkopírujte ho z existujícího — vzorec se posune korektně.',
    },
    {
      title: '4. Zadejte pokrok ve sloupci G a status v H',
      body: 'Pokrok zadávejte jako desetinné číslo 0–1 (Excel zobrazí jako 0 %–100 %). Status vyberte z rozbalovacího seznamu — barva buňky se změní automaticky podle volby.',
    },
    {
      title: '5. Časová osa se vykreslí sama',
      body: 'Sloupce J–AI obsahují 26 týdnů. Buňka se obarví modře, pokud daný týden spadá do termínu úkolu, a žlutě, pokud spadá do již splněné části (podle pokroku %). Měnit nemusíte nic — vzorec sám reaguje na termíny.',
    },
    {
      title: '6. Zkontrolujte souhrn dole',
      body: 'Na řádku „Souhrn projektu" vidíte celkový počet dnů a vážený průměr pokroku. Vážení je podle délky úkolu — krátké úkoly s vysokým pokrokem nezkreslí celkový obraz.',
    },
    {
      title: '7. Vytiskněte nebo exportujte do PDF',
      body: 'Soubor je nastavený na A4 na šířku, fit na 1 stránku. Pro export použijte Soubor → Exportovat → PDF (Excel) nebo Soubor → Stáhnout → PDF (Google Sheets).',
    },
    {
      title: '8. Pro větší projekty přejděte na AutoERP',
      body: 'Tato šablona zvládne 10–30 úkolů a jeden projekt. Pokud řídíte 5+ projektů, sdílíte úkoly s týmem nebo potřebujete kapacitní plánování zdrojů, podívejte se na modul Projektové řízení v AutoERP — autoerp.cz/projektove-rizeni.',
    },
  ];

  let rowCursor = 5;
  for (const step of steps) {
    const tCell = sh3.getCell(`B${rowCursor}`);
    tCell.value = step.title;
    tCell.font = FONTS.h2;
    tCell.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
    sh3.getRow(rowCursor).height = 22;
    rowCursor++;

    const bCell = sh3.getCell(`B${rowCursor}`);
    bCell.value = step.body;
    bCell.font = FONTS.body;
    bCell.alignment = { horizontal: 'left', vertical: 'top', wrapText: true, indent: 1 };
    sh3.getRow(rowCursor).height = 42;
    rowCursor += 2;
  }

  // CTA box
  const ctaRow = rowCursor + 1;
  sh3.mergeCells(`A${ctaRow}:B${ctaRow + 2}`);
  const ctaCell = sh3.getCell(`A${ctaRow}`);
  ctaCell.value = {
    richText: [
      { text: 'Vyzkoušejte AutoERP zdarma\n', font: { name: FONT, size: 14, bold: true, color: { argb: C.NAVY } } },
      { text: 'Projektové řízení s Ganttovým diagramem, kapacitním plánováním a napojením na fakturaci.\n', font: { name: FONT, size: 10, color: { argb: C.TEXT_BODY } } },
      { text: 'autoerp.cz/projektove-rizeni', font: { name: FONT, size: 11, bold: true, color: { argb: C.ACCENT_HOVER }, underline: 'single' } },
    ],
  };
  ctaCell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true, indent: 2 };
  ctaCell.fill = fillSolid(C.ACCENT_SOFT);
  ctaCell.border = {
    top:    { style: 'medium', color: { argb: C.ACCENT } },
    bottom: { style: 'medium', color: { argb: C.ACCENT } },
    left:   { style: 'medium', color: { argb: C.ACCENT } },
    right:  { style: 'medium', color: { argb: C.ACCENT } },
  };
  sh3.getRow(ctaRow).height = 24;
  sh3.getRow(ctaRow + 1).height = 24;
  sh3.getRow(ctaRow + 2).height = 24;

  /* ──────────────────────────────────────────────────────────────────────── */
  /*  SHEET 4 — O šabloně                                                      */
  /* ──────────────────────────────────────────────────────────────────────── */
  const sh4 = wb.addWorksheet('O šabloně', {
    views: [{ showGridLines: false }],
    pageSetup: { orientation: 'portrait', paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
  });
  sh4.getColumn(1).width = 4;
  sh4.getColumn(2).width = 90;

  sh4.mergeCells('A1:B3');
  sh4.getCell('A1').value = {
    richText: [
      { text: 'O této šabloně\n', font: FONTS.title },
      { text: 'AutoERP — open-source nástroje pro české firmy', font: FONTS.subtitle },
    ],
  };
  sh4.getCell('A1').alignment = { horizontal: 'left', vertical: 'middle', wrapText: true, indent: 1 };
  sh4.getCell('A1').fill = fillSolid(C.NAVY);
  sh4.getRow(1).height = 22;
  sh4.getRow(2).height = 22;
  sh4.getRow(3).height = 22;

  const aboutBlocks: { h: string; b: string }[] = [
    {
      h: 'Co tato šablona umí',
      b: 'Hotový Ganttův diagram pro Excel, LibreOffice Calc a Google Sheets. 17 ukázkových úkolů ve fázích Příprava → Analýza → Implementace → Testování → Spuštění. Automatický výpočet délky úkolu, vážený pokrok projektu, podmíněné formátování pro status a časovou osu.',
    },
    {
      h: 'Pro koho je určená',
      b: 'Pro projektové manažery, výrobní ředitele a CEO menších firem (do 50 zaměstnanců), kteří plánují implementaci ERP, stavební zakázku, marketingovou kampaň nebo libovolný projekt s 10–30 úkoly. Šablona je v češtině a respektuje českou legislativu (data dd.mm.rrrr, úkoly v češtině).',
    },
    {
      h: 'Licence',
      b: 'MIT licence — můžete šablonu volně používat, upravovat, šířit a používat komerčně. Nemusíte uvádět zdroj, ale potěšíte nás. Plný text licence: github.com/Draivix/ganttuv-diagram-generator/blob/main/LICENSE',
    },
    {
      h: 'Open source',
      b: 'Zdrojový kód generátoru i celá šablona jsou na GitHubu: github.com/Draivix/ganttuv-diagram-generator. Našli jste chybu nebo máte nápad na vylepšení? Otevřete issue — odpovídáme do 24 hodin.',
    },
    {
      h: 'Vytvořil',
      b: 'Apertia Tech s.r.o. (IČO 27117758) — provozovatel AutoERP. Praha, ČR. Více o nás: autoerp.cz/o-nas',
    },
    {
      h: 'Potřebujete víc než šablonu?',
      b: 'AutoERP je modulární ERP/CRM pro české a slovenské firmy. Projektové řízení s Ganttovým diagramem, kapacitním plánováním zdrojů, napojením na fakturaci a sklad. Od 3 450 Kč/měsíc, bez licencí za uživatele. 14 dní zdarma na vyzkoušení. Více: autoerp.cz/projektove-rizeni',
    },
  ];

  rowCursor = 5;
  for (const block of aboutBlocks) {
    const tCell = sh4.getCell(`B${rowCursor}`);
    tCell.value = block.h;
    tCell.font = FONTS.h2;
    tCell.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
    sh4.getRow(rowCursor).height = 22;
    rowCursor++;

    const bCell = sh4.getCell(`B${rowCursor}`);
    bCell.value = block.b;
    bCell.font = FONTS.body;
    bCell.alignment = { horizontal: 'left', vertical: 'top', wrapText: true, indent: 1 };
    sh4.getRow(rowCursor).height = 56;
    rowCursor += 2;
  }

  // CTA at bottom
  const aboutCta = rowCursor + 1;
  sh4.mergeCells(`A${aboutCta}:B${aboutCta + 2}`);
  const aboutCtaCell = sh4.getCell(`A${aboutCta}`);
  aboutCtaCell.value = {
    richText: [
      { text: 'Vyzkoušejte AutoERP zdarma — 14 dní bez závazku\n', font: { name: FONT, size: 14, bold: true, color: { argb: C.NAVY } } },
      { text: 'autoerp.cz/projektove-rizeni', font: { name: FONT, size: 12, bold: true, color: { argb: C.ACCENT_HOVER }, underline: 'single' } },
    ],
  };
  aboutCtaCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  aboutCtaCell.fill = fillSolid(C.ACCENT_SOFT);
  aboutCtaCell.border = {
    top:    { style: 'medium', color: { argb: C.ACCENT } },
    bottom: { style: 'medium', color: { argb: C.ACCENT } },
    left:   { style: 'medium', color: { argb: C.ACCENT } },
    right:  { style: 'medium', color: { argb: C.ACCENT } },
  };
  sh4.getRow(aboutCta).height = 24;
  sh4.getRow(aboutCta + 1).height = 24;
  sh4.getRow(aboutCta + 2).height = 24;

  // Save
  await wb.xlsx.writeFile(OUT_FILE);
  console.log(`Excel template generated: ${OUT_FILE}`);
}

build().catch((err) => {
  console.error('Failed to generate Gantt template:', err);
  process.exit(1);
});
