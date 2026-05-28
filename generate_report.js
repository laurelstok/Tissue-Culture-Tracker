#!/usr/bin/env node
/**
 * TC Tracker — Experiment Report Generator
 * Usage: node generate_report.js '<json>' <output.pptx>
 */

const pptxgen = require('pptxgenjs');
const fs = require('fs');
const path = require('path');

// ── Palette ───────────────────────────────────────────────────────────────────
const C = {
  navy:    '1B2A4A',
  teal:    '0F6E56',
  amber:   '854F0B',
  blue:    '185FA5',
  purple:  '534AB7',
  red:     '993C1D',
  grey:    '6B6A65',
  light:   'F4F3EF',
  white:   'FFFFFF',
  ink:     '1A1915',
};

const ACTION_COLOR = {
  thaw: C.blue, inherit: C.purple, passage: C.teal,
  freeze: C.amber, experiment: C.red, discontinued: '999999',
};

function actionColor(action) { return ACTION_COLOR[action] || C.grey; }

// ── Helpers ───────────────────────────────────────────────────────────────────
function sciToNum(s) {
  if (!s) return null;
  const str = String(s).replace(/[×x\*]/g,'e').replace(/\u00d7/g,'e');
  const n = parseFloat(str);
  if (!isNaN(n)) return n;
  const m = str.match(/^([\d.]+)?\s*[eE]\s*(.+)$/);
  if (m) return (parseFloat(m[1]||'1')) * Math.pow(10, parseFloat(m[2]));
  return null;
}

function fmtNum(n) {
  if (n === null || n === undefined) return '—';
  if (n >= 1e9) return (n/1e9).toFixed(1) + 'B';
  if (n >= 1e6) return (n/1e6).toFixed(1) + 'M';
  if (n >= 1e3) return (n/1e3).toFixed(0) + 'K';
  return String(Math.round(n));
}

function pL(rec) {
  const a = rec.action;
  const p = rec.passageNum || 0;
  if (a === 'thaw') return `Thaw P${p}`;
  if (a === 'freeze') return `P${p} ❄`;
  if (a === 'experiment') return `P${p} Exp`;
  return `P${p}`;
}

function imageBase64(imgPath) {
  try {
    const buf = fs.readFileSync(imgPath);
    const ext = path.extname(imgPath).toLowerCase().replace('.','');
    const mime = ext === 'jpg' || ext === 'jpeg' ? 'jpeg' : ext === 'png' ? 'png' : ext;
    return `image/${mime};base64,` + buf.toString('base64');
  } catch { return null; }
}

// Find images in a folder whose filename contains a well ID like A1, B2 etc.
function findWellImages(folder) {
  if (!folder || !fs.existsSync(folder)) return {};
  const WELL_RE = /\b([A-P])([0-9]{1,2})\b/i;
  const EXTS = ['.jpg','.jpeg','.png','.tif','.tiff','.bmp'];
  const map = {};
  try {
    const files = fs.readdirSync(folder);
    for (const f of files) {
      const ext = path.extname(f).toLowerCase();
      if (!EXTS.includes(ext)) continue;
      const m = WELL_RE.exec(path.basename(f, ext));
      if (m) {
        const wid = m[1].toUpperCase() + String(parseInt(m[2]));
        if (!map[wid]) map[wid] = path.join(folder, f);
      }
    }
  } catch {}
  return map;
}

// ── Slide builders ─────────────────────────────────────────────────────────────

function addCoverSlide(pres, rec, passages) {
  const s = pres.addSlide();
  s.background = { color: C.navy };

  // Left accent bar
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 0.18, h: 5.625, fill: { color: C.teal }, line: { type: 'none' }
  });

  // Experiment ID pill
  const expId = rec.expId || rec.id;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.45, y: 0.38, w: Math.min(expId.length * 0.13 + 0.4, 4), h: 0.32,
    fill: { color: C.teal }, line: { type: 'none' }, rectRadius: 0.05
  });
  s.addText(expId, {
    x: 0.45, y: 0.38, w: 4, h: 0.32,
    fontSize: 11, color: C.white, bold: true, fontFace: 'Calibri', margin: 0,
    align: 'left', valign: 'middle'
  });

  // Title
  s.addText(rec.assay || rec.line || 'Experiment Report', {
    x: 0.45, y: 0.9, w: 9.1, h: 1.5,
    fontSize: 40, bold: true, color: C.white, fontFace: 'Calibri',
    align: 'left', valign: 'middle'
  });

  // Meta row
  const meta = [
    rec.line && `Cell line: ${rec.line}`,
    rec.date && `Date: ${rec.date}`,
    rec.user && `Operator: ${rec.user}`,
    pL(rec),
  ].filter(Boolean).join('   ·   ');
  s.addText(meta, {
    x: 0.45, y: 2.5, w: 9.1, h: 0.4,
    fontSize: 13, color: 'AADDD0', fontFace: 'Calibri',
  });

  // Treatment
  if (rec.treatment) {
    s.addText('Treatment: ' + rec.treatment, {
      x: 0.45, y: 3.0, w: 9.1, h: 0.35,
      fontSize: 12, color: 'C8C6BF', fontFace: 'Calibri', italic: true,
    });
  }

  // Timepoints
  if (rec.timepoints) {
    s.addText('Timepoints: ' + rec.timepoints, {
      x: 0.45, y: 3.4, w: 9.1, h: 0.35,
      fontSize: 12, color: 'C8C6BF', fontFace: 'Calibri',
    });
  }

  // Note
  if (rec.note) {
    s.addText(rec.note, {
      x: 0.45, y: 3.9, w: 9.1, h: 0.8,
      fontSize: 11, color: '888880', fontFace: 'Calibri', italic: true, wrap: true,
    });
  }

  // Footer
  s.addText(`Generated ${new Date().toLocaleDateString()}  ·  TC Tracker`, {
    x: 0.45, y: 5.25, w: 9.1, h: 0.28,
    fontSize: 9, color: '555550', fontFace: 'Calibri',
  });
}

function addSummarySlide(pres, rec) {
  const s = pres.addSlide();
  s.background = { color: C.light };

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.65, fill: { color: C.navy }, line: { type: 'none' }
  });
  s.addText('Experiment Summary', {
    x: 0.3, y: 0, w: 9.4, h: 0.65,
    fontSize: 20, bold: true, color: C.white, fontFace: 'Calibri', valign: 'middle',
  });

  // Stat cards
  const stats = [
    { label: 'Cell line', val: rec.line || '—' },
    { label: 'Passage', val: pL(rec) },
    { label: 'Date', val: rec.date || '—' },
    { label: 'Assay', val: rec.assay || '—' },
    { label: 'Viability', val: rec.viabilityPct ? rec.viabilityPct + '%' : '—' },
    { label: 'Total cells', val: rec.totalViableCells || '—' },
    { label: 'Confluency', val: rec.confluency ? rec.confluency + '%' : '—' },
    { label: 'Operator', val: rec.user || '—' },
  ];

  const cols = 4, cardW = 2.2, cardH = 1.0, gapX = 0.1, gapY = 0.15;
  const startX = 0.3, startY = 0.85;

  stats.forEach((st, i) => {
    const col = i % cols, row = Math.floor(i / cols);
    const x = startX + col * (cardW + gapX);
    const y = startY + row * (cardH + gapY);
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: cardW, h: cardH,
      fill: { color: C.white },
      shadow: { type: 'outer', blur: 4, offset: 1, angle: 135, color: '000000', opacity: 0.08 },
      line: { type: 'none' },
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 0.06, h: cardH, fill: { color: C.teal }, line: { type: 'none' }
    });
    s.addText(st.label, {
      x: x + 0.12, y: y + 0.08, w: cardW - 0.18, h: 0.28,
      fontSize: 9, color: C.grey, fontFace: 'Calibri', bold: true,
    });
    s.addText(String(st.val), {
      x: x + 0.12, y: y + 0.36, w: cardW - 0.18, h: 0.52,
      fontSize: 15, color: C.ink, fontFace: 'Calibri', bold: true, wrap: true,
    });
  });

  // Note box
  if (rec.note) {
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.3, y: 3.0, w: 9.4, h: 1.0,
      fill: { color: 'EAF5F1' }, line: { color: 'B0D8CC', pt: 1 },
    });
    s.addText('Notes: ' + rec.note, {
      x: 0.45, y: 3.05, w: 9.1, h: 0.9,
      fontSize: 11, color: C.teal, fontFace: 'Calibri', italic: true, wrap: true,
    });
  }

  // Treatment
  if (rec.treatment) {
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.3, y: 4.1, w: 9.4, h: 0.6,
      fill: { color: 'FDF3E3' }, line: { color: 'E8D5A0', pt: 1 },
    });
    s.addText('Treatment: ' + rec.treatment, {
      x: 0.45, y: 4.15, w: 9.1, h: 0.5,
      fontSize: 11, color: C.amber, fontFace: 'Calibri', wrap: true,
    });
  }
}

function addPlateSlide(pres, rec, plateKey, plateData, imageFolder, slideTitle) {
  const s = pres.addSlide();
  s.background = { color: C.white };

  // Header bar
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.55, fill: { color: C.navy }, line: { type: 'none' }
  });
  s.addText(slideTitle || `Plate data — ${rec.id}`, {
    x: 0.3, y: 0, w: 9.4, h: 0.55,
    fontSize: 16, bold: true, color: C.white, fontFace: 'Calibri', valign: 'middle',
  });

  // Find images for this plate
  const wellImages = imageFolder ? findWellImages(imageFolder) : {};

  // Get occupied wells
  const wells = Object.keys(plateData).filter(wid => plateData[wid] && plateData[wid].occupied);
  if (!wells.length && !Object.keys(wellImages).length) {
    s.addText('No well data recorded for this plate.', {
      x: 0.5, y: 2, w: 9, h: 1,
      fontSize: 14, color: C.grey, fontFace: 'Calibri', align: 'center',
    });
    return;
  }

  // Determine grid layout — up to 6 wells per slide
  const allWells = [...new Set([...wells, ...Object.keys(wellImages)])].sort();
  const COLS = Math.min(3, allWells.length);
  const ROWS = Math.ceil(allWells.length / COLS);
  const cellW = COLS <= 2 ? 4.4 : 3.0;
  const cellH = ROWS <= 1 ? 3.5 : 2.2;
  const startX = (10 - COLS * cellW - (COLS - 1) * 0.15) / 2;
  const startY = 0.7;

  allWells.slice(0, 9).forEach((wid, i) => {
    const col = i % COLS;
    const row = Math.floor(i / COLS);
    const x = startX + col * (cellW + 0.15);
    const y = startY + row * (cellH + 0.15);
    const wd = plateData[wid] || {};

    // Card background
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: cellW, h: cellH,
      fill: { color: 'F8F8F7' },
      shadow: { type: 'outer', blur: 5, offset: 2, angle: 135, color: '000000', opacity: 0.10 },
      line: { type: 'none' },
    });

    // Well label pill
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 0.45, h: 0.25, fill: { color: C.navy }, line: { type: 'none' }
    });
    s.addText(wid, {
      x, y, w: 0.45, h: 0.25,
      fontSize: 9, bold: true, color: C.white, fontFace: 'Calibri',
      align: 'center', valign: 'middle', margin: 0,
    });

    // Image
    const imgPath = wellImages[wid];
    const imgH = cellH - 0.6;
    const imgY = y + 0.28;
    if (imgPath) {
      const imgData = imageBase64(imgPath);
      if (imgData) {
        s.addImage({ data: imgData, x: x + 0.05, y: imgY, w: cellW - 0.1, h: imgH - 0.3,
          sizing: { type: 'contain', w: cellW - 0.1, h: imgH - 0.3 } });
      }
    } else {
      // Placeholder
      s.addShape(pres.shapes.RECTANGLE, {
        x: x + 0.05, y: imgY, w: cellW - 0.1, h: imgH - 0.3,
        fill: { color: 'EBEBEA' }, line: { color: 'D5D4CE', pt: 1 },
      });
      s.addText('No image', {
        x: x + 0.05, y: imgY, w: cellW - 0.1, h: imgH - 0.3,
        fontSize: 9, color: 'AAAAAA', fontFace: 'Calibri', align: 'center', valign: 'middle',
      });
    }

    // Metadata row at bottom of card
    const metaParts = [
      wd.confluency != null && `${wd.confluency}% conf.`,
      wd.seeding && `Seed: ${wd.seeding}`,
      wd.count && `Count: ${wd.count}`,
      wd.cond && wd.cond,
      wd.treat && wd.treat,
    ].filter(Boolean);

    s.addText(metaParts.join('  ·  ') || '—', {
      x, y: y + cellH - 0.3, w: cellW, h: 0.28,
      fontSize: 8, color: C.grey, fontFace: 'Calibri',
      align: 'center', wrap: true,
    });
  });
}

function addGrowthCurveSlide(pres, rec, passages) {
  // Collect passage records in lineage
  const lineageIds = new Set();
  const queue = [rec.id];
  while (queue.length) {
    const id = queue.shift();
    if (lineageIds.has(id)) continue;
    lineageIds.add(id);
    const r = passages.find(p => p.id === id);
    if (r && r.parent) queue.push(r.parent);
  }
  const linePts = passages.filter(p =>
    p.line === rec.line && p.totalViableCells && sciToNum(p.totalViableCells) !== null
  ).sort((a, b) => a.date < b.date ? -1 : 1);

  if (linePts.length < 2) return;

  const s = pres.addSlide();
  s.background = { color: C.white };

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.55, fill: { color: C.navy }, line: { type: 'none' }
  });
  s.addText(`Growth curve — ${rec.line}`, {
    x: 0.3, y: 0, w: 9.4, h: 0.55,
    fontSize: 16, bold: true, color: C.white, fontFace: 'Calibri', valign: 'middle',
  });

  // Native line chart
  const labels = linePts.map(p => `${p.date.slice(5)}\n${pL(p)}`);
  const values = linePts.map(p => {
    const v = sciToNum(p.totalViableCells);
    return v ? Math.round(v) : 0;
  });

  s.addChart(pres.charts.LINE, [{
    name: rec.line + ' viable cells',
    labels,
    values,
  }], {
    x: 0.5, y: 0.7, w: 9, h: 3.8,
    lineSize: 2.5,
    lineSmooth: true,
    chartColors: [C.teal],
    showValue: true,
    dataLabelColor: C.teal,
    dataLabelFontSize: 9,
    catAxisLabelColor: C.grey,
    valAxisLabelColor: C.grey,
    valGridLine: { color: 'E8E7E3', size: 0.5 },
    catGridLine: { style: 'none' },
    chartArea: { fill: { color: C.white }, roundedCorners: false },
    showLegend: false,
    valAxisTitle: 'Total viable cells',
    showValAxisTitle: true,
  });

  // Stat row
  const maxCount = Math.max(...values);
  const lastCount = values[values.length - 1];
  s.addText(`Peak: ${fmtNum(maxCount)} cells   ·   Latest: ${fmtNum(lastCount)} cells   ·   ${linePts.length} records`, {
    x: 0.5, y: 4.7, w: 9, h: 0.35,
    fontSize: 11, color: C.grey, fontFace: 'Calibri', align: 'center',
  });
}

function addWellCurveSlide(pres, rec) {
  if (!rec.plateData) return;
  const plates = Object.keys(rec.plateData);
  if (plates.length < 2) return;

  // Collect wells with counts across plates
  const wellData = {};
  plates.forEach((pk, pi) => {
    const plate = rec.plateData[pk];
    Object.keys(plate).forEach(wid => {
      const w = plate[wid];
      const v = sciToNum(w.count);
      if (v) {
        if (!wellData[wid]) wellData[wid] = [];
        wellData[wid].push({ plate: pi + 1, count: v, cond: w.cond || '', treat: w.treat || '' });
      }
    });
  });

  const wellIds = Object.keys(wellData);
  if (!wellIds.length) return;

  const s = pres.addSlide();
  s.background = { color: C.white };

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.55, fill: { color: C.navy }, line: { type: 'none' }
  });
  s.addText(`Well counts over time — ${rec.expId || rec.id}`, {
    x: 0.3, y: 0, w: 9.4, h: 0.55,
    fontSize: 16, bold: true, color: C.white, fontFace: 'Calibri', valign: 'middle',
  });

  const COLORS = ['0F6E56','185FA5','854F0B','534AB7','993C1D','1D9E75','D85A30'];
  const labels = plates.map((_, i) => `Plate ${i+1}`);

  const chartData = wellIds.map((wid, wi) => ({
    name: wid,
    labels,
    values: plates.map((_, pi) => {
      const pt = (wellData[wid] || []).find(d => d.plate === pi + 1);
      return pt ? Math.round(pt.count) : 0;
    }),
  }));

  s.addChart(pres.charts.LINE, chartData, {
    x: 0.5, y: 0.7, w: 9, h: 3.8,
    lineSize: 2,
    lineSmooth: true,
    chartColors: COLORS,
    showValue: false,
    catAxisLabelColor: C.grey,
    valAxisLabelColor: C.grey,
    valGridLine: { color: 'E8E7E3', size: 0.5 },
    catGridLine: { style: 'none' },
    chartArea: { fill: { color: C.white } },
    showLegend: true,
    legendPos: 'r',
    valAxisTitle: 'Cell count',
    showValAxisTitle: true,
  });
}

// ── Main ──────────────────────────────────────────────────────────────────────

async function main() {
  const args = process.argv.slice(2);
  if (args.length < 2) {
    console.error('Usage: node generate_report.js <json_data> <output.pptx>');
    process.exit(1);
  }

  const data = JSON.parse(args[0]);
  const outputPath = args[1];

  const { rec, passages, imageFolder } = data;
  if (!rec) { console.error('No rec in data'); process.exit(1); }

  const pres = new pptxgen();
  pres.layout = 'LAYOUT_16x9';
  pres.author = rec.user || 'TC Tracker';
  pres.title = rec.assay || rec.expId || rec.id;

  // 1. Cover
  addCoverSlide(pres, rec, passages || []);

  // 2. Summary
  addSummarySlide(pres, rec);

  // 3. Plate slides (one per vessel plate)
  if (rec.plateData) {
    const plateKeys = Object.keys(rec.plateData);
    plateKeys.forEach((pk, i) => {
      const title = plateKeys.length > 1
        ? `Plate ${i+1} of ${plateKeys.length}  ·  ${rec.expId || rec.id}`
        : `Plate data  ·  ${rec.expId || rec.id}`;
      addPlateSlide(pres, rec, pk, rec.plateData[pk] || {}, imageFolder, title);
    });
  } else if (imageFolder) {
    // No plate data but we have images — make one image grid slide
    addPlateSlide(pres, rec, 'images', {}, imageFolder, `Images  ·  ${rec.expId || rec.id}`);
  }

  // 4. Well count curves (if multiple plates / timepoints)
  addWellCurveSlide(pres, rec);

  // 5. Growth curve
  addGrowthCurveSlide(pres, rec, passages || []);

  await pres.writeFile({ fileName: outputPath });
  console.log('OK:' + outputPath);
}

main().catch(e => { console.error(e); process.exit(1); });