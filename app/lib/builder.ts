import ExcelJS from 'exceljs'
import { ParsedSpec, ModelData, Component } from './parser'

// Colors
const NAVY   = 'FF1F4E79'
const WHITE  = 'FFFFFFFF'
const LBLUE  = 'FFD6E4F0'
const ALT    = 'FFF2F7FC'
const BORDER = 'FFB4C6E7'
const YELLOW = 'FFFFFF2CC'
const GREEN  = 'FFE2EFDA'
const ORANGE = 'FFFCE4D6'

function thin() {
  const s = { style: 'thin' as const, color: { argb: BORDER } }
  return { top: s, left: s, bottom: s, right: s }
}

function styleCell(c: ExcelJS.Cell, opts: {
  bold?: boolean; bg?: string; fg?: string; size?: number;
  hAlign?: ExcelJS.Alignment['horizontal']; vAlign?: ExcelJS.Alignment['vertical'];
  border?: boolean; wrap?: boolean;
}) {
  if (opts.bold !== undefined || opts.fg || opts.size) {
    c.font = { bold: opts.bold, color: { argb: opts.fg || '000000' }, size: opts.size || 10, name: 'Arial' }
  }
  if (opts.bg) c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: opts.bg } }
  if (opts.hAlign || opts.vAlign || opts.wrap) {
    c.alignment = { horizontal: opts.hAlign || 'left', vertical: opts.vAlign || 'middle', wrapText: opts.wrap }
  }
  if (opts.border !== false) c.border = thin()
}

function getColLetter(n: number): string {
  let col = ''
  while (n > 0) { const r = (n - 1) % 26; col = String.fromCharCode(65 + r) + col; n = Math.floor((n - 1) / 26) }
  return col
}

// ─── Output format matching 整理範例 ─────────────────────────────────────────
// Single sheet, vertical layout:
//   Row 1: title bar (版次 | filename | version)
//   Row 2: header row (空 | 空 | 項目 | model1 | model2 | ...)
//   Row 3-9: 規格 rows
//   Row 10: 重量 section header (Name | Mass(g) per model)
//   Row 11+: each component
//   ...density section
// ─────────────────────────────────────────────────────────────────────────────

export async function buildOutputExcel(spec: ParsedSpec): Promise<Buffer> {
  const wb = new ExcelJS.Workbook()
  wb.creator = 'CLW-SPC-TRANSFORM'

  const ws = wb.addWorksheet('整理')
  buildMainSheet(ws, spec)

  return Buffer.from(await wb.xlsx.writeBuffer())
}

export async function buildCompareExcel(specA: ParsedSpec, specB: ParsedSpec, labelA: string, labelB: string): Promise<Buffer> {
  const wb = new ExcelJS.Workbook()
  wb.creator = 'CLW-SPC-TRANSFORM'

  const ws = wb.addWorksheet('版次比對')
  buildCompareSheet(ws, specA, specB, labelA, labelB)

  return Buffer.from(await wb.xlsx.writeBuffer())
}

function buildMainSheet(ws: ExcelJS.Worksheet, spec: ParsedSpec) {
  const models = spec.models
  const M = models.length
  if (M === 0) return

  // Column layout:
  // A=大類, B=小類, C=項目, D...(D+M-1)=model values
  const colOffset = 4  // D = col 4

  // ── Row 1: Title ────────────────────────────────────────────────────
  ws.mergeCells(1, 1, 1, 3 + M)
  const t = ws.getCell(1, 1)
  t.value = `版次　${spec.fileName}　　${spec.version}`
  styleCell(t, { bold: true, fg: WHITE, bg: NAVY, size: 12, hAlign: 'center', vAlign: 'middle' })
  ws.getRow(1).height = 26

  // ── Row 2: Header (型號 row) ─────────────────────────────────────────
  ws.mergeCells(2, 1, 2, 2)
  ws.getCell(2, 1).value = ''
  ws.getCell(2, 3).value = '項目'
  styleCell(ws.getCell(2, 1), { bold: true, bg: NAVY, fg: WHITE, hAlign: 'center', vAlign: 'middle' })
  styleCell(ws.getCell(2, 3), { bold: true, bg: NAVY, fg: WHITE, hAlign: 'center', vAlign: 'middle' })
  ws.getRow(2).height = 22

  models.forEach((m, i) => {
    const c = ws.getCell(2, colOffset + i)
    c.value = m.modelName
    styleCell(c, { bold: true, bg: NAVY, fg: WHITE, hAlign: 'center', vAlign: 'middle' })
  })

  let row = 3

  // ── 規格 section ────────────────────────────────────────────────────
  const specDefs: [string, (m: ModelData) => string | number | null][] = [
    ['Final Head Weight', m => r4(m.finalHeadWeight)],
    ['Loft Angle', m => r4(m.loftAngle)],
    ['Lie Angle', m => r4(m.lieAngle)],
    ['Hosel OD(英制)', m => r4(m.hoselOD_in)],
    ['Hosel OD(公制)', m => r4(m.hoselOD_mm)],
    ['F1/E  Distance(英制)', m => r4(m.f1e_in)],
    ['F1/E  Distance(公制)', m => r4(m.f1e_mm)],
  ]

  specDefs.forEach(([label, getter], si) => {
    const isFirst = si === 0
    // Col A: merge "規格" for entire section
    if (isFirst) {
      ws.mergeCells(row, 1, row + specDefs.length - 1, 1)
      const ac = ws.getCell(row, 1)
      ac.value = '規格'
      styleCell(ac, { bold: true, bg: LBLUE, hAlign: 'center', vAlign: 'middle' })
    }
    // Col B: empty (used for sub-category if needed)
    const bc = ws.getCell(row, 2)
    bc.value = ''
    styleCell(bc, { bg: ALT })

    // Col C: label
    const lc = ws.getCell(row, 3)
    lc.value = label
    styleCell(lc, { bg: row % 2 === 0 ? ALT : WHITE })

    models.forEach((m, mi) => {
      const c = ws.getCell(row, colOffset + mi)
      c.value = getter(m) ?? ''
      styleCell(c, { hAlign: 'center', bg: row % 2 === 0 ? ALT : WHITE })
    })
    row++
  })

  // ── 重量 section ─────────────────────────────────────────────────────
  // Section header row
  ws.mergeCells(row, 1, row, 2)
  ws.getCell(row, 1).value = '重量'
  styleCell(ws.getCell(row, 1), { bold: true, bg: LBLUE, hAlign: 'center' })
  ws.getCell(row, 3).value = 'Name'
  styleCell(ws.getCell(row, 3), { bold: true, bg: LBLUE, hAlign: 'center' })
  models.forEach((m, mi) => {
    const c = ws.getCell(row, colOffset + mi)
    c.value = 'Mass (g)'
    styleCell(c, { bold: true, bg: LBLUE, hAlign: 'center' })
  })
  row++

  for (const compName of spec.allComponentNames) {
    ws.mergeCells(row, 1, row, 2)
    ws.getCell(row, 1).value = ''
    styleCell(ws.getCell(row, 1), { bg: row % 2 === 0 ? ALT : WHITE })
    ws.getCell(row, 3).value = compName
    styleCell(ws.getCell(row, 3), { bg: row % 2 === 0 ? ALT : WHITE })
    models.forEach((m, mi) => {
      const comp = m.components.find(c => c.name === compName)
      const c = ws.getCell(row, colOffset + mi)
      c.value = comp ? comp.mass_g : ''
      styleCell(c, { hAlign: 'center', bg: row % 2 === 0 ? ALT : WHITE })
    })
    row++
  }

  // ── 密度(英制) section ────────────────────────────────────────────────
  ws.mergeCells(row, 1, row, 2)
  ws.getCell(row, 1).value = '密度  (英制)'
  styleCell(ws.getCell(row, 1), { bold: true, bg: LBLUE, hAlign: 'center' })
  ws.getCell(row, 3).value = 'Name'
  styleCell(ws.getCell(row, 3), { bold: true, bg: LBLUE, hAlign: 'center' })
  models.forEach((m, mi) => {
    const c = ws.getCell(row, colOffset + mi)
    c.value = 'Density (g/in3)'
    styleCell(c, { bold: true, bg: LBLUE, hAlign: 'center' })
  })
  row++

  for (const compName of spec.allComponentNames) {
    ws.mergeCells(row, 1, row, 2)
    ws.getCell(row, 1).value = ''
    styleCell(ws.getCell(row, 1), { bg: row % 2 === 0 ? ALT : WHITE })
    ws.getCell(row, 3).value = compName
    styleCell(ws.getCell(row, 3), { bg: row % 2 === 0 ? ALT : WHITE })
    models.forEach((m, mi) => {
      const comp = m.components.find(c => c.name === compName)
      const c = ws.getCell(row, colOffset + mi)
      c.value = comp?.density_g_in3 ?? ''
      styleCell(c, { hAlign: 'center', bg: row % 2 === 0 ? ALT : WHITE })
    })
    row++
  }

  // ── 密度(公制) section ────────────────────────────────────────────────
  ws.mergeCells(row, 1, row, 2)
  ws.getCell(row, 1).value = '密度  (公制)'
  styleCell(ws.getCell(row, 1), { bold: true, bg: LBLUE, hAlign: 'center' })
  ws.getCell(row, 3).value = 'Name'
  styleCell(ws.getCell(row, 3), { bold: true, bg: LBLUE, hAlign: 'center' })
  models.forEach((m, mi) => {
    const c = ws.getCell(row, colOffset + mi)
    c.value = 'Density (g/cm3)'
    styleCell(c, { bold: true, bg: LBLUE, hAlign: 'center' })
  })
  row++

  for (const compName of spec.allComponentNames) {
    ws.mergeCells(row, 1, row, 2)
    ws.getCell(row, 1).value = ''
    styleCell(ws.getCell(row, 1), { bg: row % 2 === 0 ? ALT : WHITE })
    ws.getCell(row, 3).value = compName
    styleCell(ws.getCell(row, 3), { bg: row % 2 === 0 ? ALT : WHITE })
    models.forEach((m, mi) => {
      const comp = m.components.find(c => c.name === compName)
      const c = ws.getCell(row, colOffset + mi)
      c.value = comp?.density_g_cm3 ?? ''
      styleCell(c, { hAlign: 'center', bg: row % 2 === 0 ? ALT : WHITE })
    })
    row++
  }

  // Column widths
  ws.getColumn(1).width = 14
  ws.getColumn(2).width = 4
  ws.getColumn(3).width = 26
  models.forEach((_, i) => { ws.getColumn(colOffset + i).width = 18 })
}

function buildCompareSheet(
  ws: ExcelJS.Worksheet,
  specA: ParsedSpec, specB: ParsedSpec,
  labelA: string, labelB: string
) {
  const allModels = Array.from(new Set([
    ...specA.models.map(m => m.modelName),
    ...specB.models.map(m => m.modelName),
  ]))
  const allComps = Array.from(new Set([
    ...specA.allComponentNames,
    ...specB.allComponentNames,
  ]))

  // Title
  ws.mergeCells(1, 1, 1, 6)
  const t = ws.getCell(1, 1)
  t.value = `版次比對　${labelA}  vs  ${labelB}`
  styleCell(t, { bold: true, fg: WHITE, bg: NAVY, size: 12, hAlign: 'center', vAlign: 'middle' })
  ws.getRow(1).height = 26

  // Legend
  const legend = ws.getRow(2)
  legend.height = 18
  ;[
    { col: 1, text: '🟡 數值變更', bg: YELLOW },
    { col: 2, text: '🟢 新增', bg: GREEN },
    { col: 3, text: '🔴 移除', bg: ORANGE },
  ].forEach(({ col, text, bg }) => {
    const c = legend.getCell(col)
    c.value = text
    styleCell(c, { bg, size: 9 })
  })

  let row = 3

  for (const modelName of allModels) {
    const mA = specA.models.find(m => m.modelName === modelName)
    const mB = specB.models.find(m => m.modelName === modelName)

    // Model header
    ws.mergeCells(row, 1, row, 6)
    const mh = ws.getCell(row, 1)
    mh.value = `▶ ${modelName}`
    styleCell(mh, { bold: true, fg: WHITE, bg: 'FF2E75B6', size: 11, vAlign: 'middle' })
    ws.getRow(row).height = 20
    row++

    // Sub headers
    ;['大類', '項目', '單位', labelA, labelB, '差異'].forEach((v, i) => {
      const c = ws.getCell(row, i + 1)
      c.value = v
      styleCell(c, { bold: true, bg: LBLUE, hAlign: 'center' })
    })
    row++

    // Spec rows
    const specFields: [string, string, (m: ModelData) => number | null][] = [
      ['規格', 'Final Head Weight', m => m.finalHeadWeight],
      ['規格', 'Loft Angle', m => m.loftAngle],
      ['規格', 'Lie Angle', m => m.lieAngle],
      ['規格', 'Hosel OD(英制)', m => m.hoselOD_in],
      ['規格', 'Hosel OD(公制)', m => m.hoselOD_mm],
      ['規格', 'F1/E Distance(英制)', m => m.f1e_in],
      ['規格', 'F1/E Distance(公制)', m => m.f1e_mm],
    ]

    for (const [cat, label, getter] of specFields) {
      const vA = mA ? getter(mA) : null
      const vB = mB ? getter(mB) : null
      const diff = vA !== null && vB !== null ? r4(vB - vA) : null
      const changed = diff !== null && diff !== 0
      const bg = changed ? YELLOW : row % 2 === 0 ? ALT : WHITE
      writeRow(ws, row++, [cat, label, '', vA ?? '', vB ?? '', diff ?? ''], bg, changed)
    }

    // Component weight rows
    ws.getCell(row, 1).value = '重量'
    ws.getCell(row, 2).value = 'Name'
    ws.getCell(row, 3).value = 'g'
    ws.getCell(row, 4).value = 'Mass A'
    ws.getCell(row, 5).value = 'Mass B'
    ws.getCell(row, 6).value = '差異'
    ;[1,2,3,4,5,6].forEach(i => styleCell(ws.getCell(row, i), { bold: true, bg: LBLUE, hAlign: 'center' }))
    row++

    for (const compName of allComps) {
      const cA = mA?.components.find(c => c.name === compName)
      const cB = mB?.components.find(c => c.name === compName)
      const vA = cA?.mass_g ?? null
      const vB = cB?.mass_g ?? null
      const isNew = vA === null && vB !== null
      const isRemoved = vA !== null && vB === null
      const diff = vA !== null && vB !== null ? r4(vB - vA) : null
      const changed = diff !== null && diff !== 0
      const bg = isNew ? GREEN : isRemoved ? ORANGE : changed ? YELLOW : row % 2 === 0 ? ALT : WHITE
      writeRow(ws, row++, ['', compName, 'g', vA ?? '', vB ?? '', diff ?? ''], bg, changed || isNew || isRemoved)
    }

    row++ // blank row between models
  }

  ws.getColumn(1).width = 10
  ws.getColumn(2).width = 26
  ws.getColumn(3).width = 8
  ws.getColumn(4).width = 16
  ws.getColumn(5).width = 16
  ws.getColumn(6).width = 12
}

function writeRow(ws: ExcelJS.Worksheet, rowIdx: number, vals: (string | number)[], bg: string, highlight: boolean) {
  vals.forEach((v, i) => {
    const c = ws.getCell(rowIdx, i + 1)
    c.value = v
    styleCell(c, {
      bg,
      bold: highlight && i >= 3,
      hAlign: i >= 3 ? 'center' : 'left',
    })
  })
}

function r4(v: number | null): number | null {
  if (v === null) return null
  return Math.round(v * 10000) / 10000
}
