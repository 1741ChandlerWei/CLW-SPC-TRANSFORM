import ExcelJS from 'exceljs'
import { ParsedSpec, ModelData } from './parser'

const IN_TO_MM = 25.4

// Color palette
const CLR = {
  headerBg: 'FF1F4E79',   // dark navy
  headerFg: 'FFFFFFFF',
  subHeaderBg: 'FF2E75B6', // medium blue
  subHeaderFg: 'FFFFFFFF',
  sectionBg: 'FFD6E4F0',  // light blue
  altRow: 'FFF2F7FC',
  changed: 'FFFFF2CC',    // yellow highlight for diffs
  added: 'FFE2EFDA',      // green for added
  removed: 'FFFCE4D6',    // red/orange for removed
  border: 'FFB4C6E7',
}

function headerStyle(wb: ExcelJS.Workbook, dark = true) {
  return {
    font: { bold: true, color: { argb: dark ? CLR.headerFg : 'FF000000' }, size: 10, name: 'Arial' },
    fill: { type: 'pattern' as const, pattern: 'solid' as const, fgColor: { argb: dark ? CLR.headerBg : CLR.sectionBg } },
    alignment: { horizontal: 'center' as const, vertical: 'middle' as const, wrapText: true },
    border: thinBorder(),
  }
}

function thinBorder() {
  const s = { style: 'thin' as const, color: { argb: CLR.border } }
  return { top: s, left: s, bottom: s, right: s }
}

function applyStyle(cell: ExcelJS.Cell, style: Partial<ExcelJS.Style>) {
  Object.assign(cell, style)
}

function numFmt(cell: ExcelJS.Cell, fmt: string) {
  cell.numFmt = fmt
}

export async function buildOutputExcel(spec: ParsedSpec): Promise<Buffer> {
  const wb = new ExcelJS.Workbook()
  wb.creator = 'CLW-SPC-TRANSFORM'

  buildSummarySheet(wb, spec)
  buildPartsSheet(wb, spec)
  buildDensitySheet(wb, spec)

  const buf = await wb.xlsx.writeBuffer()
  return Buffer.from(buf)
}

export async function buildCompareExcel(
  specA: ParsedSpec,
  specB: ParsedSpec,
  labelA: string,
  labelB: string
): Promise<Buffer> {
  const wb = new ExcelJS.Workbook()
  wb.creator = 'CLW-SPC-TRANSFORM'

  buildCompareSheet(wb, specA, specB, labelA, labelB)

  const buf = await wb.xlsx.writeBuffer()
  return Buffer.from(buf)
}

// ──────────────────────────────────────────────
// Summary Sheet: specs for all models
// ──────────────────────────────────────────────
function buildSummarySheet(wb: ExcelJS.Workbook, spec: ParsedSpec) {
  const ws = wb.addWorksheet('規格總覽')

  const models = spec.models
  if (models.length === 0) return

  // Row 1: title
  ws.mergeCells(1, 1, 1, 3 + models.length)
  const titleCell = ws.getCell(1, 1)
  titleCell.value = `CLW 規格整理 ｜ ${spec.fileName} ｜ ${spec.version}`
  titleCell.font = { bold: true, size: 13, name: 'Arial', color: { argb: 'FFFFFFFF' } }
  titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: CLR.headerBg } }
  titleCell.alignment = { horizontal: 'center', vertical: 'middle' }
  ws.getRow(1).height = 28

  // Row 2: headers
  const hRow = ws.getRow(2)
  hRow.height = 30
  ;['類別', '項目', '單位', ...models.map(m => m.modelName)].forEach((v, i) => {
    const c = hRow.getCell(i + 1)
    c.value = v
    Object.assign(c, headerStyle(wb, true))
  })

  const specRows: [string, string, string, (m: ModelData) => number | string | null][] = [
    ['規格', 'Final Head Weight', 'g', m => fmt4(m.finalHeadWeight)],
    ['規格', 'Loft Angle', '°', m => fmt4(m.loftAngle)],
    ['規格', 'Lie Angle', '°', m => fmt4(m.lieAngle)],
    ['規格', 'Hosel OD', 'in', m => fmt4(m.hoselOD_in)],
    ['規格', 'Hosel OD', 'mm', m => fmt4(m.hoselOD_mm)],
    ['規格', 'F1/E Distance', 'in', m => fmt4(m.f1e_in)],
    ['規格', 'F1/E Distance', 'mm', m => fmt4(m.f1e_mm)],
  ]

  let rowIdx = 3
  let lastCat = ''

  for (const [cat, label, unit, getter] of specRows) {
    const row = ws.getRow(rowIdx++)
    const isCatChange = cat !== lastCat
    lastCat = cat

    const catCell = row.getCell(1)
    catCell.value = isCatChange ? cat : ''
    catCell.font = { bold: true, size: 10, name: 'Arial' }
    catCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: CLR.sectionBg } }
    catCell.border = thinBorder()
    catCell.alignment = { vertical: 'middle', horizontal: 'center' }

    row.getCell(2).value = label
    row.getCell(3).value = unit
    ;[row.getCell(2), row.getCell(3)].forEach(c => {
      c.font = { size: 10, name: 'Arial' }
      c.border = thinBorder()
      c.alignment = { vertical: 'middle' }
    })

    models.forEach((m, i) => {
      const c = row.getCell(4 + i)
      const val = getter(m)
      c.value = val === null ? '' : val
      c.font = { size: 10, name: 'Arial' }
      c.border = thinBorder()
      c.alignment = { horizontal: 'center', vertical: 'middle' }
      if ((rowIdx - 1) % 2 === 0) c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: CLR.altRow } }
    })
  }

  // Column widths
  ws.getColumn(1).width = 8
  ws.getColumn(2).width = 22
  ws.getColumn(3).width = 8
  models.forEach((_, i) => { ws.getColumn(4 + i).width = 16 })
}

// ──────────────────────────────────────────────
// Parts Weight Sheet
// ──────────────────────────────────────────────
function buildPartsSheet(wb: ExcelJS.Workbook, spec: ParsedSpec) {
  const ws = wb.addWorksheet('部件重量')
  const models = spec.models
  const allParts = spec.allPartNames
  if (models.length === 0) return

  ws.mergeCells(1, 1, 1, 2 + models.length)
  const tc = ws.getCell(1, 1)
  tc.value = `部件重量明細 ｜ ${spec.fileName} ｜ ${spec.version}`
  tc.font = { bold: true, size: 13, name: 'Arial', color: { argb: 'FFFFFFFF' } }
  tc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: CLR.headerBg } }
  tc.alignment = { horizontal: 'center', vertical: 'middle' }
  ws.getRow(1).height = 28

  const hRow = ws.getRow(2)
  hRow.height = 28
  ;['部位名稱 (Name)', '單位', ...models.map(m => m.modelName)].forEach((v, i) => {
    const c = hRow.getCell(i + 1)
    c.value = v
    Object.assign(c, headerStyle(wb, true))
  })

  allParts.forEach((part, pi) => {
    const row = ws.getRow(3 + pi)
    row.getCell(1).value = part
    row.getCell(2).value = 'g'
    ;[row.getCell(1), row.getCell(2)].forEach(c => {
      c.font = { size: 10, name: 'Arial' }
      c.border = thinBorder()
      if (pi % 2 === 1) c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: CLR.altRow } }
    })
    models.forEach((m, mi) => {
      const c = row.getCell(3 + mi)
      const val = m.parts[part] ?? null
      c.value = val === null ? '' : val
      c.font = { size: 10, name: 'Arial' }
      c.border = thinBorder()
      c.alignment = { horizontal: 'center' }
      if (pi % 2 === 1) c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: CLR.altRow } }
    })
  })

  // Total row
  const totalRow = ws.getRow(3 + allParts.length)
  totalRow.getCell(1).value = '合計 Total'
  totalRow.getCell(2).value = 'g'
  ;[totalRow.getCell(1), totalRow.getCell(2)].forEach(c => {
    c.font = { bold: true, size: 10, name: 'Arial' }
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: CLR.sectionBg } }
    c.border = thinBorder()
  })
  models.forEach((_, mi) => {
    const startRow = 3
    const endRow = 2 + allParts.length
    const col = getColLetter(3 + mi)
    const c = totalRow.getCell(3 + mi)
    c.value = { formula: `SUM(${col}${startRow}:${col}${endRow})` }
    c.font = { bold: true, size: 10, name: 'Arial' }
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: CLR.sectionBg } }
    c.border = thinBorder()
    c.alignment = { horizontal: 'center' }
  })

  ws.getColumn(1).width = 24
  ws.getColumn(2).width = 8
  models.forEach((_, i) => { ws.getColumn(3 + i).width = 16 })
}

// ──────────────────────────────────────────────
// Density Sheet
// ──────────────────────────────────────────────
function buildDensitySheet(wb: ExcelJS.Workbook, spec: ParsedSpec) {
  const ws = wb.addWorksheet('密度')
  const models = spec.models
  const allParts = spec.allPartNames
  if (models.length === 0) return

  ws.mergeCells(1, 1, 1, 3 + models.length)
  const tc = ws.getCell(1, 1)
  tc.value = `密度明細 ｜ ${spec.fileName} ｜ ${spec.version}`
  tc.font = { bold: true, size: 13, name: 'Arial', color: { argb: 'FFFFFFFF' } }
  tc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: CLR.headerBg } }
  tc.alignment = { horizontal: 'center', vertical: 'middle' }
  ws.getRow(1).height = 28

  const hRow = ws.getRow(2)
  hRow.height = 28
  ;['部位名稱', '單位類型', '單位', ...models.map(m => m.modelName)].forEach((v, i) => {
    const c = hRow.getCell(i + 1)
    c.value = v
    Object.assign(c, headerStyle(wb, true))
  })

  let rowIdx = 3
  for (const part of allParts) {
    // Imperial row
    const impRow = ws.getRow(rowIdx++)
    impRow.getCell(1).value = part
    impRow.getCell(2).value = '英制'
    impRow.getCell(3).value = 'g/in³'
    ;[impRow.getCell(1), impRow.getCell(2), impRow.getCell(3)].forEach(c => {
      c.font = { size: 10, name: 'Arial' }
      c.border = thinBorder()
    })
    models.forEach((m, mi) => {
      const c = impRow.getCell(4 + mi)
      c.value = m.density_imperial[part] ?? ''
      c.font = { size: 10, name: 'Arial' }
      c.border = thinBorder()
      c.alignment = { horizontal: 'center' }
    })

    // Metric row
    const metRow = ws.getRow(rowIdx++)
    metRow.getCell(1).value = ''
    metRow.getCell(2).value = '公制'
    metRow.getCell(3).value = 'g/cm³'
    ;[metRow.getCell(1), metRow.getCell(2), metRow.getCell(3)].forEach(c => {
      c.font = { size: 10, name: 'Arial' }
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: CLR.altRow } }
      c.border = thinBorder()
    })
    models.forEach((m, mi) => {
      const c = metRow.getCell(4 + mi)
      const val = m.density_metric[part] ?? null
      c.value = val === null ? '' : val
      c.font = { size: 10, name: 'Arial' }
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: CLR.altRow } }
      c.border = thinBorder()
      c.alignment = { horizontal: 'center' }
    })
  }

  ws.getColumn(1).width = 24
  ws.getColumn(2).width = 10
  ws.getColumn(3).width = 10
  models.forEach((_, i) => { ws.getColumn(4 + i).width = 16 })
}

// ──────────────────────────────────────────────
// Compare Sheet
// ──────────────────────────────────────────────
function buildCompareSheet(
  wb: ExcelJS.Workbook,
  specA: ParsedSpec,
  specB: ParsedSpec,
  labelA: string,
  labelB: string
) {
  const ws = wb.addWorksheet('版次比對')

  const allModelNames = Array.from(
    new Set([...specA.models.map(m => m.modelName), ...specB.models.map(m => m.modelName)])
  )
  const allParts = Array.from(
    new Set([...specA.allPartNames, ...specB.allPartNames])
  )

  // Title
  ws.mergeCells(1, 1, 1, 5)
  const tc = ws.getCell(1, 1)
  tc.value = `版次比對 ｜ ${labelA} vs ${labelB}`
  tc.font = { bold: true, size: 13, name: 'Arial', color: { argb: 'FFFFFFFF' } }
  tc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: CLR.headerBg } }
  tc.alignment = { horizontal: 'center', vertical: 'middle' }
  ws.getRow(1).height = 28

  // Legend
  ws.getRow(2).height = 18
  ;[
    { col: 1, text: '🟡 黃色 = 數值變更', bg: CLR.changed },
    { col: 2, text: '🟢 綠色 = 新增', bg: CLR.added },
    { col: 3, text: '🔴 橘色 = 移除', bg: CLR.removed },
  ].forEach(({ col, text, bg }) => {
    const c = ws.getCell(2, col)
    c.value = text
    c.font = { size: 9, name: 'Arial' }
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } }
    c.border = thinBorder()
  })

  // Headers per model
  let rowIdx = 3

  for (const modelName of allModelNames) {
    const mA = specA.models.find(m => m.modelName === modelName)
    const mB = specB.models.find(m => m.modelName === modelName)

    // Model name header
    ws.mergeCells(rowIdx, 1, rowIdx, 5)
    const mHeader = ws.getCell(rowIdx, 1)
    mHeader.value = `▶ ${modelName}`
    mHeader.font = { bold: true, size: 11, name: 'Arial', color: { argb: 'FFFFFFFF' } }
    mHeader.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: CLR.subHeaderBg } }
    mHeader.alignment = { vertical: 'middle' }
    ws.getRow(rowIdx).height = 22
    rowIdx++

    // Sub headers
    const sh = ws.getRow(rowIdx++)
    ;['類別', '項目', '單位', labelA, labelB].forEach((v, i) => {
      const c = sh.getCell(i + 1)
      c.value = v
      Object.assign(c, headerStyle(wb, false))
    })

    // Spec rows
    const specFields: [string, string, (m: ModelData) => number | null][] = [
      ['規格', 'Final Head Weight (g)', m => m.finalHeadWeight],
      ['規格', 'Loft Angle (°)', m => m.loftAngle],
      ['規格', 'Lie Angle (°)', m => m.lieAngle],
      ['規格', 'Hosel OD (in)', m => m.hoselOD_in],
      ['規格', 'Hosel OD (mm)', m => m.hoselOD_mm],
      ['規格', 'F1/E Distance (in)', m => m.f1e_in],
      ['規格', 'F1/E Distance (mm)', m => m.f1e_mm],
    ]

    for (const [cat, label, getter] of specFields) {
      const valA = mA ? getter(mA) : null
      const valB = mB ? getter(mB) : null
      const changed = valA !== null && valB !== null && Math.abs((valA ?? 0) - (valB ?? 0)) > 0.0001
      writeCompareRow(ws, rowIdx++, cat, label, '', valA, valB, changed, !mA, !mB)
    }

    // Parts
    for (const part of allParts) {
      const valA = mA?.parts[part] ?? null
      const valB = mB?.parts[part] ?? null
      const isNew = valA === null && valB !== null
      const isRemoved = valA !== null && valB === null
      const changed = valA !== null && valB !== null && Math.abs(valA - valB) > 0.0001
      writeCompareRow(ws, rowIdx++, '重量', part, 'g', valA, valB, changed, isNew, isRemoved)
    }

    rowIdx++ // blank row between models
  }

  ws.getColumn(1).width = 10
  ws.getColumn(2).width = 28
  ws.getColumn(3).width = 8
  ws.getColumn(4).width = 18
  ws.getColumn(5).width = 18
}

function writeCompareRow(
  ws: ExcelJS.Worksheet,
  rowIdx: number,
  cat: string,
  label: string,
  unit: string,
  valA: number | null,
  valB: number | null,
  changed: boolean,
  isNew: boolean,
  isRemoved: boolean
) {
  const row = ws.getRow(rowIdx)
  const bgColor = isNew ? CLR.added : isRemoved ? CLR.removed : changed ? CLR.changed : null

  const vals = [cat, label, unit, valA === null ? '' : valA, valB === null ? '' : valB]
  vals.forEach((v, i) => {
    const c = row.getCell(i + 1)
    c.value = v
    c.font = { size: 10, name: 'Arial', bold: (changed || isNew || isRemoved) && i >= 3 }
    c.border = thinBorder()
    c.alignment = i >= 3 ? { horizontal: 'center' } : { horizontal: 'left' }
    if (bgColor) c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgColor } }
  })
}

function fmt4(v: number | null): number | string {
  if (v === null) return ''
  return Math.round(v * 10000) / 10000
}

function getColLetter(n: number): string {
  let col = ''
  while (n > 0) {
    const rem = (n - 1) % 26
    col = String.fromCharCode(65 + rem) + col
    n = Math.floor((n - 1) / 26)
  }
  return col
}
