import ExcelJS from 'exceljs'
import { ParsedSpec, ModelData } from './parser'

// ── Colors ──
const NAVY    = 'FF1F4E79'
const NAVYFG  = 'FFFFFFFF'
const LBLUE   = 'FFD6E4F0'
const SECBG   = 'FFBDD7EE'
const ALT     = 'FFDEEAF1'
const WHITE   = 'FFFFFFFF'
const YELLOW  = 'FFFFFF00'
const GREEN   = 'FF92D050'
const ORANGE  = 'FFFF7F7F'
const BORDER  = 'FF9DC3E6'

function thin() {
  const s = { style: 'thin' as const, color: { argb: BORDER } }
  return { top: s, left: s, bottom: s, right: s }
}

function styleCell(cell: ExcelJS.Cell, opts: {
  bold?: boolean
  color?: string
  bg?: string
  hAlign?: ExcelJS.Alignment['horizontal']
  vAlign?: ExcelJS.Alignment['vertical']
  wrap?: boolean
  size?: number
}) {
  cell.font = { name: 'Arial', size: opts.size ?? 10, bold: opts.bold ?? false, color: { argb: opts.color ?? 'FF000000' } }
  if (opts.bg) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: opts.bg } }
  cell.alignment = { horizontal: opts.hAlign ?? 'left', vertical: opts.vAlign ?? 'middle', wrapText: opts.wrap ?? false }
  cell.border = thin()
}

function r4(v: number | null | undefined): number | string {
  if (v === null || v === undefined) return ''
  return Math.round(v * 10000) / 10000
}

// ─────────────────────────────────────────────────────
// V1.8 分群輸出：相同部件清單的型號放同一個 Sheet
// ─────────────────────────────────────────────────────

/** 以部件名稱清單（排序後 join）為 key，將型號動態分群 */
function groupModelsByParts(models: ModelData[]): ModelData[][] {
  const groups = new Map<string, ModelData[]>()
  for (const m of models) {
    const key = Object.keys(m.parts).sort().join('|')
    if (!groups.has(key)) groups.set(key, [])
    groups.get(key)!.push(m)
  }
  return Array.from(groups.values())
}

/** 產生 Sheet 名稱（最多 31 字元，Excel 限制）*/
function makeSheetName(groupModels: ModelData[]): string {
  const joined = groupModels.map(m => m.modelName).join(' / ')
  return joined.length <= 31 ? joined : joined.slice(0, 28) + '...'
}

/** 將一群型號寫入單一 Worksheet */
function writeGroupSheet(
  ws: ExcelJS.Worksheet,
  models: ModelData[],
  spec: ParsedSpec
) {
  // 該群的部件清單：依解析時的出現順序
  const partSet = new Set<string>()
  for (const m of models) Object.keys(m.parts).forEach(p => partSet.add(p))
  const groupParts = Array.from(partSet)

  const nM = models.length

  ws.getColumn(1).width = 12
  ws.getColumn(2).width = 26
  for (let i = 0; i < nM; i++) ws.getColumn(3 + i).width = 16

  let r = 1

  // ── Row 1: 版次 header ──
  ws.getRow(r).height = 24
  const a1 = ws.getCell(r, 1)
  a1.value = '版次'
  styleCell(a1, { bold: true, color: NAVYFG, bg: NAVY, hAlign: 'center' })
  const b1 = ws.getCell(r, 2)
  b1.value = ''
  styleCell(b1, { bold: true, color: NAVYFG, bg: NAVY, hAlign: 'center' })
  ws.mergeCells(r, 3, r, 2 + nM)
  const c1 = ws.getCell(r, 3)
  c1.value = spec.fileName + (spec.version ? `  ｜  ${spec.version}` : '')
  styleCell(c1, { bold: true, color: NAVYFG, bg: NAVY, hAlign: 'center', size: 11 })
  r++

  // ── Row 2: 型號 ──
  ws.getRow(r).height = 22
  const a2 = ws.getCell(r, 1)
  a2.value = '型號'
  styleCell(a2, { bold: true, bg: LBLUE, hAlign: 'center' })
  const b2 = ws.getCell(r, 2)
  b2.value = ''
  styleCell(b2, { bold: true, bg: LBLUE, hAlign: 'center' })
  models.forEach((m, i) => {
    const c = ws.getCell(r, 3 + i)
    c.value = m.modelName
    styleCell(c, { bold: true, bg: LBLUE, hAlign: 'center', wrap: true })
  })
  r++

  // ── 規格 rows ──
  const specDefs: [string, (m: ModelData) => number | string][] = [
    ['Final Head Weight', m => r4(m.finalHeadWeight)],
    ['Loft Angle',        m => r4(m.loftAngle)],
    ['Lie Angle',         m => r4(m.lieAngle)],
    ['Hosel OD(英制)',    m => r4(m.hoselOD_in)],
    ['Hosel OD(公制)',    m => r4(m.hoselOD_mm)],
    ['F1/E  Distance(英制)', m => r4(m.f1e_in)],
    ['F1/E  Distance(公制)', m => r4(m.f1e_mm)],
  ]

  specDefs.forEach(([label, getter], idx) => {
    ws.getRow(r).height = 18
    const bg = idx % 2 === 1 ? ALT : WHITE
    const ca = ws.getCell(r, 1)
    ca.value = idx === 0 ? '規格' : ''
    styleCell(ca, { bold: true, bg: SECBG, hAlign: 'center' })
    const cb = ws.getCell(r, 2)
    cb.value = label
    styleCell(cb, { bg })
    models.forEach((m, i) => {
      const cv = ws.getCell(r, 3 + i)
      cv.value = getter(m)
      styleCell(cv, { bg, hAlign: 'center' })
    })
    r++
  })

  // ── 重量 section header ──
  ws.getRow(r).height = 20
  ws.getCell(r, 1).value = '重量'
  styleCell(ws.getCell(r, 1), { bold: true, bg: SECBG, hAlign: 'center' })
  ws.getCell(r, 2).value = 'Name'
  styleCell(ws.getCell(r, 2), { bold: true, bg: LBLUE, hAlign: 'center' })
  models.forEach((_, i) => {
    ws.getCell(r, 3 + i).value = 'Mass (g)'
    styleCell(ws.getCell(r, 3 + i), { bold: true, bg: LBLUE, hAlign: 'center' })
  })
  r++

  groupParts.forEach((part, pi) => {
    ws.getRow(r).height = 18
    const bg = pi % 2 === 1 ? ALT : WHITE
    ws.getCell(r, 1).value = ''
    styleCell(ws.getCell(r, 1), { bg: SECBG })
    ws.getCell(r, 2).value = part
    styleCell(ws.getCell(r, 2), { bg })
    models.forEach((m, i) => {
      const cv = ws.getCell(r, 3 + i)
      cv.value = m.parts[part]?.mass ?? ''
      styleCell(cv, { bg, hAlign: 'center' })
    })
    r++
  })

  // ── 密度(英制) header ──
  ws.getRow(r).height = 20
  ws.getCell(r, 1).value = '密度  (英制)'
  styleCell(ws.getCell(r, 1), { bold: true, bg: SECBG, hAlign: 'center' })
  ws.getCell(r, 2).value = 'Name'
  styleCell(ws.getCell(r, 2), { bold: true, bg: LBLUE, hAlign: 'center' })
  models.forEach((_, i) => {
    ws.getCell(r, 3 + i).value = 'Density (g/in3)'
    styleCell(ws.getCell(r, 3 + i), { bold: true, bg: LBLUE, hAlign: 'center' })
  })
  r++

  groupParts.forEach((part, pi) => {
    ws.getRow(r).height = 18
    const bg = pi % 2 === 1 ? ALT : WHITE
    ws.getCell(r, 1).value = ''
    styleCell(ws.getCell(r, 1), { bg: SECBG })
    ws.getCell(r, 2).value = part
    styleCell(ws.getCell(r, 2), { bg })
    models.forEach((m, i) => {
      const cv = ws.getCell(r, 3 + i)
      cv.value = m.parts[part]?.density_imp ?? ''
      styleCell(cv, { bg, hAlign: 'center' })
    })
    r++
  })

  // ── 密度(公制) header ──
  ws.getRow(r).height = 20
  ws.getCell(r, 1).value = '密度  (公制)'
  styleCell(ws.getCell(r, 1), { bold: true, bg: SECBG, hAlign: 'center' })
  ws.getCell(r, 2).value = 'Name'
  styleCell(ws.getCell(r, 2), { bold: true, bg: LBLUE, hAlign: 'center' })
  models.forEach((_, i) => {
    ws.getCell(r, 3 + i).value = 'Density (g/cm3)'
    styleCell(ws.getCell(r, 3 + i), { bold: true, bg: LBLUE, hAlign: 'center' })
  })
  r++

  groupParts.forEach((part, pi) => {
    ws.getRow(r).height = 18
    const bg = pi % 2 === 1 ? ALT : WHITE
    ws.getCell(r, 1).value = ''
    styleCell(ws.getCell(r, 1), { bg: SECBG })
    ws.getCell(r, 2).value = part
    styleCell(ws.getCell(r, 2), { bg })
    models.forEach((m, i) => {
      const cv = ws.getCell(r, 3 + i)
      cv.value = m.parts[part]?.density_met ?? ''
      styleCell(cv, { bg, hAlign: 'center' })
    })
    r++
  })

  ws.views = [{ state: 'frozen', xSplit: 2, ySplit: 2, activeCell: 'C3' }]
}

export async function buildOutputExcel(spec: ParsedSpec): Promise<Buffer> {
  const wb = new ExcelJS.Workbook()
  wb.creator = 'CLW-SPC-TRANSFORM'

  const groups = groupModelsByParts(spec.models)

  // 確保 Sheet 名稱不重複
  const usedNames = new Set<string>()
  for (const group of groups) {
    let name = makeSheetName(group)
    let suffix = 2
    while (usedNames.has(name)) {
      name = makeSheetName(group).slice(0, 28) + `(${suffix++})`
    }
    usedNames.add(name)
    const ws = wb.addWorksheet(name)
    writeGroupSheet(ws, group, spec)
  }

  const buf = await wb.xlsx.writeBuffer()
  return Buffer.from(buf)
}


// ─────────────────────────────────────────────────────
// Compare output
// ─────────────────────────────────────────────────────
export async function buildCompareExcel(
  specA: ParsedSpec,
  specB: ParsedSpec,
  labelA: string,
  labelB: string
): Promise<Buffer> {
  const wb = new ExcelJS.Workbook()
  wb.creator = 'CLW-SPC-TRANSFORM'
  const ws = wb.addWorksheet('版次比對')

  const allModelNames = Array.from(new Set([
    ...specA.models.map(m => m.modelName),
    ...specB.models.map(m => m.modelName)
  ]))
  const allParts = Array.from(new Set([
    ...specA.allPartNames,
    ...specB.allPartNames
  ]))
  const nM = allModelNames.length

  ws.getColumn(1).width = 12
  ws.getColumn(2).width = 26
  for (let i = 0; i < nM; i++) {
    ws.getColumn(3 + i * 2).width = 16
    ws.getColumn(4 + i * 2).width = 16
  }

  let r = 1

  // Row 1: title
  ws.getRow(r).height = 24
  ws.getCell(r, 1).value = '版次比對'
  styleCell(ws.getCell(r, 1), { bold: true, color: NAVYFG, bg: NAVY, hAlign: 'center' })
  ws.getCell(r, 2).value = ''
  styleCell(ws.getCell(r, 2), { bold: true, color: NAVYFG, bg: NAVY })
  ws.mergeCells(r, 3, r, 2 + nM * 2)
  ws.getCell(r, 3).value = `${labelA}  vs  ${labelB}`
  styleCell(ws.getCell(r, 3), { bold: true, color: NAVYFG, bg: NAVY, hAlign: 'center' })
  r++

  // Row 2: model names (each spans 2 cols)
  ws.getRow(r).height = 22
  ws.getCell(r, 1).value = '型號'
  styleCell(ws.getCell(r, 1), { bold: true, bg: LBLUE, hAlign: 'center' })
  ws.getCell(r, 2).value = ''
  styleCell(ws.getCell(r, 2), { bold: true, bg: LBLUE })
  allModelNames.forEach((name, mi) => {
    ws.mergeCells(r, 3 + mi * 2, r, 4 + mi * 2)
    ws.getCell(r, 3 + mi * 2).value = name
    styleCell(ws.getCell(r, 3 + mi * 2), { bold: true, bg: LBLUE, hAlign: 'center', wrap: true })
  })
  r++

  // Row 3: A/B sub-headers
  ws.getRow(r).height = 18
  ws.getCell(r, 1).value = ''
  styleCell(ws.getCell(r, 1), { bg: SECBG })
  ws.getCell(r, 2).value = '項目'
  styleCell(ws.getCell(r, 2), { bold: true, bg: LBLUE, hAlign: 'center' })
  allModelNames.forEach((_, mi) => {
    ws.getCell(r, 3 + mi * 2).value = specA.version || '版本A'
    styleCell(ws.getCell(r, 3 + mi * 2), { bold: true, bg: LBLUE, hAlign: 'center' })
    ws.getCell(r, 4 + mi * 2).value = specB.version || '版本B'
    styleCell(ws.getCell(r, 4 + mi * 2), { bold: true, bg: LBLUE, hAlign: 'center' })
  })
  r++

  function writeCompRow(
    label: string, cat: string, showCat: boolean,
    getA: (m: ModelData | undefined) => number | string | null,
    getB: (m: ModelData | undefined) => number | string | null,
    alt: boolean
  ) {
    ws.getRow(r).height = 18
    const baseBg = alt ? ALT : WHITE
    ws.getCell(r, 1).value = showCat ? cat : ''
    styleCell(ws.getCell(r, 1), { bold: showCat, bg: SECBG, hAlign: 'center' })
    ws.getCell(r, 2).value = label
    styleCell(ws.getCell(r, 2), { bg: baseBg })

    allModelNames.forEach((name, mi) => {
      const mA = specA.models.find(m => m.modelName === name)
      const mB = specB.models.find(m => m.modelName === name)
      const vA = getA(mA)
      const vB = getB(mB)
      const nA = parseFloat(String(vA))
      const nB = parseFloat(String(vB))
      const changed = !isNaN(nA) && !isNaN(nB) && Math.abs(nA - nB) > 0.0001
      const isNew = (vA === null || vA === '') && vB !== null && vB !== ''
      const isRemoved = vA !== null && vA !== '' && (vB === null || vB === '')

      const bgA = isRemoved ? ORANGE : changed ? YELLOW : baseBg
      const bgB = isNew ? GREEN : changed ? YELLOW : baseBg

      ws.getCell(r, 3 + mi * 2).value = vA ?? ''
      styleCell(ws.getCell(r, 3 + mi * 2), { bg: bgA, hAlign: 'center', bold: changed || isRemoved })
      ws.getCell(r, 4 + mi * 2).value = vB ?? ''
      styleCell(ws.getCell(r, 4 + mi * 2), { bg: bgB, hAlign: 'center', bold: changed || isNew })
    })
    r++
  }

  // Spec rows
  const specDefs: [string, (m: ModelData | undefined) => number | string | null][] = [
    ['Final Head Weight', m => r4(m?.finalHeadWeight)],
    ['Loft Angle',        m => r4(m?.loftAngle)],
    ['Lie Angle',         m => r4(m?.lieAngle)],
    ['Hosel OD(英制)',    m => r4(m?.hoselOD_in)],
    ['Hosel OD(公制)',    m => r4(m?.hoselOD_mm)],
    ['F1/E  Distance(英制)', m => r4(m?.f1e_in)],
    ['F1/E  Distance(公制)', m => r4(m?.f1e_mm)],
  ]
  specDefs.forEach(([label, getter], idx) => {
    writeCompRow(label, '規格', idx === 0, getter, getter, idx % 2 === 1)
  })

  // Weight header
  ws.getRow(r).height = 20
  ws.getCell(r, 1).value = '重量'
  styleCell(ws.getCell(r, 1), { bold: true, bg: SECBG, hAlign: 'center' })
  ws.getCell(r, 2).value = 'Name'
  styleCell(ws.getCell(r, 2), { bold: true, bg: LBLUE, hAlign: 'center' })
  allModelNames.forEach((_, mi) => {
    ws.getCell(r, 3 + mi * 2).value = 'Mass (g)'
    styleCell(ws.getCell(r, 3 + mi * 2), { bold: true, bg: LBLUE, hAlign: 'center' })
    ws.getCell(r, 4 + mi * 2).value = 'Mass (g)'
    styleCell(ws.getCell(r, 4 + mi * 2), { bold: true, bg: LBLUE, hAlign: 'center' })
  }); r++
  allParts.forEach((part, pi) => {
    writeCompRow(part, '重量', false,
      m => m?.parts[part]?.mass ?? null,
      m => m?.parts[part]?.mass ?? null,
      pi % 2 === 1)
  })

  // Density imperial header
  ws.getRow(r).height = 20
  ws.getCell(r, 1).value = '密度  (英制)'
  styleCell(ws.getCell(r, 1), { bold: true, bg: SECBG, hAlign: 'center' })
  ws.getCell(r, 2).value = 'Name'
  styleCell(ws.getCell(r, 2), { bold: true, bg: LBLUE, hAlign: 'center' })
  allModelNames.forEach((_, mi) => {
    ws.getCell(r, 3 + mi * 2).value = 'Density (g/in3)'
    styleCell(ws.getCell(r, 3 + mi * 2), { bold: true, bg: LBLUE, hAlign: 'center' })
    ws.getCell(r, 4 + mi * 2).value = 'Density (g/in3)'
    styleCell(ws.getCell(r, 4 + mi * 2), { bold: true, bg: LBLUE, hAlign: 'center' })
  }); r++
  allParts.forEach((part, pi) => {
    writeCompRow(part, '密度  (英制)', false,
      m => m?.parts[part]?.density_imp ?? null,
      m => m?.parts[part]?.density_imp ?? null,
      pi % 2 === 1)
  })

  // Density metric header
  ws.getRow(r).height = 20
  ws.getCell(r, 1).value = '密度  (公制)'
  styleCell(ws.getCell(r, 1), { bold: true, bg: SECBG, hAlign: 'center' })
  ws.getCell(r, 2).value = 'Name'
  styleCell(ws.getCell(r, 2), { bold: true, bg: LBLUE, hAlign: 'center' })
  allModelNames.forEach((_, mi) => {
    ws.getCell(r, 3 + mi * 2).value = 'Density (g/cm3)'
    styleCell(ws.getCell(r, 3 + mi * 2), { bold: true, bg: LBLUE, hAlign: 'center' })
    ws.getCell(r, 4 + mi * 2).value = 'Density (g/cm3)'
    styleCell(ws.getCell(r, 4 + mi * 2), { bold: true, bg: LBLUE, hAlign: 'center' })
  }); r++
  allParts.forEach((part, pi) => {
    writeCompRow(part, '密度  (公制)', false,
      m => m?.parts[part]?.density_met ?? null,
      m => m?.parts[part]?.density_met ?? null,
      pi % 2 === 1)
  })

  ws.views = [{ state: 'frozen', xSplit: 2, ySplit: 3, activeCell: 'C4' }]
  const buf = await wb.xlsx.writeBuffer()
  return Buffer.from(buf)
}
