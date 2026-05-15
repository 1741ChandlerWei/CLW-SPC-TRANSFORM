import ExcelJS from 'exceljs'
import { ParsedSpec, ModelData } from './parser'

const C = {
  navy:      'FF1F4E79',
  navyFg:    'FFFFFFFF',
  lightBlue: 'FFD6E4F0',
  sectionBg: 'FFBDD7EE',
  rowAlt:    'FFDEEAF1',
  changed:   'FFFFFF00',
  added:     'FF92D050',
  removed:   'FFFF7F7F',
  border:    'FF9DC3E6',
}

function thin() {
  const s = { style: 'thin' as const, color: { argb: C.border } }
  return { top: s, left: s, bottom: s, right: s }
}
function applyHeader(cell: ExcelJS.Cell, dark = true) {
  cell.font = { bold: true, size: 10, name: 'Arial', color: { argb: dark ? C.navyFg : 'FF000000' } }
  cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: dark ? C.navy : C.lightBlue } }
  cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }
  cell.border = thin()
}
function applySection(cell: ExcelJS.Cell) {
  cell.font = { bold: true, size: 10, name: 'Arial' }
  cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.sectionBg } }
  cell.alignment = { horizontal: 'center', vertical: 'middle' }
  cell.border = thin()
}
function applyLabel(cell: ExcelJS.Cell) {
  cell.font = { size: 10, name: 'Arial' }
  cell.alignment = { horizontal: 'left', vertical: 'middle' }
  cell.border = thin()
}
function applyValue(cell: ExcelJS.Cell, alt = false) {
  cell.font = { size: 10, name: 'Arial' }
  cell.alignment = { horizontal: 'center', vertical: 'middle' }
  cell.border = thin()
  if (alt) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.rowAlt } }
}
function r4(v: number | null | undefined): number | string {
  if (v === null || v === undefined) return ''
  return Math.round(v * 10000) / 10000
}

export async function buildOutputExcel(spec: ParsedSpec): Promise<Buffer> {
  const wb = new ExcelJS.Workbook()
  wb.creator = 'CLW-SPC-TRANSFORM'
  const ws = wb.addWorksheet('整理')
  const models = spec.models
  const allParts = spec.allPartNames
  const nM = models.length

  ws.getColumn(1).width = 12
  ws.getColumn(2).width = 26
  for (let i = 0; i < nM; i++) ws.getColumn(3 + i).width = 18

  let r = 1

  // Row 1: 版次 + 檔名
  ws.getRow(r).height = 22
  ws.mergeCells(r, 1, r, 2)
  applyHeader(ws.getCell(r, 1), true)
  ws.getCell(r, 1).value = '版次'
  ws.mergeCells(r, 3, r, 2 + nM)
  applyHeader(ws.getCell(r, 3), true)
  ws.getCell(r, 3).value = spec.fileName + (spec.version ? `  ｜  ${spec.version}` : '')
  r++

  // Row 2: 型號 header
  ws.getRow(r).height = 22
  ws.mergeCells(r, 1, r, 2)
  applyHeader(ws.getCell(r, 1), false)
  ws.getCell(r, 1).value = '型號'
  models.forEach((m, i) => {
    applyHeader(ws.getCell(r, 3 + i), false)
    ws.getCell(r, 3 + i).value = m.modelName
  })
  r++

  // 規格 rows
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
    applySection(ws.getCell(r, 1))
    ws.getCell(r, 1).value = idx === 0 ? '規格' : ''
    applyLabel(ws.getCell(r, 2))
    ws.getCell(r, 2).value = label
    models.forEach((m, i) => {
      applyValue(ws.getCell(r, 3 + i), idx % 2 === 1)
      ws.getCell(r, 3 + i).value = getter(m)
    })
    r++
  })

  // 重量 header
  ws.getRow(r).height = 20
  applySection(ws.getCell(r, 1)); ws.getCell(r, 1).value = '重量'
  applyHeader(ws.getCell(r, 2), false); ws.getCell(r, 2).value = 'Name'
  models.forEach((_, i) => { applyHeader(ws.getCell(r, 3 + i), false); ws.getCell(r, 3 + i).value = 'Mass (g)' })
  r++
  allParts.forEach((part, pi) => {
    ws.getRow(r).height = 18
    applySection(ws.getCell(r, 1)); ws.getCell(r, 1).value = ''
    applyLabel(ws.getCell(r, 2)); ws.getCell(r, 2).value = part
    models.forEach((m, i) => { applyValue(ws.getCell(r, 3 + i), pi % 2 === 1); ws.getCell(r, 3 + i).value = m.parts[part]?.mass ?? "" })
    r++
  })

  // 密度(英制) header
  ws.getRow(r).height = 20
  applySection(ws.getCell(r, 1)); ws.getCell(r, 1).value = '密度  (英制)'
  applyHeader(ws.getCell(r, 2), false); ws.getCell(r, 2).value = 'Name'
  models.forEach((_, i) => { applyHeader(ws.getCell(r, 3 + i), false); ws.getCell(r, 3 + i).value = 'Density (g/in3)' })
  r++
  allParts.forEach((part, pi) => {
    ws.getRow(r).height = 18
    applySection(ws.getCell(r, 1)); ws.getCell(r, 1).value = ''
    applyLabel(ws.getCell(r, 2)); ws.getCell(r, 2).value = part
    models.forEach((m, i) => { applyValue(ws.getCell(r, 3 + i), pi % 2 === 1); ws.getCell(r, 3 + i).value = m.parts[part]?.density_imp ?? '' })
    r++
  })

  // 密度(公制) header
  ws.getRow(r).height = 20
  applySection(ws.getCell(r, 1)); ws.getCell(r, 1).value = '密度  (公制)'
  applyHeader(ws.getCell(r, 2), false); ws.getCell(r, 2).value = 'Name'
  models.forEach((_, i) => { applyHeader(ws.getCell(r, 3 + i), false); ws.getCell(r, 3 + i).value = 'Density (g/cm3)' })
  r++
  allParts.forEach((part, pi) => {
    ws.getRow(r).height = 18
    applySection(ws.getCell(r, 1)); ws.getCell(r, 1).value = ''
    applyLabel(ws.getCell(r, 2)); ws.getCell(r, 2).value = part
    models.forEach((m, i) => { applyValue(ws.getCell(r, 3 + i), pi % 2 === 1); ws.getCell(r, 3 + i).value = m.parts[part]?.density_met ?? '' })
    r++
  })

  ws.views = [{ state: 'frozen', xSplit: 2, ySplit: 2, activeCell: 'C3' }]
  const buf = await wb.xlsx.writeBuffer()
  return Buffer.from(buf)
}

export async function buildCompareExcel(specA: ParsedSpec, specB: ParsedSpec, labelA: string, labelB: string): Promise<Buffer> {
  const wb = new ExcelJS.Workbook()
  wb.creator = 'CLW-SPC-TRANSFORM'
  const ws = wb.addWorksheet('版次比對')

  const allModelNames = Array.from(new Set([...specA.models.map(m => m.modelName), ...specB.models.map(m => m.modelName)]))
  const allParts = Array.from(new Set([...specA.allPartNames, ...specB.allPartNames]))
  const nM = allModelNames.length

  ws.getColumn(1).width = 12
  ws.getColumn(2).width = 26
  for (let i = 0; i < nM; i++) { ws.getColumn(3 + i * 2).width = 16; ws.getColumn(4 + i * 2).width = 16 }

  let r = 1

  ws.getRow(r).height = 22
  ws.mergeCells(r, 1, r, 2); applyHeader(ws.getCell(r, 1), true); ws.getCell(r, 1).value = '版次比對'
  ws.mergeCells(r, 3, r, 2 + nM * 2); applyHeader(ws.getCell(r, 3), true); ws.getCell(r, 3).value = `${labelA}  vs  ${labelB}`
  r++

  ws.getRow(r).height = 22
  ws.mergeCells(r, 1, r, 2); applyHeader(ws.getCell(r, 1), false); ws.getCell(r, 1).value = '型號'
  allModelNames.forEach((name, mi) => {
    ws.mergeCells(r, 3 + mi * 2, r, 4 + mi * 2)
    applyHeader(ws.getCell(r, 3 + mi * 2), false); ws.getCell(r, 3 + mi * 2).value = name
  })
  r++

  ws.getRow(r).height = 18
  applySection(ws.getCell(r, 1)); ws.getCell(r, 1).value = ''
  applyHeader(ws.getCell(r, 2), false); ws.getCell(r, 2).value = '項目'
  allModelNames.forEach((_, mi) => {
    applyHeader(ws.getCell(r, 3 + mi * 2), false); ws.getCell(r, 3 + mi * 2).value = specA.version || '版本A'
    applyHeader(ws.getCell(r, 4 + mi * 2), false); ws.getCell(r, 4 + mi * 2).value = specB.version || '版本B'
  })
  r++

  function writeCompRow(label: string, cat: string, showCat: boolean,
    getA: (m: ModelData | undefined) => number | string | null,
    getB: (m: ModelData | undefined) => number | string | null, alt: boolean) {
    ws.getRow(r).height = 18
    applySection(ws.getCell(r, 1)); ws.getCell(r, 1).value = showCat ? cat : ''
    applyLabel(ws.getCell(r, 2)); ws.getCell(r, 2).value = label
    allModelNames.forEach((name, mi) => {
      const mA = specA.models.find(m => m.modelName === name)
      const mB = specB.models.find(m => m.modelName === name)
      const vA = getA(mA); const vB = getB(mB)
      const nA = parseFloat(String(vA)); const nB = parseFloat(String(vB))
      const changed = !isNaN(nA) && !isNaN(nB) && Math.abs(nA - nB) > 0.0001
      const isNew = (vA === null || vA === '') && vB !== null && vB !== ''
      const isRemoved = vA !== null && vA !== '' && (vB === null || vB === '')
      const cA = ws.getCell(r, 3 + mi * 2); const cB = ws.getCell(r, 4 + mi * 2)
      cA.value = vA ?? ''; cB.value = vB ?? ''
      applyValue(cA, alt); applyValue(cB, alt)
      if (changed) {
        cA.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.changed } }
        cB.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.changed } }
        cB.font = { bold: true, size: 10, name: 'Arial' }
      }
      if (isNew) cB.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.added } }
      if (isRemoved) cA.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.removed } }
    })
    r++
  }

  const specDefs: [string, (m: ModelData | undefined) => number | string | null][] = [
    ['Final Head Weight', m => r4(m?.finalHeadWeight)],
    ['Loft Angle',        m => r4(m?.loftAngle)],
    ['Lie Angle',         m => r4(m?.lieAngle)],
    ['Hosel OD(英制)',    m => r4(m?.hoselOD_in)],
    ['Hosel OD(公制)',    m => r4(m?.hoselOD_mm)],
    ['F1/E  Distance(英制)', m => r4(m?.f1e_in)],
    ['F1/E  Distance(公制)', m => r4(m?.f1e_mm)],
  ]
  specDefs.forEach(([label, getter], idx) => writeCompRow(label, '規格', idx === 0, getter, getter, idx % 2 === 1))

  ws.getRow(r).height = 20
  applySection(ws.getCell(r, 1)); ws.getCell(r, 1).value = '重量'
  applyHeader(ws.getCell(r, 2), false); ws.getCell(r, 2).value = 'Name'
  allModelNames.forEach((_, mi) => {
    applyHeader(ws.getCell(r, 3 + mi * 2), false); ws.getCell(r, 3 + mi * 2).value = 'Mass (g)'
    applyHeader(ws.getCell(r, 4 + mi * 2), false); ws.getCell(r, 4 + mi * 2).value = 'Mass (g)'
  }); r++
  allParts.forEach((part, pi) => writeCompRow(part, '重量', false, m => m?.parts[part]?.mass ?? null, m => m?.parts[part]?.mass ?? null, pi % 2 === 1))

  ws.getRow(r).height = 20
  applySection(ws.getCell(r, 1)); ws.getCell(r, 1).value = '密度  (英制)'
  applyHeader(ws.getCell(r, 2), false); ws.getCell(r, 2).value = 'Name'
  allModelNames.forEach((_, mi) => {
    applyHeader(ws.getCell(r, 3 + mi * 2), false); ws.getCell(r, 3 + mi * 2).value = 'Density (g/in3)'
    applyHeader(ws.getCell(r, 4 + mi * 2), false); ws.getCell(r, 4 + mi * 2).value = 'Density (g/in3)'
  }); r++
  allParts.forEach((part, pi) => writeCompRow(part, '密度  (英制)', false, m => m?.parts[part]?.density_imp ?? null, m => m?.parts[part]?.density_imp ?? null, pi % 2 === 1))

  ws.getRow(r).height = 20
  applySection(ws.getCell(r, 1)); ws.getCell(r, 1).value = '密度  (公制)'
  applyHeader(ws.getCell(r, 2), false); ws.getCell(r, 2).value = 'Name'
  allModelNames.forEach((_, mi) => {
    applyHeader(ws.getCell(r, 3 + mi * 2), false); ws.getCell(r, 3 + mi * 2).value = 'Density (g/cm3)'
    applyHeader(ws.getCell(r, 4 + mi * 2), false); ws.getCell(r, 4 + mi * 2).value = 'Density (g/cm3)'
  }); r++
  allParts.forEach((part, pi) => writeCompRow(part, '密度  (公制)', false, m => m?.parts[part]?.density_met ?? null, m => m?.parts[part]?.density_met ?? null, pi % 2 === 1))

  ws.views = [{ state: 'frozen', xSplit: 2, ySplit: 3, activeCell: 'C4' }]
  const buf = await wb.xlsx.writeBuffer()
  return Buffer.from(buf)
}
