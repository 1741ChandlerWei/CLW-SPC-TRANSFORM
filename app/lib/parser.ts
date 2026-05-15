// Parser runs server-side only (Node.js) using xlsx with data_only equivalent
// Strategy: use xlsx with cellFormula:false to get computed values

import * as XLSX from 'xlsx'

const IN_TO_MM = 25.4
const G_IN3_TO_G_CM3 = 0.0610237441

export interface PartData {
  mass: number | null
  density_imp: number | null   // g/in³
  density_met: number | null   // g/cm³
}

export interface ModelData {
  modelName: string
  finalHeadWeight: number | null
  loftAngle: number | null
  lieAngle: number | null
  hoselOD_in: number | null
  hoselOD_mm: number | null
  f1e_in: number | null
  f1e_mm: number | null
  parts: Record<string, PartData>
}

export interface ParsedSpec {
  fileName: string
  version: string
  models: ModelData[]
  allPartNames: string[]
}

function toNum(v: unknown): number | null {
  if (v === null || v === undefined || v === '') return null
  const n = typeof v === 'number' ? v : parseFloat(String(v))
  return isNaN(n) ? null : n
}

function r4(v: number | null): number | null {
  if (v === null) return null
  return Math.round(v * 10000) / 10000
}

export function parseSpecFile(buffer: Buffer): ParsedSpec {
  // Read with cellDates and raw values (computed, not formulas)
  const wb = XLSX.read(buffer, {
    type: 'buffer',
    cellFormula: false,   // don't parse formulas, use cached values
    cellNF: false,
    cellStyles: false,
  })

  // ── Version from Revision Record ──
  let version = ''
  if (wb.SheetNames.includes('Revision Record')) {
    const ws = wb.Sheets['Revision Record']
    const rows = XLSX.utils.sheet_to_json<(string | number)[]>(ws, { header: 1, defval: null }) as (string | number | null)[][]
    for (let i = rows.length - 1; i >= 1; i--) {
      const row = rows[i]
      const rev = row?.[0]
      const desc = row?.[1]
      if (rev && String(rev).trim() && desc && String(desc).trim()) {
        version = `Rev ${String(rev).trim()}`
        break
      }
    }
  }

  // ── Identify model sheets: start with 'i', have 'Final Head' in B2 area ──
  const modelSheets = wb.SheetNames.filter(name => {
    if (!name.startsWith('i')) return false
    const ws = wb.Sheets[name]
    if (!ws) return false
    // Check if B2 or nearby has a numeric Final Head Weight
    const rows = XLSX.utils.sheet_to_json<(string | number | null)[]>(ws, {
      header: 1, defval: null, range: 0
    }) as (string | number | null)[][]

    // Look for 'Final Head' label and a numeric value
    for (let i = 0; i < Math.min(15, rows.length); i++) {
      const label = String(rows[i]?.[0] || '')
      if (label.includes('Final Head')) {
        const val = rows[i]?.[1]
        // Must have a numeric cached value (not a formula string)
        if (typeof val === 'number' && val > 0) return true
        // If val is null/string but hosel is present = still a valid model sheet
        const hosel = rows.find(r => String(r?.[0] || '').includes('Hosel OD'))?.[1]
        if (typeof hosel === 'number') return true
      }
    }
    return false
  })

  const models: ModelData[] = []

  for (const sname of modelSheets) {
    const ws = wb.Sheets[sname]
    const rows = XLSX.utils.sheet_to_json<(string | number | null)[]>(ws, {
      header: 1, defval: null
    }) as (string | number | null)[][]

    const model: ModelData = {
      modelName: sname.trim(),
      finalHeadWeight: null,
      loftAngle: null,
      lieAngle: null,
      hoselOD_in: null,
      hoselOD_mm: null,
      f1e_in: null,
      f1e_mm: null,
      parts: {},
    }

    // ── Spec values: rows 0-14, col A = label, col B = value ──
    for (let i = 0; i < Math.min(20, rows.length); i++) {
      const label = String(rows[i]?.[0] || '').trim()
      const val = toNum(rows[i]?.[1])

      if (label.includes('Final Head')) model.finalHeadWeight = val
      else if (label === 'Loft Angle') model.loftAngle = r4(val)
      else if (label === 'Lie Angle') model.lieAngle = r4(val)
      else if (label === 'Hosel OD') {
        model.hoselOD_in = val
        model.hoselOD_mm = val !== null ? r4(val * IN_TO_MM) : null
      }
      else if (label.startsWith('F1/E')) {
        model.f1e_in = r4(val)
        model.f1e_mm = val !== null ? r4(val * IN_TO_MM) : null
      }
    }

    // Skip models with no valid Final Head Weight (e.g. iSW with formulas only)
    if (model.finalHeadWeight === null || model.finalHeadWeight === 0) {
      // Try Eng Targeted Mass as fallback
      const engRow = rows.find(r => String(r?.[0] || '').includes('Eng Targeted'))
      const engVal = toNum(engRow?.[1])
      if (engVal && engVal > 0) model.finalHeadWeight = engVal
      else continue // skip this model
    }

    // ── Components table: find row with col[0]='Components' col[2]='Name' ──
    let compRowIdx = -1
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i]?.[0] || '') === 'Components' && String(rows[i]?.[2] || '') === 'Name') {
        compRowIdx = i
        break
      }
    }

    if (compRowIdx >= 0) {
      for (let i = compRowIdx + 1; i < rows.length; i++) {
        const row = rows[i]
        const name = String(row?.[2] || '').trim()
        // Stop when name is empty and no component number
        if (!name) break
        const mass = toNum(row?.[3])
        const density_imp = toNum(row?.[5])
        const density_met = density_imp !== null
          ? r4(density_imp * G_IN3_TO_G_CM3)
          : null

        model.parts[name] = { mass, density_imp, density_met }
      }
    }

    models.push(model)
  }

  // ── Collect all unique part names (preserve insertion order) ──
  const partSet = new Set<string>()
  for (const m of models) {
    Object.keys(m.parts).forEach(p => partSet.add(p))
  }

  return {
    fileName: '',
    version,
    models,
    allPartNames: Array.from(partSet),
  }
}
