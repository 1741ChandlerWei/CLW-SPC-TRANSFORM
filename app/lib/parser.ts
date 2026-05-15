import * as XLSX from 'xlsx'

const IN_TO_MM = 25.4
const G_IN3_TO_G_CM3 = 0.0610237441

export interface PartData {
  mass: number | null
  density_imp: number | null
  density_met: number | null
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
  const wb = XLSX.read(buffer, {
    type: 'buffer',
    cellFormula: false,
    cellNF: false,
    cellStyles: false,
  })

  // ── Version from Revision Record ──
  let version = ''
  if (wb.SheetNames.includes('Revision Record')) {
    const ws = wb.Sheets['Revision Record']
    const rows = XLSX.utils.sheet_to_json<(string | number | null)[]>(ws, {
      header: 1, defval: null
    }) as (string | number | null)[][]
    for (let i = rows.length - 1; i >= 1; i--) {
      const rev = rows[i]?.[0]
      const desc = rows[i]?.[1]
      if (rev && String(rev).trim() && desc && String(desc).trim()) {
        version = `Rev ${String(rev).trim()}`
        break
      }
    }
  }

  // ── Detect model sheets: starts with 'i', visible, has 'Final Head' label with numeric value ──
  const modelSheetNames = wb.SheetNames.filter(name => {
    if (!name.startsWith('i')) return false
    // 跳過隱藏 Sheet（Hidden=1 or 2）
    const sheetMeta = wb.Workbook?.Sheets?.find((s: {name: string, Hidden?: number}) => s.name === name)
    if (sheetMeta?.Hidden) return false
    const ws = wb.Sheets[name]
    if (!ws) return false
    const rows = XLSX.utils.sheet_to_json<(string | number | null)[]>(ws, {
      header: 1, defval: null
    }) as (string | number | null)[][]

    for (let i = 0; i < Math.min(15, rows.length); i++) {
      const label = String(rows[i]?.[0] || '').trim()
      if (label.includes('Final Head')) {
        const val = rows[i]?.[1]
        if (typeof val === 'number' && val > 0) return true
      }
    }
    return false
  })

  const models: ModelData[] = []

  for (const sname of modelSheetNames) {
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

    let inComp = false

    for (let i = 0; i < rows.length; i++) {
      const row = rows[i]
      const label = String(row?.[0] || '').trim()
      const val = row?.[1]

      // Spec values
      if (label.includes('Final Head')) model.finalHeadWeight = toNum(val)
      else if (label === 'Loft Angle') model.loftAngle = r4(toNum(val))
      else if (label === 'Lie Angle') model.lieAngle = r4(toNum(val))
      else if (label === 'Hosel OD') {
        const n = toNum(val)
        model.hoselOD_in = n
        model.hoselOD_mm = n !== null ? r4(n * IN_TO_MM) : null
      }
      else if (label.startsWith('F1/E')) {
        const n = toNum(val)
        model.f1e_in = r4(n)
        model.f1e_mm = n !== null ? r4(n * IN_TO_MM) : null
      }

      // Components table header
      if (label === 'Components' && String(row?.[2] || '').trim() === 'Name') {
        inComp = true
        continue
      }

      // Components rows
      if (inComp) {
        // Stop at "CAD generated Mass" row
        if (String(row?.[2] || '').includes('CAD generated')) {
          inComp = false
          continue
        }
        const name = String(row?.[2] || '').trim()
        if (!name) continue  // skip empty rows, don't stop
        const mass = toNum(row?.[3])
        const dens_imp = toNum(row?.[5])
        const dens_met = dens_imp !== null ? r4(dens_imp * G_IN3_TO_G_CM3) : null
        model.parts[name] = { mass, density_imp: dens_imp, density_met: dens_met }
      }
    }

    models.push(model)
  }

  // All unique part names in order of first appearance
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
