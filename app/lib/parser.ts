import * as XLSX from 'xlsx'

const IN_TO_MM = 25.4
const G_PER_IN3_TO_G_PER_CM3 = 0.0610237441

export interface Component {
  name: string
  mass_g: number
  volume_in3: number | null
  density_g_in3: number | null
  density_g_cm3: number | null
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
  components: Component[]
}

export interface ParsedSpec {
  fileName: string
  version: string
  models: ModelData[]
  allComponentNames: string[]
}

export function parseSpecFile(buffer: Buffer): ParsedSpec {
  const wb = XLSX.read(buffer, { type: 'buffer' })

  let version = ''
  if (wb.SheetNames.includes('Revision Record')) {
    const rows = sheetToRows(wb.Sheets['Revision Record'])
    for (let i = rows.length - 1; i >= 1; i--) {
      const r = rows[i]
      if (r[0] && r[1]) { version = 'Rev ' + String(r[0]).trim(); break }
    }
  }

  const modelSheets: string[] = []
  for (const sName of wb.SheetNames) {
    const rows = sheetToRows(wb.Sheets[sName])
    if (rows.some(r => String(r[0] || '').trim() === 'Final Head Weight')) {
      modelSheets.push(sName)
    }
  }

  const models: ModelData[] = modelSheets.map(sName =>
    parseModelSheet(sName, sheetToRows(wb.Sheets[sName]))
  )

  const seen = new Set<string>()
  const allComponentNames: string[] = []
  for (const m of models) {
    for (const c of m.components) {
      if (!seen.has(c.name)) { seen.add(c.name); allComponentNames.push(c.name) }
    }
  }

  return { fileName: '', version, models, allComponentNames }
}

function parseModelSheet(modelName: string, rows: (string | number | null)[][]): ModelData {
  const model: ModelData = {
    modelName, finalHeadWeight: null, loftAngle: null, lieAngle: null,
    hoselOD_in: null, hoselOD_mm: null, f1e_in: null, f1e_mm: null, components: [],
  }

  let inComponents = false

  for (const row of rows) {
    const col0 = String(row[0] || '').trim()

    if (col0 === 'Final Head Weight') model.finalHeadWeight = toNum(row[1])
    else if (col0 === 'Loft Angle') model.loftAngle = toNum(row[1])
    else if (col0 === 'Lie Angle') model.lieAngle = toNum(row[1])
    else if (col0 === 'Hosel OD') {
      const v = toNum(row[1])
      model.hoselOD_in = v
      model.hoselOD_mm = v !== null ? round4(v * IN_TO_MM) : null
    } else if (col0.startsWith('F1/E')) {
      const v = toNum(row[1])
      model.f1e_in = v
      model.f1e_mm = v !== null ? round4(v * IN_TO_MM) : null
    }

    if (col0 === 'Components') { inComponents = true; continue }
    if (inComponents && col0 === 'World Frame Results') { inComponents = false; continue }

    if (inComponents) {
      // col[2]=Name, col[3]=Mass(g), col[4]=Volume(in^3), col[5]=Density(g/in3)
      const name = row[2] !== null && row[2] !== undefined ? String(row[2]).trim() : ''
      if (!name) continue
      const mass = toNum(row[3])
      const volume = toNum(row[4])
      const density_in3 = toNum(row[5])
      const density_g_cm3 = density_in3 !== null ? round4(density_in3 * G_PER_IN3_TO_G_PER_CM3) : null
      if (mass !== null) {
        model.components.push({ name, mass_g: mass, volume_in3: volume, density_g_in3: density_in3, density_g_cm3 })
      }
    }
  }

  return model
}

function sheetToRows(ws: XLSX.WorkSheet): (string | number | null)[][] {
  return XLSX.utils.sheet_to_json<(string | number | null)[]>(ws, { header: 1, defval: null }) as (string | number | null)[][]
}

function toNum(v: string | number | null | undefined): number | null {
  if (v === null || v === undefined || v === '') return null
  const n = typeof v === 'number' ? v : parseFloat(String(v))
  return isNaN(n) ? null : n
}

function round4(n: number): number {
  return Math.round(n * 10000) / 10000
}
