import * as XLSX from 'xlsx'

const IN_TO_MM = 25.4
const G_PER_IN3_TO_G_PER_CM3 = 0.0610237

export interface ModelData {
  modelName: string
  // Specs
  finalHeadWeight: number | null       // g
  loftAngle: number | null             // °
  lieAngle: number | null              // °
  hoselOD_in: number | null            // in
  hoselOD_mm: number | null            // mm
  f1e_in: number | null                // in
  f1e_mm: number | null                // mm
  // Parts - dynamic, name -> mass
  parts: Record<string, number>        // g
  // Density - dynamic, name -> {imperial, metric}
  density_imperial: Record<string, number>  // g/in3
  density_metric: Record<string, number>    // g/cm3
}

export interface ParsedSpec {
  fileName: string
  version: string
  models: ModelData[]
  allPartNames: string[]
}

export function parseSpecFile(buffer: Buffer): ParsedSpec {
  const wb = XLSX.read(buffer, { type: 'buffer' })

  // Try to get version from Revision Record sheet
  let version = ''
  if (wb.SheetNames.includes('Revision Record')) {
    const revSheet = wb.Sheets['Revision Record']
    const rows = XLSX.utils.sheet_to_json<string[]>(revSheet, { header: 1 }) as string[][]
    // Find last non-empty rev row
    for (let i = rows.length - 1; i >= 1; i--) {
      const row = rows[i]
      if (row && row[0] && String(row[0]).trim() && String(row[1] || '').trim()) {
        version = `Rev ${String(row[0]).trim()}`
        break
      }
    }
  }

  // Detect model sheets: any sheet that has "Final Head Weight" row
  const modelSheets: string[] = []
  for (const sName of wb.SheetNames) {
    const s = wb.Sheets[sName]
    const txt = XLSX.utils.sheet_to_csv(s)
    if (txt.includes('Final Head Weight')) {
      modelSheets.push(sName)
    }
  }

  // Also try to use the 整理範例 style: a summary sheet with multiple models in columns
  // Check for a sheet named "Dimension Summary" which has all models in rows
  // And individual model sheets for weight/density breakdowns
  const models: ModelData[] = []

  // Parse each model sheet
  for (const sName of modelSheets) {
    const s = wb.Sheets[sName]
    const rows = XLSX.utils.sheet_to_json<(string | number | null)[]>(s, {
      header: 1,
      defval: null,
    }) as (string | number | null)[][]

    const model: ModelData = {
      modelName: sName,
      finalHeadWeight: null,
      loftAngle: null,
      lieAngle: null,
      hoselOD_in: null,
      hoselOD_mm: null,
      f1e_in: null,
      f1e_mm: null,
      parts: {},
      density_imperial: {},
      density_metric: {},
    }

    // Parse sections
    let section: 'spec' | 'parts' | 'density_imp' | 'density_met' | null = null

    for (const row of rows) {
      const cells = row.map(c => (c == null ? '' : String(c).trim()))
      const firstNonEmpty = cells.find(c => c !== '')

      if (!firstNonEmpty) continue

      // Detect section headers
      if (firstNonEmpty === 'Measurement') { section = 'spec'; continue }

      if (section === 'spec') {
        const label = firstNonEmpty
        const val = parseFloat(String(cells[1] || ''))
        if (label === 'Final Head Weight') model.finalHeadWeight = isNaN(val) ? null : val
        else if (label === 'Loft Angle') model.loftAngle = isNaN(val) ? null : val
        else if (label === 'Lie Angle') model.lieAngle = isNaN(val) ? null : val
        else if (label === 'Hosel OD') {
          model.hoselOD_in = isNaN(val) ? null : val
          model.hoselOD_mm = isNaN(val) ? null : val * IN_TO_MM
        } else if (label.startsWith('F1/E')) {
          model.f1e_in = isNaN(val) ? null : val
          model.f1e_mm = isNaN(val) ? null : val * IN_TO_MM
        } else if (label === 'CAD generated Mass') {
          // end of spec section
          section = null
        }
      }
    }

    // Now parse components table - look for "Name" + "Mass" header
    let inComponents = false
    let inDensityImp = false
    let inDensityMet = false

    for (const row of rows) {
      const cells = row.map(c => (c == null ? '' : String(c).trim()))

      // Detect components section
      if (cells.includes('Name') && cells.includes('Mass (g)')) {
        inComponents = true
        inDensityImp = false
        inDensityMet = false
        continue
      }
      // Detect density imperial section
      if (cells.includes('Name') && cells.some(c => c.includes('g/in3'))) {
        // Check if we already parsed imperial
        if (!inDensityImp && Object.keys(model.density_imperial).length === 0) {
          inDensityImp = true
          inComponents = false
          inDensityMet = false
          continue
        } else if (inDensityImp) {
          // Switch to metric
          inDensityMet = true
          inDensityImp = false
          continue
        }
      }

      if (inComponents) {
        const name = cells.find(c => c && c !== '')
        if (!name) { inComponents = false; continue }
        // Find numeric value in row
        const val = findNumericInRow(row)
        if (val !== null && name && !name.startsWith('CAD') && !name.startsWith('Eng')) {
          model.parts[name] = val
        }
        if (name?.startsWith('CAD') || name?.startsWith('Eng')) inComponents = false
      }

      if (inDensityImp) {
        const name = cells.find(c => c && c !== '')
        if (!name) { inDensityImp = false; continue }
        const val = findNumericInRow(row)
        if (val !== null && name) {
          model.density_imperial[name] = val
        }
      }

      if (inDensityMet) {
        const name = cells.find(c => c && c !== '')
        if (!name) { inDensityMet = false; continue }
        const val = findNumericInRow(row)
        if (val !== null && name) {
          model.density_metric[name] = val
        }
      }
    }

    models.push(model)
  }

  // Collect all unique part names across models
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

function findNumericInRow(row: (string | number | null)[]): number | null {
  for (let i = 1; i < row.length; i++) {
    const v = row[i]
    if (v === null || v === '') continue
    const n = parseFloat(String(v))
    if (!isNaN(n)) return n
  }
  return null
}
