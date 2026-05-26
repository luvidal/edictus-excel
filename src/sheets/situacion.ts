import type { Workbook, Worksheet } from 'exceljs'
import type { ExcelApplicant, ExcelColumn, SituacionTipo } from '../types'
import { numFmtFor } from '../format'

const TABLE_NAMES: Record<SituacionTipo, string> = {
  deudas: 'TableDeudas',
  propiedades: 'TablePropiedades',
  vehiculos: 'TableVehiculos',
  inversiones: 'TableInversiones',
}

const SHEET_NAMES: Record<SituacionTipo, string> = {
  deudas: 'Deudas',
  propiedades: 'Propiedades',
  vehiculos: 'Vehículos',
  inversiones: 'Inversiones',
}

export function buildSituacionSheet(
  workbook: Workbook,
  applicants: ExcelApplicant[],
  tipo: SituacionTipo,
): Worksheet {
  const sheet = workbook.addWorksheet(SHEET_NAMES[tipo])
  sheet.views = [{ state: 'frozen', ySplit: 1 }]

  // Use the first applicant that has columns defined as the canonical column list.
  const canonicalColumns: ExcelColumn[] = (() => {
    for (const applicant of applicants) {
      const cols = applicant.situacion[tipo].columns
      if (cols && cols.length > 0) return cols
    }
    return []
  })()

  const allColumns: ExcelColumn[] = [
    { key: '__persona', label: 'Persona', format: 'text' },
    { key: '__rut', label: 'RUT', format: 'text' },
    ...canonicalColumns,
  ]

  sheet.columns = allColumns.map((col) => ({
    width: col.key === '__persona' ? 30 : col.key === '__rut' ? 14 : 18,
  }))

  // Flatten rows across all applicants.
  const flatRows: Array<{ persona: string; rut: string; row: Record<string, string | number | null> }> = []
  applicants.forEach((applicant) => {
    applicant.situacion[tipo].rows.forEach((row) => {
      flatRows.push({ persona: applicant.label, rut: applicant.rut, row })
    })
  })

  if (flatRows.length === 0) {
    // Header-only sheet (Excel Tables require ≥1 data row).
    const headerRow = sheet.getRow(1)
    allColumns.forEach((col, idx) => {
      headerRow.getCell(idx + 1).value = col.label
    })
    headerRow.font = { bold: true }
    return sheet
  }

  const tableRows = flatRows.map(({ persona, rut, row }) =>
    allColumns.map((col) => {
      if (col.key === '__persona') return persona
      if (col.key === '__rut') return rut
      const v = row[col.key]
      return v === null || v === undefined ? '' : v
    }),
  )

  sheet.addTable({
    name: TABLE_NAMES[tipo],
    ref: 'A1',
    headerRow: true,
    columns: allColumns.map((col) => ({ name: col.label, filterButton: true })),
    rows: tableRows,
  })

  // Apply numFmt to data cells based on each column's format.
  allColumns.forEach((col, colIdx) => {
    const numFmt = numFmtFor(col.format)
    if (!numFmt) return
    for (let r = 0; r < tableRows.length; r += 1) {
      const cell = sheet.getRow(r + 2).getCell(colIdx + 1)
      cell.numFmt = numFmt
    }
  })

  return sheet
}
