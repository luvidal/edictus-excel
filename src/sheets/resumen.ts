import type { Workbook, Worksheet } from 'exceljs'
import type { ExcelInput, ExcelResumenRow } from '../types'
import { numFmtFor } from '../format'

interface ResumenBlock {
  title: string
  headers: string[]
  rows: ExcelResumenRow[]
  expectedValueCount: number
}

export function buildResumenSheet(workbook: Workbook, resumen: ExcelInput['resumen']): Worksheet {
  const sheet = workbook.addWorksheet('Resumen')
  sheet.views = [{ state: 'frozen', ySplit: 1 }]

  const blocks: ResumenBlock[] = [
    {
      title: 'Antecedentes Financieros',
      headers: ['Concepto', 'Titular', 'Codeudor', 'Conjunto'],
      rows: resumen.financierosRows,
      expectedValueCount: 3,
    },
    {
      title: 'Estado Situación',
      headers: ['Concepto', 'Titular', 'Codeudor', 'Total'],
      rows: resumen.situacionRows,
      expectedValueCount: 3,
    },
    {
      title: 'Indicadores',
      headers: ['Concepto', 'Individual', 'Conjunto'],
      rows: resumen.indicadoresRows,
      expectedValueCount: 2,
    },
  ]

  // Column widths sized for the widest block (4 columns).
  sheet.columns = [
    { width: 40 },
    { width: 15 },
    { width: 15 },
    { width: 15 },
  ]

  let cursor = 1
  blocks.forEach((block, idx) => {
    if (idx > 0) {
      cursor += 1 // blank separator row
    }

    const headerRow = sheet.getRow(cursor)
    block.headers.forEach((label, colIdx) => {
      headerRow.getCell(colIdx + 1).value = label
    })
    headerRow.font = { bold: true }
    cursor += 1

    block.rows.forEach((row) => {
      const dataRow = sheet.getRow(cursor)
      dataRow.getCell(1).value = row.label
      const numFmt = numFmtFor(row.format)
      for (let i = 0; i < block.expectedValueCount; i += 1) {
        const cell = dataRow.getCell(i + 2)
        const value = row.values[i] ?? null
        cell.value = value
        if (numFmt) {
          cell.numFmt = numFmt
        }
      }
      cursor += 1
    })
  })

  return sheet
}
