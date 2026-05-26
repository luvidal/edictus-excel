import type { ExcelFormat } from './types'

export function numFmtFor(format: ExcelFormat | undefined): string | null {
  switch (format) {
    case 'currency':
      return '#,##0'
    case 'integer':
      return '#,##0'
    case 'percent':
      return '0.00%'
    case 'text':
    case undefined:
      return null
  }
}
