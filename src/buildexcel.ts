import ExcelJS from 'exceljs'
import type { ExcelInput, SituacionTipo } from './types'
import { buildResumenSheet } from './sheets/resumen'
import { buildPerfilSheet } from './sheets/perfil'
import { buildSituacionSheet } from './sheets/situacion'

const SITUACION_ORDER: SituacionTipo[] = ['deudas', 'propiedades', 'vehiculos', 'inversiones']

export async function buildExcel(input: ExcelInput): Promise<ArrayBuffer> {
  const workbook = new ExcelJS.Workbook()
  workbook.creator = 'Jogi'
  workbook.title = input.meta.requestLabel
  workbook.created = new Date(input.meta.generatedAt)

  buildResumenSheet(workbook, input.resumen)
  buildPerfilSheet(workbook, input.applicants)
  for (const tipo of SITUACION_ORDER) {
    buildSituacionSheet(workbook, input.applicants, tipo)
  }

  const buffer = await workbook.xlsx.writeBuffer()
  // ExcelJS returns a Node Buffer or a Uint8Array depending on env. Normalize to ArrayBuffer.
  if (buffer instanceof ArrayBuffer) return buffer
  const view = buffer as unknown as Uint8Array
  return view.buffer.slice(view.byteOffset, view.byteOffset + view.byteLength) as ArrayBuffer
}
