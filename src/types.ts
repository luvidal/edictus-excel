export type ExcelFormat = 'currency' | 'integer' | 'percent' | 'text'

export interface ExcelColumn {
  key: string
  label: string
  format?: ExcelFormat
}

export interface ExcelPerfilEntry {
  section: string
  subsection?: string
  label: string
  value: string
}

export interface ExcelSituacionTable {
  columns: ExcelColumn[]
  rows: Array<Record<string, string | number | null>>
}

export interface ExcelApplicant {
  role: 'titular' | 'codeudor'
  rut: string
  label: string
  edad: number | null
  perfil: ExcelPerfilEntry[]
  situacion: {
    deudas: ExcelSituacionTable
    propiedades: ExcelSituacionTable
    vehiculos: ExcelSituacionTable
    inversiones: ExcelSituacionTable
  }
}

export interface ExcelResumenRow {
  label: string
  values: Array<number | null>
  format: ExcelFormat
}

export interface ExcelInput {
  meta: {
    requestLabel: string
    generatedAt: string
    ufValue: number | null
    ufDate?: string
  }
  cliente: { nombre: string; rut: string } | null
  applicants: ExcelApplicant[]
  resumen: {
    financierosRows: ExcelResumenRow[]
    situacionRows: ExcelResumenRow[]
    indicadoresRows: ExcelResumenRow[]
  }
}

export type SituacionTipo = 'deudas' | 'propiedades' | 'vehiculos' | 'inversiones'
