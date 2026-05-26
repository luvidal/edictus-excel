import { describe, it, expect } from 'vitest'
import ExcelJS from 'exceljs'
import { buildExcel } from './buildexcel'
import type { ExcelApplicant, ExcelColumn, ExcelInput } from './types'

const deudasColumns: ExcelColumn[] = [
  { key: 'banco', label: 'Banco', format: 'text' },
  { key: 'tipo', label: 'Tipo', format: 'text' },
  { key: 'saldo', label: 'Saldo', format: 'currency' },
]

const propiedadesColumns: ExcelColumn[] = [
  { key: 'direccion', label: 'Dirección', format: 'text' },
  { key: 'avaluo', label: 'Avalúo', format: 'currency' },
]

const vehiculosColumns: ExcelColumn[] = [
  { key: 'marca', label: 'Marca', format: 'text' },
  { key: 'anio', label: 'Año', format: 'integer' },
]

const inversionesColumns: ExcelColumn[] = [
  { key: 'tipo', label: 'Tipo', format: 'text' },
  { key: 'monto', label: 'Monto', format: 'currency' },
]

function makeApplicant(role: 'titular' | 'codeudor', name: string, rut: string): ExcelApplicant {
  return {
    role,
    rut,
    label: name,
    edad: 40,
    perfil: [
      { section: 'Identificación', label: 'Estado civil', value: 'Casado' },
      { section: 'Laboral', subsection: 'Empleado', label: 'Empresa', value: 'Acme SA' },
    ],
    situacion: {
      deudas: {
        columns: deudasColumns,
        rows: [
          { banco: 'Banco X', tipo: 'Consumo', saldo: 1_500_000 },
          { banco: 'Banco Y', tipo: 'Hipotecario', saldo: 50_000_000 },
        ],
      },
      propiedades: {
        columns: propiedadesColumns,
        rows: [{ direccion: 'Calle 1 # 100', avaluo: 80_000_000 }],
      },
      vehiculos: {
        columns: vehiculosColumns,
        rows: [{ marca: 'Toyota', anio: 2020 }],
      },
      inversiones: {
        columns: inversionesColumns,
        rows: [{ tipo: 'Fondo mutuo', monto: 5_000_000 }],
      },
    },
  }
}

function makeEmptyApplicant(role: 'titular' | 'codeudor', name: string, rut: string): ExcelApplicant {
  return {
    role,
    rut,
    label: name,
    edad: 35,
    perfil: [{ section: 'Identificación', label: 'Estado civil', value: 'Soltero' }],
    situacion: {
      deudas: { columns: deudasColumns, rows: [] },
      propiedades: { columns: propiedadesColumns, rows: [] },
      vehiculos: { columns: vehiculosColumns, rows: [] },
      inversiones: { columns: inversionesColumns, rows: [] },
    },
  }
}

function makeInput(applicants: ExcelApplicant[]): ExcelInput {
  return {
    meta: {
      requestLabel: 'Crédito Hipotecario — Juan Pérez',
      generatedAt: '2026-05-26T12:00:00.000Z',
      ufValue: 38000,
      ufDate: '2026-05-26',
    },
    cliente: { nombre: 'Juan Pérez', rut: '11.111.111-1' },
    applicants,
    resumen: {
      financierosRows: [
        { label: 'Renta líquida', values: [3_000_000, 2_000_000, 5_000_000], format: 'currency' },
        { label: 'Dividendo', values: [800_000, 500_000, 1_300_000], format: 'currency' },
      ],
      situacionRows: [
        { label: 'Total activos', values: [85_000_000, 60_000_000, 145_000_000], format: 'currency' },
      ],
      indicadoresRows: [
        { label: 'Carga financiera', values: [0.27, 0.26], format: 'percent' },
      ],
    },
  }
}

async function loadWorkbook(input: ExcelInput): Promise<ExcelJS.Workbook> {
  const buffer = await buildExcel(input)
  const wb = new ExcelJS.Workbook()
  await wb.xlsx.load(buffer)
  return wb
}

describe('buildExcel', () => {
  it('produces 6 sheets in the correct order with the correct names', async () => {
    const input = makeInput([
      makeApplicant('titular', 'Juan Pérez', '11.111.111-1'),
      makeApplicant('codeudor', 'María López', '22.222.222-2'),
    ])
    const wb = await loadWorkbook(input)
    const names = wb.worksheets.map((s) => s.name)
    expect(names).toEqual(['Resumen', 'Perfil', 'Deudas', 'Propiedades', 'Vehículos', 'Inversiones'])
  })

  it('Resumen contains all three sub-tables with headers and data', async () => {
    const input = makeInput([
      makeApplicant('titular', 'Juan Pérez', '11.111.111-1'),
      makeApplicant('codeudor', 'María López', '22.222.222-2'),
    ])
    const wb = await loadWorkbook(input)
    const sheet = wb.getWorksheet('Resumen')!

    // Block 1: Antecedentes Financieros (header at row 1, 2 data rows → rows 2-3)
    expect(sheet.getRow(1).getCell(1).value).toBe('Concepto')
    expect(sheet.getRow(1).getCell(2).value).toBe('Titular')
    expect(sheet.getRow(1).getCell(3).value).toBe('Codeudor')
    expect(sheet.getRow(1).getCell(4).value).toBe('Conjunto')
    expect(sheet.getRow(2).getCell(1).value).toBe('Renta líquida')
    expect(sheet.getRow(2).getCell(2).value).toBe(3_000_000)
    expect(sheet.getRow(3).getCell(1).value).toBe('Dividendo')

    // One blank separator row → row 4. Block 2 header → row 5, data → row 6.
    expect(sheet.getRow(5).getCell(1).value).toBe('Concepto')
    expect(sheet.getRow(5).getCell(4).value).toBe('Total')
    expect(sheet.getRow(6).getCell(1).value).toBe('Total activos')
    expect(sheet.getRow(6).getCell(2).value).toBe(85_000_000)

    // Blank → row 7. Block 3 header → row 8 (3 columns), data → row 9.
    expect(sheet.getRow(8).getCell(1).value).toBe('Concepto')
    expect(sheet.getRow(8).getCell(2).value).toBe('Individual')
    expect(sheet.getRow(8).getCell(3).value).toBe('Conjunto')
    expect(sheet.getRow(9).getCell(1).value).toBe('Carga financiera')
    expect(sheet.getRow(9).getCell(2).value).toBe(0.27)
  })

  it('Perfil has one row per applicant with expected base columns', async () => {
    const input = makeInput([
      makeApplicant('titular', 'Juan Pérez', '11.111.111-1'),
      makeApplicant('codeudor', 'María López', '22.222.222-2'),
    ])
    const wb = await loadWorkbook(input)
    const sheet = wb.getWorksheet('Perfil')!

    expect(sheet.getRow(1).getCell(1).value).toBe('Persona')
    expect(sheet.getRow(1).getCell(2).value).toBe('RUT')
    expect(sheet.getRow(1).getCell(3).value).toBe('Edad')
    expect(sheet.getRow(1).getCell(4).value).toBe('Role')
    // First perfil column header concatenates section + label.
    expect(sheet.getRow(1).getCell(5).value).toBe('Identificación — Estado civil')
    // Subsection variant.
    expect(sheet.getRow(1).getCell(6).value).toBe('Laboral — Empleado — Empresa')

    expect(sheet.getRow(2).getCell(1).value).toBe('Juan Pérez')
    expect(sheet.getRow(2).getCell(2).value).toBe('11.111.111-1')
    expect(sheet.getRow(2).getCell(3).value).toBe(40)
    expect(sheet.getRow(2).getCell(4).value).toBe('titular')
    expect(sheet.getRow(2).getCell(5).value).toBe('Casado')

    expect(sheet.getRow(3).getCell(1).value).toBe('María López')
    expect(sheet.getRow(3).getCell(4).value).toBe('codeudor')
  })

  it('situacion sheets carry the Persona + RUT prefix columns', async () => {
    const input = makeInput([
      makeApplicant('titular', 'Juan Pérez', '11.111.111-1'),
      makeApplicant('codeudor', 'María López', '22.222.222-2'),
    ])
    const wb = await loadWorkbook(input)

    for (const name of ['Deudas', 'Propiedades', 'Vehículos', 'Inversiones']) {
      const sheet = wb.getWorksheet(name)!
      expect(sheet.getRow(1).getCell(1).value).toBe('Persona')
      expect(sheet.getRow(1).getCell(2).value).toBe('RUT')
    }

    const deudas = wb.getWorksheet('Deudas')!
    expect(deudas.getRow(1).getCell(3).value).toBe('Banco')
    // First applicant: 2 deudas rows → rows 2-3. Second applicant: 2 deudas rows → rows 4-5.
    expect(deudas.getRow(2).getCell(1).value).toBe('Juan Pérez')
    expect(deudas.getRow(2).getCell(2).value).toBe('11.111.111-1')
    expect(deudas.getRow(2).getCell(3).value).toBe('Banco X')
    expect(deudas.getRow(2).getCell(5).value).toBe(1_500_000)
    expect(deudas.getRow(4).getCell(1).value).toBe('María López')
  })

  it('empty situacion still produces the four sheets header-only', async () => {
    const input = makeInput([makeEmptyApplicant('titular', 'Juan Pérez', '11.111.111-1')])
    const wb = await loadWorkbook(input)

    for (const name of ['Deudas', 'Propiedades', 'Vehículos', 'Inversiones']) {
      const sheet = wb.getWorksheet(name)!
      expect(sheet.getRow(1).getCell(1).value).toBe('Persona')
      expect(sheet.getRow(1).getCell(2).value).toBe('RUT')
      // No data rows.
      expect(sheet.getRow(2).getCell(1).value).toBeFalsy()
    }
  })

  it('applies numFmt per cell on Resumen and situacion currency/percent columns', async () => {
    const input = makeInput([makeApplicant('titular', 'Juan Pérez', '11.111.111-1')])
    const wb = await loadWorkbook(input)

    const resumen = wb.getWorksheet('Resumen')!
    // Row 2 = Renta líquida (currency) — value cell numFmt.
    expect(resumen.getRow(2).getCell(2).numFmt).toBe('#,##0')
    // Indicadores row — percent.
    expect(resumen.getRow(9).getCell(2).numFmt).toBe('0.00%')

    const deudas = wb.getWorksheet('Deudas')!
    // Column 5 = Saldo (currency) on row 2.
    expect(deudas.getRow(2).getCell(5).numFmt).toBe('#,##0')
  })

  it('exposes workbook metadata from input.meta', async () => {
    const input = makeInput([makeApplicant('titular', 'Juan Pérez', '11.111.111-1')])
    const wb = await loadWorkbook(input)
    expect(wb.creator).toBe('Jogi')
    expect(wb.title).toBe('Crédito Hipotecario — Juan Pérez')
  })
})
