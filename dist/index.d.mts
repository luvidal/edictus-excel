type ExcelFormat = 'currency' | 'integer' | 'percent' | 'text';
interface ExcelColumn {
    key: string;
    label: string;
    format?: ExcelFormat;
}
interface ExcelPerfilEntry {
    section: string;
    subsection?: string;
    label: string;
    value: string;
}
interface ExcelSituacionTable {
    columns: ExcelColumn[];
    rows: Array<Record<string, string | number | null>>;
}
interface ExcelApplicant {
    role: 'titular' | 'codeudor';
    rut: string;
    label: string;
    edad: number | null;
    perfil: ExcelPerfilEntry[];
    situacion: {
        deudas: ExcelSituacionTable;
        propiedades: ExcelSituacionTable;
        vehiculos: ExcelSituacionTable;
        inversiones: ExcelSituacionTable;
    };
}
interface ExcelResumenRow {
    label: string;
    values: Array<number | null>;
    format: ExcelFormat;
}
interface ExcelInput {
    meta: {
        requestLabel: string;
        generatedAt: string;
        ufValue: number | null;
        ufDate?: string;
    };
    cliente: {
        nombre: string;
        rut: string;
    } | null;
    applicants: ExcelApplicant[];
    resumen: {
        financierosRows: ExcelResumenRow[];
        situacionRows: ExcelResumenRow[];
        indicadoresRows: ExcelResumenRow[];
    };
}
type SituacionTipo = 'deudas' | 'propiedades' | 'vehiculos' | 'inversiones';

declare function buildExcel(input: ExcelInput): Promise<ArrayBuffer>;

export { type ExcelApplicant, type ExcelColumn, type ExcelFormat, type ExcelInput, type ExcelPerfilEntry, type ExcelResumenRow, type ExcelSituacionTable, type SituacionTipo, buildExcel };
