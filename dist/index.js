'use strict';

var ExcelJS = require('exceljs');

function _interopDefault (e) { return e && e.__esModule ? e : { default: e }; }

var ExcelJS__default = /*#__PURE__*/_interopDefault(ExcelJS);

// src/buildexcel.ts

// src/format.ts
function numFmtFor(format) {
  switch (format) {
    case "currency":
      return "#,##0";
    case "integer":
      return "#,##0";
    case "percent":
      return "0.00%";
    case "text":
    case void 0:
      return null;
  }
}

// src/sheets/resumen.ts
function buildResumenSheet(workbook, resumen) {
  const sheet = workbook.addWorksheet("Resumen");
  sheet.views = [{ state: "frozen", ySplit: 1 }];
  const blocks = [
    {
      title: "Antecedentes Financieros",
      headers: ["Concepto", "Titular", "Codeudor", "Conjunto"],
      rows: resumen.financierosRows,
      expectedValueCount: 3
    },
    {
      title: "Estado Situaci\xF3n",
      headers: ["Concepto", "Titular", "Codeudor", "Total"],
      rows: resumen.situacionRows,
      expectedValueCount: 3
    },
    {
      title: "Indicadores",
      headers: ["Concepto", "Individual", "Conjunto"],
      rows: resumen.indicadoresRows,
      expectedValueCount: 2
    }
  ];
  sheet.columns = [
    { width: 40 },
    { width: 15 },
    { width: 15 },
    { width: 15 }
  ];
  let cursor = 1;
  blocks.forEach((block, idx) => {
    if (idx > 0) {
      cursor += 1;
    }
    const headerRow = sheet.getRow(cursor);
    block.headers.forEach((label, colIdx) => {
      headerRow.getCell(colIdx + 1).value = label;
    });
    headerRow.font = { bold: true };
    cursor += 1;
    block.rows.forEach((row) => {
      const dataRow = sheet.getRow(cursor);
      dataRow.getCell(1).value = row.label;
      const numFmt = numFmtFor(row.format);
      for (let i = 0; i < block.expectedValueCount; i += 1) {
        const cell = dataRow.getCell(i + 2);
        const value = row.values[i] ?? null;
        cell.value = value;
        if (numFmt) {
          cell.numFmt = numFmt;
        }
      }
      cursor += 1;
    });
  });
  return sheet;
}

// src/sheets/perfil.ts
function perfilKey(entry) {
  return `${entry.section}\0${entry.subsection ?? ""}\0${entry.label}`;
}
function perfilHeader(entry) {
  const prefix = entry.subsection ? `${entry.section} \u2014 ${entry.subsection} \u2014 ` : `${entry.section} \u2014 `;
  return `${prefix}${entry.label}`;
}
function collectPerfilColumns(applicants) {
  const seen = /* @__PURE__ */ new Map();
  applicants.forEach((applicant) => {
    applicant.perfil.forEach((entry) => {
      const key = perfilKey(entry);
      if (!seen.has(key)) {
        seen.set(key, {
          key,
          header: perfilHeader(entry),
          section: entry.section,
          subsection: entry.subsection,
          label: entry.label
        });
      }
    });
  });
  return Array.from(seen.values());
}
function buildPerfilSheet(workbook, applicants) {
  const sheet = workbook.addWorksheet("Perfil");
  sheet.views = [{ state: "frozen", ySplit: 1 }];
  const perfilColumns = collectPerfilColumns(applicants);
  const baseHeaders = ["Persona", "RUT", "Edad", "Role"];
  const headers = [...baseHeaders, ...perfilColumns.map((c) => c.header)];
  const widths = [30, 14, 8, 12, ...perfilColumns.map(() => 20)];
  sheet.columns = widths.map((width) => ({ width }));
  const headerRow = sheet.getRow(1);
  headers.forEach((label, idx) => {
    headerRow.getCell(idx + 1).value = label;
  });
  headerRow.font = { bold: true };
  applicants.forEach((applicant, applicantIdx) => {
    const row = sheet.getRow(applicantIdx + 2);
    row.getCell(1).value = applicant.label;
    row.getCell(2).value = applicant.rut;
    row.getCell(3).value = applicant.edad;
    row.getCell(4).value = applicant.role;
    const valuesByKey = /* @__PURE__ */ new Map();
    applicant.perfil.forEach((entry) => {
      valuesByKey.set(perfilKey(entry), entry.value);
    });
    perfilColumns.forEach((col, colIdx) => {
      const value = valuesByKey.get(col.key) ?? "";
      row.getCell(baseHeaders.length + colIdx + 1).value = value;
    });
  });
  return sheet;
}

// src/sheets/situacion.ts
var TABLE_NAMES = {
  deudas: "TableDeudas",
  propiedades: "TablePropiedades",
  vehiculos: "TableVehiculos",
  inversiones: "TableInversiones"
};
var SHEET_NAMES = {
  deudas: "Deudas",
  propiedades: "Propiedades",
  vehiculos: "Veh\xEDculos",
  inversiones: "Inversiones"
};
function buildSituacionSheet(workbook, applicants, tipo) {
  const sheet = workbook.addWorksheet(SHEET_NAMES[tipo]);
  sheet.views = [{ state: "frozen", ySplit: 1 }];
  const canonicalColumns = (() => {
    for (const applicant of applicants) {
      const cols = applicant.situacion[tipo].columns;
      if (cols && cols.length > 0) return cols;
    }
    return [];
  })();
  const allColumns = [
    { key: "__persona", label: "Persona", format: "text" },
    { key: "__rut", label: "RUT", format: "text" },
    ...canonicalColumns
  ];
  sheet.columns = allColumns.map((col) => ({
    width: col.key === "__persona" ? 30 : col.key === "__rut" ? 14 : 18
  }));
  const flatRows = [];
  applicants.forEach((applicant) => {
    applicant.situacion[tipo].rows.forEach((row) => {
      flatRows.push({ persona: applicant.label, rut: applicant.rut, row });
    });
  });
  if (flatRows.length === 0) {
    const headerRow = sheet.getRow(1);
    allColumns.forEach((col, idx) => {
      headerRow.getCell(idx + 1).value = col.label;
    });
    headerRow.font = { bold: true };
    return sheet;
  }
  const tableRows = flatRows.map(
    ({ persona, rut, row }) => allColumns.map((col) => {
      if (col.key === "__persona") return persona;
      if (col.key === "__rut") return rut;
      const v = row[col.key];
      return v === null || v === void 0 ? "" : v;
    })
  );
  sheet.addTable({
    name: TABLE_NAMES[tipo],
    ref: "A1",
    headerRow: true,
    columns: allColumns.map((col) => ({ name: col.label, filterButton: true })),
    rows: tableRows
  });
  allColumns.forEach((col, colIdx) => {
    const numFmt = numFmtFor(col.format);
    if (!numFmt) return;
    for (let r = 0; r < tableRows.length; r += 1) {
      const cell = sheet.getRow(r + 2).getCell(colIdx + 1);
      cell.numFmt = numFmt;
    }
  });
  return sheet;
}

// src/buildexcel.ts
var SITUACION_ORDER = ["deudas", "propiedades", "vehiculos", "inversiones"];
async function buildExcel(input) {
  const workbook = new ExcelJS__default.default.Workbook();
  workbook.creator = "Jogi";
  workbook.title = input.meta.requestLabel;
  workbook.created = new Date(input.meta.generatedAt);
  buildResumenSheet(workbook, input.resumen);
  buildPerfilSheet(workbook, input.applicants);
  for (const tipo of SITUACION_ORDER) {
    buildSituacionSheet(workbook, input.applicants, tipo);
  }
  const buffer = await workbook.xlsx.writeBuffer();
  if (buffer instanceof ArrayBuffer) return buffer;
  const view = buffer;
  return view.buffer.slice(view.byteOffset, view.byteOffset + view.byteLength);
}

exports.buildExcel = buildExcel;
//# sourceMappingURL=index.js.map
//# sourceMappingURL=index.js.map