# @edictus/excel — Excel Export Satellite

Pure data-in / Buffer-out library that renders an `.xlsx` workbook from a pre-shaped `ExcelInput`. Extracted from [jogi](../jogi) to isolate `exceljs` and all Excel-format quirks from the host's bundle and review surface.

## Compact Instructions

When compacting, preserve: file paths changed, errors found, decisions made, contract changes. Drop: full file contents already read, tool output bodies.

## Communication Style

- **No emotional validation** — never say "I understand your frustration". Results matter, not words.
- **No excessive apologies** — don't apologize repeatedly. Fix the problem.
- **Be direct** — state facts, propose solutions, execute. Skip the fluff.
- **Ask for input** — when stuck or facing multiple approaches, ask rather than guessing.

## Tech Stack

- **TypeScript** (no React, no JSX, no DOM)
- **exceljs** runtime dep — bundled, not external. Consumers get it transitively.
- **tsup** for bundling (ESM + CJS + .d.ts)
- **Vitest** (node env) for unit tests that load the produced buffer back with exceljs

## Project Structure

```
src/
├── index.ts              # Re-exports: buildExcel + ExcelInput + ExcelColumn + sub-types
├── buildexcel.ts         # buildExcel(input): workbook setup + sheet orchestration
├── types.ts              # ExcelInput, ExcelColumn, ExcelApplicant, ExcelResumenRow, etc.
├── format.ts             # ExcelFormat -> numFmt mapping
├── sheets/
│   ├── resumen.ts        # 3 stacked tables (Antecedentes / Estado / Indicadores)
│   ├── perfil.ts         # Wide: one row per applicant; dynamic perfil columns
│   └── situacion.ts      # Long: shared builder for the 4 tipo sheets
└── buildexcel.test.ts    # Vitest: load workbook back, assert headers/cells/formats
```

## Code Rules

1. **File naming** → lowercase, no hyphens/underscores
2. **No `@/` imports** → relative within `src/`
3. **No host code, no `@edictus/*` deps, no React** — pure library
4. **API stability** — `buildExcel` + `ExcelInput` + `ExcelColumn` are the exported surface; breaking changes update the host in the same handoff
5. **Sheet order is contract** — `Resumen, Perfil, Deudas, Propiedades, Vehículos, Inversiones` (accents preserved)
6. **CLAUDE.md maintenance** — update on contract/behavior changes

## Commands

```bash
npm run build        # Build dist/ (ESM + CJS + types)
npm run dev          # Build in watch mode
npm test             # Run vitest suite
npm run test:watch   # Watch mode
```

## Consumer Integration

```json
"@edictus/excel": "github:luvidal/edictus-excel#<sha>"
```

Host calls:
```ts
import { buildExcel } from '@edictus/excel'
const buffer = await buildExcel(toExcelInput(snapshot))
```
