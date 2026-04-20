// fixtures/shift.xlsx: Data!A1:B10 of revenue/cost, plus formulas that should shift when a row is inserted at 5.
await xlsx.addSheet(wb, "Data")

const cells = [
  {address:"Data!A1", value:"Revenue"}, {address:"Data!B1", value:"Cost"},
]
for (let r = 2; r <= 10; r++) {
  cells.push({address:`Data!A${r}`, value: r * 100})   // 200..1000
  cells.push({address:`Data!B${r}`, value: r *  50})   // 100..500
}

// Formulas that reference the data
cells.push({address:"Data!C1", value:"Profit"})
for (let r = 2; r <= 10; r++) {
  cells.push({address:`Data!C${r}`, formula:`=A${r}-B${r}`})          // row-local
}
cells.push({address:"Data!E1", value:"Totals"})
cells.push({address:"Data!E2", formula:"=SUM(A2:A10)"})               // range crosses insertion point
cells.push({address:"Data!E3", formula:"=SUM(B2:B10)"})
cells.push({address:"Data!E4", formula:"=SUM(C2:C10)"})
cells.push({address:"Data!E5", formula:"=AVERAGE(A2:A10)"})

// Dynamic-array formula (SORT, spills down) — also crosses insertion point
cells.push({address:"Data!G1", value:"Sorted A"})
cells.push({address:"Data!G2", formula:"=SORT(A2:A10)"})

// Named range that crosses insertion point
await xlsx.setCells(wb, cells)
await xlsx.addDefinedName(wb, "RevenueRange", "Data!$A$2:$A$10")
await xlsx.setCells(wb, [{address:"Data!E6", formula:"=SUM(RevenueRange)"}])

return await xlsx.readRangeTsv(wb, "Data!A1:G12", {includeEmpty:true, includeFormulas:true})
