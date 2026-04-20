// Seed the same data as the openpyxl side
await xlsx.addSheet(wb, "Data")
const vals_A = [98, 115, 50, 5, 200, 14, 3, 175, 94, 193]
const vals_D = [194, 79, 18, 150, 77, 66, 124, 101, 8, 24]
const cells = []
for (let r = 1; r <= 10; r++) {
  cells.push({address: `Data!A${r}`, value: vals_A[r-1]})
  cells.push({address: `Data!D${r}`, value: vals_D[r-1]})
}
await xlsx.setCells(wb, cells)

// Single rule with a discontiguous address — space-separated union, matching
// Excel's native `sqref="A1:A10 D1:D10"` form.
await xlsx.setConditionalFormatting(wb, "Data", [{
  type: "cellValue",
  address: "A1:A10 D1:D10",
  operator: "greaterThan",
  formula: "100",
  style: {fill: {color: "#FFFF00"}},
}], {clear: true})

const rules = await xlsx.getConditionalFormatting(wb, "Data")
return { ruleCount: rules.length, rules }
