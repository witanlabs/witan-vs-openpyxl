// witan idiom for a discontiguous CF: emit one rule per contiguous range
// with identical style/formula/operator.
await xlsx.addSheet(wb, "Data")
const vals_A = [98, 115, 50, 5, 200, 14, 3, 175, 94, 193]
const vals_D = [194, 79, 18, 150, 77, 66, 124, 101, 8, 24]
const cells = []
for (let r = 1; r <= 10; r++) {
  cells.push({address: `Data!A${r}`, value: vals_A[r-1]})
  cells.push({address: `Data!D${r}`, value: vals_D[r-1]})
}
await xlsx.setCells(wb, cells)

const sharedRule = {
  type: "cellValue",
  operator: "greaterThan",
  formula: "100",
  style: {fill: {color: "#FFFF00"}},
}
await xlsx.setConditionalFormatting(wb, "Data", [
  {...sharedRule, address: "A1:A10"},
  {...sharedRule, address: "D1:D10"},
], {clear: true})

const rules = await xlsx.getConditionalFormatting(wb, "Data")
return { ruleCount: rules.length, rules }
