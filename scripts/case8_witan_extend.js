// Extend the data table: add price=90 so the table covers 40..90.
// Current data table spans D1:I6 (rows for prices 40,50,60,70,80).
// Need to delete + re-add with the new columnInputValues list.
const before = await xlsx.getDataTable(wb, "Model!D1:I6")

await xlsx.deleteDataTable(wb, "Model!D1:I6")
await xlsx.addDataTable(wb, "Model", {
  type: "twoVariable",
  ref: "D1:I7",
  rowInputCell: before.rowInputCell.split("!")[1],
  columnInputCell: before.columnInputCell.split("!")[1],
  rowInputValues: before.rowInputValues,
  columnInputValues: [...before.columnInputValues, 90],
  formula: before.formula,
})

const after = await xlsx.getDataTable(wb, "Model!D1:I7")
const grid = await xlsx.readRangeTsv(wb, "Model!D1:I7")
return { before, after, grid }
