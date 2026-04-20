// Build sensitivity2d.xlsx with a 2-variable What-If Data Table
await xlsx.addSheet(wb, "Model")
await xlsx.setCells(wb, [
  {address:"Model!A1", value:"Price"},       {address:"Model!B1", value:50},
  {address:"Model!A2", value:"Volume"},      {address:"Model!B2", value:1000},
  {address:"Model!A3", value:"Unit cost"},   {address:"Model!B3", value:30},
  {address:"Model!A4", value:"Fixed cost"},  {address:"Model!B4", value:5000},
  {address:"Model!A5", value:"Profit"},      {address:"Model!B5", formula:"=B1*B2 - B3*B2 - B4"},
])

await xlsx.addDataTable(wb, "Model", {
  type: "twoVariable",
  ref: "D1:I6",
  rowInputCell: "B2",
  columnInputCell: "B1",
  rowInputValues: [500, 750, 1000, 1250, 1500],
  columnInputValues: [40, 50, 60, 70, 80],
  formula: "=B5",
})

const dt = await xlsx.getDataTable(wb, "Model!D1:I6")
const range = await xlsx.readRangeTsv(wb, "Model!D1:I6", {includeEmpty:true})
return { dataTable: dt, grid: range }
