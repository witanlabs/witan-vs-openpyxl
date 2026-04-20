// Build report.xlsx with raw transactions + summary sheet.
await xlsx.addSheet(wb, "Raw")
await xlsx.addSheet(wb, "Summary")

const cats = ["Food","Travel","Rent","Food","Supplies","Travel","Rent","Food","Supplies","Food","Travel","Supplies"]
const amts = [120, -50, 1500, 80, 45, 200, 1500, 60, 0, 95, -20, 30]
const cells = [
  {address:"Raw!A1", value:"Category"}, {address:"Raw!B1", value:"Amount"},
]
for (let i = 0; i < cats.length; i++) {
  cells.push({address: `Raw!A${i+2}`, value: cats[i]})
  cells.push({address: `Raw!B${i+2}`, value: amts[i]})
}
await xlsx.setCells(wb, cells)
return await xlsx.readRangeTsv(wb, "Raw!A1:B13")
