// Build report_spillref.xlsx:
// - Raw!A2:B13 transactions
// - Summary!D2 = UNIQUE(FILTER(...)) dynamic array that spills D2:D5
// - Summary!F1 = "Count of categories"
// - Summary!F2 = =COUNTA(Summary!D2#)                     consumes the spill
// - Summary!G1 = "Food matches"
// - Summary!G2 = =COUNTIF(Summary!D2#, "Food")            consumes the spill
// - Summary!H1 = "All categories joined"
// - Summary!H2 = =TEXTJOIN(", ", TRUE, Summary!D2#)       consumes the spill
await xlsx.addSheet(wb, "Raw")
await xlsx.addSheet(wb, "Summary")

const cats = ["Food","Travel","Rent","Food","Supplies","Travel","Rent","Food","Supplies","Food","Travel","Supplies"]
const amts = [120, -50, 1500, 80, 45, 200, 1500, 60, 0, 95, -20, 30]
const cells = [
  {address:"Raw!A1", value:"Category"}, {address:"Raw!B1", value:"Amount"},
]
for (let i = 0; i < cats.length; i++) {
  cells.push({address:`Raw!A${i+2}`, value: cats[i]})
  cells.push({address:`Raw!B${i+2}`, value: amts[i]})
}

cells.push({address:"Summary!D1", value:"Unique pos cats"})
cells.push({address:"Summary!D2", formula:"=UNIQUE(FILTER(Raw!A2:A13, Raw!B2:B13>0))"})
cells.push({address:"Summary!F1", value:"Count"})
cells.push({address:"Summary!F2", formula:"=COUNTA(Summary!D2#)"})
cells.push({address:"Summary!G1", value:"Food matches"})
cells.push({address:"Summary!G2", formula:"=COUNTIF(Summary!D2#, \"Food\")"})
cells.push({address:"Summary!H1", value:"Joined"})
cells.push({address:"Summary!H2", formula:"=TEXTJOIN(\", \", TRUE, Summary!D2#)"})

await xlsx.setCells(wb, cells)
return await xlsx.readRangeTsv(wb, "Summary!D1:H5", {includeEmpty:true, includeFormulas:true})
