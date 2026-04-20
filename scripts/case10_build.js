// fixtures/rename.xlsx: sheet "Inputs" referenced from "Summary" by
//   - single-cell refs
//   - range refs
//   - a dynamic-array formula (UNIQUE spill)
//   - a defined name that points into Inputs
await xlsx.addSheet(wb, "Inputs")
await xlsx.addSheet(wb, "Summary")

const inputs = [
  {address:"Inputs!A1", value:"Name"},    {address:"Inputs!B1", value:"Value"},
  {address:"Inputs!A2", value:"alpha"},   {address:"Inputs!B2", value:100},
  {address:"Inputs!A3", value:"beta"},    {address:"Inputs!B3", value:200},
  {address:"Inputs!A4", value:"alpha"},   {address:"Inputs!B4", value:150},   // duplicate on purpose for UNIQUE
  {address:"Inputs!A5", value:"gamma"},   {address:"Inputs!B5", value:300},
]
await xlsx.setCells(wb, inputs)

await xlsx.addDefinedName(wb, "InputsBeta", "Inputs!$B$3")

const summary = [
  {address:"Summary!A1", value:"Sum of B1+B2"},
  {address:"Summary!B1", formula:"=Inputs!B2 + Inputs!B3"},

  {address:"Summary!A2", value:"SUM(B2:B5)"},
  {address:"Summary!B2", formula:"=SUM(Inputs!B2:B5)"},

  {address:"Summary!A3", value:"Named (InputsBeta)"},
  {address:"Summary!B3", formula:"=InputsBeta"},

  {address:"Summary!A4", value:"UNIQUE spill"},
  {address:"Summary!B4", formula:"=UNIQUE(Inputs!A2:A5)"},
]
await xlsx.setCells(wb, summary)

return await xlsx.readRangeTsv(wb, "Summary!A1:B8", {includeEmpty:true, includeFormulas:true})
