await xlsx.setCells(wb, [
  {address:"Summary!D2", formula:"=UNIQUE(FILTER(Raw!A2:A13, Raw!B2:B13>0))"},
])
// Read the spill region
const tsv = await xlsx.readRangeTsv(wb, "Summary!D2:D10", {includeEmpty:false})
const cells = await xlsx.readRange(wb, {sheet:"Summary", from:{row:2,col:4}, to:{row:10,col:4}})
return {
  tsv,
  values: cells.map(r => r[0].value),
}
