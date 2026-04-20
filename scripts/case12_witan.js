// Equivalent to the openpyxl CellRichText fixture: three cells with bold TextBlocks
// separated by whitespace.
await xlsx.addSheet(wb, "Data")

// witan's richText is set via setStyle (richText is a property of StyleObj).
// Each entry is {text, style?}. Write plain text via setCells first, then apply
// the richText via setStyle.
await xlsx.setCells(wb, [
  {address:"Data!A1", value:""},
  {address:"Data!A2", value:""},
  {address:"Data!A3", value:""},
])

const bold = {bold: true}
const strikeRed = {strike: true, color:"#990000"}

await xlsx.setStyle(wb, "Data!A1", {
  richText: [
    {text: "Normal text "},
    {text: "Bold1", style: bold},
    {text: " "},                      // whitespace-only run
    {text: "Bold2", style: bold},
    {text: " more normal"},
  ],
})

await xlsx.setStyle(wb, "Data!A2", {
  richText: [
    {text: "Leading "},
    {text: "Bold1", style: bold},
    {text: "   "},                    // three-space run
    {text: "Bold2", style: bold},
    {text: " trailing"},
  ],
})

await xlsx.setStyle(wb, "Data!A3", {
  richText: [
    {text: "Some text"},
    {text: " ", style: strikeRed},    // whitespace-only styled run (the issue's case)
    {text: "and some more."},
  ],
})

return {
  A1: (await xlsx.readCell(wb, "Data!A1")).value,
  A2: (await xlsx.readCell(wb, "Data!A2")).value,
  A3: (await xlsx.readCell(wb, "Data!A3")).value,
}
