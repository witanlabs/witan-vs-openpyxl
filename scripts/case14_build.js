// fixtures/merge_borders.xlsx: merge A1:B2, where each cell in the merge has
// a distinctly different border stored in the XML. D1 is an unmerged reference.
await xlsx.addSheet(wb, "Data")

// Distinct borders, one per cell
const thinRed    = { style: "thin",   color: "#FF0000" }
const thickBlue  = { style: "thick",  color: "#0000FF" }
const medGreen   = { style: "medium", color: "#00FF00" }
const dashedMag  = { style: "dashed", color: "#FF00FF" }

await xlsx.setCells(wb, [
  {address:"Data!A1", value:"A1"},
  {address:"Data!B1", value:"B1"},
  {address:"Data!A2", value:"A2"},
  {address:"Data!B2", value:"B2"},
  {address:"Data!D1", value:"D1"},
])

await xlsx.setStyle(wb, "Data!A1", { border: {top: thinRed,   bottom: thinRed,   left: thinRed,   right: thinRed  } })
await xlsx.setStyle(wb, "Data!B1", { border: {top: thickBlue, bottom: thickBlue, left: thickBlue, right: thickBlue} })
await xlsx.setStyle(wb, "Data!A2", { border: {top: medGreen,  bottom: medGreen,  left: medGreen,  right: medGreen } })
await xlsx.setStyle(wb, "Data!B2", { border: {top: dashedMag, bottom: dashedMag, left: dashedMag, right: dashedMag} })
await xlsx.setStyle(wb, "Data!D1", { border: {top: thickBlue, bottom: thickBlue} })

await xlsx.setSheetProperties(wb, "Data", {merges: ["A1:B2"]})

// Read back through witan — should reflect per-cell XML
const out = {}
for (const a of ["A1","B1","A2","B2","D1"]) {
  out[a] = (await xlsx.getStyle(wb, `Data!${a}`)).border
}
return out
