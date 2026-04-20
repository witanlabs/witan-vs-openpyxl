await xlsx.renameSheet(wb, "Inputs", "Parameters")

const out = {}
for (const addr of ["Summary!B1","Summary!B2","Summary!B3","Summary!B4","Summary!B5","Summary!B6"]) {
  const c = await xlsx.readCell(wb, addr)
  out[addr] = { value: c.value, formula: c.formula ?? null }
}
const names = await xlsx.listDefinedNames(wb)
const sheets = (await xlsx.listSheets(wb)).map(s => s.sheet)
return { sheets, cells: out, definedNames: names }
