// Insert 1 row after row 4 (→ current row 5 becomes row 6), fill new row with Revenue=525/Cost=250.
await xlsx.insertRowAfter(wb, "Data", 4, 1)
await xlsx.setCells(wb, [
  {address:"Data!A5", value:525},
  {address:"Data!B5", value:250},
  {address:"Data!C5", formula:"=A5-B5"},
])

const out = {}
for (const addr of ["A5","B5","C2","C5","C6","C11","E2","E3","E4","E5","E6","G2","G3","G10","G11"]) {
  const c = await xlsx.readCell(wb, `Data!${addr}`)
  out[addr] = { value: c.value, formula: c.formula ?? null }
}
const names = await xlsx.listDefinedNames(wb)
return { cells: out, definedNames: names }
