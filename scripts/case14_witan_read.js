const out = {}
for (const a of ["A1","B1","A2","B2","D1"]) {
  out[a] = (await xlsx.getStyle(wb, `Data!${a}`)).border
}
return out
