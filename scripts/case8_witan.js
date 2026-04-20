// Describe the data table on the Model sheet.
const sheets = await xlsx.listSheets(wb)
const withDTs = sheets.filter(s => s.dataTables && s.dataTables.length)
const out = []
for (const s of withDTs) {
  for (const ref of s.dataTables) {
    const dt = await xlsx.getDataTable(wb, ref)
    out.push(dt)
  }
}
return out
