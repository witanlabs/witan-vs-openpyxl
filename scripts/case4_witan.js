// Read existing threaded comments, then add a resolved thread on B3 by 'Auditor'.
const existing = []
for (const addr of ["Data!B2", "Data!C2", "Data!B3"]) {
  const c = await xlsx.readCell(wb, addr)
  existing.push({address: addr, thread: c.thread ?? null})
}

await xlsx.setCells(wb, [
  {address:"Data!B3", thread: {add:[{author:"Auditor", text:"Verified against ledger"}], resolved: true}},
])

const after = []
for (const addr of ["Data!B2", "Data!C2", "Data!B3"]) {
  const c = await xlsx.readCell(wb, addr)
  after.push({address: addr, thread: c.thread ?? null})
}

return { before: existing, after }
