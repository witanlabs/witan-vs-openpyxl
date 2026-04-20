// Build review.xlsx with 2 threaded comments on Data!B2 and Data!C2 — one resolved, one open.
await xlsx.addSheet(wb, "Data")
await xlsx.setCells(wb, [
  {address:"Data!A1", value:"Account"}, {address:"Data!B1", value:"Balance"}, {address:"Data!C1", value:"Status"},
  {address:"Data!A2", value:"Cash"},    {address:"Data!B2", value:12345},     {address:"Data!C2", value:"Open"},
  {address:"Data!A3", value:"AR"},      {address:"Data!B3", value:67890},     {address:"Data!C3", value:"Open"},
])
// Add two threaded comments (one resolved)
await xlsx.setCells(wb, [
  {address:"Data!B2", thread: {add: [{author:"Reviewer A", text:"Initial balance confirmed"}], resolved: true}},
  {address:"Data!C2", thread: {add: [{author:"Reviewer B", text:"Needs follow-up on Q3 close"}]}},
])
const b2 = await xlsx.readCell(wb, "Data!B2")
const c2 = await xlsx.readCell(wb, "Data!C2")
return { b2_thread: b2.thread, c2_thread: c2.thread }
