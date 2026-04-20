// Same task: on Data sheet, merge A1:B2 and A2:C3 (overlapping).
await xlsx.addSheet(wb, "Data")

// Seed labels
const cells = []
for (let r = 1; r <= 3; r++) {
  for (let c = 1; c <= 3; c++) {
    cells.push({address:`Data!${String.fromCharCode(64+c)}${r}`, value:`${String.fromCharCode(64+c)}${r}`})
  }
}
await xlsx.setCells(wb, cells)

const attempts = []

// Try applying both merges in a single call
try {
  await xlsx.setSheetProperties(wb, "Data", {merges: ["A1:B2", "A2:C3"]})
  attempts.push({form:"both-at-once", ok:true})
} catch (e) {
  attempts.push({form:"both-at-once", ok:false, error:String(e.message ?? e)})
}

// Try sequential
try {
  await xlsx.setSheetProperties(wb, "Data", {merges: ["A1:B2"]})
  attempts.push({form:"first-only", ok:true})
} catch (e) {
  attempts.push({form:"first-only", ok:false, error:String(e.message ?? e)})
}
try {
  await xlsx.setSheetProperties(wb, "Data", {merges: ["A2:C3"]})
  attempts.push({form:"second-added", ok:true})
} catch (e) {
  attempts.push({form:"second-added", ok:false, error:String(e.message ?? e)})
}

const props = await xlsx.getSheetProperties(wb, "Data")
return { attempts, merges: props.merges }
