const before = await xlsx.readCell(wb, "Summary!E23")
const r = await xlsx.setCells(wb, [{address:"Assumptions!B5", value:0.12}])
return {
  before_rate: 0.08,
  before_npv: before.value,
  after_rate: 0.12,
  after_npv: r.touched["Summary!E23"],
  errors: r.errors,
}
