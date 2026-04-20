const before = await xlsx.readCell(wb, "Model!B7")
const r = await xlsx.setCells(wb, [{address:"Inputs!B4", value:0.2}])
return {
  before_bonus_rate: 0.10,
  before_net_income: before.value,
  after_bonus_rate: 0.20,
  after_net_income: r.touched["Model!B7"],
  profit: r.touched["Model!B3"],
  bonus: r.touched["Model!B4"],
  tax: r.touched["Model!B6"],
  errors: r.errors,
}
