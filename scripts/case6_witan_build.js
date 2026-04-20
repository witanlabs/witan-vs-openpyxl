// Build same fixture + single-series line chart via witan
await xlsx.addSheet(wb, "Data")
await xlsx.setCells(wb, [
  {address:"Data!A1", value:"Services"},
  {address:"Data!B1", value:"2025-01"}, {address:"Data!C1", value:"2025-02"}, {address:"Data!D1", value:"2025-03"}, {address:"Data!E1", value:"2025-04"}, {address:"Data!F1", value:"2025-05"},
  {address:"Data!A2", value:"Compute"},
  {address:"Data!B2", value:100},      {address:"Data!C2", value:200},      {address:"Data!D2", value:150},      {address:"Data!E2", value:180},      {address:"Data!F2", value:220},
])

await xlsx.addChart(wb, "Data", {
  name: "Compute",
  position: {from: {cell: "A5"}, to: {cell: "H22"}},
  title: {text: "Compute usage (witan single-series)"},
  legend: {position: "right"},
  groups: [{
    type: "line",
    series: [{
      name: {ref: "Data!A2"},
      categories: "Data!B1:F1",
      values: "Data!B2:F2",
    }],
  }],
})

const charts = await xlsx.listCharts(wb)
return { charts }
