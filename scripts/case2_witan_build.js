await xlsx.setWorkbookProperties(wb, {iterativeCalculation: {enabled: true, maxIterations: 200, maxChange: 0.0001}})
await xlsx.addSheet(wb, "Inputs")
await xlsx.addSheet(wb, "Model")
await xlsx.setCells(wb, [
  {address:"Inputs!A1", value:"Revenue"},    {address:"Inputs!B1", value:100000},
  {address:"Inputs!A2", value:"Opex ratio"}, {address:"Inputs!B2", value:0.4},
  {address:"Inputs!A3", value:"Tax rate"},   {address:"Inputs!B3", value:0.3},
  {address:"Inputs!A4", value:"Bonus rate"}, {address:"Inputs!B4", value:0.1},

  {address:"Model!A1", value:"Revenue"},             {address:"Model!B1", formula:"=Inputs!B1"},
  {address:"Model!A2", value:"Opex"},                {address:"Model!B2", formula:"=Model!B1*Inputs!B2"},
  {address:"Model!A3", value:"Profit before bonus"}, {address:"Model!B3", formula:"=Model!B1-Model!B2-Model!B4"},
  {address:"Model!A4", value:"Bonus"},               {address:"Model!B4", formula:"=Model!B3*Inputs!B4"},
  {address:"Model!A5", value:"Profit after bonus"},  {address:"Model!B5", formula:"=Model!B3"},
  {address:"Model!A6", value:"Tax"},                 {address:"Model!B6", formula:"=Model!B5*Inputs!B3"},
  {address:"Model!A7", value:"Net income"},          {address:"Model!B7", formula:"=Model!B5-Model!B6"},
])
const r = await xlsx.readRange(wb, {sheet:"Model", from:{row:1,col:2}, to:{row:7,col:2}})
return r.map(row => row[0].value)
