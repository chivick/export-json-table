const docDefinition = {
  content: [{
    table: {
      body: [
        [
          "First", "Second", "Third","The last one"
        ],
        [
          "Value 1", "Value 2", "Value 3", "Value 4",
        ],
        [
          { text: "Bold Value", bold: true}, "Value 2", "Value 3", "Value 4",
        ],
      ]
    }
  }]
}

const pdf = pdfMake.createPdf(docDefinition)

document.getElementById("download-pdf").addEventListener("click", () => {
  pdf.download("Report.pdf")
})



const table = document.getElementById("myTable")


const exportTableToExcel = () => {
  let wb = XLSX.utils.table_to_book(table)

  XLSX.writeFile(wb, "SheetJSTable.xlsx", { compression: true })
}

const exportTableToCSV = () => {
  let wb = XLSX.utils.table_to_book(table)

  XLSX.writeFile(wb, "SheetJSTable.csv", { type: "csv", compression: true })
}

const exportJsonToExcel = () => {
  const jsonData = [
    {
      name: "John Doe",
      age: 29
    },
    {
      name: "Michael Jordan",
      age: 34
    },
    {
      name: "Sarah Smith",
      age: 23
    }
  ]

  const worksheet = XLSX.utils.json_to_sheet(jsonData)
  const workbook = XLSX.utils.book_new()

  XLSX.utils.book_append_sheet(workbook, worksheet, "Users")

  // at this point the excel file can now be exported
  // but a little more cleaning up could be done for improvement

  // change the headers of the excel sheet, from cell "A1"
  XLSX.utils.sheet_add_aoa(worksheet, [["Name","Age"]], { origin: "A1" })

  // changing column width
  const max_width = jsonData.reduce((w,r) => Math.max(w, r.name.length), 10)
  // setting the width of column A
  worksheet["!cols"] = [{ wch : max_width}]

  // Export file
  XLSX.writeFile(workbook, "JsonData.xlsx", { compression: true })
}

const exportJsonToCSV = () => {
  const jsonData = [
    {
      name: "John Doe",
      age: 29
    },
    {
      name: "Michael Jordan",
      age: 34
    },
    {
      name: "Sarah Smith",
      age: 23
    }
  ]

  const worksheet = XLSX.utils.json_to_sheet(jsonData)
  const workbook = XLSX.utils.book_new()

  XLSX.utils.book_append_sheet(workbook, worksheet, "Users")

  // at this point the excel file can now be exported
  // but a little more cleaning up could be done for improvement

  // change the headers of the excel sheet, from cell "A1"
  XLSX.utils.sheet_add_aoa(worksheet, [["Name","Age"]], { origin: "A1" })

  // changing column width
  const max_width = jsonData.reduce((w,r) => Math.max(w, r.name.length), 10)
  // setting the width of column A
  worksheet["!cols"] = [{ wch : max_width}]

  // Export file
  XLSX.writeFile(workbook, "JsonData.csv", { type: "csv", compression: true })
}

const printTable = () => {
  const printWindow = window.open(",","PRINT")
  
  // add some css
  printWindow.document.write("<html><head><style>") 
  printWindow.document.write("table { border-collapse: collapse; border-spacing: 0; border: 1px solid #ddd; width: 100%; }") 
  printWindow.document.write("th, td { text-align: left; padding: 16px; border: 1px solid #ddd; }")
  printWindow.document.write("tr:nth-child(even) { background-color: #f2f2f2; }")
  printWindow.document.write("th { font-size: 20px; }")
  printWindow.document.write("</style></head><body>")
  printWindow.document.write(table.outerHTML)
  printWindow.document.write("</body></html>")
  printWindow.document.close()
  printWindow.focus()
  printWindow.print()
}

const printJsonTable = () => {
  const printWindow = window.open(",","PRINT")

  const jsonData = [
    {
      name: "John Doe",
      age: 29
    },
    {
      name: "Michael Jordan",
      age: 34
    },
    {
      name: "Sarah Smith",
      age: 23
    }
  ]

  let tableHtml = `<table><tr>`

  Object.keys(jsonData[0]).forEach(key => {
    tableHtml += `<th>${key.toUpperCase()}</th>`
  })
  
  tableHtml += `</tr>`

  jsonData.forEach(row => {
    tableHtml += `<tr><td>${row.name}</td><td>${row.age}</td></tr>`
  })

  tableHtml += `</table>`
  

  // add some css
  printWindow.document.write("<html><head><style>") 
  printWindow.document.write("table { border-collapse: collapse; border-spacing: 0; border: 1px solid #ddd; width: 100%; }") 
  printWindow.document.write("th, td { text-align: left; padding: 16px; border: 1px solid #ddd; }")
  printWindow.document.write("tr:nth-child(even) { background-color: #f2f2f2; }")
  printWindow.document.write("th { font-size: 20px; }") 
  printWindow.document.write("</style></head><body>")
  printWindow.document.write(tableHtml)
  printWindow.document.write("</body></html>")
  printWindow.document.close()
  printWindow.focus()
  printWindow.print()
}


const copyTableToClipboard = () => {
  // let range = document.createRange()
  // range.selectNode(table)

  // window.getSelection().addRange(range)

  // document.execCommand()

  let text = table.textContent
  navigator.clipboard.writeText(text).then(() => {
    alert("Copied to clipboard")
  }, (error) => {
    alert("Error")
  })
}

document.getElementById("download-excel").addEventListener("click", exportTableToExcel)
document.getElementById("download-json-excel").addEventListener("click", exportJsonToExcel)
document.getElementById("download-json-csv").addEventListener("click", exportJsonToCSV)
document.getElementById("download-csv").addEventListener("click", exportTableToCSV)
document.getElementById("download-print").addEventListener("click", printTable)
document.getElementById("download-json-print").addEventListener("click", printJsonTable)
document.getElementById("copy-to-clipboard").addEventListener("click", copyTableToClipboard)
