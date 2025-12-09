const fs = require('fs')
const path = require('path')
require('dotenv').config()

module.exports.doConvertXlsx = async function (inputFilePath) {
  const ExcelJS = require('exceljs')
  const outDir = process.env.XLSX_OUT || path.join(__dirname, '../../temp')
  if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true })
  const outputFileName = path.basename(inputFilePath)
  const outputFilePath = path.join(outDir, outputFileName)

  // Map of text markers to URLs
  const hyperlinksMap = {
    '[Фото]': 'https://drive.google.com/file/d/1lVqn2_zexcaSruT3c4VC7ItrTS0JXuXz/view',
    '[Мапа]': 'https://www.google.com/maps/place/50%C2%B023\'30.5%22N+30%C2%B022\'39.8%22E/@50.3917937,30.3766247,18z'
  }

  try {
    // Read workbook
    const workbook = new ExcelJS.Workbook()
    await workbook.xlsx.readFile(inputFilePath)

    // Process each worksheet
    workbook.worksheets.forEach(worksheet => {
      worksheet.eachRow(row => {
        row.eachCell(cell => {
          const cellValue = cell.value ? cell.value.toString().trim() : ''

          // Check if cell contains a hyperlink marker
          for (const [marker, url] of Object.entries(hyperlinksMap)) {
            if (cellValue === marker) {
              // Convert to proper hyperlink
              cell.value = marker
              cell.hyperlink = url
              cell.font = { color: { argb: 'FF0563C1' }, underline: true }
              break
            }
          }
        })
      })
    })

    // Write converted workbook
    await workbook.xlsx.writeFile(outputFilePath)
    return outputFilePath
  } catch (err) {
    console.error('Error converting XLSX:', err.message)
    throw err
  }
}