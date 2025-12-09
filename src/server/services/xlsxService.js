const fs = require('fs')
const path = require('path')
require('dotenv').config()

module.exports.doConvertXlsx = async function (inputFilePath) {
  const originalFile = inputFilePath
  const outDir = process.env.XLSX_OUT || path.join(__dirname, '../../temp')
  if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true })
  const outputFileName = path.basename(originalFile)
  const outputFilePath = path.join(outDir, outputFileName)

  fs.copyFileSync(originalFile, outputFilePath)

  const tempDir = path.join(outDir, 'temp_final')
  fs.mkdirSync(tempDir, { recursive: true })

  // Extract XLSX archive (platform-independent)
  const AdmZip = require('adm-zip')
  const zip = new AdmZip(outputFilePath)
  zip.extractAllTo(tempDir, true)

  // Remove HYPERLINK formulas and extract links
  const sharedStringsPath = path.join(tempDir, 'xl', 'sharedStrings.xml')
  let content = fs.readFileSync(sharedStringsPath, 'utf8')
  
  // Extract URLs from HYPERLINK formulas
  const hyperlinksMap = {}
  const hyperlinkRegex = /=HYPERLINK\("([^"]+)","([^"]+)"\)/g
  let match
  while ((match = hyperlinkRegex.exec(content)) !== null) {
    hyperlinksMap[match[2]] = match[1]  // marker -> URL
  }

  // Replace HYPERLINK formulas with just the display text
  content = content.replace(/=HYPERLINK\("([^"]+)","([^"]+)"\)/g, '$2')
  fs.writeFileSync(sharedStringsPath, content, 'utf8')

  // Keep sharedStrings as is - just remove HYPERLINK formulas  
  content = content.replace(/=HYPERLINK\("([^"]+)","([^"]+)"\)/g, '$2')
  fs.writeFileSync(sharedStringsPath, content, 'utf8')

  // Modify worksheet cells to contain HYPERLINK formulas directly
  const worksheetPath = path.join(tempDir, 'xl', 'worksheets', 'sheet1.xml')
  let worksheet = fs.readFileSync(worksheetPath, 'utf8')

  for (const [marker, url] of Object.entries(hyperlinksMap)) {
    // Find cells that reference sharedStrings containing the marker
    const cellPattern = /<c r="([A-Z]+\d+)"([^>]*) t="s"><v>(\d+)<\/v><\/c>/g
    let cellMatch
    
    while ((cellMatch = cellPattern.exec(worksheet)) !== null) {
      const cellRef = cellMatch[1]
      const cellAttrs = cellMatch[2]
      const stringIndex = cellMatch[3]
      
      // Check if this sharedString index contains our marker
      const siPattern = /<si[^>]*>[\s\S]*?<\/si>/g
      let siMatch
      let siCount = 0
      
      while ((siMatch = siPattern.exec(content)) !== null) {
        if (siCount === parseInt(stringIndex)) {
          if (siMatch[0].includes(marker)) {
            // Replace this cell with HYPERLINK formula
            const hyperlinkFormula = `=HYPERLINK("${url}","${marker}")`
            const newCell = `<c r="${cellRef}"${cellAttrs}><f>${hyperlinkFormula}</f></c>`
            worksheet = worksheet.replace(cellMatch[0], newCell)
          }
          break
        }
        siCount++
      }
    }
  }

  fs.writeFileSync(worksheetPath, worksheet, 'utf8')
  zip.updateFile('xl/worksheets/sheet1.xml', Buffer.from(worksheet, 'utf8'))

  zip.writeZip(outputFilePath)

  fs.rmSync(tempDir, { recursive: true, force: true })

  return outputFilePath
}