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

  // SIMPLE APPROACH: Create inline HYPERLINK formulas in sharedStrings
  // Don't touch relationships at all - just put back working HYPERLINK formulas
  
  for (const [marker, url] of Object.entries(hyperlinksMap)) {
    const hyperlinkFormula = `=HYPERLINK("${url}","${marker}")`
    content = content.replace(new RegExp(marker.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g'), hyperlinkFormula)
  }
  
  fs.writeFileSync(sharedStringsPath, content, 'utf8')

  zip.writeZip(outputFilePath)

  fs.rmSync(tempDir, { recursive: true, force: true })

  return outputFilePath
}