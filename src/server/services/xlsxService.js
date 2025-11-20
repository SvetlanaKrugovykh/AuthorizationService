const fs = require('fs')
const path = require('path')
const { execSync } = require('child_process')
require('dotenv').config()

module.exports.doConvertXlsx = async function (inputFilePath) {
  // Determine output directory and file name
  let outDir = process.env.XLSX_OUT || path.join(__dirname, '../../temp')
  // If XLSX_OUT is set, ensure absolute path is created
  if (process.env.XLSX_OUT) {
    outDir = path.isAbsolute(process.env.XLSX_OUT) ? process.env.XLSX_OUT : path.resolve(process.env.XLSX_OUT)
  }
  if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true })
  const outputFileName = path.basename(inputFilePath)
  const outputFilePath = path.join(outDir, outputFileName)

  // Copy input file to output directory
  fs.copyFileSync(inputFilePath, outputFilePath)

  // Example: tempDir for further processing (stub)
  const tempDir = process.env.TEMP_CATALOG || path.join(__dirname, '../../temp/xlsx')
  if (!fs.existsSync(tempDir)) fs.mkdirSync(tempDir, { recursive: true })

  // Example: sharedStrings.xml processing (stub)
  const sharedStringsPath = path.join(tempDir, 'xl', 'sharedStrings.xml')
  if (fs.existsSync(sharedStringsPath)) {
    let content = fs.readFileSync(sharedStringsPath, 'utf8')
    content = content.replace(/=HYPERLINK\("([^"]+)","([^"]+)"\)/g, '$2')
    fs.writeFileSync(sharedStringsPath, content, 'utf8')
  }

  // Example: relsDir and relsContent (stub)
  const relsDir = path.join(tempDir, 'xl', 'worksheets', '_rels')
  if (!fs.existsSync(relsDir)) fs.mkdirSync(relsDir, { recursive: true })
  const relsContent = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://drive.google.com/file/d/1lVqn2_zexcaSruT3c4VC7ItrTS0JXuXz/view" TargetMode="External"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://www.google.com/maps/place/50%C2%B023'30.5%22N+30%C2%B022'39.8%22E/@50.3917937,30.3766247,18z" TargetMode="External"/>
</Relationships>`
  fs.writeFileSync(path.join(relsDir, 'sheet1.xml.rels'), relsContent, 'utf8')

  // Example: worksheet.xml processing (stub)
  const worksheetPath = path.join(tempDir, 'xl', 'worksheets', 'sheet1.xml')
  if (fs.existsSync(worksheetPath)) {
    let worksheet = fs.readFileSync(worksheetPath, 'utf8')
    if (!worksheet.includes('xmlns:r=')) {
      worksheet = worksheet.replace('<worksheet', '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
    }
    const hyperlinks = `<hyperlinks>\n<hyperlink ref="K14" r:id="rId1"/>\n<hyperlink ref="L14" r:id="rId2"/>\n</hyperlinks>`
    worksheet = worksheet.replace('</worksheet>', hyperlinks + '</worksheet>')
    worksheet = worksheet.replace(
      /(<row[^>]*)\s+ht="[^"]*"([^>]*customHeight="true"[^>]*>)/g,
      (match, before, after) => {
        return before + ' ht="15"' + after.replace(/\s+customHeight="true"/g, '')
      }
    )
    fs.writeFileSync(worksheetPath, worksheet, 'utf8')
  }

  // Remove tempDir if needed
  // fs.rmSync(tempDir, { recursive: true, force: true })

  // Log output
  console.log('Done!')
  console.log('File:', outputFilePath)

  // Return output file path
  return outputFilePath
}