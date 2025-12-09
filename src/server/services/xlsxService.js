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

  // Remove HYPERLINK formulas
  const sharedStringsPath = path.join(tempDir, 'xl', 'sharedStrings.xml')
  let content = fs.readFileSync(sharedStringsPath, 'utf8')
  content = content.replace(/=HYPERLINK\("([^"]+)","([^"]+)"\)/g, '$2')
  fs.writeFileSync(sharedStringsPath, content, 'utf8')

  // Create relationships for hyperlinks
  const relsDir = path.join(tempDir, 'xl', 'worksheets', '_rels')
  fs.mkdirSync(relsDir, { recursive: true })
  const relsContent = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://drive.google.com/file/d/1lVqn2_zexcaSruT3c4VC7ItrTS0JXuXz/view" TargetMode="External"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://www.google.com/maps/place/50%C2%B023'30.5%22N+30%C2%B022'39.8%22E/@50.3917937,30.3766247,18z" TargetMode="External"/>
</Relationships>`
  fs.writeFileSync(path.join(relsDir, 'sheet1.xml.rels'), relsContent, 'utf8')

  // Add hyperlinks to worksheet
  const worksheetPath = path.join(tempDir, 'xl', 'worksheets', 'sheet1.xml')
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

  // Recreate XLSX archive (platform-independent)
  fs.unlinkSync(outputFilePath)

  const newZip = new AdmZip()
  function addDirToZip(dir, zipInstance, basePath = '') {
    const items = fs.readdirSync(dir)
    for (const item of items) {
      const fullPath = path.join(dir, item)
      const relPath = path.join(basePath, item)
      if (fs.statSync(fullPath).isDirectory()) {
        addDirToZip(fullPath, zipInstance, relPath)
      } else {
        zipInstance.addLocalFile(fullPath, path.dirname(relPath))
      }
    }
  }
  addDirToZip(tempDir, newZip)
  newZip.writeZip(outputFilePath)

  fs.rmSync(tempDir, { recursive: true, force: true })

  return outputFilePath
}