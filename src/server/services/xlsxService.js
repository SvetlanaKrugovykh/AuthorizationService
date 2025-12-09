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

  // STEP 1: Add only relationships, no worksheet hyperlinks yet
  const relsDir = path.join(tempDir, 'xl', 'worksheets', '_rels')
  fs.mkdirSync(relsDir, { recursive: true })
  const relsPath = path.join(relsDir, 'sheet1.xml.rels')
  let existingRels = ''
  let maxId = 0

  if (fs.existsSync(relsPath)) {
    existingRels = fs.readFileSync(relsPath, 'utf8')
    const idMatches = existingRels.match(/Id="rId(\d+)"/g)
    if (idMatches) {
      idMatches.forEach(match => {
        const num = parseInt(match.replace(/\D/g, ''))
        if (num > maxId) maxId = num
      })
    }
  }

  let newRelsXml = ''
  let nextId = maxId + 1
  for (const [marker, url] of Object.entries(hyperlinksMap)) {
    newRelsXml += `        <Relationship Id="rId${nextId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${url}" TargetMode="External"/>\n`
    nextId++
  }

  if (newRelsXml) {
    let relsContent
    if (existingRels) {
      relsContent = existingRels.replace('</Relationships>', newRelsXml + '</Relationships>')
    } else {
      relsContent = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${newRelsXml}</Relationships>`
    }
    fs.writeFileSync(relsPath, relsContent, 'utf8')

    // Add relationships to ZIP
    const relsFileInZip = 'xl/worksheets/_rels/sheet1.xml.rels'
    if (zip.getEntry(relsFileInZip)) {
      zip.updateFile(relsFileInZip, Buffer.from(relsContent, 'utf8'))
    } else {
      zip.addFile(relsFileInZip, Buffer.from(relsContent, 'utf8'))
    }
  }

  // STEP 2: Add hyperlinks block to worksheet
  const worksheetPath = path.join(tempDir, 'xl', 'worksheets', 'sheet1.xml')
  let worksheet = fs.readFileSync(worksheetPath, 'utf8')

  // Find cells with hyperlink markers and build hyperlinks element
  let hyperlinkElements = ''
  nextId = maxId + 1 // Reset nextId for hyperlinks
  
  for (const [marker, url] of Object.entries(hyperlinksMap)) {
    const relId = `rId${nextId}`
    
    // Find cells containing the marker text
    const cellPattern = /<c r="([A-Z]+\d+)"[^>]*>[\s\S]*?<v>(\d+)<\/v>[\s\S]*?<\/c>/g
    let cellMatch
    
    while ((cellMatch = cellPattern.exec(worksheet)) !== null) {
      const cellRef = cellMatch[1]
      const stringIndex = cellMatch[2]
      
      // Check if this index in sharedStrings contains our marker
      const siPattern = /<si[^>]*>[\s\S]*?<\/si>/g
      let siMatch
      let siCount = 0
      
      while ((siMatch = siPattern.exec(content)) !== null) {
        if (siCount === parseInt(stringIndex)) {
          if (siMatch[0].includes(marker)) {
            hyperlinkElements += `          <hyperlink ref="${cellRef}" r:id="${relId}"/>\n`
          }
          break
        }
        siCount++
      }
    }
    nextId++
  }

  // Add hyperlinks block if we found any
  if (hyperlinkElements) {
    const hyperlinks = `        <hyperlinks>\n${hyperlinkElements}        </hyperlinks>`
    worksheet = worksheet.replace(/(<legacyDrawingHF[^>]*\/>)/, `$1\n${hyperlinks}`)
  }

  fs.writeFileSync(worksheetPath, worksheet, 'utf8')
  zip.updateFile('xl/worksheets/sheet1.xml', Buffer.from(worksheet, 'utf8'))

  zip.writeZip(outputFilePath)

  fs.rmSync(tempDir, { recursive: true, force: true })

  return outputFilePath
}