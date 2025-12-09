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

  // Read existing relationships and find max ID
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

  // Create relationship entries for all hyperlinks
  let newRelsXml = ''
  const hyperlinksToAdd = []
  let nextId = maxId + 1

  for (const [marker, url] of Object.entries(hyperlinksMap)) {
    const relId = `rId${nextId}`
    newRelsXml += `        <Relationship Id="${relId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${url}" TargetMode="External"/>\n`
    hyperlinksToAdd.push({ marker, relId, url })
    nextId++
  }

  // Update or create relationships file
  let relsContent
  if (existingRels) {
    relsContent = existingRels.replace('</Relationships>', newRelsXml + '</Relationships>')
  } else {
    relsContent = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${newRelsXml}</Relationships>`
  }
  fs.writeFileSync(relsPath, relsContent, 'utf8')

  // Find cells with hyperlink markers and build hyperlinks element
  const worksheetPath = path.join(tempDir, 'xl', 'worksheets', 'sheet1.xml')
  let worksheet = fs.readFileSync(worksheetPath, 'utf8')
  if (!worksheet.includes('xmlns:r=')) {
    worksheet = worksheet.replace('<worksheet', '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
  }

  // Parse worksheet to find cells with markers and build hyperlinks
  let hyperlinkElements = ''
  
  // For each marker, find cells that contain it
  for (const { marker, relId } of hyperlinksToAdd) {
    // Look for cells containing the marker text value
    // The marker might be wrapped as: <c r="K14" ... t="s"><v>123</v></c> where shared string 123 contains the marker
    // But we need to find which cells actually display the marker
    // Simple approach: find all cells and check their content
    
    // Match pattern: <c r="CELLREF" ... >...<v>INDEX</v>...</c>
    // Then check if that index in sharedStrings contains our marker
    const cellPattern = /<c r="([A-Z]+\d+)"[^>]*>[\s\S]*?<v>(\d+)<\/v>[\s\S]*?<\/c>/g
    let cellMatch
    
    while ((cellMatch = cellPattern.exec(worksheet)) !== null) {
      const cellRef = cellMatch[1]
      const stringIndex = cellMatch[2]
      
      // Check if this index in sharedStrings contains our marker
      // Find the stringIndex-th <si> block
      const siPattern = /<si[^>]*>[\s\S]*?<\/si>/g
      let siMatch
      let siCount = 0
      
      while ((siMatch = siPattern.exec(content)) !== null) {
        if (siCount === parseInt(stringIndex)) {
          // Found the matching <si> block
          if (siMatch[0].includes(marker)) {
            // This cell contains our marker!
            hyperlinkElements += `          <hyperlink ref="${cellRef}" r:id="${relId}"/>\n`
          }
          break
        }
        siCount++
      }
    }
  }

  if (hyperlinkElements) {
    const hyperlinks = `        <hyperlinks>\n${hyperlinkElements}        </hyperlinks>`
    worksheet = worksheet.replace(/(<\/worksheet>)$/m, hyperlinks + '\n$1')
  }

  fs.writeFileSync(worksheetPath, worksheet, 'utf8')

  // Update files inside existing archive (preserves original structure and MS Office compatibility)  
  zip.updateFile('xl/sharedStrings.xml', Buffer.from(content, 'utf8'))
  zip.updateFile('xl/worksheets/sheet1.xml', Buffer.from(worksheet, 'utf8'))
  
  // Add or update relationships file
  const relsFileInZip = 'xl/worksheets/_rels/sheet1.xml.rels'
  if (zip.getEntry(relsFileInZip)) {
    zip.updateFile(relsFileInZip, Buffer.from(relsContent, 'utf8'))
  } else {
    zip.addFile(relsFileInZip, Buffer.from(relsContent, 'utf8'))
  }

  zip.writeZip(outputFilePath)

  fs.rmSync(tempDir, { recursive: true, force: true })

  return outputFilePath
}