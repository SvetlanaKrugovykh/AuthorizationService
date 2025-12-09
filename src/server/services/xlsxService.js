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
  const hyperlinksMap = {
    '[Фото]': 'https://drive.google.com/file/d/1lVqn2_zexcaSruT3c4VC7ItrTS0JXuXz/view',
    '[Мапа]': 'https://www.google.com/maps/place/50%C2%B023\'30.5%22N+30%C2%B022\'39.8%22E/@50.3917937,30.3766247,18z'
  }

  // Replace HYPERLINK formulas
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
    // Find max rId number
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
    newRelsXml += `<Relationship Id="${relId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${url}" TargetMode="External"/>\n`
    hyperlinksToAdd.push({ marker, relId })
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
  
  // First, find indices of markers in sharedStrings
  const sharedStringsContent = fs.readFileSync(sharedStringsPath, 'utf8')
  const markerIndices = {}
  
  for (const [marker, url] of Object.entries(hyperlinksMap)) {
    // Find marker in sharedStrings: <t>marker</t> or <t>marker</t> with variations
    const markerRegex = new RegExp(`<t[^>]*>${marker}</t>`, 'g')
    let match
    let index = 0
    let count = 0
    
    // Count which index this marker appears at
    const tempContent = sharedStringsContent
    const allMatches = tempContent.match(/<si[^>]*>[\s\S]*?<\/si>/g) || []
    for (let i = 0; i < allMatches.length; i++) {
      if (allMatches[i].includes(marker)) {
        if (!markerIndices[marker]) {
          markerIndices[marker] = i
        }
      }
    }
  }

  // Now find cells in worksheet that reference these indices
  for (const { marker, relId } of hyperlinksToAdd) {
    const markerIndex = markerIndices[marker]
    if (markerIndex !== undefined) {
      // Find cells with v={markerIndex}
      const cellRegex = new RegExp(`<c r="([A-Z]+\\d+)"[^>]*t="s"[^>]*>\\s*<v>${markerIndex}</v>\\s*</c>`, 'g')
      let match
      while ((match = cellRegex.exec(worksheet)) !== null) {
        hyperlinkElements += `<hyperlink ref="${match[1]}" r:id="${relId}"/>\n`
      }
    }
  }

  if (hyperlinkElements) {
    const hyperlinks = `<hyperlinks>\n${hyperlinkElements}</hyperlinks>`
    worksheet = worksheet.replace('</worksheet>', hyperlinks + '</worksheet>')
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