const fs = require('fs')
const path = require('path')
require('dotenv').config()

function convertToHyperlinks(inputFilePath, outputFilePath) {
  if (!outputFilePath) {
    const outDir = process.env.XLSX_OUT || path.join(__dirname, '../../temp')
    if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true })
    const outputFileName = path.basename(inputFilePath)
    outputFilePath = path.join(outDir, outputFileName)
  }

  // Ensure output directory exists
  const outDir = path.dirname(outputFilePath)
  if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true })

  fs.copyFileSync(inputFilePath, outputFilePath)

  const tempDir = path.join(outDir, 'temp_final')
  fs.mkdirSync(tempDir, { recursive: true })

  // Extract XLSX archive (platform-independent)
  const AdmZip = require('adm-zip')
  const zip = new AdmZip(outputFilePath)
  zip.extractAllTo(tempDir, true)

  // 1. SCAN SHAREDSTRINGS.XML - extract HYPERLINK formulas dynamically
  const sharedStringsPath = path.join(tempDir, 'xl', 'sharedStrings.xml')
  let sharedStrings = fs.readFileSync(sharedStringsPath, 'utf8')
  
  // Extract ALL HYPERLINK formulas dynamically
  const hyperlinkMap = {}
  const hyperlinkPattern = /=HYPERLINK\("([^"]+)","([^"]+)"\)/g
  let match

  while ((match = hyperlinkPattern.exec(sharedStrings)) !== null) {
    const url = match[1]
    const displayText = match[2]
    hyperlinkMap[displayText] = url
  }

  if (Object.keys(hyperlinkMap).length === 0) {
    fs.rmSync(tempDir, { recursive: true, force: true })
    return outputFilePath
  }

  // 2. SCAN WORKSHEET.XML - find cells containing markers BEFORE cleaning
  const worksheetPath = path.join(tempDir, 'xl', 'worksheets', 'sheet1.xml')
  let worksheet = fs.readFileSync(worksheetPath, 'utf8')

  // Find ALL cells and check if their sharedString values contain our markers
  const cellsToConvert = []
  // FIXED: Allow multiline content with whitespace
  const cellPattern = /<c r="([A-Z]+\d+)"([^>]*)>\s*<v>(\d+)<\/v>\s*<\/c>/g
  let cellMatch

  // Get all sharedString entries for quick lookup (from ORIGINAL content)
  const originalSharedStringEntries = []
  const siPattern = /<si[^>]*>([\s\S]*?)<\/si>/g
  let siMatch
  while ((siMatch = siPattern.exec(sharedStrings)) !== null) {
    originalSharedStringEntries.push(siMatch[1])
  }



  // Check each cell
  let cellCount = 0
  while ((cellMatch = cellPattern.exec(worksheet)) !== null) {
    const cellRef = cellMatch[1]
    const cellAttrs = cellMatch[2]
    const stringIndex = parseInt(cellMatch[3])
    cellCount++
    
    // Debug removed for production
    
    if (stringIndex < originalSharedStringEntries.length) {
      const sharedStringContent = originalSharedStringEntries[stringIndex]
      
      // Check if this sharedString contains any of our markers
      for (const [marker, url] of Object.entries(hyperlinkMap)) {
        if (sharedStringContent.includes(marker)) {
          cellsToConvert.push({
            cellRef: cellRef,
            marker: marker,
            url: url,
            originalCell: cellMatch[0],
            cellAttrs: cellAttrs,
            stringIndex: stringIndex
          })
          break // Only one marker per cell
        }
      }
    }
  }



  // 3. NOW CLEAN HYPERLINK formulas from sharedStrings
  let sharedStringsModified = false
  for (const [marker, url] of Object.entries(hyperlinkMap)) {
    const hyperlinkFormula = `=HYPERLINK("${url}","${marker}")`
    if (sharedStrings.includes(hyperlinkFormula)) {
      sharedStrings = sharedStrings.replace(hyperlinkFormula, marker)
      sharedStringsModified = true
    }
  }

  if (cellsToConvert.length === 0) {
    fs.rmSync(tempDir, { recursive: true, force: true })
    return outputFilePath
  }

  // 4. CREATE RELATIONSHIPS
  const relsDir = path.join(tempDir, 'xl', 'worksheets', '_rels')
  const relsFilePath = path.join(relsDir, 'sheet1.xml.rels')
  
  if (!fs.existsSync(relsDir)) {
    fs.mkdirSync(relsDir, { recursive: true })
  }

  let relationships = ''
  if (fs.existsSync(relsFilePath)) {
    relationships = fs.readFileSync(relsFilePath, 'utf8')
  } else {
    relationships = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`
  }

  // Find highest existing ID
  let maxId = 0
  const existingIds = relationships.match(/Id="rId(\d+)"/g) || []
  existingIds.forEach(idMatch => {
    const id = parseInt(idMatch.match(/rId(\d+)/)[1])
    if (id > maxId) maxId = id
  })

  // Add relationships
  let relId = maxId + 1
  const cellToRelationshipId = {}
  
  for (const cell of cellsToConvert) {
    const relationshipId = `rId${relId}`
    cellToRelationshipId[cell.cellRef] = relationshipId
    
    const relationshipXml = `<Relationship Id="${relationshipId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${cell.url}" TargetMode="External"/>`
    relationships = relationships.replace('</Relationships>', `${relationshipXml}\n</Relationships>`)
    
    relId++
  }

  fs.writeFileSync(relsFilePath, relationships, 'utf8')

  // 5. UPDATE SHAREDSTRINGS  
  if (sharedStringsModified) {
    fs.writeFileSync(sharedStringsPath, sharedStrings, 'utf8')
  }

  // 6. ADD HYPERLINKS SECTION TO WORKSHEET (MS Office compatible)
  if (cellsToConvert.length > 0) {
    // Ensure worksheet has proper namespace declarations
    if (!worksheet.includes('xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')) {
      // Add namespace to worksheet root element if missing
      worksheet = worksheet.replace(
        /<worksheet([^>]*)>/,
        '<worksheet$1 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
      )
    }
    
    const hyperlinksXml = cellsToConvert.map(cell => 
      `<hyperlink ref="${cell.cellRef}" r:id="${cellToRelationshipId[cell.cellRef]}"/>`
    ).join('\n')
    
    const hyperlinksSection = `<hyperlinks>\n${hyperlinksXml}\n</hyperlinks>\n</worksheet>`
    worksheet = worksheet.replace('</worksheet>', hyperlinksSection)
  }

  // 7. ENSURE SINGLE PROPER XML DECLARATION
  // First, remove ALL XML declarations
  const xmlDeclarationPattern = /<\?xml[^>]*>\s*/g
  const cleanWorksheet = worksheet.replace(xmlDeclarationPattern, '').trim()
  
  // Then add exactly one at the beginning
  worksheet = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + cleanWorksheet

  // 7. SAVE EVERYTHING
  fs.writeFileSync(worksheetPath, worksheet, 'utf8')
  zip.updateFile('xl/worksheets/sheet1.xml', Buffer.from(worksheet, 'utf8'))
  zip.updateFile('xl/sharedStrings.xml', Buffer.from(sharedStrings, 'utf8'))

  // Add relationships file to ZIP
  if (fs.existsSync(relsFilePath)) {
    zip.addLocalFile(relsFilePath, 'xl/worksheets/_rels/', 'sheet1.xml.rels')
  }

  zip.writeZip(outputFilePath)
  
  fs.rmSync(tempDir, { recursive: true, force: true })

  return outputFilePath
}

// Alternative method using xlsx (SheetJS) library for better MS Office compatibility
function convertWithSheetJS(inputFilePath, outputFilePath) {
  const XLSX = require('xlsx')
  
  if (!outputFilePath) {
    const outDir = process.env.XLSX_OUT || path.join(__dirname, '../../temp')
    if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true })
    const outputFileName = 'SheetJS_' + path.basename(inputFilePath)
    outputFilePath = path.join(outDir, outputFileName)
  }

  try {
    // Step 1: Extract hyperlinks using our proven method
    const AdmZip = require('adm-zip')
    const zip = new AdmZip(inputFilePath)
    const tempDir = path.join(path.dirname(outputFilePath), 'temp_sheetjs')
    fs.mkdirSync(tempDir, { recursive: true })
    zip.extractAllTo(tempDir, true)

    const sharedStringsPath = path.join(tempDir, 'xl', 'sharedStrings.xml')
    const sharedStrings = fs.readFileSync(sharedStringsPath, 'utf8')
    
    const hyperlinkMap = {}
    const hyperlinkPattern = /=HYPERLINK\("([^"]+)","([^"]+)"\)/g
    let match

    while ((match = hyperlinkPattern.exec(sharedStrings)) !== null) {
      const url = match[1]
      const displayText = match[2]
      hyperlinkMap[displayText] = url
    }

    fs.rmSync(tempDir, { recursive: true, force: true })

    if (Object.keys(hyperlinkMap).length === 0) {
      return null
    }

    // Step 2: Read with SheetJS and convert
    const workbook = XLSX.readFile(inputFilePath)
    const sheetName = workbook.SheetNames[0]
    const worksheet = workbook.Sheets[sheetName]
    const range = XLSX.utils.decode_range(worksheet['!ref'])

    let convertedCount = 0
    
    for (let row = range.s.r; row <= range.e.r; row++) {
      for (let col = range.s.c; col <= range.e.c; col++) {
        const cellAddress = XLSX.utils.encode_cell({ r: row, c: col })
        const cell = worksheet[cellAddress]
        
        if (cell && cell.v && typeof cell.v === 'string') {
          for (const [marker, url] of Object.entries(hyperlinkMap)) {
            if (cell.v.includes(marker)) {
              cell.v = marker
              cell.t = 's'
              
              if (!worksheet['!links']) worksheet['!links'] = {}
              worksheet['!links'][cellAddress] = {
                Target: url,
                Tooltip: marker
              }
              
              convertedCount++
              break
            }
          }
        }
      }
    }

    if (convertedCount === 0) {
      return null
    }

    // Write with SheetJS
    XLSX.writeFile(workbook, outputFilePath)
    return outputFilePath

  } catch (error) {
    return null
  }
}

module.exports.doConvertXlsx = async function (inputFilePath) {
  const originalFile = inputFilePath
  const outDir = process.env.XLSX_OUT || path.join(__dirname, '../../temp')
  if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true })
  const outputFileName = path.basename(originalFile)
  const outputFilePath = path.join(outDir, outputFileName)

  // Try SheetJS method first (better MS Office compatibility)
  const sheetJSResult = convertWithSheetJS(originalFile, path.join(outDir, 'SheetJS_' + outputFileName))
  if (sheetJSResult) {
    return sheetJSResult
  }

  // Fallback to custom XML method
  return convertToHyperlinks(originalFile, outputFilePath)
}

module.exports.convertToHyperlinks = convertToHyperlinks
module.exports.convertWithSheetJS = convertWithSheetJS