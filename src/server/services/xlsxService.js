const fs = require('fs')
const path = require('path')
const XLSX = require('xlsx')

/**
 * MS Office compatible XLSX hyperlink converter
 * Uses professional XLSX library for guaranteed compatibility
 * Converts HYPERLINK formulas from 1C to real clickable hyperlinks
 */
function convertToHyperlinks(inputFilePath, outputFilePath) {
  console.log('🚀 Starting MS Office compatible conversion...')
  
  if (!outputFilePath) {
    const outDir = process.env.XLSX_OUT || path.join(__dirname, '../../temp')
    if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true })
    const outputFileName = path.basename(inputFilePath)
    outputFilePath = path.join(outDir, outputFileName)
  }

  // Ensure output directory exists
  const outDir = path.dirname(outputFilePath)
  if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true })

  console.log('📖 Reading file:', inputFilePath)
  
  // Read the workbook with XLSX library (MS Office compatible)
  const workbook = XLSX.readFile(inputFilePath, {
    cellHTML: false,
    cellNF: false,
    cellStyles: false,
    sheetStubs: false,
    cellFormula: true,
    cellText: false
  })

  console.log('📋 Found sheets:', workbook.SheetNames)
  
  const sheetName = workbook.SheetNames[0]
  const sheet = workbook.Sheets[sheetName]
  
  console.log('📏 Sheet range:', sheet['!ref'])

  // Find all cells that need hyperlink conversion
  const hyperlinksToCreate = []
  const range = XLSX.utils.decode_range(sheet['!ref'])
  
  console.log('🔍 Scanning for HYPERLINK in cell VALUES (1C format)...')
  
  for (let R = range.s.r; R <= range.e.r; R++) {
    for (let C = range.s.c; C <= range.e.c; C++) {
      const cellAddress = XLSX.utils.encode_cell({r: R, c: C})
      const cell = sheet[cellAddress]
      
      if (cell && cell.v) {
        const cellValue = cell.v.toString()
        
        // Check if cell value contains HYPERLINK formula (1C saves them as values!)
        if (cellValue.includes('=HYPERLINK(')) {
          console.log('📎 Found HYPERLINK in value at', cellAddress + ':', cellValue.substring(0, 100) + '...')
          
          // Parse HYPERLINK formula: =HYPERLINK("url","display_text")
          const match = cellValue.match(/=HYPERLINK\("([^"]+)","([^"]+)"\)/)
          if (match) {
            hyperlinksToCreate.push({
              cellAddress: cellAddress,
              url: match[1],
              displayText: match[2],
              row: R,
              col: C,
              originalValue: cellValue
            })
          }
        }
      }
    }
  }
  
  console.log('🎯 Found', hyperlinksToCreate.length, 'hyperlinks to create')
  
  if (hyperlinksToCreate.length === 0) {
    console.log('❌ No hyperlinks found - copying original file')
    fs.copyFileSync(inputFilePath, outputFilePath)
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

  // 8. FIX BROKEN ATTRIBUTES THAT SPAN MULTIPLE LINES (MS Office fix)
  // Fix broken xmlns attributes that cause MS Office XML errors
  worksheet = worksheet.replace(/xmlns="http:\/\/schemas\.openxmlformats\.org\/spreadsheetml\/\n2006\/main"/g, 
    'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"');
  worksheet = worksheet.replace(/xmlns:r="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/rela\n\s*tionships"/g, 
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"');
  
  // Generic fix for any attribute values that span multiple lines
  worksheet = worksheet.replace(/="([^"]*)\n\s*([^"]*)"/g, '="$1$2"');
  
  // Fix broken XML tags that span multiple lines  
  worksheet = worksheet.replace(/<([^>]*)\n\s*([^>]*)>/g, '<$1 $2>');

  // 9. FORMAT XML PROPERLY FOR MS OFFICE (add proper line breaks)
  // MS Office is very strict about XML formatting and line counting
  
  // Simple but effective approach: add line breaks after every closing tag
  let formattedWorksheet = '';
  let i = 0;
  while (i < worksheet.length) {
    const char = worksheet[i];
    formattedWorksheet += char;
    
    // Add \r\n after every closing tag >
    if (char === '>') {
      // Don't add line break if next chars are already \r\n or \n
      const nextChar = worksheet[i + 1];
      const nextTwoChars = worksheet.substring(i + 1, i + 3);
      
      if (nextChar !== '\r' && nextChar !== '\n' && nextTwoChars !== '\r\n') {
        formattedWorksheet += '\r\n';
      }
    }
    i++;
  }
  
  worksheet = formattedWorksheet;
  
  // Clean up multiple consecutive line breaks
  worksheet = worksheet.replace(/\r\n\r\n+/g, '\r\n');

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

module.exports.doConvertXlsx = async function (inputFilePath) {
  const originalFile = inputFilePath
  const outDir = process.env.XLSX_OUT || path.join(__dirname, '../../temp')
  if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true })
  const outputFileName = path.basename(originalFile)
  const outputFilePath = path.join(outDir, outputFileName)

  return convertToHyperlinks(originalFile, outputFilePath)
}

module.exports.convertToHyperlinks = convertToHyperlinks