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

  // Convert formulas to actual hyperlinks using XLSX library
  for (const hyperlinkData of hyperlinksToCreate) {
    const { cellAddress, url, displayText } = hyperlinkData
    
    console.log('🔗 Creating hyperlink:', cellAddress, '->', url)
    
    // Replace the formula with the display text and hyperlink
    sheet[cellAddress] = {
      v: displayText,   // Display value
      t: 's',           // String type
      l: {              // Link object - MS Office compatible
        Target: url,
        Tooltip: displayText
      }
    }
  }

  // Write the new file using XLSX library (guaranteed MS Office compatibility)
  console.log('💾 Writing MS Office compatible file:', outputFilePath)
  
  XLSX.writeFile(workbook, outputFilePath, {
    bookType: 'xlsx',
    type: 'binary',
    cellStyles: true,
    bookSST: true,  // Use shared strings table
    compression: true
  })
  
  console.log('✅ MS Office compatible conversion complete!')

  return outputFilePath
}

// Async wrapper for backward compatibility
module.exports.doConvertXlsx = async function (inputFilePath) {
  const originalFile = inputFilePath
  const outDir = process.env.XLSX_OUT || path.join(__dirname, '../../temp')
  if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true })
  const outputFileName = path.basename(originalFile)
  const outputFilePath = path.join(outDir, outputFileName)

  return convertToHyperlinks(originalFile, outputFilePath)
}

// Main export - MS Office compatible version
module.exports.convertToHyperlinks = convertToHyperlinks

// Additional exports for compatibility
module.exports.convertToHyperlinksV2 = convertToHyperlinks