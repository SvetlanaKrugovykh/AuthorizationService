const fs = require('fs')
const path = require('path')
const XLSX = require('xlsx')

/**
 * NEW CLEAN APPROACH using XLSX library for MS Office compatibility
 * Converts HYPERLINK formulas to real clickable hyperlinks
 */
function convertToHyperlinksV2(inputFilePath, outputFilePath) {
  console.log('🚀 Starting V2 conversion with XLSX library...')
  
  if (!outputFilePath) {
    const outDir = process.env.XLSX_OUT || path.join(__dirname, '../../temp')
    if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true })
    const outputFileName = path.basename(inputFilePath, '.xlsx') + '_V2.xlsx'
    outputFilePath = path.join(outDir, outputFileName)
  }

  // Ensure output directory exists
  const outDir = path.dirname(outputFilePath)
  if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true })

  console.log('📖 Reading file:', inputFilePath)
  
  // Step 1: Read the workbook with all options for hyperlinks
  const workbook = XLSX.readFile(inputFilePath, {
    cellHTML: false,
    cellNF: false,
    cellStyles: false,
    sheetStubs: false,
    cellFormula: true,  // Important: keep formulas
    cellText: false
  })

  console.log('📋 Found sheets:', workbook.SheetNames)
  
  const sheetName = workbook.SheetNames[0]  // Usually 'TDSheet'
  const sheet = workbook.Sheets[sheetName]
  
  console.log('📏 Sheet range:', sheet['!ref'])

  // Step 2: Find all cells that need hyperlink conversion
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

  // Step 3: Convert formulas to actual hyperlinks
  for (const hyperlinkData of hyperlinksToCreate) {
    const { cellAddress, url, displayText } = hyperlinkData
    
    console.log('🔗 Creating hyperlink:', cellAddress, '->', url)
    
    // Replace the formula with the display text
    sheet[cellAddress] = {
      v: displayText,   // Display value
      t: 's',           // String type
      l: {              // Link object
        Target: url,
        Tooltip: displayText
      }
    }
  }

  // Step 4: Write the new file using XLSX library (MS Office compatible)
  console.log('💾 Writing MS Office compatible file:', outputFilePath)
  
  XLSX.writeFile(workbook, outputFilePath, {
    bookType: 'xlsx',
    type: 'binary',
    cellStyles: true,
    bookSST: true,  // Use shared strings table
    compression: true
  })
  
  console.log('✅ V2 conversion complete!')
  return outputFilePath
}

module.exports = {
  convertToHyperlinksV2
}