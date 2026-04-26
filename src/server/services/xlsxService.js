const fs = require("fs")
const path = require("path")
const XLSX = require("xlsx-js-style")

function convertToHyperlinks(inputFilePath, outputFilePath) {
	console.log("🚀 Starting MS Office compatible conversion...")

	if (!outputFilePath) {
		const outDir = process.env.XLSX_OUT || path.join(__dirname, "../../temp")
		if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true })
		const outputFileName = path.basename(inputFilePath)
		outputFilePath = path.join(outDir, outputFileName)
	}

	const outDir = path.dirname(outputFilePath)
	if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true })

	console.log("📖 Reading file:", inputFilePath)

	const workbook = XLSX.readFile(inputFilePath, {
		cellHTML: false,
		cellNF: false,
		cellStyles: false,
		sheetStubs: false,
		cellFormula: true,
		cellText: false,
	})

	console.log("📋 Found sheets:", workbook.SheetNames)

	const sheetName = workbook.SheetNames[0]
	const sheet = workbook.Sheets[sheetName]

	console.log("📏 Sheet range:", sheet["!ref"])

	const hyperlinksToCreate = []
	const range = XLSX.utils.decode_range(sheet["!ref"])

	console.log("🔍 Scanning for HYPERLINK in cell VALUES (1C format)...")

	for (let R = range.s.r; R <= range.e.r; R++) {
		for (let C = range.s.c; C <= range.e.c; C++) {
			const cellAddress = XLSX.utils.encode_cell({ r: R, c: C })
			const cell = sheet[cellAddress]

			if (cell && cell.v) {
				const cellValue = cell.v.toString()

				if (cellValue.includes("=HYPERLINK(")) {
					console.log(
						"📎 Found HYPERLINK in value at",
						cellAddress + ":",
						cellValue.substring(0, 100) + "...",
					)

					const match = cellValue.match(/=HYPERLINK\("([^"]+)","([^"]+)"\)/)
					if (match) {
						hyperlinksToCreate.push({
							cellAddress: cellAddress,
							url: match[1],
							displayText: match[2],
							row: R,
							col: C,
							originalValue: cellValue,
						})
					}
				}
			}
		}
	}

	console.log("🎯 Found", hyperlinksToCreate.length, "hyperlinks to create")

	console.log("🎨 Scanning cells...")
	const cellsToColor = []
	const cellsVilno = []
	const cellsNpunkt = []
	const cellsGps = []

	for (let R = range.s.r; R <= range.e.r; R++) {
		for (let C = range.s.c; C <= range.e.c; C++) {
			const cellAddress = XLSX.utils.encode_cell({ r: R, c: C })
			const cell = sheet[cellAddress]

			if (
				cell &&
				cell.v &&
				typeof cell.v === "string" &&
				cell.v.trim() === "Продано"
			) {
				cellsToColor.push({ cellAddress, originalValue: cell.v })
			}

			if (
				cell &&
				cell.v &&
				typeof cell.v === "string" &&
				cell.v.trim() === "Вільно"
			) {
				cellsVilno.push({ cellAddress, originalValue: cell.v })
			}

			if (C === 2 || C === 3) {
				cellsNpunkt.push({ cellAddress, cell })
			}

			if (C === 11) {
				cellsGps.push({ cellAddress, cell })
			}
		}
	}

	console.log("🟢 Found", cellsToColor.length, 'cells with "Продано"')
	console.log("⬜ Found", cellsVilno.length, 'cells with "Вільно"')
	console.log("🔵 Found", cellsNpunkt.length, "cells in columns C/D")
	console.log("📍 Found", cellsGps.length, "cells in column L (GPS)")

	if (
		hyperlinksToCreate.length === 0 &&
		cellsToColor.length === 0 &&
		cellsVilno.length === 0
	) {
		console.log(
			"❌ No hyperlinks or colorable cells found - copying original file",
		)
		fs.copyFileSync(inputFilePath, outputFilePath)
		return outputFilePath
	}

	// Convert formulas to actual hyperlinks
	for (const hyperlinkData of hyperlinksToCreate) {
		const { cellAddress, url, displayText } = hyperlinkData
		console.log("🔗 Creating hyperlink:", cellAddress, "->", url)
		sheet[cellAddress] = {
			v: displayText,
			t: "s",
			l: {
				Target: url,
				Tooltip: displayText,
			},
		}
	}

	// Apply green color to "Продано" cells
	for (const colorData of cellsToColor) {
		const { cellAddress, originalValue } = colorData
		console.log("🟢 Coloring cell:", cellAddress, "-> green #a8d2a8")
		const existingCell = sheet[cellAddress] || {}
		const existingStyles = existingCell.s || {}
		sheet[cellAddress] = {
			...existingCell,
			v: originalValue,
			t: typeof originalValue === "number" ? "n" : "s",
			s: {
				...existingStyles,
				fill: { fgColor: { rgb: "A8D2A8" } },
			},
		}
	}

	// Apply grey color to "Вільно" cells
	for (const colorData of cellsVilno) {
		const { cellAddress, originalValue } = colorData
		console.log("⬜ Coloring cell:", cellAddress, "-> grey #ebebeb")
		const existingCell = sheet[cellAddress] || {}
		const existingStyles = existingCell.s || {}
		sheet[cellAddress] = {
			...existingCell,
			v: originalValue,
			t: "s",
			s: {
				...existingStyles,
				fill: { fgColor: { rgb: "EBEBEB" } },
			},
		}
	}

	// Apply light blue + wrap to columns C and D
	for (const npunkt of cellsNpunkt) {
		const { cellAddress, cell } = npunkt
		console.log(
			"🔵 Coloring cell:",
			cellAddress,
			"-> light blue #ccffff + wrap",
		)
		const existingCell = cell || {}
		const existingStyles = existingCell.s || {}
		sheet[cellAddress] = {
			...existingCell,
			v: existingCell.v !== undefined ? existingCell.v : "",
			t: existingCell.t || "s",
			s: {
				...existingStyles,
				fill: { fgColor: { rgb: "CCFFFF" } },
				alignment: {
					...(existingStyles.alignment || {}),
					wrapText: true,
				},
			},
		}
	}

	// Apply wrap only to column L (GPS)
	for (const gps of cellsGps) {
		const { cellAddress, cell } = gps
		console.log("📍 Wrapping cell:", cellAddress, "-> wrapText")
		const existingCell = cell || {}
		const existingStyles = existingCell.s || {}
		sheet[cellAddress] = {
			...existingCell,
			v: existingCell.v !== undefined ? existingCell.v : "",
			t: existingCell.t || "s",
			s: {
				...existingStyles,
				alignment: {
					...(existingStyles.alignment || {}),
					wrapText: true,
				},
			},
		}
	}

	// Set column widths — fix column L (GPS, index 11) to fit "50.3883571," + 3 chars
	const colsCount = range.e.c + 1
	const existingCols = sheet["!cols"] || []
	const cols = Array.from(
		{ length: colsCount },
		(_, i) => existingCols[i] || {},
	)
	cols[11] = { ...cols[11], wch: 15 }
	sheet["!cols"] = cols
	console.log("📐 Set column L width to 15 chars")

	console.log("💾 Writing MS Office compatible file:", outputFilePath)

	XLSX.writeFile(workbook, outputFilePath, {
		bookType: "xlsx",
		cellStyles: true,
		type: "buffer",
		bookSST: true,
		compression: false,
	})

	console.log("✅ MS Office compatible conversion complete!")

	return outputFilePath
}

// Async wrapper for backward compatibility
module.exports.doConvertXlsx = async function (inputFilePath) {
	const originalFile = inputFilePath
	const outDir = process.env.XLSX_OUT || path.join(__dirname, "../../temp")
	if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true })
	const outputFileName = path.basename(originalFile)
	const outputFilePath = path.join(outDir, outputFileName)

	return convertToHyperlinks(originalFile, outputFilePath)
}

module.exports.convertToHyperlinks = convertToHyperlinks
module.exports.convertToHyperlinksV2 = convertToHyperlinks
