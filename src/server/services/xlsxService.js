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
							cellAddress,
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

	// ─── Helper: is this cell value a pure group number? ─────────────────────
	function isGroupNumber(val) {
		if (val === undefined || val === null || val === "") return false
		return /^\d+$/.test(val.toString().trim())
	}

	// ─── Detect group structure ───────────────────────────────────────────────
	console.log("🔍 Detecting group structure...")
	const groupHeaders = []

	for (let R = range.s.r; R <= range.e.r; R++) {
		const cellA = sheet[XLSX.utils.encode_cell({ r: R, c: 0 })]
		if (isGroupNumber(cellA ? cellA.v : undefined)) {
			groupHeaders.push({ row: R, groupNum: Number(cellA.v) })
		}
	}

	console.log(
		"📂 Found",
		groupHeaders.length,
		"groups:",
		groupHeaders.map((g) => g.groupNum).join(", "),
	)

	const groups = groupHeaders.map((g, i) => ({
		headerRow: g.row,
		groupNum: g.groupNum,
		dataStart: g.row + 1,
		dataEnd:
			i + 1 < groupHeaders.length ? groupHeaders[i + 1].row - 1 : range.e.r,
	}))

	// ─── Scan cells ───────────────────────────────────────────────────────────
	console.log("🎨 Scanning cells...")
	const cellsToColor = []
	const cellsVilno = []
	const cellsNpunkt = [] // C, D  — col index 2, 3
	const cellsGps = [] // L     — col index 11
	const cellsG = [] // G     — col index 6
	const cellsI = [] // I     — col index 8

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
			if (C === 2 || C === 3) cellsNpunkt.push({ cellAddress, cell })
			if (C === 11) cellsGps.push({ cellAddress, cell })
			if (C === 6) cellsG.push({ cellAddress, cell })
			if (C === 8) cellsI.push({ cellAddress, cell })
		}
	}

	console.log("🟢 Found", cellsToColor.length, 'cells with "Продано"')
	console.log("⬜ Found", cellsVilno.length, 'cells with "Вільно"')
	console.log("🔵 Found", cellsNpunkt.length, "cells in columns C/D")
	console.log("📍 Found", cellsGps.length, "cells in column L (GPS)")
	console.log("🔷 Found", cellsG.length, "cells in column G")
	console.log("🔷 Found", cellsI.length, "cells in column I")

	if (
		hyperlinksToCreate.length === 0 &&
		cellsToColor.length === 0 &&
		cellsVilno.length === 0 &&
		groups.length === 0
	) {
		console.log("❌ Nothing to process - copying original file")
		fs.copyFileSync(inputFilePath, outputFilePath)
		return outputFilePath
	}

	// ─── Convert hyperlinks ───────────────────────────────────────────────────
	for (const { cellAddress, url, displayText } of hyperlinksToCreate) {
		console.log("🔗 Creating hyperlink:", cellAddress, "->", url)
		sheet[cellAddress] = {
			v: displayText,
			t: "s",
			l: { Target: url, Tooltip: displayText },
		}
	}

	// ─── Color "Продано" green ────────────────────────────────────────────────
	for (const { cellAddress, originalValue } of cellsToColor) {
		const existingCell = sheet[cellAddress] || {}
		sheet[cellAddress] = {
			...existingCell,
			v: originalValue,
			t: typeof originalValue === "number" ? "n" : "s",
			s: {
				...(existingCell.s || {}),
				fill: { fgColor: { rgb: "A8D2A8" } },
			},
		}
	}

	// ─── Color "Вільно" grey ──────────────────────────────────────────────────
	for (const { cellAddress, originalValue } of cellsVilno) {
		const existingCell = sheet[cellAddress] || {}
		sheet[cellAddress] = {
			...existingCell,
			v: originalValue,
			t: "s",
			s: {
				...(existingCell.s || {}),
				fill: { fgColor: { rgb: "EBEBEB" } },
			},
		}
	}

	// ─── Light blue + wrap for columns C/D ───────────────────────────────────
	for (const { cellAddress, cell } of cellsNpunkt) {
		if (!cell) continue
		sheet[cellAddress] = {
			...cell,
			s: {
				...(cell.s || {}),
				fill: { fgColor: { rgb: "CCFFFF" } },
				alignment: { ...(cell.s?.alignment || {}), wrapText: true },
			},
		}
	}

	// ─── Column G: blue color + bold ─────────────────────────────────────────
	for (const { cellAddress, cell } of cellsG) {
		if (!cell) continue
		console.log("🔷 Styling cell:", cellAddress, "-> blue #0000ea + bold")
		sheet[cellAddress] = {
			...cell,
			s: {
				...(cell.s || {}),
				font: {
					...(cell.s?.font || {}),
					bold: true,
					color: { rgb: "0000EA" },
				},
			},
		}
	}

	// ─── Column I: blue color ─────────────────────────────────────────────────
	for (const { cellAddress, cell } of cellsI) {
		if (!cell) continue
		console.log("🔷 Styling cell:", cellAddress, "-> blue #0000ea")
		sheet[cellAddress] = {
			...cell,
			s: {
				...(cell.s || {}),
				font: {
					...(cell.s?.font || {}),
					color: { rgb: "0000EA" },
				},
			},
		}
	}

	// ─── Column L: blue color + wrap + wrapText rows capped at 2 ─────────────
	for (const { cellAddress, cell } of cellsGps) {
		if (!cell) continue
		console.log("📍 Styling cell:", cellAddress, "-> blue #0000ea + wrap")
		sheet[cellAddress] = {
			...cell,
			s: {
				...(cell.s || {}),
				font: {
					...(cell.s?.font || {}),
					color: { rgb: "0000EA" },
				},
				alignment: {
					...(cell.s?.alignment || {}),
					wrapText: true,
				},
			},
		}
	}

	// ─── Color group header rows C-L with #f2e297 ────────────────────────────
	console.log("🟡 Applying group header colors...")
	for (const group of groups) {
		const R = group.headerRow
		console.log("🟡 Group", group.groupNum, "- coloring header row", R + 1)
		for (let C = 2; C <= 11; C++) {
			const cellAddress = XLSX.utils.encode_cell({ r: R, c: C })
			const existingCell = sheet[cellAddress] || {}
			sheet[cellAddress] = {
				...existingCell,
				v: existingCell.v !== undefined ? existingCell.v : "",
				t: existingCell.t || "s",
				s: {
					...(existingCell.s || {}),
					fill: { fgColor: { rgb: "F2E297" } },
				},
			}
		}
	}

	// ─── Excel row groupings ──────────────────────────────────────────────────
	console.log("📊 Setting Excel row groupings...")
	const totalRows = range.e.r + 1
	const rowsMeta = sheet["!rows"] ? [...sheet["!rows"]] : []
	while (rowsMeta.length < totalRows) rowsMeta.push(undefined)

	for (const group of groups) {
		for (let R = group.dataStart; R <= group.dataEnd; R++) {
			rowsMeta[R] = { ...(rowsMeta[R] || {}), level: 1, hidden: false }
		}
	}
	sheet["!rows"] = rowsMeta

	// ─── Column widths ────────────────────────────────────────────────────────
	// A=0, B=1, C=2, D=3, G=6, I=8, L=11
	const colsCount = range.e.c + 1
	const existingCols = sheet["!cols"] || []
	const cols = Array.from(
		{ length: colsCount },
		(_, i) => existingCols[i] || {},
	)

	cols[0] = { ...cols[0], wch: 8 } // A — max 8 chars (merged A-B)
	cols[1] = { ...cols[1], wch: 8 } // B — max 8 chars (merged A-B)
	cols[6] = { ...cols[6], wch: 1 } // G — 1 char wide
	cols[8] = { ...cols[8], wch: 8 } // I — max 8 chars
	cols[11] = { ...cols[11], wch: 10 } // L — exactly 10 chars

	sheet["!cols"] = cols
	console.log("📐 Column widths set: A/B=8, G=1, I=8, L=10")

	// ─── Row heights: cap GPS column (L) to 2-line height ────────────────────
	// Standard row height ~15pt, 2 lines = ~30pt
	// We set it on ALL data rows so wrap never exceeds 2 visible lines
	const ROW_HEIGHT_2_LINES = 30 // points
	for (const group of groups) {
		for (let R = group.dataStart; R <= group.dataEnd; R++) {
			rowsMeta[R] = {
				...(rowsMeta[R] || {}),
				hpt: ROW_HEIGHT_2_LINES, // height in points
				hpx: ROW_HEIGHT_2_LINES, // height in pixels (same value works)
			}
		}
	}
	sheet["!rows"] = rowsMeta
	console.log("📐 Row heights set to 2-line max (30pt) for all data rows")

	// ─── Write file ───────────────────────────────────────────────────────────
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
