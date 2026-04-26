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

	// ─── Helper: is this cell value a pure group number? (1, 2, 99, etc.) ───
	function isGroupNumber(val) {
		if (val === undefined || val === null || val === "") return false
		const s = val.toString().trim()
		return /^\d+$/.test(s) // only digits, no letters
	}

	// ─── First pass: identify group header rows and data row ranges ───────────
	console.log("🔍 Detecting group structure...")

	// We look at column A (index 0) — merged A-B means value lives in A
	const groupHeaders = [] // { row, groupNum }

	for (let R = range.s.r; R <= range.e.r; R++) {
		const cellA = sheet[XLSX.utils.encode_cell({ r: R, c: 0 })]
		const valA = cellA ? cellA.v : undefined
		if (isGroupNumber(valA)) {
			groupHeaders.push({ row: R, groupNum: Number(valA) })
		}
	}

	console.log(
		"📂 Found",
		groupHeaders.length,
		"groups:",
		groupHeaders.map((g) => g.groupNum).join(", "),
	)

	// Build groups: each group owns rows from (headerRow+1) to (nextHeaderRow-1)
	const groups = groupHeaders.map((g, i) => ({
		headerRow: g.row,
		groupNum: g.groupNum,
		dataStart: g.row + 1,
		dataEnd:
			i + 1 < groupHeaders.length ? groupHeaders[i + 1].row - 1 : range.e.r,
	}))

	// ─── Scan cells for coloring ──────────────────────────────────────────────
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
		console.log("🟢 Coloring cell:", cellAddress, "-> green #a8d2a8")
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
		console.log("⬜ Coloring cell:", cellAddress, "-> grey #ebebeb")
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
		console.log(
			"🔵 Coloring cell:",
			cellAddress,
			"-> light blue #ccffff + wrap",
		)
		sheet[cellAddress] = {
			...cell,
			s: {
				...(cell.s || {}),
				fill: { fgColor: { rgb: "CCFFFF" } },
				alignment: { ...(cell.s?.alignment || {}), wrapText: true },
			},
		}
	}

	// ─── Wrap only for column L (GPS) ────────────────────────────────────────
	for (const { cellAddress, cell } of cellsGps) {
		if (!cell) continue
		console.log("📍 Wrapping cell:", cellAddress, "-> wrapText")
		sheet[cellAddress] = {
			...cell,
			s: {
				...(cell.s || {}),
				alignment: { ...(cell.s?.alignment || {}), wrapText: true },
			},
		}
	}

	// ─── Color group header rows (C-L) with #f2e297 ──────────────────────────
	console.log("🟡 Applying group header colors...")
	for (const group of groups) {
		const R = group.headerRow
		console.log(
			"🟡 Group",
			group.groupNum,
			"- coloring header row",
			R + 1,
			"(columns C-L)",
		)

		// Color columns C through L (indices 2–11) in the header row
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

	// ─── Set Excel row groupings (outline) ───────────────────────────────────
	// Excel groupings live in sheet['!rows'] as { level, hidden } per row index
	// Header row = outside group (level 0), data rows = level 1, expanded
	console.log("📊 Setting Excel row groupings...")

	const totalRows = range.e.r + 1
	const rowsMeta = sheet["!rows"] ? [...sheet["!rows"]] : []

	// Ensure array is long enough
	while (rowsMeta.length < totalRows) rowsMeta.push(undefined)

	for (const group of groups) {
		console.log(
			"📊 Group",
			group.groupNum,
			"- data rows",
			group.dataStart + 1,
			"to",
			group.dataEnd + 1,
			"(Excel rows",
			group.dataStart + 1,
			"-",
			group.dataEnd + 1,
			")",
		)
		for (let R = group.dataStart; R <= group.dataEnd; R++) {
			rowsMeta[R] = {
				...(rowsMeta[R] || {}),
				level: 1, // outline level 1
				hidden: false, // expanded (not collapsed)
			}
		}
	}

	sheet["!rows"] = rowsMeta

	// ─── Column widths ────────────────────────────────────────────────────────
	const colsCount = range.e.c + 1
	const existingCols = sheet["!cols"] || []
	const cols = Array.from(
		{ length: colsCount },
		(_, i) => existingCols[i] || {},
	)
	cols[11] = { ...cols[11], wch: 15 }
	sheet["!cols"] = cols
	console.log("📐 Set column L width to 15 chars")

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
