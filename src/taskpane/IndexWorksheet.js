// Helper function to build the pipe-delimited string for a row
function buildRowArrayString(valuesRow, formulasRow) {
    // ... (Helper function remains the same)
    // Indices: A=0, B=1, C=2, G=6, H=7, I=8, K=10, L=11, M=12, N=13, O=14, P=15, R=17
    const cols = [0, 1, 2, 6, 7, 8, 10, 11, 12, 13, 14, 15, 17];
    const formulaCols = [10, 11, 12, 13, 14, 15, 17]; // Columns to check for formulas
    let parts = [];

    if (!valuesRow) return ""; // Handle case where row might be completely empty

    for (const colIndex of cols) {
        // Ensure the row has enough columns
        let value = (valuesRow.length > colIndex && valuesRow[colIndex] !== null && valuesRow[colIndex] !== undefined) ? valuesRow[colIndex] : "";
        let formula = (formulasRow && formulasRow.length > colIndex) ? formulasRow[colIndex] : null;

        if (formulaCols.includes(colIndex)) {
            // Check formula first
            if (typeof formula === 'string' && formula.startsWith('=')) {
                value = "F";
            }
        }
        const valueString = String(value).replace(/"/g, '""');
        parts.push(valueString);
    }
    return parts.join("|");
}

// Helper function to process a standard code block based on Column D
function processCodeBlock(values, formulas, startDataRowIndex, endDataRowIndex, codeName, rangeStartRowExcel) {
    // console.log(`Processing block: ${codeName} from Excel row ${startDataRowIndex + rangeStartRowExcel} to ${endDataRowIndex + rangeStartRowExcel}`);
    let blockString = `<${codeName};`;
    let rowCount = 1;
    // The indices i here are 0-based relative to the START of the loaded data range
    for (let i = startDataRowIndex; i <= endDataRowIndex; i++) {
        if (values && i < values.length && formulas && i < formulas.length) {
            const rowArrayString = buildRowArrayString(values[i], formulas[i]);
            blockString += `row${rowCount}="${rowArrayString}";`;
            rowCount++;
        } else {
             // console.warn(`Skipping data row index ${i} in block ${codeName} due to missing data.`);
        }
    }
    blockString += ">";
    // console.log(`Generated block string for ${codeName}: ${blockString.substring(0,100)}...`);
    return blockString;
}

export async function generateTabString() {
    try {
        await Excel.run(async (context) => {
            console.log("Starting generateTabString (dynamic range)...");
            const sheets = context.workbook.worksheets;
            sheets.load("items/name");
            await context.sync();
            console.log(`Found ${sheets.items.length} sheets.`);

            let finalResultString = "";

            for (const sheet of sheets.items) {
                console.log(`Processing sheet: ${sheet.name}`);
                let sheetCodeBlocks = "";
                let lastUsedRowInB = 0;

                try {
                    // --- Find the last used row in Column B ---
                    // Get the bottom-most cell in column B (using a large row number)
                    const bottomCellInB = sheet.getRange("B1048576");
                    // Find the last used cell *up* from the absolute bottom
                    const lastUsedCellInB = bottomCellInB.getRangeEdge(Excel.KeyboardDirection.up);
                    lastUsedCellInB.load("rowIndex");
                    await context.sync();

                    // Add 1 because rowIndex is 0-based
                    lastUsedRowInB = lastUsedCellInB.rowIndex + 1;
                    console.log(`Sheet ${sheet.name}: Last used row in Col B found at: ${lastUsedRowInB}`);

                    // --- Validate last row and define dynamic range ---
                    if (lastUsedRowInB < 9) {
                        console.log(`Sheet ${sheet.name}: Last used row (${lastUsedRowInB}) is before row 9. Skipping code generation.`);
                        finalResultString += `<TAB; label1="${sheet.name}";>\n\n`; // Add TAB tag even if no codes
                        continue; // Skip to the next sheet
                    }

                    const dynamicRangeAddress = `A9:R${lastUsedRowInB}`;
                    console.log(`Sheet ${sheet.name}: Loading dynamic range: ${dynamicRangeAddress}`);

                    // --- Load data from the dynamic range --- Load data from the dynamic range ---
                    const range = sheet.getRange(dynamicRangeAddress);
                    range.load(["values", "formulas", "rowCount", "rowIndex"]);
                    await context.sync();

                    const values = range.values;
                    const formulas = range.formulas;
                    const loadedRowCount = range.rowCount;
                    const rangeStartRowExcel = range.rowIndex + 1; // Should be 9

                    if (rangeStartRowExcel !== 9) {
                         console.warn(`Sheet ${sheet.name}: Loaded range started at ${rangeStartRowExcel} instead of 9.`);
                    }
                    if (!values || loadedRowCount === 0) {
                        console.log(`Sheet ${sheet.name}: No values loaded from range ${dynamicRangeAddress}. Skipping code generation.`);
                        finalResultString += `<TAB; label1="${sheet.name}";>\n\n`;
                        continue; // Skip to the next sheet
                    }

                    // --- Process the loaded data --- Process the loaded data ---
                    // Check if Column D (index 3 in the loaded array) has significant data
                    let hasDataInD = false;
                    for (let r = 0; r < loadedRowCount; r++) {
                         if (values[r] && values[r].length > 3 && values[r][3] && String(values[r][3]).trim() !== "") {
                            hasDataInD = true;
                            break;
                        }
                    }
                    console.log(`Sheet ${sheet.name}: Has significant data in Column D (in loaded range)? ${hasDataInD}`);

                    if (!hasDataInD) {
                        // --- Generate MANUAL-ER Block ---
                        console.log(`Sheet ${sheet.name}: Generating MANUAL-ER block.`);
                        sheetCodeBlocks += "<MANUAL-ER;";
                        let manualRowCount = 1;
                        for (let r = 0; r < loadedRowCount; r++) {
                            if (values[r] && formulas[r]) { // Ensure row data exists
                                const rowArrayString = buildRowArrayString(values[r], formulas[r]);
                                sheetCodeBlocks += `row${manualRowCount}="${rowArrayString}";`;
                                manualRowCount++;
                            }
                        }
                        sheetCodeBlocks += ">";

                    } else {
                        // --- Generate Blocks Based on Column D ---
                        console.log(`Sheet ${sheet.name}: Generating blocks based on Column D changes.`);
                        let currentBlockStartDataRow = -1;
                        let currentBlockCodeName = "";

                        for (let r = 0; r < loadedRowCount; r++) {
                            const dValue = (values[r] && values[r].length > 3 && values[r][3]) ? String(values[r][3]).trim() : "";

                            if (dValue !== "") {
                                if (currentBlockStartDataRow === -1) {
                                    currentBlockStartDataRow = r;
                                    currentBlockCodeName = dValue;
                                } else if (dValue !== currentBlockCodeName) {
                                    // Add newline before appending if sheetCodeBlocks isn't empty
                                    if (sheetCodeBlocks) sheetCodeBlocks += "\n";
                                    sheetCodeBlocks += processCodeBlock(values, formulas, currentBlockStartDataRow, r - 1, currentBlockCodeName, rangeStartRowExcel);
                                    currentBlockStartDataRow = r;
                                    currentBlockCodeName = dValue;
                                }
                            } else { // Empty D value
                                if (currentBlockStartDataRow !== -1) {
                                    // Add newline before appending if sheetCodeBlocks isn't empty
                                    if (sheetCodeBlocks) sheetCodeBlocks += "\n";
                                    sheetCodeBlocks += processCodeBlock(values, formulas, currentBlockStartDataRow, r - 1, currentBlockCodeName, rangeStartRowExcel);
                                    currentBlockStartDataRow = -1;
                                    currentBlockCodeName = "";
                                }
                            }

                            // Handle the last block if the loop finishes while inside a block
                            if (r === loadedRowCount - 1 && currentBlockStartDataRow !== -1) {
                                // Add newline before appending if sheetCodeBlocks isn't empty
                                if (sheetCodeBlocks) sheetCodeBlocks += "\n";
                                sheetCodeBlocks += processCodeBlock(values, formulas, currentBlockStartDataRow, r, currentBlockCodeName, rangeStartRowExcel);
                            }
                        } // end for loop r
                    } // end else (hasDataInD)

                } catch (sheetError) {
                     console.error(`Error processing sheet ${sheet.name}: ${sheetError}`);
                     if (sheetError instanceof OfficeExtension.Error) {
                         console.error("Debug info: " + JSON.stringify(sheetError.debugInfo));
                     }
                     sheetCodeBlocks = "<!-- Error processing sheet data -->";
                }

                // Append the result for this sheet
                // Add a newline between TAB and blocks only if blocks exist
                let tabLine = `<TAB; label1="${sheet.name}";>`;
                if (sheetCodeBlocks) {
                    tabLine += "\n" + sheetCodeBlocks; // Add newline before the blocks
                }
                finalResultString += tabLine + "\n\n"; // Add double newline after each sheet entry
            } // end loop sheets

            console.log("--- FINAL GENERATED STRING ---");
            console.log(finalResultString);
            console.log("--- END FINAL GENERATED STRING ---");

        }); // end Excel.run
    } catch (error) {
        console.error("Error in generateTabString top level: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.error("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }
}
