/**
 * CodeCollection.js
 * Functions for processing and managing code collections
 */

import { convertKeysToCamelCase } from "@pinecone-database/pinecone/dist/utils";

/**
 * Parses code strings and creates a code collection
 * @param {string} inputText - The input text containing code strings
 * @returns {Array} - An array of code objects with type and parameters
 */
export function populateCodeCollection(inputText) {
    try {
        console.log("Processing input text for code collection");
        
        // Initialize an empty code collection
        const codeCollection = [];
        
        // Split the input text by newlines to process each line
        const lines = inputText.split('\n');
        
        for (const line of lines) {
            // Skip empty lines
            if (!line.trim()) continue;
            
            // Extract the code type and parameters
            const codeMatch = line.match(/<([^;>]+);(.*?)>/);
            if (!codeMatch) continue;
            
            const codeType = codeMatch[1].trim();
            const paramsString = codeMatch[2].trim();
            
            // Parse parameters
            const params = {};
            
            // Handle special case for row parameters with asterisks
            const rowMatches = paramsString.matchAll(/row(\d+)\s*=\s*"([^"]*)"/g);
            for (const match of rowMatches) {
                const rowNum = match[1];
                const rowValue = match[2];
                params[`row${rowNum}`] = rowValue;
            }
            
            // Parse other parameters
            const paramMatches = paramsString.matchAll(/(\w+)\s*=\s*"([^"]*)"/g);
            for (const match of paramMatches) {
                const paramName = match[1].trim();
                const paramValue = match[2].trim();
                
                // Skip row parameters as they're already handled
                if (paramName.startsWith('row')) continue;
                
                params[paramName] = paramValue;
            }
            
            // Add the code to the collection
            codeCollection.push({
                type: codeType,
                params: params
            });
        }
        
        console.log(`Processed ${codeCollection.length} codes`);
        return codeCollection;
    } catch (error) {
        console.error("Error in populateCodeCollection:", error);
        throw error;
    }
}

/**
 * Exports a code collection to text format
 * @param {Array} codeCollection - The code collection to export
 * @returns {string} - A formatted text representation of the code collection
 */
export function exportCodeCollectionToText(codeCollection) {
    try {
        if (!codeCollection || !Array.isArray(codeCollection)) {
            throw new Error("Invalid code collection");
        }
        
        let result = "Code Collection:\n";
        result += "================\n\n";
        
        codeCollection.forEach((code, index) => {
            result += `Code ${index + 1}: ${code.type}\n`;
            result += "Parameters:\n";
            
            // First display non-row parameters
            for (const [key, value] of Object.entries(code.params)) {
                if (!key.startsWith('row')) {
                    result += `  ${key}: ${value}\n`;
                }
            }
            
            // Then display row parameters
            const rowParams = Object.entries(code.params)
                .filter(([key]) => key.startsWith('row'))
                .sort((a, b) => {
                    const numA = parseInt(a[0].replace('row', ''));
                    const numB = parseInt(b[0].replace('row', ''));
                    return numA - numB;
                });
            
            if (rowParams.length > 0) {
                result += "  Rows:\n";
                for (const [key, value] of rowParams) {
                    result += `    ${key}: ${value}\n`;
                }
            }
            
            result += "\n";
        });
        
        return result;
    } catch (error) {
        console.error("Error in exportCodeCollectionToText:", error);
        throw error;
    }
} 

/**
 * Processes a code collection and performs operations based on code types
 * @param {Array} codeCollection - The code collection to process
 * @returns {Object} - Results of processing the code collection
 */
export async function runCodes(codeCollection) {
    try {
        console.log("Running code collection processing");
        
        if (!codeCollection || !Array.isArray(codeCollection)) {
            throw new Error("Invalid code collection");
        }
        
        // Initialize result object
        const result = {
            processedCodes: 0,
            createdTabs: [],
            errors: []
        };
        
        // Initialize state variables (similar to VBA variables)
        let currentWorksheetName = null;
        const assumptionTabs = [];
        
        // Process each code in the collection
        for (let i = 0; i < codeCollection.length; i++) {
            const code = codeCollection[i];
            const codeType = code.type;
            
            try {
                // Handle MODEL code type
                if (codeType === "MODEL") {
                    // Skip for now as mentioned in the original VBA code
                    console.log("MODEL code type encountered - skipping for now");
                    continue;
                }
                
                // Handle TAB code type
                if (codeType === "TAB") {
                    // Accept both label1 and Label1 for backward compatibility
                    const tabName = code.params.label1 || code.params.Label1 || `Tab_${i}`;
                    
                    // Check if worksheet exists and delete it
                    await Excel.run(async (context) => {
                        try {
                            // Get all worksheets
                            const sheets = context.workbook.worksheets;
                            sheets.load("items/name");
                            console.log("sheets", sheets);
                            await context.sync();
                            
                            // Check if worksheet exists
                            const existingSheet = sheets.items.find(sheet => sheet.name === tabName);
                            console.log("existingSheet", existingSheet);
                            if (existingSheet) {
                                // Delete the worksheet if it exists
                                existingSheet.delete();
                                await context.sync();
                            }
                            console.log("existingSheet deleted");
                            
                            // Get the Calcs worksheet
                            const sourceCalcsWS = context.workbook.worksheets.getItem("Calcs");
                            console.log("sourceCalcsWS", sourceCalcsWS);
                            
                            // Create a new worksheet by copying the Calcs worksheet
                            const newSheet = sourceCalcsWS.copy();
                            console.log("newSheet created by copying Calcs worksheet");
                            
                            // Rename it
                            newSheet.name = tabName;
                            console.log("newSheet renamed to", tabName);
                            
                            // Set the first row
                            const firstRow = 9; // Equivalent to calcsfirstrow in VBA
                            console.log("firstRow", firstRow);
                            
                            // Clear all cells including and below the first row
                            const usedRange = newSheet.getUsedRange();
                            usedRange.load("rowCount");
                            await context.sync();
                            
                            if (usedRange.rowCount >= firstRow) {
                                const clearRange = newSheet.getRange(`${firstRow}:${usedRange.rowCount}`);
                                clearRange.clear();
                                console.log(`Cleared all cells from row ${firstRow} to ${usedRange.rowCount}`);
                            }
                            
                            // Add to assumption tabs collection
                            assumptionTabs.push({
                                name: tabName,
                                worksheet: newSheet
                            });
                            
                            // Set the current worksheet name
                            currentWorksheetName = tabName;
                            
                            await context.sync();
                            
                            result.createdTabs.push(tabName);
                            console.log("Tab created successfully:", tabName);
                        } catch (error) {
                            console.error("Detailed error in TAB processing:", error);
                            throw error;
                        }
                    }).catch(error => {
                        console.error(`Error processing TAB code: ${error.message}`);
                        result.errors.push({
                            codeIndex: i,
                            codeType: codeType,
                            error: error.message
                        });
                    });
                    
                    continue;
                }
                
                // Handle non-TAB codes
                if (codeType !== "TAB") {
                    await Excel.run(async (context) => {
                        try {
                            // Get the Codes worksheet
                            const codesWS = context.workbook.worksheets.getItem("Codes");
                            console.log("Got Codes worksheet");
                            
                            // Get the used range of the Codes worksheet
                            const usedRange = codesWS.getUsedRange();
                            usedRange.load("rowCount");
                            usedRange.load("columnCount");
                            await context.sync();
                            console.log(`Used range: ${usedRange.rowCount} rows x ${usedRange.columnCount} columns`);
                            
                            // Get the current worksheet
                            const currentWS = context.workbook.worksheets.getItem(currentWorksheetName);
                            console.log("Got current worksheet:", currentWorksheetName);
                            
                            // Get the last row in the current worksheet
                            const lastUsedRow = currentWS.getUsedRange().getLastRow();
                            lastUsedRow.load("rowIndex");
                            await context.sync();
                            const pasteRow = lastUsedRow.rowIndex + 2; // Adjusted to paste one row lower
                            console.log("Paste row:", pasteRow);
                            
                            // Search for the code type in column D (index 3)
                            let firstRow = -1;
                            let lastRow = -1;
                            
                            // Load the values of column D
                            const columnD = codesWS.getRange(`D1:D${usedRange.rowCount}`);
                            columnD.load("values");
                            await context.sync();
                            
                            console.log("Loaded column D values");
                            
                            // Check if values are loaded properly
                            if (!columnD.values) {
                                console.error("columnD.values is null or undefined");
                                throw new Error(`Failed to load values from column D in Codes worksheet`);
                            }
                            
                            console.log(`columnD.values length: ${columnD.values.length}`);
                            
                            // Debug print the first few values in column D
                            console.log("First 10 values in column D:");
                            for (let i = 0; i < Math.min(10, columnD.values.length); i++) {
                                console.log(`Row ${i+1}: ${columnD.values[i][0]}`);
                            }
                            
                            // Find the first and last row with the code
                            for (let row = 0; row < columnD.values.length; row++) {
                                if (columnD.values[row][0] === codeType) {
                                    if (firstRow === -1) {
                                        firstRow = row + 1; // Excel rows are 1-indexed
                                    }
                                    lastRow = row + 1;
                                }
                            }
                            
                            // Check if the code type was found
                            const codeTypeFound = firstRow !== -1 && lastRow !== -1;
                            
                            if (!codeTypeFound) {
                                console.warn(`Code type ${codeType} not found in Codes worksheet. Skipping this code.`);
                                result.errors.push({
                                    codeIndex: i,
                                    codeType: codeType,
                                    error: `Code type ${codeType} not found in Codes worksheet`
                                });
                                // Skip to the next code
                                result.processedCodes++;
                            } else {
                                console.log(`Found code type ${codeType} in rows ${firstRow} to ${lastRow}`);
                                
                                // Try the suggested approach to copy the range with all properties
                                await Excel.run(async (context) => {
                                    // Get the source range
                                    const sourceRange = context.workbook.worksheets.getItem("Codes").getRange(`A${firstRow}:CX${lastRow}`);
                                    
                                    // Get the destination range
                                    const destinationRange = context.workbook.worksheets.getItem(currentWorksheetName).getRange(`A${pasteRow}`);
                                    
                                    // Copy the range with all properties
                                    destinationRange.copyFrom(sourceRange, Excel.RangeCopyType.all);
                                    
                                    await context.sync();
                                });
                                
                                await context.sync();
                                
                                // Apply the driver and assumption inputs function to the current worksheet
                                try {
                                    console.log(`Applying driver and assumption inputs to worksheet: ${currentWorksheetName}`);
                                    
                                    // Get the current worksheet and load its properties
                                    const currentWorksheet = context.workbook.worksheets.getItem(currentWorksheetName);
                                    currentWorksheet.load('name');
                                    await context.sync();
                                    
                                    await driverAndAssumptionInputs(
                                        currentWorksheet,
                                        pasteRow,
                                        code
                                    );
                                    console.log(`Successfully applied driver and assumption inputs to worksheet: ${currentWorksheetName}`);
                                } catch (error) {
                                    console.error(`Error applying driver and assumption inputs: ${error.message}`);
                                    result.errors.push({
                                        codeIndex: i,
                                        codeType: codeType,
                                        error: `Error applying driver and assumption inputs: ${error.message}`
                                    });
                                }
                                
                                result.processedCodes++;
                            }
                        } catch (error) {
                            console.error(`Error processing code ${codeType}:`, error);
                            throw error;
                        }
                    }).catch(error => {
                        console.error(`Error processing code ${codeType}: ${error.message}`);
                        result.errors.push({
                            codeIndex: i,
                            codeType: codeType,
                            error: error.message
                        });
                    });
                }
            } catch (error) {
                console.error(`Error processing code ${i}:`, error);
                result.errors.push({
                    codeIndex: i,
                    codeType: codeType,
                    error: error.message
                });
            }
        }
        
        // Prepare the final result object, including the names of assumption tabs
        const finalResult = {
            ...result, // Includes processedCodes, errors
            assumptionTabs: assumptionTabs.map(tab => tab.name) // Return only the names
        };

        console.log("runCodes finished. Returning:", finalResult);
        return finalResult; // Return the modified result object

    } catch (error) {
        console.error("Error in runCodes:", error);
        // Consider how to return errors. Throwing stops execution.
        // Returning them in the result allows the caller to decide.
        throw error; // Or return { errors: [error.message], assumptionTabs: [] }
    }
}

/**
 * Helper function to update cell references in formulas when rows are inserted
 * @param {string} formula - The formula to update
 * @param {number} rowOffset - The number of rows to offset
 * @returns {string} - The updated formula
 */
function updateFormulaReferences(formula, rowOffset) {
    if (!formula || !formula.startsWith('=')) {
        return formula;
    }
    
    // Regular expression to match cell references (e.g., A1, B2, etc.)
    const cellRefRegex = /([A-Z]+)([0-9]+)/g;
    
    // Replace each cell reference with an updated one
    return formula.replace(cellRefRegex, (match, col, row) => {
        const rowNum = parseInt(row);
        return `${col}${rowNum + rowOffset}`;
    });
}

/**
 * Tests if the active cell's fill color is #CCFFCC (light green)
 * @returns {Promise<boolean>} - True if the active cell is green, false otherwise
 */
export async function isActiveCellGreen() {
    try {
        console.log("Testing if cell B2 is green (#CCFFCC)");
        
        return await Excel.run(async (context) => {
            // Get cell B2 instead of the active cell
            const cellB2 = context.workbook.worksheets.getActiveWorksheet().getRange("B2");
            
            // Load the fill color property and address
            cellB2.load(["format/fill/color", "address"]);
            
            // Execute the request
            await context.sync();
            
            // Check if the color is #CCFFCC
            const isGreen = cellB2.format.fill.color === "#CCFFCC";
            
            console.log(`Cell B2 address: ${cellB2.address}, color: ${cellB2.format.fill.color}, Is green: ${isGreen}`);
            
            return isGreen;
        });
    } catch (error) {
        console.error("Error in isActiveCellGreen:", error);
        throw error;
    }
}


/**
 * Processes driver and assumption inputs for a worksheet based on code parameters,
 * replicating the logic from the VBA Driver_and_Assumption_Inputs function.
 * @param {Excel.Worksheet} worksheet - The initial Excel worksheet object.
 * @param {number} calcsPasteRow - The starting row for finding the code block.
 * @param {Object} code - The code object with type and parameters.
 * @returns {Promise<void>}
 */
export async function driverAndAssumptionInputs(worksheet, calcsPasteRow, code) {
    try {
        // --- Load worksheet name before calling helper ---
        // This requires its own context if worksheet object might not have name loaded yet
        let worksheetName = 'unknown';
        try {
             await Excel.run(async (context) => {
                 worksheet.load('name');
                 await context.sync();
                 worksheetName = worksheet.name;
                 
             });
         } catch(nameLoadError) {
             console.error("Failed to load worksheet name before calling helper", nameLoadError);
             throw new Error("Cannot determine worksheet name to proceed.");
         }

        // Define variable to store lastRow outside Excel.run scope so we can use it later
        let lastRow = 1000; // Default value in case of failure

        try {
            // Get a fresh worksheet reference and find the last row within a proper Excel.run context
            lastRow = await Excel.run(async (context) => {    
                // Get worksheet reference within THIS context by name
                const currentWorksheet = context.workbook.worksheets.getItem(worksheetName);
                
                // Get the used range of the worksheet
                const usedRange = currentWorksheet.getUsedRange();

                // Get the last row within the used range
                const lastRowRange = usedRange.getLastRow();

                // Load the rowIndex property of the last row
                lastRowRange.load("rowIndex");

                // Synchronize the state with the Excel document
                await context.sync();

                // Calculate the 1-based index of the last row
                const result = lastRowRange.rowIndex + 1;
                console.log('lastRow', result);
                
                // Return the value so it's accessible outside this Excel.run
                return result;
            });
        } catch(lastRowError) {
            console.error("Failed to determine last row", lastRowError);
            throw new Error("Cannot determine last row to proceed.");
        }

        // Ensure lastRow is a valid number (helper should return 1000 on error)
        if (typeof lastRow !== 'number' || lastRow <= 0) {
            console.error(`Last row determination failed or returned invalid value (${lastRow}). Cannot proceed safely.`);
            throw new Error("Failed to determine a valid last row for processing.");
        }
        // --- End Determine Last Row ---

        // Now, proceed with the main logic within its own Excel.run
        await Excel.run(async (context) => {
            // Pass the determined lastRow into this context
            const determinedLastRow = lastRow; 
            
            // Get worksheet reference within THIS context by name
            const currentWorksheet = context.workbook.worksheets.getItem(worksheetName);
            
            // USE calcsPasteRow in console log
            console.log(`Processing driver/assumption inputs for worksheet: ${worksheetName}, Code: ${code.type}, Start Row: ${calcsPasteRow}, Using Last Row: ${determinedLastRow}`);

            const columnSequence = ['A', 'B', 'C', 'G', 'H', 'I', 'K', 'L', 'M', 'N', 'O', 'P', 'R'];
            
            // Get the code value
            const codeValue = code.type;

            // Find the search row (first row >= calcsPasteRow where CodeValue is found in Col D)
            // Note: Using determinedLastRow obtained from the helper function
            let searchRow = -1; // 1-based index
            let searchRange = null; 
            let searchRangeAddress = ''; 

            try {
                // USE calcsPasteRow in console log
                console.log(`Attempting to get searchRange. calcsPasteRow: ${calcsPasteRow}, determinedLastRow: ${determinedLastRow}`);
                // USE calcsPasteRow in condition
                if (typeof calcsPasteRow !== 'number' || typeof determinedLastRow !== 'number' || calcsPasteRow <= 0 || determinedLastRow < calcsPasteRow) {
                     console.error(`Invalid range parameters for searchRange: calcsPasteRow=${calcsPasteRow}, determinedLastRow=${determinedLastRow}. Skipping search.`);
                     searchRow = -1; 
                 } else {
                     // USE calcsPasteRow for search range address
                     searchRangeAddress = `D${calcsPasteRow}:D${determinedLastRow}`;
                     console.log(`Creating searchRange with address: ${searchRangeAddress}`);
                     // Need to use the worksheet object available in *this* context
                     searchRange = currentWorksheet.getRange(searchRangeAddress);

                     console.log(`Loading values for searchRange: ${searchRangeAddress}`);
                     searchRange.load('values');
                     await context.sync();
                     console.log(`Successfully loaded values for searchRange: ${searchRangeAddress}`);
                 }

            } catch (loadError) {
                 console.error(`Error loading/syncing searchRange (${searchRangeAddress}): ${loadError.message}`, loadError);
                 searchRow = -1; 
            }

            // Check if searchRange was successfully created and loaded before accessing .values
            if (searchRange && searchRange.values) { // Check searchRange first!
                 console.log(`SearchRange (${searchRangeAddress}) has loaded values. Searching for codeValue: ${codeValue}`);
                 for (let i = 0; i < searchRange.values.length; i++) {
                    if (searchRange.values[i][0] === codeValue) {
                        // USE calcsPasteRow to calculate searchRow
                        searchRow = calcsPasteRow + i; // Found the row (1-based)
                        console.log(`Found codeValue '${codeValue}' at index ${i}, resulting searchRow: ${searchRow}`);
                        break;
                    }
                }
                 if (searchRow === -1) { // If loop finished without finding
                     console.log(`CodeValue '${codeValue}' not found within the loaded values of searchRange (${searchRangeAddress}).`);
                 }
            } else if (searchRow !== -1) { // Only log warning if we didn't already hit the loadError or invalid params
                 console.warn(`searchRange (${searchRangeAddress}) object exists but '.values' property is not available after sync. Search cannot be performed.`);
                 searchRow = -1; // Ensure we trigger the "not found" logic
            }

            if (searchRow === -1) {
                 console.warn(`Code type ${codeValue} not found or could not be searched for in column D (Range: ${searchRangeAddress || 'Invalid'}). Skipping inputs for this code.`);
                 return; // Exit if code not found or search failed
            }
            console.log(`Found code ${codeValue} at search row: ${searchRow}`);


            // Find the check row (first row >= searchrow where Col B is not light green #CCFFCC)
            // VBA: Do While .Cells(checkrow, 2).Interior.Color = RGB(204, 255, 204)
            let checkRow = searchRow;
            let isGreen = true;
            while (isGreen) {
                const cellB = currentWorksheet.getRange(`B${checkRow}`);
                cellB.load('format/fill/color');
                await context.sync();
                 // Handle potential errors if cell color can't be loaded
                 if (cellB.format && cellB.format.fill) {
                    isGreen = cellB.format.fill.color === '#CCFFCC';
                 } else {
                     console.warn(`Could not read fill color for cell B${checkRow}. Assuming not green.`);
                     isGreen = false; // Assume not green if color cannot be determined
                 }

                if (isGreen) {
                    checkRow++;
                    // Add a safety break condition?
                    if (checkRow > determinedLastRow + 50) { // Use determinedLastRow
                         console.error("Check row exceeded expected limits. Breaking loop.");
                         throw new Error("Failed to find non-green check row within reasonable bounds.");
                    }
                }
            }
            console.log(`Found check row (first non-green row in B at/after search row): ${checkRow}`);


            // Process drivers, labels, and financialsdriver (relative to searchRow)
            for (let k = 1; k <= 9; k++) {
                const targetRow = searchRow + k - 1;
                if (targetRow > determinedLastRow + 20) { // Safety check: Don't write way past the data
                    console.warn(`Target row ${targetRow} seems too high. Skipping write for k=${k}.`);
                    continue;
                }

                // Financials Driver (only for k=1)
                if (k === 1 && code.params.financialsdriver) {
                    const finDriverCell = currentWorksheet.getRange(`I${targetRow}`);
                    finDriverCell.values = [[code.params.financialsdriver]];
                    console.log(`Set financialsdriver at I${targetRow}: ${code.params.financialsdriver}`);
                }

                // Driver
                const driverParam = code.params[`driver${k}`];
                if (driverParam) {
                    const driverCell = currentWorksheet.getRange(`F${targetRow}`);
                    driverCell.values = [[driverParam]];
                     console.log(`Set driver${k} at F${targetRow}: ${driverParam}`);
                }

                // Label
                const labelParam = code.params[`label${k}`];
                if (labelParam) {
                    const labelCell = currentWorksheet.getRange(`B${targetRow}`);
                    labelCell.values = [[labelParam]];
                     console.log(`Set label${k} at B${targetRow}: ${labelParam}`);
                }
            }
            await context.sync(); // Sync after loop for efficiency


            // Process row items (inserting rows relative to checkRow)
            let currentCheckRowForInserts = checkRow; // Use a separate variable to track cumulative insertions correctly
            for (let g = 1; g <= 200; g++) { // Max 200 row parameters as in VBA
                const rowParam = code.params[`row${g}`];
                if (!rowParam) continue; // Skip if rowg parameter doesn't exist

                 console.log(`Processing row${g}: ${rowParam}`);

                const rowItems = rowParam.split('*');
                const numNewRows = rowItems.length - 1; // Number of rows to insert

                // Calculate the 1-based row number *before* potential insertions for this 'g' iteration
                // This takes into account rows inserted by previous 'g' loops via currentCheckRowForInserts
                const baseRowForThisG = currentCheckRowForInserts + g - 1;
                console.log(`Base row for row${g}: ${baseRowForThisG}, numNewRows: ${numNewRows}`);

                if (numNewRows > 0) {
                    // Insert new rows below the baseRowForThisG
                    const insertStartAddress = `${baseRowForThisG + 1}:${baseRowForThisG + numNewRows}`;
                    console.log(`Inserting ${numNewRows} rows at ${insertStartAddress}`);
                    const insertRange = currentWorksheet.getRange(insertStartAddress);
                    insertRange.insert(Excel.InsertShiftDirection.down);
                    await context.sync(); // Sync after insert

                    // Sequentially copy formats and formulas from the previous row to the newly inserted ones
                    // This helps ensure relative formulas are adjusted correctly step-by-step
                    console.log(`Copying formats/formulas sequentially for inserted rows.`);
                    for (let i = 0; i < numNewRows; i++) {
                        const sourceRowNum = baseRowForThisG + i;
                        const targetRowNum = baseRowForThisG + i + 1; // The newly inserted row
                        const sourceRowRange = currentWorksheet.getRange(`${sourceRowNum}:${sourceRowNum}`);
                        const targetRowRange = currentWorksheet.getRange(`${targetRowNum}:${targetRowNum}`);

                        // Copy formats
                        console.log(`  Copying formats from row ${sourceRowNum} to ${targetRowNum}`);
                        targetRowRange.copyFrom(sourceRowRange, Excel.RangeCopyType.formats);

                        // Copy formulas (should adjust relative references)
                        console.log(`  Copying formulas from row ${sourceRowNum} to ${targetRowNum}`);
                        targetRowRange.copyFrom(sourceRowRange, Excel.RangeCopyType.formulas);

                        // We could use RangeCopyType.all, but separate copy ensures population step overrides values cleanly.
                    }
                    await context.sync(); // Sync after all copies for this 'g' group are done
                    console.log("Finished sequential copy for inserted rows.");
                }

                // Populate the row(s) (original row + inserted rows)
                // This runs AFTER rows are inserted and structure (formats/formulas) is copied.
                for (let yy = 0; yy <= numNewRows; yy++) {
                    const currentRowNum = baseRowForThisG + yy; // 1-based row number to write to
                    const splitArray = rowItems[yy].split('|');
                    console.log(`Populating row ${currentRowNum} with items: ${rowItems[yy]}`);

                    for (let x = 0; x < splitArray.length; x++) {
                        // Check bounds for columnSequence
                        if (x >= columnSequence.length) {
                            console.warn(`Data item index ${x} exceeds columnSequence length (${columnSequence.length}). Skipping.`);
                            continue;
                        }

                        const valueToWrite = splitArray[x];
                        // VBA check: If splitArray(x) <> "" And splitArray(x) <> "F" Then
                        // 'F' likely means "Formula", so we don't overwrite if the value is 'F'.
                        if (valueToWrite && valueToWrite.toUpperCase() !== 'F') {
                            const colLetter = columnSequence[x];
                            const cellToWrite = currentWorksheet.getRange(`${colLetter}${currentRowNum}`);
                            // Attempt to infer data type (basic number check)
                            const numValue = Number(valueToWrite);
                            if (!isNaN(numValue) && valueToWrite.trim() !== '') {
                                cellToWrite.values = [[numValue]];
                            } else {
                                // Preserve existing value if empty string, otherwise write text
                                if (valueToWrite.trim() !== '') {
                                    cellToWrite.values = [[valueToWrite]];
                                }
                            }
                            // console.log(`  Wrote '${valueToWrite}' to ${colLetter}${currentRowNum}`);
                        }
                    }
                }
                await context.sync(); // Sync after populating each 'g' group

                // Adjust the base check row marker for subsequent 'g' iterations
                // by adding the number of rows inserted in *this* iteration.
                currentCheckRowForInserts += numNewRows;
                console.log(`Finished processing row${g}. currentCheckRowForInserts is now ${currentCheckRowForInserts}`);

            } // End for g loop

            console.log(`Completed processing driver and assumption inputs for code ${codeValue} in worksheet ${worksheetName}`);
        }); // End main Excel.run
    } catch (error) {
        console.error(`Error in driverAndAssumptionInputs MAIN CATCH for code '${code.type}' in worksheet '${worksheet?.name || 'unknown'}': ${error.message}`, error);
        throw error;
    }
}


/**
 * Finds the last used row in a specific column of a worksheet.
 * @param {Excel.Worksheet} worksheet - The worksheet to search in.
 * @param {string} columnLetter - The column letter (e.g., "B").
 * @returns {Promise<number>} - The 1-based index of the last used row, or 0 if the column is empty or an error occurs.
 */
async function getLastUsedRow(worksheet, columnLetter) {
    // Re-use worksheet object passed into the function within this Excel.run
    // Need context from the caller's Excel.run or wrap this in its own
    console.log(`Attempting to get last used row for column ${columnLetter} in sheet ${worksheet.name}`);
    try {
        // It's safer to re-get the worksheet by name if this is called outside the main loop's context
        // However, if called within the loop's context, using the passed object is fine.
        // For simplicity assuming it's called within a valid context for now.
        const fullColumn = worksheet.getRange(`${columnLetter}:${columnLetter}`);
        const usedRange = fullColumn.getUsedRange(true); // Use 'true' for valuesOnly parameter
        const lastCell = usedRange.getLastCell();
        lastCell.load("rowIndex");
        await worksheet.context.sync(); // Use the context associated with the worksheet object
        const lastRowIndex = lastCell.rowIndex + 1; // Convert 0-based index to 1-based row number
        console.log(`Last used row in column ${columnLetter} is ${lastRowIndex}`);
        return lastRowIndex;
    } catch (error) {
        // Handle cases where the column might be completely empty or other errors
        if (error.code === "ItemNotFound" || error.code === "GeneralException") {
            console.warn(`Could not find used range or last cell in column ${columnLetter} of sheet ${worksheet.name}. Assuming empty or header only (returning 0).`);
            return 0; // Return 0 if column is empty or error occurs
        }
        console.error(`Error in getLastUsedRow for column ${columnLetter} on sheet ${worksheet.name}:`, error);
        // It's often better to let the caller handle the error if it's unexpected.
        throw error; // Re-throw other errors
    }
    // Note: Removed the inner Excel.run as it complicates context management.
    // This function now expects to be called *within* an existing Excel.run context.
}

/**
 * Adjusts driver references in column AE based on lookups in column A.
 * Replicates the core logic of VBA Adjust_Drivers.
 * @param {Excel.Worksheet} worksheet - The assumption worksheet (within an Excel.run context).
 * @param {number} lastRow - The last row to process (inclusive).
 */
async function adjustDriversJS(worksheet, lastRow) {
    const START_ROW = 9;
    const DRIVER_CODE_COL = "F"; // Column containing the driver code to look up
    const LOOKUP_COL = "A";      // Column to search for the driver code
    const TARGET_COL = "AE";     // Column where the result address string is written

    console.log(`Running adjustDriversJS for sheet: ${worksheet.name} from row ${START_ROW} to ${lastRow}`);

    // Ensure lastRow is valid before proceeding
    if (lastRow < START_ROW) {
        console.warn(`adjustDriversJS: lastRow (${lastRow}) is less than START_ROW (${START_ROW}). Skipping.`);
        return;
    }

    try {
        // Define the ranges to load
        const driverCodeRangeAddress = `${DRIVER_CODE_COL}${START_ROW}:${DRIVER_CODE_COL}${lastRow}`;
        const lookupRangeAddress = `${LOOKUP_COL}${START_ROW}:${LOOKUP_COL}${lastRow}`;
        const driverCodeRange = worksheet.getRange(driverCodeRangeAddress);
        const lookupRange = worksheet.getRange(lookupRangeAddress);

        // Load values from both columns
        driverCodeRange.load("values");
        lookupRange.load("values");
        await worksheet.context.sync(); // Sync to get the values

        const driverCodeValues = driverCodeRange.values;
        const lookupValues = lookupRange.values;

        // Create a map for efficient lookup: { lookupValue: rowIndex }
        // Note: rowIndex here is the 1-based Excel row number
        const lookupMap = new Map();
        for (let i = 0; i < lookupValues.length; i++) {
            const value = lookupValues[i][0];
            // Only add non-empty values to the map. Handle potential duplicates?
            // VBA's .Find typically finds the first match. Map naturally stores the last encountered.
            if (value !== null && value !== "") {
                 // The row number in Excel is START_ROW + index
                lookupMap.set(value, START_ROW + i);
            }
        }
        console.log(`Built lookup map from ${LOOKUP_COL}${START_ROW}:${LOOKUP_COL}${lastRow} with ${lookupMap.size} entries.`);

        // Prepare the output values for the target column AE
        // Initialize with nulls or empty strings to clear previous values potentially
        const outputValues = []; // Array of arrays for Excel range: [[value1], [value2], ...]
        let foundCount = 0;
        let notFoundCount = 0;

        for (let i = 0; i < driverCodeValues.length; i++) {
            const driverCode = driverCodeValues[i][0];
            const currentRow = START_ROW + i; // Current Excel row being processed

            if (driverCode !== null && driverCode !== "") {
                if (lookupMap.has(driverCode)) {
                    const foundRow = lookupMap.get(driverCode);
                    const targetAddress = `${TARGET_COL}${foundRow}`;
                    outputValues.push([targetAddress]); // Store as [[value]] for range write
                    foundCount++;
                    // console.log(`Row ${currentRow} (${DRIVER_CODE_COL}): Found '${driverCode}' in ${LOOKUP_COL} at row ${foundRow}. Setting ${TARGET_COL}${currentRow} = '${targetAddress}'`);
                } else {
                    // Value in F not found in A
                    console.warn(`adjustDriversJS: Driver code '${driverCode}' from cell ${DRIVER_CODE_COL}${currentRow} not found in range ${lookupRangeAddress}.`);
                    outputValues.push([null]); // Or [""] or keep existing? VBA doesn't explicitly clear. Using null.
                    notFoundCount++;
                }
            } else {
                // Empty cell in F, write null to corresponding AE cell
                outputValues.push([null]);
            }
        }

        // Write the results back to column AE
        if (outputValues.length > 0) {
            const targetRangeAddress = `${TARGET_COL}${START_ROW}:${TARGET_COL}${lastRow}`;
            const targetRange = worksheet.getRange(targetRangeAddress);
            console.log(`Writing ${foundCount} results (${notFoundCount} not found) to ${targetRangeAddress}`);
            targetRange.values = outputValues;
            // Sync will happen in the caller's context
        } else {
             console.log(`adjustDriversJS: No values to write to ${TARGET_COL}.`);
        }

    } catch (error) {
        console.error(`Error in adjustDriversJS for sheet ${worksheet.name}:`, error);
        // Decide if error should be re-thrown to stop the whole process
        // throw error;
    }
    // No context.sync() here - it should be handled by the calling function (processAssumptionTabs)
}

/**
 * Replaces INDIRECT functions in a specified column range with their evaluated values.
 * Mimics the VBA Replace_Indirects logic using batched range value lookups.
 * @param {Excel.Worksheet} worksheet - The assumption worksheet (within an Excel.run context).
 * @param {number} lastRow - The last row to process.
 */
async function replaceIndirectsJS(worksheet, lastRow) {
    const START_ROW = 9;
    const TARGET_COL = "AE";

    console.log(`Running replaceIndirectsJS for sheet: ${worksheet.name} from row ${START_ROW} to ${lastRow}`);

    if (lastRow < START_ROW) {
        console.warn(`replaceIndirectsJS: lastRow (${lastRow}) is less than START_ROW (${START_ROW}). Skipping.`);
        return;
    }

    const targetRangeAddress = `${TARGET_COL}${START_ROW}:${TARGET_COL}${lastRow}`;
    const targetRange = worksheet.getRange(targetRangeAddress);

    try {
        // 1. Load formulas from the target range
        targetRange.load("formulas");
        await worksheet.context.sync();

        const originalFormulas = targetRange.formulas; // 2D array [[f1], [f2], ...]
        const referencesToLookup = new Map(); // Map<string, { range: Excel.Range | null, value: any }>
        const formulaData = []; // Array<{ originalFormula: string, index: number }>

        // 2. First Pass: Identify all unique INDIRECT arguments
        console.log("Replace_Indirects: Pass 1 - Identifying INDIRECT arguments");
        for (let i = 0; i < originalFormulas.length; i++) {
            let formula = originalFormulas[i][0];
            formulaData.push({ originalFormula: formula, index: i }); // Store original formula and index

            if (typeof formula === 'string') {
                // Use a loop to find all INDIRECT occurrences in a single formula
                let searchStartIndex = 0;
                while (true) {
                    const upperFormula = formula.toUpperCase();
                    const indirectStartIndex = upperFormula.indexOf("INDIRECT(", searchStartIndex);

                    // Stop if no more INDIRECT found or if it might be part of INDEX
                    if (indirectStartIndex === -1 || upperFormula.includes("INDEX(")) {
                        break;
                    }

                    // Find the matching closing parenthesis (simple approach)
                    const parenStartIndex = indirectStartIndex + "INDIRECT(".length;
                    const parenEndIndex = formula.indexOf(")", parenStartIndex);

                    if (parenEndIndex === -1) {
                        console.warn(`Row ${START_ROW + i}: Malformed INDIRECT found in formula: ${formula}`);
                        break; // Cannot process this INDIRECT
                    }

                    const argString = formula.substring(parenStartIndex, parenEndIndex).trim();

                    // Validate argString looks like a cell/range reference (basic check)
                    // This helps avoid trying to load ranges like "Sheet1!A:A" which might fail or be slow
                    if (argString && /^[A-Za-z0-9_!$:'". ]+$/.test(argString) && !referencesToLookup.has(argString)) {
                         console.log(`  Found reference to lookup: ${argString}`);
                         referencesToLookup.set(argString, { range: null, value: undefined }); // Placeholder
                    }

                    // Continue searching after this INDIRECT
                    searchStartIndex = parenEndIndex + 1;
                }
            }
        }

        // 3. Batch Load Values for identified references
        console.log(`Replace_Indirects: Loading values for ${referencesToLookup.size} unique references.`);
        if (referencesToLookup.size > 0) {
            for (const [refString, data] of referencesToLookup.entries()) {
                try {
                    // Attempt to get the range and load its value
                    data.range = worksheet.getRange(refString);
                    // Load values. Consider loading formulas too if INDIRECT might point to a formula cell.
                    // Loading numberFormat might help distinguish between 0 and empty.
                    data.range.load(["values", "text"]); // Load text to handle "DELETE" easily
                } catch (rangeError) {
                    console.warn(`Replace_Indirects: Error getting range for reference "${refString}". It might be invalid or on another sheet.`, rangeError.debugInfo || rangeError.message);
                     // Keep data.range as null, will be handled later
                    referencesToLookup.set(refString, { range: null, value: '#REF!' }); // Mark as error
                }
            }
            await worksheet.context.sync(); // Sync all loaded values

            // Populate the values in the map
            for (const [refString, data] of referencesToLookup.entries()) {
                 if (data.range) { // If range was successfully retrieved
                     try {
                         // Use .text to directly compare with "DELETE"
                         // Use .values for the actual numeric/boolean value if not "DELETE"
                        const cellText = data.range.text[0][0];
                        if (cellText === "DELETE") {
                            data.value = "0"; // Replace "DELETE" with "0" string as per VBA
                        } else {
                             // Use the actual value (could be string, number, boolean)
                             // Prefer values[0][0] as it respects data types better than text
                             data.value = data.range.values[0][0];
                        }
                     } catch (valueError) {
                         console.warn(`Replace_Indirects: Error reading value for reference "${refString}" after sync.`, valueError.debugInfo || valueError.message);
                         data.value = '#VALUE!'; // Or another suitable error indicator
                     }
                 }
                 // If data.range was null or value fetch failed, data.value remains '#REF!' or '#VALUE!'
            }
             console.log("Replace_Indirects: Finished loading reference values.");
        }


        // 4. Second Pass: Replace INDIRECT with looked-up values
        console.log("Replace_Indirects: Pass 2 - Replacing INDIRECT calls.");
        const newFormulas = []; // Array of arrays: [[newF1], [newF2], ...]
        for (const item of formulaData) {
            let currentFormula = item.originalFormula;

            if (typeof currentFormula === 'string') {
                let loopCount = 0; // Safety break
                const MAX_LOOPS = 20; // Prevent infinite loops for complex/circular cases

                while (loopCount < MAX_LOOPS) {
                    const upperFormula = currentFormula.toUpperCase();
                    const indirectStartIndex = upperFormula.indexOf("INDIRECT(");

                    if (indirectStartIndex === -1 || upperFormula.includes("INDEX(")) {
                        break; // No more INDIRECTs (or INDEX present)
                    }

                    const parenStartIndex = indirectStartIndex + "INDIRECT(".length;
                    const parenEndIndex = currentFormula.indexOf(")", parenStartIndex);

                    if (parenEndIndex === -1) {
                         // Already warned in pass 1, just break here
                        break;
                    }

                    const indString = currentFormula.substring(indirectStartIndex, parenEndIndex + 1); // The full INDIRECT(...)
                    const argString = currentFormula.substring(parenStartIndex, parenEndIndex).trim();

                    let directRef = '#REF!'; // Default if lookup fails
                     if (referencesToLookup.has(argString)) {
                         directRef = referencesToLookup.get(argString).value;
                     } else {
                         // Argument wasn't identified/loaded (maybe invalid?)
                         console.warn(`Row ${START_ROW + item.index}: INDIRECT argument "${argString}" not found in lookup map during replacement.`);
                     }

                    // Handle potential null/undefined values from lookup - treat as 0? VBA doesn't explicitly handle this.
                    // Let's treat null/undefined as 0 for replacement to avoid inserting 'null' or 'undefined' into formulas.
                     // Empty string "" should probably remain "" unless it was "DELETE".
                     if (directRef === null || typeof directRef === 'undefined') {
                         directRef = 0; // Replace null/undefined with numeric 0
                     } else if (directRef === "") {
                          // Keep empty string as empty string unless it was originally "DELETE"
                          // The map handles "DELETE" -> "0" already
                     } else if (typeof directRef === 'string') {
                         // If the resolved value is a string, potentially needs quoting if replacing in a formula context?
                         // VBA seems to just concatenate the value directly. Let's follow that.
                         // Example: =SUM(INDIRECT("A1")) where A1 contains "B2" becomes =SUM(B2)
                         // Example: =CONCATENATE("Result: ",INDIRECT("A1")) where A1 contains "Success" becomes =CONCATENATE("Result: ","Success") - requires quotes?
                         // VBA appears to handle this implicitly. JS replace won't add quotes.
                         // Let's test behavior, may need adjustment if it breaks formulas expecting strings.
                         // For now, direct replacement. Consider adding quotes if `directRef` is text AND the context requires it.
                     } else if (typeof directRef === 'boolean') {
                         directRef = directRef ? 'TRUE' : 'FALSE'; // Convert boolean to formula text
                     }
                     // Numeric values are fine as is.

                    // Perform the replacement. Use replace directly on the found indString.
                    currentFormula = currentFormula.replace(indString, String(directRef));
                    loopCount++;

                } // End while loop for single formula processing

                if (loopCount === MAX_LOOPS) {
                    console.warn(`Row ${START_ROW + item.index}: Max replacement loops reached for formula. Result might be incomplete: ${currentFormula}`);
                }
            }
            // Add the processed formula (or original if not string/no INDIRECT) to the result array
            newFormulas.push([currentFormula]);

        } // End for loop processing all formulas

        // 5. Write the modified formulas back to the range
        console.log(`Replace_Indirects: Writing ${newFormulas.length} updated formulas back to ${targetRangeAddress}`);
        targetRange.formulas = newFormulas;

        // Sync is handled by the caller (processAssumptionTabs)

    } catch (error) {
        console.error(`Error in replaceIndirectsJS for sheet ${worksheet.name} range ${targetRangeAddress}:`, error.debugInfo || error);
        // Re-throw the error to allow the calling function to handle it
        throw error;
    }
}

/**
 * Placeholder for Populate_Financials VBA logic.
 * Populates the "Financials" sheet based on codes in the assumption sheet.
 * @param {Excel.Worksheet} worksheet - The assumption worksheet (within an Excel.run context).
 * @param {number} lastRow - The last row to process in the assumption sheet.
 * @param {Excel.Worksheet} financialsSheet - The "Financials" worksheet (within the same Excel.run context).
 */
async function populateFinancialsJS(worksheet, lastRow, financialsSheet) {
    console.log(`Running populateFinancialsJS for sheet: ${worksheet.name} (lastRow: ${lastRow}) -> ${financialsSheet.name}`);
    // This function MUST be called within an Excel.run context.

    const CALCS_FIRST_ROW = 9; // Same as START_ROW elsewhere
    const ASSUMPTION_CODE_COL = "C"; // Column with code to lookup on assumption sheet
    const ASSUMPTION_LINK_COL_B = "B";
    const ASSUMPTION_LINK_COL_D = "D";
    // Column on assumption sheet to link for monthly data
    const ASSUMPTION_MONTHS_START_COL = "AE";

    const FINANCIALS_CODE_COLUMN = "I"; // Column to search for code on Financials sheet
    const FINANCIALS_TARGET_COL_B = "B";
    const FINANCIALS_TARGET_COL_D = "D";
    const FINANCIALS_ANNUALS_START_COL = "J"; // Annuals start here
    const FINANCIALS_MONTHS_START_COL = "AE"; // Months start here

    // --- Updated Column Definitions ---
    const ANNUALS_END_COL = "P";       // Annuals end here
    const MONTHS_END_COL = "CX";       // Months end here
    // --- End Updated Column Definitions ---

    // Formatting constants
    // const PURPLE_COLOR = "#800080"; // RGB(128, 0, 128) - Removed as Actuals section is removed
    const GREEN_COLOR = "#008000";  // RGB(0, 128, 0)
    const CURRENCY_FORMAT = '_(* $#,##0_);_(* $(#,##0);_(* "$" -_);_(@_)';

    // Ensure lastRow is valid
    if (lastRow < CALCS_FIRST_ROW) {
        console.warn(`populateFinancialsJS: lastRow (${lastRow}) is less than CALCS_FIRST_ROW (${CALCS_FIRST_ROW}). Skipping.`);
        return;
    }

    try {
        // 1. Load data from Assumption Sheet
        console.log(`populateFinancialsJS: Loading assumption data up to row ${lastRow}`);
        const assumptionCodeRange = worksheet.getRange(`${ASSUMPTION_CODE_COL}${CALCS_FIRST_ROW}:${ASSUMPTION_CODE_COL}${lastRow}`);
        // No need to load B, D, AE addresses/values here anymore if only used for linking

        assumptionCodeRange.load("values");

        // 2. Load data from Financials Sheet (Find last row in code column I)
        const financialsSearchCol = financialsSheet.getRange(`${FINANCIALS_CODE_COLUMN}:${FINANCIALS_CODE_COLUMN}`);
        const financialsUsedRange = financialsSearchCol.getUsedRange(true);
        financialsUsedRange.load("rowCount");
        // It's okay to sync assumption and initial financials loads together
        // await worksheet.context.sync(); // Removed intermediate sync

        let financialsLastRow = 0;
        // Sync financials rowCount load before calculating financialsLastRow
        await worksheet.context.sync();
        if (financialsUsedRange.rowCount > 0) {
           try {
              const lastCell = financialsUsedRange.getLastCell();
              lastCell.load("rowIndex");
               await worksheet.context.sync();
              financialsLastRow = lastCell.rowIndex + 1;
           } catch(e) {
               console.warn(`Could not get last cell directly for Financials col ${FINANCIALS_CODE_COLUMN}. Error: ${e.message}. Attempting fallback range loading.`);
               try {
                   const fallbackRange = financialsSheet.getRange(`${FINANCIALS_CODE_COLUMN}1:${FINANCIALS_CODE_COLUMN}10000`);
                   fallbackRange.load("values");
                   await worksheet.context.sync();
                   for (let i = fallbackRange.values.length - 1; i >= 0; i--) {
                       if (fallbackRange.values[i][0] !== null && fallbackRange.values[i][0] !== "") {
                           financialsLastRow = i + 1;
                           break;
                       }
                   }
                   if (financialsLastRow === 0) console.warn(`Fallback range load for Financials col ${FINANCIALS_CODE_COLUMN} also yielded no data.`);
               } catch (fallbackError) {
                    console.error(`Error during fallback range loading for Financials col ${FINANCIALS_CODE_COLUMN}:`, fallbackError);
                    financialsLastRow = 0;
               }
           }
        }
        console.log(`Financials last relevant row in column ${FINANCIALS_CODE_COLUMN}: ${financialsLastRow}`);

        // 3. Create Map of Financials Codes (Col I) -> Row Number
        const financialsCodeMap = new Map();
        if (financialsLastRow > 0) {
            const financialsCodeRange = financialsSheet.getRange(`${FINANCIALS_CODE_COLUMN}1:${FINANCIALS_CODE_COLUMN}${financialsLastRow}`);
            financialsCodeRange.load("values");
            await worksheet.context.sync(); // Sync map data load
            for (let i = 0; i < financialsCodeRange.values.length; i++) {
                const code = financialsCodeRange.values[i][0];
                if (code !== null && code !== "") {
                    financialsCodeMap.set(code, i + 1);
                }
            }
            console.log(`Built Financials code map with ${financialsCodeMap.size} entries.`);
        } else {
            console.warn(`Financials sheet column ${FINANCIALS_CODE_COLUMN} appears empty or last row not found. No codes loaded.`);
        }

        // 4. Identify rows to insert and prepare task data
        const tasks = [];
        console.log("populateFinancialsJS: Syncing assumption codes load...");
        await worksheet.context.sync(); // Sync needed for assumptionCodeRange.values

        const assumptionCodes = assumptionCodeRange.values;
        console.log(`populateFinancialsJS: Processing ${assumptionCodes?.length ?? 0} assumption rows.`);

        // --- REMOVED Debug logging for row 17 values/addresses ---

        for (let i = 0; i < (assumptionCodes?.length ?? 0); i++) {
            const code = assumptionCodes[i][0];
            const assumptionRow = CALCS_FIRST_ROW + i; // This is the correct Excel row number

            if (code !== null && code !== "") {
                if (financialsCodeMap.has(code)) {
                    const targetRow = financialsCodeMap.get(code);

                    // --- Manually construct the address strings ---
                    const cellAddressB = `${ASSUMPTION_LINK_COL_B}${assumptionRow}`;
                    const cellAddressD = `${ASSUMPTION_LINK_COL_D}${assumptionRow}`;
                    const cellAddressMonths = `${ASSUMPTION_MONTHS_START_COL}${assumptionRow}`;

                    // Construct the full formula links directly
                    const formulaLinkB = `='${worksheet.name}'!${cellAddressB}`;
                    const formulaLinkD = `='${worksheet.name}'!${cellAddressD}`;
                    const formulaLinkMonths = `='${worksheet.name}'!${cellAddressMonths}`;

                    // --- REMOVED getSimpleAddress helper function and related checks ---
                    console.log(`  Task Prep: Row ${assumptionRow}, Code ${code}`); // Simplified log

                    tasks.push({
                        targetRow: targetRow,
                        assumptionRow: assumptionRow,
                        code: code,
                        addressB: formulaLinkB,     // Use the constructed formula link
                        addressD: formulaLinkD,     // Use the constructed formula link
                        addressMonths: formulaLinkMonths // Use the constructed formula link
                    });
                }
            }
        }

        if (tasks.length === 0) {
            console.log("No matching codes found. Nothing to insert or populate.");
            return;
        }

        // 5. Sort tasks by targetRow DESCENDING
        tasks.sort((a, b) => b.targetRow - a.targetRow);
        console.log(`Sorted ${tasks.length} tasks for insertion.`);
        // --- DEBUG: Log the tasks array --- 
        // console.log("Tasks array (sorted desc by targetRow):", JSON.stringify(tasks)); // REMOVED DEBUG
        // --- END DEBUG ---

        // 6. Perform Insertions (bottom-up)
        console.log("Performing row insertions...");
        for (const task of tasks) { // Uses the DESCENDING sorted tasks
            financialsSheet.getRange(`${task.targetRow}:${task.targetRow}`).insert(Excel.InsertShiftDirection.down);
            // *** It's generally more efficient to sync less often, but syncing after each insert
            // ensures the row model is updated for potential complex dependencies if they existed.
            // Keep sync here for now unless performance becomes an issue. ***
            // await worksheet.context.sync(); // Sync after EACH insertion -- REMOVED THIS LINE
        }
        await worksheet.context.sync(); // Sync AFTER all insertions are queued
        console.log("Finished row insertions.");

        // Pre-calculate the final adjusted row for each task after all insertions
        console.log("Calculating final adjusted rows for population/autofill...");
        // Get unique original target rows, sorted ascending
        const originalTargetRowsAsc = [...new Set(tasks.map(t => t.targetRow))].sort((a, b) => a - b);
        const taskAdjustedRows = new Map(); // Map to store { assumptionRow: adjustedRow }
        let totalShift = 0; // Total shift accumulated from previous rows

        // --- DEBUG: Log originalTargetRowsAsc ---
        // console.log("Original Target Rows (unique, asc):", originalTargetRowsAsc); // REMOVED DEBUG
        // --- END DEBUG ---

        originalTargetRowsAsc.forEach(uniqueRow => {
            // --- DEBUG: Log current uniqueRow ---
            // console.log(`Processing uniqueRow: ${uniqueRow}`); // REMOVED DEBUG
            // --- END DEBUG ---

            // Find all tasks that originally targeted this unique row
            // CORRECTED PROPERTY NAME IN FILTER: task.targetRow instead of task.originalTargetRow
            const tasksAtThisRow = tasks.filter(task => task.targetRow === uniqueRow);

            // --- DEBUG: Log tasks found for this uniqueRow ---
            // console.log(`  Tasks found for uniqueRow ${uniqueRow}:`, JSON.stringify(tasksAtThisRow)); // REMOVED DEBUG
            // --- END DEBUG ---

            // Optional: Sort tasksAtThisRow by assumptionRow for deterministic order, though might not be strictly necessary
            // tasksAtThisRow.sort((a, b) => a.assumptionRow - b.assumptionRow);

            let currentAdjustedRowForGroup = uniqueRow + totalShift; // Starting adjusted row for this group

            // Assign consecutive adjusted rows to each task in this group
            tasksAtThisRow.forEach(task => {
                taskAdjustedRows.set(task.assumptionRow, currentAdjustedRowForGroup); // Use assumptionRow as key
                console.log(`  Mapping: Code ${task.code}, Assumption Row ${task.assumptionRow}, Original Target ${uniqueRow}, Final Adjusted Row ${currentAdjustedRowForGroup}`);
                currentAdjustedRowForGroup++; // Increment for the next task inserting at the same original spot
            });

            // Update the total shift for subsequent unique rows
            totalShift += tasksAtThisRow.length;
        });

        // --- DEBUG: Log the contents of the map --- 
        // console.log("taskAdjustedRows map contents:", taskAdjustedRows); // REMOVED DEBUG
        // --- END DEBUG ---

        // 7. Populate and Format inserted rows using ADJUSTED row numbers
        console.log("Populating inserted rows (using adjusted rows)...");
        for (const task of tasks) { // Iterates descending sorted tasks (order doesn't strictly matter here, but using the same loop)
            // const originalTargetRow = task.targetRow; // No longer needed for lookup
            const populateRow = taskAdjustedRows.get(task.assumptionRow); // Get the calculated adjusted row using assumptionRow

            // Check if populateRow was found
            if (typeof populateRow === 'undefined' || populateRow === null) {
                console.error(`Could not find adjusted row for task with Assumption Row ${task.assumptionRow}, Code ${task.code}. Skipping population.`);
                continue; // Skip this task if mapping failed
            }

            // Use populateRow instead of task.targetRow for getRange calls
            const cellB = financialsSheet.getRange(`${FINANCIALS_TARGET_COL_B}${populateRow}`);
            const cellD = financialsSheet.getRange(`${FINANCIALS_TARGET_COL_D}${populateRow}`);
            const cellAnnualsStart = financialsSheet.getRange(`${FINANCIALS_ANNUALS_START_COL}${populateRow}`);
            const cellMonthsStart = financialsSheet.getRange(`${FINANCIALS_MONTHS_START_COL}${populateRow}`);

            // --- Populate Column B ---
            cellB.formulas = [[task.addressB]]; // Set formula directly
            cellB.format.font.bold = false;
            cellB.format.font.italic = false;
            cellB.format.indentLevel = 2;

            // --- Populate Column D ---
            cellD.formulas = [[task.addressD]]; // Set formula directly
            cellD.format.font.bold = false;
            cellD.format.font.italic = false;
            cellD.format.indentLevel = 2;

            // --- Populate Annuals Start Column (J) with SUMIF ---
            const codePrefix = String(task.code).substring(0, 2).toUpperCase();
            let formulaJ = "";
            if (codePrefix === "IS" || codePrefix === "CF") {
                 // Corrected R2C[1] to R2C: Criteria should reference current column's header (J$2)
                 formulaJ = `=SUMIF(R3,R2C,R[0])`;
            } else {
                 // Corrected R2C[1] to R2C: Criteria should reference current column's header (J$2)
                 formulaJ = `=SUMIF(R4,R2C,R[0])`;
            }
            cellAnnualsStart.formulasR1C1 = [[formulaJ]]; // Use formulasR1C1 for SUMIF
            cellAnnualsStart.format.font.bold = false;
            cellAnnualsStart.format.font.italic = false;
            cellAnnualsStart.format.numberFormat = CURRENCY_FORMAT;

            // --- Populate Months Start Column (AE) with Link ---
            cellMonthsStart.formulas = [[task.addressMonths]]; // Set formula directly
            cellMonthsStart.format.font.bold = false;
            cellMonthsStart.format.font.italic = false;
            cellMonthsStart.format.font.color = GREEN_COLOR; // Keep green color for month links
            cellMonthsStart.format.numberFormat = CURRENCY_FORMAT;

            // Removed Actuals column population (was L in previous version)
        }
        console.log("Finished setting values/formulas/formats for inserted rows.");
        await worksheet.context.sync(); // Sync all population and formatting


        // 8. Perform Autofills using ADJUSTED row numbers
        console.log("Performing autofills (using adjusted rows)...");
        for (const task of tasks) { // Uses the DESCENDING sorted tasks again
             // const originalTargetRow = task.targetRow; // No longer needed for lookup
             const populateRow = taskAdjustedRows.get(task.assumptionRow); // Get the calculated adjusted row using assumptionRow

             // Check if populateRow was found
            if (typeof populateRow === 'undefined' || populateRow === null) {
                console.error(`Could not find adjusted row for task with Assumption Row ${task.assumptionRow}, Code ${task.code}. Skipping autofill.`);
                continue; // Skip this task if mapping failed
            }

             try {
                // Use populateRow for autofill ranges
                // Autofill Annuals: J -> P
                const sourceAnnuals = financialsSheet.getRange(`${FINANCIALS_ANNUALS_START_COL}${populateRow}`); // Use adjusted row
                const destAnnuals = financialsSheet.getRange(`${FINANCIALS_ANNUALS_START_COL}${populateRow}:${ANNUALS_END_COL}${populateRow}`); // Use adjusted row
                sourceAnnuals.autoFill(destAnnuals, Excel.AutoFillType.fillDefault);
                // console.log(`  Autofilled ${FINANCIALS_ANNUALS_START_COL}${populateRow} to ${ANNUALS_END_COL}${populateRow}`);

                // Autofill Months: AE -> CX
                const sourceMonths = financialsSheet.getRange(`${FINANCIALS_MONTHS_START_COL}${populateRow}`); // Use adjusted row
                const destMonths = financialsSheet.getRange(`${FINANCIALS_MONTHS_START_COL}${populateRow}:${MONTHS_END_COL}${populateRow}`); // Use adjusted row
                sourceMonths.autoFill(destMonths, Excel.AutoFillType.fillDefault);
                // console.log(`  Autofilled ${FINANCIALS_MONTHS_START_COL}${populateRow} to ${MONTHS_END_COL}${populateRow}`);

                // Removed Actuals autofill
             } catch(autofillError) {
                 // Update error message to use adjusted row
                 console.error(`Error during autofill for adjusted row ${populateRow} (Code: ${task.code}, Original Target: ${task.targetRow}):`, autofillError.debugInfo || autofillError);
             }
        }
        console.log("Finished setting up autofills.");
        await worksheet.context.sync(); // Sync all autofill operations
        console.log("Autofills synced.");

        console.log(`populateFinancialsJS successfully completed for ${worksheet.name} -> ${financialsSheet.name}`);

    } catch (error) {
        console.error(`Error in populateFinancialsJS for sheet ${worksheet.name} -> ${financialsSheet.name}:`, error.debugInfo || error);
        throw error;
    }
}

/**
 * Placeholder for Link_Fin_References VBA logic.
 * Links assumption sheet cells (K onwards) to the "Financials" sheet based on codes in col E.
 * @param {Excel.Worksheet} worksheet - The assumption worksheet (within an Excel.run context).
 * @param {number} lastRow - The last row to process.
 * @param {Excel.Worksheet} financialsSheet - The "Financials" worksheet (within the same Excel.run context).
 */
async function linkFinReferencesJS(worksheet, lastRow, financialsSheet) {
    console.log(`Running linkFinReferencesJS for sheet: ${worksheet.name} (lastRow: ${lastRow}) -> ${financialsSheet.name}`);
    // This function MUST be called within an Excel.run context.
    const START_ROW = 9;
    const CHECK_COLUMN = "E";
    const FINANCIALS_LOOKUP_COLUMN = "B";
    const TARGET_COLUMN_START = "K";
    const NON_GREEN_COLOR = "#CCFFCC"; // VBA: RGB(204, 255, 204)
    // Add constant for monthscolumnend (e.g., "CX")

    // TODO: Implement Link_Fin_References logic
    // 1. Load values & formats from E<START_ROW>:E<lastRow> (assumption)
    // 2. Load values from FINANCIALS_LOOKUP_COLUMN (Financials, use getUsedRange().getIntersection)
    // 3. Create map of Financials names (col B) to row numbers
    // 4. Iterate assumption rows (i)
    // 5. Check conditions: value not empty, not "DASS", not 0, cell color NOT NON_GREEN_COLOR
    // 6. If conditions met:
    //    a. Find referenceRow in Financials map
    //    b. Set formula K<i> = ='Financials'!K<referenceRow>
    //    c. Autofill K<i> to monthscolumnend<i>
    // 7. Clear column "Q" (Q<START_ROW>:Q<lastRow>)
    worksheet.load('name'); // Keep references valid if needed later in the SAME context
    financialsSheet.load('name');
    await worksheet.context.sync();
     console.warn(`linkFinReferencesJS for ${worksheet.name} -> ${financialsSheet.name} not implemented yet.`);

}

/**
 * Placeholder for Format_Changes_In_Working_Capital VBA logic.
 * Inserts a row and adjusts formatting in "Financials" based on specific codes.
 * @param {Excel.Worksheet} financialsSheet - The "Financials" worksheet (within an Excel.run context).
 */
async function formatChangesInWorkingCapitalJS(financialsSheet) {
    console.log(`Running formatChangesInWorkingCapitalJS for sheet: ${financialsSheet.name}`);
    // This function MUST be called within an Excel.run context.
    const FIND_TEXT_1 = "CF: Non-cash";
    const FIND_TEXT_2 = "CF: WC";
    const SEARCH_COLUMN = "I";
     // Define ranges for border removal based on VBA (e.g., "K" + (foundRow + 1) + ":P" + (foundRow + 1))

    // TODO: Implement Format_Changes_In_Working_Capital logic
    // 1. Find FIND_TEXT_1 in SEARCH_COLUMN using range.find
    // 2. If found:
    //    a. Get cell above (offset -1, 0)
    //    b. Load its value
    //    c. Sync
    //    d. Check if value is FIND_TEXT_2
    //    e. If yes:
    //       i. Get the entire row of the found cell (.getEntireRow())
    //       ii. Insert a new row above it (insert(Excel.InsertShiftDirection.down))
    //       iii. Get ranges in the *original* row (now shifted down one) and remove borders.
    financialsSheet.load('name'); // Keep reference valid if needed later in the SAME context
    await financialsSheet.context.sync();
     console.warn(`formatChangesInWorkingCapitalJS on ${financialsSheet.name} not implemented yet.`);
}

/**
 * Processes assumption tabs after runCodes, replicating VBA logic.
 * Expects an array of assumption tab names.
 * @param {string[]} assumptionTabNames - Array of assumption tab names created by runCodes.
 */
export async function processAssumptionTabs(assumptionTabNames) {
    console.log(`Starting processing for ${assumptionTabNames.length} assumption tabs:`, assumptionTabNames);
    if (!assumptionTabNames || assumptionTabNames.length === 0) {
        console.log("No assumption tabs provided to process.");
        return;
    }

    const FINANCIALS_SHEET_NAME = "Financials"; // Define constant
    const AUTOFILL_START_COLUMN = "AE";
    const AUTOFILL_END_COLUMN = "CX";
    const START_ROW = 9; // Standard start row for processing

    try {
        // --- Loop through each assumption tab name ---
        for (const worksheetName of assumptionTabNames) {
             console.log(`\nProcessing Assumption Tab: ${worksheetName}`);

            try {
                 // Perform operations for a single tab within one Excel.run for efficiency
                 await Excel.run(async (context) => {
                     // Get worksheet and financials sheet references within THIS context
                     const currentWorksheet = context.workbook.worksheets.getItem(worksheetName);
                     const financialsSheet = context.workbook.worksheets.getItem(FINANCIALS_SHEET_NAME);
                     currentWorksheet.load('name'); // Load basic properties
                     financialsSheet.load('name');
                     await context.sync(); // Ensure sheets are loaded

                     console.log(`Successfully got references for ${currentWorksheet.name} and ${financialsSheet.name}`);

                     // 1. Get Last Row for the current assumption tab
                     // getLastUsedRow needs context, so call it here
                     const lastRow = await getLastUsedRow(currentWorksheet, "B"); // Pass currentWorksheet from this context
                     if (lastRow < START_ROW) {
                         console.warn(`Skipping tab ${worksheetName} as last used row in Col B (${lastRow}) is before start row (${START_ROW}).`);
                         return; // Skip to next tab if empty or too short
                     }
                     console.log(`Last row in Col B for ${worksheetName}: ${lastRow}`);

                     // --- Call helper functions IN ORDER, passing worksheet objects from this context ---
                     // These helpers now expect to run within this context

                     // 2. Adjust Drivers
                     await adjustDriversJS(currentWorksheet, lastRow);

                     // 3. Replace Indirects
                     await replaceIndirectsJS(currentWorksheet, lastRow);

                     // 4. Get Last Row Again (if Replace_Indirects might change it)
                     // const updatedLastRow = await getLastUsedRow(currentWorksheet, "B"); // Recalculate if necessary
                     const updatedLastRow = lastRow; // Assuming Replace_Indirects doesn't change last row for now
                     console.log(`Using last row for subsequent steps: ${updatedLastRow}`);
                     if (updatedLastRow < START_ROW) {
                         console.warn(`Skipping remaining steps for ${worksheetName} as updated last row (${updatedLastRow}) is invalid.`);
                         return;
                     }

                     // 5. Populate Financials
                     await populateFinancialsJS(currentWorksheet, updatedLastRow, financialsSheet);

                     // 6. Link Financial References
                     await linkFinReferencesJS(currentWorksheet, updatedLastRow, financialsSheet);

                     // 7. Autofill AE9:AE<lastRow> -> CX<lastRow> on Assumption Tab
                     console.log(`Autofilling ${AUTOFILL_START_COLUMN}${START_ROW}:${AUTOFILL_START_COLUMN}${updatedLastRow} to ${AUTOFILL_END_COLUMN} on ${worksheetName}`);
                     const sourceRange = currentWorksheet.getRange(`${AUTOFILL_START_COLUMN}${START_ROW}:${AUTOFILL_START_COLUMN}${updatedLastRow}`);
                     const fillRange = currentWorksheet.getRange(`${AUTOFILL_START_COLUMN}${START_ROW}:${AUTOFILL_END_COLUMN}${updatedLastRow}`);
                     sourceRange.autoFill(fillRange, Excel.AutoFillType.fillDefault);

                     // Sync all batched operations for this tab
                     await context.sync();
                     console.log(`Finished processing and syncing for tab ${worksheetName}`);

                 }); // End Excel.run for single tab processing

             } catch (tabError) {
                 console.error(`Error processing tab ${worksheetName}:`, tabError);
                 // Optionally add to an error list and continue with the next tab
                 // Be mindful that subsequent tabs might depend on this one succeeding.
             }
        } // --- End loop through assumption tabs ---

        // --- Final Operations on Financials Sheet ---
        console.log(`\nPerforming final operations on ${FINANCIALS_SHEET_NAME}`);
        try {
             await Excel.run(async (context) => {
                 const finSheet = context.workbook.worksheets.getItem(FINANCIALS_SHEET_NAME);
                 finSheet.load('name'); // Load name for logging
                 await context.sync(); // Ensure sheet is loaded

                 // 1. Format Changes In Working Capital
                 await formatChangesInWorkingCapitalJS(finSheet); // Pass sheet from this context

                 // 2. Get Last Row for Financials
                 const financialsLastRow = await getLastUsedRow(finSheet, "B"); // Pass sheet from this context
                 if (financialsLastRow < START_ROW) {
                     console.warn(`Skipping final autofill on ${FINANCIALS_SHEET_NAME} as last row (${financialsLastRow}) is before start row (${START_ROW}).`);
                     return;
                 }
                 console.log(`Last row in Col B for ${FINANCIALS_SHEET_NAME}: ${financialsLastRow}`);

                 // 3. Autofill AE9:AE<lastRow> -> CX<lastRow> on Financials
                 console.log(`Autofilling ${AUTOFILL_START_COLUMN}${START_ROW}:${AUTOFILL_START_COLUMN}${financialsLastRow} to ${AUTOFILL_END_COLUMN} on ${FINANCIALS_SHEET_NAME}`);
                 const sourceRangeFin = finSheet.getRange(`${AUTOFILL_START_COLUMN}${START_ROW}:${AUTOFILL_START_COLUMN}${financialsLastRow}`);
                 const fillRangeFin = finSheet.getRange(`${AUTOFILL_START_COLUMN}${START_ROW}:${AUTOFILL_END_COLUMN}${financialsLastRow}`);
                 sourceRangeFin.autoFill(fillRangeFin, Excel.AutoFillType.fillDefault);

                 // Commented out VBA Retained Earnings logic - implement if needed
                 // console.log("Finding rows for Retained Earnings calculation...");
                 // const retainedEarningsRow = await findRowByValue(context, finSheet, "B:B", "Retained Earnings"); // Requires findRowByValue helper
                 // const assetsRow = await findRowByValue(context, finSheet, "B:B", "Total Assets");
                 // const liabilitiesRow = await findRowByValue(context, finSheet, "B:B", "Total Liabilities");
                 // console.log(`Found Rows - RE: ${retainedEarningsRow}, Assets: ${assetsRow}, Liab: ${liabilitiesRow}`);
                 // if (retainedEarningsRow > 0 && assetsRow > 0 && liabilitiesRow > 0) {
                 //     const targetCell = finSheet.getRange(`AD${retainedEarningsRow}`); // Assuming AD is the target column
                 //     const formula = `=AD${assetsRow}-AD${liabilitiesRow}`;
                 //     console.log(`Setting Retained Earnings formula in AD${retainedEarningsRow}: ${formula}`);
                 //     targetCell.formulas = [[formula]];
                 // } else {
                 //     console.warn("Could not find one or more rows required for Retained Earnings calculation.");
                 // }

                 // Sync final Financials sheet operations
                 await context.sync();
                 console.log(`Finished final operations on ${FINANCIALS_SHEET_NAME}`);
             });
         } catch (financialsError) {
             console.error(`Error during final operations on ${FINANCIALS_SHEET_NAME}:`, financialsError);
         }

        console.log("Finished processing all assumption tabs.");

    } catch (error) {
        console.error("Error in processAssumptionTabs main function:", error);
        // Potentially re-throw or handle top-level errors
    }
}

// TODO: Implement the actual logic within the JS helper functions (adjustDriversJS, replaceIndirectsJS, etc.).
// TODO: Implement findRowByValue helper function if Retained Earnings logic is needed.
// TODO: Update the calling code (e.g., button handler in taskpane.js) to call `processAssumptionTabs` after `runCodes`.