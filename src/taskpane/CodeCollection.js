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
 * Toggles Excel calculation mode between Manual and Automatic.
 * @param {boolean} manual - If true, sets calculation to Manual; otherwise sets to Automatic.
 * @returns {Promise<void>}
 */
export async function toggleManualCalculation(manual) {
    try {
        await Excel.run(async (context) => {
            const application = context.workbook.application;
            
            // Set calculation mode
            application.calculationMode = manual 
                ? Excel.CalculationMode.manual 
                : Excel.CalculationMode.automatic;

            console.log(`Calculation mode set to: ${application.calculationMode}`);
            
            await context.sync();
        });
    } catch (error) {
        console.error(`Error toggling calculation mode: ${error}`);
        // Decide if this should throw or just log
        // throw error;
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
        
        // --- Set calculation to Manual before starting --- 
        await toggleManualCalculation(true);
        
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
        


        // --- Set calculation back to Automatic after finishing --- 
        await toggleManualCalculation(false);

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
 * A simple test function.
 */
export async function testTextFunction() {
    // Put your code to test here

    console.log("hello world");

    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        sheet.getRange("A1:B20").getLastRow().select();
        await context.sync();
    });

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
 * Placeholder for Adjust_Drivers VBA logic.
 * Finds driver codes in col F, looks up in col A, writes target address string in col AE.
 * @param {Excel.Worksheet} worksheet - The assumption worksheet (within an Excel.run context).
 * @param {number} lastRow - The last row to process.
 */
async function adjustDriversJS(worksheet, lastRow) {
    console.log(`Running adjustDriversJS for sheet: ${worksheet.name} up to row ${lastRow}`);
    // This function MUST be called within an Excel.run context that includes the worksheet object.
    // TODO: Implement Adjust_Drivers logic
    // 1. Load range F9:F<lastRow> and A9:A<lastRow>
    // 2. Create map of Col A values to row numbers
    // 3. Iterate F column values
    // 4. If value exists, find corresponding row in map
    // 5. Prepare target address string (e.g., "AE" + foundRow)
    // 6. Write strings to AE9:AE<lastRow>
    worksheet.load('name'); // Keep worksheet reference valid if needed later in the SAME context
    await worksheet.context.sync(); // Sync actions within the caller's context
    console.warn(`adjustDriversJS on ${worksheet.name} not implemented yet.`);

}

/**
 * Placeholder for Replace_Indirects VBA logic.
 * Replaces INDIRECT functions in column AE with their evaluated values.
 * @param {Excel.Worksheet} worksheet - The assumption worksheet (within an Excel.run context).
 * @param {number} lastRow - The last row to process.
 */
async function replaceIndirectsJS(worksheet, lastRow) {
    console.log(`Running replaceIndirectsJS for sheet: ${worksheet.name} up to row ${lastRow}`);
    // This function MUST be called within an Excel.run context.
    // TODO: Implement Replace_Indirects logic (Complex)
    // 1. Load formulas from AE9:AE<lastRow>
    // 2. Use regex to find INDIRECT(...)
    // 3. For each match:
    //    a. Extract argument string
    //    b. Get range for argument string: worksheet.getRange(argString)
    //    c. Load value of the argument range (may need sync within loop or batch requests)
    //    d. Replace INDIRECT(...) in formula string with the value (handle "DELETE")
    // 4. Write modified formulas back to AE9:AE<lastRow>
    worksheet.load('name');
    await worksheet.context.sync();
     console.warn(`replaceIndirectsJS on ${worksheet.name} not implemented yet.`);

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
    const CALCS_FIRST_ROW = 9; // Assuming same as VBA calcsfirstrow
    const FINANCIALS_CODE_COLUMN = "I";
    const FINANCIALS_TARGET_COLUMN_B = "B";
    const FINANCIALS_TARGET_COLUMN_D = "D";
    const FINANCIALS_TARGET_COLUMN_J = "J";
     // Add other constants as needed (K, month start/end, actual start/end, annual end)

    // TODO: Implement Populate_Financials logic
    // 1. Define constants (column letters, row numbers)
    // 2. Load codes from C<CALCS_FIRST_ROW>:C<lastRow> (assumption sheet)
    // 3. Load codes from FINANCIALS_CODE_COLUMN (Financials sheet, use getUsedRange().getIntersection)
    // 4. Create map of Financials codes to row numbers
    // 5. Iterate assumption sheet codes (row `i`)
    // 6. If code found in Financials map (pasterow):
    //    a. Store insertion task { targetRow: pasterow, assumptionRow: i, code: code }
    // 7. Sort tasks by targetRow DESCENDING (crucial for correct insertion indices)
    // 8. Perform insertions in order (requires careful context.sync management)
    // 9. Iterate original tasks (adjusting pasterow for insertions)
    //    a. Populate B, D (links)
    //    b. Populate J, K (SUMIF - adapt INDIRECT?)
    //    c. Populate actual/month columns (SUMIFS, links)
    //    d. Set formats
    //    e. Perform autofills (might need separate syncs)
    worksheet.load('name'); // Keep references valid if needed later in the SAME context
    financialsSheet.load('name');
    await worksheet.context.sync();
     console.warn(`populateFinancialsJS for ${worksheet.name} -> ${financialsSheet.name} not implemented yet.`);

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