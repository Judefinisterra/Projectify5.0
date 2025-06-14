/**
 * CodeCollection.js
 * Functions for processing and managing code collections
 */

import { convertKeysToCamelCase } from "@pinecone-database/pinecone/dist/utils";

// >>> ADDED: Import the logic validation function
import { validateLogicOnly } from './Validation.js';
// >>> ADDED: Import the format validation function
import { validateFormatErrors } from './ValidationFormat.js';

/*
 * LOGIC VALIDATION INTEGRATION GUIDE
 * 
 * This module provides logic validation functions to detect errors before LogicGPT API calls.
 * The logic validation checks for:
 * - Invalid code types
 * - Duplicate row drivers within tabs
 * - Invalid financial statement codes
 * - Duplicate financial statement items
 * - Custom formula syntax errors
 * - Driver references not found
 * - Invalid row formats
 * - TAB validation errors
 * 
 * INTEGRATION WORKFLOW:
 * 
 * 1. Basic Usage (Formatted for GPT prompt):
 *    const logicErrors = await getLogicErrorsForPrompt(inputCodeStrings);
 *    const enhancedPrompt = originalPrompt + logicErrors;
 *    // Call LogicGPT API with enhancedPrompt
 * 
 * 2. Conditional Processing:
 *    if (await hasLogicErrors(inputCodeStrings)) {
 *        const logicErrors = await getLogicErrorsForPrompt(inputCodeStrings);
 *        // Include errors in LogicGPT prompt
 *    } else {
 *        // Proceed with normal LogicGPT call
 *    }
 * 
 * 3. Custom Error Processing:
 *    const rawErrors = await getLogicErrors(inputCodeStrings);
 *    // Process errors as needed for custom formatting
 * 
 * NOTE: This validation focuses on LOGIC errors only (LERR codes), not format errors (FERR codes).
 * Format validation is handled separately in the full validation system.
 */

/*
 * FORMAT VALIDATION INTEGRATION GUIDE
 * 
 * This module provides format validation functions to detect errors before FormatGPT API calls.
 * The format validation checks for:
 * - LABELH1/H2/H3 column formatting rules
 * - Indent and bold parameter requirements
 * - Adjacent code formatting conflicts
 * - BR code adjacency rules
 * - Total row formatting requirements
 * 
 * INTEGRATION WORKFLOW:
 * 
 * 1. Basic Usage (Formatted for GPT prompt):
 *    const formatErrors = await getFormatErrorsForPrompt(inputCodeStrings);
 *    const enhancedPrompt = originalPrompt + formatErrors;
 *    // Call FormatGPT API with enhancedPrompt
 * 
 * 2. Conditional Processing:
 *    if (await hasFormatErrors(inputCodeStrings)) {
 *        const formatErrors = await getFormatErrorsForPrompt(inputCodeStrings);
 *        // Include errors in FormatGPT prompt
 *    } else {
 *        // Proceed with normal FormatGPT call
 *    }
 * 
 * 3. Custom Error Processing:
 *    const rawErrors = await getFormatErrors(inputCodeStrings);
 *    // Process errors as needed for custom formatting
 * 
 * NOTE: This validation focuses on FORMAT errors only (FERR codes), not logic errors (LERR codes).
 */

// <<< END ADDED

/**
 * Runs logic validation on input text and returns formatted errors for LogicGPT prompt inclusion
 * @param {string} inputText - The input text containing code strings
 * @returns {Promise<string>} - Formatted logic errors string for prompt inclusion
 * 
 * USAGE EXAMPLE:
 * // Before making LogicGPT API call:
 * const logicErrors = await getLogicErrorsForPrompt(inputCodeStrings);
 * const enhancedPrompt = originalPrompt + logicErrors;
 * // Then call LogicGPT API with enhancedPrompt
 */
export async function getLogicErrorsForPrompt(inputText) {
    try {
        console.log("╔" + "═".repeat(78) + "╗");
        console.log("║ [CodeCollection] LOGIC VALIDATION FOR GPT PROMPT                          ║");
        console.log("╚" + "═".repeat(78) + "╝");
        console.log("[CodeCollection] Starting logic validation for LogicGPT prompt enhancement...");
        console.log("[CodeCollection] Input type:", typeof inputText);
        console.log("[CodeCollection] Input size:", inputText?.length || 0, "characters");
        
        // Run logic-only validation
        const logicErrors = await validateLogicOnly(inputText);
        
        if (logicErrors.length === 0) {
            console.log("[CodeCollection] ✅ No logic errors found - prompt will not be enhanced");
            console.log("[CodeCollection] Logic validation complete - returning empty string");
            return ""; // Return empty string if no errors
        }
        
        console.log(`[CodeCollection] ⚠️  Found ${logicErrors.length} logic errors to include in prompt`);
        
        // Format errors for GPT prompt inclusion
        let formattedErrors = "\n\nLOGIC VALIDATION ERRORS DETECTED:\n";
        formattedErrors += "Please address the following logic errors in your response:\n\n";
        
        logicErrors.forEach((error, index) => {
            formattedErrors += `${index + 1}. ${error}\n`;
        });
        
        formattedErrors += "\nPlease ensure your corrected codestrings resolve these logic issues.\n";
        
        console.log("[CodeCollection] Formatted error prompt enhancement:");
        console.log(formattedErrors);
        console.log("[CodeCollection] Logic validation complete - returning formatted errors for prompt");
        
        return formattedErrors;
        
    } catch (error) {
        console.error("[CodeCollection] ❌ Error during logic validation:", error);
        console.error("[CodeCollection] Returning error message for prompt inclusion");
        return "\n\nLOGIC VALIDATION ERROR: Could not complete logic validation due to technical error.\n";
    }
}

/**
 * Gets raw logic errors array for custom processing
 * @param {string} inputText - The input text containing code strings
 * @returns {Promise<Array>} - Array of logic error strings
 */
export async function getLogicErrors(inputText) {
    try {
        console.log("[CodeCollection] Getting raw logic errors array...");
        console.log("[CodeCollection] Input type:", typeof inputText);
        console.log("[CodeCollection] Input size:", inputText?.length || 0, "characters");
        
        const errors = await validateLogicOnly(inputText);
        
        console.log(`[CodeCollection] Raw logic validation complete - returning ${errors.length} errors`);
        return errors;
    } catch (error) {
        console.error("[CodeCollection] ❌ Error getting logic errors:", error);
        const errorArray = [`Logic validation error: ${error.message}`];
        console.log("[CodeCollection] Returning error array with 1 error message");
        return errorArray;
    }
}

/**
 * Checks if there are any logic errors without returning details
 * @param {string} inputText - The input text containing code strings
 * @returns {Promise<boolean>} - True if logic errors exist, false otherwise
 */
export async function hasLogicErrors(inputText) {
    try {
        console.log("[CodeCollection] Checking if logic errors exist...");
        console.log("[CodeCollection] Input type:", typeof inputText);
        console.log("[CodeCollection] Input size:", inputText?.length || 0, "characters");
        
        const errors = await validateLogicOnly(inputText);
        const hasErrors = errors.length > 0;
        
        console.log(`[CodeCollection] Logic error check complete - result: ${hasErrors ? '❌ Has errors' : '✅ No errors'} (${errors.length} errors found)`);
        return hasErrors;
    } catch (error) {
        console.error("[CodeCollection] ❌ Error checking for logic errors:", error);
        console.log("[CodeCollection] Assuming errors exist due to validation failure");
        return true; // Assume errors exist if validation fails
    }
}

/**
 * Runs format validation on input text and returns formatted errors for FormatGPT prompt inclusion
 * @param {string} inputText - The input text containing code strings
 * @returns {Promise<string>} - Formatted format errors string for prompt inclusion
 * 
 * USAGE EXAMPLE:
 * // Before making FormatGPT API call:
 * const formatErrors = await getFormatErrorsForPrompt(inputCodeStrings);
 * const enhancedPrompt = originalPrompt + formatErrors;
 * // Then call FormatGPT API with enhancedPrompt
 */
export async function getFormatErrorsForPrompt(inputText) {
    try {
        console.log("╔" + "═".repeat(78) + "╗");
        console.log("║ [CodeCollection] FORMAT VALIDATION FOR GPT PROMPT                         ║");
        console.log("╚" + "═".repeat(78) + "╝");
        console.log("[CodeCollection] Starting format validation for FormatGPT prompt enhancement...");
        console.log("[CodeCollection] Input type:", typeof inputText);
        console.log("[CodeCollection] Input size:", inputText?.length || 0, "characters");
        
        // Run format-only validation
        const formatErrors = await validateFormatErrors(inputText);
        
        if (formatErrors.length === 0) {
            console.log("[CodeCollection] ✅ No format errors found - prompt will not be enhanced");
            console.log("[CodeCollection] Format validation complete - returning empty string");
            return ""; // Return empty string if no errors
        }
        
        console.log(`[CodeCollection] ⚠️  Found ${formatErrors.length} format errors to include in prompt`);
        
        // Format errors for GPT prompt inclusion
        let formattedErrors = "\n\nFORMAT VALIDATION ERRORS DETECTED:\n";
        formattedErrors += "Please address the following format errors in your response:\n\n";
        
        formatErrors.forEach((error, index) => {
            formattedErrors += `${index + 1}. ${error}\n`;
        });
        
        formattedErrors += "\nPlease ensure your corrected codestrings resolve these format issues.\n";
        
        console.log("[CodeCollection] Formatted error prompt enhancement:");
        console.log(formattedErrors);
        console.log("[CodeCollection] Format validation complete - returning formatted errors for prompt");
        
        return formattedErrors;
        
    } catch (error) {
        console.error("[CodeCollection] ❌ Error during format validation:", error);
        console.error("[CodeCollection] Returning error message for prompt inclusion");
        return "\n\nFORMAT VALIDATION ERROR: Could not complete format validation due to technical error.\n";
    }
}

/**
 * Gets raw format errors array for custom processing
 * @param {string} inputText - The input text containing code strings
 * @returns {Promise<Array>} - Array of format error strings
 */
export async function getFormatErrors(inputText) {
    try {
        console.log("[CodeCollection] Getting raw format errors array...");
        console.log("[CodeCollection] Input type:", typeof inputText);
        console.log("[CodeCollection] Input size:", inputText?.length || 0, "characters");
        
        const errors = await validateFormatErrors(inputText);
        
        console.log(`[CodeCollection] Raw format validation complete - returning ${errors.length} errors`);
        return errors;
    } catch (error) {
        console.error("[CodeCollection] ❌ Error getting format errors:", error);
        const errorArray = [`Format validation error: ${error.message}`];
        console.log("[CodeCollection] Returning error array with 1 error message");
        return errorArray;
    }
}

/**
 * Checks if there are any format errors without returning details
 * @param {string} inputText - The input text containing code strings
 * @returns {Promise<boolean>} - True if format errors exist, false otherwise
 */
export async function hasFormatErrors(inputText) {
    try {
        console.log("[CodeCollection] Checking if format errors exist...");
        console.log("[CodeCollection] Input type:", typeof inputText);
        console.log("[CodeCollection] Input size:", inputText?.length || 0, "characters");
        
        const errors = await validateFormatErrors(inputText);
        const hasErrors = errors.length > 0;
        
        console.log(`[CodeCollection] Format error check complete - result: ${hasErrors ? '❌ Has errors' : '✅ No errors'} (${errors.length} errors found)`);
        return hasErrors;
    } catch (error) {
        console.error("[CodeCollection] ❌ Error checking for format errors:", error);
        console.log("[CodeCollection] Assuming errors exist due to validation failure");
        return true; // Assume errors exist if validation fails
    }
}

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
        
        // Extract all code strings using regex pattern /<[^>]+>/g
        // This allows multiple code strings on the same line
        const codeStringMatches = inputText.match(/<[^>]+>/g);
        
        if (!codeStringMatches) {
            console.log("No code strings found in input text");
            return codeCollection;
        }
        
        console.log(`Found ${codeStringMatches.length} code strings to process`);
        
        for (const codeString of codeStringMatches) {
            // Skip empty code strings
            if (!codeString.trim()) continue;
            
            // Extract content within < >
            const content = codeString.substring(1, codeString.length - 1);
            
            // Extract the code type and parameters
            const codeMatch = content.match(/^([^;>]+);(.*?)$/);
            if (!codeMatch) {
                console.warn(`Skipping malformed code string: ${codeString}`);
                continue;
            }
            
            const codeType = codeMatch[1].trim();
            // Remove potential leftover newline/carriage return characters from the params string
            const paramsString = codeMatch[2].replace(/[\r\n]+/g, '').trim();
            
            // Parse parameters
            const params = {};
            
            // Handle special case for row parameters with asterisks
            // Use the cleaned paramsString
            const rowMatches = paramsString.matchAll(/row(\d+)\s*=\s*"([^"]*)"/g);
            for (const match of rowMatches) {
                const rowNum = match[1];
                const originalRowValue = match[2];
                let rowValue = originalRowValue;
                
                // Remove column identifiers in parentheses (e.g., (D), (L), (C1), (Y1), etc.)
                rowValue = rowValue.replace(/\([^)]*\)/g, '');
                
                // Console log the cleaning process
                if (originalRowValue !== rowValue) {
                    console.log(`Cleaned row${rowNum}:`);
                    console.log(`  Before: "${originalRowValue}"`);
                    console.log(`  After:  "${rowValue}"`);
                }
                
                params[`row${rowNum}`] = rowValue;
            }
            
            // Parse other parameters, including the new "format" parameter
            // Use the cleaned paramsString
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
                            // if (existingSheet) {
                            //     // Delete the worksheet if it exists
                            //     existingSheet.delete();
                            //     await context.sync();
                            // }
                            // console.log("existingSheet deleted");
                            
                            // Get the Financials worksheet (needed for position and as fallback template)
                            const financialsSheet = context.workbook.worksheets.getItem("Financials");
                            financialsSheet.load("position"); // Load Financials sheet position
                            await context.sync(); // Sync to get Financials position
                            console.log(`Financials sheet is at position ${financialsSheet.position}`);
                            
                            // Check if the target tab already exists
                            if (!existingSheet) {
                                let newSheet;
                                let sourceSheetName;

                                // Try to get the Calcs worksheet
                                try {
                                    const sourceCalcsWS = context.workbook.worksheets.getItem("Calcs");
                                    await context.sync(); // Ensure it's loaded if found
                                    console.log("Using Calcs worksheet as template.");
                                    console.log(`[SHEET OPERATION] Copying sheet 'Calcs' to create new sheet for tab '${tabName}'`);
                                    newSheet = sourceCalcsWS.copy();
                                    sourceSheetName = "Calcs";
                                } catch (calcsError) {
                                    // If Calcs doesn't exist, use Financials as the template
                                    console.warn("Calcs worksheet not found. Using Financials as template.");
                                    console.log(`[SHEET OPERATION] Copying sheet 'Financials' to create new sheet for tab '${tabName}'`);
                                    newSheet = financialsSheet.copy();
                                    sourceSheetName = "Financials";
                                    // Sync needed *after* copy to reference the new sheet object reliably
                                    await context.sync(); 
                                    
                                    // --- Load name before accessing it ---
                                    newSheet.load("name");
                                    await context.sync();
                                    // --- End Load name ---
                                    
                                    // --- Clear rows 10 down if copied from Financials ---
                                    console.log(`Clearing contents and formats from row 10 down in new sheet ${newSheet.name} copied from ${sourceSheetName}`);
                                    // Use a reasonable large row number or get last row if needed, 10000 should suffice
                                    const clearRange = newSheet.getRange("10:10000"); 
                                    clearRange.clear(Excel.ClearApplyTo.all);
                                    // Do NOT sync clear yet, batch with linking below

                                    // --- Link non-empty cells in rows 1-8 back to Financials ---
                                    console.log(`Linking header rows (1-8) in ${newSheet.name} back to Financials`);
                                    // Get used range of the new sheet to find last column
                                    const usedRange = newSheet.getUsedRange(true); // Use valuesOnly = true
                                    usedRange.load(["columnCount", "rowCount"]);
                                    // Sync to get the used range info *before* calculating link range address
                                    await context.sync();

                                    const lastColIndex = usedRange.columnCount > 0 ? usedRange.columnCount - 1 : 0; 
                                    const lastColLetter = columnIndexToLetter(lastColIndex);
                                    // Process only up to row 8
                                    const linkRangeAddress = `A1:${lastColLetter}8`;

                                    console.log(`Processing header link range: ${linkRangeAddress}`);
                                    const linkRange = newSheet.getRange(linkRangeAddress);
                                    linkRange.load("values");
                                    // Sync to load the values *before* iterating and setting formulas
                                    await context.sync();

                                    const values = linkRange.values;
                                    // Batch formula setting directly
                                    for (let r = 0; r < values.length; r++) {
                                        const rowNum = r + 1;
                                        for (let c = 0; c < values[r].length; c++) {
                                            const cellValue = values[r][c];
                                            if (cellValue !== null && cellValue !== "") {
                                                const colLetter = columnIndexToLetter(c);
                                                const cellAddress = `${colLetter}${rowNum}`;
                                                const formula = `=Financials!${cellAddress}`;
                                                // Get the specific cell and queue the formula update
                                                const targetCell = newSheet.getRange(cellAddress);
                                                targetCell.formulas = [[formula]];
                                                // console.log(`  Queueing formula for ${cellAddress} to ${formula}`); 
                                            }
                                        }
                                    }
                                    // The sync for these formula changes will happen later, along with rename/position.
                                    // --- End Link header rows ---

                                    // --- Set font color for rows 2-8 ---
                                    console.log(`Setting font color for rows 2-8 in ${newSheet.name}`);
                                    const headerFormatRangeAddress = `A2:${lastColLetter}8`;
                                    const headerFormatRange = newSheet.getRange(headerFormatRangeAddress);
                                    headerFormatRange.format.font.color = "#008000"; // Green
                                    // --- End Set font color ---

                                    // --- Set tab color ---
                                    console.log(`Setting tab color for ${newSheet.name}`);
                                    newSheet.tabColor = "#4472C4"; // Blue
                                    // --- End Set tab color ---
                                }

                                // Sync copy operation if not already synced (e.g., if Calcs was used)
                                // If Financials was used, sync happened before clear. If Calcs was used, sync happens here.
                                if (sourceSheetName === "Calcs") {
                                     await context.sync();
                                }

                                console.log(`newSheet created by copying ${sourceSheetName} worksheet`);

                                // Rename it
                                console.log(`[SHEET OPERATION] Renaming copied sheet to '${tabName}'`);
                                newSheet.name = tabName;
                                console.log("newSheet renamed to", tabName);

                                // <<< NEW: Set position relative to Financials sheet >>>
                                newSheet.position = financialsSheet.position + 1;
                                console.log(`Set position of ${tabName} to ${newSheet.position}`);
                                // Add to assumption tabs collection
                                assumptionTabs.push({
                                    name: tabName,
                                    worksheet: newSheet
                                }); // <-- Added closing brace and semicolon here

                                currentWorksheetName = tabName;

                                await context.sync(); // Sync rename and position changes

                                result.createdTabs.push(tabName);
                                console.log("Tab created successfully:", tabName);
                            // }); <-- Removed this closing parenthesis, it belongs to Excel.run below

                            }

                            else {
                                console.log("Worksheet already exists:", tabName);
                                console.log(`[SHEET OPERATION] No sheet creation needed - tab '${tabName}' already exists`);
                                assumptionTabs.push({
                                    name: tabName,
                                    worksheet: existingSheet
                                });
                                // Need to set currentWorksheetName here too if the sheet exists
                                currentWorksheetName = tabName; 
                            }
                      
       
                            
                         
                            
                            // Set the current worksheet name <-- This comment is now redundant/misplaced
                       
                        } catch (error) {
                            console.error("Detailed error in TAB processing:", error);
                            throw error;
                        }
                    }).catch(error => { // <-- This is the correct closing for Excel.run
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
                            const pasteRow = Math.max(lastUsedRow.rowIndex + 2, 10); // Ensure paste starts at row 10 or later
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
                                
                                // Get the source range from codesWS (already available in this context)
                                const sourceRange = codesWS.getRange(`A${firstRow}:CX${lastRow}`);
                                
                                // Get the destination range in currentWS (already available in this context)
                                const destinationRange = currentWS.getRange(`A${pasteRow}`);
                                
                                // Copy the range with all properties
                                destinationRange.copyFrom(sourceRange, Excel.RangeCopyType.all);
                                
                                await context.sync(); // Sync the copy operation

                                // NEW: Apply bold formatting if specified
                                if (code.params.bold !== undefined) { // Check if the parameter exists
                                    const boldValue = String(code.params.bold).toLowerCase();
                                    const numPastedRows = lastRow - firstRow + 1;
                                    // Ensure endPastedRow is at least pasteRow and calculated correctly
                                    const endPastedRow = pasteRow + Math.max(0, numPastedRows - 1);
                                    
                                    // Assuming CX is a sufficiently wide column, as used in the copy.
                                    const rangeAddressToBold = `A${pasteRow}:CX${endPastedRow}`;
                                    const rangeToBold = currentWS.getRange(rangeAddressToBold);
                                    
                                    console.log(`Processing "bold" parameter: "${boldValue}" for range ${rangeAddressToBold}`);

                                    if (boldValue === "true") {
                                        console.log(`Applying bold formatting to ${rangeAddressToBold} in ${currentWorksheetName} for code ${codeType}`);
                                        rangeToBold.format.font.bold = true;
                                        await context.sync(); // Sync the bold formatting
                                        console.log(`Bold formatting applied and synced for ${rangeAddressToBold}`);
                                    } else if (boldValue === "false") {
                                        console.log(`Removing bold formatting from ${rangeAddressToBold} in ${currentWorksheetName} for code ${codeType}`);
                                        rangeToBold.format.font.bold = false;
                                        await context.sync(); // Sync the bold formatting removal
                                        console.log(`Bold formatting removed and synced for ${rangeAddressToBold}`);
                                    } else {
                                        console.log(`"bold" parameter value "${boldValue}" is not recognized as boolean. No bold formatting change applied.`);
                                    }
                                }

                                // NEW: Apply top border formatting if specified
                                if (code.params.topborder && String(code.params.topborder).toUpperCase() === "TRUE") {
                                    const numPastedRows = lastRow - firstRow + 1;
                                    const endPastedRow = pasteRow + Math.max(0, numPastedRows - 1);

                                    console.log(`Applying top border to J${pasteRow}:P${endPastedRow} and S${pasteRow}:CX${endPastedRow} in ${currentWorksheetName} for code ${codeType}`);

                                    for (let r = pasteRow; r <= endPastedRow; r++) {
                                        const rangeJtoP = currentWS.getRange(`J${r}:P${r}`);
                                        rangeJtoP.format.borders.getItem('EdgeTop').style = 'Continuous';
                                        rangeJtoP.format.borders.getItem('EdgeTop').weight = 'Thin';
                                        // Color defaults to Automatic (usually black)

                                        const rangeStoCX = currentWS.getRange(`S${r}:CX${r}`);
                                        rangeStoCX.format.borders.getItem('EdgeTop').style = 'Continuous';
                                        rangeStoCX.format.borders.getItem('EdgeTop').weight = 'Thin';
                                    }
                                    await context.sync(); // Sync the top border formatting
                                    console.log(`Top border formatting applied and synced for J${pasteRow}:P${endPastedRow} and S${pasteRow}:CX${endPastedRow}`);
                                }
                                
                                // NEW: Apply indent formatting if specified
                                if (code.params.indent) {
                                    const indentValue = parseInt(code.params.indent, 10);
                                    if (!isNaN(indentValue) && indentValue > 0) {
                                        const numPastedRows = lastRow - firstRow + 1;
                                        const endPastedRow = pasteRow + Math.max(0, numPastedRows - 1);
                                        const indentRangeAddress = `B${pasteRow}:B${endPastedRow}`;

                                        console.log(`Applying indent of ${indentValue} to ${indentRangeAddress} in ${currentWorksheetName} for code ${codeType}`);
                                        const rangeToIndent = currentWS.getRange(indentRangeAddress);
                                        rangeToIndent.format.indentLevel = indentValue;
                                        await context.sync(); // Sync the indent formatting
                                        console.log(`Indent formatting applied and synced for ${indentRangeAddress}`);
                                    } else {
                                        console.warn(`Invalid indent value: "${code.params.indent}" for code ${codeType}. Indent must be a positive integer.`);
                                    }
                                }

                                // NEW: Apply negative formatting if specified
                                if (code.params.negative && String(code.params.negative).toUpperCase() === "TRUE") {
                                    const numPastedRows = lastRow - firstRow + 1;
                                    const endPastedRow = pasteRow + Math.max(0, numPastedRows - 1); // This is the last row of the pasted block

                                    console.log(`Applying negative transformation to formulas in AE${endPastedRow}:CX${endPastedRow} for code ${codeType}`);
                                    const formulaRange = currentWS.getRange(`AE${endPastedRow}:CX${endPastedRow}`);
                                    formulaRange.load("formulas");
                                    await context.sync();

                                    const originalFormulasRow = formulaRange.formulas[0]; // Get the single row of formulas
                                    const newFormulasRow = [];
                                    let formulasChanged = false;

                                    for (let i = 0; i < originalFormulasRow.length; i++) {
                                        const currentCellFormula = originalFormulasRow[i];
                                        if (typeof currentCellFormula === 'string' && currentCellFormula.startsWith('=')) {
                                            // Construct the new formula: =-(original_content)
                                            newFormulasRow.push(`=-(${currentCellFormula.substring(1)})`);
                                            formulasChanged = true;
                                        } else {
                                            newFormulasRow.push(currentCellFormula); // Keep non-formulas or empty strings as is
                                        }
                                    }

                                    if (formulasChanged) {
                                        formulaRange.formulas = [newFormulasRow]; // Set as a 2D array
                                        await context.sync();
                                        console.log(`Negative transformation applied and synced for AE${endPastedRow}:CX${endPastedRow}`);
                                    } else {
                                        console.log(`No formulas found to transform in AE${endPastedRow}:CX${endPastedRow}`);
                                    }
                                }
                                
                                // NEW: Apply "format" parameter for number formatting and italics
                                if (code.params.format) {
                                    const formatValue = String(code.params.format).toLowerCase();
                                    const numPastedRows = lastRow - firstRow + 1;
                                    const endPastedRow = pasteRow + Math.max(0, numPastedRows - 1);
                                    // Apply to J through CX (including column J now)
                                    const formatRangeAddress = `J${pasteRow}:CX${endPastedRow}`;
                                    const rangeToFormat = currentWS.getRange(formatRangeAddress);
                                    let numberFormatString = null;
                                    // Removed applyItalics variable as direct checks on formatValue are clearer for B:CX range

                                    console.log(`Processing "format" parameter: "${formatValue}" for range ${formatRangeAddress}`);

                                    if (formatValue === "dollar" || formatValue === "dollaritalic") {
                                        numberFormatString = '_(* $ #,##0_);_(* $ (#,##0);_(* "$" -""?_);_(@_)';
                                    } else if (formatValue === "volume") {
                                        numberFormatString = '_(* #,##0_);_(* (#,##0);_(* " -"?_);_(@_)';
                                    } else if (formatValue === "percent") {
                                        numberFormatString = '_(* #,##0.0%;_(* (#,##0.0)%;_(* " -"?_)';
                                    } else if (formatValue === "factor") {
                                        numberFormatString = '_(* #,##0.0x;_(* (#,##0.0)x;_(* "   -"?_)';
                                    } else if (formatValue === "date") {
                                        numberFormatString = 'mmm-yy';
                                    }

                                    if (numberFormatString) {
                                        console.log(`Applying number format: "${numberFormatString}" to ${formatRangeAddress}`); // J:CX
                                        rangeToFormat.numberFormat = [[numberFormatString]]; // J:CX
                                        
                                        // Italicization logic based on formatValue for the B:CX range
                                        const fullItalicRangeAddress = `B${pasteRow}:CX${endPastedRow}`;
                                        const fullRangeToHandleItalics = currentWS.getRange(fullItalicRangeAddress);

                                        if (formatValue === "dollaritalic" || formatValue === "volume" || formatValue === "percent" || formatValue === "factor") {
                                            console.log(`Applying italics to ${fullItalicRangeAddress} due to format type (${formatValue})`);
                                            fullRangeToHandleItalics.format.font.italic = true;
                                        } else if (formatValue === "dollar" || formatValue === "date") {
                                            console.log(`Ensuring ${fullItalicRangeAddress} is NOT italicized due to format type (${formatValue})`);
                                            fullRangeToHandleItalics.format.font.italic = false;
                                        } else {
                                            // For unrecognized formats that still had a numberFormatString (e.g. if logic changes later),
                                            // or if K:CX needs explicit non-italic default when no B:CX rule applies.
                                            // However, current logic implies if numberFormatString is set, formatValue is one of the known ones.
                                            // If K:CX (rangeToFormat) needs specific non-italic handling for other cases, it would go here.
                                            // For now, this 'else' might not be hit if numberFormatString implies a known formatValue.
                                            // The primary `italic` parameter handles general italic override later anyway.
                                            console.log(`Format type ${formatValue} has number format but no specific B:CX italic rule. K:CX italics remain as previously set or default.`);
                                        }
                                        await context.sync();
                                        console.log(`"format" parameter processing (number format and B:CX italics) synced for ${formatRangeAddress}`);
                                    } else {
                                        console.log(`"format" parameter value "${formatValue}" is not recognized. No formatting applied.`);
                                    }
                                }

                                // NEW: Apply "italic" parameter for font style
                                if (code.params.italic !== undefined) { // Check if the parameter exists
                                    const italicValue = String(code.params.italic).toLowerCase();
                                    const numPastedRows = lastRow - firstRow + 1;
                                    const endPastedRow = pasteRow + Math.max(0, numPastedRows - 1);
                                    const italicRangeAddress = `B${pasteRow}:CX${endPastedRow}`;
                                    const rangeToItalicize = currentWS.getRange(italicRangeAddress);

                                    console.log(`Processing "italic" parameter: "${italicValue}" for range ${italicRangeAddress}`);

                                    if (italicValue === "true") {
                                        console.log(`Applying italics to ${italicRangeAddress}`);
                                        rangeToItalicize.format.font.italic = true;
                                        await context.sync();
                                        console.log(`"italic" parameter (true) processing synced for ${italicRangeAddress}`);
                                    } else if (italicValue === "false") {
                                        console.log(`Removing italics from ${italicRangeAddress}`);
                                        rangeToItalicize.format.font.italic = false;
                                        await context.sync();
                                        console.log(`"italic" parameter (false) processing synced for ${italicRangeAddress}`);
                                    } else {
                                        console.log(`"italic" parameter value "${italicValue}" is not recognized as boolean. No italicization change applied.`);
                                    }
                                }

                                // Apply the driver and assumption inputs function to the current worksheet
                                try {
                                    console.log(`Applying driver and assumption inputs to worksheet: ${currentWorksheetName}`);
                                    
                                    // Get the current worksheet and load its properties
                                    const currentWorksheet = context.workbook.worksheets.getItem(currentWorksheetName);
                                    currentWorksheet.load('name');
                                    await context.sync();
                                    
                                    const rowInfo = await driverAndAssumptionInputs(
                                        currentWorksheet,
                                        pasteRow,
                                        code
                                    );
                                    console.log(`Successfully applied driver and assumption inputs to worksheet: ${currentWorksheetName}`);
                                    
                                    // Store the actual row information for use by FIXED-E/FIXED-S
                                    code.actualRowInfo = rowInfo;
                                } catch (error) {
                                    console.error(`Error applying driver and assumption inputs: ${error.message}`);
                                    result.errors.push({
                                        codeIndex: i,
                                        codeType: codeType,
                                        error: `Error applying driver and assumption inputs: ${error.message}`
                                    });
                                }

                                // NEW: Apply "sumif" parameter for custom SUMIF/AVERAGEIF formulas in columns J-P
                                // NOTE: This must be AFTER driverAndAssumptionInputs to avoid being overwritten
                                if (code.params.sumif !== undefined) {
                                    const sumifValue = String(code.params.sumif).toLowerCase();
                                    const numPastedRows = lastRow - firstRow + 1;
                                    const endPastedRow = pasteRow + Math.max(0, numPastedRows - 1);
                                    const sumifRangeAddress = `J${pasteRow}:P${endPastedRow}`;
                                    const rangeToModify = currentWS.getRange(sumifRangeAddress);

                                    console.log(`Processing "sumif" parameter: "${sumifValue}" for range ${sumifRangeAddress}`);

                                    let formulaTemplate = null;
                                    if (sumifValue === "year") {
                                        formulaTemplate = "=SUMIF($3:$3, INDIRECT(ADDRESS(2,COLUMN(),2)), INDIRECT(ROW() & \":\" & ROW()))";
                                    } else if (sumifValue === "yearend") {
                                        formulaTemplate = "=SUMIF($4:$4, INDIRECT(ADDRESS(2,COLUMN(),2)), INDIRECT(ROW() & \":\" & ROW()))";
                                    } else if (sumifValue === "average") {
                                        formulaTemplate = "=AVERAGEIF($3:$3, INDIRECT(ADDRESS(2,COLUMN(),2)), INDIRECT(ROW() & \":\" & ROW()))";
                                    } else if (sumifValue === "offsetyear") {
                                        formulaTemplate = "=INDEX(INDIRECT(ROW() & \":\" & ROW()),1,MATCH(INDIRECT(ADDRESS(2,COLUMN()-1,2)),$4:$4,0)+1)";
                                    }

                                    if (formulaTemplate) {
                                        console.log(`Applying ${sumifValue} formula template to ${sumifRangeAddress}: ${formulaTemplate}`);
                                        
                                        if (sumifValue === "offsetyear") {
                                            // Special handling for offsetyear: set column J to 0, apply formula to K-P
                                            const columnJRange = currentWS.getRange(`J${pasteRow}:J${endPastedRow}`);
                                            const columnKPRange = currentWS.getRange(`K${pasteRow}:P${endPastedRow}`);
                                            
                                            // Set column J to 0
                                            const zeroArray = [];
                                            for (let r = 0; r < numPastedRows; r++) {
                                                zeroArray.push([0]);
                                            }
                                            columnJRange.values = zeroArray;
                                            
                                            // Apply formula to columns K-P (6 columns)
                                            const formulaArray = [];
                                            for (let r = 0; r < numPastedRows; r++) {
                                                const rowFormulas = [];
                                                for (let c = 0; c < 6; c++) { // K through P is 6 columns
                                                    rowFormulas.push(formulaTemplate);
                                                }
                                                formulaArray.push(rowFormulas);
                                            }
                                            columnKPRange.formulas = formulaArray;
                                            
                                            console.log(`Set column J to 0 and applied ${sumifValue} formula to K${pasteRow}:P${endPastedRow}`);
                                        } else {
                                            // Standard handling for other sumif types: apply formula to entire J-P range
                                            const formulaArray = [];
                                            for (let r = 0; r < numPastedRows; r++) {
                                                const rowFormulas = [];
                                                for (let c = 0; c < 7; c++) { // J through P is 7 columns
                                                    rowFormulas.push(formulaTemplate);
                                                }
                                                formulaArray.push(rowFormulas);
                                            }
                                            rangeToModify.formulas = formulaArray;
                                        }
                                        
                                        await context.sync();
                                        console.log(`"sumif" parameter (${sumifValue}) processing synced for ${sumifRangeAddress}`);
                                    } else {
                                        console.log(`"sumif" parameter value "${sumifValue}" is not recognized. Valid values are: year, yearend, average, offsetyear. No formula changes applied.`);
                                    }
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
 * Parses and converts FORMULA-S custom formula syntax to Excel formula
 * Handles special functions: SPREAD, BEG, END, RAISE, ONETIMEDATE, SPREADDATES
 * @param {string} formulaString - The formula string from the customformula parameter
 * @param {number} targetRow - The row number where the formula will be placed
 * @param {Excel.Worksheet} worksheet - The worksheet object (optional, needed for driver lookups)
 * @param {Excel.RequestContext} context - The Excel context (optional, needed for driver lookups)
 * @returns {Promise<string>} - The converted Excel formula
 */
export async function parseFormulaSCustomFormula(formulaString, targetRow, worksheet = null, context = null) {
    if (!formulaString || typeof formulaString !== 'string') {
        return formulaString;
    }
    
    let result = formulaString;
    
    // Build driver map if worksheet is provided (for cd{1-V1} syntax)
    let driverMap = null;
    if (worksheet && context) {
        try {
            // Use the same range approach as processFormulaSRows
            const startRow = 10; // Same as processFormulaSRows
            
            // Get the last used row in the worksheet
            const usedRange = worksheet.getUsedRange();
            usedRange.load("rowIndex, rowCount");
            await context.sync();
            const lastRow = usedRange.rowIndex + usedRange.rowCount;
            
            // Load column A values to create driver lookup map
            const colARangeAddress = `A${startRow}:A${lastRow}`;
            const colARange = worksheet.getRange(colARangeAddress);
            colARange.load("values");
            await context.sync();
            
            driverMap = new Map();
            const colAValues = colARange.values;
            
            // Build driver map: driver name -> row number (same logic as processFormulaSRows)
            for (let i = 0; i < colAValues.length; i++) {
                const value = colAValues[i][0];
                if (value !== null && value !== "") {
                    const rowNum = startRow + i; // Same calculation as processFormulaSRows
                    driverMap.set(String(value), rowNum);
                    console.log(`    Driver map: ${value} -> row ${rowNum}`);
                }
            }
            console.log(`    Built driver map with ${driverMap.size} entries for cd lookups`);
        } catch (error) {
            console.warn(`    Could not build driver map: ${error.message}`);
            driverMap = null;
        }
    }
    
    // Process rd{} references first if driver map is available
    if (driverMap) {
        const rdPattern = /rd\{([^}]+)\}/g;
        result = result.replace(rdPattern, (match, driverName) => {
            const driverRow = driverMap.get(driverName);
            if (driverRow) {
                const replacement = `AE$${driverRow}`;
                console.log(`    Replacing rd{${driverName}} with ${replacement}`);
                return replacement;
            } else {
                console.warn(`    Driver '${driverName}' not found in column A, keeping as is`);
                return match; // Keep the original if driver not found
            }
        });
    }

    // Process cd{} references before functions too
    // Define column mapping for cd: 1=D, 2=E, 3=F, etc. (excluding A, B, C)
    const columnMapping = {
        '1': 'D',
        '2': 'E',
        '3': 'F',
        '4': 'G',
        '5': 'H',
        '6': 'I',
        '7': 'J',
        '8': 'K',
        '9': 'L'
    };
    
    // Replace all cd{number} or cd{number-driverName} or cd{driverName-number} patterns with column references
    const cdPattern = /cd\{([^}]+)\}/g;
    result = result.replace(cdPattern, (match, content) => {
        // Check if content contains a dash
        const dashIndex = content.indexOf('-');
        let columnNum, driverName;
        
        if (dashIndex !== -1) {
            const firstPart = content.substring(0, dashIndex).trim();
            const secondPart = content.substring(dashIndex + 1).trim();
            
            // Check if first part is a number to determine syntax format
            if (/^\d+$/.test(firstPart)) {
                // Format: {columnNumber-driverName} (existing syntax)
                columnNum = firstPart;
                driverName = secondPart;
                console.log(`    Parsing cd{${content}} as columnNumber-driverName format`);
            } else {
                // Format: {driverName-columnNumber} (new syntax)
                driverName = firstPart;
                columnNum = secondPart;
                console.log(`    Parsing cd{${content}} as driverName-columnNumber format`);
            }
        } else {
            // Just column number, no driver name
            columnNum = content.trim();
            driverName = null;
            console.log(`    Parsing cd{${content}} as columnNumber only format`);
        }
        
        const column = columnMapping[columnNum];
        if (!column) {
            console.warn(`    Column number '${columnNum}' not valid (must be 1-9), keeping as is`);
            return match; // Keep the original if column number not valid
        }
        
        // Determine which row to use
        let rowToUse = targetRow;
        if (driverName && driverMap) {
            const driverRow = driverMap.get(driverName);
            if (driverRow) {
                rowToUse = driverRow;
                console.log(`    Replacing cd{${content}} with $${column}${rowToUse} (driver '${driverName}' found at row ${driverRow})`);
            } else {
                console.warn(`    Driver '${driverName}' not found in column A, using current row ${targetRow}`);
            }
        } else if (driverName && !driverMap) {
            console.warn(`    Cannot look up driver '${driverName}' - driver map not available, using current row ${targetRow}`);
        } else {
            console.log(`    Replacing cd{${columnNum}} with $${column}${rowToUse}`);
        }
        
        const replacement = `$${column}${rowToUse}`;
        return replacement;
    });
    
    // Process SPREAD function: SPREAD(driver) -> driver/AE$7
    result = result.replace(/SPREAD\(([^)]+)\)/gi, (match, driver) => {
        console.log(`    Converting SPREAD(${driver}) to ${driver}/AE$7`);
        return `(${driver}/AE$7)`;
    });
    
    // Process BEG function: BEG(driver) -> (EOMONTH(driver,0)<=AE$2)
    result = result.replace(/BEG\(([^)]+)\)/gi, (match, driver) => {
        console.log(`    Converting BEG(${driver}) to (EOMONTH(${driver},0)<=AE$2)`);
        return `(EOMONTH(${driver},0)<=AE$2)`;
    });
    
    // Process END function: END(driver) -> (EOMONTH(driver,0)>AE$2)
    result = result.replace(/END\(([^)]+)\)/gi, (match, driver) => {
        console.log(`    Converting END(${driver}) to (EOMONTH(${driver},0)>AE$2)`);
        return `(EOMONTH(${driver},0)>AE$2)`;
    });
    
    // Process RAISE function: RAISE(driver1,driver2) -> (1 + (driver1)) ^ (AE$3 - max(year(driver2), $AE3))
    result = result.replace(/RAISE\(([^,]+),([^)]+)\)/gi, (match, driver1, driver2) => {
        // Trim whitespace from drivers
        driver1 = driver1.trim();
        driver2 = driver2.trim();
        console.log(`    Converting RAISE(${driver1},${driver2}) to (1 + (${driver1})) ^ (AE$3 - max(year(${driver2}), $AE3))`);
        return `(1 + (${driver1})) ^ (AE$3 - max(year(${driver2}), $AE3))`;
    });
    
    // Process ONETIMEDATE function: ONETIMEDATE(driver) -> (EOMONTH((driver),0)=AE$2)
    result = result.replace(/ONETIMEDATE\(([^)]+)\)/gi, (match, driver) => {
        console.log(`    Converting ONETIMEDATE(${driver}) to (EOMONTH((${driver}),0)=AE$2)`);
        return `(EOMONTH((${driver}),0)=AE$2)`;
    });
    
    // Process SPREADDATES function: SPREADDATES(driver1,driver2,driver3) -> IF(AND(EOMONTH(AE$2,0)>=EOMONTH(driver2,0),EOMONTH(driver3,0)<=EOMONTH($I12,0)),driver1/(DATEDIF(driver2,driver3,"m")+1),0)
    // Note: Need to handle nested parentheses and comma separation
    result = result.replace(/SPREADDATES\(([^,]+),([^,]+),([^)]+)\)/gi, (match, driver1, driver2, driver3) => {
        // Trim whitespace from drivers
        driver1 = driver1.trim();
        driver2 = driver2.trim();
        driver3 = driver3.trim();
        const newFormula = `IF(AND(EOMONTH(AE$2,0)>=EOMONTH(${driver2},0),EOMONTH(AE$2,0)<=EOMONTH(${driver3},0)),${driver1}/(DATEDIF(${driver2},${driver3},"m")+1),0)`;
        console.log(`    Converting SPREADDATES(${driver1},${driver2},${driver3}) to ${newFormula}`);
        return newFormula;
    });


    
    // Replace timeseriesdivisor with AE$7
    result = result.replace(/timeseriesdivisor/gi, (match) => {
        const replacement = 'AE$7';
        console.log(`    Replacing ${match} with ${replacement}`);
        return replacement;
    });
    
    // Replace currentmonth with AE$2
    result = result.replace(/currentmonth/gi, (match) => {
        const replacement = 'AE$2';
        console.log(`    Replacing ${match} with ${replacement}`);
        return replacement;
    });
    
    // Replace beginningmonth with $AE$2
    result = result.replace(/beginningmonth/gi, (match) => {
        const replacement = '$AE$2';
        console.log(`    Replacing ${match} with ${replacement}`);
        return replacement;
    });
    
    // Replace currentyear with AE$3
    result = result.replace(/currentyear/gi, (match) => {
        const replacement = 'AE$3';
        console.log(`    Replacing ${match} with ${replacement}`);
        return replacement;
    });
    
    // Replace yearend with AE$4
    result = result.replace(/yearend/gi, (match) => {
        const replacement = 'AE$4';
        console.log(`    Replacing ${match} with ${replacement}`);
        return replacement;
    });
    
    // Replace beginningyear with $AE$2
    result = result.replace(/beginningyear/gi, (match) => {
        const replacement = '$AE$2';
        console.log(`    Replacing ${match} with ${replacement}`);
        return replacement;
    });
    
    return result;
}

/**
 * Parses a value string and extracts formatting information based on symbols
 * @param {string} valueString - The value string that may contain formatting symbols
 * @returns {Object} - Object containing cleaned value, format type, and italic setting
 */
function parseValueWithSymbols(valueString) {
    if (!valueString || typeof valueString !== 'string') {
        return { cleanValue: valueString, formatType: null, isItalic: false };
    }

    let cleanValue = valueString;
    let formatType = null;
    let isItalic = false;
    let currencySymbol = null;

    // Check for ~currency prefix (currency with italic)
    if (cleanValue.startsWith('~$')) {
        formatType = 'dollaritalic';
        currencySymbol = '$';
        isItalic = true;
        cleanValue = cleanValue.substring(2);
    }
    else if (cleanValue.startsWith('~£')) {
        formatType = 'pounditalic';
        currencySymbol = '£';
        isItalic = true;
        cleanValue = cleanValue.substring(2);
    }
    else if (cleanValue.startsWith('~€')) {
        formatType = 'euroitalic';
        currencySymbol = '€';
        isItalic = true;
        cleanValue = cleanValue.substring(2);
    }
    else if (cleanValue.startsWith('~¥')) {
        formatType = 'yenitalic';
        currencySymbol = '¥';
        isItalic = true;
        cleanValue = cleanValue.substring(2);
    }
    // Check for currency prefixes (currency without italic)
    else if (cleanValue.startsWith('$')) {
        formatType = 'dollar';
        currencySymbol = '$';
        cleanValue = cleanValue.substring(1);
    }
    else if (cleanValue.startsWith('£')) {
        formatType = 'pound';
        currencySymbol = '£';
        cleanValue = cleanValue.substring(1);
    }
    else if (cleanValue.startsWith('€')) {
        formatType = 'euro';
        currencySymbol = '€';
        cleanValue = cleanValue.substring(1);
    }
    else if (cleanValue.startsWith('¥')) {
        formatType = 'yen';
        currencySymbol = '¥';
        cleanValue = cleanValue.substring(1);
    }
    // Check for ~ prefix (italic)
    else if (cleanValue.startsWith('~')) {
        isItalic = true;
        cleanValue = cleanValue.substring(1); // Remove ~ prefix
        
        // If after removing ~, it's a number or "F", apply volume formatting
        if (cleanValue === 'F' || (!isNaN(Number(cleanValue)) && cleanValue.trim() !== '')) {
            formatType = 'volume';
        }
    }
    // Check if it's just "F" (volume formatting)
    else if (cleanValue === 'F') {
        formatType = 'volume';
    }
    // Check if it ends with "x" for factor formatting (e.g., "10x")
    else if (cleanValue.toLowerCase().endsWith('x') && cleanValue.length > 1) {
        const numberPart = cleanValue.substring(0, cleanValue.length - 1);
        if (!isNaN(Number(numberPart)) && numberPart.trim() !== '') {
            formatType = 'factor';
            cleanValue = numberPart; // Remove the 'x' from the clean value
        }
    }
    // Check if it's a number without symbols (volume formatting)
    else if (!isNaN(Number(cleanValue)) && cleanValue.trim() !== '' && cleanValue !== '') {
        formatType = 'volume';
    }
    // Check if it's a date (basic date pattern check)
    else if (isDateString(cleanValue)) {
        formatType = 'date';
    }

    return { cleanValue, formatType, isItalic, currencySymbol };
}

/**
 * Enhanced date string detection
 * @param {string} value - The value to check
 * @returns {boolean} - True if the value appears to be a date
 */
function isDateString(value) {
    if (!value || typeof value !== 'string') return false;
    
    // Check for common date patterns
    const datePatterns = [
        /^\d{1,2}\/\d{1,2}\/\d{4}$/,        // MM/DD/YYYY or M/D/YYYY
        /^\d{1,2}-\d{1,2}-\d{4}$/,          // MM-DD-YYYY or M-D-YYYY
        /^\d{4}\/\d{1,2}\/\d{1,2}$/,        // YYYY/MM/DD or YYYY/M/D
        /^\d{4}-\d{1,2}-\d{1,2}$/,          // YYYY-MM-DD or YYYY-M-D
        /^\d{1,2}\/\d{1,2}\/\d{2}$/,        // MM/DD/YY or M/D/YY
        /^\d{1,2}-\d{1,2}-\d{2}$/           // MM-DD-YY or M-D-YY
    ];
    
    for (const pattern of datePatterns) {
        if (pattern.test(value)) {
            const date = new Date(value);
            return !isNaN(date.getTime()) && date.getFullYear() > 1900 && date.getFullYear() < 2100;
        }
    }
    
    return false;
}

/**
 * Applies symbol-based formatting to a row of cells
 * @param {Excel.Worksheet} worksheet - The worksheet containing the cells
 * @param {number} rowNum - The row number to format
 * @param {Array} splitArray - Array of values from the row parameter
 * @param {Array} columnSequence - Array of column letters
 * @returns {Promise<void>}
 */
async function applyRowSymbolFormatting(worksheet, rowNum, splitArray, columnSequence) {
    console.log(`Applying symbol-based formatting to row ${rowNum}`);
    
    // Define format configurations
    const formatConfigs = {
        'dollaritalic': {
            numberFormat: '_(* $ #,##0_);_(* $ (#,##0);_(* "$" -""?_);_(@_)',
            italic: true,
            bold: false
        },
        'dollar': {
            numberFormat: '_(* $ #,##0_);_(* $ (#,##0);_(* "$" -""?_);_(@_)',
            italic: false,
            bold: false
        },
        'pounditalic': {
            numberFormat: '_(* [$£-809] #,##0_);_(* [$£-809] (#,##0);_(* [$£-809] "-"?_);_(@_)',
            italic: true,
            bold: false
        },
        'pound': {
            numberFormat: '_(* [$£-809] #,##0_);_(* [$£-809] (#,##0);_(* [$£-809] "-"?_);_(@_)',
            italic: false,
            bold: false
        },
        'euroitalic': {
            numberFormat: '_(* [$€-407] #,##0_);_(* [$€-407] (#,##0);_(* [$€-407] "-"?_);_(@_)',
            italic: true,
            bold: false
        },
        'euro': {
            numberFormat: '_(* [$€-407] #,##0_);_(* [$€-407] (#,##0);_(* [$€-407] "-"?_);_(@_)',
            italic: false,
            bold: false
        },
        'yenitalic': {
            numberFormat: '_(* [$¥-411] #,##0_);_(* [$¥-411] (#,##0);_(* [$¥-411] "-"?_);_(@_)',
            italic: true,
            bold: false
        },
        'yen': {
            numberFormat: '_(* [$¥-411] #,##0_);_(* [$¥-411] (#,##0);_(* [$¥-411] "-"?_);_(@_)',
            italic: false,
            bold: false
        },
        'volume': {
            numberFormat: '_(* #,##0_);_(* (#,##0);_(* " -"?_);_(@_)',
            italic: true,
            bold: false
        },
        'date': {
            numberFormat: 'mmm-yy',
            italic: false,
            bold: false
        },
        'factor': {
            numberFormat: '_(* #,##0.0x;_(* (#,##0.0)x;_(* "   -"?_)',
            italic: true,
            bold: false
        }
    };

    for (let x = 0; x < splitArray.length && x < columnSequence.length; x++) {
        const originalValue = splitArray[x];
        const colLetter = columnSequence[x];
        
        if (!originalValue) continue; // Skip empty values
        
        // Parse the value and extract formatting information
        const parsed = parseValueWithSymbols(originalValue);
        console.log(`  Column ${colLetter}: "${originalValue}" -> cleanValue: "${parsed.cleanValue}", format: ${parsed.formatType}, italic: ${parsed.isItalic}, currency: ${parsed.currencySymbol || 'none'}`);
        
        const cellRange = worksheet.getRange(`${colLetter}${rowNum}`);
        
                 // Apply number format if specified
         if (parsed.formatType && formatConfigs[parsed.formatType]) {
             const config = formatConfigs[parsed.formatType];
             cellRange.numberFormat = [[config.numberFormat]];
             cellRange.format.font.italic = config.italic;
             // Don't override bold formatting - preserve existing bold setting
             // cellRange.format.font.bold = config.bold;  // Removed to preserve existing bold
             console.log(`    Applied ${parsed.formatType} formatting to ${colLetter}${rowNum}`);
         }
         
         // Apply italic formatting if ~ symbol was present (overrides format config)
         if (parsed.isItalic) {
             cellRange.format.font.italic = true;
             console.log(`    Applied italic to ${colLetter}${rowNum} due to ~ symbol`);
         }
         // For regular text with no symbols and no specific format, ensure it's not italic
         else if (!parsed.formatType) {
             cellRange.format.font.italic = false;
             console.log(`    Set non-italic for regular text in ${colLetter}${rowNum}`);
         }
    }
    
    // Sync the formatting changes
    await worksheet.context.sync();
    console.log(`Completed symbol-based formatting for row ${rowNum}`);
}

/**
 * Copies complete formatting from column O to column J and columns S through CX for a specific row
 * This includes number format (currency) and all font formatting (italic, bold, etc.)
 * @param {Excel.Worksheet} worksheet - The worksheet containing the cells
 * @param {number} rowNum - The row number to copy formatting for
 * @returns {Promise<void>}
 */
async function copyColumnPFormattingToJAndSCX(worksheet, rowNum) {
    console.log(`Copying complete column O formatting to J and S:CX for row ${rowNum}`);
    
    try {
        // Get the source cell and target ranges
        const sourceCellO = worksheet.getRange(`O${rowNum}`);
        const targetCellJ = worksheet.getRange(`J${rowNum}`);
        const targetRangeSCX = worksheet.getRange(`S${rowNum}:CX${rowNum}`);
        
        // Load complete formatting from source cell O
        sourceCellO.load(["numberFormat", "format/font/italic", "format/font/bold"]);
        await worksheet.context.sync();
        
        // Copy complete formatting from O to J
        targetCellJ.copyFrom(sourceCellO, Excel.RangeCopyType.formats);
        
        // Copy complete formatting from O to S:CX range
        targetRangeSCX.copyFrom(sourceCellO, Excel.RangeCopyType.formats);
        
        // Sync the formatting changes
        await worksheet.context.sync();
        
        const numberFormat = sourceCellO.numberFormat[0][0];
        console.log(`Successfully copied complete column O formatting to J${rowNum} and S${rowNum}:CX${rowNum}`);
        console.log(`  Applied number format: ${numberFormat}`);
        console.log(`  Applied font italic: ${sourceCellO.format.font.italic}`);
        console.log(`  Applied font bold: ${sourceCellO.format.font.bold}`);
        
    } catch (error) {
        console.error(`Error copying column O formatting to J and S:CX for row ${rowNum}: ${error.message}`);
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

            // NEW SECTION: Convert row references to absolute for columns >= AE
            // First, load all formulas from columns AE to CX for the range of interest
            console.log("Making row references absolute for cell references in columns >= AE before row insertion");
            
            const START_ROW = 10;
            const TARGET_COL = "AE";
            const END_COL = "CX";
            
            // Define range to process
            let processStartRow = Math.min(calcsPasteRow, START_ROW);
            let processEndRow = determinedLastRow;
            
            const formulaRangeAddress = `${TARGET_COL}${processStartRow}:${END_COL}${processEndRow}`;
            console.log(`Loading formulas from range: ${formulaRangeAddress}`);
            
            try {
                const formulaRange = currentWorksheet.getRange(formulaRangeAddress);
                formulaRange.load("formulas");
                await context.sync();
                
                // Calculate TARGET_COL index for reference comparisons
                const targetColIndex = columnLetterToIndex(TARGET_COL);
                console.log(`Target column ${TARGET_COL} has index ${targetColIndex}`);
                
                let formulasUpdated = false;
                const origFormulas = formulaRange.formulas;
                const newFormulas = [];
                
                // Process each formula in the range
                for (let r = 0; r < origFormulas.length; r++) {
                    const rowFormulas = [];
                    
                    for (let c = 0; c < origFormulas[r].length; c++) {
                        let formula = origFormulas[r][c];
                        
                        // Only process string formulas
                        if (typeof formula === 'string') {
                            // Skip if it's not a formula
                            if (!formula.startsWith('=')) {
                                rowFormulas.push(formula);
                                continue;
                            }
                            
                            // Find cell references (e.g., A1, B2, AA34) but exclude already absolute refs (e.g., A$1, $A$1)
                            // This regex captures: group 1 = column letter(s), group 2 = row number
                            // It skips references that already have $ before the row number
                            const cellRefRegex = /([A-Z]+)(\d+)(?![^\W_])/g;
                            
                            // Replace with absolute row references where needed
                            const originalFormula = formula;
                            formula = formula.replace(cellRefRegex, (match, col, row) => {
                                // Get column index
                                const colIndex = columnLetterToIndex(col);
                                
                                // If column index is >= target column index, make row reference absolute
                                if (colIndex >= targetColIndex) {
                                    return `${col}$${row}`;
                                }
                                return match; // Keep as is for columns before TARGET_COL
                            });
                            
                            if (formula !== originalFormula) {
                                formulasUpdated = true;
                                //console.log(`  Row ${processStartRow + r}, Col ${columnIndexToLetter(c + targetColIndex)}: Formula changed from '${originalFormula}' to '${formula}'`);
                            }
                        }
                        
                        rowFormulas.push(formula);
                    }
                    
                    newFormulas.push(rowFormulas);
                }
                
                // Only update if changes were made
                if (formulasUpdated) {
                    console.log(`Updating formulas with absolute row references in range ${formulaRangeAddress}`);
                    formulaRange.formulas = newFormulas;
                    await context.sync();
                    console.log("Formula updates completed");
                } else {
                    console.log("No formulas needed absolute row reference updates");
                }
            } catch (formulaError) {
                console.error(`Error processing formulas for absolute row references: ${formulaError.message}`, formulaError);
                // Continue with the function, don't let this conversion stop the flow
            }
            // END NEW SECTION

            const columnSequence = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'K', 'L', 'M', 'N', 'O', 'P', 'R'];
            
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

                        const originalValue = splitArray[x];
                        const colLetter = columnSequence[x];
                        const cellToWrite = currentWorksheet.getRange(`${colLetter}${currentRowNum}`);
                        
                        // Parse the value to extract clean value without formatting symbols
                        const parsed = parseValueWithSymbols(originalValue);
                        const valueToWrite = parsed.cleanValue;
                        
                        // VBA check: If splitArray(x) <> "" And splitArray(x) <> "F" Then
                        // 'F' likely means "Formula", so we don't overwrite if the value is 'F'.
                        if (valueToWrite && valueToWrite.toUpperCase() !== 'F') {
                            // Attempt to infer data type (basic number check)
                            const numValue = Number(valueToWrite);
                            if (!isNaN(numValue) && valueToWrite.trim() !== '') {
                                cellToWrite.values = [[numValue]];
                            } else {
                                // Write text value if not empty
                                if (valueToWrite.trim() !== '') {
                                    cellToWrite.values = [[valueToWrite]];
                                } else {
                                    // Clear the cell if the value is empty/blank
                                    cellToWrite.clear(Excel.ClearApplyTo.contents);
                                }
                            }
                            // console.log(`  Wrote '${valueToWrite}' to ${colLetter}${currentRowNum}`);
                        } else if (valueToWrite === '' || valueToWrite === null || valueToWrite === undefined) {
                            // Clear the cell if the value is explicitly empty (blank between ||| separators)
                            cellToWrite.clear(Excel.ClearApplyTo.contents);
                            // console.log(`  Cleared contents of ${colLetter}${currentRowNum} due to blank value`);
                        }
                    }
                    
                    // Apply customformula parameter to column AE for row1 (FORMULA-S codes only)
                    if (g === 1 && yy === 0 && code.type === "FORMULA-S" && code.params.customformula && code.params.customformula !== "0") {
                        try {
                            console.log(`  Applying customformula to AE${currentRowNum} for FORMULA-S: ${code.params.customformula}`);
                            
                            // For FORMULA-S, set it as a value that will be converted to formula later by processFormulaSRows
                            const customFormulaCell = currentWorksheet.getRange(`AE${currentRowNum}`);
                            customFormulaCell.values = [[code.params.customformula]];
                            console.log(`  Set customformula as value for FORMULA-S processing`);
                            
                            // Track this row for processFormulaSRows since column D will be overwritten
                            addFormulaSRow(worksheetName, currentRowNum);
                        } catch (customFormulaError) {
                            console.error(`  Error applying customformula: ${customFormulaError.message}`);
                        }
                    }

                    // Apply symbol-based formatting for each cell in the row
                    try {
                        await applyRowSymbolFormatting(currentWorksheet, currentRowNum, splitArray, columnSequence);
                    } catch (formatError) {
                        console.error(`  Error applying symbol-based formatting: ${formatError.message}`);
                    }

                    // Apply columnformula parameter to column AE for row1 (all code types)
                    if (g === 1 && yy === 0 && code.params.columnformula && code.params.columnformula !== "0") {
                        try {
                            console.log(`  Processing columnformula for AE${currentRowNum}: ${code.params.columnformula}`);
                            
                            // Process the formula through parseFormulaSCustomFormula to handle BEG, END, etc.
                            let processedFormula = await parseFormulaSCustomFormula(code.params.columnformula, currentRowNum, currentWorksheet, context);
                            console.log(`  Processed columnformula result: ${processedFormula}`);
                            
                            // Check for negative parameter and apply it to the processed formula
                            if (code.params.negative && String(code.params.negative).toUpperCase() === "TRUE") {
                                console.log(`  Applying negative transformation to columnformula`);
                                processedFormula = `-(${processedFormula})`;
                                console.log(`  Columnformula after negative transformation: ${processedFormula}`);
                            }
                            
                            // Apply the processed formula to the cell
                            const columnFormulaCell = currentWorksheet.getRange(`AE${currentRowNum}`);
                            if (processedFormula && processedFormula !== code.params.columnformula) {
                                // If the formula was modified, set it as a formula
                                columnFormulaCell.formulas = [['=' + processedFormula]];
                                console.log(`  Set processed columnformula as formula: =${processedFormula}`);
                            } else {
                                // If no processing occurred, set as original value
                                columnFormulaCell.values = [[code.params.columnformula]];
                                console.log(`  Set original columnformula as value: ${code.params.columnformula}`);
                            }
                        } catch (columnFormulaError) {
                            console.error(`  Error applying columnformula: ${columnFormulaError.message}`);
                        }
                    }

                    // NOTE: columnformat parameter replaced by symbol-based formatting
                    // The formatting is now applied through symbols in row data (e.g., ~$120000, $F, etc.)

                    // NEW: Apply columncomment parameter to specified columns for row1
                    if (g === 1 && yy === 0 && code.params.columncomment) {
                        try {
                            console.log(`  Applying columncomment to row ${currentRowNum}: ${code.params.columncomment}`);
                            
                            // Split the comment string by backslash
                            const comments = code.params.columncomment.split('\\');
                            console.log(`  Comment array: [${comments.join(', ')}]`);
                            
                            // Define column mapping (excluding A, B, C columns)
                            const columnMapping = {
                                '1': 'D',
                                '2': 'E',
                                '3': 'F',
                                '4': 'G',
                                '5': 'H',
                                '6': 'I',
                                '7': 'J',
                                '8': 'K',
                                '9': 'L'
                            };
                            
                            // Apply each comment to its corresponding column
                            for (let i = 0; i < comments.length && i < 9; i++) {
                                const commentText = comments[i].trim();
                                const columnNum = String(i + 1);
                                const columnLetter = columnMapping[columnNum];
                                
                                if (!columnLetter) {
                                    console.warn(`    Column number ${columnNum} exceeds mapping range, skipping`);
                                    continue;
                                }
                                
                                // Skip empty comments
                                if (!commentText) {
                                    console.log(`    Empty comment for column ${columnNum}, skipping`);
                                    continue;
                                }
                                
                                const cellAddress = `${columnLetter}${currentRowNum}`;
                                
                                // Get the cell to verify it exists and has content
                                const targetCell = currentWorksheet.getRange(cellAddress);
                                targetCell.load(["values"]);
                                await context.sync();
                                
                                const currentValue = targetCell.values[0][0];
                                console.log(`    Cell ${cellAddress}: Current value = '${currentValue}', Adding comment: '${commentText}'`);
                                
                                // Add comment to the cell using the worksheet's comments collection
                                currentWorksheet.comments.add(cellAddress, commentText);
                                
                                console.log(`    Added comment '${commentText}' to ${cellAddress}`);
                            }
                            
                            await context.sync(); // Sync the comment changes
                            console.log(`  Finished applying columncomment to row ${currentRowNum}`);
                            
                        } catch (columnCommentError) {
                            console.error(`  Error applying columncomment: ${columnCommentError.message}`);
                        }
                    }

                    // FINAL STEP: Copy complete column P formatting to J and S:CX (after all other formatting)
                    try {
                        await copyColumnPFormattingToJAndSCX(currentWorksheet, currentRowNum);
                    } catch (copyFormatError) {
                        console.error(`  Error copying column P formatting to J and S:CX: ${copyFormatError.message}`);
                    }
                }
                await context.sync(); // Sync after populating each 'g' group

                // Adjust the base check row marker for subsequent 'g' iterations
                // by adding the number of rows inserted in *this* iteration.
                currentCheckRowForInserts += numNewRows;
                console.log(`Finished processing row${g}. currentCheckRowForInserts is now ${currentCheckRowForInserts}`);

            } // End for g loop

            console.log(`Completed processing driver and assumption inputs for code ${codeValue} in worksheet ${worksheetName}`);
            
            // Return information about the actual rows populated
            const actualFirstRow = searchRow;
            const actualLastRow = currentCheckRowForInserts - 1; // The last row populated
            const totalRowsPopulated = actualLastRow - actualFirstRow + 1;
            
            return {
                firstRow: actualFirstRow,
                lastRow: actualLastRow,
                totalRows: totalRowsPopulated
            };
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
    const START_ROW = 10; // <<< CHANGED FROM 9
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
    const START_ROW = 10; // <<< CHANGED FROM 9
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
        
        // Calculate TARGET_COL index for reference comparisons
        const targetColIndex = columnLetterToIndex(TARGET_COL);
        console.log(`Target column ${TARGET_COL} has index ${targetColIndex}`);
        
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
                
                // NEW SECTION: Convert row references to absolute for columns >= TARGET_COL
                if (typeof currentFormula === 'string') {
                    console.log(`Making row references absolute for cell references in columns >= ${TARGET_COL} in row ${START_ROW + item.index}`);
                    
                    // Find cell references (e.g., A1, B2, AA34) but exclude already absolute refs (e.g., A$1, $A$1)
                    // This regex captures: group 1 = column letter(s), group 2 = row number
                    // It skips references that already have $ before the row number
                    const cellRefRegex = /([A-Z]+)(\d+)(?![^\W_])/g;
                    
                    // Replace with absolute row references where needed
                    currentFormula = currentFormula.replace(cellRefRegex, (match, col, row) => {
                        // Get column index
                        const colIndex = columnLetterToIndex(col);
                        
                        // If column index is >= target column index, make row reference absolute
                        if (colIndex >= targetColIndex) {
                            return `${col}$${row}`;
                        }
                        return match; // Keep as is for columns before TARGET_COL
                    });
                    
                    console.log(`  Formula after converting to absolute row refs: ${currentFormula}`);
                }
                // END NEW SECTION
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

    const CALCS_FIRST_ROW = 10; // <<< CHANGED FROM 9 // Same as START_ROW elsewhere
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
                   // Use a potentially more reliable column like B for last row fallback
                   const fallbackRange = financialsSheet.getRange(`${FINANCIALS_TARGET_COL_B}1:${FINANCIALS_TARGET_COL_B}10000`); // Check Col B
                   fallbackRange.load("values");
                   await worksheet.context.sync();
                   for (let i = fallbackRange.values.length - 1; i >= 0; i--) {
                       if (fallbackRange.values[i][0] !== null && fallbackRange.values[i][0] !== "") {
                           financialsLastRow = i + 1;
                           break;
                       }
                   }
                   if (financialsLastRow === 0) console.warn(`Fallback range load for Financials col ${FINANCIALS_TARGET_COL_B} also yielded no data.`);
               } catch (fallbackError) {
                    console.error(`Error during fallback range loading for Financials col ${FINANCIALS_TARGET_COL_B}:`, fallbackError);
                    financialsLastRow = 0; // Keep it 0 if fallback fails
               }
           }
        }
        // Recalculate financialsLastRow based on Col B if it's potentially larger
        try {
            const lastRowB = await getLastUsedRow(financialsSheet, FINANCIALS_TARGET_COL_B);
            financialsLastRow = Math.max(financialsLastRow, lastRowB);
        } catch (lastRowBErr) {
            console.warn(`Could not get last row from Col B: ${lastRowBErr.message}`);
        }

        console.log(`Financials last relevant row used for processing: ${financialsLastRow}`);


        // 3. Create Map of Financials Codes (Col I) -> Row Number
        // MODIFIED: Use a case-insensitive map for codes
        const financialsCodeMap = new Map();
        if (financialsLastRow > 0) {
            const financialsCodeRange = financialsSheet.getRange(`${FINANCIALS_CODE_COLUMN}1:${FINANCIALS_CODE_COLUMN}${financialsLastRow}`);
            financialsCodeRange.load("values");
            await worksheet.context.sync(); // Sync map data load
            for (let i = 0; i < financialsCodeRange.values.length; i++) {
                const code = financialsCodeRange.values[i][0];
                if (code !== null && code !== "") {
                    // Convert code to uppercase for case-insensitive comparison
                    const upperCode = String(code).toUpperCase();
                    // Only map the first occurrence of a code, like .Find would
                    if (!financialsCodeMap.has(upperCode)) {
                         financialsCodeMap.set(upperCode, i + 1);
                    }
                }
            }
            console.log(`Built Financials code map with ${financialsCodeMap.size} entries.`);
        } else {
            console.warn(`Financials sheet column ${FINANCIALS_CODE_COLUMN} appears empty or last row not found. No codes loaded for map.`);
        }

        // *** REMOVED: Logic for existingDataLinks Set ***
        // const existingDataLinks = new Set();
        // if (financialsLastRow > 0) {
        //     ... load formulas from Financials Col B ...
        //     ... populate existingDataLinks set ...
        // }
        // *** END REMOVED ***

        // 4. Identify rows to insert and prepare task data
        const tasks = [];
        console.log("populateFinancialsJS: Syncing assumption codes load...");
        await worksheet.context.sync(); // Sync needed for assumptionCodeRange.values

        // *** RELOAD assumption codes here AFTER the sync above, just in case ***
        // It's safer to reload after any potential sync/modification, though unlikely needed here.
        // Keeping the original load before the Financials code map creation seems okay.
        const assumptionCodes = assumptionCodeRange.values; // Use the already loaded values

        console.log(`populateFinancialsJS: Processing ${assumptionCodes?.length ?? 0} assumption rows.`);

        // --- REMOVED Debug logging for row 17 values/addresses ---

        for (let i = 0; i < (assumptionCodes?.length ?? 0); i++) {
            const code = assumptionCodes[i][0];
            const assumptionRow = CALCS_FIRST_ROW + i; // This is the correct Excel row number

            if (code !== null && code !== "") {
                // Construct the potential link formulas first
                const linkFormulaB = `='${worksheet.name}'!${ASSUMPTION_LINK_COL_B}${assumptionRow}`;
                const linkFormulaD = `='${worksheet.name}'!${ASSUMPTION_LINK_COL_D}${assumptionRow}`;
                const linkFormulaMonths = `='${worksheet.name}'!${ASSUMPTION_MONTHS_START_COL}${assumptionRow}`;

                // *** REMOVED CHECK 1: Skip if this assumption row link already exists in Financials Col B ***
                // if (existingDataLinks.has(linkFormulaB)) {
                //     console.log(`  Skipping Code ${code} (Assumption Row ${assumptionRow}): Link ${linkFormulaB} already exists in Financials!${FINANCIALS_TARGET_COL_B}.`);
                //     continue; // Skip to next assumption code
                // }

                // *** MODIFIED: Use case-insensitive check for code existence ***
                const upperCode = String(code).toUpperCase();
                if (!financialsCodeMap.has(upperCode)) {
                     console.log(`  Skipping Code ${code} (Assumption Row ${assumptionRow}): Code not found in Financials template column ${FINANCIALS_CODE_COLUMN}. Cannot determine target row.`);
                     continue; // Skip if no template row found
                }

                // If the code exists in the map, proceed to create the task
                const targetRow = financialsCodeMap.get(upperCode); // Get the row number from the map
                console.log(`  Task Prep: Code ${code} (Assumption Row ${assumptionRow}) -> Target Financials Row (for insertion): ${targetRow}`);

                tasks.push({
                    targetRow: targetRow,
                    assumptionRow: assumptionRow,
                    code: code,
                    addressB: linkFormulaB,     // Use the constructed formula link
                    addressD: linkFormulaD,     // Use the constructed formula link
                    addressMonths: linkFormulaMonths // Use the constructed formula link
                });
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
            // MODIFIED: Make code prefix comparison case-insensitive
            let codePrefix = String(task.code).substring(0, 2).toUpperCase();
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
            
            // --- NEW: Populate Actuals Columns S:AD with SUMIFS formula ---
            try {
                const actualsRange = financialsSheet.getRange(`S${populateRow}:AD${populateRow}`);
                const sumifsFormula = "=SUMIFS('actuals'!$B:$B,'actuals'!$D:$D,EOMONTH(INDIRECT(ADDRESS(2,COLUMN())),0),'actuals'!$E:$E,@INDIRECT(ADDRESS(ROW(),2)))";
                
                // Create a 2D array matching the range dimensions
                const numCols = columnLetterToIndex('AD') - columnLetterToIndex('S') + 1;
                const formulasArray = [Array(numCols).fill(sumifsFormula)];
                actualsRange.formulas = formulasArray;
                
                // Apply formatting
                actualsRange.format.numberFormat = CURRENCY_FORMAT;
                actualsRange.format.font.bold = false;
                actualsRange.format.font.italic = false;
                actualsRange.format.font.color = "#7030A0"; // Set font color
                console.log(`  Set SUMIFS formula for S${populateRow}:AD${populateRow}`);
            } catch (sumifsError) {
                console.error(`Error setting SUMIFS formula for row ${populateRow} (Code: ${task.code}):`, sumifsError.debugInfo || sumifsError);
            }
            // --- END NEW SECTION ---
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

        // *** NEW STEP: Modify codes in Assumption Sheet Column C ***
        console.log(`Modifying codes in ${worksheet.name} column ${ASSUMPTION_CODE_COL} (${CALCS_FIRST_ROW}:${lastRow}) by prepending '-'...`);
        try {
            // Re-get the range and load values (ensure we have the latest state)
            const codeColRange = worksheet.getRange(`${ASSUMPTION_CODE_COL}${CALCS_FIRST_ROW}:${ASSUMPTION_CODE_COL}${lastRow}`);
            codeColRange.load("values");
            await worksheet.context.sync(); // Load the values

            const currentCodeValues = codeColRange.values;
            const modifiedCodeValues = [];
            let modifiedCount = 0;

            for (let i = 0; i < currentCodeValues.length; i++) {
                const originalValue = currentCodeValues[i][0];
                if (originalValue !== null && originalValue !== "" && !String(originalValue).startsWith('-')) {
                    modifiedCodeValues.push(["-" + originalValue]); // Prepend "-"
                    modifiedCount++;
                } else {
                    modifiedCodeValues.push([originalValue]); // Keep original if empty, null, or already starts with '-'
                }
            }

            // Write the modified values back if any changes were made
            if (modifiedCount > 0) {
                 console.log(`  Writing ${modifiedCount} modified codes back to ${ASSUMPTION_CODE_COL}${CALCS_FIRST_ROW}:${ASSUMPTION_CODE_COL}${lastRow}`);
                 codeColRange.values = modifiedCodeValues;
                 await worksheet.context.sync(); // Sync the code modifications
                 console.log("  Synced code modifications.");
            } else {
                console.log("  No codes needed modification.");
            }
        } catch (modifyError) {
             console.error(`Error modifying codes in ${worksheet.name} column ${ASSUMPTION_CODE_COL}:`, modifyError.debugInfo || modifyError);
             // Continue even if modification fails? Or throw? Let's log and continue.
        }
        // *** END NEW STEP ***


        console.log(`populateFinancialsJS successfully completed for ${worksheet.name} -> ${financialsSheet.name}`);

    } catch (error) {
        console.error(`Error in populateFinancialsJS for sheet ${worksheet.name} -> ${financialsSheet.name}:`, error.debugInfo || error);
        throw error;
    }
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
    const START_ROW = 10; // <<< CHANGED FROM 9 // Standard start row for processing

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

                     // 6.5 Set font color to white in column A
                     // We use updatedLastRow here, as deletions haven't happened yet
                     await setColumnAFontWhite(currentWorksheet, START_ROW, updatedLastRow); 
                     console.log(`Set font color to white in column A from rows ${START_ROW}-${updatedLastRow}`);
  
                     // // Force recalculation before Index Growth Curve (especially if manual calc mode)
                     // console.log(`Performing full workbook recalculation before Index Growth Curve for ${worksheetName}...`);
                     // context.workbook.application.calculate(Excel.CalculationType.fullRebuild);
                     // await context.sync(); // Sync the calculation
                     // console.log(`Recalculation complete for ${worksheetName}.`);
 
                      // 6.8 Apply Index Growth Curve logic (if applicable)
                     // Run Index Growth *before* deleting rows. Use updatedLastRow as the boundary.
                     await applyIndexGrowthCurveJS(currentWorksheet, updatedLastRow); 
                     
                     // 6.9 Process FORMULA-S rows - Convert driver references to cell references BEFORE deleting green rows
                     console.log(`Processing FORMULA-S rows in ${worksheetName} (before green row deletion)...`);
                     await processFormulaSRows(currentWorksheet, START_ROW, updatedLastRow);
                     console.log(`Finished processing FORMULA-S rows`);
                     
                     // Clear the tracked FORMULA-S rows for this worksheet (cleanup)
                     clearFormulaSRows(worksheetName);
                     
                     // 7. Delete rows with green background (#CCFFCC) - AFTER FORMULA-S processing
                     console.log(`Deleting green rows in ${worksheetName}...`);
                     // Changed START_ROW to START_ROW - 1 to include row 9
                     const finalLastRow = await deleteGreenRows(currentWorksheet, START_ROW - 1, updatedLastRow); // Get the new last row AFTER deletions
                     console.log(`After deleting green rows, last row is now: ${finalLastRow}`);
 
                     // 8. Autofill AE9:AE<lastRow> -> CX<lastRow> on Assumption Tab - Use finalLastRow
                     console.log(`Autofilling ${AUTOFILL_START_COLUMN}${START_ROW}:${AUTOFILL_START_COLUMN}${finalLastRow} to ${AUTOFILL_END_COLUMN} on ${worksheetName}`);
                     const sourceRange = currentWorksheet.getRange(`${AUTOFILL_START_COLUMN}${START_ROW}:${AUTOFILL_START_COLUMN}${finalLastRow}`);
                     const fillRange = currentWorksheet.getRange(`${AUTOFILL_START_COLUMN}${START_ROW}:${AUTOFILL_END_COLUMN}${finalLastRow}`);
                     sourceRange.autoFill(fillRange, Excel.AutoFillType.fillDefault);
 
                     // 9. Set Row 9 interior color to none
                     console.log(`Setting row 9 interior color to none for ${worksheetName}`);
                     const row9Range = currentWorksheet.getRange("9:9");
                     row9Range.format.fill.clear();

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
                 // await formatChangesInWorkingCapitalJS(finSheet); // <<< COMMENTED OUT

                 // 2. Get Last Row for Financials
                 const financialsLastRow = await getLastUsedRow(finSheet, "B"); // Pass sheet from this context
                 if (financialsLastRow < START_ROW) {
                     console.warn(`Skipping final autofill on ${FINANCIALS_SHEET_NAME} as last row (${financialsLastRow}) is before start row (${START_ROW}).`);
                     return;
                 }
                 console.log(`Last row in Col B for ${FINANCIALS_SHEET_NAME}: ${financialsLastRow}`);

                //  // 3. Autofill AE9:AE<lastRow> -> CX<lastRow> on Financials
                //  console.log(`Autofilling ${AUTOFILL_START_COLUMN}${START_ROW}:${AUTOFILL_START_COLUMN}${financialsLastRow} to ${AUTOFILL_END_COLUMN} on ${FINANCIALS_SHEET_NAME}`);
                //  const sourceRangeFin = finSheet.getRange(`${AUTOFILL_START_COLUMN}${START_ROW}:${AUTOFILL_START_COLUMN}${financialsLastRow}`);
                //  const fillRangeFin = finSheet.getRange(`${AUTOFILL_START_COLUMN}${START_ROW}:${AUTOFILL_END_COLUMN}${financialsLastRow}`);
                //  sourceRangeFin.autoFill(fillRangeFin, Excel.AutoFillType.fillDefault);


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

/**
 * Deletes rows with light green background (#CCFFCC) in column B
 * @param {Excel.Worksheet} worksheet - The worksheet to process
 * @param {number} startRow - The first row to check
 * @param {number} lastRow - The last row to check
 * @returns {Promise<number>} - The new last row after deletions
 */
async function deleteGreenRows(worksheet, startRow, lastRow) {
    console.log(`Deleting green rows (#CCFFCC) in ${worksheet.name} from row ${startRow} to ${lastRow}`);
    
    try {
        // Create an array to store rows that need deletion (in descending order)
        const rowsToDelete = [];
        
        // Process each row individually instead of as a range to avoid collection issues
        for (let rowNum = startRow; rowNum <= lastRow; rowNum++) {
            const cellAddress = `B${rowNum}`;
            const cell = worksheet.getRange(cellAddress);
            cell.load("format/fill/color");
            
            try {
                await worksheet.context.sync();
                
                // Safely check if properties exist and if color matches
                if (cell.format && 
                    cell.format.fill && 
                    cell.format.fill.color === "#CCFFCC") {
                    rowsToDelete.push(rowNum);
                }
            } catch (cellError) {
                console.warn(`Error checking color for ${cellAddress}: ${cellError.message}`);
                // Continue to next cell if there's an error with this one
            }
        }
        
        // Sort in descending order to delete from bottom to top
        rowsToDelete.sort((a, b) => b - a);
        
        console.log(`Found ${rowsToDelete.length} green rows to delete`);
        
        // Delete each row (from bottom to top)
        if (rowsToDelete.length > 0) {
            for (const rowNum of rowsToDelete) {
                console.log(`Deleting row ${rowNum}`);
                const rowRange = worksheet.getRange(`${rowNum}:${rowNum}`);
                rowRange.delete(Excel.DeleteShiftDirection.up);
            }
            
            await worksheet.context.sync();
            
            // Recalculate the last row
            const newLastRow = await getLastUsedRow(worksheet, "B");
            console.log(`New last row after deletions: ${newLastRow}`);
            
            return newLastRow;
        } else {
            console.log("No green rows found to delete");
            return lastRow; // Return original lastRow if no rows deleted
        }
    } catch (error) {
        console.error(`Error in deleteGreenRows: ${error.message}`, error);
        // Return the original lastRow on error
        return lastRow;
    }
}

/**
 * Sets the font color to white for all cells in column A
 * @param {Excel.Worksheet} worksheet - The worksheet to process
 * @param {number} startRow - The first row to format
 * @param {number} lastRow - The last row to format
 * @returns {Promise<void>}
 */
async function setColumnAFontWhite(worksheet, startRow, lastRow) {
    console.log(`Setting font color to white in column A for ${worksheet.name} from row ${startRow} to ${lastRow}`);
    
    try {
        // Get the entire range for column A from startRow to lastRow
        const columnARange = worksheet.getRange(`A${startRow}:A${lastRow}`);
        
        // Set the font color to white
        columnARange.format.font.color = "#FFFFFF";
        
        await worksheet.context.sync();
        console.log(`Successfully set font color to white in column A for rows ${startRow}-${lastRow}`);
    } catch (error) {
        console.error(`Error in setColumnAFontWhite: ${error.message}`, error);
    }
}

// --- Helper Functions for Column Conversion ---

/**
 * Converts a 0-based column index into a column letter (e.g., 0 -> A, 1 -> B, 26 -> AA).
 * @param {number} index - The 0-based column index.
 * @returns {string} The column letter.
 */
function columnIndexToLetter(index) {
    let letter = '';
    while (index >= 0) {
        letter = String.fromCharCode(index % 26 + 'A'.charCodeAt(0)) + letter;
        index = Math.floor(index / 26) - 1;
    }
    return letter;
}

/**
 * Converts a column letter into a 0-based column index (e.g., A -> 0, B -> 1, AA -> 26).
 * @param {string} letter - The column letter (case-insensitive).
 * @returns {number} The 0-based column index.
 */
function columnLetterToIndex(letter) {
    letter = letter.toUpperCase();
    let index = 0;
    for (let i = 0; i < letter.length; i++) {
        index = index * 26 + (letter.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
    }
    return index - 1; // Adjust to 0-based
}

/**
 * Hides specific rows and columns on all worksheets except for specified exclusions.
 * Hides rows 1-8, columns C-I (3-9), and columns S-AC (19-29).
 * @param {string[]} excludedSheetNames - An array of sheet names to exclude from hiding.
 * @returns {Promise<void>}
 */
export async function hideRowsAndColumnsOnSheets(excludedSheetNames = ["Actuals Data", "Actuals Categorization"]) {
    try {
        console.log(`Hiding rows/columns on sheets, excluding: ${excludedSheetNames.join(', ')}`);

        await Excel.run(async (context) => {
            const worksheets = context.workbook.worksheets;
            worksheets.load("items/name");
            await context.sync();

            for (const worksheet of worksheets.items) {
                const sheetName = worksheet.name;
                if (excludedSheetNames.includes(sheetName)) {
                    console.log(`Skipping sheet: ${sheetName} (excluded)`);
                    continue;
                }

                console.log(`Processing sheet: ${sheetName}`);

                try {
                    // Hide Rows 1-8
                    const rowRange = worksheet.getRange("1:8");
                    rowRange.rowHidden = true;
                    console.log(`  Hiding rows 1-8`);

                    // Hide Columns C-E (Changed from C:I)
                    const colRange1 = worksheet.getRange("C:E"); // Changed range
                    colRange1.columnHidden = true;
                    console.log(`  Hiding columns C-E`); // Update log message

                    // Hide Columns S-AC
                    const colRange2 = worksheet.getRange("S:AC");
                    colRange2.columnHidden = true;
                    console.log(`  Hiding columns S-AC`);

                    // It's often more efficient to batch sync operations,
                    // but sometimes hiding needs immediate effect or separate syncs.
                    // Let's sync after hiding for this sheet.
                    await context.sync();
                    console.log(`  Finished hiding for ${sheetName}`);

                } catch (hideError) {
                    console.error(`  Error hiding rows/columns on sheet ${sheetName}: ${hideError.message}`, {
                        code: hideError.code,
                        debugInfo: hideError.debugInfo ? JSON.stringify(hideError.debugInfo) : 'N/A'
                    });
                    // Continue to the next sheet even if one fails
                }
            }

            console.log("Finished processing all sheets for hiding rows/columns.");
        }); // End Excel.run

    } catch (error) {
        console.error("Critical error in hideRowsAndColumnsOnSheets:", error);
        throw error;
    }
}

// TODO: Implement the actual logic within the JS helper functions (adjustDriversJS, replaceIndirectsJS, etc.).
// TODO: Implement findRowByValue helper function if Retained Earnings logic is needed.
// TODO: Update the calling code (e.g., button handler in taskpane.js) to call `processAssumptionTabs` after `runCodes`.

/**
 * Inserts worksheets from a base64-encoded Excel file into the current workbook
 * @param {string} base64String - Base64-encoded string of the source Excel file
 * @param {string[]} [sheetNames] - Optional array of sheet names to insert. If not provided, all sheets will be inserted.
 * @returns {Promise<void>}
 */
export async function handleInsertWorksheetsFromBase64(base64String, sheetNames = null) {
    try {
        // Validate base64 string
        if (!base64String || typeof base64String !== 'string') {
            throw new Error("Invalid base64 string provided");
        }

        // Validate base64 format
        if (!/^[A-Za-z0-9+/]*={0,2}$/.test(base64String)) {
            throw new Error("Invalid base64 format");
        }

        await Excel.run(async (context) => {
            const workbook = context.workbook;
            
            // Check if we have the required API version
            if (!workbook.insertWorksheetsFromBase64) {
                throw new Error("This feature requires Excel API requirement set 1.13 or later");
            }
            
            // Insert the worksheets with error handling
            try {
                console.log(`[SHEET OPERATION] Inserting worksheets from base64 string. Sheet names: ${sheetNames ? sheetNames.join(', ') : 'All sheets from source file'}`);
                await workbook.insertWorksheetsFromBase64(base64String, {
                    sheetNames: sheetNames
                });
                
                await context.sync();
                console.log("Worksheets inserted successfully");
                console.log(`[SHEET OPERATION] Successfully inserted worksheets: ${sheetNames ? sheetNames.join(', ') : 'All sheets from source file'}`);
            } catch (error) {
                console.error("Error during worksheet insertion:", error);
                throw new Error(`Failed to insert worksheets: ${error.message}`);
            }
        });
    } catch (error) {
        console.error("Error inserting worksheets from base64:", error);
        throw error;
    }
}

/**
 * Applies the Index Growth Curve logic to a worksheet, mimicking VBA Function IndexGrowthCurve.
 * Finds INDEXBEGIN/INDEXEND, inserts rows, populates data and formulas, applies formatting.
 * @param {Excel.Worksheet} worksheet - The assumption worksheet (within an Excel.run context).
 * @param {number} initialLastRow - The last row determined before this function runs.
 */
async function applyIndexGrowthCurveJS(worksheet, initialLastRow) {
    console.log(`Running applyIndexGrowthCurveJS for sheet: ${worksheet.name}`);
    const START_ROW = 10; // Row to start searching for INDEXBEGIN
    const BEGIN_MARKER = "INDEXBEGIN";
    const END_MARKER = "INDEXEND";
    const DATA_COL = "C";
    const SEARCH_COL = "D";
    const OUTPUT_COL_B = "B";
    const OUTPUT_COL_C = "C";
    const OUTPUT_COL_D = "D";
    const CHECK_COL_B = "B"; // Column B for green check
    const VALUE_COL_A = "A"; // Column A for BS/AV check
    const DRIVER_REF_COL = "AE"; // Column containing driver range ref in END_MARKER row
    const SUMIF_START_COL = "K"; // K
    const SUMIF_END_COL = "P"; // P
    const SUMPRODUCT_COL = "AE"; // AE (VBA used AE, not S)
    const MONTHS_START_COL = "AE"; // AE
    const MONTHS_END_COL = "CX"; // CX
    const LIGHT_BLUE_COLOR = "#D9E1F2"; // RGB(217, 225, 242)
    const LIGHT_GREEN_COLOR = "#CCFFCC"; // RGB(204, 255, 204)
 
    try {
        // Re-get worksheet reference within this context to ensure freshness
        const context = worksheet.context; // Get context from the passed object
        const worksheetName = worksheet.name; // Get name from potentially stale object
        const currentWorksheet = context.workbook.worksheets.getItem(worksheetName);
        // We assume the context itself is valid for this entire Excel.run block
 
        // --- 1. Find INDEXBEGIN and INDEXEND rows ---
        console.log(`Searching for ${BEGIN_MARKER} and ${END_MARKER} in column ${SEARCH_COL} of ${worksheetName}`);
        const searchRangeAddress = `${SEARCH_COL}${START_ROW}:${SEARCH_COL}${initialLastRow}`; 
        const searchRange = currentWorksheet.getRange(searchRangeAddress); // Use refreshed worksheet object
        searchRange.load("values");
        await context.sync(); // Use the context variable
 
        let firstRow = -1;
        let lastRow = -1;
        let indexEndRow = -1; // Keep track of the original END_MARKER row
 
        if (searchRange.values) {
            for (let i = 0; i < searchRange.values.length; i++) {
                const currentRow = START_ROW + i;
                const cellValue = searchRange.values[i][0];
                if (cellValue === BEGIN_MARKER && firstRow === -1) {
                    firstRow = currentRow;
                }
                if (cellValue === END_MARKER) {
                    lastRow = currentRow; // This will be the last END_MARKER found
                    indexEndRow = currentRow; // Store the original row index
                }
            }
        }
 
        if (firstRow === -1 || lastRow === -1 || lastRow < firstRow) {
            console.log(`Markers ${BEGIN_MARKER}/${END_MARKER} not found or in wrong order in ${searchRangeAddress}. Skipping Index Growth Curve.`);
            return; // Exit if markers not found or invalid
        }
        console.log(`Found ${BEGIN_MARKER} at row ${firstRow}, ${END_MARKER} at row ${lastRow}`);
 
        // --- 2. Collect Index Rows (Rows between markers where Col C is not empty) ---
        const indexRows = [];
        // CHANGE DATA_COL here if needed, e.g. const DATA_COL_TO_CHECK = "B";
        const DATA_COL_TO_CHECK = "B"; // Or "A", etc.
        const dataColRangeAddress = `${DATA_COL_TO_CHECK}${firstRow}:${DATA_COL_TO_CHECK}${lastRow}`;
        const dataColRange = currentWorksheet.getRange(dataColRangeAddress);
        // ... rest of the loading and checking logic ...
 
        if (indexRows.length === 0) {
            console.log(`No data rows found between ${BEGIN_MARKER} and ${END_MARKER} in column ${DATA_COL}. Skipping rest of Index Growth Curve.`);
            return; // Exit if no data rows found
        }
        console.log(`Collected ${indexRows.length} index rows:`, indexRows);
 
        // --- 3. Set Background Color for non-green rows ---
        // Range: B(firstRow+2) to CX(lastRow-2) in VBA, but logic only checks B color. Let's adjust row color based on B.
        const formatCheckStartRow = firstRow + 2;
        const formatCheckEndRow = lastRow - 2;
        console.log(`Setting background color for non-green rows between ${formatCheckStartRow} and ${formatCheckEndRow}`);
        if (formatCheckStartRow <= formatCheckEndRow) {
             // Load colors first
             const checkColorRange = currentWorksheet.getRange(`${CHECK_COL_B}${formatCheckStartRow}:${CHECK_COL_B}${formatCheckEndRow}`);
             checkColorRange.load("format/fill/color");
             await context.sync();
 
              // Queue formatting changes
              for (let i = 0; i < checkColorRange.values.length; i++) { // checkColorRange.values isn't loaded, use index
                 const currentRow = formatCheckStartRow + i;
                  // Use loaded format object
                 if (checkColorRange.format.fill.color !== LIGHT_GREEN_COLOR) {
                     console.log(`  Setting row ${currentRow} background to ${LIGHT_BLUE_COLOR}`);
                     const rowRange = currentWorksheet.getRange(`${currentRow}:${currentRow}`);
                     rowRange.format.fill.color = LIGHT_BLUE_COLOR;
                     // Clear fill in column A specifically
                     const cellARange = currentWorksheet.getRange(`A${currentRow}`);
                     cellARange.format.fill.clear();
                 }
              }
         }
 
         // --- 4. Insert Rows ---
         const newRowStart = lastRow + 2;
         const numNewRows = indexRows.length;
         const newRowEnd = newRowStart + numNewRows - 1;
         console.log(`Inserting ${numNewRows} rows at range ${newRowStart}:${newRowEnd}`);
         const insertRange = currentWorksheet.getRange(`${newRowStart}:${newRowEnd}`);
         insertRange.insert(Excel.InsertShiftDirection.down);
         // Sync required before populating new rows
         await context.sync();
 
         // --- 5. Populate New Rows (B, C, D) ---
         console.log(`Populating columns ${OUTPUT_COL_B}, ${OUTPUT_COL_C}, ${OUTPUT_COL_D} in new rows ${newRowStart}:${newRowEnd}`);
         // Load source data from original index rows
         const sourceDataAddresses = indexRows.map(r => `${OUTPUT_COL_B}${r}:${OUTPUT_COL_C}${r}`);
         // Cannot load disjoint ranges easily this way. Load columns B and C for the whole original block.
         const sourceBlockRange = currentWorksheet.getRange(`${OUTPUT_COL_B}${firstRow}:${OUTPUT_COL_C}${lastRow}`);
         sourceBlockRange.load("values");
         await context.sync();
 
         const outputDataBC = [];
         const outputDataD = [];
         const sourceValues = sourceBlockRange.values;
         for (const originalRow of indexRows) {
             const rowIndexInBlock = originalRow - firstRow; // 0-based index within the loaded block
             const valB = sourceValues[rowIndexInBlock][0]; // Col B is index 0
             const valC = sourceValues[rowIndexInBlock][1]; // Col C is index 1
             outputDataBC.push([valB, valC]);
             outputDataD.push([END_MARKER]);
         }
 
         const outputRangeBC = currentWorksheet.getRange(`${OUTPUT_COL_B}${newRowStart}:${OUTPUT_COL_C}${newRowEnd}`);
         outputRangeBC.values = outputDataBC;
         const outputRangeD = currentWorksheet.getRange(`${OUTPUT_COL_D}${newRowStart}:${OUTPUT_COL_D}${newRowEnd}`);
         outputRangeD.values = outputDataD;
 
         // --- 6. Apply SUMIF Formulas (K-P) ---
         console.log(`Applying SUMIF formulas to ${SUMIF_START_COL}${newRowStart}:${SUMIF_END_COL}${newRowEnd}`);
         // Load necessary data: Col C and Col A values from original index rows
         const sourceColCRange = currentWorksheet.getRange(`${DATA_COL}${firstRow}:${DATA_COL}${lastRow}`);
         const sourceColARange = currentWorksheet.getRange(`${VALUE_COL_A}${firstRow}:${VALUE_COL_A}${lastRow}`);
         sourceColCRange.load("values");
         sourceColARange.load("values");
         await context.sync();
 
         const sourceValuesC = sourceColCRange.values;
         const sourceValuesA = sourceColARange.values;
         const numSumifCols = columnLetterToIndex(SUMIF_END_COL) - columnLetterToIndex(SUMIF_START_COL) + 1;
         const sumifFormulas = [];
 
         for (let i = 0; i < indexRows.length; i++) {
             const originalRow = indexRows[i];
             const rowIndexInBlock = originalRow - firstRow; // 0-based index within the loaded block
             const codeC = sourceValuesC[rowIndexInBlock][0] || ""; // Ensure string
             const valueA = sourceValuesA[rowIndexInBlock][0];
             const targetRowNum = newRowStart + i; // Row where formula will be placed
 
             let baseFormula;
             // Check if Col C starts with "BS" or Col A is "AV"
             if (codeC.toUpperCase().startsWith("BS") || String(valueA).toUpperCase() === "AV") {
                 // "=SUMIF($3:$3,@ INDIRECT(ADDRESS(ROW($A$2),COLUMN(),2)), INDIRECT(ROW() & "":"" & ROW()))"
                  baseFormula = `=SUMIF($3:$3, INDIRECT(ADDRESS(2,COLUMN())), ${targetRowNum}:${targetRowNum})`;
             } else {
                 // "=SUMIF($4:$4,@ INDIRECT(ADDRESS(ROW($A$2),COLUMN(),2)), INDIRECT(ROW() & "":"" & ROW()))"
                  baseFormula = `=SUMIF($4:$4, INDIRECT(ADDRESS(2,COLUMN())), ${targetRowNum}:${targetRowNum})`;
             }
             // Create array for the row
             sumifFormulas.push(Array(numSumifCols).fill(baseFormula));
         }
 
         const sumifRange = currentWorksheet.getRange(`${SUMIF_START_COL}${newRowStart}:${SUMIF_END_COL}${newRowEnd}`);
         sumifRange.formulas = sumifFormulas;
 
         // --- 7. Apply SUMPRODUCT Formulas (AE) ---
         console.log(`Applying SUMPRODUCT formulas to ${SUMPRODUCT_COL}${newRowStart}:${SUMPRODUCT_COL}${newRowEnd}`);
         // Get the driver range string from the original END_MARKER row, column AE
         const driverCell = currentWorksheet.getRange(`${DRIVER_REF_COL}${indexEndRow}`);
         driverCell.load("values");
         await context.sync();
         const driverRangeString = driverCell.values[0][0];
 
         if (!driverRangeString || typeof driverRangeString !== 'string') {
             console.warn(`Driver range string not found or invalid in cell ${DRIVER_REF_COL}${indexEndRow}. Skipping SUMPRODUCT.`);
         } else {
              console.log(`Using driver range: ${driverRangeString}`);
              // Iterate and set formula for each cell individually (mimics FormulaArray)
              for (let i = 0; i < indexRows.length; i++) {
                  const originalRow = indexRows[i];
                  const targetRow = newRowStart + i;
                  const dataRangeString = `$${MONTHS_START_COL}$${originalRow}:$${MONTHS_END_COL}$${originalRow}`;
                  // Formula: =SUMPRODUCT(INDEX(driverRange, N(IF({1}, MAX(COLUMN(driverRange)) - COLUMN(driverRange) + 1))), dataRange)
                  const sumproductFormula = `=SUMPRODUCT(INDEX(${driverRangeString},N(IF({1},MAX(COLUMN(${driverRangeString}))-COLUMN(${driverRangeString})+1))), ${dataRangeString})`;
 
                  const targetCell = currentWorksheet.getRange(`${SUMPRODUCT_COL}${targetRow}`);
                  targetCell.formulas = [[sumproductFormula]];
                  // console.log(`  Set formula for ${SUMPRODUCT_COL}${targetRow}: ${sumproductFormula}`);
             }
         }
 
         // --- 8. Copy Formats and Adjust ---
         console.log(`Copying formats and adjusting for new rows ${newRowStart}:${newRowEnd}`);
         for (let i = 0; i < indexRows.length; i++) {
             const sourceRow = indexRows[i];
             const targetRow = newRowStart + i;
 
             const sourceRowRange = currentWorksheet.getRange(`${sourceRow}:${sourceRow}`);
             const targetRowRange = currentWorksheet.getRange(`${targetRow}:${targetRow}`);
 
             // Copy formats first
             targetRowRange.copyFrom(sourceRowRange, Excel.RangeCopyType.formats);
              await context.sync(); // Sync after each copy maybe needed? Let's try one sync after loop.
 
             // Apply format overrides
             targetRowRange.format.font.color = "#000000"; // Black font
             targetRowRange.format.borders.load('items'); // Load borders collection
              await context.sync(); // Need to sync load before clearing
 
              targetRowRange.format.borders.items.forEach(border => border.style = 'None');
             // Explicitly clear all borders (simpler?)
             // targetRowRange.format.borders.getItem('EdgeTop').style = 'None';
             // targetRowRange.format.borders.getItem('EdgeBottom').style = 'None';
             // targetRowRange.format.borders.getItem('EdgeLeft').style = 'None';
             // targetRowRange.format.borders.getItem('EdgeRight').style = 'None';
             // targetRowRange.format.borders.getItem('InsideVertical').style = 'None';
             // targetRowRange.format.borders.getItem('InsideHorizontal').style = 'None';
 
             targetRowRange.format.fill.clear(); // Clear interior color
             targetRowRange.format.font.bold = false; // Remove bold
 
             // Set indent level for column B
             const targetCellB = currentWorksheet.getRange(`${OUTPUT_COL_B}${targetRow}`);
             targetCellB.format.indentLevel = 2;
         }
          await context.sync(); // Sync format changes
 
         // --- 9. Clear Original Column C values ---
         console.log(`Clearing values in original index rows (${indexRows.join(', ')}) column ${DATA_COL}`);
         // It's safer to clear individually if rows aren't contiguous
         for (const originalRow of indexRows) {
             const cellToClear = currentWorksheet.getRange(`${DATA_COL}${originalRow}`);
             cellToClear.clear(Excel.ClearApplyTo.contents);
         }
          await context.sync(); // Sync clears
 
         console.log(`applyIndexGrowthCurveJS completed successfully for sheet: ${worksheetName}`);
 
     } catch (error) {
         console.error(`Error in applyIndexGrowthCurveJS for sheet ${worksheet.name}:`, error);
         // Decide if error should be re-thrown
         // throw error; // Optional: re-throw to stop processAssumptionTabs if critical
     }
     // Note: This function runs within the context of the calling Excel.run in processAssumptionTabs.
     // Syncs are added within the function for critical steps like after insertion.
 }

/**
 * Parses and converts FIXED-E/FIXED-S formula syntax to Excel formula
 * @param {string} formulaString - The formula string from the code parameter (e.g., "DOLANN(C3)*BEG(C2)*END(C1)")
 * @param {number} targetRow - The row number where the formula will be placed
 * @returns {string} - The converted Excel formula
 */

/**
 * Applies formatting to a column range based on format type
 * @param {Excel.Worksheet} worksheet - The worksheet to apply formatting to
 * @param {string} column - The column letter
 * @param {number} startRow - The starting row
 * @param {number} endRow - The ending row
 * @param {string} formatType - The format type (dollaritalic, date, etc.)
 * @returns {Promise<void>}
 */
async function applyColumnFormatting(worksheet, column, startRow, endRow, formatType) {
    const rangeAddress = `${column}${startRow}:${column}${endRow}`;
    const range = worksheet.getRange(rangeAddress);
    
    let numberFormatString = null;
    let applyItalics = false;
    
    switch (formatType) {
        case "dollaritalic":
            numberFormatString = '_(* $ #,##0_);_(* $ (#,##0);_(* "$" -""?_);_(@_)';
            applyItalics = true;
            break;
        case "dollar":
            numberFormatString = '_(* $ #,##0_);_(* $ (#,##0);_(* "$" -""?_);_(@_)';
            applyItalics = false;
            break;
        case "date":
            numberFormatString = 'mmm-yy';
            applyItalics = false;
            break;
        case "volume":
            numberFormatString = '_(* #,##0_);_(* (#,##0);_(* " -"?_);_(@_)';
            applyItalics = true;
            break;
        case "percent":
            numberFormatString = '_(* #,##0.0%;_(* (#,##0.0)%;_(* " -"?_)';
            applyItalics = true;
            break;
        case "factor":
            numberFormatString = '_(* #,##0.0x;_(* (#,##0.0)x;_(* "   -"?_)';
            applyItalics = true;
            break;
    }
    
    if (numberFormatString) {
        console.log(`Applying ${formatType} format to ${rangeAddress}`);
        range.numberFormat = [[numberFormatString]];
        range.format.font.italic = applyItalics;
    }
}

/**
 * Applies FIXED-E or FIXED-S formula to column AE when code type matches and formula parameter exists
 * @param {Excel.Worksheet} worksheet - The worksheet to apply the formula to
 * @param {number} pasteRow - The starting row where the code was pasted
 * @param {number} lastRow - The last row from the Codes worksheet
 * @param {number} firstRow - The first row from the Codes worksheet  
 * @param {Object} code - The code object with type and parameters
 * @returns {Promise<void>}
 */

/**
 * Processes FORMULA-S rows by converting driver references in column AE to Excel formulas
 * @param {Excel.Worksheet} worksheet - The worksheet to process
 * @param {number} startRow - The starting row to search from
 * @param {number} lastRow - The last row to search to
 * @returns {Promise<void>}
 */
async function processFormulaSRows(worksheet, startRow, lastRow) {
    console.log(`Processing FORMULA-S rows in ${worksheet.name} from row ${startRow} to ${lastRow}`);
    
    try {
        // Get stored FORMULA-S row positions (since column D may have been overwritten)
        const formulaSRows = getFormulaSRows(worksheet.name);
        
        if (formulaSRows.length === 0) {
            console.log("No FORMULA-S rows found in tracker");
            return;
        }
        
        console.log(`Found ${formulaSRows.length} FORMULA-S rows at: ${formulaSRows.join(", ")}`);
        
        // Load column A values to create a driver lookup map
        const colARangeAddress = `A${startRow}:A${lastRow}`;
        const colARange = worksheet.getRange(colARangeAddress);
        colARange.load("values");
        await worksheet.context.sync();
        
        const colAValues = colARange.values;
        const driverMap = new Map();
        
        // Check each row individually for green background color to filter driver map
        const greenRows = new Set();
        for (let i = 0; i < colAValues.length; i++) {
            const rowNum = startRow + i;
            try {
                const cellB = worksheet.getRange(`B${rowNum}`);
                cellB.load('format/fill/color');
                await worksheet.context.sync();
                
                if (cellB.format && cellB.format.fill && cellB.format.fill.color === '#CCFFCC') {
                    greenRows.add(rowNum);
                }
            } catch (colorError) {
                // If we can't check color, assume it's not green
                console.warn(`  Could not check color for row ${rowNum}, assuming not green`);
            }
        }
        
        // Build driver map: driver name -> row number (excluding green rows)
        for (let i = 0; i < colAValues.length; i++) {
            const value = colAValues[i][0];
            const rowNum = startRow + i;
            
            if (value !== null && value !== "" && !greenRows.has(rowNum)) {
                driverMap.set(String(value), rowNum);
                console.log(`  Driver map: ${value} -> row ${rowNum} (non-green)`);
            } else if (value !== null && value !== "" && greenRows.has(rowNum)) {
                console.log(`  Skipping green row driver: ${value} at row ${rowNum} (will be deleted)`);
            }
        }
        
        // Define column mapping for cd: 1=D, 2=E, 3=F, etc. (excluding A, B, C)
        const columnMapping = {
            '1': 'D',
            '2': 'E',
            '3': 'F',
            '4': 'G',
            '5': 'H',
            '6': 'I',
            '7': 'J',
            '8': 'K',
            '9': 'L',
            '10': 'M',
            '11': 'N',
            '12': 'O',
            '13': 'P',
            '14': 'Q',
            '15': 'R',
            '16': 'S'
        };
        
        // Process each FORMULA-S row
        for (const rowNum of formulaSRows) {
            // Get the current value in column AE
            const aeCell = worksheet.getRange(`AE${rowNum}`);
            aeCell.load("values");
            await worksheet.context.sync();
            
            const originalValue = aeCell.values[0][0];
            if (!originalValue || originalValue === "") {
                console.log(`  Row ${rowNum}: No value in AE, skipping`);
                continue;
            }
            
            console.log(`  Row ${rowNum}: Processing formula string: "${originalValue}"`);
            
            // Convert the string to a formula
            let formula = String(originalValue);
            
            // Replace all rd{driverName} patterns with cell references
            const rdPattern = /rd\{([^}]+)\}/g;
            formula = formula.replace(rdPattern, (match, driverName) => {
                const driverRow = driverMap.get(driverName);
                if (driverRow) {
                    const replacement = `AE$${driverRow}`;
                    console.log(`    Replacing rd{${driverName}} with ${replacement}`);
                    return replacement;
                } else {
                    console.warn(`    Driver '${driverName}' not found in column A, keeping as is`);
                    return match; // Keep the original if driver not found
                }
            });
            
                    // Replace all cd{number} or cd{number-driverName} or cd{driverName-number} patterns with column references
        const cdPattern = /cd\{([^}]+)\}/g;
        formula = formula.replace(cdPattern, (match, content) => {
            // Check if content contains a dash
            const dashIndex = content.indexOf('-');
            let columnNum, driverName;
            
            if (dashIndex !== -1) {
                const firstPart = content.substring(0, dashIndex).trim();
                const secondPart = content.substring(dashIndex + 1).trim();
                
                // Check if first part is a number to determine syntax format
                if (/^\d+$/.test(firstPart)) {
                    // Format: {columnNumber-driverName} (existing syntax)
                    columnNum = firstPart;
                    driverName = secondPart;
                    console.log(`    Parsing cd{${content}} as columnNumber-driverName format`);
                } else {
                    // Format: {driverName-columnNumber} (new syntax)
                    driverName = firstPart;
                    columnNum = secondPart;
                    console.log(`    Parsing cd{${content}} as driverName-columnNumber format`);
                }
            } else {
                // Just column number, no driver name
                columnNum = content.trim();
                driverName = null;
                console.log(`    Parsing cd{${content}} as columnNumber only format`);
            }
            
            const column = columnMapping[columnNum];
            if (!column) {
                console.warn(`    Column number '${columnNum}' not valid (must be 1-16), keeping as is`);
                return match; // Keep the original if column number not valid
            }
            
            // Determine which row to use
            let rowToUse = rowNum;
            if (driverName) {
                const driverRow = driverMap.get(driverName);
                if (driverRow) {
                    rowToUse = driverRow;
                    console.log(`    Replacing cd{${content}} with $${column}${rowToUse} (driver '${driverName}' found at row ${driverRow})`);
                } else {
                    console.warn(`    Driver '${driverName}' not found in column A, using current row ${rowNum}`);
                }
            } else {
                console.log(`    Replacing cd{${columnNum}} with $${column}${rowToUse}`);
            }
            
            const replacement = `$${column}${rowToUse}`;
            return replacement;
        });
            
            // Replace timeseriesdivisor with AE$7
            formula = formula.replace(/timeseriesdivisor/gi, (match) => {
                const replacement = 'AE$7';
                console.log(`    Replacing ${match} with ${replacement}`);
                return replacement;
            });
            
            // Replace currentmonth with AE$2
            formula = formula.replace(/currentmonth/gi, (match) => {
                const replacement = 'AE$2';
                console.log(`    Replacing ${match} with ${replacement}`);
                return replacement;
            });
            
            // Replace beginningmonth with $AE$2
            formula = formula.replace(/beginningmonth/gi, (match) => {
                const replacement = '$AE$2';
                console.log(`    Replacing ${match} with ${replacement}`);
                return replacement;
            });
            
            // Replace currentyear with AE$3
            formula = formula.replace(/currentyear/gi, (match) => {
                const replacement = 'AE$3';
                console.log(`    Replacing ${match} with ${replacement}`);
                return replacement;
            });
            
            // Replace yearend with AE$4
            formula = formula.replace(/yearend/gi, (match) => {
                const replacement = 'AE$4';
                console.log(`    Replacing ${match} with ${replacement}`);
                return replacement;
            });
            
            // Replace beginningyear with $AE$2
            formula = formula.replace(/beginningyear/gi, (match) => {
                const replacement = '$AE$3';
                console.log(`    Replacing ${match} with ${replacement}`);
                return replacement;
            });
            
            // Now parse special functions (SPREAD, BEG, END, etc.)
            console.log(`  Parsing special functions in formula...`);
            
            // Process SPREAD function: SPREAD(driver) -> driver/AE$7
            formula = formula.replace(/SPREAD\(([^)]+)\)/gi, (match, driver) => {
                console.log(`    Converting SPREAD(${driver}) to ${driver}/AE$7`);
                return `(${driver}/AE$7)`;
            });
            
            // Process BEG function: BEG(driver) -> (EOMONTH(driver,0)<=AE$2)
            formula = formula.replace(/BEG\(([^)]+)\)/gi, (match, driver) => {
                console.log(`    Converting BEG(${driver}) to (EOMONTH(${driver},0)<=AE$2)`);
                return `(EOMONTH(${driver},0)<=AE$2)`;
            });
            
            // Process END function: END(driver) -> (EOMONTH(driver,0)>AE$2)
            formula = formula.replace(/END\(([^)]+)\)/gi, (match, driver) => {
                console.log(`    Converting END(${driver}) to (EOMONTH(${driver},0)>AE$2)`);
                return `(EOMONTH(${driver},0)>AE$2)`;
            });
            
            // Process RAISE function: RAISE(driver1,driver2) -> (1 + (driver1)) ^ (AE$3 - max(year(driver2), $AE3))
            formula = formula.replace(/RAISE\(([^,]+),([^)]+)\)/gi, (match, driver1, driver2) => {
                // Trim whitespace from drivers
                driver1 = driver1.trim();
                driver2 = driver2.trim();
                console.log(`    Converting RAISE(${driver1},${driver2}) to (1 + (${driver1})) ^ (AE$3 - max(year(${driver2}), $AE3))`);
                return `(1 + (${driver1})) ^ (AE$3 - max(year(${driver2}), $AE3))`;
            });
            
            // Process ONETIMEDATE function: ONETIMEDATE(driver) -> (EOMONTH((driver),0)=AE$2)
            formula = formula.replace(/ONETIMEDATE\(([^)]+)\)/gi, (match, driver) => {
                console.log(`    Converting ONETIMEDATE(${driver}) to (EOMONTH((${driver}),0)=AE$2)`);
                return `(EOMONTH((${driver}),0)=AE$2)`;
            });
            
            // Process SPREADDATES function: SPREADDATES(driver1,driver2,driver3) -> IF(AND(EOMONTH(AE$2,0)>=EOMONTH(driver2,0),EOMONTH(driver3,0)<=EOMONTH($I12,0)),driver1/(DATEDIF(driver2,driver3,"m")+1),0)
            // Note: Need to handle nested parentheses and comma separation
            formula = formula.replace(/SPREADDATES\(([^,]+),([^,]+),([^)]+)\)/gi, (match, driver1, driver2, driver3) => {
                // Trim whitespace from drivers
                driver1 = driver1.trim();
                driver2 = driver2.trim();
                driver3 = driver3.trim();
                const newFormula = `IF(AND(EOMONTH(AE$2,0)>=EOMONTH(${driver2},0),EOMONTH(AE$2,0)<=EOMONTH(${driver3},0)),${driver1}/(DATEDIF(${driver2},${driver3},"m")+1),0)`;
                console.log(`    Converting SPREADDATES(${driver1},${driver2},${driver3}) to ${newFormula}`);
                return newFormula;
            });
            
            // Ensure the formula starts with '='
            if (!formula.startsWith('=')) {
                formula = '=' + formula;
            }
            
            console.log(`  Row ${rowNum}: Final formula: ${formula}`);
            
            // Set the formula in the cell
            aeCell.formulas = [[formula]];
        }
        
        // Sync all formula changes
        await worksheet.context.sync();
        console.log(`Successfully processed ${formulaSRows.length} FORMULA-S rows`);
        
    } catch (error) {
        console.error(`Error in processFormulaSRows: ${error.message}`, error);
        throw error;
    }
}

/**
 * Hides Columns C-I, Rows 2-8, and specific Actuals columns on specified sheets,
 * then navigates to cell A1 of the Financials sheet.
 * @param {string[]} assumptionTabNames - Array of assumption tab names created by runCodes.
 * @returns {Promise<void>}
 */
export async function hideColumnsAndNavigate(assumptionTabNames) { // Renamed and added parameter
    // Define Actuals columns
    const ACTUALS_START_COL = "S";
    const ACTUALS_END_COL = "AD";

    try {
        const targetSheetNames = [...assumptionTabNames, "Financials"]; // Combine assumption tabs and Financials
        console.log(`Attempting to hide specific rows/columns on sheets [${targetSheetNames.join(', ')}] and navigate...`);

        await Excel.run(async (context) => {
            // Get all worksheets
            const worksheets = context.workbook.worksheets;
            // Load only names needed for matching
            worksheets.load("items/name");
            await context.sync();

            console.log(`Found ${worksheets.items.length} worksheets. Targeting ${targetSheetNames.length} specific sheets.`);
            let hideAttempted = false;

            // Calculate actuals end column for assumption tabs
            const actualsEndIndex = columnLetterToIndex(ACTUALS_END_COL);
            const actualsEndMinusOneCol = actualsEndIndex > 0 ? columnIndexToLetter(actualsEndIndex - 1) : ACTUALS_START_COL; // Handle edge case

            // --- Queue hiding operations for target sheets ---
            for (const worksheet of worksheets.items) {
                const sheetName = worksheet.name;
                if (targetSheetNames.includes(sheetName)) { // Check if sheet is in our target list
                    console.log(`Queueing hide operations for: ${sheetName}`);
                    try {
                        // Hide Rows 2:8 (Applies to both)
                        const rows28 = worksheet.getRange("2:8");
                        rows28.rowHidden = true;

                        // Conditional Column Hiding
                        if (sheetName === "Financials") {
                            console.log(`  -> Hiding Columns C:I for Financials`);
                            const colsCI = worksheet.getRange("C:I");
                            colsCI.columnHidden = true;
                        } else {
                            // Hide Columns C:E for Assumption Tabs
                            console.log(`  -> Hiding Columns C:E for ${sheetName}`);
                            const colsCE = worksheet.getRange("C:E");
                            colsCE.columnHidden = true;
                        }

                        // Hide Actuals Columns based on sheet type
                        if (sheetName === "Financials") {
                            console.log(`  -> Hiding Actuals range ${ACTUALS_START_COL}:${ACTUALS_END_COL}`);
                            const actualsRangeFin = worksheet.getRange(`${ACTUALS_START_COL}:${ACTUALS_END_COL}`);
                            actualsRangeFin.columnHidden = true;
                        } else if (assumptionTabNames.includes(sheetName)) {
                             console.log(`  -> Hiding Actuals range ${ACTUALS_START_COL}:${actualsEndMinusOneCol}`);
                             const actualsRangeAssum = worksheet.getRange(`${ACTUALS_START_COL}:${actualsEndMinusOneCol}`);
                             actualsRangeAssum.columnHidden = true;
                        }

                        hideAttempted = true; // Mark that at least one hide was queued
                    } catch (error) {
                        // Log unexpected errors during the queuing attempt
                        console.error(`  Error queuing hide operations for ${sheetName}: ${error.message}`, {
                            code: error.code,
                            debugInfo: error.debugInfo ? JSON.stringify(error.debugInfo) : 'N/A'
                        });
                    }
                }
            }

            // --- Sync all queued hide operations ---
            if (hideAttempted) {
                console.log(`Attempting to sync hide columns/rows operations...`);
                try {
                    await context.sync();
                    console.log("Successfully synced hide columns/rows operations.");
                } catch (syncError) {
                    console.error(`Error syncing hide columns/rows operations: ${syncError.message}`, {
                        code: syncError.code,
                        debugInfo: syncError.debugInfo ? JSON.stringify(syncError.debugInfo) : 'N/A'
                    });
                     // Report failure but continue to navigation attempt
                }
            } else {
                 console.log("No target sheets found or no hide operations were queued.");
            }

            // --- Activate and Select A1 on each assumption tab (mimic Ctrl+Home) ---
            console.log("Activating and selecting A1 on assumption tabs...");
            // Note: This requires syncing inside the loop for the activate/select effect per sheet
            for (const sheetName of assumptionTabNames) {
                try {
                    console.log(`  Activating and selecting A1 for: ${sheetName}`);
                    const worksheet = context.workbook.worksheets.getItem(sheetName);
                    worksheet.activate(); // Activate the sheet first
                    const rangeA1 = worksheet.getRange("A1");
                    rangeA1.select(); // Then select A1
                    await context.sync(); // Sync *immediately* to apply activation and selection for this sheet
                    console.log(`  Synced A1 view reset for ${sheetName}.`);
                } catch (error) {
                     console.error(`  Error resetting view for ${sheetName}: ${error.message}`);
                     // Optionally continue to the next sheet even if one fails
                }
            }
            // No final sync needed for this loop as it happens inside

            // --- Activate and Select J9 on each assumption tab ---
            console.log("Activating and selecting J9 on assumption tabs...");
            // Note: This requires syncing inside the loop for the activate/select effect per sheet
            for (const sheetName of assumptionTabNames) {
                try {
                    console.log(`  Activating and selecting J9 for: ${sheetName}`);
                    const worksheet = context.workbook.worksheets.getItem(sheetName);
                    worksheet.activate(); // Activate the sheet first
                    const rangeJ9 = worksheet.getRange("J9"); // Get J9
                    rangeJ9.select(); // Then select J9
                    await context.sync(); // Sync *immediately* to apply activation and selection for this sheet
                    console.log(`  Synced J9 view reset for ${sheetName}.`);
                } catch (error) {
                     console.error(`  Error resetting view to J9 for ${sheetName}: ${error.message}`);
                     // Optionally continue to the next sheet even if one fails
                }
            }
            // No final sync needed for this loop as it happens inside

            

            // --- Delete sheets that begin with "Codes" or "Calcs" ---
            console.log("Deleting sheets that begin with 'Codes' or 'Calcs'...");
            console.log("[SHEET OPERATION] Scanning for sheets to delete (those beginning with 'Codes' or 'Calcs')");
            try {
                // Get all worksheets again to ensure we have the latest list
                const allWorksheets = context.workbook.worksheets;
                allWorksheets.load("items/name");
                await context.sync();
                
                const sheetsToDelete = [];
                
                // Identify sheets to delete
                for (const worksheet of allWorksheets.items) {
                    const sheetName = worksheet.name;
                    if (sheetName.startsWith("Codes") || sheetName.startsWith("Calcs")) {
                        sheetsToDelete.push(sheetName);
                    }
                }
                
                // Delete identified sheets
                if (sheetsToDelete.length > 0) {
                    console.log(`Found ${sheetsToDelete.length} sheet(s) to delete: ${sheetsToDelete.join(', ')}`);
                    
                    for (const sheetName of sheetsToDelete) {
                        try {
                            const sheetToDelete = context.workbook.worksheets.getItem(sheetName);
                            console.log(`[SHEET OPERATION] Deleting sheet '${sheetName}'`);
                            sheetToDelete.delete();
                            console.log(`  Queued deletion of sheet: ${sheetName}`);
                        } catch (deleteError) {
                            console.error(`  Error queuing deletion of sheet ${sheetName}: ${deleteError.message}`);
                            // Continue with other deletions even if one fails
                        }
                    }
                    
                    // Sync all deletions
                    try {
                        await context.sync();
                        console.log(`Successfully deleted ${sheetsToDelete.length} sheet(s).`);
                    } catch (syncError) {
                        console.error(`Error syncing sheet deletions: ${syncError.message}`, {
                            code: syncError.code,
                            debugInfo: syncError.debugInfo ? JSON.stringify(syncError.debugInfo) : 'N/A'
                        });
                        // Continue even if sync fails
                    }
                } else {
                    console.log("No sheets found starting with 'Codes' or 'Calcs' to delete.");
                }
            } catch (deletionError) {
                console.error(`Error during sheet deletion process: ${deletionError.message}`, {
                    code: deletionError.code,
                    debugInfo: deletionError.debugInfo ? JSON.stringify(deletionError.debugInfo) : 'N/A'
                });
                // Do not throw here, allow the function to finish
            }

            // --- Reorder tabs according to the model plan ---
            console.log("Reordering tabs according to model plan...");
            try {
                // Define the desired tab order
                // First is always Financials, followed by assumption tabs in their creation order,
                // then any other existing sheets (like Misc., Actuals Data, etc.)
                const priorityOrder = ["Financials", ...assumptionTabNames];
                
                // Get all worksheets one more time to ensure we have the latest list after deletions
                const finalWorksheets = context.workbook.worksheets;
                finalWorksheets.load("items/name");
                await context.sync();
                
                // Create a list of all sheet names
                const allSheetNames = finalWorksheets.items.map(ws => ws.name);
                
                // Separate sheets into priority order and others
                const orderedSheets = [];
                const otherSheets = [];
                
                // Add sheets in priority order first
                for (const priorityName of priorityOrder) {
                    if (allSheetNames.includes(priorityName)) {
                        orderedSheets.push(priorityName);
                    }
                }
                
                // Add remaining sheets that aren't in priority order
                for (const sheetName of allSheetNames) {
                    if (!orderedSheets.includes(sheetName)) {
                        otherSheets.push(sheetName);
                    }
                }
                
                // Combine the lists
                const finalOrder = [...orderedSheets, ...otherSheets];
                console.log(`Final tab order: ${finalOrder.join(', ')}`);
                
                // Reorder sheets by setting their positions
                for (let i = 0; i < finalOrder.length; i++) {
                    try {
                        const worksheet = context.workbook.worksheets.getItem(finalOrder[i]);
                        worksheet.position = i;
                        console.log(`  Set ${finalOrder[i]} to position ${i}`);
                    } catch (positionError) {
                        console.error(`  Error setting position for sheet ${finalOrder[i]}: ${positionError.message}`);
                        // Continue with other sheets even if one fails
                    }
                }
                
                // Sync all position changes
                try {
                    await context.sync();
                    console.log("Successfully reordered all tabs.");
                } catch (syncError) {
                    console.error(`Error syncing tab reordering: ${syncError.message}`, {
                        code: syncError.code,
                        debugInfo: syncError.debugInfo ? JSON.stringify(syncError.debugInfo) : 'N/A'
                    });
                    // Continue even if sync fails
                }
                
            } catch (reorderError) {
                console.error(`Error during tab reordering process: ${reorderError.message}`, {
                    code: reorderError.code,
                    debugInfo: reorderError.debugInfo ? JSON.stringify(reorderError.debugInfo) : 'N/A'
                });
                // Do not throw here, allow the function to finish
            }

            // --- Navigate to Financials sheet and select cell J10 with view reset ---
            // (This ensures Financials is the final active sheet with proper view)
            try {
                console.log("Navigating to Financials sheet and selecting J10 with view reset...");
                const financialsSheet = context.workbook.worksheets.getItem("Financials");
                
                // Activate the Financials sheet
                financialsSheet.activate();
                
                // Get J10 range
                const rangeJ10 = financialsSheet.getRange("J10");
                
                // Select J10
                rangeJ10.select();
                
                // Reset the view to top of sheet (like CTRL+Home)
                try {
                    // First, select A1 to scroll to top-left
                    const rangeA1 = financialsSheet.getRange("A1");
                    rangeA1.select();
                    
                    // Sync to ensure the scroll happens
                    await context.sync();
                    
                    // Now select J10 as the final selection
                    rangeJ10.select();
                    
                    // Optional: Reset zoom to 100%
                    try {
                        const activeWindow = context.workbook.getActiveCell().getWorksheet().getActiveView();
                        if (activeWindow) {
                            activeWindow.zoomLevel = 100;
                        }
                    } catch (zoomError) {
                        console.log("Could not reset zoom level (requires Excel API 1.7+):", zoomError.message);
                    }
                } catch (viewError) {
                    console.log("Could not fully reset view:", viewError.message);
                    // Continue anyway - selecting J10 is the most important part
                }
                
                await context.sync(); // Sync the final state
                console.log("Successfully navigated to Financials!J10 and reset view to top.");
            } catch (navError) {
                console.error(`Error navigating to Financials sheet J10: ${navError.message}`, {
                    code: navError.code,
                    debugInfo: navError.debugInfo ? JSON.stringify(navError.debugInfo) : 'N/A'
                });
                // Do not throw here, allow the function to finish
            }

            console.log("Finished hideColumnsAndNavigate function.");

        }); // End Excel.run
    } catch (error) {
        // Catch errors from the Excel.run call itself
        console.error("Critical error in hideColumnsAndNavigate:", error);
        throw error; // Re-throw critical errors
    }
}

// Global storage for FORMULA-S row positions per worksheet
const formulaSRowTracker = new Map(); // Map<worksheetName, Set<rowNumber>>

/**
 * Stores a FORMULA-S row position for later processing
 * @param {string} worksheetName - Name of the worksheet
 * @param {number} rowNumber - Row number containing customformula
 */
function addFormulaSRow(worksheetName, rowNumber) {
    if (!formulaSRowTracker.has(worksheetName)) {
        formulaSRowTracker.set(worksheetName, new Set());
    }
    formulaSRowTracker.get(worksheetName).add(rowNumber);
    console.log(`  Tracked FORMULA-S row ${rowNumber} for worksheet ${worksheetName}`);
}

/**
 * Gets stored FORMULA-S row positions for a worksheet
 * @param {string} worksheetName - Name of the worksheet
 * @returns {number[]} Array of row numbers
 */
function getFormulaSRows(worksheetName) {
    const rows = formulaSRowTracker.get(worksheetName);
    return rows ? Array.from(rows) : [];
}

/**
 * Clears stored FORMULA-S row positions for a worksheet (cleanup)
 * @param {string} worksheetName - Name of the worksheet
 */
function clearFormulaSRows(worksheetName) {
    formulaSRowTracker.delete(worksheetName);
}