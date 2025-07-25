/**
 * CodeCollection.js
 * Functions for processing and managing code collections
 */

import { convertKeysToCamelCase } from "@pinecone-database/pinecone/dist/utils";

// >>> ADDED: Import the logic validation function
import { validateLogicOnly } from './Validation.js';
// >>> ADDED: Import the format validation function
import { validateFormatErrors } from './ValidationFormat.js';

// === TIMING UTILITIES ===
const functionTimers = new Map();
const functionTimes = new Map();
let globalStartTime = null;

function startTimer(functionName) {
    const now = performance.now();
    functionTimers.set(functionName, now);
    console.log(`⏱️  [TIMER] Starting: ${functionName}`);
}

function endTimer(functionName) {
    const startTime = functionTimers.get(functionName);
    if (startTime) {
        const duration = (performance.now() - startTime) / 1000; // Convert to seconds
        functionTimes.set(functionName, (functionTimes.get(functionName) || 0) + duration);
        functionTimers.delete(functionName);
        console.log(`⏱️  [TIMER] Completed: ${functionName} (${duration.toFixed(3)}s)`);
    }
}

function startGlobalTimer() {
    globalStartTime = performance.now();
    console.log("🚀 [TIMER] Starting global runCodes timing");
}

function printTimingSummary() {
    const globalDuration = globalStartTime ? (performance.now() - globalStartTime) / 1000 : 0;
    
    console.log("\n" + "=".repeat(80));
    console.log("⏱️  FUNCTION TIMING SUMMARY - runCodes Process");
    console.log("=".repeat(80));
    
    const sortedTimes = Array.from(functionTimes.entries()).sort((a, b) => b[1] - a[1]);
    
    for (const [functionName, time] of sortedTimes) {
        const percentage = globalDuration > 0 ? ((time / globalDuration) * 100).toFixed(1) : "0.0";
        console.log(`${functionName.padEnd(45)} ${time.toFixed(3)}s (${percentage}%)`);
    }
    
    const measuredTotal = Array.from(functionTimes.values()).reduce((sum, time) => sum + time, 0);
    const unmeasuredTime = globalDuration - measuredTotal;
    
    console.log("-".repeat(80));
    console.log(`${"MEASURED FUNCTIONS TOTAL".padEnd(45)} ${measuredTotal.toFixed(3)}s`);
    console.log(`${"UNMEASURED TIME (overhead)".padEnd(45)} ${unmeasuredTime.toFixed(3)}s`);
    console.log(`${"GLOBAL TOTAL TIME".padEnd(45)} ${globalDuration.toFixed(3)}s`);
    console.log("=".repeat(80));
}

function resetTimers() {
    functionTimers.clear();
    functionTimes.clear();
    globalStartTime = null;
    console.log("🔄 [TIMER] Reset all timers");
}

function getTimingData() {
    return {
        functionTimes: Array.from(functionTimes.entries()),
        globalDuration: globalStartTime ? (performance.now() - globalStartTime) / 1000 : 0
    };
}

// Export timing functions for external use
export { startTimer, endTimer, startGlobalTimer, printTimingSummary, resetTimers, getTimingData };
// === END TIMING UTILITIES ===

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
        startTimer("populateCodeCollection");
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
        endTimer("populateCodeCollection");
        return codeCollection;
    } catch (error) {
        endTimer("populateCodeCollection");
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
        startGlobalTimer();
        startTimer("runCodes-initialization");
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
        
        endTimer("runCodes-initialization");
        startTimer("runCodes-main-loop");
        
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

                // Handle ACTUALS code type - process immediately
                if (codeType === "ACTUALS") {
                    console.log("ACTUALS code type encountered - processing immediately");
                    startTimer(`ACTUALS-processing-${codeType}-${i}`);
                    await processActualsSingle(code);
                    endTimer(`ACTUALS-processing-${codeType}-${i}`);
                    result.processedCodes++;
                    continue;
                }
                
                // Handle TAB code type
                if (codeType === "TAB") {
                    startTimer(`TAB-processing-${codeType}-${i}`);
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
                                // Clean up any references to this tab in the Financials sheet BEFORE deleting
                                await cleanupFinancialsReferences(tabName, context);
                                
                                // Delete the worksheet if it exists
                                existingSheet.delete();
                                await context.sync();
                            }
                            console.log("existingSheet deleted");
                            
                            // Get the Financials worksheet (needed for position and as fallback template)
                            const financialsSheet = context.workbook.worksheets.getItem("Financials");
                            financialsSheet.load("position"); // Load Financials sheet position
                            await context.sync(); // Sync to get Financials position
                            console.log(`Financials sheet is at position ${financialsSheet.position}`);
                            
                            // Always create new sheet since we deleted any existing one
                            {
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
                    
                    endTimer(`TAB-processing-${codeType}-${i}`);
                    continue;
                }
                
                // Handle non-TAB codes
                if (codeType !== "TAB") {
                    startTimer(`non-TAB-processing-${codeType}-${i}`);
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

                                // NEW: Row Grouping for INDEXBEGIN codes (simplified)
                                if (codeType === "INDEXBEGIN") {
                                    try {
                                        const numCopiedRows = lastRow - firstRow + 1;
                                        console.log(`🔍 [INDEXBEGIN GROUPING] Processing INDEXBEGIN code with ${numCopiedRows} copied rows`);
                                        console.log(`🔍 [INDEXBEGIN GROUPING] Copied from Codes rows ${firstRow}-${lastRow} to worksheet rows ${pasteRow}-${pasteRow + numCopiedRows - 1}`);
                                        
                                        if (numCopiedRows >= 3) {
                                            // Group the bottom 3 rows of the copied block
                                            const groupStartRow = pasteRow + (numCopiedRows - 3);
                                            const groupEndRow = pasteRow + (numCopiedRows - 1);
                                            
                                            console.log(`🎯 [INDEXBEGIN GROUPING] Grouping bottom 3 rows:`);
                                            console.log(`    Total copied rows: ${numCopiedRows}`);
                                            console.log(`    Group start row: ${groupStartRow}`);
                                            console.log(`    Group end row: ${groupEndRow}`);
                                            console.log(`    Rows to group: ${groupEndRow - groupStartRow + 1}`);
                                            
                                            console.log(`🔧 [INDEXBEGIN GROUPING] Creating range ${groupStartRow}:${groupEndRow}...`);
                                            const indexGroupRange = currentWS.getRange(`${groupStartRow}:${groupEndRow}`);
                                            
                                            console.log(`🔧 [INDEXBEGIN GROUPING] Applying group() method...`);
                                            indexGroupRange.group(Excel.GroupOption.byRows);
                                            
                                            console.log(`🔧 [INDEXBEGIN GROUPING] Applying hideGroupDetails() method...`);
                                            indexGroupRange.hideGroupDetails(Excel.GroupOption.byRows);
                                            
                                            console.log(`🔧 [INDEXBEGIN GROUPING] Syncing grouping changes...`);
                                            await context.sync();
                                            
                                            console.log(`🎉 [INDEXBEGIN GROUPING] Successfully grouped and collapsed rows ${groupStartRow}-${groupEndRow}`);
                                        } else {
                                            console.log(`⏭️ [INDEXBEGIN GROUPING] Skipping grouping - only ${numCopiedRows} rows copied (need at least 3)`);
                                        }
                                    } catch (indexGroupError) {
                                        console.error(`❌ [INDEXBEGIN GROUPING] Error processing INDEXBEGIN rows:`);
                                        console.error(`    Error message: ${indexGroupError.message}`);
                                        console.error(`    Error code: ${indexGroupError.code || 'N/A'}`);
                                        console.error(`    Full error:`, indexGroupError);
                                        // Continue processing even if grouping fails
                                    }
                                }

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

                                    console.log(`Applying negative transformation to formulas in U${endPastedRow}:CN${endPastedRow} for code ${codeType}`);
                                    const formulaRange = currentWS.getRange(`U${endPastedRow}:CN${endPastedRow}`);
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
                                        console.log(`Negative transformation applied and synced for U${endPastedRow}:CN${endPastedRow}`);
                                    } else {
                                        console.log(`No formulas found to transform in U${endPastedRow}:CN${endPastedRow}`);
                                    }
                                }
                                
                                // NEW: Apply "format" parameter for number formatting and italics
                                if (code.params.format) {
                                    const formatValue = String(code.params.format).toLowerCase();
                                    const numPastedRows = lastRow - firstRow + 1;
                                    const endPastedRow = pasteRow + Math.max(0, numPastedRows - 1);
                                    // Apply to J through CN (including column J now)
                                    const formatRangeAddress = `J${pasteRow}:CN${endPastedRow}`;
                                    const rangeToFormat = currentWS.getRange(formatRangeAddress);
                                    let numberFormatString = null;
                                    // Removed applyItalics variable as direct checks on formatValue are clearer for B:CN range

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
                                        console.log(`Applying number format: "${numberFormatString}" to ${formatRangeAddress}`); // J:CN
                                        rangeToFormat.numberFormat = [[numberFormatString]]; // J:CN
                                        
                                        // Italicization logic based on formatValue for the B:CN range
                                        const fullItalicRangeAddress = `B${pasteRow}:CN${endPastedRow}`;
                                        const fullRangeToHandleItalics = currentWS.getRange(fullItalicRangeAddress);

                                        if (formatValue === "dollaritalic" || formatValue === "volume" || formatValue === "percent" || formatValue === "factor") {
                                            console.log(`Applying italics to ${fullItalicRangeAddress} due to format type (${formatValue})`);
                                            fullRangeToHandleItalics.format.font.italic = true;
                                        } else if (formatValue === "dollar" || formatValue === "date") {
                                            console.log(`Ensuring ${fullItalicRangeAddress} is NOT italicized due to format type (${formatValue})`);
                                            fullRangeToHandleItalics.format.font.italic = false;
                                        } else {
                                            // For unrecognized formats that still had a numberFormatString (e.g. if logic changes later),
                                            // or if K:CN needs explicit non-italic default when no B:CN rule applies.
                                            // However, current logic implies if numberFormatString is set, formatValue is one of the known ones.
                                            // If K:CN (rangeToFormat) needs specific non-italic handling for other cases, it would go here.
                                            // For now, this 'else' might not be hit if numberFormatString implies a known formatValue.
                                            // The primary `italic` parameter handles general italic override later anyway.
                                            console.log(`Format type ${formatValue} has number format but no specific B:CN italic rule. K:CN italics remain as previously set or default.`);
                                        }
                                        await context.sync();
                                        console.log(`"format" parameter processing (number format and B:CN italics) synced for ${formatRangeAddress}`);
                                    } else {
                                        console.log(`"format" parameter value "${formatValue}" is not recognized. No formatting applied.`);
                                    }
                                }

                                // NEW: Apply "italic" parameter for font style
                                if (code.params.italic !== undefined) { // Check if the parameter exists
                                    const italicValue = String(code.params.italic).toLowerCase();
                                    const numPastedRows = lastRow - firstRow + 1;
                                    const endPastedRow = pasteRow + Math.max(0, numPastedRows - 1);
                                    const italicRangeAddress = `B${pasteRow}:CN${endPastedRow}`;
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

                                                // Check for monthsr parameter and auto-add appropriate sumif type if missing
                const hasMonthsrParam = Object.keys(code.params).some(key => key.startsWith('monthsr'));
                if (hasMonthsrParam && !code.params.sumif) {
                    // Determine appropriate sumif type based on code type
                    let autoSumifType = "year"; // Default for SPREAD-E
                    if (codeType === "CONST-E" || codeType === "ENDPOINT-E") {
                        autoSumifType = "average";
                    }
                    
                    console.log(`🗓️ [AUTO-SUMIF] Detected monthsr parameter without sumif - automatically adding sumif="${autoSumifType}" for code type ${codeType}`);
                    code.params.sumif = autoSumifType;
                    console.log(`🗓️ [AUTO-SUMIF] Code parameters after auto-add:`, code.params);
                } else if (hasMonthsrParam && code.params.sumif) {
                    console.log(`🗓️ [AUTO-SUMIF] monthsr parameter detected, sumif already present: "${code.params.sumif}"`);
                }

                // Apply the driver and assumption inputs function to the current worksheet
                try {
                    startTimer(`driverAndAssumptionInputs-${codeType}-${i}`);
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
                    endTimer(`driverAndAssumptionInputs-${codeType}-${i}`);
                } catch (error) {
                    endTimer(`driverAndAssumptionInputs-${codeType}-${i}`);
                    console.error(`Error applying driver and assumption inputs: ${error.message}`);
                    result.errors.push({
                        codeIndex: i,
                        codeType: codeType,
                        error: `Error applying driver and assumption inputs: ${error.message}`
                    });
                }

                                // NEW: Apply "sumif" parameter for custom SUMIF/AVERAGEIF formulas only to Y1-Y6 columns containing "F"
                                // NOTE: This must be AFTER driverAndAssumptionInputs to avoid being overwritten
                                console.log(`🔍 [SUMIF CHECK] Checking for sumif parameter. code.params.sumif = "${code.params.sumif}"`);
                                console.log(`🔍 [SUMIF CHECK] All code parameters:`, code.params);
                                if (code.params.sumif !== undefined) {
                                    const sumifValue = String(code.params.sumif).toLowerCase();
                                    const numPastedRows = lastRow - firstRow + 1;
                                    const endPastedRow = pasteRow + Math.max(0, numPastedRows - 1);

                                    console.log(`✅ [SUMIF PROCESSING] Processing "sumif" parameter: "${sumifValue}" - analyzing row1 for Y1-Y6 columns with "F"`);

                                    // Parse row1 parameter to find which Y1-Y6 positions contain "F"
                                    const row1Param = code.params.row1;
                                    const columnsWithF = [];
                                    
                                    if (row1Param) {
                                        const row1Parts = row1Param.split('|');
                                        console.log(`  Analyzing row1 with ${row1Parts.length} parts`);
                                        
                                        // Y1-Y6 are at indices 9-14 in the row parameter (0-based)
                                        // They map to columns K-P respectively (J is not mapped to any Y position)
                                        const y1ToY6Indices = [9, 10, 11, 12, 13, 14]; // Y1, Y2, Y3, Y4, Y5, Y6
                                        const columnLetters = ['K', 'L', 'M', 'N', 'O', 'P']; // Corresponding columns
                                        
                                        for (let i = 0; i < y1ToY6Indices.length; i++) {
                                            const partIndex = y1ToY6Indices[i];
                                            const columnLetter = columnLetters[i];
                                            const yLabel = `Y${i + 1}`;
                                            
                                            if (partIndex < row1Parts.length) {
                                                const part = row1Parts[partIndex];
                                                console.log(`    ${yLabel} (index ${partIndex}): "${part}" -> Column ${columnLetter}`);
                                                
                                                // Check if this part contains "F" (case-insensitive)
                                                if (part && part.toUpperCase().includes('F')) {
                                                    columnsWithF.push({
                                                        yLabel: yLabel,
                                                        columnLetter: columnLetter,
                                                        columnIndex: i, // 0-based index for J-P
                                                        part: part
                                                    });
                                                    console.log(`      ✅ Found "F" in ${yLabel} -> will apply SUMIF to column ${columnLetter}`);
                                                } else {
                                                    console.log(`      ❌ No "F" found in ${yLabel} -> will skip column ${columnLetter}`);
                                                }
                                            }
                                        }
                                    }

                                    if (columnsWithF.length === 0) {
                                        console.log(`  No Y1-Y6 columns contain "F" - skipping sumif parameter application`);
                                    } else {
                                        console.log(`  Found ${columnsWithF.length} Y1-Y6 columns with "F": ${columnsWithF.map(c => c.yLabel + '(' + c.columnLetter + ')').join(', ')}`);

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
                                            console.log(`  Applying ${sumifValue} formula template: ${formulaTemplate}`);
                                            
                                            if (sumifValue === "offsetyear") {
                                                // Special handling for offsetyear: set column J to 0, set non-F columns in K-P to 0, apply formula to F columns
                                                console.log(`  Special offsetyear handling: setting J and non-F columns to 0, applying formula to F columns only`);
                                                
                                                // Set column J to 0 (as in original logic)
                                                const columnJRange = currentWS.getRange(`J${pasteRow}:J${endPastedRow}`);
                                                const zeroArrayJ = [];
                                                for (let r = 0; r < numPastedRows; r++) {
                                                    zeroArrayJ.push([0]);
                                                }
                                                columnJRange.values = zeroArrayJ;
                                                console.log(`    Set column J to 0`);
                                                
                                                // Set all K-P columns to 0 first
                                                const allKPRange = currentWS.getRange(`K${pasteRow}:P${endPastedRow}`); // K through P (Y1-Y6 columns)
                                                const zeroArrayKP = [];
                                                for (let r = 0; r < numPastedRows; r++) {
                                                    zeroArrayKP.push([0, 0, 0, 0, 0, 0]); // 6 columns K-P
                                                }
                                                allKPRange.values = zeroArrayKP;
                                                console.log(`    Set columns K-P to 0`);
                                                
                                                // Then apply formula only to columns with F
                                                for (const colInfo of columnsWithF) {
                                                    const colRange = currentWS.getRange(`${colInfo.columnLetter}${pasteRow}:${colInfo.columnLetter}${endPastedRow}`);
                                                    const colFormulas = [];
                                                    for (let r = 0; r < numPastedRows; r++) {
                                                        colFormulas.push([formulaTemplate]);
                                                    }
                                                    colRange.formulas = colFormulas;
                                                    console.log(`    Applied offsetyear formula to column ${colInfo.columnLetter} (${colInfo.yLabel})`);
                                                }
                                                
                                                // Apply black font color to all columns containing "F" (offsetyear case)
                                                console.log(`    Applying black font color to ${columnsWithF.length} columns with "F" (offsetyear)`);
                                                for (const colInfo of columnsWithF) {
                                                    const colRange = currentWS.getRange(`${colInfo.columnLetter}${pasteRow}:${colInfo.columnLetter}${endPastedRow}`);
                                                    colRange.format.font.color = "#000000"; // Black font color
                                                    console.log(`      Applied black font color to column ${colInfo.columnLetter} (${colInfo.yLabel})`);
                                                }
                                            } else {
                                                // Standard handling: apply formula only to columns with F
                                                console.log(`  Standard handling: applying formula only to columns with F`);
                                                
                                                for (const colInfo of columnsWithF) {
                                                    const colRange = currentWS.getRange(`${colInfo.columnLetter}${pasteRow}:${colInfo.columnLetter}${endPastedRow}`);
                                                    const colFormulas = [];
                                                    for (let r = 0; r < numPastedRows; r++) {
                                                        colFormulas.push([formulaTemplate]);
                                                    }
                                                    colRange.formulas = colFormulas;
                                                    console.log(`    Applied ${sumifValue} formula to column ${colInfo.columnLetter} (${colInfo.yLabel}): "${colInfo.part}"`);
                                                }
                                            }
                                            
                                            // Apply black font color to all columns containing "F"
                                            console.log(`  Applying black font color to ${columnsWithF.length} columns with "F"`);
                                            for (const colInfo of columnsWithF) {
                                                const colRange = currentWS.getRange(`${colInfo.columnLetter}${pasteRow}:${colInfo.columnLetter}${endPastedRow}`);
                                                colRange.format.font.color = "#000000"; // Black font color
                                                console.log(`    Applied black font color to column ${colInfo.columnLetter} (${colInfo.yLabel})`);
                                            }
                                            
                                            await context.sync();
                                            console.log(`  "sumif" parameter (${sumifValue}) processing synced for ${columnsWithF.length} columns with F`);
                                        } else {
                                            console.log(`  "sumif" parameter value "${sumifValue}" is not recognized. Valid values are: year, yearend, average, offsetyear. No formula changes applied.`);
                                        }
                                    }
                                } else {
                                    console.log(`❌ [SUMIF CHECK] No sumif parameter found - sumif processing skipped`);
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
                    endTimer(`non-TAB-processing-${codeType}-${i}`);
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
        endTimer("runCodes-main-loop");
        startTimer("runCodes-finalization");
        
        const finalResult = {
            ...result, // Includes processedCodes, errors
            assumptionTabs: assumptionTabs.map(tab => tab.name) // Return only the names
        };

        endTimer("runCodes-finalization");
        console.log("runCodes finished. Returning:", finalResult);
        
        // Restore left borders to column S cells with interior color on all assumption tabs
        console.log("Restoring column S left borders on all assumption tabs...");
        for (const assumptionTab of assumptionTabs) {
            try {
                await restoreColumnSLeftBorders(assumptionTab);
            } catch (borderError) {
                console.error(`Error restoring column S borders on ${assumptionTab.name}: ${borderError.message}`);
            }
        }
        console.log("Completed column S border restoration.");

        // ACTUALS codes are now processed immediately when encountered (no batch processing needed)
        
        // Print timing summary for runCodes only
        printTimingSummary();
        
        return finalResult; // Return the modified result object

    } catch (error) {
        console.error("Error in runCodes:", error);
        // Print timing summary even on error
        printTimingSummary();
        // Consider how to return errors. Throwing stops execution.
        // Returning them in the result allows the caller to decide.
        throw error; // Or return { errors: [error.message], assumptionTabs: [] }
    }
}

/**
 * Helper function to delete rows in Financials sheet that reference a deleted tab
 * @param {string} deletedTabName - The name of the tab that was deleted
 * @param {Excel.RequestContext} context - The Excel request context
 */
async function cleanupFinancialsReferences(deletedTabName, context) {
    try {
        console.log(`[CLEANUP] Starting cleanup of Financials references to deleted tab: ${deletedTabName}`);
        
        // Get the Financials worksheet
        const financialsSheet = context.workbook.worksheets.getItem("Financials");
        
        // Get used range to determine how many rows to check
        const usedRange = financialsSheet.getUsedRange(true);
        if (!usedRange) {
            console.log("[CLEANUP] No used range found in Financials sheet");
            return;
        }
        
        usedRange.load("rowCount");
        await context.sync();
        
        const lastRow = usedRange.rowCount;
        console.log(`[CLEANUP] Checking ${lastRow} rows in Financials sheet`);
        
        // Get column U formulas for all rows
        const columnURange = financialsSheet.getRange(`U1:U${lastRow}`);
        columnURange.load("formulas");
        await context.sync();
        
        const formulas = columnURange.formulas;
        const rowsToDelete = [];
        
        // Check each formula in column U
        for (let i = 0; i < formulas.length; i++) {
            const formula = formulas[i][0]; // Get the formula from the 2D array
            if (typeof formula === 'string' && formula.startsWith('=')) {
                // Check if formula references the deleted tab
                const tabReference = `${deletedTabName}!`;
                if (formula.includes(tabReference)) {
                    const rowNumber = i + 1; // Convert 0-based index to 1-based row number
                    rowsToDelete.push(rowNumber);
                    console.log(`[CLEANUP] Found reference to ${deletedTabName} in row ${rowNumber}: ${formula}`);
                }
            }
        }
        
        // Delete rows in reverse order (highest row number first) to avoid shifting issues
        rowsToDelete.reverse();
        
        for (const rowNumber of rowsToDelete) {
            console.log(`[CLEANUP] Deleting row ${rowNumber} from Financials sheet`);
            const rowRange = financialsSheet.getRangeByIndexes(rowNumber - 1, 0, 1, 1000); // Get entire row
            rowRange.delete(Excel.DeleteShiftDirection.up);
        }
        
        if (rowsToDelete.length > 0) {
            await context.sync();
            console.log(`[CLEANUP] Successfully deleted ${rowsToDelete.length} rows from Financials sheet that referenced ${deletedTabName}`);
        } else {
            console.log(`[CLEANUP] No rows found in Financials sheet that reference ${deletedTabName}`);
        }
        
    } catch (error) {
        console.error(`[CLEANUP] Error cleaning up Financials references for ${deletedTabName}:`, error);
        // Don't throw the error to avoid breaking the main flow
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
 * Handles special functions: SPREAD, BEG, END, RAISE, ONETIMEDATE, SPREADDATES, ONETIMEINDEX
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
    console.log(`🔧 [FORMULA-S PARSER] Starting with: "${result}"`);
    
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
                const replacement = `U$${driverRow}`;
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
    
    console.log(`🔧 [FORMULA-S PARSER] After rd{}/cd{} processing: "${result}"`);
    
    // Process SPREAD function: SPREAD(driver) -> driver/U$7
    result = result.replace(/SPREAD\(([^)]+)\)/gi, (match, driver) => {
        console.log(`    Converting SPREAD(${driver}) to ${driver}/U$7`);
        return `(${driver}/U$7)`;
    });
    
    // Process BEG function: BEG(driver) -> (EOMONTH(driver,0)<=EOMONTH(U$2,0))
    result = result.replace(/BEG\(([^)]+)\)/gi, (match, driver) => {
        console.log(`    Converting BEG(${driver}) to (EOMONTH(${driver},0)<=EOMONTH(U$2,0))`);
        return `(EOMONTH(${driver},0)<=EOMONTH(U$2,0))`;
    });
    
    // Process END function: END(driver) -> (EOMONTH(driver,0)>EOMONTH(U$2,0))
    result = result.replace(/END\(([^)]+)\)/gi, (match, driver) => {
        console.log(`    Converting END(${driver}) to (EOMONTH(${driver},0)>EOMONTH(U$2,0))`);
        return `(EOMONTH(${driver},0)>EOMONTH(U$2,0))`;
    });
    
    // Process RAISE function: RAISE(driver1,driver2) -> (1 + (driver1)) ^ (U$3 - max(year(driver2), $U3))
    result = result.replace(/RAISE\(([^,]+),([^)]+)\)/gi, (match, driver1, driver2) => {
        // Trim whitespace from drivers
        driver1 = driver1.trim();
        driver2 = driver2.trim();
        console.log(`    Converting RAISE(${driver1},${driver2}) to (1 + (${driver1})) ^ (U$3 - max(year(${driver2}), $U3))`);
        return `(1 + (${driver1})) ^ (U$3 - max(year(${driver2}), $U3))`;
    });
    
    // Process ANNBONUS function: ANNBONUS(driver1,driver2) -> (driver1*(MONTH(EOMONTH(U$2,0))=driver2))
    result = result.replace(/ANNBONUS\(([^,]+),([^)]+)\)/gi, (match, driver1, driver2) => {
        // Trim whitespace from drivers
        driver1 = driver1.trim();
        driver2 = driver2.trim();
        console.log(`    Converting ANNBONUS(${driver1},${driver2}) to (${driver1}*(MONTH(EOMONTH(U$2,0))=${driver2}))`);
        return `(${driver1}*(MONTH(EOMONTH(U$2,0))=${driver2}))`;
    });
    
    // Process QUARTERBONUS function: QUARTERBONUS(driver1) -> (driver1*(U$6<>0))
    result = result.replace(/QUARTERBONUS\(([^)]+)\)/gi, (match, driver1) => {
        // Trim whitespace from driver
        driver1 = driver1.trim();
        console.log(`    Converting QUARTERBONUS(${driver1}) to (${driver1}*(U$6<>0))`);
        return `(${driver1}*(U$6<>0))`;
    });
    
    // Process ONETIMEDATE function: ONETIMEDATE(driver) -> (EOMONTH((driver),0)=EOMONTH(U$2,0))
    result = result.replace(/ONETIMEDATE\(([^)]+)\)/gi, (match, driver) => {
        console.log(`    Converting ONETIMEDATE(${driver}) to (EOMONTH((${driver}),0)=EOMONTH(U$2,0))`);
        return `(EOMONTH((${driver}),0)=EOMONTH(U$2,0))`;
    });
    
    // Process SPREADDATES function: SPREADDATES(driver1,driver2,driver3) -> IF(AND(EOMONTH(EOMONTH(U$2,0),0)>=EOMONTH(driver2,0),EOMONTH(EOMONTH(U$2,0),0)<=EOMONTH(driver3,0)),driver1/(DATEDIF(driver2,driver3,"m")+1),0)
    // Note: Need to handle nested parentheses and comma separation
    result = result.replace(/SPREADDATES\(([^,]+),([^,]+),([^)]+)\)/gi, (match, driver1, driver2, driver3) => {
        // Trim whitespace from drivers
        driver1 = driver1.trim();
        driver2 = driver2.trim();
        driver3 = driver3.trim();
        const newFormula = `IF(AND(EOMONTH(EOMONTH(U$2,0),0)>=EOMONTH(${driver2},0),EOMONTH(EOMONTH(U$2,0),0)<=EOMONTH(${driver3},0)),${driver1}/(DATEDIF(${driver2},${driver3},"m")+1),0)`;
        console.log(`    Converting SPREADDATES(${driver1},${driver2},${driver3}) to ${newFormula}`);
        return newFormula;
    });
    
    // Process AMORT function: AMORT(driver1,driver2,driver3,driver4) -> IFERROR(PPMT(driver1/U$7,1,driver2-U$8+1,SUM(driver3,OFFSET(driver4,-1,0),0,0),0),0)
    result = result.replace(/AMORT\(([^,]+),([^,]+),([^,]+),([^)]+)\)/gi, (match, driver1, driver2, driver3, driver4) => {
        // Trim whitespace from drivers
        driver1 = driver1.trim();
        driver2 = driver2.trim();
        driver3 = driver3.trim();
        driver4 = driver4.trim();
        const newFormula = `IFERROR(PPMT(${driver1}/U$7,1,${driver2}-U$8+1,SUM(${driver3},OFFSET(${driver4},-1,0),0,0),0),0)`;
        console.log(`    Converting AMORT(${driver1},${driver2},${driver3},${driver4}) to ${newFormula}`);
        return newFormula;
    });
    
    // Process BULLET function: BULLET(driver1,driver2,driver3) -> IF(driver2=U$8,SUM(driver1,OFFSET(driver3,-1,0)),0)
    result = result.replace(/BULLET\(([^,]+),([^,]+),([^)]+)\)/gi, (match, driver1, driver2, driver3) => {
        // Trim whitespace from drivers
        driver1 = driver1.trim();
        driver2 = driver2.trim();
        driver3 = driver3.trim();
        const newFormula = `IF(${driver2}=U$8,SUM(${driver1},OFFSET(${driver3},-1,0)),0)`;
        console.log(`    Converting BULLET(${driver1},${driver2},${driver3}) to ${newFormula}`);
        return newFormula;
    });
    
    // Process PBULLET function: PBULLET(driver1,driver2) -> (driver1*(driver2=U$8)) (Fixed partial principal payment)
    result = result.replace(/PBULLET\(([^,]+),([^)]+)\)/gi, (match, driver1, driver2) => {
        // Trim whitespace from drivers
        driver1 = driver1.trim();
        driver2 = driver2.trim();
        const newFormula = `(${driver1}*(${driver2}=U$8))`;
        console.log(`    Converting PBULLET(${driver1},${driver2}) to ${newFormula}`);
        return newFormula;
    });
    
    // Process INTONLY function: INTONLY(driver1) -> (U$8>driver1) (Interest only period)
    result = result.replace(/INTONLY\(([^)]+)\)/gi, (match, driver1) => {
        // Trim whitespace from driver
        driver1 = driver1.trim();
        const newFormula = `(U$8>${driver1})`;
        console.log(`    Converting INTONLY(${driver1}) to ${newFormula}`);
        return newFormula;
    });
    
    // Process ONETIMEINDEX function: ONETIMEINDEX(driver1) -> (U$8=driver1) (One-time event at specific index month)
    result = result.replace(/ONETIMEINDEX\(([^)]+)\)/gi, (match, driver1) => {
        // Trim whitespace from driver
        driver1 = driver1.trim();
        const newFormula = `(U$8=${driver1})`;
        console.log(`    Converting ONETIMEINDEX(${driver1}) to ${newFormula}`);
        return newFormula;
    });
    
    // Process BEGINDEX function: BEGINDEX(driver1) -> (U$8>=driver1) (Start from index month and continue)
    result = result.replace(/BEGINDEX\(([^)]+)\)/gi, (match, driver1) => {
        // Trim whitespace from driver
        driver1 = driver1.trim();
        const newFormula = `(U$8>=${driver1})`;
        console.log(`    Converting BEGINDEX(${driver1}) to ${newFormula}`);
        return newFormula;
    });
    
    // Process ENDINDEX function: ENDINDEX(driver1) -> (U$8<=driver1) (End at index month)
    result = result.replace(/ENDINDEX\(([^)]+)\)/gi, (match, driver1) => {
        // Trim whitespace from driver
        driver1 = driver1.trim();
        const newFormula = `(U$8<=${driver1})`;
        console.log(`    Converting ENDINDEX(${driver1}) to ${newFormula}`);
        return newFormula;
    });
    
    // Process RANGE function: RANGE(driver1) -> U{row}:CN{row} where {row} is the row number from driver1
    console.log(`    Checking for RANGE function in: "${result}"`);
    const rangeMatches = result.match(/RANGE\(([^)]+)\)/gi);
    if (rangeMatches) {
        console.log(`    Found RANGE matches: ${rangeMatches.join(', ')}`);
    } else {
        console.log(`    No RANGE matches found`);
    }
    
    result = result.replace(/RANGE\(([^)]+)\)/gi, (match, driver1) => {
        // Trim whitespace from driver
        driver1 = driver1.trim();
        
        // Extract row number from driver1 (e.g., U$15 -> 15, U15 -> 15)
        const rowMatch = driver1.match(/[A-Z]+\$?(\d+)/);
        if (rowMatch) {
            const rowNumber = rowMatch[1];
            const newFormula = `U${rowNumber}:CN${rowNumber}`;
            console.log(`    Converting RANGE(${driver1}) to ${newFormula}`);
            return newFormula;
        } else {
            console.warn(`    Could not extract row number from ${driver1}, keeping as is`);
            return match; // Keep original if we can't parse
        }
    });

    // Process TABLEMIN function: TABLEMIN(driver1,driver2) -> MIN(SUM(driver1:OFFSET(INDIRECT(ADDRESS(ROW(),COLUMN())),-1,0)),driver2)
    console.log(`    Checking for TABLEMIN function in: "${result}"`);
    const tableminMatches = result.match(/TABLEMIN\(([^,]+),([^)]+)\)/gi);
    if (tableminMatches) {
        console.log(`    Found TABLEMIN matches: ${tableminMatches.join(', ')}`);
    } else {
        console.log(`    No TABLEMIN matches found`);
    }
    
    result = result.replace(/TABLEMIN\(([^,]+),([^)]+)\)/gi, (match, driver1, driver2) => {
        // Trim whitespace from drivers
        driver1 = driver1.trim();
        driver2 = driver2.trim();
        const newFormula = `MIN(SUM(${driver1}:OFFSET(INDIRECT(ADDRESS(ROW(),COLUMN())),-1,0)),${driver2})`;
        console.log(`    Converting TABLEMIN(${driver1},${driver2}) to ${newFormula}`);
        return newFormula;
    });

    
    // Replace timeseriesdivisor with U$7
    result = result.replace(/timeseriesdivisor/gi, (match) => {
        const replacement = 'U$7';
        console.log(`    Replacing ${match} with ${replacement}`);
        return replacement;
    });
    
    // Replace currentmonth with EOMONTH(U$2,0)
    result = result.replace(/currentmonth/gi, (match) => {
        const replacement = 'EOMONTH(U$2,0)';
        console.log(`    Replacing ${match} with ${replacement}`);
        return replacement;
    });
    
    // Replace beginningmonth with EOMONTH($U$2,0)
    result = result.replace(/beginningmonth/gi, (match) => {
        const replacement = 'EOMONTH($U$2,0)';
        console.log(`    Replacing ${match} with ${replacement}`);
        return replacement;
    });
    
    // Replace currentyear with U$3
    result = result.replace(/currentyear/gi, (match) => {
        const replacement = 'U$3';
        console.log(`    Replacing ${match} with ${replacement}`);
        return replacement;
    });
    
    // Replace yearend with U$4
    result = result.replace(/yearend/gi, (match) => {
        const replacement = 'U$4';
        console.log(`    Replacing ${match} with ${replacement}`);
        return replacement;
    });
    
    // Replace beginningyear with $U$3
    result = result.replace(/beginningyear/gi, (match) => {
        const replacement = '$U$3';
        console.log(`    Replacing ${match} with ${replacement}`);
        return replacement;
    });
    
    console.log(`🔧 [FORMULA-S PARSER] Final result: "${result}"`);
    return result;
}

/**
 * Parses a value string and extracts comments from square brackets
 * @param {string} valueString - The value string that may contain comments in square brackets
 * @returns {Object} - Object containing cleaned value and extracted comment
 */
function parseCommentFromBrackets(valueString) {
    if (!valueString || typeof valueString !== 'string') {
        return { cleanValue: valueString, comment: null };
    }

    // Look for content in square brackets [comment]
    const bracketMatch = valueString.match(/\[([^\]]+)\]/);
    if (bracketMatch) {
        const comment = bracketMatch[1].trim();
        const cleanValue = valueString.replace(/\[[^\]]+\]/, ''); // Remove the bracket content
        console.log(`    Extracted comment: "${comment}" from value: "${valueString}"`);
        console.log(`    Clean value after comment removal: "${cleanValue}"`);
        return { cleanValue, comment };
    }

    return { cleanValue: valueString, comment: null };
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
        
        // Check if it ends with % for percentage formatting
        if (cleanValue.endsWith('%')) {
            formatType = 'percent';
            cleanValue = cleanValue.substring(0, cleanValue.length - 1); // Remove % from clean value
        }
        // Check if it ends with "x" for factor formatting (e.g., "~10x")
        else if (cleanValue.toLowerCase().endsWith('x') && cleanValue.length > 1) {
            const numberPart = cleanValue.substring(0, cleanValue.length - 1);
            if (!isNaN(Number(numberPart)) && numberPart.trim() !== '') {
                formatType = 'factor';
                cleanValue = numberPart; // Remove the 'x' from the clean value
            }
        }
        // If after removing ~, it's a number or "F", apply volume formatting
        else if (cleanValue === 'F' || (!isNaN(Number(cleanValue)) && cleanValue.trim() !== '')) {
            formatType = 'volume';
        }
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
 * Generates a sequence of column letters starting from U for the specified length
 * @param {number} length - Number of columns needed
 * @returns {string[]} - Array of column letters starting from U
 */
function generateMonthsColumnSequence(length) {
    const sequence = [];
    const startColumn = 'U'; // Starting from column U
    const startIndex = columnLetterToIndex(startColumn); // Convert U to index
    
    for (let i = 0; i < length; i++) {
        const columnIndex = startIndex + i;
        const columnLetter = columnIndexToLetter(columnIndex);
        sequence.push(columnLetter);
    }
    
    console.log(`Generated months column sequence starting from U for ${length} values: ${sequence.join(', ')}`);
    return sequence;
}

/**
 * Auto-populates the entire year's worth of monthly data based on code type and year detection
 * @param {Excel.Worksheet} worksheet - The worksheet to populate
 * @param {number} currentRowNum - The current row number being processed
 * @param {string[]} monthValues - Array of initial month values from monthsr parameter
 * @param {string[]} initialColumnSequence - Array of initial column letters that were populated
 * @param {string} codeType - The code type (SPREAD-E, CONST-E, ENDPOINT-E)
 * @param {Excel.RequestContext} context - The Excel context
 * @returns {Promise<void>}
 */
async function autoPopulateEntireYear(worksheet, currentRowNum, monthValues, initialColumnSequence, codeType, context) {
    console.log(`🗓️ [YEAR AUTO-POPULATE] Starting year-wide population for row ${currentRowNum}, code type: ${codeType}`);
    
    try {
        // Step 1: Determine the year from row 3 of the last populated month column
        if (initialColumnSequence.length === 0) {
            console.log(`🗓️ [YEAR AUTO-POPULATE] No initial columns populated, skipping year auto-population`);
            return;
        }
        
        const lastPopulatedColumn = initialColumnSequence[initialColumnSequence.length - 1];
        console.log(`🗓️ [YEAR AUTO-POPULATE] Last populated column: ${lastPopulatedColumn}`);
        
        // Get the year from row 3 of the last populated column
        const yearCell = worksheet.getRange(`${lastPopulatedColumn}3`);
        yearCell.load("values");
        await context.sync();
        
        const yearValue = yearCell.values[0][0];
        console.log(`🗓️ [YEAR AUTO-POPULATE] Year detected from ${lastPopulatedColumn}3: ${yearValue}`);
        
        if (!yearValue || isNaN(Number(yearValue))) {
            console.warn(`🗓️ [YEAR AUTO-POPULATE] Invalid year value '${yearValue}' in ${lastPopulatedColumn}3, skipping year auto-population`);
            return;
        }
        
        // Step 2: Find all columns to the right that have the same year in row 3
        const yearColumns = await findColumnsWithYear(worksheet, yearValue, lastPopulatedColumn, context);
        console.log(`🗓️ [YEAR AUTO-POPULATE] Found ${yearColumns.length} columns with year ${yearValue}: ${yearColumns.join(', ')}`);
        
        if (yearColumns.length === 0) {
            console.log(`🗓️ [YEAR AUTO-POPULATE] No additional columns found with year ${yearValue}, skipping year auto-population`);
            return;
        }
        
        // Step 3: Determine the fill value based on code type
        // All code types (SPREAD-E, CONST-E, ENDPOINT-E) now use the last defined month value
        const lastDefinedValue = getLastDefinedMonthValue(monthValues);
        let fillValue = lastDefinedValue;
        
        console.log(`🗓️ [YEAR AUTO-POPULATE] Code type ${codeType}: using last defined value ${fillValue}`);
        
        // Step 4: Populate the year columns with the fill value
        console.log(`🗓️ [YEAR AUTO-POPULATE] Populating ${yearColumns.length} columns with value: ${fillValue}`);
        
        for (const column of yearColumns) {
            const cellToFill = worksheet.getRange(`${column}${currentRowNum}`);
            cellToFill.values = [[fillValue]];
            console.log(`    Set ${column}${currentRowNum} = ${fillValue}`);
        }
        
        // Step 5: Apply blue font formatting to all populated year columns
        const allYearColumns = [...yearColumns];
        for (const column of allYearColumns) {
            const cellToFormat = worksheet.getRange(`${column}${currentRowNum}`);
            cellToFormat.format.font.color = "#0000FF"; // Blue font color
        }
        
        await context.sync();
        console.log(`🗓️ [YEAR AUTO-POPULATE] Applied blue font formatting to ${allYearColumns.length} year columns`);
        
        // Step 6: Find and update the corresponding annual column (K:P) with SUMIF
        try {
            await updateAnnualColumnWithSumif(worksheet, currentRowNum, yearValue, codeType, context);
        } catch (annualError) {
            console.error(`🗓️ [YEAR AUTO-POPULATE] Error updating annual column: ${annualError.message}`);
        }
        
        console.log(`✅ [YEAR AUTO-POPULATE] Completed year-wide population for row ${currentRowNum}`);
        
    } catch (error) {
        console.error(`❌ [YEAR AUTO-POPULATE] Error in autoPopulateEntireYear: ${error.message}`, error);
        throw error;
    }
}

/**
 * Finds all columns to the right of a starting column that have a specific year in row 3
 * @param {Excel.Worksheet} worksheet - The worksheet to search
 * @param {any} yearValue - The year value to search for
 * @param {string} startColumn - The starting column letter
 * @param {Excel.RequestContext} context - The Excel context
 * @returns {Promise<string[]>} - Array of column letters that have the year
 */
async function findColumnsWithYear(worksheet, yearValue, startColumn, context) {
    const yearColumns = [];
    const startIndex = columnLetterToIndex(startColumn);
    const endIndex = columnLetterToIndex("CN"); // Search up to CN
    
    console.log(`🔍 [YEAR SEARCH] Searching for year ${yearValue} from column ${startColumn} to CN`);
    
    // Load row 3 values for the entire search range
    const searchRangeStart = columnIndexToLetter(startIndex + 1); // Start after the initial column
    const searchRange = worksheet.getRange(`${searchRangeStart}3:CN3`);
    searchRange.load("values");
    await context.sync();
    
    const row3Values = searchRange.values[0]; // Get the single row array
    
    // Check each column for the year value
    for (let i = 0; i < row3Values.length; i++) {
        const cellValue = row3Values[i];
        const columnIndex = startIndex + 1 + i; // Calculate actual column index
        const columnLetter = columnIndexToLetter(columnIndex);
        
        if (cellValue == yearValue) { // Use == for type-flexible comparison
            yearColumns.push(columnLetter);
            console.log(`    Found year ${yearValue} in column ${columnLetter}`);
        }
    }
    
    return yearColumns;
}

/**
 * Gets the last defined (non-empty, non-zero) month value from the monthValues array
 * @param {string[]} monthValues - Array of month values
 * @returns {number} - The last defined value or 0 if none found
 */
function getLastDefinedMonthValue(monthValues) {
    // Iterate backwards through the month values to find the last defined one
    for (let i = monthValues.length - 1; i >= 0; i--) {
        const value = monthValues[i];
        if (value && value.trim() !== '' && value.trim() !== '0') {
            // Parse value to extract clean numeric value (remove symbols)
            const parsed = parseValueWithSymbols(value);
            const numValue = Number(parsed.cleanValue);
            if (!isNaN(numValue)) {
                console.log(`📊 [LAST DEFINED] Last defined month value: ${numValue} from "${value}"`);
                return numValue;
            }
        }
    }
    
    console.log(`📊 [LAST DEFINED] No defined month values found, using 0`);
    return 0;
}

/**
 * Updates the corresponding annual column (K:P) with a SUMIF formula based on code type
 * @param {Excel.Worksheet} worksheet - The worksheet to update
 * @param {number} currentRowNum - The current row number
 * @param {any} yearValue - The year value to match
 * @param {string} codeType - The code type (SPREAD-E, CONST-E, ENDPOINT-E)
 * @param {Excel.RequestContext} context - The Excel context
 * @returns {Promise<void>}
 */
async function updateAnnualColumnWithSumif(worksheet, currentRowNum, yearValue, codeType, context) {
    console.log(`📊 [ANNUAL SUMIF] Finding annual column for year ${yearValue} and updating with SUMIF`);
    
    try {
        // Load row 2 values for annual columns K:P to find the matching year
        const annualRange = worksheet.getRange("K2:P2");
        annualRange.load("values");
        await context.sync();
        
        const row2Values = annualRange.values[0]; // Get the single row array
        const annualColumns = ['K', 'L', 'M', 'N', 'O', 'P'];
        
        let targetColumn = null;
        
        // Find the column with the matching year in row 2
        for (let i = 0; i < row2Values.length; i++) {
            const cellValue = row2Values[i];
            if (cellValue == yearValue) { // Use == for type-flexible comparison
                targetColumn = annualColumns[i];
                console.log(`📊 [ANNUAL SUMIF] Found year ${yearValue} in annual column ${targetColumn}2`);
                break;
            }
        }
        
        if (!targetColumn) {
            console.warn(`📊 [ANNUAL SUMIF] Year ${yearValue} not found in annual columns K2:P2, skipping SUMIF update`);
            return;
        }
        
        // Determine SUMIF type based on code type
        let sumifType = "year"; // Default for SPREAD-E
        if (codeType === "CONST-E" || codeType === "ENDPOINT-E") {
            sumifType = "average";
        }
        
        console.log(`📊 [ANNUAL SUMIF] Code type ${codeType} -> SUMIF type: ${sumifType}`);
        
        // Create the SUMIF formula based on type
        let sumifFormula = "";
        if (sumifType === "year") {
            sumifFormula = "=SUMIF($3:$3, INDIRECT(ADDRESS(2,COLUMN(),2)), INDIRECT(ROW() & \":\" & ROW()))";
        } else if (sumifType === "average") {
            sumifFormula = "=AVERAGEIF($3:$3, INDIRECT(ADDRESS(2,COLUMN(),2)), INDIRECT(ROW() & \":\" & ROW()))";
        }
        
        // Apply the formula to the target annual column
        const targetCell = worksheet.getRange(`${targetColumn}${currentRowNum}`);
        targetCell.formulas = [[sumifFormula]];
        
        // Apply black font color to match other SUMIF formulas
        targetCell.format.font.color = "#000000";
        
        await context.sync();
        
        console.log(`✅ [ANNUAL SUMIF] Applied ${sumifType} SUMIF formula to ${targetColumn}${currentRowNum}: ${sumifFormula}`);
        
    } catch (error) {
        console.error(`❌ [ANNUAL SUMIF] Error updating annual column: ${error.message}`, error);
        throw error;
    }
}

/**
 * Applies symbol-based formatting to MonthsRow cells with blue font color
 * @param {Excel.Worksheet} worksheet - The worksheet containing the cells
 * @param {number} rowNum - The row number to format
 * @param {Array} monthValues - Array of values from the MonthsRow parameter
 * @param {Array} columnSequence - Array of column letters for monthly data
 * @returns {Promise<void>}
 */
async function applyMonthsRowSymbolFormatting(worksheet, rowNum, monthValues, columnSequence) {
    console.log(`Applying MonthsRow symbol-based formatting with blue font to row ${rowNum}`);
    
    // Define format configurations (same as regular formatting but with blue font)
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

    const PERCENTAGE_FORMAT = '_(* #,##0.0%;_(* (#,##0.0)%;_(* " -"?_)';
    const BLUE_FONT_COLOR = "#0000FF"; // Blue font color for MonthsRow values

    for (let x = 0; x < monthValues.length && x < columnSequence.length; x++) {
        const originalValue = monthValues[x];
        const colLetter = columnSequence[x];
        
        if (!originalValue) continue; // Skip empty values
        
        // Parse the value and extract formatting information
        const parsed = parseValueWithSymbols(originalValue);
        console.log(`  MonthsRow Column ${colLetter}: "${originalValue}" -> cleanValue: "${parsed.cleanValue}", format: ${parsed.formatType}, italic: ${parsed.isItalic}, currency: ${parsed.currencySymbol || 'none'}`);
        
        const cellRange = worksheet.getRange(`${colLetter}${rowNum}`);
        
        // Apply number format if specified
        if (parsed.formatType && formatConfigs[parsed.formatType]) {
            const config = formatConfigs[parsed.formatType];
            cellRange.numberFormat = [[config.numberFormat]];
            cellRange.format.font.italic = config.italic;
            console.log(`    Applied ${parsed.formatType} formatting to ${colLetter}${rowNum}`);
        }
        
        // IMMEDIATELY override with percentage format if this is a percentage
        if (parsed.formatType === 'percent') {
            console.log(`    Immediately overriding with percentage format for ${colLetter}${rowNum}`);
            cellRange.numberFormat = [[PERCENTAGE_FORMAT]];
            cellRange.format.font.italic = true; // Percentage values with ~ should be italic
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
        
        // ALWAYS apply blue font color for MonthsRow values
        cellRange.format.font.color = BLUE_FONT_COLOR;
        console.log(`    Applied blue font color to MonthsRow value in ${colLetter}${rowNum}`);
    }
    
    // Single sync for all formatting changes
    await worksheet.context.sync();
    console.log(`Completed MonthsRow symbol-based formatting with blue font for row ${rowNum}`);
}

/**
 * Applies symbol-based formatting to a row of cells, with immediate percentage format override
 * This combines volume formatting (based on ~) with percentage formatting in a single synchronous operation
 * @param {Excel.Worksheet} worksheet - The worksheet containing the cells
 * @param {number} rowNum - The row number to format
 * @param {Array} splitArray - Array of values from the row parameter
 * @param {Array} columnSequence - Array of column letters
 * @returns {Promise<void>}
 */
async function applyRowSymbolFormatting(worksheet, rowNum, splitArray, columnSequence) {
    console.log(`Applying symbol-based formatting with percentage override to row ${rowNum}`);
    
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

    const PERCENTAGE_FORMAT = '_(* #,##0.0%;_(* (#,##0.0)%;_(* " -"?_)';

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
            console.log(`    Applied ${parsed.formatType} formatting to ${colLetter}${rowNum}`);
        }
        
        // IMMEDIATELY override with percentage format if this is a percentage
        if (parsed.formatType === 'percent') {
            console.log(`    Immediately overriding with percentage format for ${colLetter}${rowNum}`);
            cellRange.numberFormat = [[PERCENTAGE_FORMAT]];
            cellRange.format.font.italic = true; // Percentage values with ~ should be italic
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
    
    // Single sync for all formatting changes
    await worksheet.context.sync();
    console.log(`Completed symbol-based formatting with percentage override for row ${rowNum}`);
}

/**
 * Processes ACTUALS codes and populates the Actuals tab with data
 * @param {Array} actualsCodes - Array of ACTUALS code objects to process
 * @returns {Promise<void>}
 */
async function processActualsCodes(actualsCodes) {
    if (!actualsCodes || actualsCodes.length === 0) {
        console.log("No ACTUALS codes to process");
        return;
    }

    console.log(`Processing ${actualsCodes.length} ACTUALS codes`);

    try {
        await Excel.run(async (context) => {
            // Get or create the Actuals worksheet
            let actualsSheet;
            try {
                actualsSheet = context.workbook.worksheets.getItem("Actuals");
                console.log("Found existing Actuals worksheet");
            } catch (error) {
                console.log("Actuals worksheet not found, creating new one");
                actualsSheet = context.workbook.worksheets.add("Actuals");
            }

            actualsSheet.load("name");
            await context.sync();

            // Find the next available row starting from row 2
            let currentRow = 2;
            const usedRange = actualsSheet.getUsedRange();
            if (usedRange) {
                usedRange.load("rowIndex, rowCount");
                await context.sync();
                const lastUsedRow = usedRange.rowIndex + usedRange.rowCount;
                currentRow = Math.max(2, lastUsedRow + 1); // Start from row 2 or after last used row
            }
            
            console.log(`Starting to insert ACTUALS data at row ${currentRow}`);

            // Process each ACTUALS code
            for (const code of actualsCodes) {
                console.log(`Processing ACTUALS code with parameters:`, code.params);
                
                // Get the values parameter (support both 'values' and 'data' for flexibility)
                const valuesData = code.params.values || code.params.data;
                
                if (!valuesData) {
                    console.warn("ACTUALS code missing 'values' parameter, skipping");
                    continue;
                }

                console.log(`Processing values data: ${valuesData}`);

                // Parse the CSV-like data: rows delimited by *, cells delimited by |
                const rows = valuesData.split('*');
                const dataToInsert = [];

                for (const row of rows) {
                    if (row.trim() === '') continue; // Skip empty rows
                    
                    const cells = row.split('|');
                    if (cells.length === 4) {
                        // Parse the amount as a number if possible
                        let amount = cells[1];
                        if (!isNaN(amount) && amount.trim() !== '') {
                            amount = parseFloat(amount);
                        }
                        
                        // Parse the date and convert to last day of the month
                        let date = cells[2];
                        if (date && date.trim() !== '') {
                            const parsedDate = new Date(date);
                            if (!isNaN(parsedDate.getTime())) {
                                // Convert to last day of the month
                                const year = parsedDate.getFullYear();
                                const month = parsedDate.getMonth();
                                // Create date for first day of next month, then subtract 1 day
                                const lastDayOfMonth = new Date(year, month + 1, 0);
                                
                                // Format as mm/dd/yyyy string
                                const formattedMonth = String(lastDayOfMonth.getMonth() + 1).padStart(2, '0');
                                const formattedDay = String(lastDayOfMonth.getDate()).padStart(2, '0');
                                const formattedYear = lastDayOfMonth.getFullYear();
                                date = `${formattedMonth}/${formattedDay}/${formattedYear}`;
                                
                                console.log(`  Converted date ${cells[2]} to last day of month: ${date}`);
                            }
                        }

                        dataToInsert.push([
                            cells[0], // Description
                            amount,   // Amount (number or string)
                            date,     // Date (Date object or string)
                            cells[3]  // Category
                        ]);
                        
                        console.log(`  Parsed row: ${cells[0]} | ${amount} | ${date} | ${cells[3]}`);
                    } else {
                        console.warn(`Invalid row format (expected 4 columns, got ${cells.length}): ${row}`);
                    }
                }

                // Insert the data into the worksheet
                if (dataToInsert.length > 0) {
                    const endRow = currentRow + dataToInsert.length - 1;
                    const dataRange = actualsSheet.getRange(`A${currentRow}:D${endRow}`);
                    dataRange.values = dataToInsert;
                    
                    // Format the data range
                    dataRange.format.borders.getItemAt(Excel.BorderIndex.insideHorizontal).style = "Continuous";
                    dataRange.format.borders.getItemAt(Excel.BorderIndex.insideVertical).style = "Continuous";
                    dataRange.format.borders.getItemAt(Excel.BorderIndex.edgeTop).style = "Continuous";
                    dataRange.format.borders.getItemAt(Excel.BorderIndex.edgeBottom).style = "Continuous";
                    dataRange.format.borders.getItemAt(Excel.BorderIndex.edgeLeft).style = "Continuous";
                    dataRange.format.borders.getItemAt(Excel.BorderIndex.edgeRight).style = "Continuous";
                    
                    // Format amount column (B) as currency
                    const amountRange = actualsSheet.getRange(`B${currentRow}:B${endRow}`);
                    amountRange.numberFormat = [["_(* $ #,##0_);_(* $ (#,##0);_(* \"$\" -\"\"?_);_(@_)"]];
                    
                    // Format date column (C) as date
                    const dateRange = actualsSheet.getRange(`C${currentRow}:C${endRow}`);
                    dateRange.numberFormat = [["mm/dd/yyyy"]];
                    
                    console.log(`Inserted ${dataToInsert.length} rows of data starting at row ${currentRow}`);
                    currentRow = endRow + 1;
                }
            }

            // Auto-fit columns
            actualsSheet.getRange("A:D").format.autofitColumns();

            await context.sync();
            console.log("Completed processing ACTUALS codes and populating Actuals worksheet");
        });
    } catch (error) {
        console.error("Error processing ACTUALS codes:", error);
        throw error;
    }
}

/**
 * Processes a single ACTUALS code immediately when encountered
 * @param {Object} code - The ACTUALS code object with parameters
 * @returns {Promise<void>}
 */
async function processActualsSingle(code) {
    console.log(`Processing single ACTUALS code with parameters:`, code.params);

    try {
        await Excel.run(async (context) => {
            // Get or create the Actuals worksheet
            let actualsSheet;
            try {
                actualsSheet = context.workbook.worksheets.getItem("Actuals");
                console.log("Found existing Actuals worksheet");
            } catch (error) {
                console.log("Actuals worksheet not found, creating new one");
                actualsSheet = context.workbook.worksheets.add("Actuals");
            }

            actualsSheet.load("name");
            await context.sync();

            // Find the next available row starting from row 2
            let currentRow = 2;
            const usedRange = actualsSheet.getUsedRange();
            if (usedRange) {
                usedRange.load("rowIndex, rowCount");
                await context.sync();
                const lastUsedRow = usedRange.rowIndex + usedRange.rowCount;
                currentRow = Math.max(2, lastUsedRow + 1); // Start from row 2 or after last used row
            }
            
            console.log(`Inserting ACTUALS data at row ${currentRow}`);

            // Get the values parameter (support both 'values' and 'data' for flexibility)
            const valuesData = code.params.values || code.params.data;
            
            if (!valuesData) {
                console.warn("ACTUALS code missing 'values' parameter, skipping");
                return;
            }

            console.log(`Processing values data: ${valuesData}`);

            // Parse the CSV-like data: rows delimited by *, cells delimited by |
            const rows = valuesData.split('*');
            const dataToInsert = [];

            for (const row of rows) {
                if (row.trim() === '') continue; // Skip empty rows
                
                const cells = row.split('|');
                if (cells.length === 4) {
                    // Parse the amount as a number if possible
                    let amount = cells[1];
                    if (!isNaN(amount) && amount.trim() !== '') {
                        amount = parseFloat(amount);
                    }
                    
                    // Parse the date and convert to last day of the month
                    let date = cells[2];
                    if (date && date.trim() !== '') {
                        const parsedDate = new Date(date);
                        if (!isNaN(parsedDate.getTime())) {
                            // Convert to last day of the month
                            const year = parsedDate.getFullYear();
                            const month = parsedDate.getMonth();
                            // Create date for first day of next month, then subtract 1 day
                            const lastDayOfMonth = new Date(year, month + 1, 0);
                            
                            // Format as mm/dd/yyyy string
                            const formattedMonth = String(lastDayOfMonth.getMonth() + 1).padStart(2, '0');
                            const formattedDay = String(lastDayOfMonth.getDate()).padStart(2, '0');
                            const formattedYear = lastDayOfMonth.getFullYear();
                            date = `${formattedMonth}/${formattedDay}/${formattedYear}`;
                            
                            console.log(`  Converted date ${cells[2]} to last day of month: ${date}`);
                        }
                    }

                    dataToInsert.push([
                        cells[0], // Description
                        amount,   // Amount (number or string)
                        date,     // Date (Date object or string)
                        cells[3]  // Category
                    ]);
                    
                    console.log(`  Parsed row: ${cells[0]} | ${amount} | ${date} | ${cells[3]}`);
                } else {
                    console.warn(`Invalid row format (expected 4 columns, got ${cells.length}): ${row}`);
                }
            }

            // Insert the data into the worksheet
            if (dataToInsert.length > 0) {
                const endRow = currentRow + dataToInsert.length - 1;
                const dataRange = actualsSheet.getRange(`A${currentRow}:D${endRow}`);
                dataRange.values = dataToInsert;
                
                // Format the data range
                dataRange.format.borders.getItemAt(Excel.BorderIndex.insideHorizontal).style = "Continuous";
                dataRange.format.borders.getItemAt(Excel.BorderIndex.insideVertical).style = "Continuous";
                dataRange.format.borders.getItemAt(Excel.BorderIndex.edgeTop).style = "Continuous";
                dataRange.format.borders.getItemAt(Excel.BorderIndex.edgeBottom).style = "Continuous";
                dataRange.format.borders.getItemAt(Excel.BorderIndex.edgeLeft).style = "Continuous";
                dataRange.format.borders.getItemAt(Excel.BorderIndex.edgeRight).style = "Continuous";
                
                // Format amount column (B) as currency
                const amountRange = actualsSheet.getRange(`B${currentRow}:B${endRow}`);
                amountRange.numberFormat = [["_(* $ #,##0_);_(* $ (#,##0);_(* \"$\" -\"\"?_);_(@_)"]];
                
                // Format date column (C) as date
                const dateRange = actualsSheet.getRange(`C${currentRow}:C${endRow}`);
                dateRange.numberFormat = [["mm/dd/yyyy"]];
                
                console.log(`Inserted ${dataToInsert.length} rows of data starting at row ${currentRow}`);
            }

            // Auto-fit columns
            actualsSheet.getRange("A:D").format.autofitColumns();

            await context.sync();
            console.log("Completed processing single ACTUALS code and populating Actuals worksheet");
        });
    } catch (error) {
        console.error("Error processing single ACTUALS code:", error);
        throw error;
    }
}

/**
 * Restores left borders to all cells in column S that have an interior color set
 * @param {Excel.Worksheet} worksheet - The worksheet to process
 * @returns {Promise<void>}
 */
async function restoreColumnSLeftBorders(worksheet) {
    console.log(`Restoring left borders to column S cells with interior color in ${worksheet.name}`);
    
    try {
        // Get the used range to determine how many rows to check
        const usedRange = worksheet.getUsedRange();
        if (!usedRange) {
            console.log(`  No used range found in ${worksheet.name}`);
            return;
        }
        
        usedRange.load("rowCount");
        await worksheet.context.sync();
        
        const totalRows = usedRange.rowCount;
        console.log(`  Checking ${totalRows} rows in column S`);
        
        // Process in batches to avoid performance issues
        const batchSize = 100;
        for (let startRow = 1; startRow <= totalRows; startRow += batchSize) {
            const endRow = Math.min(startRow + batchSize - 1, totalRows);
            const columnSRange = worksheet.getRange(`S${startRow}:S${endRow}`);
            
            // Load the interior color for each cell
            columnSRange.load("format/fill/color");
            await worksheet.context.sync();
            
            // Check each cell in this batch
            for (let i = 0; i < columnSRange.values.length; i++) {
                const rowNum = startRow + i;
                const cellColor = columnSRange.format.fill.color[i][0];
                
                // If the cell has an interior color set (not empty/default/white)
                if (cellColor && cellColor !== "" && cellColor !== "#FFFFFF") {
                    const singleCell = worksheet.getRange(`S${rowNum}`);
                    singleCell.format.borders.getItem('EdgeLeft').style = 'Continuous';
                    singleCell.format.borders.getItem('EdgeLeft').weight = 'Thin';
                }
            }
            
            await worksheet.context.sync();
        }
        
        console.log(`  Completed restoring left borders to column S in ${worksheet.name}`);
        
    } catch (error) {
        console.error(`Error restoring column S left borders in ${worksheet.name}: ${error.message}`);
    }
}

/**
 * Copies selective formatting from column O to column J and columns T through CN for a specific row
 * Only copies number format and italic format, preserving font colors
 * @param {Excel.Worksheet} worksheet - The worksheet containing the cells
 * @param {number} rowNum - The row number to copy formatting for
 * @returns {Promise<void>}
 */
async function copyColumnPFormattingToJAndTCN(worksheet, rowNum) {
    console.log(`Copying selective column O formatting (number format + italic) to J and T:CN for row ${rowNum}`);
    
    try {
        // Get the source cell and target ranges
        const sourceCellO = worksheet.getRange(`O${rowNum}`);
        const targetCellJ = worksheet.getRange(`J${rowNum}`);
        const targetRangeTCN = worksheet.getRange(`T${rowNum}:CN${rowNum}`);
        
        // Load specific formatting from source cell O
        sourceCellO.load(["numberFormat", "format/font/italic"]);
        await worksheet.context.sync();
        
        // Get the values to apply
        const numberFormat = sourceCellO.numberFormat[0][0];
        const isItalic = sourceCellO.format.font.italic;
        
        // Apply number format and italic to column J
        targetCellJ.numberFormat = [[numberFormat]];
        targetCellJ.format.font.italic = isItalic;
        
        // Apply number format and italic to T:CN range
        targetRangeTCN.numberFormat = [[numberFormat]];
        targetRangeTCN.format.font.italic = isItalic;
        
        // Sync the formatting changes
        await worksheet.context.sync();
        
        console.log(`Successfully copied selective column O formatting to J${rowNum} and T${rowNum}:CN${rowNum}`);
        console.log(`  Applied number format: ${numberFormat}`);
        console.log(`  Applied font italic: ${isItalic}`);
        console.log(`  Font colors preserved (not copied)`);
        
    } catch (error) {
        console.error(`Error copying selective column O formatting to J and T:CN for row ${rowNum}: ${error.message}`);
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

                    // COMMENTED OUT: Convert row references to absolute for columns >= AE
        // This section was making row references absolute, but user requested it to be disabled
        /*
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
        */
        // END COMMENTED OUT SECTION

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

                    // Storage for comments to be added after cell population
                    const cellComments = new Map(); // Map<columnLetter, comment>

                    for (let x = 0; x < splitArray.length; x++) {
                        // Check bounds for columnSequence
                        if (x >= columnSequence.length) {
                            console.warn(`Data item index ${x} exceeds columnSequence length (${columnSequence.length}). Skipping.`);
                            continue;
                        }

                        const originalValue = splitArray[x];
                        const colLetter = columnSequence[x];
                        const cellToWrite = currentWorksheet.getRange(`${colLetter}${currentRowNum}`);
                        
                        // Parse the value to extract comments from square brackets first
                        const commentParsed = parseCommentFromBrackets(originalValue);
                        const valueAfterCommentRemoval = commentParsed.cleanValue;
                        
                        // Store comment for later application if one exists
                        if (commentParsed.comment) {
                            cellComments.set(colLetter, commentParsed.comment);
                            console.log(`  Stored comment for ${colLetter}${currentRowNum}: "${commentParsed.comment}"`);
                        }
                        
                        // Parse the value to extract clean value without formatting symbols
                        const parsed = parseValueWithSymbols(valueAfterCommentRemoval);
                        const valueToWrite = parsed.cleanValue;
                        
                        // VBA check: If splitArray(x) <> "" And splitArray(x) <> "F" Then
                        // 'F' likely means "Formula", so we don't overwrite if the value is 'F'.
                        if (valueToWrite && valueToWrite.toUpperCase() !== 'F') {
                            // Attempt to infer data type (basic number check)
                            let numValue = Number(valueToWrite);
                            if (!isNaN(numValue) && valueToWrite.trim() !== '') {
                                // Special handling for percentage values
                                if (parsed.formatType === 'percent') {
                                    // Convert percentage to decimal (5% -> 0.05)
                                    numValue = numValue / 100;
                                    console.log(`  Converting percentage: ${valueToWrite}% -> ${numValue} (decimal)`);
                                }
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
                    
                                            // Apply customformula parameter to column U for row1 (FORMULA-S codes only)
                if (g === 1 && yy === 0 && code.type === "FORMULA-S" && code.params.customformula && code.params.customformula !== "0") {
                    try {
                        console.log(`  Applying customformula to U${currentRowNum} for FORMULA-S: ${code.params.customformula}`);
                        
                        // Parse comments from square brackets in customformula
                        const customFormulaParsed = parseCommentFromBrackets(code.params.customformula);
                        const cleanCustomFormula = customFormulaParsed.cleanValue;
                        
                        // For FORMULA-S, set it as a value that will be converted to formula later by processFormulaSRows
                        const customFormulaCell = currentWorksheet.getRange(`U${currentRowNum}`);
                        customFormulaCell.values = [[cleanCustomFormula]];
                        console.log(`  Set customformula as value for FORMULA-S processing (cleaned of comments)`);
                        
                        // Apply comment to U cell if one was extracted
                        if (customFormulaParsed.comment) {
                            console.log(`  Adding customformula comment "${customFormulaParsed.comment}" to U${currentRowNum}`);
                            try {
                                currentWorksheet.comments.add(`U${currentRowNum}`, customFormulaParsed.comment);
                                await context.sync(); // Sync the comment addition
                                console.log(`  Successfully applied customformula comment to U${currentRowNum}`);
                            } catch (commentError) {
                                console.error(`  Error applying customformula comment: ${commentError.message}`);
                            }
                        }
                        
                        // Track this row for processFormulaSRows since column D will be overwritten
                        addFormulaSRow(worksheetName, currentRowNum);
                    } catch (customFormulaError) {
                        console.error(`  Error applying customformula: ${customFormulaError.message}`);
                    }
                }

                // Apply customformula parameter to column I for row1 (COLUMNFORMULA-S codes only)
                if (g === 1 && yy === 0 && code.type === "COLUMNFORMULA-S" && code.params.customformula && code.params.customformula !== "0") {
                    try {
                        console.log(`  Applying customformula to I${currentRowNum} for COLUMNFORMULA-S: ${code.params.customformula}`);
                        
                        // Parse comments from square brackets in customformula
                        const customFormulaParsed = parseCommentFromBrackets(code.params.customformula);
                        const cleanCustomFormula = customFormulaParsed.cleanValue;
                        
                        // For COLUMNFORMULA-S, set it as a value that will be converted to formula later by processColumnFormulaSRows
                        const customFormulaCell = currentWorksheet.getRange(`I${currentRowNum}`);
                        customFormulaCell.values = [[cleanCustomFormula]];
                        console.log(`  Set customformula as value for COLUMNFORMULA-S processing (cleaned of comments)`);
                        
                        // Apply comment to I cell if one was extracted
                        if (customFormulaParsed.comment) {
                            console.log(`  Adding customformula comment "${customFormulaParsed.comment}" to I${currentRowNum}`);
                            try {
                                currentWorksheet.comments.add(`I${currentRowNum}`, customFormulaParsed.comment);
                                await context.sync(); // Sync the comment addition
                                console.log(`  Successfully applied customformula comment to I${currentRowNum}`);
                            } catch (commentError) {
                                console.error(`  Error applying customformula comment: ${commentError.message}`);
                            }
                        }
                        
                        // Track this row for processColumnFormulaSRows since column D will be overwritten
                        addColumnFormulaSRow(worksheetName, currentRowNum);
                    } catch (customFormulaError) {
                        console.error(`  Error applying customformula: ${customFormulaError.message}`);
                    }
                }

                    // Apply symbol-based formatting for each cell in the row
                    try {
                        // Create a cleaned split array without square bracket comments for formatting
                        const cleanedSplitArray = splitArray.map(value => {
                            const commentParsed = parseCommentFromBrackets(value);
                            return commentParsed.cleanValue;
                        });
                        await applyRowSymbolFormatting(currentWorksheet, currentRowNum, cleanedSplitArray, columnSequence);
                    } catch (formatError) {
                        console.error(`  Error applying symbol-based formatting: ${formatError.message}`);
                    }

                    // Apply comments extracted from square brackets in row data
                    if (cellComments.size > 0) {
                        console.log(`  Applying ${cellComments.size} comments extracted from row data for row ${currentRowNum}`);
                        try {
                            for (const [columnLetter, comment] of cellComments) {
                                const cellAddress = `${columnLetter}${currentRowNum}`;
                                console.log(`    Adding comment "${comment}" to ${cellAddress}`);
                                currentWorksheet.comments.add(cellAddress, comment);
                            }
                            await context.sync(); // Sync the comment additions
                            console.log(`  Successfully applied all row data comments for row ${currentRowNum}`);
                        } catch (commentError) {
                            console.error(`  Error applying row data comments: ${commentError.message}`);
                        }
                    }

                    // Process MonthsRow for this specific row (yy=0 means first row of the group)
                    if (yy === 0) {
                        // Check for MonthsRow parameter for this g value with "monthsr" prefix
                        const monthsRowParam = code.params[`monthsr${g}`];
                        
                        if (monthsRowParam) {
                            console.log(`  Processing monthsr${g} for row ${currentRowNum}: ${monthsRowParam}`);
                            console.log(`  Code type: ${code.type} - will auto-populate entire year based on type`);

                            // Split the values by pipes
                            const monthValues = monthsRowParam.split('|');
                            
                            // Create column sequence starting from U
                            const monthsColumnSequence = generateMonthsColumnSequence(monthValues.length);
                            
                            // Storage for comments to be added after cell population
                            const monthsCellComments = new Map(); // Map<columnLetter, comment>

                            // First, populate the initial monthsr values as before
                            for (let x = 0; x < monthValues.length; x++) {
                                if (x >= monthsColumnSequence.length) {
                                    console.warn(`  monthsr${g} value index ${x} exceeds available monthly columns. Skipping.`);
                                    continue;
                                }

                                const originalValue = monthValues[x];
                                const colLetter = monthsColumnSequence[x];
                                const cellToWrite = currentWorksheet.getRange(`${colLetter}${currentRowNum}`);
                                
                                // Parse the value to extract comments from square brackets first
                                const commentParsed = parseCommentFromBrackets(originalValue);
                                const valueAfterCommentRemoval = commentParsed.cleanValue;
                                
                                // Store comment for later application if one exists
                                if (commentParsed.comment) {
                                    monthsCellComments.set(colLetter, commentParsed.comment);
                                    console.log(`    Stored monthsr${g} comment for ${colLetter}${currentRowNum}: "${commentParsed.comment}"`);
                                }
                                
                                // Parse the value to extract clean value without formatting symbols
                                const parsed = parseValueWithSymbols(valueAfterCommentRemoval);
                                const valueToWrite = parsed.cleanValue;
                                
                                // Write the value if it's not empty and not 'F'
                                if (valueToWrite && valueToWrite.toUpperCase() !== 'F') {
                                    // Attempt to infer data type (basic number check)
                                    let numValue = Number(valueToWrite);
                                    if (!isNaN(numValue) && valueToWrite.trim() !== '') {
                                        // Special handling for percentage values
                                        if (parsed.formatType === 'percent') {
                                            // Convert percentage to decimal (5% -> 0.05)
                                            numValue = numValue / 100;
                                            console.log(`    Converting monthsr${g} percentage: ${valueToWrite}% -> ${numValue} (decimal)`);
                                        }
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
                                    console.log(`    Wrote monthsr${g} value '${valueToWrite}' to ${colLetter}${currentRowNum}`);
                                } else if (valueToWrite === '' || valueToWrite === null || valueToWrite === undefined) {
                                    // Clear the cell if the value is explicitly empty
                                    cellToWrite.clear(Excel.ClearApplyTo.contents);
                                    console.log(`    Cleared contents of ${colLetter}${currentRowNum} due to blank monthsr${g} value`);
                                }
                            }
                            
                            // Sync the initial values before proceeding with year-wide population
                            await context.sync();
                            
                            // Now auto-populate the entire year based on code type
                            try {
                                await autoPopulateEntireYear(currentWorksheet, currentRowNum, monthValues, monthsColumnSequence, code.type, context);
                            } catch (yearPopulateError) {
                                console.error(`    Error auto-populating entire year for monthsr${g}: ${yearPopulateError.message}`);
                            }
                            
                            // Apply symbol-based formatting for MonthsRow values (entire year)
                            try {
                                // Create a cleaned split array without square bracket comments for formatting
                                const cleanedMonthValues = monthValues.map(value => {
                                    const commentParsed = parseCommentFromBrackets(value);
                                    return commentParsed.cleanValue;
                                });
                                await applyMonthsRowSymbolFormatting(currentWorksheet, currentRowNum, cleanedMonthValues, monthsColumnSequence);
                            } catch (formatError) {
                                console.error(`    Error applying monthsr${g} symbol-based formatting: ${formatError.message}`);
                            }

                            // Apply comments extracted from square brackets in MonthsRow data
                            if (monthsCellComments.size > 0) {
                                console.log(`    Applying ${monthsCellComments.size} comments extracted from monthsr${g} data for row ${currentRowNum}`);
                                try {
                                    for (const [columnLetter, comment] of monthsCellComments) {
                                        const cellAddress = `${columnLetter}${currentRowNum}`;
                                        console.log(`      Adding monthsr${g} comment "${comment}" to ${cellAddress}`);
                                        currentWorksheet.comments.add(cellAddress, comment);
                                    }
                                    await context.sync(); // Sync the comment additions
                                    console.log(`    Successfully applied all monthsr${g} comments for row ${currentRowNum}`);
                                } catch (commentError) {
                                    console.error(`    Error applying monthsr${g} comments: ${commentError.message}`);
                                }
                            }
                        }
                    }

                    // Apply columnformula parameter to column U for row1 (all code types)
                    if (g === 1 && yy === 0 && code.params.columnformula && code.params.columnformula !== "0") {
                        try {
                            console.log(`  Processing columnformula for U${currentRowNum}: ${code.params.columnformula}`);
                            
                            // Parse comments from square brackets in columnformula
                            const columnFormulaParsed = parseCommentFromBrackets(code.params.columnformula);
                            const cleanColumnFormula = columnFormulaParsed.cleanValue;
                            
                            // Process the formula through parseFormulaSCustomFormula to handle BEG, END, etc.
                            let processedFormula = await parseFormulaSCustomFormula(cleanColumnFormula, currentRowNum, currentWorksheet, context);
                            console.log(`  Processed columnformula result: ${processedFormula}`);
                            
                            // Check for negative parameter and apply it to the processed formula
                            if (code.params.negative && String(code.params.negative).toUpperCase() === "TRUE") {
                                console.log(`  Applying negative transformation to columnformula`);
                                processedFormula = `-(${processedFormula})`;
                                console.log(`  Columnformula after negative transformation: ${processedFormula}`);
                            }
                            
                            // Apply the processed formula to the cell
                            const columnFormulaCell = currentWorksheet.getRange(`U${currentRowNum}`);
                            if (processedFormula && processedFormula !== cleanColumnFormula) {
                                // If the formula was modified, set it as a formula
                                columnFormulaCell.formulas = [['=' + processedFormula]];
                                console.log(`  Set processed columnformula as formula: =${processedFormula}`);
                            } else {
                                // If no processing occurred, set as original value
                                columnFormulaCell.values = [[cleanColumnFormula]];
                                console.log(`  Set original columnformula as value: ${cleanColumnFormula}`);
                            }
                            
                            // Apply comment to U cell if one was extracted
                            if (columnFormulaParsed.comment) {
                                console.log(`  Adding columnformula comment "${columnFormulaParsed.comment}" to U${currentRowNum}`);
                                try {
                                    currentWorksheet.comments.add(`U${currentRowNum}`, columnFormulaParsed.comment);
                                    await context.sync(); // Sync the comment addition
                                    console.log(`  Successfully applied columnformula comment to U${currentRowNum}`);
                                } catch (commentError) {
                                    console.error(`  Error applying columnformula comment: ${commentError.message}`);
                                }
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
                                targetCell.load("values");
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

                    // FINAL STEP: Copy selective column O formatting to J and T:CN (number format + italic only)
                    try {
                        await copyColumnPFormattingToJAndTCN(currentWorksheet, currentRowNum);
                    } catch (copyFormatError) {
                        console.error(`  Error copying selective column O formatting to J and T:CN: ${copyFormatError.message}`);
                    }

                    // NOTE: Percentage formatting override is now handled within applyRowSymbolFormatting
                    // No need for separate reapplyPercentageFormatting call
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
    const TARGET_COL = "U";      // Column where the result address string is written

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
    const TARGET_COL = "U";

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
                
                // COMMENTED OUT: Convert row references to absolute for columns >= TARGET_COL
                // This section was making row references absolute, but user requested it to be disabled
                /*
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
                */
                // END COMMENTED OUT SECTION
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
    const ASSUMPTION_LINK_COL_C = "C";
    const ASSUMPTION_LINK_COL_D = "D";
    // Column on assumption sheet to link for monthly data
    const ASSUMPTION_MONTHS_START_COL = "U";

    const FINANCIALS_CODE_COLUMN = "I"; // Column to search for code on Financials sheet
    const FINANCIALS_TARGET_COL_B = "B";
    const FINANCIALS_TARGET_COL_C = "C";
    const FINANCIALS_TARGET_COL_D = "D";
    const FINANCIALS_ANNUALS_START_COL = "J"; // Annuals start here
    const FINANCIALS_MONTHS_START_COL = "U"; // Months start here

    // --- Updated Column Definitions ---
    const ANNUALS_END_COL = "P";       // Annuals end here
    const MONTHS_END_COL = "CN";       // Months end here
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
                const linkFormulaC = `='${worksheet.name}'!${ASSUMPTION_LINK_COL_C}${assumptionRow}`;
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
                    addressC: linkFormulaC,     // Use the constructed formula link
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
            const cellC = financialsSheet.getRange(`${FINANCIALS_TARGET_COL_C}${populateRow}`);
            const cellD = financialsSheet.getRange(`${FINANCIALS_TARGET_COL_D}${populateRow}`);
            const cellAnnualsStart = financialsSheet.getRange(`${FINANCIALS_ANNUALS_START_COL}${populateRow}`);
            const cellMonthsStart = financialsSheet.getRange(`${FINANCIALS_MONTHS_START_COL}${populateRow}`);

            // --- Populate Column B ---
            cellB.formulas = [[task.addressB]]; // Set formula directly
            cellB.format.font.bold = false;
            cellB.format.font.italic = false;
            cellB.format.indentLevel = 2;

            // --- Populate Column C ---
            cellC.formulas = [[task.addressC]]; // Set formula directly
            cellC.format.font.bold = false;
            cellC.format.font.italic = false;
            cellC.format.indentLevel = 2;

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
            
            // --- NEW: Populate Actuals Column T with SUMIFS formula ---
            try {
                const actualsCell = financialsSheet.getRange(`T${populateRow}`);
                const sumifsFormula = "=SUMIFS(Actuals!$B:$B,Actuals!$C:$C,EOMONTH(INDIRECT(ADDRESS(2,COLUMN())),0),Actuals!$D:$D,@INDIRECT(ADDRESS(ROW(),2)))";
                
                // Set the formula for column T only
                actualsCell.formulas = [[sumifsFormula]];
                
                // Apply formatting
                actualsCell.format.numberFormat = CURRENCY_FORMAT;
                actualsCell.format.font.bold = false;
                actualsCell.format.font.italic = false;
                actualsCell.format.font.color = "#7030A0"; // Set font color
                console.log(`  Set SUMIFS formula for T${populateRow}`);
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
    startTimer("processAssumptionTabs-total");
    console.log(`Starting processing for ${assumptionTabNames.length} assumption tabs:`, assumptionTabNames);
    if (!assumptionTabNames || assumptionTabNames.length === 0) {
        console.log("No assumption tabs provided to process.");
        endTimer("processAssumptionTabs-total");
        return;
    }

    const FINANCIALS_SHEET_NAME = "Financials"; // Define constant
    const AUTOFILL_START_COLUMN = "U";
    const AUTOFILL_END_COLUMN = "CN";
    const START_ROW = 10; // <<< CHANGED FROM 9 // Standard start row for processing

    try {
        // --- Loop through each assumption tab name ---
        for (const worksheetName of assumptionTabNames) {
             console.log(`\nProcessing Assumption Tab: ${worksheetName}`);
             startTimer(`processTab-${worksheetName}`);

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
                     startTimer(`adjustDriversJS-${worksheetName}`);
                     await adjustDriversJS(currentWorksheet, lastRow);
                     endTimer(`adjustDriversJS-${worksheetName}`);

                     // 3. Replace Indirects
                     startTimer(`replaceIndirectsJS-${worksheetName}`);
                     await replaceIndirectsJS(currentWorksheet, lastRow);
                     endTimer(`replaceIndirectsJS-${worksheetName}`);

                     // 4. Get Last Row Again (if Replace_Indirects might change it)
                     // const updatedLastRow = await getLastUsedRow(currentWorksheet, "B"); // Recalculate if necessary
                     const updatedLastRow = lastRow; // Assuming Replace_Indirects doesn't change last row for now
                     console.log(`Using last row for subsequent steps: ${updatedLastRow}`);
                     if (updatedLastRow < START_ROW) {
                         console.warn(`Skipping remaining steps for ${worksheetName} as updated last row (${updatedLastRow}) is invalid.`);
                         return;
                     }

                     // 5. Apply Index Growth Curve logic (MOVED UP - must run before Populate Financials)
                     // Run Index Growth *before* populating financials so financials links to aggregate rows
                     startTimer(`applyIndexGrowthCurveJS-${worksheetName}`);
                     await applyIndexGrowthCurveJS(currentWorksheet, updatedLastRow);
                     endTimer(`applyIndexGrowthCurveJS-${worksheetName}`);
                     
                     // 6. Get updated last row after Index Growth Curve (may have inserted rows)
                     const postIndexLastRow = await getLastUsedRow(currentWorksheet, "B");
                     console.log(`Last row after Index Growth Curve: ${postIndexLastRow}`);

                     // 7. Populate Financials (MOVED DOWN - now runs after Index Growth Curve)
                     startTimer(`populateFinancialsJS-${worksheetName}`);
                     await populateFinancialsJS(currentWorksheet, postIndexLastRow, financialsSheet);
                     endTimer(`populateFinancialsJS-${worksheetName}`);

                     // 8. Set font color to white in column A
                     // We use postIndexLastRow here, as Index Growth may have added rows
                     startTimer(`setColumnAFontWhite-${worksheetName}`);
                     await setColumnAFontWhite(currentWorksheet, START_ROW, postIndexLastRow); 
                     endTimer(`setColumnAFontWhite-${worksheetName}`);
                     console.log(`Set font color to white in column A from rows ${START_ROW}-${postIndexLastRow}`); 
                     
                     // 9. Process FORMULA-S rows - Convert driver references to cell references BEFORE deleting green rows
                     console.log(`Processing FORMULA-S rows in ${worksheetName} (before green row deletion)...`);
                     startTimer(`processFormulaSRows-${worksheetName}`);
                     await processFormulaSRows(currentWorksheet, START_ROW, postIndexLastRow);
                     endTimer(`processFormulaSRows-${worksheetName}`);
                     console.log(`Finished processing FORMULA-S rows`);
                     
                     // Clear the tracked FORMULA-S rows for this worksheet (cleanup)
                     clearFormulaSRows(worksheetName);
                     
                     // 9.5. Process COLUMNFORMULA-S rows - Convert driver references to cell references BEFORE deleting green rows
                     console.log(`Processing COLUMNFORMULA-S rows in ${worksheetName} (before green row deletion)...`);
                     startTimer(`processColumnFormulaSRows-${worksheetName}`);
                     await processColumnFormulaSRows(currentWorksheet, START_ROW, postIndexLastRow);
                     endTimer(`processColumnFormulaSRows-${worksheetName}`);
                     console.log(`Finished processing COLUMNFORMULA-S rows`);
                     
                     // Clear the tracked COLUMNFORMULA-S rows for this worksheet (cleanup)
                     clearColumnFormulaSRows(worksheetName);
                     
                     // 10. Delete rows with green background (#CCFFCC) - AFTER FORMULA-S and COLUMNFORMULA-S processing
                     console.log(`Deleting green rows in ${worksheetName}...`);
                     startTimer(`deleteGreenRows-${worksheetName}`);
                     // Changed START_ROW to START_ROW - 1 to include row 9
                     const finalLastRow = await deleteGreenRows(currentWorksheet, START_ROW - 1, postIndexLastRow); // Get the new last row AFTER deletions
                     endTimer(`deleteGreenRows-${worksheetName}`);
                     console.log(`After deleting green rows, last row is now: ${finalLastRow}`);
                     
                     // 11. Update SUMIF formulas after green row deletion
                     console.log(`Updating SUMIF formulas after green row deletion in ${worksheetName}...`);
                     startTimer(`updateSumifFormulasAfterGreenDeletion-${worksheetName}`);
                     await updateSumifFormulasAfterGreenDeletion(currentWorksheet, START_ROW, finalLastRow);
                     endTimer(`updateSumifFormulasAfterGreenDeletion-${worksheetName}`);
                     console.log(`Completed SUMIF formula updates for ${worksheetName}`);
 
                     // 12. Autofill U<startRow>:U<lastRow> -> CN<lastRow> on Assumption Tab - Only rows with formulas in U
                     console.log(`Checking for formulas in ${AUTOFILL_START_COLUMN}${START_ROW}:${AUTOFILL_START_COLUMN}${finalLastRow} and autofilling to ${AUTOFILL_END_COLUMN} on ${worksheetName}`);
                     
                     startTimer(`autofillFormulas-${worksheetName}`);
                     // Load formulas from column U to check which rows have formulas
                     const aeFormulaRange = currentWorksheet.getRange(`${AUTOFILL_START_COLUMN}${START_ROW}:${AUTOFILL_START_COLUMN}${finalLastRow}`);
                     aeFormulaRange.load("formulas");
                     await context.sync();
                     
                     let autofillCount = 0;
                     let skippedCount = 0;
                     // Check each row and autofill only those with formulas in U (skip hardcoded values)
                     for (let rowIndex = 0; rowIndex < aeFormulaRange.formulas.length; rowIndex++) {
                         const formula = aeFormulaRange.formulas[rowIndex][0];
                         const currentRowNum = START_ROW + rowIndex;
                         
                         // Only autofill if there's an actual formula (starts with =), skip hardcoded values
                         if (formula && typeof formula === 'string' && formula.startsWith('=')) {
                             console.log(`  Autofilling formula row ${currentRowNum}: ${AUTOFILL_START_COLUMN}${currentRowNum} -> ${AUTOFILL_END_COLUMN}${currentRowNum} (Formula: ${formula.substring(0, 50)}...)`);
                             const sourceCell = currentWorksheet.getRange(`${AUTOFILL_START_COLUMN}${currentRowNum}`);
                             const fillRowRange = currentWorksheet.getRange(`${AUTOFILL_START_COLUMN}${currentRowNum}:${AUTOFILL_END_COLUMN}${currentRowNum}`);
                             sourceCell.autoFill(fillRowRange, Excel.AutoFillType.fillDefault);
                             autofillCount++;
                         } else if (formula && typeof formula === 'string' && formula !== '') {
                             // Log when we skip hardcoded values
                             console.log(`  Skipping hardcoded value row ${currentRowNum}: U${currentRowNum} contains "${formula}" (not a formula)`);
                             skippedCount++;
                         } else if (formula === '' || formula === null || formula === undefined) {
                             // Log when we skip empty cells
                             console.log(`  Skipping empty cell row ${currentRowNum}: U${currentRowNum} is empty`);
                             skippedCount++;
                         } else {
                             // Log unexpected cases
                             console.log(`  Skipping unexpected value row ${currentRowNum}: U${currentRowNum} contains ${typeof formula}: "${formula}"`);
                             skippedCount++;
                         }
                     }
                     console.log(`Autofilled ${autofillCount} formula rows, skipped ${skippedCount} non-formula rows, out of ${aeFormulaRange.formulas.length} total rows`);
                     endTimer(`autofillFormulas-${worksheetName}`);
 
                     // 13. Set Row 9 interior color to none
                     console.log(`Setting row 9 interior color to none for ${worksheetName}`);
                     const row9Range = currentWorksheet.getRange("9:9");
                     row9Range.format.fill.clear();

                     // Sync all batched operations for this tab
                     await context.sync();
                     console.log(`Finished processing and syncing for tab ${worksheetName}`);

                 }); // End Excel.run for single tab processing
                 
                 endTimer(`processTab-${worksheetName}`);

             } catch (tabError) {
                 endTimer(`processTab-${worksheetName}`);
                 console.error(`Error processing tab ${worksheetName}:`, tabError);
                 // Optionally add to an error list and continue with the next tab
                 // Be mindful that subsequent tabs might depend on this one succeeding.
             }
        } // --- End loop through assumption tabs ---

        // --- Final Operations on Financials Sheet ---
        console.log(`\nPerforming final operations on ${FINANCIALS_SHEET_NAME}`);
        startTimer("finalOperations-Financials");
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
             endTimer("finalOperations-Financials");
         } catch (financialsError) {
             endTimer("finalOperations-Financials");
             console.error(`Error during final operations on ${FINANCIALS_SHEET_NAME}:`, financialsError);
         }

        console.log("Finished processing all assumption tabs.");
        endTimer("processAssumptionTabs-total");
        
        // Print timing summary for processAssumptionTabs
        console.log("\n" + "=".repeat(80));
        console.log("⏱️  PROCESS ASSUMPTION TABS TIMING SUMMARY");
        console.log("=".repeat(80));
        
        const tabTimes = Array.from(functionTimes.entries())
            .filter(([name]) => name.includes('processTab-') || name.includes('adjustDriversJS-') || 
                               name.includes('replaceIndirectsJS-') || name.includes('applyIndexGrowthCurveJS-') ||
                               name.includes('populateFinancialsJS-') || name.includes('processFormulaSRows-') ||
                               name.includes('deleteGreenRows-') || name.includes('updateSumifFormulasAfterGreenDeletion-') ||
                               name.includes('autofillFormulas-') || name.includes('setColumnAFontWhite-') ||
                               name.includes('finalOperations-') || name === 'processAssumptionTabs-total')
            .sort((a, b) => b[1] - a[1]);
        
        for (const [functionName, time] of tabTimes) {
            console.log(`${functionName.padEnd(50)} ${time.toFixed(3)}s`);
        }
        console.log("=".repeat(80));

    } catch (error) {
        endTimer("processAssumptionTabs-total");
        console.error("Error in processAssumptionTabs main function:", error);
        // Print partial timing summary even on error
        printTimingSummary();
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

                    // Hide Columns T-AC (keep S visible)
                    const colRange2 = worksheet.getRange("T:AC");
                    colRange2.columnHidden = true;
                    console.log(`  Hiding columns T-AC`);

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
// Test function to verify Excel API works with a simple workbook
export async function testExcelInsertion() {
    try {
        console.log(`[testExcelInsertion] Creating minimal test workbook...`);
        
        // Create a minimal Excel workbook programmatically
        await Excel.run(async (context) => {
            const newWorksheet = context.workbook.worksheets.add("TestSheet");
            newWorksheet.getCell(0, 0).values = [["Test Data"]];
            await context.sync();
            console.log(`[testExcelInsertion] ✅ Successfully created test worksheet programmatically`);
        });
        
        return true;
    } catch (error) {
        console.error(`[testExcelInsertion] ❌ Failed to create test worksheet:`, error);
        return false;
    }
}

// Test function to create minimal valid Excel base64 for testing
export function createMinimalExcelBase64() {
    // This is a minimal valid XLSX file (empty workbook) encoded as base64
    // Generated from a real Excel file with just one empty worksheet
    return "UEsDBBQACAgIAJuV1lYAAAAAAAAAAAAAAAALAAAAX3JlbHMvLnJlbHOkkdFqwzAMRf+F8H5r3GxbCyNEaAfbwx5KG8a2xzBLlmO1+fdxU9o19KEQhBDce3WlwzOD1pFYczfFYqDmOVvIbpqGlIqOtNR8yJE5TwCPGwMGBa1p5F3B+fYeQJhWBaJjZNQ9yFJbf7tTDR5zZXe4o1A8w0y3G6+xq1VeGjZCWZJRKCu5zUd4Q/Xq3wZpj13IjINPl9m3Nh9d5vfTMw3NdtqElTaABbN2Bpf1Nt7/a2rV8q/yWi+32lZy7N2G6yktP6+5L7vpOl0s5xfZy49K1B/nAQAA//8DAFBLAwQUAAgICACbldZWAAAAAAAAAAAAAAAAEQAAAGRvY1Byb3BzL2FwcC54bWytkdFqwzAMRf+F8H5rbY0S09Km0NKydQ9ls8e2DyAPOZbaOMaWwf9+7gOBwsYY2JPVPV5dcnQPEbmfJE1wSQWJLdHEIXS9bKJuQnOpJqtXc6e1RzZNxI9v/nKWlHWUyE8jTjWCNNQ/plKZ6kNrWDU0XFGJSpsZcfGJyqNpfJgSKw6S1tVhpzJkfzydtCRbQAQGJKzTc3QUJwQ20pJPVktGEfZrXH/PNk9WjdGG7XhQqb/8Z8jtQ3xj/2K7m47fzLKWMfT8AzaVPx4jFQQ2V/X6AwAA//8DAFBLAwQUAAgICACbldZWAAAAAAAAAAAAAAAAEwAAAGRvY1Byb3BzL2NvcmUueG1stZNNSwMxEIb/SljevE7SH7Zpu1aQKh6KiFi8hm22M82H2SRtu/z3Zluxtuj15NvJzLzvzJv0LlnKsZ4cJ1JXFZ6NIxwJqlkrqn6Fn55vh2MOJm9JLZQK/ZJEIuNDlIi+qbBnXSC7WJvI5yOLqYqwCcEWOY59p6kJcWlKJU0InSr/1JtIvjBUtzfryJHt3uGK9wz9HDf3i4KVJJQDLqpR4lUdcnj7/e2RBn2eS7f3+5XxJr/k4UPgF9PY20ysRJEPOKt9LRK9Qs4V9rCe+cj2RKO9aBV4YVe8LbYHcg2TQ7rMHFYfqWHjlWrwlq/6CvTDZDO/Gq5DjH8DKjBTCQlMYNuoF8U1u5xgm4qKTQY1y4uJO8lDfn1DY+3WfLNDZWJb69WfzZHLlXX6+oItmznKdNLFVsq29CQh+1QHSP44IaVYkxfeDTY/x9d/m0yWf4X8Ij7YzrBNOaKdAZzQQdrOWCc7Q5Jb9gYAAA//8DAFBLAwQUAAgICACbldZWAAAAAAAAAAAAAAAAGAAAAGN1c3RvbVByb3BlcnRpZXMvYXBwLnhtbI2PwQqCQBRF/+VlXdNJNIcZbaJlJYGZrscnPcfHo5k1/X0TbTJ27enCPe5x7jrrcAr8bpoUMT6IIQIljaKm7QLeFwdKIqaZSgP8K5b9gKqBjFHXtl4Zmt1QJHfGHsQ2U9p8Qql5pLT8Wqlk2n9rYe3aWOOV0ZSWv8W/I0+hMJEYx/HWyZppVOjfVdW1NRr5rFEP47q4KG8aZFaWr82aJtGhLbUNuP6L1fHPQz8AAAD//wMAUEsDBBQACAgIAJuV1lYAAAAAAAAAAAAAAAAMAAAAX3JlbHMvLnJlbHOkkdFqwzAMRf9F0Hut8ZPbZFVMaQfbwh5K6wa9M5ZuWvZ2LJ3+/epmu2EQGCKQJO7V9Qi/b1hhWKGpZZGO2ORJsLQllKDCNsXqeK6iWTlNUcmQn9+/fKD7llY0qUltTl8o1B7gT3K31v1eOFNMqLKEq4sRQCOT3gCJdoAz5Xg+FXXLzIBaFPeZJaiqTcGLhT2a5DnM8mPAmllOaBNhFyHLVoQ3Aq1UtnEbXQZgeCOOk3W20iMvAmz1E2/K/jw5oM+1hZB32AO0Qjdj7M/t9T/e/VY7r4EE+yZmlfKPGvkZJTn8j7/m0tJ0DnOhbvU1vlPX6+vFNQAA//8DAFBLAwQUAAgICACbldZWAAAAAAAAAAAAAAAADwAAAHhsL3dvcmtib29rLnhtbJRSW0sDMRR+l96d5C7bN8VLF7uCFy5aFC+I4kOaZmdjzCQxyLa0f9+Mx6qgCnUgk3PO953v8sJzv1HjJdhOG13BbJJBBbTQstP1Cr5eXs1eA9XdWKsNBu5iLLSsQKEpDxkYY0wSa9bZJLHMZmHoWNZqOZKGFkhGgU8k/gfC0IZJbYyzSWI4hs66qI6IQGzPgCiSJOZDwJONjjy2S1J9n0L7S+L9G3Jb/N5rp+zMHOSajVqGjjL7LUnkbHe7LL8jVz+9V5c2rqzjl6NJJlVr7XhFPBi6lnXuJI6WMb9WwPvqtKY5Z36vBTF5Gw6MQeAeGSv8tDnZq3kBVPFIQ/f/dJ0Qp7J8hwru3atDFLPF3LXKpR6HI/I4eiD6wLhbBe4Q7JdbL36JgK7jcBtyocPJ3/hJFhz/Kcbn5zFQ1mH9MNj0VHJOhyJo5Fm7+Q7rSJFn3MQ8j/NDlUlNJUJZmqe5c/IbDf2x6+6p0HQBF0Sj1EG5GsIejJqUgWGXf7TjKMrQPl3TjZlqFcx/6hZdK1ORkjnyAcVYTlkjZwJl7kJEusjKKlz6+YlHI7bxWcSU8jJy6s+9xaNnVKwNjL4dWj39AJ5/HQAA//8DAFBLAwQUAAgICACbldZWAAAAAAAAAAAAAAAAGgAAAHhsL3NoYXJlZFN0cmluZ3MueG1s9tG7CsIwGAbg+5P0D+GeJNy03xZpBQdBBBdduEyThObSeVLapqIP4Bu7ODg4ePMqx8Whu8LjuN1pN5oqN1IqP/OT39/xR43/CQAA//8DAFBLAwQUAAgICACbldZWAAAAAAAAAAAAAAAAFAAAAHhsL3RoZW1lL3RoZW1lMS54bWy1WV1vkzAUfT+pf8DyPp8aCQFBdO0M1e1j0lZte67xNaDa2NZHZ7O9bd/z6/a20bJ+TEg9+n+qDr73nHvuudfxh9vxlH1I3qMx6nZY1qgymC9s7u6yPmzDyutGJ26eM2yxXa6wKzw8gL7LQoVdJNGgfY3d0VhzHNdGhGW9y24gvYD6juNgCdFKEfNJNQYxhS+l4E5tMGUoMb1mf1P9+qrxJJ7Lhd+pjy3r6+7Y47m+gfrzrIz3Q6kkj2w+0xsj9HbFxgGDJZQ1Ul5jRSKxLfHMxjCFsjGJj4w3M8p1TJjKrN0h3sMStgv7/7R/urvfqhT2q1WJ91mJ19oO3X2X2eaEFQlJE7c9s7nNPfNWJa5ZNPfmO8iiJcfcaJ0dPKp9ql1hzXeZ41atP49FtgOy7QfZDsOy7YKN0DqT7QjrT3adEFW/3Woe/aqzrPfb9gOyXQkjGKKL1Og9E0L7Xub44KDCRI8L7X5vHJJhGOpKdnZZUZ4pJGLCXcI2EULa3PYJfg+vCHlrJOxtbR/7W+3+8mQ7t99hEUvd3lNUNFu3x9fhXC8FJf5RWZFMWZU8H7dPRYg7dAj8o6pNZZMjZJJgHdZqkxh1K6LrDR5J6IwMI1qQGl6bC7bvO9e63bTOjhRK/X1UJyOQ6hUlRZeXVzAQl7vMOYWDVPd2YiOUjHVFP5TlZT2yT5Nx6FxgJTXfBWGQ3qh6qP6jdKxIpQ7lldaD8gSvK72M0knMxCiGOJAj1tUo8VWCRykQ5i+Fl44aDpRPq77qNjYlq4+e8zK0KvpSinI8G1vOc8U8Xa1nPM5dIxhKiS7wIQM9MxJdYKxKqP72bXRJ5x7qqmrXsKiw2Fa8VRwSRqJJ6U7D2Pl/tQkqK6TcxRFQJKfKULlQG6OBwigWuLF6cZXRl99XKgWP3fD5j1+66f7L3o1Vj5tCGJp39UKOjfj3bFLJLzPn3Ug3X7dJWUHUtKJvK6vcvHjPeX7+9FJLhzF7K7MN9lRsqYh3/4a9lX7b/jF2Kt6aRUK1f6yPcMlL3VO/PUKlPOH7VlHOOLvYjbN6cZHfwz5S5E3aB7QkIyOl0gGBl7QRWnTaUK5GmOKkf2ddkbPJoI5WoxtbJcOsR3YKixGOKdm/bXTjq10ISJJ+Sz7vfW9b/e9D7+1YZdSXeaL8O8gOc5z9C";
}

export async function handleInsertWorksheetsFromBase64(base64String, sheetNames = null) {
    try {
        // Enhanced validation and debugging
        console.log(`[handleInsertWorksheetsFromBase64] Starting worksheet insertion`);
        console.log(`[handleInsertWorksheetsFromBase64] Base64 string type: ${typeof base64String}`);
        console.log(`[handleInsertWorksheetsFromBase64] Base64 string length: ${base64String ? base64String.length : 'null'}`);
        console.log(`[handleInsertWorksheetsFromBase64] Sheet names: ${sheetNames ? sheetNames.join(', ') : 'All sheets from source file'}`);
        
        // Validate base64 string
        if (!base64String || typeof base64String !== 'string') {
            throw new Error("Invalid base64 string provided - string is null, undefined, or not a string");
        }

        if (base64String.length === 0) {
            throw new Error("Base64 string is empty");
        }

        // Enhanced base64 format validation with detailed error message
        const base64Regex = /^[A-Za-z0-9+/]*={0,2}$/;
        if (!base64Regex.test(base64String)) {
            console.error(`[handleInsertWorksheetsFromBase64] Invalid base64 format. First 100 chars: ${base64String.substring(0, 100)}`);
            throw new Error("Invalid base64 format - contains invalid characters");
        }

        // Test base64 decode capability and check decoded size
        try {
            const testDecode = atob(base64String.substring(0, 100)); // Test a small portion
            console.log(`[handleInsertWorksheetsFromBase64] Base64 format validation passed`);
            console.log(`[handleInsertWorksheetsFromBase64] Test decode length: ${testDecode.length}`);
            
            // Calculate expected decoded size
            const expectedSize = (base64String.length * 3) / 4;
            console.log(`[handleInsertWorksheetsFromBase64] Expected decoded size: ${expectedSize} bytes`);
            console.log(`[handleInsertWorksheetsFromBase64] Base64 first 50 chars: ${base64String.substring(0, 50)}`);
            console.log(`[handleInsertWorksheetsFromBase64] Base64 last 50 chars: ${base64String.substring(base64String.length - 50)}`);
            
            // Check for size limitations (Excel may have limits)
            const sizeMB = expectedSize / (1024 * 1024);
            console.log(`[handleInsertWorksheetsFromBase64] File size: ${sizeMB.toFixed(2)} MB`);
            
            if (sizeMB > 50) {
                console.warn(`[handleInsertWorksheetsFromBase64] Large file warning: ${sizeMB.toFixed(2)} MB may exceed Excel API limits`);
            }
            
            // Verify this looks like Excel file header (Excel files start with specific bytes)
            try {
                const decodedStart = atob(base64String.substring(0, 20));
                const bytes = Array.from(decodedStart).map(c => c.charCodeAt(0).toString(16).padStart(2, '0')).join(' ');
                console.log(`[handleInsertWorksheetsFromBase64] File header bytes: ${bytes}`);
                
                // Excel files typically start with PK (ZIP signature: 50 4B) for .xlsx
                if (decodedStart.startsWith('PK')) {
                    console.log(`[handleInsertWorksheetsFromBase64] ✅ Valid Excel file signature detected`);
                } else {
                    console.error(`[handleInsertWorksheetsFromBase64] ❌ Invalid Excel file signature!`);
                    console.error(`[handleInsertWorksheetsFromBase64] Expected: PK (50 4B)`);
                    console.error(`[handleInsertWorksheetsFromBase64] Actual first 10 chars: "${decodedStart.substring(0, 10)}"`);
                    console.error(`[handleInsertWorksheetsFromBase64] Actual bytes: ${bytes}`);
                    console.error(`[handleInsertWorksheetsFromBase64] This indicates file corruption during build/fetch process`);
                }
            } catch (headerError) {
                console.warn(`[handleInsertWorksheetsFromBase64] Could not verify file header:`, headerError);
            }
        } catch (decodeError) {
            console.error(`[handleInsertWorksheetsFromBase64] Base64 decode test failed:`, decodeError);
            throw new Error(`Base64 string cannot be decoded: ${decodeError.message}`);
        }

        await Excel.run(async (context) => {
            const workbook = context.workbook;
            
            // Enhanced API version checking
            console.log(`[handleInsertWorksheetsFromBase64] Checking Excel API compatibility...`);
            
            if (!workbook.insertWorksheetsFromBase64) {
                console.error(`[handleInsertWorksheetsFromBase64] insertWorksheetsFromBase64 method not available`);
                throw new Error("This feature requires Excel API requirement set 1.13 or later");
            }
            
            console.log(`[handleInsertWorksheetsFromBase64] Excel API insertWorksheetsFromBase64 method is available`);
            
            // Try to get Excel version info if available
            try {
                const host = Office.context.host;
                const platform = Office.context.platform;
                const version = Office.context.requirements;
                console.log(`[handleInsertWorksheetsFromBase64] Excel host info:`, { host, platform });
                console.log(`[handleInsertWorsheetsFromBase64] Requirements info:`, version);
            } catch (infoError) {
                console.warn(`[handleInsertWorksheetsFromBase64] Could not get Excel version info:`, infoError);
            }
            
            // Quick test to verify basic Excel API functionality within current context
            console.log(`[handleInsertWorksheetsFromBase64] Testing basic Excel API functionality...`);
            try {
                const testWorksheet = workbook.worksheets.add("TempTestSheet");
                testWorksheet.getCell(0, 0).values = [["Test"]];
                await context.sync();
                testWorksheet.delete();
                await context.sync();
                console.log(`[handleInsertWorksheetsFromBase64] ✅ Basic Excel API test passed`);
            } catch (testError) {
                console.error(`[handleInsertWorksheetsFromBase64] ❌ Basic Excel API test failed:`, testError);
                throw new Error(`Basic Excel API test failed: ${testError.message}`);
            }
            
            // Insert the worksheets with enhanced error handling
            try {
                console.log(`[SHEET OPERATION] Calling Excel insertWorksheetsFromBase64...`);
                
                // Prepare options - be careful with empty objects
                let insertOptions;
                if (sheetNames && sheetNames.length > 0) {
                    insertOptions = { sheetNames: sheetNames };
                    console.log(`[SHEET OPERATION] Requesting specific sheets: ${sheetNames.join(', ')}`);
                } else {
                    insertOptions = undefined; // Don't pass empty object
                    console.log(`[SHEET OPERATION] Requesting all sheets from workbook`);
                }
                
                console.log(`[SHEET OPERATION] Insert options:`, insertOptions);
                
                // Try the API call with proper error context
                if (insertOptions) {
                    await workbook.insertWorksheetsFromBase64(base64String, insertOptions);
                } else {
                    await workbook.insertWorksheetsFromBase64(base64String);
                }
                
                await context.sync();
                console.log("Worksheets inserted successfully");
                console.log(`[SHEET OPERATION] Successfully inserted worksheets: ${sheetNames ? sheetNames.join(', ') : 'All sheets from source file'}`);
            } catch (error) {
                console.error("Error during worksheet insertion:", error);
                console.error("Error details:", {
                    name: error.name,
                    message: error.message,
                    code: error.code,
                    debugInfo: error.debugInfo
                });
                throw new Error(`Failed to insert worksheets: ${error.message} (Code: ${error.code || 'Unknown'})`);
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
    const DRIVER_REF_COL = "U"; // Column containing driver range ref in END_MARKER row
    const SUMIF_START_COL = "K"; // K
    const SUMIF_END_COL = "P"; // P
    const SUMPRODUCT_COL = "U"; // U (changed from AE to U)
    const MONTHS_START_COL = "U"; // U
    const MONTHS_END_COL = "CN"; // CN
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
        
        // Force an extra sync to ensure all previous operations are complete
        await context.sync();
        
        const searchRangeAddress = `${SEARCH_COL}${START_ROW}:${SEARCH_COL}${initialLastRow}`; 
        const searchRange = currentWorksheet.getRange(searchRangeAddress); // Use refreshed worksheet object
        searchRange.load("values");
        await context.sync(); // Use the context variable
 
        let firstRow = -1;
        let lastRow = -1;
        let indexEndRow = -1; // Keep track of the original END_MARKER row
 
        if (searchRange.values) {
            console.log(`🔍 [INDEX GROWTH] Searching column D from row ${START_ROW} to ${initialLastRow}:`);
            for (let i = 0; i < searchRange.values.length; i++) {
                const currentRow = START_ROW + i;
                const cellValue = searchRange.values[i][0];
                
                // Debug: Show all values in column D
                if (cellValue && cellValue !== '') {
                    console.log(`   Row ${currentRow}: "${cellValue}"`);
                }
                
                if (cellValue === BEGIN_MARKER && firstRow === -1) {
                    firstRow = currentRow;
                    console.log(`🎯 [INDEX GROWTH] Found INDEXBEGIN at row ${currentRow}`);
                }
                if (cellValue === END_MARKER) {
                    lastRow = currentRow; // This will be updated to the LAST END_MARKER found
                    indexEndRow = currentRow; // Store the original row index
                    console.log(`🎯 [INDEX GROWTH] Found INDEXEND at row ${currentRow}`);
                }
            }
            console.log(`🎯 [INDEX GROWTH] Final INDEXEND position determined as row ${lastRow} (last occurrence found)`);
        }

        if (firstRow === -1 || lastRow === -1 || lastRow < firstRow) {
            console.log(`Markers ${BEGIN_MARKER}/${END_MARKER} not found or in wrong order in ${searchRangeAddress}. Skipping Index Growth Curve.`);
            return; // Exit if markers not found or invalid
        }
        console.log(`🎯 [INDEX GROWTH] Found ${BEGIN_MARKER} at row ${firstRow}, final ${END_MARKER} at row ${lastRow}`);
 
                // --- 2. Collect Index Rows (Rows between markers where Col C is not empty) ---
        const indexRows = [];
        const DATA_COL_TO_CHECK = "C"; // Column to check for data
        const dataColRangeAddress = `${DATA_COL_TO_CHECK}${firstRow}:${DATA_COL_TO_CHECK}${lastRow}`;
        const dataColRange = currentWorksheet.getRange(dataColRangeAddress);
        dataColRange.load("values");
        await context.sync();

        // Check each row between markers for data in the specified column
        for (let i = 0; i < dataColRange.values.length; i++) {
            const currentRow = firstRow + i;
            const cellValue = dataColRange.values[i][0];
            
            // Skip the marker rows themselves
            if (currentRow === firstRow || currentRow === lastRow) {
                continue;
            }
            
            // Check if there's data in this row
            if (cellValue !== null && cellValue !== "" && String(cellValue).trim() !== "") {
                indexRows.push(currentRow);
            }
        }

        if (indexRows.length === 0) {
            console.log(`No data rows found between ${BEGIN_MARKER} and ${END_MARKER} in column ${DATA_COL_TO_CHECK}. Skipping rest of Index Growth Curve.`);
            return; // Exit if no data rows found
        }
        console.log(`Collected ${indexRows.length} index rows:`, indexRows);
 
        // --- 3. Set Background Color for non-green rows ---
        // Range: B(firstRow+2) to CX(lastRow-1) - format up to the row before INDEXEND
        const formatCheckStartRow = firstRow + 2;
        const formatCheckEndRow = lastRow - 1;
                console.log(`Setting background color for non-green rows between ${formatCheckStartRow} and ${formatCheckEndRow}`);
        if (formatCheckStartRow <= formatCheckEndRow) {
             // Check each row individually for background color
             for (let currentRow = formatCheckStartRow; currentRow <= formatCheckEndRow; currentRow++) {
                 try {
                     const checkCell = currentWorksheet.getRange(`${CHECK_COL_B}${currentRow}`);
                     checkCell.load("format/fill/color");
                     checkCell.load("values");
                     await context.sync();
                     
                     // Check if the cell is not light green
                     if (checkCell.format.fill.color !== LIGHT_GREEN_COLOR) {
                         console.log(`  Setting row ${currentRow} background to ${LIGHT_BLUE_COLOR}`);
                         const rowRange = currentWorksheet.getRange(`${currentRow}:${currentRow}`);
                         rowRange.format.fill.color = LIGHT_BLUE_COLOR;
                         
                         // Check if column B contains "BR" and set font color to blue
                         const cellValue = checkCell.values[0][0];
                         if (cellValue && String(cellValue).toUpperCase().includes("BR")) {
                             console.log(`  Setting row ${currentRow} font color to blue (contains BR)`);
                             rowRange.format.font.color = LIGHT_BLUE_COLOR;
                         }
                         
                         // Change blue font color to black in columns U through CN (but preserve MonthsRow blue)
                         console.log(`  Changing blue font to black in columns U:CN for row ${currentRow} (preserving MonthsRow blue)`);
                         const aeToChRange = currentWorksheet.getRange(`U${currentRow}:CN${currentRow}`);
                         // Load current font colors first to preserve blue colors from MonthsRow
                         aeToChRange.load("format/font/color");
                         await context.sync();
                         
                         // Only change non-blue fonts to black (preserve MonthsRow blue #0000FF)
                         const currentFontColors = aeToChRange.format.font.color;
                         if (Array.isArray(currentFontColors) && Array.isArray(currentFontColors[0])) {
                             // Multiple cells - check each one
                             const newColors = currentFontColors.map(row => 
                                 row.map(color => color === "#0000FF" ? "#0000FF" : "#000000")
                             );
                             aeToChRange.format.font.color = newColors;
                         } else if (typeof currentFontColors === 'string' && currentFontColors !== "#0000FF") {
                             // Single cell or uniform color - only change if not blue
                             aeToChRange.format.font.color = "#000000";
                         }
                         
                         // Clear fill in column A specifically
                         const cellARange = currentWorksheet.getRange(`A${currentRow}`);
                         cellARange.format.fill.clear();
                         await context.sync(); // Sync the formatting changes
                     } else {
                         console.log(`  Row ${currentRow} is green, skipping background change`);
                     }
                 } catch (colorError) {
                     console.warn(`  Error checking/setting color for row ${currentRow}: ${colorError.message}`);
                 }
             }
         }
 
         // --- 4. Insert Rows ---
         // Insert AFTER INDEXEND (below all INDEXEND rows - INDEXEND creates 3 rows)
         const newRowStart = lastRow + 3;
         const numNewRows = indexRows.length;
         const newRowEnd = newRowStart + numNewRows - 1;
         console.log(`🔧 [INDEX GROWTH] INSERTION DETAILS:`);
         console.log(`   INDEXBEGIN at row: ${firstRow}`);
         console.log(`   Last INDEXEND at row: ${lastRow}`);
         console.log(`   Will insert ${numNewRows} aggregate rows at: ${newRowStart}:${newRowEnd}`);
         console.log(`   This means aggregate rows will be AFTER all 3 INDEXEND rows (+3 offset)`);
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
                 // BS items use row 4 reference
                  baseFormula = `=SUMIF($4:$4, INDIRECT(ADDRESS(2,COLUMN())), ${targetRowNum}:${targetRowNum})`;
             } else {
                 // All other items use row 3 reference
                  baseFormula = `=SUMIF($3:$3, INDIRECT(ADDRESS(2,COLUMN())), ${targetRowNum}:${targetRowNum})`;
             }
             // Create array for the row
             sumifFormulas.push(Array(numSumifCols).fill(baseFormula));
         }
 
         const sumifRange = currentWorksheet.getRange(`${SUMIF_START_COL}${newRowStart}:${SUMIF_END_COL}${newRowEnd}`);
         sumifRange.formulas = sumifFormulas;
 
                  // --- 7. Apply SUMPRODUCT Formulas (AE) ---
         console.log(`Applying SUMPRODUCT formulas to ${SUMPRODUCT_COL}${newRowStart}:${SUMPRODUCT_COL}${newRowEnd}`);
         
         // Try to get driver name from column F of the INDEXBEGIN row (driver1 parameter)
         console.log(`Looking for driver name in cell F${firstRow} (INDEXBEGIN row)`);
         const driverNameCell = currentWorksheet.getRange(`F${firstRow}`);
         driverNameCell.load("values");
         await context.sync();
         const driverName = driverNameCell.values[0][0];
         
         console.log(`Driver name found: "${driverName}" (type: ${typeof driverName})`);
         
         let driverRangeString = null;
         
         if (driverName && typeof driverName === 'string' && driverName.trim() !== '') {
             // Look up the driver name in column A to find its row
             console.log(`Looking up driver "${driverName}" in column A to find its row...`);
             
             // Load column A values to find the driver row
             const colARangeAddress = `A${START_ROW}:A${initialLastRow}`;
             const colARange = currentWorksheet.getRange(colARangeAddress);
             colARange.load("values");
             await context.sync();
             
             let driverRow = -1;
             for (let i = 0; i < colARange.values.length; i++) {
                 const cellValue = colARange.values[i][0];
                 if (cellValue === driverName) {
                     driverRow = START_ROW + i;
                     console.log(`Found driver "${driverName}" at row ${driverRow}`);
                     break;
                 }
             }
             
                           if (driverRow !== -1) {
                  // Create driver range using the found row (modified format: $AE$11:AE$11)
                  driverRangeString = `$${MONTHS_START_COL}$${driverRow}:${MONTHS_START_COL}$${driverRow}`;
                  console.log(`Created driver range from driver row: ${driverRangeString}`);
             } else {
                 console.warn(`Driver "${driverName}" not found in column A. Will use default range.`);
             }
         } else {
             console.warn(`No valid driver name found in F${firstRow}. Will use default range.`);
         }

         // Fallback to default range if no driver found
         if (!driverRangeString) {
             console.log(`Using default driver range (row 1 headers)`);
             driverRangeString = `$${MONTHS_START_COL}$1:${MONTHS_START_COL}1`;
         }
         
         console.log(`Final driver range to use: ${driverRangeString}`);
         console.log(`Setting SUMPRODUCT formulas for ${indexRows.length} target rows...`);
         
         // Build array of SUMPRODUCT formulas for batch assignment
         const sumproductFormulas = [];
         for (let i = 0; i < indexRows.length; i++) {
             const originalRow = indexRows[i];
             const targetRow = newRowStart + i;
             
             // Debug logging for data range construction
             console.log(`  DEBUG: MONTHS_START_COL = "${MONTHS_START_COL}", MONTHS_END_COL = "${MONTHS_END_COL}", originalRow = ${originalRow}`);
             const dataRangeString = `$${MONTHS_START_COL}$${originalRow}:${MONTHS_START_COL}${originalRow}`;
             console.log(`  DEBUG: dataRangeString = "${dataRangeString}"`);
             
             // Formula: =SUMPRODUCT(INDEX(driverRange, N(IF({1}, MAX(COLUMN(driverRange)) - COLUMN(driverRange) + 1))), dataRange)
             const sumproductFormula = `=SUMPRODUCT(INDEX(${driverRangeString},N(IF({1},MAX(COLUMN(${driverRangeString}))-COLUMN(${driverRangeString})+1))), ${dataRangeString})`;

             console.log(`  Prepared formula for ${SUMPRODUCT_COL}${targetRow}: ${sumproductFormula}`);
             sumproductFormulas.push([sumproductFormula]);
         }
         
         // Set all SUMPRODUCT formulas at once using array assignment
         const sumproductRange = currentWorksheet.getRange(`${SUMPRODUCT_COL}${newRowStart}:${SUMPRODUCT_COL}${newRowEnd}`);
         sumproductRange.formulas = sumproductFormulas;
         await context.sync(); // Sync the SUMPRODUCT formulas
         console.log(`Successfully set ${indexRows.length} SUMPRODUCT formulas using array assignment`);

                  // NOTE: SUMIF formula updates moved to after green row deletion
 
         // --- 9. Copy Formats and Adjust ---
         console.log(`Copying formats and adjusting for new rows ${newRowStart}:${newRowEnd}`);
         
         // Copy formats from source rows to target rows (done individually due to non-contiguous source rows)
         for (let i = 0; i < indexRows.length; i++) {
             const sourceRow = indexRows[i];
             const targetRow = newRowStart + i;
 
             const sourceRowRange = currentWorksheet.getRange(`${sourceRow}:${sourceRow}`);
             const targetRowRange = currentWorksheet.getRange(`${targetRow}:${targetRow}`);
 
             // Copy formats first
             targetRowRange.copyFrom(sourceRowRange, Excel.RangeCopyType.formats);
         }
         await context.sync(); // Sync all format copies at once
         
         // Apply bulk format overrides to all new rows at once (preserve MonthsRow blue fonts)
         const allNewRowsRange = currentWorksheet.getRange(`${newRowStart}:${newRowEnd}`);
         
         // Load current font colors to preserve blue colors from MonthsRow before overriding
         const monthsRowRange = currentWorksheet.getRange(`U${newRowStart}:CN${newRowEnd}`);
         monthsRowRange.load("format/font/color");
         await context.sync();
         
         // Set black font for the entire range first
         allNewRowsRange.format.font.color = "#000000"; // Black font for all rows
         await context.sync();
         
         // Then restore blue fonts in the U:CN range where they should be preserved
         const currentMonthsFontColors = monthsRowRange.format.font.color;
         if (Array.isArray(currentMonthsFontColors) && Array.isArray(currentMonthsFontColors[0])) {
             // Multiple cells - restore blue where it was
             const preservedColors = currentMonthsFontColors.map(row => 
                 row.map(color => color === "#0000FF" ? "#0000FF" : "#000000")
             );
             monthsRowRange.format.font.color = preservedColors;
             console.log(`  Preserved blue font colors in MonthsRow columns U:CN for aggregated rows`);
         }
         allNewRowsRange.format.fill.clear(); // Clear interior color for all rows
         allNewRowsRange.format.font.bold = false; // Remove bold for all rows
         
         // Clear all borders for the entire range at once
         try {
             allNewRowsRange.format.borders.getItem('EdgeTop').style = 'None';
             allNewRowsRange.format.borders.getItem('EdgeBottom').style = 'None';
             allNewRowsRange.format.borders.getItem('EdgeLeft').style = 'None';
             allNewRowsRange.format.borders.getItem('EdgeRight').style = 'None';
             allNewRowsRange.format.borders.getItem('InsideVertical').style = 'None';
             allNewRowsRange.format.borders.getItem('InsideHorizontal').style = 'None';
         } catch (borderError) {
             console.warn(`  Error clearing borders for range ${newRowStart}:${newRowEnd}: ${borderError.message}`);
         }
         
         // Set indent level for column B in all new rows at once
         const columnBRange = currentWorksheet.getRange(`${OUTPUT_COL_B}${newRowStart}:${OUTPUT_COL_B}${newRowEnd}`);
         columnBRange.format.indentLevel = 2;
         
         await context.sync(); // Sync all format changes at once
 
         // --- 10. Clear Original Column C values ---
         console.log(`Clearing values in original index rows (${indexRows.join(', ')}) column ${DATA_COL}`);
         // Build array of ranges to clear for batch operation
         const rangesToClear = indexRows.map(row => currentWorksheet.getRange(`${DATA_COL}${row}`));
         
         // Clear all ranges (Excel will handle this efficiently)
         rangesToClear.forEach(range => range.clear(Excel.ClearApplyTo.contents));
         await context.sync(); // Sync all clears at once

         // --- 11. Clear Column J and Columns Q-S within INDEXBEGIN/INDEXEND block ---
         console.log(`Clearing column J values and columns Q-S (values + formats) within INDEXBEGIN/INDEXEND block (rows ${firstRow} to ${lastRow})`);
         
         // Clear column J values only
         const columnJRange = currentWorksheet.getRange(`J${firstRow}:J${lastRow}`);
         columnJRange.clear(Excel.ClearApplyTo.contents);
         console.log(`  Cleared column J values from rows ${firstRow} to ${lastRow}`);
         
         // Clear columns Q-S (both values and formats)
         const columnsQSRange = currentWorksheet.getRange(`Q${firstRow}:S${lastRow}`);
         columnsQSRange.clear(Excel.ClearApplyTo.all); // Clear both contents and formats
         console.log(`  Cleared columns Q-S (values + formats) from rows ${firstRow} to ${lastRow}`);
         
         await context.sync(); // Sync the clearing operations

         // NOTE: Row grouping for INDEXBEGIN moved to runCodes function where the original rows are copied
 
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
            // Get the current value in column U
            const aeCell = worksheet.getRange(`U${rowNum}`);
            aeCell.load("values");
            await worksheet.context.sync();
            
            const originalValue = aeCell.values[0][0];
            if (!originalValue || originalValue === "") {
                console.log(`  Row ${rowNum}: No value in U, skipping`);
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
                    const replacement = `U$${driverRow}`;
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
            
            // Replace timeseriesdivisor with U$7
            formula = formula.replace(/timeseriesdivisor/gi, (match) => {
                const replacement = 'U$7';
                console.log(`    Replacing ${match} with ${replacement}`);
                return replacement;
            });
            
            // Replace currentmonth with EOMONTH(U$2,0)
            formula = formula.replace(/currentmonth/gi, (match) => {
                const replacement = 'EOMONTH(U$2,0)';
                console.log(`    Replacing ${match} with ${replacement}`);
                return replacement;
            });
            
            // Replace beginningmonth with EOMONTH($U$2,0)
            formula = formula.replace(/beginningmonth/gi, (match) => {
                const replacement = 'EOMONTH($U$2,0)';
                console.log(`    Replacing ${match} with ${replacement}`);
                return replacement;
            });
            
            // Replace currentyear with U$3
            formula = formula.replace(/currentyear/gi, (match) => {
                const replacement = 'U$3';
                console.log(`    Replacing ${match} with ${replacement}`);
                return replacement;
            });
            
            // Replace yearend with U$4
            formula = formula.replace(/yearend/gi, (match) => {
                const replacement = 'U$4';
                console.log(`    Replacing ${match} with ${replacement}`);
                return replacement;
            });
            
            // Replace beginningyear with $U$3
            formula = formula.replace(/beginningyear/gi, (match) => {
                const replacement = '$U$3';
                console.log(`    Replacing ${match} with ${replacement}`);
                return replacement;
            });
            
            // Now parse special functions (SPREAD, BEG, END, etc.)
            console.log(`  Parsing special functions in formula...`);
            
            // Process SPREAD function: SPREAD(driver) -> driver/U$7
            formula = formula.replace(/SPREAD\(([^)]+)\)/gi, (match, driver) => {
                console.log(`    Converting SPREAD(${driver}) to ${driver}/U$7`);
                return `(${driver}/U$7)`;
            });
            
            // Process BEG function: BEG(driver) -> (EOMONTH(driver,0)<=EOMONTH(U$2,0))
            formula = formula.replace(/BEG\(([^)]+)\)/gi, (match, driver) => {
                console.log(`    Converting BEG(${driver}) to (EOMONTH(${driver},0)<=EOMONTH(U$2,0))`);
                return `(EOMONTH(${driver},0)<=EOMONTH(U$2,0))`;
            });
            
            // Process END function: END(driver) -> (EOMONTH(driver,0)>EOMONTH(U$2,0))
            formula = formula.replace(/END\(([^)]+)\)/gi, (match, driver) => {
                console.log(`    Converting END(${driver}) to (EOMONTH(${driver},0)>EOMONTH(U$2,0))`);
                return `(EOMONTH(${driver},0)>EOMONTH(U$2,0))`;
            });
            
            // Process RAISE function: RAISE(driver1,driver2) -> (1 + (driver1)) ^ (U$3 - max(year(driver2), $U3))
            formula = formula.replace(/RAISE\(([^,]+),([^)]+)\)/gi, (match, driver1, driver2) => {
                // Trim whitespace from drivers
                driver1 = driver1.trim();
                driver2 = driver2.trim();
                console.log(`    Converting RAISE(${driver1},${driver2}) to (1 + (${driver1})) ^ (U$3 - max(year(${driver2}), $U3))`);
                return `(1 + (${driver1})) ^ (U$3 - max(year(${driver2}), $U3))`;
            });
            
            // Process ANNBONUS function: ANNBONUS(driver1,driver2) -> (driver1*(MONTH(EOMONTH(U$2,0))=driver2))
            formula = formula.replace(/ANNBONUS\(([^,]+),([^)]+)\)/gi, (match, driver1, driver2) => {
                // Trim whitespace from drivers
                driver1 = driver1.trim();
                driver2 = driver2.trim();
                console.log(`    Converting ANNBONUS(${driver1},${driver2}) to (${driver1}*(MONTH(EOMONTH(U$2,0))=${driver2}))`);
                return `(${driver1}*(MONTH(EOMONTH(U$2,0))=${driver2}))`;
            });
            
            // Process QUARTERBONUS function: QUARTERBONUS(driver1) -> (driver1*(U$6<>0))
            formula = formula.replace(/QUARTERBONUS\(([^)]+)\)/gi, (match, driver1) => {
                // Trim whitespace from driver
                driver1 = driver1.trim();
                console.log(`    Converting QUARTERBONUS(${driver1}) to (${driver1}*(U$6<>0))`);
                return `(${driver1}*(U$6<>0))`;
            });
            
            // Process ONETIMEDATE function: ONETIMEDATE(driver) -> (EOMONTH((driver),0)=EOMONTH(U$2,0))
            formula = formula.replace(/ONETIMEDATE\(([^)]+)\)/gi, (match, driver) => {
                console.log(`    Converting ONETIMEDATE(${driver}) to (EOMONTH((${driver}),0)=EOMONTH(U$2,0))`);
                return `(EOMONTH((${driver}),0)=EOMONTH(U$2,0))`;
            });
            
            // Process SPREADDATES function: SPREADDATES(driver1,driver2,driver3) -> IF(AND(EOMONTH(EOMONTH(U$2,0),0)>=EOMONTH(driver2,0),EOMONTH(EOMONTH(U$2,0),0)<=EOMONTH(driver3,0)),driver1/(DATEDIF(driver2,driver3,"m")+1),0)
            // Note: Need to handle nested parentheses and comma separation
            formula = formula.replace(/SPREADDATES\(([^,]+),([^,]+),([^)]+)\)/gi, (match, driver1, driver2, driver3) => {
                // Trim whitespace from drivers
                driver1 = driver1.trim();
                driver2 = driver2.trim();
                driver3 = driver3.trim();
                const newFormula = `IF(AND(EOMONTH(EOMONTH(U$2,0),0)>=EOMONTH(${driver2},0),EOMONTH(EOMONTH(U$2,0),0)<=EOMONTH(${driver3},0)),${driver1}/(DATEDIF(${driver2},${driver3},"m")+1),0)`;
                console.log(`    Converting SPREADDATES(${driver1},${driver2},${driver3}) to ${newFormula}`);
                return newFormula;
            });
            
                    // Process AMORT function: AMORT(driver1,driver2,driver3,driver4) -> IFERROR(PPMT(driver1/U$7,1,driver2-U$8+1,SUM(driver3,OFFSET(driver4,-1,0),0,0),0)
        formula = formula.replace(/AMORT\(([^,]+),([^,]+),([^,]+),([^)]+)\)/gi, (match, driver1, driver2, driver3, driver4) => {
            // Trim whitespace from drivers
            driver1 = driver1.trim();
            driver2 = driver2.trim();
            driver3 = driver3.trim();
            driver4 = driver4.trim();
            const newFormula = `IFERROR(PPMT(${driver1}/U$7,1,${driver2}-U$8+1,SUM(${driver3},OFFSET(${driver4},-1,0),0,0),0)`;
            console.log(`    Converting AMORT(${driver1},${driver2},${driver3},${driver4}) to ${newFormula}`);
            return newFormula;
        });
            
            // Process BULLET function: BULLET(driver1,driver2,driver3) -> IF(driver2=U$8,SUM(driver1,OFFSET(driver3,-1,0)),0)
            formula = formula.replace(/BULLET\(([^,]+),([^,]+),([^)]+)\)/gi, (match, driver1, driver2, driver3) => {
                // Trim whitespace from drivers
                driver1 = driver1.trim();
                driver2 = driver2.trim();
                driver3 = driver3.trim();
                const newFormula = `IF(${driver2}=U$8,SUM(${driver1},OFFSET(${driver3},-1,0)),0)`;
                console.log(`    Converting BULLET(${driver1},${driver2},${driver3}) to ${newFormula}`);
                return newFormula;
            });
            
            // Process PBULLET function: PBULLET(driver1,driver2) -> (driver1*(driver2=U$8)) (Fixed partial principal payment)
            formula = formula.replace(/PBULLET\(([^,]+),([^)]+)\)/gi, (match, driver1, driver2) => {
                // Trim whitespace from drivers
                driver1 = driver1.trim();
                driver2 = driver2.trim();
                const newFormula = `(${driver1}*(${driver2}=U$8))`;
                console.log(`    Converting PBULLET(${driver1},${driver2}) to ${newFormula}`);
                return newFormula;
            });
            
            // Process INTONLY function: INTONLY(driver1) -> (U$8>driver1) (Interest only period)
            formula = formula.replace(/INTONLY\(([^)]+)\)/gi, (match, driver1) => {
                // Trim whitespace from driver
                driver1 = driver1.trim();
                const newFormula = `(U$8>${driver1})`;
                console.log(`    Converting INTONLY(${driver1}) to ${newFormula}`);
                return newFormula;
            });
            
                    // Process ONETIMEINDEX function: ONETIMEINDEX(driver1) -> (U$8=driver1) (One-time event at specific index month)
        formula = formula.replace(/ONETIMEINDEX\(([^)]+)\)/gi, (match, driver1) => {
            // Trim whitespace from driver
            driver1 = driver1.trim();
            const newFormula = `(U$8=${driver1})`;
            console.log(`    Converting ONETIMEINDEX(${driver1}) to ${newFormula}`);
            return newFormula;
        });
        
        // Process BEGINDEX function: BEGINDEX(driver1) -> (U$8>=driver1) (Start from index month and continue)
        formula = formula.replace(/BEGINDEX\(([^)]+)\)/gi, (match, driver1) => {
            // Trim whitespace from driver
            driver1 = driver1.trim();
            const newFormula = `(U$8>=${driver1})`;
            console.log(`    Converting BEGINDEX(${driver1}) to ${newFormula}`);
            return newFormula;
        });
        
        // Process ENDINDEX function: ENDINDEX(driver1) -> (U$8<=driver1) (End at index month)
        formula = formula.replace(/ENDINDEX\(([^)]+)\)/gi, (match, driver1) => {
            // Trim whitespace from driver
            driver1 = driver1.trim();
            const newFormula = `(U$8<=${driver1})`;
            console.log(`    Converting ENDINDEX(${driver1}) to ${newFormula}`);
            return newFormula;
        });
        
            // Process RANGE function: RANGE(driver1) -> U{row}:CN{row} where {row} is the row number from driver1
    console.log(`    Checking for RANGE function in: "${formula}"`);
    const rangeMatchesParser = formula.match(/RANGE\(([^)]+)\)/gi);
    if (rangeMatchesParser) {
        console.log(`    Found RANGE matches: ${rangeMatchesParser.join(', ')}`);
    } else {
        console.log(`    No RANGE matches found`);
    }
    
    formula = formula.replace(/RANGE\(([^)]+)\)/gi, (match, driver1) => {
        // Trim whitespace from driver
        driver1 = driver1.trim();
        
        // Extract row number from driver1 (e.g., U$15 -> 15, U15 -> 15)
        const rowMatch = driver1.match(/[A-Z]+\$?(\d+)/);
        if (rowMatch) {
            const rowNumber = rowMatch[1];
            const newFormula = `U${rowNumber}:CN${rowNumber}`;
            console.log(`    Converting RANGE(${driver1}) to ${newFormula}`);
            return newFormula;
        } else {
            console.warn(`    Could not extract row number from ${driver1}, keeping as is`);
            return match; // Keep original if we can't parse
        }
    });

            // Process SUMTABLE function: SUMTABLE(driver1) -> SUM(driver1:OFFSET(INDIRECT(ADDRESS(ROW(),COLUMN())),-1,0))
    console.log(`    Checking for SUMTABLE function in: "${formula}"`);
    const sumtableMatchesParser = formula.match(/SUMTABLE\(([^)]+)\)/gi);
    if (sumtableMatchesParser) {
        console.log(`    Found SUMTABLE matches: ${sumtableMatchesParser.join(', ')}`);
    } else {
        console.log(`    No SUMTABLE matches found`);
    }
    
    formula = formula.replace(/SUMTABLE\(([^)]+)\)/gi, (match, driver1) => {
        // Trim whitespace from driver
        driver1 = driver1.trim();
        const newFormula = `SUM(${driver1}:OFFSET(INDIRECT(ADDRESS(ROW(),COLUMN())),-1,0))`;
        console.log(`    Converting SUMTABLE(${driver1}) to ${newFormula}`);
        return newFormula;
    });
    
                // Process RANGE function: RANGE(driver1) -> U{row}:CN{row} where {row} is the row number from driver1
            console.log(`    Checking for RANGE function in: "${formula}"`);
            const rangeMatchesColumnFormula = formula.match(/RANGE\(([^)]+)\)/gi);
            if (rangeMatchesColumnFormula) {
                console.log(`    Found RANGE matches: ${rangeMatchesColumnFormula.join(', ')}`);
            } else {
                console.log(`    No RANGE matches found`);
            }
            
            formula = formula.replace(/RANGE\(([^)]+)\)/gi, (match, driver1) => {
                // Trim whitespace from driver
                driver1 = driver1.trim();
                
                // Extract row number from driver1 (e.g., U$15 -> 15, U15 -> 15)
                const rowMatch = driver1.match(/[A-Z]+\$?(\d+)/);
                if (rowMatch) {
                    const rowNumber = rowMatch[1];
                    const newFormula = `U${rowNumber}:CN${rowNumber}`;
                    console.log(`    Converting RANGE(${driver1}) to ${newFormula}`);
                    return newFormula;
                } else {
                    console.warn(`    Could not extract row number from ${driver1}, keeping as is`);
                    return match; // Keep original if we can't parse
                }
            });

            // Process SUMTABLE function: SUMTABLE(driver1) -> SUM(driver1:OFFSET(INDIRECT(ADDRESS(ROW(),COLUMN())),-1,0))
            console.log(`    Checking for SUMTABLE function in: "${formula}"`);
            const sumtableMatchesFormulaSRows = formula.match(/SUMTABLE\(([^)]+)\)/gi);
            if (sumtableMatchesFormulaSRows) {
                console.log(`    Found SUMTABLE matches: ${sumtableMatchesFormulaSRows.join(', ')}`);
            } else {
                console.log(`    No SUMTABLE matches found`);
            }
            
            formula = formula.replace(/SUMTABLE\(([^)]+)\)/gi, (match, driver1) => {
                // Trim whitespace from driver
                driver1 = driver1.trim();
                const newFormula = `SUM(${driver1}:OFFSET(INDIRECT(ADDRESS(ROW(),COLUMN())),-1,0))`;
                console.log(`    Converting SUMTABLE(${driver1}) to ${newFormula}`);
                return newFormula;
            });
            
            // Process TABLEMIN function: TABLEMIN(driver1,driver2) -> MIN(SUM(driver1:OFFSET(INDIRECT(ADDRESS(ROW(),COLUMN())),-1,0)),driver2)
            console.log(`    Checking for TABLEMIN function in: "${formula}"`);
            const tableminMatches = formula.match(/TABLEMIN\(([^,]+),([^)]+)\)/gi);
            if (tableminMatches) {
                console.log(`    Found TABLEMIN matches: ${tableminMatches.join(', ')}`);
            } else {
                console.log(`    No TABLEMIN matches found`);
            }
            
            formula = formula.replace(/TABLEMIN\(([^,]+),([^)]+)\)/gi, (match, driver1, driver2) => {
                // Trim whitespace from drivers
                driver1 = driver1.trim();
                driver2 = driver2.trim();
                const newFormula = `MIN(SUM(${driver1}:OFFSET(INDIRECT(ADDRESS(ROW(),COLUMN())),-1,0)),${driver2})`;
                console.log(`    Converting TABLEMIN(${driver1},${driver2}) to ${newFormula}`);
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
 * Processes COLUMNFORMULA-S rows by converting driver references in column I to Excel formulas
 * @param {Excel.Worksheet} worksheet - The worksheet to process
 * @param {number} startRow - The starting row to search from
 * @param {number} lastRow - The last row to search to
 * @returns {Promise<void>}
 */
async function processColumnFormulaSRows(worksheet, startRow, lastRow) {
    console.log(`Processing COLUMNFORMULA-S rows in ${worksheet.name} from row ${startRow} to ${lastRow}`);
    
    try {
        // Get stored COLUMNFORMULA-S row positions (since column D may have been overwritten)
        const columnFormulaSRows = getColumnFormulaSRows(worksheet.name);
        
        if (columnFormulaSRows.length === 0) {
            console.log("No COLUMNFORMULA-S rows found in tracker");
            return;
        }
        
        console.log(`Found ${columnFormulaSRows.length} COLUMNFORMULA-S rows at: ${columnFormulaSRows.join(", ")}`);
        
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
            '1': 'D', '2': 'E', '3': 'F', '4': 'G', '5': 'H', '6': 'I', '7': 'J', '8': 'K', '9': 'L',
            '10': 'M', '11': 'N', '12': 'O', '13': 'P', '14': 'Q', '15': 'R', '16': 'S'
        };
        
        // Process each COLUMNFORMULA-S row
        for (const rowNum of columnFormulaSRows) {
            // Get the current value in column I
            const iCell = worksheet.getRange(`I${rowNum}`);
            iCell.load("values");
            await worksheet.context.sync();
            
            const originalValue = iCell.values[0][0];
            if (!originalValue || originalValue === "") {
                console.log(`  Row ${rowNum}: No value in I, skipping`);
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
                    const replacement = `U$${driverRow}`;
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
            
            // Apply the same function replacements as FORMULA-S
            formula = await parseFormulaSCustomFormula(formula, rowNum, worksheet, worksheet.context);
            
            // Ensure the formula starts with '='
            if (!formula.startsWith('=')) {
                formula = '=' + formula;
            }
            
            console.log(`  Row ${rowNum}: Final formula: ${formula}`);
            
            // Set the formula in the cell
            iCell.formulas = [[formula]];
        }
        
        // Sync all formula changes
        await worksheet.context.sync();
        console.log(`Successfully processed ${columnFormulaSRows.length} COLUMNFORMULA-S rows`);
        
    } catch (error) {
        console.error(`Error in processColumnFormulaSRows: ${error.message}`, error);
        throw error;
    }
}



/**
 * Hides Columns C-I, Rows 2-8, and specific Actuals columns on specified sheets,
 * then navigates to cell A1 of the Financials sheet and inserts model codes.
 * @param {string[]} assumptionTabNames - Array of assumption tab names created by runCodes.
 * @param {string} [originalModelCodes] - Optional: The original model codes string to insert into A1
 * @returns {Promise<void>}
 */
export async function hideColumnsAndNavigate(assumptionTabNames, originalModelCodes = null) { // Renamed and added parameter
    // Define Actuals columns
    const ACTUALS_START_COL = "T";
    const ACTUALS_END_COL = "T";

    try {
        startTimer("hideColumnsAndNavigate-total");
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

            // --- Queue hiding operations for target sheets ---
            for (const worksheet of worksheets.items) {
                const sheetName = worksheet.name;
                if (targetSheetNames.includes(sheetName)) { // Check if sheet is in our target list
                    console.log(`Queueing hide operations for: ${sheetName}`);
                    try {
                        // Hide Rows 2:10 (Applies to both)
                        const rows210 = worksheet.getRange("2:10");
                        rows210.rowHidden = true;

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
                             // Skip hiding actuals columns on assumption tabs to keep column S visible
                             console.log(`  -> Skipping actuals column hiding for assumption tab ${sheetName} to keep column S visible`);
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

            // --- Reset view to A1 on ALL worksheets to ensure proper view positioning ---
            console.log("Resetting view to A1 on ALL worksheets to ensure leftmost column visibility...");
            
            // Get all worksheets for comprehensive view reset
            const allWorksheetsForReset = context.workbook.worksheets;
            allWorksheetsForReset.load("items/name");
            await context.sync();
            
            // Reset view on every worksheet
            for (const worksheet of allWorksheetsForReset.items) {
                const sheetName = worksheet.name;
                try {
                    console.log(`  Resetting view for: ${sheetName}`);
                    
                    // Activate the sheet
                    worksheet.activate();
                    
                    // Select A1 to ensure view scrolls to top-left corner
                    const rangeA1 = worksheet.getRange("A1");
                    rangeA1.select();
                    
                    // Sync immediately to ensure view reset takes effect
                    await context.sync();
                    
                    console.log(`  ✅ View reset completed for ${sheetName} (A1 selected, view at leftmost position)`);
                } catch (error) {
                    console.error(`  ❌ Error resetting view for ${sheetName}: ${error.message}`);
                    // Continue to next sheet even if one fails
                }
            }
            
            // --- Set final selection on assumption tabs to J11 for data entry ---
            console.log("Setting final selection to J11 on assumption tabs for optimal data entry position...");
            for (const sheetName of assumptionTabNames) {
                try {
                    console.log(`  Setting final selection J11 for: ${sheetName}`);
                    const worksheet = context.workbook.worksheets.getItem(sheetName);
                    worksheet.activate();
                    const rangeJ11 = worksheet.getRange("J11"); // J11 for data entry
                    rangeJ11.select();
                    await context.sync();
                    console.log(`  ✅ Final selection J11 set for ${sheetName}`);
                } catch (error) {
                     console.error(`  ❌ Error setting final selection for ${sheetName}: ${error.message}`);
                     // Continue to next sheet even if one fails
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

            // --- Insert model codes into Financials A1 if provided ---
            if (originalModelCodes) {
                try {
                    console.log("📋 Inserting model codes into Financials!A1...");
                    const financialsSheet = context.workbook.worksheets.getItem("Financials");
                    const cellA1 = financialsSheet.getRange("A1");
                    
                    // Insert the model codes
                    cellA1.values = [[originalModelCodes]];
                    
                    // Apply hidden formatting
                    cellA1.format.font.color = "#FFFFFF"; // White font (hidden)
                    cellA1.format.font.size = 8; // Small font
                    cellA1.format.wrapText = false; // No wrap
                    
                    await context.sync();
                    console.log(`✅ Model codes inserted into Financials!A1 (${originalModelCodes.length} characters)`);
                } catch (codesError) {
                    console.error("❌ Error inserting model codes:", codesError.message);
                    // Continue - don't let this stop the process
                }
            }

            // --- Navigate to Financials sheet as final active sheet with proper view reset ---
            try {
                console.log("Setting Financials as final active sheet with proper view reset...");
                const financialsSheet = context.workbook.worksheets.getItem("Financials");
                
                // Activate the Financials sheet
                financialsSheet.activate();
                
                // First, reset view completely to A1 (ensures leftmost position)
                console.log("  Resetting Financials view to A1 for proper left alignment...");
                const rangeA1 = financialsSheet.getRange("A1");
                rangeA1.select();
                await context.sync(); // Sync to ensure view reset happens
                
                // Now set final selection to J11 for data entry
                console.log("  Setting final selection to J11 for optimal data entry...");
                const rangeJ11 = financialsSheet.getRange("J11");
                rangeJ11.select();
                
                // Optional: Try to reset zoom to 100%
                try {
                    const activeWorksheet = context.workbook.getActiveWorksheet();
                    if (activeWorksheet && activeWorksheet.load) {
                        // This is available in newer Excel API versions
                        console.log("  Attempting to reset zoom level to 100%...");
                        // Note: This may not be available in all Excel versions
                    }
                } catch (zoomError) {
                    console.log("  Zoom reset not available in this Excel version:", zoomError.message);
                }
                
                await context.sync(); // Sync the final state
                console.log("✅ Financials sheet view reset complete - A1 view position with J11 selected");
            } catch (navError) {
                console.error(`❌ Error setting up Financials sheet final view: ${navError.message}`, {
                    code: navError.code,
                    debugInfo: navError.debugInfo ? JSON.stringify(navError.debugInfo) : 'N/A'
                });
                // Do not throw here, allow the function to finish
            }

            console.log("Finished hideColumnsAndNavigate function.");

        }); // End Excel.run
        
        endTimer("hideColumnsAndNavigate-total");
        
        // Print timing summary for hideColumnsAndNavigate
        console.log("\n" + "=".repeat(80));
        console.log("⏱️  HIDE COLUMNS AND NAVIGATE TIMING SUMMARY");
        console.log("=".repeat(80));
        
        const hideNavTimes = Array.from(functionTimes.entries())
            .filter(([name]) => name.includes('hideColumnsAndNavigate'))
            .sort((a, b) => b[1] - a[1]);
        
        for (const [functionName, time] of hideNavTimes) {
            console.log(`${functionName.padEnd(50)} ${time.toFixed(3)}s`);
        }
        console.log("=".repeat(80));
        
    } catch (error) {
        endTimer("hideColumnsAndNavigate-total");
        // Catch errors from the Excel.run call itself
        console.error("Critical error in hideColumnsAndNavigate:", error);
        throw error; // Re-throw critical errors
    }
}

// Global storage for FORMULA-S row positions per worksheet
const formulaSRowTracker = new Map(); // Map<worksheetName, Set<rowNumber>>

// Global storage for COLUMNFORMULA-S row positions per worksheet
const columnFormulaSRowTracker = new Map(); // Map<worksheetName, Set<rowNumber>>

// Global storage for INDEXBEGIN special rows per worksheet
const indexBeginRowTracker = new Map(); // Map<worksheetName, {timeSeriesRow, yearRow, yearEndRow}>

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
 * Stores a COLUMNFORMULA-S row position for later processing
 * @param {string} worksheetName - Name of the worksheet
 * @param {number} rowNumber - Row number containing customformula
 */
function addColumnFormulaSRow(worksheetName, rowNumber) {
    if (!columnFormulaSRowTracker.has(worksheetName)) {
        columnFormulaSRowTracker.set(worksheetName, new Set());
    }
    columnFormulaSRowTracker.get(worksheetName).add(rowNumber);
    console.log(`  Tracked COLUMNFORMULA-S row ${rowNumber} for worksheet ${worksheetName}`);
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
 * Gets stored COLUMNFORMULA-S row positions for a worksheet
 * @param {string} worksheetName - Name of the worksheet
 * @returns {number[]} Array of row numbers
 */
function getColumnFormulaSRows(worksheetName) {
    const rows = columnFormulaSRowTracker.get(worksheetName);
    return rows ? Array.from(rows) : [];
}

/**
 * Clears stored FORMULA-S row positions for a worksheet (cleanup)
 * @param {string} worksheetName - Name of the worksheet
 */
function clearFormulaSRows(worksheetName) {
    formulaSRowTracker.delete(worksheetName);
}

/**
 * Clears stored COLUMNFORMULA-S row positions for a worksheet (cleanup)
 * @param {string} worksheetName - Name of the worksheet
 */
function clearColumnFormulaSRows(worksheetName) {
    columnFormulaSRowTracker.delete(worksheetName);
}

/**
 * Stores INDEXBEGIN special row positions for later formula updates
 * @param {string} worksheetName - Name of the worksheet
 * @param {number} timeSeriesRow - Row number of the time series row
 * @param {number} yearRow - Row number of the year row  
 * @param {number} yearEndRow - Row number of the year end row
 */
function setIndexBeginRows(worksheetName, timeSeriesRow, yearRow, yearEndRow) {
    indexBeginRowTracker.set(worksheetName, {
        timeSeriesRow: timeSeriesRow,
        yearRow: yearRow,
        yearEndRow: yearEndRow
    });
    console.log(`📋 [INDEXBEGIN TRACKER] Stored rows for ${worksheetName}: Time Series=${timeSeriesRow}, Year=${yearRow}, Year End=${yearEndRow}`);
}

/**
 * Gets stored INDEXBEGIN special row positions for a worksheet
 * @param {string} worksheetName - Name of the worksheet
 * @returns {Object|null} Object with timeSeriesRow, yearRow, yearEndRow or null if not found
 */
function getIndexBeginRows(worksheetName) {
    return indexBeginRowTracker.get(worksheetName) || null;
}

/**
 * Clears stored INDEXBEGIN row positions for a worksheet (cleanup)
 * @param {string} worksheetName - Name of the worksheet
 */
function clearIndexBeginRows(worksheetName) {
    indexBeginRowTracker.delete(worksheetName);
}

/**
 * Updates SUMIF formulas after green rows are deleted by finding INDEXBEGIN and calculating special rows
 * @param {Excel.Worksheet} worksheet - The worksheet to process
 * @param {number} startRow - The starting row to search from
 * @param {number} lastRow - The last row to search to
 * @returns {Promise<void>}
 */
async function updateSumifFormulasAfterGreenDeletion(worksheet, startRow, lastRow) {
    console.log(`🔄 [SUMIF POST-DELETE] Starting SUMIF formula updates after green row deletion`);
    console.log(`🔄 [SUMIF POST-DELETE] Searching for INDEXBEGIN in ${worksheet.name} from row ${startRow} to ${lastRow}`);
    
    try {
        // Search for the first row with "INDEXBEGIN" in column D
        let indexBeginRow = null;
        
        // Load column D values to search for INDEXBEGIN
        const columnDRange = worksheet.getRange(`D${startRow}:D${lastRow}`);
        columnDRange.load("values");
        await worksheet.context.sync();
        
        const columnDValues = columnDRange.values;
        
        for (let i = 0; i < columnDValues.length; i++) {
            const cellValue = columnDValues[i][0];
            if (cellValue === "INDEXBEGIN") {
                indexBeginRow = startRow + i;
                console.log(`📋 [SUMIF POST-DELETE] Found INDEXBEGIN at row ${indexBeginRow}`);
                break;
            }
        }
        
        if (!indexBeginRow) {
            console.log(`⚠️ [SUMIF POST-DELETE] No INDEXBEGIN found in column D from row ${startRow} to ${lastRow} - skipping SUMIF updates`);
            return;
        }
        
        // Calculate special rows based on INDEXBEGIN position
        const timeSeriesRow = indexBeginRow;        // INDEXBEGIN row is time series row
        const yearRow = indexBeginRow + 2;          // Go down 2 rows for year row
        const yearEndRow = indexBeginRow + 3;       // Go down 1 more row for year end row
        
        console.log(`📋 [SUMIF POST-DELETE] Special rows defined:`);
        console.log(`    Time Series row (INDEXBEGIN): ${timeSeriesRow}`);
        console.log(`    Year row (Time Series + 2): ${yearRow}`);
        console.log(`    Year End row (Year + 1): ${yearEndRow}`);
        
        // Create range strings for the entire rows
        const yearRowRange = `$${yearRow}:$${yearRow}`;
        const yearEndRowRange = `$${yearEndRow}:$${yearEndRow}`;
        
        console.log(`🔄 [SUMIF POST-DELETE] Range strings: Year row=${yearRowRange}, Year End row=${yearEndRowRange}`);
        
        // Now search for INDEXEND to find the range of rows to update
        let indexEndRow = null;
        for (let i = 0; i < columnDValues.length; i++) {
            const cellValue = columnDValues[i][0];
            if (cellValue === "INDEXEND") {
                indexEndRow = startRow + i;
                console.log(`📋 [SUMIF POST-DELETE] Found INDEXEND at row ${indexEndRow}`);
                break;
            }
        }
        
        if (!indexEndRow) {
            console.log(`⚠️ [SUMIF POST-DELETE] No INDEXEND found - will process all rows from INDEXBEGIN to end`);
            indexEndRow = lastRow;
        }
        
        // Process all rows between INDEXBEGIN and INDEXEND that might have SUMIF formulas
        let formulasUpdated = 0;
        for (let currentRow = indexBeginRow + 1; currentRow <= indexEndRow - 1; currentRow++) {
            console.log(`🔄 [SUMIF POST-DELETE] Processing formulas in row ${currentRow}`);
            
            // Load formulas from columns J through P for this row
            const formulaRange = worksheet.getRange(`J${currentRow}:P${currentRow}`);
            formulaRange.load("formulas");
            await worksheet.context.sync();
            
            const formulas = formulaRange.formulas[0]; // Get the single row array
            let rowFormulasUpdated = false;
            const newFormulas = [];
            
            for (let colIndex = 0; colIndex < formulas.length; colIndex++) {
                let formula = formulas[colIndex];
                let originalFormula = formula;
                
                if (typeof formula === 'string' && formula.startsWith('=')) {
                    console.log(`    🔍 Column ${String.fromCharCode(74 + colIndex)} formula: ${formula}`);
                    
                                         // Update SUMIF($3:$3 patterns (year row)
                     if (formula.toLowerCase().includes('sumif($3:$3')) {
                         console.log(`    🎯 Found SUMIF($3:$3 pattern - updating to use year row ${yearRow}`);
                         
                         // Replace SUMIF($3:$3 with SUMIF([yearRowRange]
                         formula = formula.replace(/sumif\(\$3:\$3/gi, `SUMIF(${yearRowRange}`);
                         
                         // Replace ADDRESS(2,COLUMN(),2) with ADDRESS([timeSeriesRow],COLUMN(),2)
                         formula = formula.replace(/ADDRESS\(2,COLUMN\(\),2\)/gi, `ADDRESS(${timeSeriesRow},COLUMN(),2)`);
                         
                         // Update direct column references (K$2, L$2, etc.) to point to time series row
                         const beforeDirectUpdate = formula;
                         formula = formula.replace(/\b([A-Z]{1,3})\$2\b/g, function(match, columnLetter) {
                             return columnLetter + '$' + timeSeriesRow;
                         });
                         
                         if (formula !== beforeDirectUpdate) {
                             console.log(`    🔧 Updated direct column references from row 2 to time series row ${timeSeriesRow}`);
                             console.log(`      Before direct ref update: ${beforeDirectUpdate}`);
                             console.log(`      After direct ref update:  ${formula}`);
                         }
                         
                         console.log(`    ✅ Updated SUMIF($3:$3 formula in column ${String.fromCharCode(74 + colIndex)}`);
                     }
                     
                     // Update SUMIF($4:$4 patterns (year end row)
                     if (formula.toLowerCase().includes('sumif($4:$4')) {
                         console.log(`    🎯 Found SUMIF($4:$4 pattern - updating to use year end row ${yearEndRow}`);
                         
                         // Replace SUMIF($4:$4 with SUMIF([yearEndRowRange]
                         formula = formula.replace(/sumif\(\$4:\$4/gi, `SUMIF(${yearEndRowRange}`);
                         
                         // Replace ADDRESS(2,COLUMN(),2) with ADDRESS([timeSeriesRow],COLUMN(),2)
                         formula = formula.replace(/ADDRESS\(2,COLUMN\(\),2\)/gi, `ADDRESS(${timeSeriesRow},COLUMN(),2)`);
                         
                         // Update direct column references (K$2, L$2, etc.) to point to time series row
                         const beforeDirectUpdate = formula;
                         formula = formula.replace(/\b([A-Z]{1,3})\$2\b/g, function(match, columnLetter) {
                             return columnLetter + '$' + timeSeriesRow;
                         });
                         
                         if (formula !== beforeDirectUpdate) {
                             console.log(`    🔧 Updated direct column references from row 2 to time series row ${timeSeriesRow}`);
                             console.log(`      Before direct ref update: ${beforeDirectUpdate}`);
                             console.log(`      After direct ref update:  ${formula}`);
                         }
                         
                         console.log(`    ✅ Updated SUMIF($4:$4 formula in column ${String.fromCharCode(74 + colIndex)}`);
                     }
                     
                     // Update INDEX/MATCH patterns
                     if (formula.toLowerCase().includes('index(indirect(row()')) {
                         console.log(`    🔍 Found potential INDEX/MATCH pattern, analyzing...`);
                         
                         // First, ensure we have the basic INDEX(INDIRECT(ROW() & ":" & ROW()) structure
                         const indexMatch = formula.match(/INDEX\(INDIRECT\(ROW\(\)\s*&\s*":"\s*&\s*ROW\(\)\)/i);
                         
                         if (indexMatch) {
                             console.log(`    🎯 Found INDEX/MATCH pattern - updating to use time series and year end rows`);
                             console.log(`    📋 Time series row: ${timeSeriesRow}`);
                             console.log(`    📋 Year end row range: ${yearEndRowRange}`);
                             
                             // Replace the entire formula with the new pattern, wrapped in IFERROR
                             const newFormula = `=IFERROR(INDEX(INDIRECT(ROW() & ":" & ROW()),1,MATCH(INDIRECT(ADDRESS(${timeSeriesRow},COLUMN()-1,2)),${yearEndRowRange},0)+1),0)`;
                             
                             // Only update if it's actually different
                             if (newFormula !== formula) {
                                 formula = newFormula;
                                 console.log(`    ✅ Updated INDEX/MATCH formula:`);
                                 console.log(`      Before: ${originalFormula}`);
                                 console.log(`      After:  ${formula}`);
                             } else {
                                 console.log(`    ➡️ Formula already in correct format`);
                             }
                         } else {
                             console.log(`    ➡️ Not a matching INDEX/MATCH pattern`);
                         }
                     }
                    
                    if (formula !== originalFormula) {
                        rowFormulasUpdated = true;
                        console.log(`    🔄 Column ${String.fromCharCode(74 + colIndex)} formula changed`);
                        console.log(`      Before: ${originalFormula}`);
                        console.log(`      After:  ${formula}`);
                    }
                }
                
                newFormulas.push(formula);
            }
            
            // Update the formulas if any changes were made
            if (rowFormulasUpdated) {
                formulaRange.formulas = [newFormulas];
                await worksheet.context.sync();
                formulasUpdated++;
                console.log(`    ✅ Updated formulas synced for row ${currentRow}`);
            } else {
                console.log(`    ➡️ No formula updates needed for row ${currentRow}`);
            }

            // --- Process column U for the current row ---
            console.log(`    🔍 Processing column U formula for row ${currentRow}`);
            const aeCell = worksheet.getRange(`U${currentRow}`);
            aeCell.load("formulas");
            await worksheet.context.sync();

            let aeFormula = aeCell.formulas[0][0];
            if (typeof aeFormula === 'string' && (aeFormula.startsWith('=') || aeFormula.startsWith('@'))) {
                const originalAeFormula = aeFormula;
                const timeSeriesRowRange = `$J$${timeSeriesRow}:$P$${timeSeriesRow}`;
                
                // Define the pattern to find the MATCH part of the formula
                const patternToFind = /MATCH\(@INDIRECT\(ADDRESS\(3,COLUMN\(\),2\)\),\$J\$2:\$P\$2,0\)/gi;

                if (originalAeFormula.match(patternToFind)) {
                    console.log(`    🎯 Found U formula pattern in row ${currentRow}: ${originalAeFormula}`);
                    
                    // Define what to replace it with
                    const replacementPattern = `MATCH(@INDIRECT(ADDRESS(${yearRow},COLUMN(),2)),${timeSeriesRowRange},0)`;

                    // Replace all occurrences in the formula
                    aeFormula = originalAeFormula.replace(patternToFind, replacementPattern);

                    if (aeFormula !== originalAeFormula) {
                        console.log(`    ✅ Updated U formula in row ${currentRow}`);
                        console.log(`      Before: ${originalAeFormula}`);
                        console.log(`      After:  ${aeFormula}`);
                        aeCell.formulas = [[aeFormula]];
                        await worksheet.context.sync();
                        formulasUpdated++; // Increment shared counter
                    }
                }
            } else {
                console.log(`    ➡️ No formula found in U${currentRow}`);
            }
         }
        
        console.log(`🎉 [SUMIF POST-DELETE] Completed updating SUMIF formulas. Updated ${formulasUpdated} rows between ${indexBeginRow} and ${indexEndRow}`);
        
    } catch (error) {
        console.error(`❌ [SUMIF POST-DELETE] Error updating SUMIF formulas: ${error.message}`, error);
        // Continue even if SUMIF updates fail
    }
}

/**
 * Adds missing column labels to row parameters in code strings
 * Ensures all row parameters have the complete sequence: (D), (L), (F), (C1), (C2), (C3), (C4), (C5), (C6), (Y1), (Y2), (Y3), (Y4), (Y5), (Y6)
 * @param {string} inputText - The input text containing code strings
 * @returns {string} - The text with missing column labels added
 */
export function addMissingColumnLabels(inputText) {
    try {
        console.log("=".repeat(80));
        console.log("[AddColumnLabels] STARTING MISSING COLUMN LABEL ADDITION");
        console.log("[AddColumnLabels] Input type:", typeof inputText);
        console.log("[AddColumnLabels] Input length:", inputText?.length || 0);
        
        // Define the expected column label sequence (15 positions total)
        const expectedLabels = ['(D)', '(L)', '(F)', '(C1)', '(C2)', '(C3)', '(C4)', '(C5)', '(C6)', '(Y1)', '(Y2)', '(Y3)', '(Y4)', '(Y5)', '(Y6)'];
        console.log(`[AddColumnLabels] Expected sequence: ${expectedLabels.join(' | ')}`);
        
        // Extract all code strings using regex pattern /<[^>]+>/g
        const codeStringMatches = inputText.match(/<[^>]+>/g);
        if (!codeStringMatches) {
            console.log("[AddColumnLabels] No code strings found in input");
            console.log("=".repeat(80));
            return inputText;
        }
        
        console.log(`[AddColumnLabels] Found ${codeStringMatches.length} code strings to process`);
        
        let modifiedText = inputText;
        let totalModifications = 0;
        
        // Process each code string
        for (const codeString of codeStringMatches) {
            const originalCodeString = codeString;
            let modifiedCodeString = codeString;
            let codeStringModified = false;
            
            // Find all row parameters in this code string
            const rowMatches = codeString.matchAll(/row(\d+)\s*=\s*"([^"]*)"/g);
            
            for (const match of rowMatches) {
                const rowNum = match[1];
                const originalRowValue = match[2];
                const rowParameterFull = match[0]; // The full "row1="..." string
                
                // Clean up extra pipes before processing
                const cleanedRowValue = originalRowValue.replace(/\|\|+/g, '|'); // Replace double+ pipes with single
                
                // Log pipe cleanup if changes were made
                if (originalRowValue !== cleanedRowValue) {
                    console.log(`[AddColumnLabels] Cleaned extra pipes in row${rowNum}:`);
                    console.log(`  Before: "${originalRowValue}"`);
                    console.log(`  After:  "${cleanedRowValue}"`);
                }
                
                // Split the cleaned row value by pipes
                const parts = cleanedRowValue.split('|');
                console.log(`[AddColumnLabels] Processing row${rowNum} with ${parts.length} parts`);
                
                // Create a new parts array with missing labels added
                const newParts = [];
                let partsIndex = 0;
                let addedLabels = [];
                
                for (let expectedIndex = 0; expectedIndex < expectedLabels.length; expectedIndex++) {
                    const expectedLabel = expectedLabels[expectedIndex];
                    const currentPart = partsIndex < parts.length ? parts[partsIndex] : '';
                    
                    // Check if the current part contains the expected label
                    if (currentPart.includes(expectedLabel)) {
                        // Expected label is present, use the current part
                        newParts.push(currentPart);
                        partsIndex++;
                    } else {
                        // Expected label is missing, check if any content at this position should have the label
                        if (currentPart.trim() !== '' && !currentPart.startsWith('(') && !currentPart.endsWith(')')) {
                            // There's content but no label, add the label to the content
                            newParts.push(currentPart + expectedLabel);
                            partsIndex++;
                            addedLabels.push(expectedLabel);
                        } else {
                            // No content at this position, insert just the label
                            newParts.push(expectedLabel);
                            addedLabels.push(expectedLabel);
                            // Don't increment partsIndex since we're inserting, not replacing
                        }
                    }
                }
                
                // Add any remaining parts (shouldn't happen with proper 15-part structure)
                while (partsIndex < parts.length) {
                    newParts.push(parts[partsIndex]);
                    partsIndex++;
                }
                
                // Ensure we have exactly 15 parts (or adjust if needed)
                while (newParts.length < expectedLabels.length) {
                    const missingLabel = expectedLabels[newParts.length];
                    newParts.push(missingLabel);
                    addedLabels.push(missingLabel);
                }
                
                // Add final empty part if needed (to maintain pipe count)
                if (newParts.length === expectedLabels.length) {
                    newParts.push('');
                }
                
                // Create the new row value
                const newRowValue = newParts.join('|');
                
                // Check if modifications were made (either pipe cleanup or label additions)
                if (cleanedRowValue !== newRowValue || originalRowValue !== cleanedRowValue) {
                    console.log(`[AddColumnLabels] Modified row${rowNum}:`);
                    console.log(`  Original: "${originalRowValue}"`);
                    if (originalRowValue !== cleanedRowValue) {
                        console.log(`  Pipe cleaned: "${cleanedRowValue}"`);
                    }
                    console.log(`  Final: "${newRowValue}"`);
                    if (addedLabels.length > 0) {
                        console.log(`  Added labels: ${addedLabels.join(', ')}`);
                    }
                    
                    // Replace the row parameter in the code string
                    const newRowParameter = `row${rowNum}="${newRowValue}"`;
                    modifiedCodeString = modifiedCodeString.replace(rowParameterFull, newRowParameter);
                    codeStringModified = true;
                    totalModifications++;
                }
            }
            
            // Replace the original code string with the modified one in the full text
            if (codeStringModified) {
                modifiedText = modifiedText.replace(originalCodeString, modifiedCodeString);
                console.log(`[AddColumnLabels] Code string modified: ${originalCodeString.substring(0, 50)}...`);
            }
        }
        
        console.log(`[AddColumnLabels] Processing complete. Total modifications: ${totalModifications}`);
        console.log("=".repeat(80));
        
        return modifiedText;
        
    } catch (error) {
        console.error("[AddColumnLabels] Error adding missing column labels:", error);
        console.log("=".repeat(80));
        return inputText; // Return original text if there's an error
    }
}

/**
 * Enhanced version that more intelligently detects and adds missing column labels
 * @param {string} inputText - The input text containing code strings
 * @returns {string} - The text with missing column labels added
 */
export function addMissingColumnLabelsEnhanced(inputText) {
    try {
        console.log("=".repeat(80));
        console.log("[AddColumnLabelsEnhanced] STARTING ENHANCED MISSING COLUMN LABEL ADDITION");
        
        // Define the expected column label sequence (15 positions total)
        const expectedLabels = ['(D)', '(L)', '(F)', '(C1)', '(C2)', '(C3)', '(C4)', '(C5)', '(C6)', '(Y1)', '(Y2)', '(Y3)', '(Y4)', '(Y5)', '(Y6)'];
        
        // Extract all code strings using regex pattern /<[^>]+>/g
        const codeStringMatches = inputText.match(/<[^>]+>/g);
        if (!codeStringMatches) {
            console.log("[AddColumnLabelsEnhanced] No code strings found in input");
            return inputText;
        }
        
        let modifiedText = inputText;
        let totalModifications = 0;
        
        // Process each code string
        for (const codeString of codeStringMatches) {
            let modifiedCodeString = codeString;
            
            // Find all row parameters in this code string
            const rowMatches = codeString.matchAll(/row(\d+)\s*=\s*"([^"]*)"/g);
            
            for (const match of rowMatches) {
                const rowNum = match[1];
                const originalRowValue = match[2];
                const rowParameterFull = match[0];
                
                // Clean up extra pipes before processing
                const cleanedRowValue = originalRowValue.replace(/\|\|+/g, '|'); // Replace double+ pipes with single
                
                // Log pipe cleanup if changes were made
                if (originalRowValue !== cleanedRowValue) {
                    console.log(`[AddColumnLabelsEnhanced] Cleaned extra pipes in row${rowNum}:`);
                    console.log(`  Before: "${originalRowValue}"`);
                    console.log(`  After:  "${cleanedRowValue}"`);
                }
                
                // Split by pipes
                const parts = cleanedRowValue.split('|');
                
                // Check which expected labels are present
                const presentLabels = new Set();
                parts.forEach(part => {
                    expectedLabels.forEach(label => {
                        if (part.includes(label)) {
                            presentLabels.add(label);
                        }
                    });
                });
                
                // Find missing labels
                const missingLabels = expectedLabels.filter(label => !presentLabels.has(label));
                
                if (missingLabels.length > 0) {
                    console.log(`[AddColumnLabelsEnhanced] Row${rowNum} missing labels: ${missingLabels.join(', ')}`);
                    
                    // Reconstruct the parts array with missing labels
                    const newParts = [];
                    
                    for (let i = 0; i < expectedLabels.length; i++) {
                        const expectedLabel = expectedLabels[i];
                        
                        // Find if this label exists in any part
                        const partWithLabel = parts.find(part => part.includes(expectedLabel));
                        
                        if (partWithLabel) {
                            newParts.push(partWithLabel);
                        } else {
                            // Check if there's content at this position without a label
                            if (i < parts.length && parts[i] && !parts[i].startsWith('(')) {
                                // There's content, but it might be meant for this position
                                // Check if this content doesn't belong to a later expected label
                                const hasLaterLabel = expectedLabels.slice(i + 1).some(laterLabel => 
                                    parts[i].includes(laterLabel));
                                
                                if (!hasLaterLabel) {
                                    // This content belongs here, add the label to it
                                    newParts.push(parts[i] + expectedLabel);
                                } else {
                                    // This content belongs to a later position, insert empty label
                                    newParts.push(expectedLabel);
                                }
                            } else {
                                // No content, insert just the label
                                newParts.push(expectedLabel);
                            }
                        }
                    }
                    
                    // Ensure we end with empty string if needed
                    if (newParts.length === expectedLabels.length && !originalRowValue.endsWith('|')) {
                        newParts.push('');
                    }
                    
                    const newRowValue = newParts.join('|');
                    
                    if (cleanedRowValue !== newRowValue || originalRowValue !== cleanedRowValue) {
                        console.log(`[AddColumnLabelsEnhanced] Modified row${rowNum}:`);
                        console.log(`  Original: "${originalRowValue}"`);
                        if (originalRowValue !== cleanedRowValue) {
                            console.log(`  Pipe cleaned: "${cleanedRowValue}"`);
                        }
                        console.log(`  Final: "${newRowValue}"`);
                        
                        const newRowParameter = `row${rowNum}="${newRowValue}"`;
                        modifiedCodeString = modifiedCodeString.replace(rowParameterFull, newRowParameter);
                        totalModifications++;
                    }
                }
            }
            
            // Replace in the full text
            if (modifiedCodeString !== codeString) {
                modifiedText = modifiedText.replace(codeString, modifiedCodeString);
            }
        }
        
        console.log(`[AddColumnLabelsEnhanced] Total modifications: ${totalModifications}`);
        console.log("=".repeat(80));
        
        return modifiedText;
        
    } catch (error) {
        console.error("[AddColumnLabelsEnhanced] Error:", error);
        return inputText;
    }
}

/**
 * Simple wrapper function to add missing column labels and clean up extra pipes after main encoder
 * This function processes the text to ensure all row parameters have complete column label sequences
 * and removes any extra consecutive pipe symbols (|| becomes |)
 * @param {string} encodedText - The text output from the main encoder
 * @returns {string} - The text with missing column labels added and extra pipes cleaned up
 */
export function postProcessColumnLabels(encodedText) {
    console.log("🏷️  [POST-PROCESSOR] Adding missing column labels and cleaning pipes in encoded text...");
    
    try {
        // Use the enhanced version for better accuracy
        const result = addMissingColumnLabelsEnhanced(encodedText);
        console.log("✅ [POST-PROCESSOR] Column label and pipe cleanup post-processing complete");
        return result;
    } catch (error) {
        console.error("❌ [POST-PROCESSOR] Error in column label post-processing:", error);
        console.log("⚠️  [POST-PROCESSOR] Returning original text due to error");
        return encodedText;
    }
}

/**
 * Targeted function that handles the specific missing label scenario like the user's example
 * Focuses on ensuring C1-C6 and Y1-Y6 sequences are complete
 * @param {string} inputText - The input text containing code strings
 * @returns {string} - The text with missing column labels added
 */
export function fixMissingColumnLabels(inputText) {
    try {
        console.log("🔧 [COLUMN LABEL FIX] Starting targeted column label fix");
        
        let modifiedText = inputText;
        let totalFixes = 0;
        
        // Process each code string
        const codeStringPattern = /<[^>]+>/g;
        const codeStrings = inputText.match(codeStringPattern) || [];
        
        for (const codeString of codeStrings) {
            let updatedCodeString = codeString;
            
            // Find row parameters
            const rowPattern = /row(\d+)\s*=\s*"([^"]*)"/g;
            let match;
            
            while ((match = rowPattern.exec(codeString)) !== null) {
                const rowNum = match[1];
                const rowValue = match[2];
                const fullMatch = match[0];
                
                // Clean up extra pipes before processing
                const cleanedRowValue = rowValue.replace(/\|\|+/g, '|'); // Replace double+ pipes with single
                
                // Log pipe cleanup if changes were made
                if (rowValue !== cleanedRowValue) {
                    console.log(`🔧 [COLUMN LABEL FIX] Cleaned extra pipes in row${rowNum}:`);
                    console.log(`   Before: "${rowValue}"`);
                    console.log(`   After:  "${cleanedRowValue}"`);
                }
                
                // Split by pipes
                const parts = cleanedRowValue.split('|');
                
                // Check for missing column labels in the sequence
                const updatedParts = insertMissingLabelsInSequence(parts);
                
                if (updatedParts.join('|') !== cleanedRowValue || rowValue !== cleanedRowValue) {
                    const newRowValue = updatedParts.join('|');
                    const newRowParam = `row${rowNum}="${newRowValue}"`;
                    
                    console.log(`🔧 [COLUMN LABEL FIX] Fixed row${rowNum}:`);
                    console.log(`   Original: "${rowValue}"`);
                    if (rowValue !== cleanedRowValue) {
                        console.log(`   Pipe cleaned: "${cleanedRowValue}"`);
                    }
                    console.log(`   Final: "${newRowValue}"`);
                    
                    updatedCodeString = updatedCodeString.replace(fullMatch, newRowParam);
                    totalFixes++;
                }
            }
            
            // Update the main text
            if (updatedCodeString !== codeString) {
                modifiedText = modifiedText.replace(codeString, updatedCodeString);
            }
        }
        
        console.log(`✅ [COLUMN LABEL FIX] Complete. Total fixes: ${totalFixes}`);
        return modifiedText;
        
    } catch (error) {
        console.error("❌ [COLUMN LABEL FIX] Error:", error);
        return inputText;
    }
}

/**
 * Helper function to insert missing labels in the correct sequence
 * @param {string[]} parts - Array of pipe-separated parts
 * @returns {string[]} - Updated array with missing labels inserted
 */
function insertMissingLabelsInSequence(parts) {
    // Expected sequence: (D), (L), (F), (C1), (C2), (C3), (C4), (C5), (C6), (Y1), (Y2), (Y3), (Y4), (Y5), (Y6)
    const expectedSequence = ['(D)', '(L)', '(F)', '(C1)', '(C2)', '(C3)', '(C4)', '(C5)', '(C6)', '(Y1)', '(Y2)', '(Y3)', '(Y4)', '(Y5)', '(Y6)'];
    
    const result = [];
    let partIndex = 0;
    
    for (let seqIndex = 0; seqIndex < expectedSequence.length; seqIndex++) {
        const expectedLabel = expectedSequence[seqIndex];
        
        if (partIndex < parts.length) {
            const currentPart = parts[partIndex];
            
            // Check if current part has the expected label
            if (currentPart.includes(expectedLabel)) {
                // Label is present, use this part
                result.push(currentPart);
                partIndex++;
            } else {
                // Check if current part has content but wrong/missing label
                if (currentPart.trim() !== '' && !currentPart.match(/\([A-Z]\d*\)/)) {
                    // Content without proper label - this might be meant for this position
                    // But first check if it belongs to a later position
                    const belongsToLaterPosition = expectedSequence.slice(seqIndex + 1).some(laterLabel =>
                        currentPart.includes(laterLabel));
                    
                    if (!belongsToLaterPosition) {
                        // Content belongs here, add the expected label
                        result.push(currentPart + expectedLabel);
                        partIndex++;
                    } else {
                        // Content belongs later, insert missing label
                        result.push(expectedLabel);
                        // Don't increment partIndex
                    }
                } else {
                    // Insert missing label
                    result.push(expectedLabel);
                    // Don't increment partIndex
                }
            }
        } else {
            // No more parts, insert missing label
            result.push(expectedLabel);
        }
    }
    
    // Add any remaining parts (like the final empty string)
    while (partIndex < parts.length) {
        result.push(parts[partIndex]);
        partIndex++;
    }
    
    return result;
}

/**
 * Test function to demonstrate column label fixing with the user's example
 * This shows how the missing (C4) label gets added automatically
 */
export function testColumnLabelFix() {
    console.log("🧪 [TEST] Testing column label fix functionality");
    
    // User's original example - missing (C4)
    const testInput = `<COLUMNHEADER-E; row1="(D)|~Hourly Employees:(L)|(F)|(C1)|(C2)|(C3)|~# of hours per week(C5)|~Hourly rate(C6)|(Y1)|(Y2)|(Y3)|(Y4)|(Y5)|(Y6)|">`;
    
    console.log("📝 [TEST] Input:");
    console.log(`   ${testInput}`);
    
    // Apply the fix
    const result = fixMissingColumnLabels(testInput);
    
    console.log("📝 [TEST] Output:");
    console.log(`   ${result}`);
    
    // Expected output should have (C4) inserted
    const expected = `<COLUMNHEADER-E; row1="(D)|~Hourly Employees:(L)|(F)|(C1)|(C2)|(C3)|(C4)|~# of hours per week(C5)|~Hourly rate(C6)|(Y1)|(Y2)|(Y3)|(Y4)|(Y5)|(Y6)|">`;
    
    console.log("🎯 [TEST] Expected:");
    console.log(`   ${expected}`);
    
    // Check if it matches
    const success = result === expected;
    console.log(`${success ? '✅' : '❌'} [TEST] Test ${success ? 'PASSED' : 'FAILED'}`);
    
    if (!success) {
        console.log("🔍 [TEST] Difference analysis:");
        console.log(`   Result length: ${result.length}`);
        console.log(`   Expected length: ${expected.length}`);
    }
    
    return { input: testInput, result, expected, success };
}

/**
 * Test function to demonstrate comment parsing from square brackets
 * Shows extraction of comments from both row data and formula fields
 */
export function testCommentParsing() {
    console.log("🧪 [TEST] Testing comment parsing from square brackets");
    
    // Test row data with comments
    const testRowData = "~# of Bookings[Per year](L)";
    console.log(`📝 [TEST] Testing row data: "${testRowData}"`);
    
    const rowResult = parseCommentFromBrackets(testRowData);
    console.log(`📝 [TEST] Row parsing result:`);
    console.log(`   Clean value: "${rowResult.cleanValue}"`);
    console.log(`   Comment: "${rowResult.comment}"`);
    
    // Test customformula with comments
    const testFormula = "rd{V1}*rd{V2}*30[Monthly calculation]";
    console.log(`📝 [TEST] Testing formula: "${testFormula}"`);
    
    const formulaResult = parseCommentFromBrackets(testFormula);
    console.log(`📝 [TEST] Formula parsing result:`);
    console.log(`   Clean value: "${formulaResult.cleanValue}"`);
    console.log(`   Comment: "${formulaResult.comment}"`);
    
    // Test value without comments
    const testNoComment = "~$120000(C5)";
    console.log(`📝 [TEST] Testing value without comments: "${testNoComment}"`);
    
    const noCommentResult = parseCommentFromBrackets(testNoComment);
    console.log(`📝 [TEST] No comment parsing result:`);
    console.log(`   Clean value: "${noCommentResult.cleanValue}"`);
    console.log(`   Comment: "${noCommentResult.comment}"`);
    
    // Expected results validation
    const tests = [
        {
            input: testRowData,
            expectedClean: "~# of Bookings(L)",
            expectedComment: "Per year",
            result: rowResult
        },
        {
            input: testFormula,
            expectedClean: "rd{V1}*rd{V2}*30",
            expectedComment: "Monthly calculation",
            result: formulaResult
        },
        {
            input: testNoComment,
            expectedClean: "~$120000(C5)",
            expectedComment: null,
            result: noCommentResult
        }
    ];
    
    let allPassed = true;
    tests.forEach((test, index) => {
        const cleanMatch = test.result.cleanValue === test.expectedClean;
        const commentMatch = test.result.comment === test.expectedComment;
        const passed = cleanMatch && commentMatch;
        
        console.log(`${passed ? '✅' : '❌'} [TEST] Test ${index + 1}: ${passed ? 'PASSED' : 'FAILED'}`);
        if (!passed) {
            console.log(`   Expected clean: "${test.expectedClean}", got: "${test.result.cleanValue}"`);
            console.log(`   Expected comment: "${test.expectedComment}", got: "${test.result.comment}"`);
            allPassed = false;
        }
    });
    
    console.log(`${allPassed ? '✅' : '❌'} [TEST] Overall comment parsing test: ${allPassed ? 'PASSED' : 'FAILED'}`);
    
    return { tests, allPassed };
}

/**
 * Test function to demonstrate both missing column labels and extra pipe cleanup
 * This shows the new functionality of removing extra pipes
 */
export function testColumnLabelAndPipeCleanup() {
    console.log("🧪 [TEST] Testing column label fix and pipe cleanup functionality");
    
    // Test case with both missing labels and extra pipes
    const testInput = `<FORMULA-S; row1="V1(D)|Revenue(L)||(C1)||(C2)||(C3)||~$120000(C5)|~Hourly rate(C6)|(Y1)|(Y2)|(Y3)|(Y4)|(Y5)|(Y6)|">`;
    
    console.log("📝 [TEST] Input (missing C4 and extra pipes):");
    console.log(`   ${testInput}`);
    
    // Apply the fix
    const result = fixMissingColumnLabels(testInput);
    
    console.log("📝 [TEST] Output:");
    console.log(`   ${result}`);
    
    // Expected output should have extra pipes removed and (C4) inserted
    const expected = `<FORMULA-S; row1="V1(D)|Revenue(L)|(F)|(C1)|(C2)|(C3)|(C4)|~$120000(C5)|~Hourly rate(C6)|(Y1)|(Y2)|(Y3)|(Y4)|(Y5)|(Y6)|">`;
    
    console.log("🎯 [TEST] Expected (pipes cleaned and C4 added):");
    console.log(`   ${expected}`);
    
    // Check if it matches
    const success = result === expected;
    console.log(`${success ? '✅' : '❌'} [TEST] Pipe cleanup test ${success ? 'PASSED' : 'FAILED'}`);
    
    if (!success) {
        console.log("🔍 [TEST] Difference analysis:");
        console.log(`   Result length: ${result.length}`);
        console.log(`   Expected length: ${expected.length}`);
    }
    
    return { input: testInput, result, expected, success };
}

/**
 * Inserts the entire model codes into cell A1 of the Financials tab
 * This preserves the original code strings that were used to create the model
 * @param {string} modelCodes - The complete model codes string that was used to build the model
 * @returns {Promise<void>}
 */
export async function insertModelCodesIntoFinancials(modelCodes) {
    try {
        console.log("📋 [MODEL CODES] Inserting model codes into Financials!A1");
        console.log(`📋 [MODEL CODES] Code string length: ${modelCodes?.length || 0} characters`);
        
        if (!modelCodes || typeof modelCodes !== 'string') {
            console.warn("⚠️ [MODEL CODES] Invalid or empty model codes provided, skipping insertion");
            return;
        }
        
        await Excel.run(async (context) => {
            try {
                // Get the Financials worksheet
                const financialsSheet = context.workbook.worksheets.getItem("Financials");
                console.log("📋 [MODEL CODES] Successfully got Financials worksheet reference");
                
                // Get cell A1
                const cellA1 = financialsSheet.getRange("A1");
                
                // Insert the model codes as a value
                cellA1.values = [[modelCodes]];
                console.log("📋 [MODEL CODES] Model codes inserted into Financials!A1");
                
                // Apply formatting to make it easier to identify
                cellA1.format.font.color = "#FFFFFF"; // White font (hidden)
                cellA1.format.font.size = 8; // Small font size
                cellA1.format.wrapText = false; // Don't wrap text to keep it contained
                console.log("📋 [MODEL CODES] Applied formatting (white font, size 8, no wrap)");
                
                // Sync the changes
                await context.sync();
                console.log("✅ [MODEL CODES] Successfully inserted and formatted model codes in Financials!A1");
                
            } catch (error) {
                console.error("❌ [MODEL CODES] Error inserting model codes into Financials!A1:", error);
                throw error;
            }
        });
        
    } catch (error) {
        console.error("❌ [MODEL CODES] Critical error in insertModelCodesIntoFinancials:", error);
        // Don't throw the error to avoid breaking the model completion process
        console.log("⚠️ [MODEL CODES] Continuing despite error to avoid breaking model completion");
    }
}

/**
 * Enhanced version of hideColumnsAndNavigate that also inserts model codes
 * @param {string[]} assumptionTabNames - Array of assumption tab names created by runCodes
 * @param {string} originalModelCodes - The original model codes string used to build the model
 * @returns {Promise<void>}
 */
export async function hideColumnsAndNavigateAndInsertCodes(assumptionTabNames, originalModelCodes) {
    try {
        console.log("🎯 [MODEL COMPLETION] Starting final model completion steps...");
        
        // First, run the standard hide columns and navigate function
        await hideColumnsAndNavigate(assumptionTabNames);
        console.log("✅ [MODEL COMPLETION] Completed hideColumnsAndNavigate");
        
        // Then insert the model codes into Financials!A1
        if (originalModelCodes) {
            console.log("📋 [MODEL COMPLETION] Inserting original model codes into Financials...");
            await insertModelCodesIntoFinancials(originalModelCodes);
            console.log("✅ [MODEL COMPLETION] Model codes insertion completed");
        } else {
            console.warn("⚠️ [MODEL COMPLETION] No original model codes provided, skipping insertion");
        }
        
        console.log("🎉 [MODEL COMPLETION] All model completion steps finished successfully");
        
    } catch (error) {
        console.error("❌ [MODEL COMPLETION] Error in hideColumnsNavigateAndInsertCodes:", error);
        throw error;
    }
}

/**
 * Retrieves the model codes from cell A1 of the Financials tab
 * Useful for debugging or extracting the codes that were used to build the model
 * @returns {Promise<string|null>} - The model codes string or null if not found/error
 */
export async function getModelCodesFromFinancials() {
    try {
        console.log("📋 [MODEL CODES RETRIEVAL] Retrieving model codes from Financials!A1");
        
        return await Excel.run(async (context) => {
            try {
                // Get the Financials worksheet
                const financialsSheet = context.workbook.worksheets.getItem("Financials");
                
                // Get cell A1
                const cellA1 = financialsSheet.getRange("A1");
                cellA1.load("values");
                
                // Sync to load the value
                await context.sync();
                
                const modelCodes = cellA1.values[0][0];
                
                if (modelCodes && typeof modelCodes === 'string' && modelCodes.trim() !== '') {
                    console.log(`✅ [MODEL CODES RETRIEVAL] Retrieved model codes (${modelCodes.length} characters)`);
                    return modelCodes;
                } else {
                    console.log("⚠️ [MODEL CODES RETRIEVAL] No model codes found in Financials!A1");
                    return null;
                }
                
            } catch (error) {
                console.error("❌ [MODEL CODES RETRIEVAL] Error retrieving model codes:", error);
                return null;
            }
        });
        
    } catch (error) {
        console.error("❌ [MODEL CODES RETRIEVAL] Critical error in getModelCodesFromFinancials:", error);
        return null;
    }
}

/**
 * Test function for TABLEMIN and SUMTABLE conversion
 */
export async function testTableminConversion() {
    console.log("🧪 Testing TABLEMIN and SUMTABLE conversion...");
    
    // Test TABLEMIN - the exact case from the user
    const testFormula = "TABLEMIN(rd{V1},cd{6-V5})";
    console.log(`Original TABLEMIN formula: ${testFormula}`);
    
    // Test without worksheet context first (rd{} and cd{} won't be converted)
    const result1 = await parseFormulaSCustomFormula(testFormula, 15);
    console.log(`Result without worksheet context: ${result1}`);
    
    // Test with simplified formula where references are already resolved
    const testFormula2 = "TABLEMIN(U$15,$I15)";
    console.log(`\nSimplified TABLEMIN formula: ${testFormula2}`);
    
    const result2 = await parseFormulaSCustomFormula(testFormula2, 15);
    console.log(`Result with resolved references: ${result2}`);
    
    // Test SUMTABLE
    console.log("\n🧪 Testing SUMTABLE function:");
    const sumtableFormula1 = "SUMTABLE(rd{V1})";
    console.log(`Original SUMTABLE formula: ${sumtableFormula1}`);
    
    const sumtableResult1 = await parseFormulaSCustomFormula(sumtableFormula1, 15);
    console.log(`Result without worksheet context: ${sumtableResult1}`);
    
    const sumtableFormula2 = "SUMTABLE(U$15)";
    console.log(`\nSimplified SUMTABLE formula: ${sumtableFormula2}`);
    
    const sumtableResult2 = await parseFormulaSCustomFormula(sumtableFormula2, 15);
    console.log(`Result with resolved references: ${sumtableResult2}`);
    
    // Test RANGE function
    console.log("\n🧪 Testing RANGE function:");
    const rangeFormula1 = "RANGE(rd{V1})";
    console.log(`Original RANGE formula: ${rangeFormula1}`);
    
    const rangeResult1 = await parseFormulaSCustomFormula(rangeFormula1, 15);
    console.log(`Result without worksheet context: ${rangeResult1}`);
    
    const rangeFormula2 = "RANGE(U$3)";
    console.log(`\nSimplified RANGE formula: ${rangeFormula2}`);
    
    const rangeResult2 = await parseFormulaSCustomFormula(rangeFormula2, 15);
    console.log(`Result with resolved references: ${rangeResult2}`);
    
    const rangeFormula3 = "RANGE(V4)";
    console.log(`\nAnother RANGE formula: ${rangeFormula3}`);
    
    const rangeResult3 = await parseFormulaSCustomFormula(rangeFormula3, 15);
    console.log(`Result: ${rangeResult3}`);
    
    // Test edge cases
    const testCases = [
        "TABLEMIN(A1,B1)",
        "MIN(TABLEMIN(A1,B1),C1)",
        "TABLEMIN(rd{Driver1},cd{5})",
        "=TABLEMIN(U$15,$I15)",
        "SUMTABLE(A1)",
        "SUMTABLE(rd{Driver1})",
        "=SUMTABLE(U$15)",
        "SUMTABLE(A1)+TABLEMIN(B1,C1)",
        "RANGE(U$5)",
        "RANGE(V3)",
        "SUM(RANGE(U$10))",
        "=RANGE(rd{V1})",
        "RANGE(U15)",
        "RANGE(A$7)"
    ];
    
    console.log("\n🧪 Testing additional cases:");
    for (const testCase of testCases) {
        console.log(`Input: ${testCase}`);
        const result = await parseFormulaSCustomFormula(testCase, 15);
        console.log(`Output: ${result}\n`);
    }
    
    console.log("✅ TABLEMIN, SUMTABLE, and RANGE conversion test complete");
}

/**
 * Test function to demonstrate COLUMNFORMULA-S functionality  
 * @returns {Promise<void>}
 */
export async function testColumnFormulaSConversion() {
    console.log("=".repeat(80));
    console.log("🧪 TESTING COLUMNFORMULA-S FUNCTIONALITY");
    console.log("=".repeat(80));
    
    try {
        const testInput = `<COLUMNFORMULA-S; customformula="rd{V1}*cd{6-V2}/cd{6-V3}"; row1="A1(D)|Test Formula(L)|is: revenue(F)|(C1)|(C2)|(C3)|(C4)|(C5)|(C6)|(Y1)|(Y2)|(Y3)|(Y4)|(Y5)|(Y6)|">`;
        
        console.log("📋 Test Input:");
        console.log(testInput);
        
        // Parse the code string
        const codeCollection = populateCodeCollection(testInput);
        console.log("✅ Parsed code collection:");
        console.log(JSON.stringify(codeCollection, null, 2));
        
        // Verify the code type and customformula parameter
        if (codeCollection.length > 0) {
            const code = codeCollection[0];
            console.log(`📊 Code type: ${code.type}`);
            console.log(`📊 Custom formula: ${code.params.customformula}`);
            
            if (code.type === "COLUMNFORMULA-S" && code.params.customformula) {
                console.log("✅ COLUMNFORMULA-S code parsed correctly");
                console.log("✅ Custom formula parameter extracted successfully");
                
                // Test formula processing (simulate what would happen in column I)
                const testFormula = code.params.customformula;
                console.log(`🔧 Testing formula processing on: "${testFormula}"`);
                
                // Test the formula parsing with parseFormulaSCustomFormula
                const processedFormula = await parseFormulaSCustomFormula(testFormula, 10);
                console.log(`🎯 Processed formula result: "${processedFormula}"`);
                console.log("🎯 This formula would be placed in column I during code execution");
                console.log("🎯 processColumnFormulaSRows would convert it to Excel formula");
                
                // Show comparison
                console.log("\n📊 COLUMNFORMULA-S vs FORMULA-S Comparison:");
                console.log("   FORMULA-S: customformula -> column U");
                console.log("   COLUMNFORMULA-S: customformula -> column I");
                console.log("   Both use the same cd{} and rd{} syntax");
                console.log("   Both are processed by parseFormulaSCustomFormula");
            } else {
                console.log("❌ COLUMNFORMULA-S parsing failed");
            }
        }
        
        console.log("=".repeat(80));
        console.log("🎉 COLUMNFORMULA-S test completed successfully");
        console.log("=".repeat(80));
        
    } catch (error) {
        console.error("❌ Error in testColumnFormulaSConversion:", error.message);
        console.log("=".repeat(80));
    }
}

/**
 * Test function to demonstrate RANGE custom function functionality
 * @returns {Promise<void>}
 */
export async function testRangeFunction() {
    console.log("=".repeat(80));
    console.log("🧪 TESTING RANGE CUSTOM FUNCTION");
    console.log("=".repeat(80));
    
    try {
        console.log("📋 RANGE Function Purpose:");
        console.log("   Converts RANGE(driver1) to U{row}:CN{row}");
        console.log("   Where {row} is extracted from the driver1 cell reference");
        console.log("   Examples: RANGE(rd{V1}) where V1 is row 3 → U3:CN3");
        console.log("");
        
        // Test cases for RANGE function
        const testCases = [
            {
                name: "Basic cell reference",
                formula: "RANGE(U$3)",
                expected: "U3:CN3"
            },
            {
                name: "Cell reference without dollar sign",
                formula: "RANGE(V4)",
                expected: "U4:CN4"
            },
            {
                name: "Mixed reference",
                formula: "RANGE(A$15)",
                expected: "U15:CN15"
            },
            {
                name: "Within SUM function",
                formula: "SUM(RANGE(U$10))",
                expected: "SUM(U10:CN10)"
            },
            {
                name: "With rd{} syntax (won't resolve without worksheet context)",
                formula: "RANGE(rd{V1})",
                expected: "RANGE(rd{V1})" // Will remain unchanged without worksheet context
            },
            {
                name: "Multiple RANGE functions",
                formula: "RANGE(U$5)+RANGE(V6)",
                expected: "U5:CN5+U6:CN6"
            },
            {
                name: "As Excel formula",
                formula: "=RANGE(U$12)",
                expected: "=U12:CN12"
            }
        ];
        
        console.log("🧪 Testing RANGE function conversions:");
        console.log("-".repeat(60));
        
        for (const testCase of testCases) {
            console.log(`\n📝 Test: ${testCase.name}`);
            console.log(`   Input:    "${testCase.formula}"`);
            console.log(`   Expected: "${testCase.expected}"`);
            
            const result = await parseFormulaSCustomFormula(testCase.formula, 15);
            console.log(`   Actual:   "${result}"`);
            
            const success = result === testCase.expected;
            console.log(`   ${success ? '✅ PASS' : '❌ FAIL'}`);
            
            if (!success) {
                console.log(`   ⚠️  Expected "${testCase.expected}" but got "${result}"`);
            }
        }
        
        console.log("\n" + "=".repeat(80));
        console.log("📊 RANGE Function Usage Examples:");
        console.log("   • RANGE(U$3) → U3:CN3 (columns U through CN of row 3)");
        console.log("   • SUM(RANGE(U$5)) → SUM(U5:CN5) (sum all values in row 5 from U to CN)");
        console.log("   • AVERAGE(RANGE(V10)) → AVERAGE(U10:CN10) (average of row 10)");
        console.log("   • RANGE(rd{V1}) → Depends on what row V1 refers to");
        console.log("");
        console.log("✅ Useful for operations across entire time series (columns U:CN)");
        console.log("✅ Works with any Excel function that accepts ranges");
        console.log("✅ Supports both absolute ($) and relative references");
        console.log("=".repeat(80));
        
    } catch (error) {
        console.error("❌ Error in testRangeFunction:", error.message);
        console.log("=".repeat(80));
    }
}