// PipeValidation.js - Automatically corrects pipe count in row parameters
// This sits outside the normal validation process and corrects codestrings automatically

/**
 * Extracts all row parameters from a codestring
 * @param {string} inputString - The codestring to extract from
 * @returns {Object} - Object with rowX as keys and quoted content as values
 */
function extractRowParameters(inputString) {
    const rowParams = {};
    
    if (!inputString || inputString.length === 0) {
        return rowParams;
    }
    
    // Look for all row parameters (row1, row2, etc.)
    for (let i = 1; i <= 999; i++) {
        const rowPattern = `row${i}`;
        const regex = new RegExp(`${rowPattern}\\s*=\\s*"([^"]*)"`, 'g');
        const matches = [...inputString.matchAll(regex)];
        
        matches.forEach((match, index) => {
            const key = matches.length > 1 ? `${rowPattern}_${index + 1}` : rowPattern;
            rowParams[key] = match[1]; // The content between quotes
        });
    }
    
    return rowParams;
}

/**
 * Checks if a string contains row parameters (row1, row2, etc.)
 * @param {string} inputString - The string to check
 * @returns {boolean} - True if contains row parameters
 */
function containsRowParameters(inputString) {
    if (!inputString || inputString.length === 0) {
        console.log("containsRowParameters: Empty or null input string");
        return false;
    }
    
    // Look for pattern like "row1 = " or "row2=" etc.
    if (inputString.includes('row') && inputString.includes('=') && inputString.includes('"')) {
        // Additional check to ensure it's actually a row parameter
        for (let i = 1; i <= 999; i++) {
            if (inputString.includes(`row${i}`)) {
                console.log(`containsRowParameters: Found row parameter row${i} in string: ${inputString.substring(0, 100)}...`);
                return true;
            }
        }
    }
    
    console.log(`containsRowParameters: No row parameters found in: ${inputString.substring(0, 100)}...`);
    return false;
}

/**
 * Fixes the pipe count in a single row parameter to exactly 15 pipes
 * @param {string} inputString - The string containing a row parameter
 * @returns {string} - The corrected string
 */
function fixPipeCountTo15(inputString) {
    console.log(`fixPipeCountTo15: Processing input: ${inputString}`);
    
    // Find the quoted section
    const startQuotePos = inputString.indexOf('"');
    if (startQuotePos === -1) {
        console.log("fixPipeCountTo15: No opening quote found, returning original");
        return inputString;
    }
    
    const endQuotePos = inputString.indexOf('"', startQuotePos + 1);
    if (endQuotePos === -1) {
        console.log("fixPipeCountTo15: No closing quote found, returning original");
        return inputString;
    }
    
    // Extract the content between quotes
    const quotedContent = inputString.substring(startQuotePos + 1, endQuotePos);
    console.log(`fixPipeCountTo15: Quoted content: "${quotedContent}"`);
    
    // Count pipes in the quoted content
    const pipeCount = (quotedContent.match(/\|/g) || []).length;
    console.log(`fixPipeCountTo15: Found ${pipeCount} pipes (target: 15)`);
    
    // If already exactly 15 pipes, no change needed
    if (pipeCount === 15) {
        console.log("fixPipeCountTo15: Already has exactly 15 pipes, no changes needed");
        return inputString;
    }
    
    // Find the position after the 3rd pipe
    let pipeFoundCount = 0;
    let insertPosition = -1;
    
    for (let j = 0; j < quotedContent.length; j++) {
        if (quotedContent[j] === '|') {
            pipeFoundCount++;
            if (pipeFoundCount === 3) {
                insertPosition = j;
                console.log(`fixPipeCountTo15: Found 3rd pipe at position ${j}`);
                break;
            }
        }
    }
    
    // If we couldn't find the 3rd pipe, return original
    if (insertPosition === -1) {
        console.log("fixPipeCountTo15: Could not find 3rd pipe, returning original");
        return inputString;
    }
    
    let newQuotedContent;
    
    if (pipeCount < 15) {
        // Add pipes after the 3rd pipe
        const pipesToAdd = 15 - pipeCount;
        console.log(`fixPipeCountTo15: Adding ${pipesToAdd} pipes after 3rd pipe`);
        const additionalPipes = '|'.repeat(pipesToAdd);
        
        newQuotedContent = quotedContent.substring(0, insertPosition + 1) + 
                          additionalPipes + 
                          quotedContent.substring(insertPosition + 1);
                          
    } else if (pipeCount > 15) {
        // Remove pipes after the 3rd pipe
        const pipesToRemove = pipeCount - 15;
        console.log(`fixPipeCountTo15: Removing ${pipesToRemove} pipes after 3rd pipe`);
        let afterThirdPipe = quotedContent.substring(insertPosition + 1);
        
        // Remove pipes from the beginning of afterThirdPipe
        let pipesRemoved = 0;
        let k = 0;
        
        while (k < afterThirdPipe.length && pipesRemoved < pipesToRemove) {
            if (afterThirdPipe[k] === '|') {
                pipesRemoved++;
                k++;
            } else {
                break;
            }
        }
        
        console.log(`fixPipeCountTo15: Actually removed ${pipesRemoved} pipes`);
        afterThirdPipe = afterThirdPipe.substring(k);
        newQuotedContent = quotedContent.substring(0, insertPosition + 1) + afterThirdPipe;
    }
    
    // Reconstruct the full string
    const beforeQuote = inputString.substring(0, startQuotePos + 1);
    const afterQuote = inputString.substring(endQuotePos);
    const result = beforeQuote + newQuotedContent + afterQuote;
    
    console.log(`fixPipeCountTo15: New quoted content: "${newQuotedContent}"`);
    console.log(`fixPipeCountTo15: Final result: ${result}`);
    
    // Verify the pipe count in the result
    const finalPipeCount = (newQuotedContent.match(/\|/g) || []).length;
    console.log(`fixPipeCountTo15: Verification - Final pipe count: ${finalPipeCount}`);
    
    return result;
}

/**
 * Fixes all row parameter pipes in the input string
 * @param {string} inputString - The string containing codestrings with row parameters
 * @returns {string} - The corrected string
 */
function fixAllRowParameterPipes(inputString) {
    console.log(`fixAllRowParameterPipes: Starting to process input: ${inputString.substring(0, 200)}...`);
    let result = inputString;
    let totalChanges = 0;
    
    // Find and fix each row parameter (row1, row2, row3, etc.)
    for (let i = 1; i <= 999; i++) {
        const rowPattern = `row${i}`;
        let searchPos = 0;
        let instancesFound = 0;
        
        while (true) {
            // Find the next instance of this row parameter
            const startPos = result.indexOf(rowPattern, searchPos);
            
            // If no more instances found, move to next row number
            if (startPos === -1) break;
            
            console.log(`fixAllRowParameterPipes: Found ${rowPattern} at position ${startPos}`);
            
            // Make sure this is actually a row parameter (has = and quotes after it)
            const checkStart = startPos + rowPattern.length;
            const hasEquals = result.indexOf('=', checkStart) > checkStart && 
                             result.indexOf('=', checkStart) < result.indexOf('row', checkStart + 1) || 
                             result.indexOf('row', checkStart + 1) === -1;
            const hasQuotes = result.indexOf('"', checkStart) > checkStart && 
                             result.indexOf('"', checkStart) < result.indexOf('row', checkStart + 1) || 
                             result.indexOf('row', checkStart + 1) === -1;
            
            if (!(hasEquals && hasQuotes)) {
                // Not a valid row parameter, continue searching after this position
                console.log(`fixAllRowParameterPipes: ${rowPattern} at position ${startPos} is not a valid row parameter (missing = or quotes)`);
                searchPos = startPos + rowPattern.length;
                continue;
            }
            
            instancesFound++;
            console.log(`fixAllRowParameterPipes: Valid ${rowPattern} instance #${instancesFound} found`);
            
            // Find the end of this row parameter instance
            let endPos;
            let nextRowPos = result.length;
            
            // Look for the next row parameter of ANY number
            for (let j = 1; j <= 999; j++) {
                const nextPattern = `row${j}`;
                const tempPos = result.indexOf(nextPattern, checkStart);
                if (tempPos > 0 && tempPos < nextRowPos) {
                    // Make sure this next row parameter is valid too
                    const nextCheckStart = tempPos + nextPattern.length;
                    const nextHasEquals = result.indexOf('=', nextCheckStart) > nextCheckStart;
                    const nextHasQuotes = result.indexOf('"', nextCheckStart) > nextCheckStart;
                    if (nextHasEquals && nextHasQuotes) {
                        nextRowPos = tempPos;
                    }
                }
            }
            
            endPos = nextRowPos - 1;
            
            // Extract this single row parameter instance
            const singleRowParam = result.substring(startPos, endPos + 1);
            console.log(`fixAllRowParameterPipes: Extracted ${rowPattern} instance: ${singleRowParam}`);
            
            // Fix the pipes in this single row parameter
            const fixedRowParam = fixPipeCountTo15(singleRowParam);
            
            // Replace the original with the fixed version
            if (fixedRowParam !== singleRowParam) {
                console.log(`fixAllRowParameterPipes: ${rowPattern} instance was modified`);
                console.log(`fixAllRowParameterPipes: Before: ${singleRowParam}`);
                console.log(`fixAllRowParameterPipes: After:  ${fixedRowParam}`);
                
                const lengthDifference = fixedRowParam.length - singleRowParam.length;
                result = result.substring(0, startPos) + fixedRowParam + result.substring(endPos + 1);
                // Adjust end position for the length change
                endPos = endPos + lengthDifference;
                totalChanges++;
            } else {
                console.log(`fixAllRowParameterPipes: ${rowPattern} instance required no changes`);
            }
            
            // Continue searching after this instance
            searchPos = endPos + 1;
        }
        
        if (instancesFound > 0) {
            console.log(`fixAllRowParameterPipes: Processed ${instancesFound} instances of ${rowPattern}`);
        }
    }
    
    console.log(`fixAllRowParameterPipes: Completed processing. Total changes made: ${totalChanges}`);
    if (totalChanges > 0) {
        console.log(`fixAllRowParameterPipes: Final result: ${result.substring(0, 200)}...`);
    }
    
    return result;
}

/**
 * Main function to automatically correct pipe counts in codestrings
 * @param {string|Array} inputText - Input text or array of codestrings
 * @returns {Object} - Object containing corrected text and change count
 */
export function autoCorrectPipeCounts(inputText) {
    console.log("=== PIPE CORRECTION PROCESS STARTED ===");
    console.log(`autoCorrectPipeCounts: Input type: ${typeof inputText}`);
    console.log(`autoCorrectPipeCounts: Input preview: ${typeof inputText === 'string' ? inputText.substring(0, 200) + '...' : `Array with ${inputText.length} elements`}`);
    
    let inputCodeStrings = [];
    let changesMade = 0;
    
    // Extract all code strings using regex pattern /<[^>]+>/g
    if (typeof inputText === 'string') {
        const codeStringMatches = inputText.match(/<[^>]+>/g);
        inputCodeStrings = codeStringMatches || [];
        console.log(`autoCorrectPipeCounts: Extracted ${inputCodeStrings.length} code strings from string input`);
    } else if (Array.isArray(inputText)) {
        // If already an array, extract code strings from each element
        inputCodeStrings = [];
        for (const item of inputText) {
            if (typeof item === 'string') {
                const matches = item.match(/<[^>]+>/g);
                if (matches) {
                    console.log(`autoCorrectPipeCounts: Found ${matches.length} code strings in array element: ${item.substring(0, 100)}...`);
                    inputCodeStrings.push(...matches);
                }
            }
        }
        console.log(`autoCorrectPipeCounts: Total code strings extracted from array: ${inputCodeStrings.length}`);
    }
    
    if (inputCodeStrings.length === 0) {
        console.log("autoCorrectPipeCounts: No code strings found in input for pipe correction");
        console.log("=== PIPE CORRECTION PROCESS COMPLETED (NO WORK) ===");
        return {
            correctedText: inputText,
            changesMade: 0,
            details: []
        };
    }
    
    console.log(`autoCorrectPipeCounts: Found ${inputCodeStrings.length} code strings for pipe correction`);
    
    // Remove line breaks within angle brackets before processing
    inputCodeStrings = inputCodeStrings.map((str, index) => {
        const cleaned = str.replace(/<[^>]*>/g, match => {
            return match.replace(/[\r\n]+/g, ' ');
        });
        if (cleaned !== str) {
            console.log(`autoCorrectPipeCounts: Cleaned line breaks from code string ${index + 1}`);
        }
        return cleaned;
    });
    
    // Clean up input strings by removing anything outside angle brackets
    inputCodeStrings = inputCodeStrings.map((str, index) => {
        const match = str.match(/<[^>]+>/);
        const result = match ? match[0] : str;
        if (result !== str) {
            console.log(`autoCorrectPipeCounts: Cleaned up code string ${index + 1}: ${str} -> ${result}`);
        }
        return result;
    });
    
    let correctedCodeStrings = [];
    let correctionDetails = [];
    
    console.log("autoCorrectPipeCounts: Starting to process individual code strings...");
    
    // Process each codestring
    for (let i = 0; i < inputCodeStrings.length; i++) {
        const codeString = inputCodeStrings[i];
        console.log(`\nautoCorrectPipeCounts: Processing code string ${i + 1}/${inputCodeStrings.length}`);
        
        if (containsRowParameters(codeString)) {
            console.log(`autoCorrectPipeCounts: Code string ${i + 1} contains row parameters, applying corrections...`);
            const correctedCodeString = fixAllRowParameterPipes(codeString);
            
            if (correctedCodeString !== codeString) {
                changesMade++;
                correctionDetails.push({
                    original: codeString,
                    corrected: correctedCodeString,
                    changed: true
                });
                console.log(`autoCorrectPipeCounts: ✓ Code string ${i + 1} was MODIFIED`);
                
                // Extract and log the specific row parameters that changed
                console.log(`autoCorrectPipeCounts: 📋 DETAILED ROW PARAMETER CHANGES:`);
                const originalRows = extractRowParameters(codeString);
                const correctedRows = extractRowParameters(correctedCodeString);
                
                // Compare each row parameter
                Object.keys(originalRows).forEach(rowKey => {
                    if (originalRows[rowKey] !== correctedRows[rowKey]) {
                        console.log(`autoCorrectPipeCounts: 🔄 ${rowKey} CHANGED:`);
                        console.log(`autoCorrectPipeCounts:   BEFORE: ${rowKey} = "${originalRows[rowKey]}"`);
                        console.log(`autoCorrectPipeCounts:   AFTER:  ${rowKey} = "${correctedRows[rowKey]}"`);
                        
                        // Count pipes in before and after
                        const beforePipes = (originalRows[rowKey].match(/\|/g) || []).length;
                        const afterPipes = (correctedRows[rowKey].match(/\|/g) || []).length;
                        console.log(`autoCorrectPipeCounts:   PIPES:  ${beforePipes} → ${afterPipes} (${afterPipes - beforePipes >= 0 ? '+' : ''}${afterPipes - beforePipes})`);
                    }
                });
                
                console.log(`autoCorrectPipeCounts: 📊 SUMMARY: ${codeString.substring(0, 100)}... → ${correctedCodeString.substring(0, 100)}...`);
            } else {
                correctionDetails.push({
                    original: codeString,
                    corrected: correctedCodeString,
                    changed: false
                });
                console.log(`autoCorrectPipeCounts: ✓ Code string ${i + 1} required NO CHANGES`);
            }
            
            correctedCodeStrings.push(correctedCodeString);
        } else {
            console.log(`autoCorrectPipeCounts: Code string ${i + 1} has no row parameters, skipping...`);
            correctedCodeStrings.push(codeString);
            correctionDetails.push({
                original: codeString,
                corrected: codeString,
                changed: false
            });
        }
    }
    
    console.log(`\nautoCorrectPipeCounts: Individual processing complete. Building final result...`);
    
    // Reconstruct the corrected text
    let correctedText = inputText;
    let replacements = 0;
    
    if (typeof inputText === 'string') {
        // Replace original codestrings with corrected ones in the original text
        for (let i = 0; i < inputCodeStrings.length; i++) {
            if (correctionDetails[i].changed) {
                correctedText = correctedText.replace(
                    correctionDetails[i].original, 
                    correctionDetails[i].corrected
                );
                replacements++;
            }
        }
        console.log(`autoCorrectPipeCounts: Made ${replacements} text replacements in original string`);
    } else if (Array.isArray(inputText)) {
        correctedText = correctedCodeStrings;
        console.log(`autoCorrectPipeCounts: Returned corrected array with ${correctedCodeStrings.length} elements`);
    }
    
    // Summary logging
    console.log(`\n=== PIPE CORRECTION SUMMARY ===`);
    console.log(`Total code strings processed: ${inputCodeStrings.length}`);
    console.log(`Code strings with row parameters: ${correctionDetails.filter(d => d.original !== d.corrected || containsRowParameters(d.original)).length}`);
    console.log(`Code strings actually modified: ${changesMade}`);
    console.log(`Changes made breakdown:`);
    correctionDetails.forEach((detail, index) => {
        if (detail.changed) {
            console.log(`  - Code string ${index + 1}: MODIFIED`);
        }
    });
    console.log("=== PIPE CORRECTION PROCESS COMPLETED ===\n");
    
    return {
        correctedText: correctedText,
        changesMade: changesMade,
        details: correctionDetails
    };
}

/**
 * Utility function to test the pipe correction with various scenarios
 */
export function testPipeCorrection() {
    console.log("\n🧪 === TESTING PIPE CORRECTION FUNCTION === 🧪");
    
    // Test case 1: Single row parameter with 12 pipes (needs 3 added) - Your example
    console.log("\n📋 TEST CASE 1: Single row with 12 pipes (needs 3 added)");
    const test1 = '<CONST-E; row1 = "V5|# of Employees|||||F|F|F|F|F|F|";>';
    console.log("🔍 Input: ", test1);
    console.log("📊 Expected: Add 3 pipes after column 3");
    console.log("📊 Expected detailed logs showing BEFORE/AFTER for row1");
    const result1 = autoCorrectPipeCounts(test1);
    console.log("✅ Output:", result1.correctedText);
    console.log("📈 Changes Made:", result1.changesMade);
    console.log("🔢 Pipe count verification:", (result1.correctedText.match(/\|/g) || []).length, "pipes total");
    
    // Test case 2: Multiple row parameters with different pipe counts
    console.log("\n📋 TEST CASE 2: Multiple rows with different pipe counts");
    const test2 = '<CONST-E; row1 = "A1|CEO||||||$100000|F|F|F|F|F|F|"; row2 = "B2|Manager||||||$50000|G|G|";>';
    console.log("🔍 Input: ", test2);
    console.log("📊 Expected: row1 has 14 pipes (needs 1), row2 has 10 pipes (needs 5)");
    const result2 = autoCorrectPipeCounts(test2);
    console.log("✅ Output:", result2.correctedText);
    console.log("📈 Changes Made:", result2.changesMade);
    
    // Test case 3: Already correct pipe counts (15 pipes)
    console.log("\n📋 TEST CASE 3: Already correct pipe count");
    const test3 = '<CONST-E; row1 = "A1|CEO|||||||$100000|F|F|F|F|F|F|";>';
    console.log("🔍 Input: ", test3);
    console.log("📊 Expected: No changes needed (already 15 pipes)");
    const result3 = autoCorrectPipeCounts(test3);
    console.log("✅ Output:", result3.correctedText);
    console.log("📈 Changes Made:", result3.changesMade);
    console.log("🔢 Pipe count verification:", (result3.correctedText.match(/\|/g) || []).length, "pipes total");
    
    // Test case 4: Too many pipes (needs pipes removed)
    console.log("\n📋 TEST CASE 4: Too many pipes (needs removal)");
    const test4 = '<CONST-E; row1 = "A1|CEO|||||||||||||$100000|F|F|F|F|F|F|";>';
    console.log("🔍 Input: ", test4);
    console.log("📊 Expected: Remove excess pipes after column 3");
    const result4 = autoCorrectPipeCounts(test4);
    console.log("✅ Output:", result4.correctedText);
    console.log("📈 Changes Made:", result4.changesMade);
    console.log("🔢 Pipe count verification:", (result4.correctedText.match(/\|/g) || []).length, "pipes total");
    
    // Test case 5: Multiple codestrings in a single input
    console.log("\n📋 TEST CASE 5: Multiple codestrings in one input");
    const test5 = '<BR; row1 = "|||";><CONST-E; row1 = "V1|Sales|||||100|200|";><CONST-E; row1 = "V2|Price|||||||$10|$15|$20|$25|$30|$35|";>';
    console.log("🔍 Input: ", test5);
    console.log("📊 Expected: BR has 3 pipes (needs 12), first CONST-E has 8 pipes (needs 7), second CONST-E has 15 pipes (no change)");
    const result5 = autoCorrectPipeCounts(test5);
    console.log("✅ Output:", result5.correctedText);
    console.log("📈 Changes Made:", result5.changesMade);
    
    // Test case 6: No row parameters
    console.log("\n📋 TEST CASE 6: No row parameters");
    const test6 = 'Some text without any row parameters at all';
    console.log("🔍 Input: ", test6);
    console.log("📊 Expected: No changes (no row parameters found)");
    const result6 = autoCorrectPipeCounts(test6);
    console.log("✅ Output:", result6.correctedText);
    console.log("📈 Changes Made:", result6.changesMade);
    
    // Test case 7: Detailed logging demonstration
    console.log("\n📋 TEST CASE 7: Detailed logging demonstration");
    const test7 = '<FORMULA-S; row1 = "A1|Salary|F|F|F|F|"; row2 = "A2|Bonus||||||||||||||||||F|F|F|";>';
    console.log("🔍 Input: ", test7);
    console.log("📊 Expected: row1 has 6 pipes (needs 9), row2 has 20 pipes (needs 5 removed)");
    console.log("📊 Expected: Detailed BEFORE/AFTER logs for both row1 and row2");
    const result7 = autoCorrectPipeCounts(test7);
    console.log("✅ Output:", result7.correctedText);
    console.log("📈 Changes Made:", result7.changesMade);
    
    // Summary
    console.log("\n📊 === TEST SUMMARY ===");
    const totalTests = 7;
    const results = [result1, result2, result3, result4, result5, result6, result7];
    const testsWithChanges = results.filter(r => r.changesMade > 0).length;
    
    console.log(`✅ Total tests run: ${totalTests}`);
    console.log(`🔧 Tests that made changes: ${testsWithChanges}`);
    console.log(`✨ Tests with no changes needed: ${totalTests - testsWithChanges}`);
    console.log(`📝 Total changes across all tests: ${results.reduce((sum, r) => sum + r.changesMade, 0)}`);
    
    console.log("\n🎉 === TESTING COMPLETE === 🎉");
    
    return {
        totalTests,
        testsWithChanges,
        totalChanges: results.reduce((sum, r) => sum + r.changesMade, 0),
        results,
        detailedLoggingDemo: result7 // Highlight the detailed logging test case
    };
} 