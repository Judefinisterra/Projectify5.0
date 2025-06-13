// IntegratePipeValidation.js - Example integration with existing validation system

import { autoCorrectPipeCounts } from './PipeValidation.js';
import { validateFormatErrors } from './ValidationFormat.js';
import { validateLogicErrors } from './ValidationLogic.js';

/**
 * Enhanced validation function that automatically corrects pipe counts BEFORE running other validations
 * This sits outside the normal validation process and corrects codestrings automatically
 * @param {string|Array} inputText - Input text or array of codestrings
 * @returns {Object} - Validation results with corrected text and all errors
 */
export async function validateWithPipeCorrection(inputText) {
    console.log("🔧 === ENHANCED VALIDATION WITH PIPE CORRECTION ===");
    
    // Step 1: Automatically correct pipe counts (this won't trigger errors, just fixes them)
    console.log("🔧 Step 1: Auto-correcting pipe counts...");
    const pipeCorrection = autoCorrectPipeCounts(inputText);
    
    // Step 2: Run standard validations on the corrected text
    console.log("🔧 Step 2: Running format validation...");
    const formatErrors = await validateFormatErrors(pipeCorrection.correctedText);
    
    console.log("🔧 Step 3: Running logic validation...");
    const logicErrors = await validateLogicErrors(pipeCorrection.correctedText);
    
    // Combine all results
    const result = {
        // Corrected text (ready to use)
        correctedText: pipeCorrection.correctedText,
        
        // Pipe correction info
        pipeCorrection: {
            changesMade: pipeCorrection.changesMade,
            details: pipeCorrection.details
        },
        
        // Standard validation errors
        formatErrors: formatErrors,
        logicErrors: logicErrors,
        
        // Combined error list
        allErrors: [...formatErrors, ...logicErrors],
        
        // Summary
        summary: {
            pipeCorrections: pipeCorrection.changesMade,
            formatErrors: formatErrors.length,
            logicErrors: logicErrors.length,
            totalErrors: formatErrors.length + logicErrors.length,
            hasErrors: formatErrors.length > 0 || logicErrors.length > 0
        }
    };
    
    // Log summary
    console.log("📊 === VALIDATION SUMMARY ===");
    console.log(`🔧 Pipe corrections made: ${result.summary.pipeCorrections}`);
    console.log(`⚠️  Format errors found: ${result.summary.formatErrors}`);
    console.log(`⚠️  Logic errors found: ${result.summary.logicErrors}`);
    console.log(`📋 Total errors: ${result.summary.totalErrors}`);
    console.log(`✅ Validation status: ${result.summary.hasErrors ? 'FAILED' : 'PASSED'}`);
    
    return result;
}

/**
 * Standalone pipe correction function for use in other parts of the system
 * @param {string|Array} inputText - Input text or array of codestrings
 * @returns {string|Array} - Corrected text (same type as input)
 */
export function correctPipesOnly(inputText) {
    console.log("🔧 Running pipe correction only...");
    const result = autoCorrectPipeCounts(inputText);
    
    if (result.changesMade > 0) {
        console.log(`✅ Pipe correction completed: ${result.changesMade} changes made`);
    } else {
        console.log("✅ Pipe correction completed: No changes needed");
    }
    
    return result.correctedText;
}

/**
 * Example usage function demonstrating the pipe correction system
 */
export function demonstratePipeCorrection() {
    console.log("🎯 === PIPE CORRECTION DEMONSTRATION ===");
    
    // Example 1: Your specific use case
    const example1 = '<CONST-E; row1 = "V5|# of Employees|||||F|F|F|F|F|F|";>';
    console.log("\n📋 Example 1: Your specific case (12 pipes -> 15 pipes)");
    console.log("🔍 Before:", example1);
    
    const corrected1 = correctPipesOnly(example1);
    console.log("✅ After: ", corrected1);
    
    // Example 2: Multiple issues
    const example2 = `
        <CONST-E; row1 = "V1|Sales|||F|F|F|F|F|F|";>
        <CONST-E; row2 = "V2|Price|||||||||||||||||$10|$15|$20|$25|$30|$35|";>
        <BR; row1 = "|||";>
    `;
    
    console.log("\n📋 Example 2: Multiple codestrings with different pipe issues");
    console.log("🔍 Before:", example2.trim());
    
    const corrected2 = correctPipesOnly(example2);
    console.log("✅ After: ", corrected2);
    
    // Example 3: Full validation with pipe correction
    console.log("\n📋 Example 3: Full validation with automatic pipe correction");
    validateWithPipeCorrection(example1).then(result => {
        console.log("📊 Full validation result:");
        console.log("- Corrected text:", result.correctedText);
        console.log("- Pipe corrections:", result.pipeCorrection.changesMade);
        console.log("- Format errors:", result.summary.formatErrors);
        console.log("- Logic errors:", result.summary.logicErrors);
        console.log("- Overall status:", result.summary.hasErrors ? "❌ FAILED" : "✅ PASSED");
    });
}

// Auto-run demonstration if this file is loaded directly
if (typeof window !== 'undefined') {
    // Browser environment
    window.demonstratePipeCorrection = demonstratePipeCorrection;
    console.log("🎯 Pipe validation system loaded! Run demonstratePipeCorrection() to see examples.");
} 