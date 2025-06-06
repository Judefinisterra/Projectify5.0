// Remove all these imports as they're no longer needed
// import { promises as fs } from 'fs';
// import { join } from 'path';
// import { fileURLToPath } from 'url';
// import { dirname } from 'path';

// Import the new separated validation modules
import { validateLogicErrors } from './ValidationLogic.js';
import { validateFormatErrors } from './ValidationFormat.js';

// DEPRECATED: These wrapper functions are provided for backward compatibility
// New code should use validateLogicErrors and validateFormatErrors directly

// Helper function to validate custom formula syntax
function validateCustomFormula(formula) {
    const errors = [];
    
    if (!formula || formula.trim() === '') {
        return errors; // Empty formula is allowed
    }
    
    // Prepare formula for validation
    let testFormula = formula.trim();
    
    // Add = if not present (as mentioned in requirements)
    if (!testFormula.startsWith('=')) {
        testFormula = '=' + testFormula;
    }
    
    // Replace custom functions with valid Excel references for syntax checking
    // Replace dr{...}, cd{...}, etc. with placeholder cell references
    testFormula = testFormula.replace(/dr\{[^}]*\}/g, 'A1');
    testFormula = testFormula.replace(/cd\{[^}]*\}/g, 'B1');
    testFormula = testFormula.replace(/rd\{[^}]*\}/g, 'C1');
    testFormula = testFormula.replace(/cr\{[^}]*\}/g, 'D1');
    
    // Basic Excel formula syntax validation
    try {
        // Check for balanced parentheses
        let parenCount = 0;
        for (let i = 0; i < testFormula.length; i++) {
            if (testFormula[i] === '(') parenCount++;
            if (testFormula[i] === ')') parenCount--;
            if (parenCount < 0) {
                errors.push(`Unmatched closing parenthesis at position ${i + 1}`);
                break;
            }
        }
        if (parenCount > 0) {
            errors.push(`${parenCount} unmatched opening parenthesis(es)`);
        }
        
        // Check for common Excel function syntax errors
        // Look for function calls that might be missing commas
        const functionPattern = /(\w+)\s*\(/g;
        let match;
        while ((match = functionPattern.exec(testFormula)) !== null) {
            const funcName = match[1].toLowerCase();
            const startPos = match.index + match[0].length;
            
            // Find the matching closing parenthesis for this function
            let depth = 1;
            let pos = startPos;
            let args = [];
            let currentArg = '';
            
            while (pos < testFormula.length && depth > 0) {
                const char = testFormula[pos];
                if (char === '(') {
                    depth++;
                    currentArg += char;
                } else if (char === ')') {
                    depth--;
                    if (depth === 0) {
                        if (currentArg.trim()) args.push(currentArg.trim());
                    } else {
                        currentArg += char;
                    }
                } else if (char === ',' && depth === 1) {
                    args.push(currentArg.trim());
                    currentArg = '';
                } else {
                    currentArg += char;
                }
                pos++;
            }
            
            // Check specific functions that commonly need commas
            if (['offset', 'index', 'match', 'vlookup', 'hlookup', 'indirect'].includes(funcName)) {
                // Look for patterns like offset(A10) which should be offset(A1,0,0)
                if (args.length === 1 && /^[A-Z]+\d+$/.test(args[0])) {
                    errors.push(`Function ${funcName.toUpperCase()}() likely missing required arguments - found only "${args[0]}"`);
                }
                
                // Check for missing commas in arguments (like "A10" immediately followed by numbers)
                args.forEach((arg, index) => {
                    // Look for patterns like "A1123" which should be "A1,123"
                    const suspiciousPattern = /^([A-Z]+\d+)(\d+)$/;
                    const suspiciousMatch = arg.match(suspiciousPattern);
                    if (suspiciousMatch) {
                        errors.push(`Function ${funcName.toUpperCase()}() argument ${index + 1} "${arg}" appears to be missing a comma between "${suspiciousMatch[1]}" and "${suspiciousMatch[2]}"`);
                    }
                });
            }
        }
        
        // Check for other common syntax issues
        // Missing operators between terms
        if (/[A-Z]\d+[A-Z]\d+/.test(testFormula)) {
            errors.push(`Missing operator between cell references (e.g., A1B1 should be A1+B1 or A1*B1)`);
        }
        
        // Invalid characters in cell references
        const invalidCellRef = /[A-Z]+\d+[A-Za-z]+\d*/g;
        let invalidMatch;
        while ((invalidMatch = invalidCellRef.exec(testFormula)) !== null) {
            if (!/^[A-Z]+\d+$/.test(invalidMatch[0])) {
                errors.push(`Invalid cell reference format: "${invalidMatch[0]}"`);
            }
        }
        
        // >>> ADDED: Check for hardcoded numbers in custom formulas
        // This validates against using hardcoded numeric values instead of cell references
        const originalFormula = formula.trim();
        
        // Look for numeric patterns that are likely hardcoded values
        // Match integers and decimals (positive/negative) but exclude cell references
        // and exclude numbers within curly braces (driver references like rd{V12})
        
        // First, temporarily replace all driver references with placeholders to avoid false positives
        let formulaForNumberCheck = originalFormula;
        formulaForNumberCheck = formulaForNumberCheck.replace(/[a-z]{2}\{[^}]*\}/g, 'DRIVER_REF');
        
        const numberPattern = /(?<![A-Z])(-?\d+(?:\.\d+)?)(?![A-Z])/g;
        let numberMatch;
        const foundNumbers = [];
        
        while ((numberMatch = numberPattern.exec(formulaForNumberCheck)) !== null) {
            const number = numberMatch[1];
            const position = numberMatch.index;
            
            // Skip acceptable hardcoded numbers: -1, 0, and 1
            if (number === '-1' || number === '0' || number === '1') {
                continue;
            }
            
            // Get the corresponding position in the original formula to show proper context
            // Find this number in the original formula at approximately the same position
            const originalContext = originalFormula.substring(Math.max(0, position - 10), Math.min(originalFormula.length, position + number.length + 10));
            
            foundNumbers.push({
                number: number,
                position: position,
                context: originalContext.trim()
            });
        }
        
        if (foundNumbers.length > 0) {
            foundNumbers.forEach(item => {
                errors.push(`Hardcoded number "${item.number}" found in formula - consider using a cell reference instead (context: "${item.context}")`);
            });
        }
        // <<< END ADDED
        
    } catch (error) {
        errors.push(`Formula syntax error: ${error.message}`);
    }
    
    return errors;
}
// <<< END ADDED

// DEPRECATED: Wrapper function for backward compatibility
// Use validateLogicErrors and validateFormatErrors directly for new code
export async function validateCodeStrings(inputText) {
    // Combine both logic and format validation errors
    const logicErrors = await validateLogicErrors(inputText);
    const formatErrors = await validateFormatErrors(inputText);
    
    // Combine errors and return as string (for backward compatibility)
    const allErrors = [...logicErrors, ...formatErrors];
    return allErrors.join('\n');
}

// DEPRECATED: Wrapper function for backward compatibility  
// Use validateLogicErrors and validateFormatErrors directly for new code
export async function validateCodeStringsForRun(inputText) {
    // Combine both logic and format validation errors
    const logicErrors = await validateLogicErrors(inputText);
    const formatErrors = await validateFormatErrors(inputText);
    
    // Return combined errors as array (for backward compatibility)
    return [...logicErrors, ...formatErrors];
}

// Modified export function for backward compatibility
export async function runValidation(inputStrings) {
    try {
        const cleanedStrings = inputStrings
            .map(line => line.trim())
            .filter(line => line.length > 0)
            .map(line => line.replace(/^'|',$|,$|^"|"$/g, ''));

        const codeStrings = cleanedStrings.join(' ').match(/<[^>]+>/g) || [];

        if (codeStrings.length === 0) {
            console.error('Validation run with no codestrings - input is empty');
            return ['No code strings found to validate'];
        }

        const validationErrors = await validateCodeStrings(codeStrings);
        return validationErrors;
    } catch (error) {
        console.error('Validation failed:', error);
        return [`Validation failed: ${error.message}`];
    }
}

