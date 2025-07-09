// ValidationLogic.js - Handles logic validation errors (LERR codes)

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
        
        // Check for hardcoded numbers in custom formulas
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
        
    } catch (error) {
        errors.push(`Formula syntax error: ${error.message}`);
    }
    
    return errors;
}

export async function validateLogicErrors(inputText) {
    const errors = [];
    const tabLabels = new Set();
    const rowValues = new Set();
    const codeTypes = new Set();
    const tabRowDrivers = new Map(); // Track row drivers per TAB
    
    // Extract all code strings using regex pattern /<[^>]+>/g
    // This allows multiple code strings on the same line
    let inputCodeStrings = [];
    if (typeof inputText === 'string') {
        const codeStringMatches = inputText.match(/<[^>]+>/g);
        inputCodeStrings = codeStringMatches || [];
    } else if (Array.isArray(inputText)) {
        // If already an array, extract code strings from each element
        inputCodeStrings = [];
        for (const item of inputText) {
            if (typeof item === 'string') {
                const matches = item.match(/<[^>]+>/g);
                if (matches) {
                    inputCodeStrings.push(...matches);
                }
            }
        }
    }
    
    if (inputCodeStrings.length === 0) {
        console.log("No code strings found in input for logic validation");
        return errors;
    }
    
    console.log(`Found ${inputCodeStrings.length} code strings for logic validation`);
    
    // Remove line breaks within angle brackets before processing
    inputCodeStrings = inputCodeStrings.map(str => {
        // Replace line breaks within < > brackets
        return str.replace(/<[^>]*>/g, match => {
            // Remove all types of line breaks within the match
            return match.replace(/[\r\n]+/g, ' ');
        });
    });
    console.log("Preprocessed code strings (line breaks removed within <>):", inputCodeStrings);
    
    // Clean up input strings by removing anything outside angle brackets
    inputCodeStrings = inputCodeStrings.map(str => {
        const match = str.match(/<[^>]+>/);
        return match ? match[0] : str;
    });

    console.log("CleanedInput Code Strings:", inputCodeStrings);

    // Load valid codes from the correct location
    let validCodes = new Set();
    try {
        const response = await fetch('../prompts/Codes.txt');
        if (!response.ok) {
            throw new Error('Failed to load Codes.txt file');
        }
        const fileContent = await response.text();
        validCodes = new Set(fileContent.split('\n')
            .map(line => line.trim())
            .filter(line => line.length > 0));
    } catch (error) {
        errors.push(`[LERR001] Error reading Codes.txt file: ${error.message}`);
        return errors;
    }

    // Determine which tab each code belongs to
    let currentTab = 'default'; // For codes before any TAB
    const codeToTab = new Map();
    
    // Initialize the default tab
    tabRowDrivers.set('default', new Map());
    
    for (let i = 0; i < inputCodeStrings.length; i++) {
        const codeString = inputCodeStrings[i];
        const codeMatch = codeString.match(/<([^;]+);/);
        
        if (codeMatch && codeMatch[1].trim() === 'TAB') {
            currentTab = codeString; // Use the TAB code string as the tab identifier
            tabRowDrivers.set(currentTab, new Map());
        }
        
        codeToTab.set(i, currentTab);
    }

    // First pass: collect all row values, code types, and suffixes
    const financialStatementItems = new Map(); // Track financial statement items (column B) by their location
    
    // Define valid financial statement codes
    const validFinancialCodes = new Set([
        'is: revenue',
        'is: direct costs',
        'is: corporate overhead',
        'is: d&a',
        'is: interest',
        'is: other income',
        'is: net income',
        'bs: current assets',
        'bs: fixed assets',
        'bs: current liabilities',
        'bs: lt liabilities',
        'bs: equity',
        'cf: wc',
        'cf: non-cash',
        'cf: cfi',
        'cf: cff'
    ]);
    
    for (let i = 0; i < inputCodeStrings.length; i++) {
        const codeString = inputCodeStrings[i];
        const currentTabForCode = codeToTab.get(i);
        
        if (!codeString.startsWith('<') || !codeString.endsWith('>')) {
            errors.push(`[LERR002] Invalid code string format: ${codeString}`);
            continue;
        }
        if (codeString.startsWith('<BR>')) {
            continue;
        }

        // Validate customformula parameters
        const customFormulaMatches = codeString.match(/customformula\d*\s*=\s*"([^"]*)"/g);
        if (customFormulaMatches) {
            customFormulaMatches.forEach(match => {
                const formulaMatch = match.match(/customformula\d*\s*=\s*"([^"]*)"/);
                if (formulaMatch) {
                    const formula = formulaMatch[1];
                    const formulaErrors = validateCustomFormula(formula);
                    formulaErrors.forEach(error => {
                        errors.push(`[LERR014] Custom formula validation error in ${codeString}: ${error}`);
                    });
                }
            });
        }

        const rowMatches = codeString.match(/row\d+\s*=\s*"([^"]*)"/g);
        if (rowMatches) {
            rowMatches.forEach(match => {
                const rowParam = match.match(/(row\d+)\s*=\s*"([^"]*)"/);
                const rowName = rowParam[1];
                const rowContent = rowParam[2];
                // Handle spaces before/after the pipe delimiter
                const parts = rowContent.split('|');
                if (parts.length > 0) {
                    const driver = parts[0].trim();
                    
                    // Check for financial statement items
                    if (parts.length >= 3) {
                        const columnB = parts[1].trim();
                        const columnC = parts[2].trim();
                        
                        // Check if column 3 starts with IS, BS, or CF
                        if (columnC && (columnC.startsWith('IS:') || columnC.startsWith('BS:') || columnC.startsWith('CF:'))) {
                            // Validate against approved financial codes
                            if (!validFinancialCodes.has(columnC.toLowerCase())) {
                                errors.push(`[LERR013] Invalid financial statement code "${columnC}" in ${codeString} at ${rowName}. Must be one of the approved codes.`);
                            }
                            
                            if (columnB) {
                                // Track this financial statement item for duplicate checking
                                if (financialStatementItems.has(columnB)) {
                                    const existing = financialStatementItems.get(columnB);
                                    existing.push({ codeString, rowName, driver });
                                } else {
                                    financialStatementItems.set(columnB, [{ codeString, rowName, driver }]);
                                }
                            }
                        }
                    }
                    
                    // Track row drivers per tab
                    if (driver && !driver.startsWith('*') && tabRowDrivers.has(currentTabForCode)) {
                        const driversInThisTab = tabRowDrivers.get(currentTabForCode);
                        if (driversInThisTab.has(driver)) {
                            const existing = driversInThisTab.get(driver);
                            existing.occurrences.push({ codeString, rowName });
                        } else {
                            driversInThisTab.set(driver, {
                                firstOccurrence: { codeString, rowName },
                                occurrences: [{ codeString, rowName }]
                            });
                        }
                    }
                    
                    // Extract all potential row IDs including those after asterisks
                    parts.forEach(part => {
                        const trimmedPart = part.trim();
                        if (trimmedPart.startsWith('*')) {
                            // Add the ID after the asterisk
                            const afterAsterisk = trimmedPart.substring(1).trim();
                            if (afterAsterisk) {
                                rowValues.add(afterAsterisk);
                            }
                        } else if (trimmedPart) {
                            // Add regular IDs
                            rowValues.add(trimmedPart);
                        }
                    });
                }
            });
        }

        // Extract code type and suffix
        const codeMatch = codeString.match(/<([^;]+);/);
        if (codeMatch) {
            const codeType = codeMatch[1].trim();
            codeTypes.add(codeType);
        }
    }

    // Check for duplicate row drivers within each tab
    tabRowDrivers.forEach((driversInTab, tabCode) => {
        driversInTab.forEach((driverInfo, driver) => {
            if (driverInfo.occurrences.length > 1) {
                const tabLabel = tabCode === 'default' ? 'before any TAB' : 
                    (() => {
                        const labelMatch = tabCode.match(/label\d+="([^"]*)"/);
                        return labelMatch ? `TAB "${labelMatch[1]}"` : 'TAB (no label)';
                    })();
                
                const locations = driverInfo.occurrences.map(loc => {
                    const codeMatch = loc.codeString.match(/<([^;]+);/);
                    const codeType = codeMatch ? codeMatch[1] : 'Unknown';
                    return `${codeType} code (${loc.rowName})`;
                }).join(' and ');
                
                errors.push(`[LERR003] Duplicate row driver "${driver}" found within ${tabLabel}: ${locations}`);
                
                // Add details about each occurrence
                driverInfo.occurrences.forEach(loc => {
                    errors.push(`  - In ${loc.codeString.substring(0, 100)}... at ${loc.rowName}`);
                });
            }
        });
    });

    // Check for duplicate financial statement items
    financialStatementItems.forEach((occurrences, itemName) => {
        if (occurrences.length > 1) {
            // Exception: Skip depreciation and amortization items
            if (/depreciation|amortization/i.test(itemName)) {
                return; // Skip this error for depreciation/amortization items
            }
            
            // Get unique tabs for better error reporting
            const tabSet = new Set();
            occurrences.forEach(loc => {
                const tabIndex = inputCodeStrings.indexOf(loc.codeString);
                const tab = codeToTab.get(tabIndex);
                if (tab !== 'default') {
                    const labelMatch = tab.match(/label\d+="([^"]*)"/);
                    tabSet.add(labelMatch ? labelMatch[1] : 'Unnamed Tab');
                } else {
                    tabSet.add('Before any TAB');
                }
            });
            
            const tabList = Array.from(tabSet).join(', ');
            
                            errors.push(`[FERR010] Duplicate financial statement item "${itemName}" found in ${occurrences.length} locations across tabs: ${tabList}`);
            
            // Add details about each occurrence
            occurrences.forEach(loc => {
                errors.push(`  - In ${loc.codeString.substring(0, 100)}... at ${loc.rowName} (driver: ${loc.driver || 'none'})`);
            });
        }
    });

    // Second pass: detailed validation
    for (const codeString of inputCodeStrings) {
        // Skip BR tags completely
        if (codeString === '<BR>') {
            continue;
        }
        
        const codeMatch = codeString.match(/<([^;]+);/);
        if (!codeMatch) {
            errors.push(`[LERR004] Cannot extract code type from: ${codeString}`);
            continue;
        }

        const codeType = codeMatch[1].trim();
        
        // Validate code exists in description file
        if (!validCodes.has(codeType)) {
            errors.push(`[LERR005] Invalid code type: "${codeType}" not found in valid codes list - ${codeString}`);
        }

        // Validate mathematical operator codes driver parameters (LERR016)
        if (codeType.endsWith('-S')) {
            // Define the mathematical operator codes that need validation
            const mathOperatorCodes = [
                'MULT3-S', 'DIVIDE2-S', 'SUBTRACT2-S', 'SUBTOTAL2-S', 'SUBTOTAL3-S', 
                'AVGMULT3-S', 'ANNUALIZE-S', 'DEANNUALIZE-S', 'AVGDEANNUALIZE2-S', 
                'DIRECT-S', 'CHANGE-S', 'INCREASE-S', 'DECREASE-S', 'GROWTH-S', 
                'OFFSETCOLUMN-S', 'OFFSET2-S', 'SUM2-S', 'DISCOUNT2-S', 'FORMULA-S', 'COLUMNFORMULA-S'
            ];
            
            if (mathOperatorCodes.includes(codeType)) {
                // Extract existing driver parameters
                const driverParams = [];
                const driver1Match = codeString.match(/driver1\s*=\s*"([^"]*)"/);
                const driver2Match = codeString.match(/driver2\s*=\s*"([^"]*)"/);
                const driver3Match = codeString.match(/driver3\s*=\s*"([^"]*)"/);
                
                if (driver1Match) driverParams.push('driver1');
                if (driver2Match) driverParams.push('driver2');
                if (driver3Match) driverParams.push('driver3');
                
                // Determine expected drivers based on code pattern
                let expectedDrivers = [];
                
                if (codeType === 'FORMULA-S') {
                    // FORMULA-S should have no driver parameters
                    expectedDrivers = [];
                } else if (codeType === 'COLUMNFORMULA-S') {
                    // COLUMNFORMULA-S should have no driver parameters
                    expectedDrivers = [];
                } else if (codeType.match(/\d+3-S$/)) {
                    // Codes ending with "3-S" should have driver1, driver2, driver3
                    expectedDrivers = ['driver1', 'driver2', 'driver3'];
                } else if (codeType.match(/\d+2-S$/)) {
                    // Codes ending with "2-S" should have driver1, driver2
                    expectedDrivers = ['driver1', 'driver2'];
                } else if (codeType.match(/^[A-Z]+[A-Z]+-S$/)) {
                    // Codes with no number before "-S" should have only driver1
                    expectedDrivers = ['driver1'];
                }
                
                // Validate driver parameters
                if (expectedDrivers.length !== driverParams.length) {
                    const expectedStr = expectedDrivers.length === 0 ? 'no driver parameters' : expectedDrivers.join(', ');
                    const actualStr = driverParams.length === 0 ? 'no driver parameters' : driverParams.join(', ');
                    errors.push(`[LERR016] Mathematical operator code "${codeType}" requires ${expectedStr}, but found ${actualStr} - ${codeString}`);
                } else {
                    // Check if the right driver parameters are present
                    const missingDrivers = expectedDrivers.filter(driver => !driverParams.includes(driver));
                    const extraDrivers = driverParams.filter(driver => !expectedDrivers.includes(driver));
                    
                    if (missingDrivers.length > 0) {
                        errors.push(`[LERR016] Mathematical operator code "${codeType}" missing required parameters: ${missingDrivers.join(', ')} - ${codeString}`);
                    }
                    if (extraDrivers.length > 0) {
                        errors.push(`[LERR016] Mathematical operator code "${codeType}" has unexpected parameters: ${extraDrivers.join(', ')} - ${codeString}`);
                    }
                }
            }
        }

        // Validate TAB labels
        if (codeType === 'TAB') {
            const labelMatch = codeString.match(/label\d+="([^"]*)"/);
            if (!labelMatch) {
                errors.push(`[LERR006] TAB code missing label parameter - ${codeString}`);
                continue;
            }

            const label = labelMatch[1];
            
            if (label.length > 30) {
                errors.push(`[LERR007] Tab label too long (max 30 chars): "${label}" - ${codeString}`);
            }
            
            if (/[,":;]/.test(label)) {
                errors.push(`[LERR008] Tab label contains illegal characters (,":;): "${label}" - ${codeString}`);
            }
            
            if (tabLabels.has(label)) {
                errors.push(`[LERR009] Duplicate tab label: "${label}" - ${codeString}`);
            }
            tabLabels.add(label);
            
            // Initialize driver tracking for this TAB
            if (!tabRowDrivers.has(codeString)) {
                tabRowDrivers.set(codeString, new Map());
            }
        }

        // Validate row format and check for duplicate drivers within TAB
        if (codeType === 'TAB') {
            const rowMatches = codeString.match(/row\d+="([^"]*)"/g);
            if (rowMatches) {
                const driversInThisTab = tabRowDrivers.get(codeString);
                
                rowMatches.forEach(match => {
                    const rowParam = match.match(/(row\d+)="([^"]*)"/);
                    const rowName = rowParam[1];
                    const rowContent = rowParam[2];
                    const parts = rowContent.split('|');
                    
                    if (parts.length < 2) {
                        errors.push(`[LERR010] Invalid row format (missing required fields): "${rowContent}" in ${codeString}`);
                    } else {
                        // Extract the driver (first part before |)
                        const driver = parts[0].trim();
                        
                        // Check if this driver already exists in this TAB
                        if (driversInThisTab.has(driver)) {
                            const existingRow = driversInThisTab.get(driver);
                            errors.push(`[LERR003] Duplicate row driver "${driver}" found in ${codeString} - appears in both ${existingRow} and ${rowName}`);
                        } else {
                            driversInThisTab.set(driver, rowName);
                        }
                    }
                });
            }
        } else {
            // For non-TAB codes, just validate row format
            const rowMatches = codeString.match(/row\d+="([^"]*)"/g);
            if (rowMatches) {
                rowMatches.forEach(match => {
                    const rowContent = match.match(/row\d+="([^"]*)"/)[1];
                    const parts = rowContent.split('|');
                    if (parts.length < 2) {
                        errors.push(`[LERR010] Invalid row format (missing required fields): "${rowContent}" in ${codeString}`);
                    }
                });
            }
        }
    }

    // Third pass: validate driver references after all row IDs are collected
    for (const codeString of inputCodeStrings) {
        const driverMatches = codeString.match(/driver\d+\s*=\s*"([^"]*)"/g);
        if (driverMatches) {
            driverMatches.forEach(match => {
                const driverValue = match.match(/driver\d+\s*=\s*"([^"]*)"/)[1].trim();
                if (!rowValues.has(driverValue)) {
                    errors.push(`[LERR011] Driver value "${driverValue}" not found in any row - in ${codeString}`);
                }
            });
        }
    }

    // Return array of errors directly
    return errors;
}

// Modified export function for compatibility
export async function runLogicValidation(inputStrings) {
    try {
        const cleanedStrings = inputStrings
            .map(line => line.trim())
            .filter(line => line.length > 0)
            .map(line => line.replace(/^'|',$|,$|^"|"$/g, ''));

        const codeStrings = cleanedStrings.join(' ').match(/<[^>]+>/g) || [];

        if (codeStrings.length === 0) {
            console.error('Logic validation run with no codestrings - input is empty');
            return ['No code strings found to validate'];
        }

        const validationErrors = await validateLogicErrors(codeStrings);
        return validationErrors;
    } catch (error) {
        console.error('Logic validation failed:', error);
        return [`Logic validation failed: ${error.message}`];
    }
} 