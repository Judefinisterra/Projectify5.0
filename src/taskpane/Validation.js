// Remove all these imports as they're no longer needed
// import { promises as fs } from 'fs';
// import { join } from 'path';
// import { fileURLToPath } from 'url';
// import { dirname } from 'path';

export async function validateCodeStrings(inputCodeStrings) {
    const errors = [];
    const tabLabels = new Set();
    const rowValues = new Set();
    const codeTypes = new Set();
    const tabRowDrivers = new Map(); // Track row drivers per TAB
    
    // >>> ADDED: Remove line breaks within angle brackets before processing
    inputCodeStrings = inputCodeStrings.map(str => {
        // Replace line breaks within < > brackets
        return str.replace(/<[^>]*>/g, match => {
            // Remove all types of line breaks within the match
            return match.replace(/[\r\n]+/g, ' ');
        });
    });
    console.log("Preprocessed code strings (line breaks removed within <>):", inputCodeStrings);
    // <<< END ADDED
    
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
        'IS: revenue',
        'IS: direct costs',
        'IS: corporate overhead',
        'IS: D&A',
        'IS: interest',
        'IS: Other Income',
        'IS: Net Income',
        'BS: current assets',
        'BS: fixed assets',
        'BS: current liabilities',
        'BS: LT Liabilities',
        'BS: equity',
        'CF: WC',
        'CF: Non-cash',
        'CF: CFI',
        'CF: CFF'
    ]);
    
    for (let i = 0; i < inputCodeStrings.length; i++) {
        const codeString = inputCodeStrings[i];
        const currentTabForCode = codeToTab.get(i);
        
        if (!codeString.startsWith('<') || !codeString.endsWith('>')) {
            errors.push(`[LERR002] Invalid code string format: ${codeString}`);
            continue;
        }
        // Extract and store row IDs
        if (codeString.startsWith('<BR>')) {
            // Skip extraction for BR tags
            continue;
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
                            if (!validFinancialCodes.has(columnC)) {
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

        
        // console.log(rowMatches);
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
            
            errors.push(`[LERR012] Duplicate financial statement item "${itemName}" found in ${occurrences.length} locations across tabs: ${tabList}`);
            
            // Add details about each occurrence
            occurrences.forEach(loc => {
                errors.push(`  - In ${loc.codeString.substring(0, 100)}... at ${loc.rowName} (driver: ${loc.driver || 'none'})`);
            });
        }
    });

    // Format Validation Section
    const formatRequiredCodes = new Set([
        'FINANCIALS-S', 'MULT3-S', 'DIVIDE2-S', 'SUBTRACT2-S', 'SUBTOTAL2-S', 
        'SUBTOTAL3-S', 'AVGMULT3-S', 'ANNUALIZE-S', 'DEANNUALIZE-S', 
        'AVGDEANNUALIZE2-S', 'DIRECT-S', 'CHANGE-S', 'INCREASE-S', 'DECREASE-S', 
        'GROWTH-S', 'OFFSETCOLUMN-S', 'OFFSET2-S', 'SUM2-S', 'DISCOUNT2-S', 
        'CONST-E', 'SPREAD-E', 'ENDPOINT-E', 'GROWTH-E'
    ]);

    // Check format parameters and adjacency rules
    for (let i = 0; i < inputCodeStrings.length; i++) {
        const codeString = inputCodeStrings[i];
        const codeMatch = codeString.match(/<([^;]+);/);
        
        if (!codeMatch) continue;
        
        const codeType = codeMatch[1].trim();
        
        // Rule 1: Check required format parameters
        // if (formatRequiredCodes.has(codeType)) {
        //     const hasTopBorder = /topborder\s*=/.test(codeString.toLowerCase());
        //     const hasFormat = /format\s*=/.test(codeString.toLowerCase());
        //     const hasBold = /bold\s*=/.test(codeString.toLowerCase());
        //     const hasIndent = /indent\s*=/.test(codeString.toLowerCase());
            
        //     // const missingParams = [];
        //     // if (!hasTopBorder) missingParams.push('topborder');
        //     // if (!hasFormat) missingParams.push('format');
        //     // if (!hasBold) missingParams.push('bold');
        //     // if (!hasIndent) missingParams.push('indent');
            
        //     if (missingParams.length > 0) {
        //         // errors.push(`[FERR001] Format validation: ${codeType} code missing required parameters: ${missingParams.join(', ')} - ${codeString}`);
        //     }
        // }
        
        // Check LABELH1/H2/H3 column 2 must end with colon
        if (codeType === 'LABELH1' || codeType === 'LABELH2' || codeType === 'LABELH3') {
            const rowMatch = codeString.match(/row1\s*=\s*"([^"]*)"/);
            if (rowMatch) {
                const rowContent = rowMatch[1];
                const parts = rowContent.split('|');
                if (parts.length >= 2) {
                    const column2 = parts[1].trim();
                    if (column2 && !column2.endsWith(':')) {
                        errors.push(`[FERR002] Format validation: ${codeType} code column 2 must end with colon - ${codeString} should have "${column2}:" in column 2`);
                    }
                }
            }
        }
        
        // Check LABELH3 must be followed by indent="2"
        if (codeType === 'LABELH3' && i < inputCodeStrings.length - 1) {
            const nextCodeString = inputCodeStrings[i + 1];
            const nextCodeMatch = nextCodeString.match(/<([^;]+);/);
            
            if (nextCodeMatch) {
                const nextCodeType = nextCodeMatch[1].trim();
                // Skip BR codes when checking the next code
                let nextNonBRIndex = i + 1;
                while (nextNonBRIndex < inputCodeStrings.length && 
                       inputCodeStrings[nextNonBRIndex].match(/<BR[>;]/)) {
                    nextNonBRIndex++;
                }
                
                if (nextNonBRIndex < inputCodeStrings.length) {
                    const nextNonBRCode = inputCodeStrings[nextNonBRIndex];
                    const hasIndent2 = /indent\s*=\s*["']?2["']?/i.test(nextNonBRCode);
                    
                    if (!hasIndent2) {
                        errors.push(`[FERR003] Format validation: LABELH3 must be followed by a code with indent="2". Use LABELH2 instead - ${codeString} followed by ${nextNonBRCode}`);
                    }
                }
            }
        }
        
        // Check MULT, SUBTRACT, SUM, SUBTOTAL codes must have indent="1" and bold="true"
        if (codeType.startsWith('MULT') || codeType.startsWith('SUBTRACT') || 
            codeType.startsWith('SUM') || codeType.startsWith('SUBTOTAL')) {
            
            // Check indent="1"
            const hasIndent1 = /indent\s*=\s*["']?1["']?/i.test(codeString);
            const indentMatch = codeString.match(/indent\s*=\s*["']?(\d+)["']?/i);
            const indentValue = indentMatch ? indentMatch[1] : null;
            
            // Check bold="true"
            const hasBoldTrue = /bold\s*=\s*["']?true["']?/i.test(codeString);
            const boldMatch = codeString.match(/bold\s*=\s*["']?(true|false)["']?/i);
            const boldValue = boldMatch ? boldMatch[1].toLowerCase() : null;
            
            const issues = [];
            
            if (!hasIndent1) {
                if (indentValue) {
                    issues.push(`has indent="${indentValue}" instead of indent="1"`);
                } else {
                    issues.push('missing indent="1"');
                }
            }
            
            if (!hasBoldTrue) {
                if (boldValue === 'false') {
                    issues.push('has bold="false" instead of bold="true"');
                } else {
                    issues.push('missing bold="true"');
                }
            }
            
            if (issues.length > 0) {
                errors.push(`[FERR008] Format validation: ${codeType} code must have indent="1" and bold="true" - ${codeString} ${issues.join(' and ')}`);
            }
        }
        
        // Check rows with column 2 beginning with "Total"
        const rowMatch = codeString.match(/row\d+\s*=\s*"([^"]*)"/);
        if (rowMatch) {
            const rowContent = rowMatch[1];
            const parts = rowContent.split('|');
            if (parts.length >= 2) {
                const column2 = parts[1].trim();
                if (column2.startsWith('Total')) {
                    // Check if previous code is BR
                    const prevIsBR = i > 0 && inputCodeStrings[i - 1].match(/<BR[>;]/);
                    
                    // Check required parameters
                    const hasBoldTrue = /bold\s*=\s*["']?true["']?/i.test(codeString);
                    const hasIndent1 = /indent\s*=\s*["']?1["']?/i.test(codeString);
                    const hasTopBorderTrue = /topborder\s*=\s*["']?true["']?/i.test(codeString);
                    
                    const issues = [];
                    
                    // Both cases need bold="true" and indent="1"
                    if (!hasBoldTrue) {
                        issues.push('missing or incorrect bold="true"');
                    }
                    if (!hasIndent1) {
                        issues.push('missing or incorrect indent="1"');
                    }
                    
                    // Check topborder based on whether it follows BR
                    if (prevIsBR) {
                        // Should NOT have topborder="True" after BR
                        if (hasTopBorderTrue) {
                            issues.push('should not have topborder="True" when following BR code');
                        }
                    } else {
                        // Should have topborder="True" when not after BR
                        if (!hasTopBorderTrue) {
                            issues.push('missing topborder="True" (required when not following BR code)');
                        }
                    }
                    
                    if (issues.length > 0) {
                        const expectedParams = prevIsBR 
                            ? 'bold="true" and indent="1"' 
                            : 'bold="true", indent="1", and topborder="True"';
                        errors.push(`[FERR009] Format validation: Row with column 2 beginning with "Total" must have ${expectedParams} - ${codeString} ${issues.join(', ')}`);
                    }
                }
            }
        }
        
        // Rules 2, 3, 4: Check adjacency rules
        if (i > 0) {
            const prevCodeString = inputCodeStrings[i - 1];
            const prevCodeMatch = prevCodeString.match(/<([^;]+);/);
            const prevCodeType = prevCodeMatch ? prevCodeMatch[1].trim() : '';
            
            // Rule 2: Check adjacent topborder or bold
            const currTopBorder = /topborder\s*=\s*["']?true["']?/i.test(codeString);
            const prevTopBorder = /topborder\s*=\s*["']?true["']?/i.test(prevCodeString);
            const currBold = /bold\s*=\s*["']?true["']?/i.test(codeString);
            const prevBold = /bold\s*=\s*["']?true["']?/i.test(prevCodeString);
            
            if (currTopBorder && prevTopBorder) {
                errors.push(`[FERR004] Format validation: Adjacent codes both have topborder="true" - ${prevCodeString} followed by ${codeString}`);
            }
            
            if (currBold && prevBold) {
                errors.push(`[FERR005] Format validation: Adjacent codes both have bold="true" - ${prevCodeString} followed by ${codeString}`);
            }
            
            // Rule 3: Check adjacent indent="1"
            const currIndent1 = /indent\s*=\s*["']?1["']?/i.test(codeString);
            const prevIndent1 = /indent\s*=\s*["']?1["']?/i.test(prevCodeString);
            
            if (currIndent1 && prevIndent1) {
                errors.push(`[FERR006] Format validation: Adjacent codes both have indent="1" - ${prevCodeString} followed by ${codeString}`);
            }
            
            // Rule 4: BR followed by indent="2"
            if (prevCodeType === 'BR') {
                const currIndent2 = /indent\s*=\s*["']?2["']?/i.test(codeString);
                if (currIndent2) {
                    errors.push(`[FERR007] Format validation: BR code followed by code with indent="2" - ${prevCodeString} followed by ${codeString}`);
                }
            }
        }
    }

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
    } // <-- This is the end of the second pass loop

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

    // Convert array of errors to a single string with line breaks
    return errors.join('\n');
}

// Modified export function
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

// >>> ADDED: New validation function for Run Codes flow (returns array)
export async function validateCodeStringsForRun(inputCodeStrings) {
    const errors = [];
    const tabLabels = new Set();
    const rowValues = new Set();
    const codeTypes = new Set();
    const tabRowDrivers = new Map(); // Track row drivers per TAB
    
    // >>> ADDED: Remove line breaks within angle brackets before processing
    inputCodeStrings = inputCodeStrings.map(str => {
        // Replace line breaks within < > brackets
        return str.replace(/<[^>]*>/g, match => {
            // Remove all types of line breaks within the match
            return match.replace(/[\r\n]+/g, ' ');
        });
    });
    console.log("[ValidateForRun] Preprocessed code strings (line breaks removed within <>):", inputCodeStrings);
    // <<< END ADDED
    
    // Clean up input strings by removing anything outside angle brackets
    inputCodeStrings = inputCodeStrings.map(str => {
        const match = str.match(/<[^>]+>/);
        return match ? match[0] : str;
    });

    console.log("[ValidateForRun] CleanedInput Code Strings:", inputCodeStrings);

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
        return errors; // Return array on critical error
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
        'IS: revenue',
        'IS: direct costs',
        'IS: corporate overhead',
        'IS: D&A',
        'IS: interest',
        'IS: Other Income',
        'IS: Net Income',
        'BS: current assets',
        'BS: fixed assets',
        'BS: current liabilities',
        'BS: LT Liabilities',
        'BS: equity',
        'CF: WC',
        'CF: Non-cash',
        'CF: CFI',
        'CF: CFF'
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
        const rowMatches = codeString.match(/row\d+\s*=\s*"([^"]*)"/g);
        if (rowMatches) {
            rowMatches.forEach(match => {
                const rowParam = match.match(/(row\d+)\s*=\s*"([^"]*)"/);
                const rowName = rowParam[1];
                const rowContent = rowParam[2];
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
                            if (!validFinancialCodes.has(columnC)) {
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
                    
                    parts.forEach(part => {
                        const trimmedPart = part.trim();
                        if (trimmedPart.startsWith('*')) {
                            const afterAsterisk = trimmedPart.substring(1).trim();
                            if (afterAsterisk) {
                                rowValues.add(afterAsterisk);
                            }
                        } else if (trimmedPart) {
                            rowValues.add(trimmedPart);
                        }
                    });
                }
            });
        }
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
            
            errors.push(`[LERR012] Duplicate financial statement item "${itemName}" found in ${occurrences.length} locations across tabs: ${tabList}`);
            
            // Add details about each occurrence
            occurrences.forEach(loc => {
                errors.push(`  - In ${loc.codeString.substring(0, 100)}... at ${loc.rowName} (driver: ${loc.driver || 'none'})`);
            });
        }
    });

    // Format Validation Section
    const formatRequiredCodes = new Set([
        'FINANCIALS-S', 'MULT3-S', 'DIVIDE2-S', 'SUBTRACT2-S', 'SUBTOTAL2-S', 
        'SUBTOTAL3-S', 'AVGMULT3-S', 'ANNUALIZE-S', 'DEANNUALIZE-S', 
        'AVGDEANNUALIZE2-S', 'DIRECT-S', 'CHANGE-S', 'INCREASE-S', 'DECREASE-S', 
        'GROWTH-S', 'OFFSETCOLUMN-S', 'OFFSET2-S', 'SUM2-S', 'DISCOUNT2-S', 
        'CONST-E', 'SPREAD-E', 'ENDPOINT-E', 'GROWTH-E'
    ]);

    // Check format parameters and adjacency rules
    for (let i = 0; i < inputCodeStrings.length; i++) {
        const codeString = inputCodeStrings[i];
        const codeMatch = codeString.match(/<([^;]+);/);
        
        if (!codeMatch) continue;
        
        const codeType = codeMatch[1].trim();
        
        // Rule 1: Check required format parameters
        // if (formatRequiredCodes.has(codeType)) {
        //     const hasTopBorder = /topborder\s*=/.test(codeString.toLowerCase());
        //     const hasFormat = /format\s*=/.test(codeString.toLowerCase());
        //     const hasBold = /bold\s*=/.test(codeString.toLowerCase());
        //     const hasIndent = /indent\s*=/.test(codeString.toLowerCase());
            
        //     // const missingParams = [];
        //     // if (!hasTopBorder) missingParams.push('topborder');
        //     // if (!hasFormat) missingParams.push('format');
        //     // if (!hasBold) missingParams.push('bold');
        //     // if (!hasIndent) missingParams.push('indent');
            
        //     if (missingParams.length > 0) {
        //         // errors.push(`[FERR001] Format validation: ${codeType} code missing required parameters: ${missingParams.join(', ')} - ${codeString}`);
        //     }
        // }
        
        // Check LABELH1/H2/H3 column 2 must end with colon
        if (codeType === 'LABELH1' || codeType === 'LABELH2' || codeType === 'LABELH3') {
            const rowMatch = codeString.match(/row1\s*=\s*"([^"]*)"/);
            if (rowMatch) {
                const rowContent = rowMatch[1];
                const parts = rowContent.split('|');
                if (parts.length >= 2) {
                    const column2 = parts[1].trim();
                    if (column2 && !column2.endsWith(':')) {
                        errors.push(`[FERR002] Format validation: ${codeType} code column 2 must end with colon - ${codeString} should have "${column2}:" in column 2`);
                    }
                }
            }
        }
        
        // Check LABELH3 must be followed by indent="2"
        if (codeType === 'LABELH3' && i < inputCodeStrings.length - 1) {
            const nextCodeString = inputCodeStrings[i + 1];
            const nextCodeMatch = nextCodeString.match(/<([^;]+);/);
            
            if (nextCodeMatch) {
                const nextCodeType = nextCodeMatch[1].trim();
                // Skip BR codes when checking the next code
                let nextNonBRIndex = i + 1;
                while (nextNonBRIndex < inputCodeStrings.length && 
                       inputCodeStrings[nextNonBRIndex].match(/<BR[>;]/)) {
                    nextNonBRIndex++;
                }
                
                if (nextNonBRIndex < inputCodeStrings.length) {
                    const nextNonBRCode = inputCodeStrings[nextNonBRIndex];
                    const hasIndent2 = /indent\s*=\s*["']?2["']?/i.test(nextNonBRCode);
                    
                    if (!hasIndent2) {
                        errors.push(`[FERR003] Format validation: LABELH3 must be followed by a code with indent="2". Use LABELH2 instead - ${codeString} followed by ${nextNonBRCode}`);
                    }
                }
            }
        }
        
        // Check MULT, SUBTRACT, SUM, SUBTOTAL codes must have indent="1" and bold="true"
        if (codeType.startsWith('MULT') || codeType.startsWith('SUBTRACT') || 
            codeType.startsWith('SUM') || codeType.startsWith('SUBTOTAL')) {
            
            // Check indent="1"
            const hasIndent1 = /indent\s*=\s*["']?1["']?/i.test(codeString);
            const indentMatch = codeString.match(/indent\s*=\s*["']?(\d+)["']?/i);
            const indentValue = indentMatch ? indentMatch[1] : null;
            
            // Check bold="true"
            const hasBoldTrue = /bold\s*=\s*["']?true["']?/i.test(codeString);
            const boldMatch = codeString.match(/bold\s*=\s*["']?(true|false)["']?/i);
            const boldValue = boldMatch ? boldMatch[1].toLowerCase() : null;
            
            const issues = [];
            
            if (!hasIndent1) {
                if (indentValue) {
                    issues.push(`has indent="${indentValue}" instead of indent="1"`);
                } else {
                    issues.push('missing indent="1"');
                }
            }
            
            if (!hasBoldTrue) {
                if (boldValue === 'false') {
                    issues.push('has bold="false" instead of bold="true"');
                } else {
                    issues.push('missing bold="true"');
                }
            }
            
            if (issues.length > 0) {
                errors.push(`[FERR008] Format validation: ${codeType} code must have indent="1" and bold="true" - ${codeString} ${issues.join(' and ')}`);
            }
        }
        
        // Check rows with column 2 beginning with "Total"
        const rowMatch = codeString.match(/row\d+\s*=\s*"([^"]*)"/);
        if (rowMatch) {
            const rowContent = rowMatch[1];
            const parts = rowContent.split('|');
            if (parts.length >= 2) {
                const column2 = parts[1].trim();
                if (column2.startsWith('Total')) {
                    // Check if previous code is BR
                    const prevIsBR = i > 0 && inputCodeStrings[i - 1].match(/<BR[>;]/);
                    
                    // Check required parameters
                    const hasBoldTrue = /bold\s*=\s*["']?true["']?/i.test(codeString);
                    const hasIndent1 = /indent\s*=\s*["']?1["']?/i.test(codeString);
                    const hasTopBorderTrue = /topborder\s*=\s*["']?true["']?/i.test(codeString);
                    
                    const issues = [];
                    
                    // Both cases need bold="true" and indent="1"
                    if (!hasBoldTrue) {
                        issues.push('missing or incorrect bold="true"');
                    }
                    if (!hasIndent1) {
                        issues.push('missing or incorrect indent="1"');
                    }
                    
                    // Check topborder based on whether it follows BR
                    if (prevIsBR) {
                        // Should NOT have topborder="True" after BR
                        if (hasTopBorderTrue) {
                            issues.push('should not have topborder="True" when following BR code');
                        }
                    } else {
                        // Should have topborder="True" when not after BR
                        if (!hasTopBorderTrue) {
                            issues.push('missing topborder="True" (required when not following BR code)');
                        }
                    }
                    
                    if (issues.length > 0) {
                        const expectedParams = prevIsBR 
                            ? 'bold="true" and indent="1"' 
                            : 'bold="true", indent="1", and topborder="True"';
                        errors.push(`[FERR009] Format validation: Row with column 2 beginning with "Total" must have ${expectedParams} - ${codeString} ${issues.join(', ')}`);
                    }
                }
            }
        }
        
        // Rules 2, 3, 4: Check adjacency rules
        if (i > 0) {
            const prevCodeString = inputCodeStrings[i - 1];
            const prevCodeMatch = prevCodeString.match(/<([^;]+);/);
            const prevCodeType = prevCodeMatch ? prevCodeMatch[1].trim() : '';
            
            // Rule 2: Check adjacent topborder or bold
            const currTopBorder = /topborder\s*=\s*["']?true["']?/i.test(codeString);
            const prevTopBorder = /topborder\s*=\s*["']?true["']?/i.test(prevCodeString);
            const currBold = /bold\s*=\s*["']?true["']?/i.test(codeString);
            const prevBold = /bold\s*=\s*["']?true["']?/i.test(prevCodeString);
            
            if (currTopBorder && prevTopBorder) {
                errors.push(`[FERR004] Format validation: Adjacent codes both have topborder="true" - ${prevCodeString} followed by ${codeString}`);
            }
            
            if (currBold && prevBold) {
                errors.push(`[FERR005] Format validation: Adjacent codes both have bold="true" - ${prevCodeString} followed by ${codeString}`);
            }
            
            // Rule 3: Check adjacent indent="1"
            const currIndent1 = /indent\s*=\s*["']?1["']?/i.test(codeString);
            const prevIndent1 = /indent\s*=\s*["']?1["']?/i.test(prevCodeString);
            
            if (currIndent1 && prevIndent1) {
                errors.push(`[FERR006] Format validation: Adjacent codes both have indent="1" - ${prevCodeString} followed by ${codeString}`);
            }
            
            // Rule 4: BR followed by indent="2"
            if (prevCodeType === 'BR') {
                const currIndent2 = /indent\s*=\s*["']?2["']?/i.test(codeString);
                if (currIndent2) {
                    errors.push(`[FERR007] Format validation: BR code followed by code with indent="2" - ${prevCodeString} followed by ${codeString}`);
                }
            }
        }
    }

    // Second pass: detailed validation
    for (const codeString of inputCodeStrings) {
        if (codeString === '<BR>') {
            continue;
        }
        const codeMatch = codeString.match(/<([^;]+);/);
        if (!codeMatch) {
            errors.push(`[LERR004] Cannot extract code type from: ${codeString}`);
            continue;
        }
        const codeType = codeMatch[1].trim();
        if (!validCodes.has(codeType)) {
            errors.push(`[LERR005] Invalid code type: "${codeType}" not found in valid codes list - ${codeString}`);
        }
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

    // Third pass: validate driver references
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

    // Key Change: Return the array of errors directly
    return errors;
}
// <<< END ADDED FUNCTION

