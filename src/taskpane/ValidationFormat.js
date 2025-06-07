// ValidationFormat.js - Handles format validation errors (FERR codes)

export async function validateFormatErrors(inputText) {
    const errors = [];
    
    // Extract all code strings using regex pattern /<[^>]+>/g
    let inputCodeStrings = [];
    if (typeof inputText === 'string') {
        const codeStringMatches = inputText.match(/<[^>]+>/g);
        inputCodeStrings = codeStringMatches || [];
    } else if (Array.isArray(inputText)) {
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
        console.log("No code strings found in input for format validation");
        return errors;
    }
    
    console.log(`Found ${inputCodeStrings.length} code strings for format validation`);
    
    // Remove line breaks within angle brackets before processing
    inputCodeStrings = inputCodeStrings.map(str => {
        return str.replace(/<[^>]*>/g, match => {
            return match.replace(/[\r\n]+/g, ' ');
        });
    });
    
    // Clean up input strings by removing anything outside angle brackets
    inputCodeStrings = inputCodeStrings.map(str => {
        const match = str.match(/<[^>]+>/);
        return match ? match[0] : str;
    });

    // Check format parameters and adjacency rules
    for (let i = 0; i < inputCodeStrings.length; i++) {
        const codeString = inputCodeStrings[i];
        const codeMatch = codeString.match(/<([^;]+);/);
        
        if (!codeMatch) continue;
        
        const codeType = codeMatch[1].trim();
        
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
        

        
        // Check rows with column 2 beginning with "Total"
        const rowMatch = codeString.match(/row\d+\s*=\s*"([^"]*)"/);
        if (rowMatch) {
            const rowContent = rowMatch[1];
            const parts = rowContent.split('|');
            if (parts.length >= 2) {
                const column2 = parts[1].trim();
                if (column2.startsWith('Total')) {
                    const prevIsBR = i > 0 && inputCodeStrings[i - 1].match(/<BR[>;]/);
                    
                    const hasBoldTrue = /bold\s*=\s*["']?true["']?/i.test(codeString);
                    const hasIndent1 = /indent\s*=\s*["']?1["']?/i.test(codeString);
                    const hasTopBorderTrue = /topborder\s*=\s*["']?true["']?/i.test(codeString);
                    
                    const issues = [];
                    
                    if (!hasBoldTrue) {
                        issues.push('missing or incorrect bold="true"');
                    }
                    if (!hasIndent1) {
                        issues.push('missing or incorrect indent="1"');
                    }
                    
                    if (prevIsBR) {
                        if (hasTopBorderTrue) {
                            issues.push('should not have topborder="True" when following BR code');
                        }
                    } else {
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
        
        // Check adjacency rules
        if (i > 0) {
            const prevCodeString = inputCodeStrings[i - 1];
            const prevCodeMatch = prevCodeString.match(/<([^;]+);/);
            const prevCodeType = prevCodeMatch ? prevCodeMatch[1].trim() : '';
            
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
            
            const currIndent1 = /indent\s*=\s*["']?1["']?/i.test(codeString);
            const prevIndent1 = /indent\s*=\s*["']?1["']?/i.test(prevCodeString);
            
            if (currIndent1 && prevIndent1) {
                errors.push(`[FERR006] Format validation: Adjacent codes both have indent="1" - ${prevCodeString} followed by ${codeString}`);
            }
            
            if (prevCodeType === 'BR') {
                const currIndent2 = /indent\s*=\s*["']?2["']?/i.test(codeString);
                if (currIndent2) {
                    errors.push(`[FERR007] Format validation: BR code followed by code with indent="2" - ${prevCodeString} followed by ${codeString}`);
                }
            }
        }
    }

    return errors;
}

export async function runFormatValidation(inputStrings) {
    try {
        const cleanedStrings = inputStrings
            .map(line => line.trim())
            .filter(line => line.length > 0)
            .map(line => line.replace(/^'|',$|,$|^"|"$/g, ''));

        const codeStrings = cleanedStrings.join(' ').match(/<[^>]+>/g) || [];

        if (codeStrings.length === 0) {
            console.error('Format validation run with no codestrings - input is empty');
            return ['No code strings found to validate'];
        }

        const validationErrors = await validateFormatErrors(codeStrings);
        return validationErrors;
    } catch (error) {
        console.error('Format validation failed:', error);
        return [`Format validation failed: ${error.message}`];
    }
} 