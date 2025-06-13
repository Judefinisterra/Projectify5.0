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