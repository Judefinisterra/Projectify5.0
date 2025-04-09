/**
 * CodeCollection.js
 * Functions for processing and managing code collections
 */

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
            const paramMatches = paramsString.matchAll(/(\w+)\s*=\s*"([^"]*)"/g);
            
            for (const match of paramMatches) {
                const paramName = match[1].trim();
                const paramValue = match[2].trim();
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
            
            for (const [key, value] of Object.entries(code.params)) {
                result += `  ${key}: ${value}\n`;
            }
            
            result += "\n";
        });
        
        return result;
    } catch (error) {
        console.error("Error in exportCodeCollectionToText:", error);
        throw error;
    }
} 