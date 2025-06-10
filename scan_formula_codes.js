const fs = require('fs');
const path = require('path');

// File paths
const formulaCodesFile = path.join(__dirname, 'src/prompts/FORMULASCODES.txt');
const encoderSystemFile = path.join(__dirname, 'src/prompts/Encoder_System.txt');
const logicCheckerFile = path.join(__dirname, 'src/prompts/LogicCheckerGPT.txt');

// Training data file paths to check
const trainingDataPaths = [
    path.join(__dirname, '../../../Downloads/TrainingDataUTF8.jsonl'),
    path.join(__dirname, '../../../Dropbox/B - Freelance/C_Projectify/VanPC/Training Data/Fine Tuning Training Data/TrainingDataUTF8.jsonl'),
    path.join(__dirname, '../../../Dropbox/B - Freelance/C_Projectify/VanPC/Training Data/Fine Tuning Training Data/TrainingDataUTF9.jsonl')
];

/**
 * Parse FORMULASCODES.txt to extract code definitions
 */
function parseFormulaCodes() {
    try {
        const content = fs.readFileSync(formulaCodesFile, 'utf8');
        const lines = content.split('\n');
        
        const codes = {};
        
        for (const line of lines) {
            const trimmedLine = line.trim();
            if (!trimmedLine || trimmedLine.startsWith('FORMULA-S custom functions:') || trimmedLine.startsWith('Codes:')) {
                continue;
            }
            
            // Extract code from patterns like "rd{row driver}:" or "BEG(columndriver):" or "FAPPE-S:"
            let codeMatch;
            let fullDescription = trimmedLine;
            
            // Check for patterns like "rd{...}:" or "cd{...}:"
            if ((codeMatch = trimmedLine.match(/^([a-zA-Z-]+)\{[^}]*\}:/))) {
                const code = codeMatch[1];
                codes[code] = fullDescription;
            }
            // Check for patterns like "BEG(...):" or "SPREADDATES(...):"
            else if ((codeMatch = trimmedLine.match(/^([a-zA-Z-]+)\([^)]*\):/))) {
                const code = codeMatch[1];
                codes[code] = fullDescription;
            }
            // Check for patterns like "FAPPE-S:"
            else if ((codeMatch = trimmedLine.match(/^([a-zA-Z-]+):/))) {
                const code = codeMatch[1];
                codes[code] = fullDescription;
            }
        }
        
        console.log(`📋 Parsed ${Object.keys(codes).length} formula codes:`, Object.keys(codes));
        return codes;
        
    } catch (error) {
        console.error('❌ Error reading FORMULASCODES.txt:', error.message);
        return {};
    }
}

/**
 * Scan training data files for customformula parameters
 */
function scanTrainingDataForFormulaCodes(formulaCodes) {
    const usedCodes = new Set();
    const codeKeys = Object.keys(formulaCodes);
    
    for (const trainingDataPath of trainingDataPaths) {
        try {
            if (!fs.existsSync(trainingDataPath)) {
                console.log(`⚠️  Training data file not found: ${trainingDataPath}`);
                continue;
            }
            
            console.log(`🔍 Scanning: ${path.basename(trainingDataPath)}`);
            const content = fs.readFileSync(trainingDataPath, 'utf8');
            const lines = content.split('\n');
            
            let scannedEntries = 0;
            let customFormulaCount = 0;
            
            for (const line of lines) {
                if (!line.trim()) continue;
                
                try {
                    const entry = JSON.parse(line);
                    scannedEntries++;
                    
                    // Check all messages in the entry
                    if (entry.messages) {
                        for (const message of entry.messages) {
                            if (message.content) {
                                // Look for customformula parameters
                                const customFormulaMatches = message.content.match(/customformula\d*\s*=\s*"([^"]*)"/g);
                                
                                if (customFormulaMatches) {
                                    customFormulaCount++;
                                    
                                    for (const match of customFormulaMatches) {
                                        const formulaContent = match.match(/customformula\d*\s*=\s*"([^"]*)"/)[1];
                                        
                                        // Check if any of our codes are used in this formula
                                        for (const code of codeKeys) {
                                            // Check for direct usage like "rd{" or "BEG(" or "FAPPE-S"
                                            if (formulaContent.includes(code + '{') || 
                                                formulaContent.includes(code + '(') || 
                                                formulaContent.includes(code)) {
                                                usedCodes.add(code);
                                                console.log(`✅ Found "${code}" in: ${formulaContent.substring(0, 50)}...`);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                } catch (parseError) {
                    // Skip malformed JSON lines
                    continue;
                }
            }
            
            console.log(`   📊 Scanned ${scannedEntries} entries, found ${customFormulaCount} with customformula parameters`);
            
        } catch (error) {
            console.error(`❌ Error scanning ${trainingDataPath}:`, error.message);
        }
    }
    
    return Array.from(usedCodes);
}

/**
 * Update encoder system prompt with used formula code descriptions
 */
function updateEncoderSystem(usedCodes, formulaCodes) {
    try {
        let content = fs.readFileSync(encoderSystemFile, 'utf8');
        
        // Find the FORMULA-S Context section and add the used codes there
        const formulaSContextIndex = content.indexOf('FORMULA-S Context:');
        if (formulaSContextIndex === -1) {
            console.error('❌ Could not find "FORMULA-S Context:" section in Encoder_System.txt');
            return false;
        }
        
        // Build the additional content
        let additionalContent = '\n\nFORMULA-S Custom Functions (found in training data):\n';
        for (const code of usedCodes) {
            if (formulaCodes[code]) {
                additionalContent += `${formulaCodes[code]}\n`;
            }
        }
        
        // Find the end of the FORMULA-S Context section (next section starts with a capital letter section header)
        const contextEndMatch = content.substring(formulaSContextIndex).match(/\n\n[A-Z][a-zA-Z\s]+:\s*\n/);
        let insertIndex;
        
        if (contextEndMatch) {
            insertIndex = formulaSContextIndex + contextEndMatch.index;
        } else {
            // If no next section found, add at the end
            insertIndex = content.length;
        }
        
        // Check if we already have this content to avoid duplicates
        if (!content.includes('FORMULA-S Custom Functions (found in training data):')) {
            content = content.substring(0, insertIndex) + additionalContent + content.substring(insertIndex);
            
            fs.writeFileSync(encoderSystemFile, content, 'utf8');
            console.log('✅ Updated Encoder_System.txt with formula code descriptions');
            return true;
        } else {
            console.log('ℹ️  Encoder_System.txt already contains formula code descriptions');
            return true;
        }
        
    } catch (error) {
        console.error('❌ Error updating Encoder_System.txt:', error.message);
        return false;
    }
}

/**
 * Update logic checker prompt with used formula code descriptions
 */
function updateLogicChecker(usedCodes, formulaCodes) {
    try {
        let content = fs.readFileSync(logicCheckerFile, 'utf8');
        
        // Add the formula codes before the final output section
        const finalOutputIndex = content.lastIndexOf('12) Final Output:');
        if (finalOutputIndex === -1) {
            console.error('❌ Could not find "Final Output:" section in LogicCheckerGPT.txt');
            return false;
        }
        
        // Build the additional content
        let additionalContent = '\nFORMULA-S Custom Functions (found in training data):\n';
        for (const code of usedCodes) {
            if (formulaCodes[code]) {
                additionalContent += `${formulaCodes[code]}\n`;
            }
        }
        additionalContent += '\n';
        
        // Check if we already have this content to avoid duplicates
        if (!content.includes('FORMULA-S Custom Functions (found in training data):')) {
            content = content.substring(0, finalOutputIndex) + additionalContent + content.substring(finalOutputIndex);
            
            fs.writeFileSync(logicCheckerFile, content, 'utf8');
            console.log('✅ Updated LogicCheckerGPT.txt with formula code descriptions');
            return true;
        } else {
            console.log('ℹ️  LogicCheckerGPT.txt already contains formula code descriptions');
            return true;
        }
        
    } catch (error) {
        console.error('❌ Error updating LogicCheckerGPT.txt:', error.message);
        return false;
    }
}

/**
 * Main function
 */
function main() {
    console.log('🚀 Starting Formula Code Scanner...\n');
    
    // Step 1: Parse formula codes from FORMULASCODES.txt
    console.log('📖 Step 1: Parsing FORMULASCODES.txt...');
    const formulaCodes = parseFormulaCodes();
    
    if (Object.keys(formulaCodes).length === 0) {
        console.error('❌ No formula codes found. Exiting.');
        return;
    }
    
    // Step 2: Scan training data for usage of these codes
    console.log('\n🔍 Step 2: Scanning training data...');
    const usedCodes = scanTrainingDataForFormulaCodes(formulaCodes);
    
    console.log(`\n📊 Summary: Found ${usedCodes.length} formula codes used in training data:`);
    usedCodes.forEach(code => {
        console.log(`   - ${code}`);
    });
    
    if (usedCodes.length === 0) {
        console.log('ℹ️  No formula codes found in training data. No updates needed.');
        return;
    }
    
    // Step 3: Update encoder and logic prompts
    console.log('\n📝 Step 3: Updating prompts...');
    const encoderSuccess = updateEncoderSystem(usedCodes, formulaCodes);
    const logicSuccess = updateLogicChecker(usedCodes, formulaCodes);
    
    if (encoderSuccess && logicSuccess) {
        console.log('\n🎉 Formula code scanning and prompt updates completed successfully!');
    } else {
        console.log('\n⚠️  Some updates failed. Please check the error messages above.');
    }
}

// Run the main function
main(); 