const fs = require('fs');
const path = require('path');

// File paths
const inputFile = path.join(__dirname, '../../../Dropbox/B - Freelance/C_Projectify/VanPC/Chain Reaction/Chain Reaction Model_6.4.25 v3.txt');
const outputFile = path.join(__dirname, '../../../Downloads/TrainingDataUTF8.jsonl');

// System message for all training entries
const systemMessage = "You are a financial model planning assistant. You are tasked with programming the client's request with a custom financial model coding lanaguage you have been trained on";

function cleanString(str) {
    // Only remove wrapping quotes if the ENTIRE string is wrapped in quotes
    // and doesn't contain the financial modeling syntax
    const trimmed = str.trim();
    
    // Check if the entire string is wrapped in quotes AND doesn't contain financial modeling syntax
    if (trimmed.startsWith('"') && trimmed.endsWith('"') && 
        !trimmed.includes('=""') && // Financial modeling attribute syntax
        !trimmed.includes('|') && // Financial modeling data syntax
        !trimmed.includes('row1 =') && // Financial modeling row syntax
        !trimmed.includes('BR;') && // Financial modeling break syntax
        !trimmed.includes('LABELH') && // Financial modeling label syntax
        !trimmed.includes('CONST-E') && // Financial modeling constants
        !trimmed.includes('FORMULA-S') && // Financial modeling formulas
        !trimmed.includes('SPREAD-E')) { // Financial modeling spreads
        
        // This is a simple quoted string, remove outer quotes
        return trimmed.slice(1, -1);
    }
    
    // For everything else (including financial modeling syntax), keep as-is
    // The quotes will be properly escaped when we JSON.stringify
    return trimmed;
}

function convertToJsonl() {
    try {
        console.log('Reading input file...');
        const data = fs.readFileSync(inputFile, 'utf8');
        
        // Split into lines and remove the header
        const lines = data.split('\n').slice(1);
        
        // Read existing JSONL data if it exists
        let existingEntries = [];
        // Start fresh to ensure all entries have correct quote formatting
        console.log('Starting fresh to ensure proper quote formatting for all entries...');
        
        // Process each line and convert to JSONL format
        const newEntries = [];
        let processedCount = 0;
        let skippedCount = 0;
        
        for (let i = 0; i < lines.length; i++) {
            const line = lines[i].trim();
            
            // Skip empty lines
            if (!line) continue;
            
            // Check if line contains @@@ delimiter
            if (line.includes('@@@')) {
                // Find the position of @@@
                const delimiterIndex = line.indexOf('@@@');
                
                if (delimiterIndex > 0 && delimiterIndex < line.length - 3) {
                    let userInput = line.substring(0, delimiterIndex).trim();
                    let assistantOutput = line.substring(delimiterIndex + 3).trim();
                    
                    // Clean the strings
                    userInput = cleanString(userInput);
                    assistantOutput = cleanString(assistantOutput);
                    
                    // Skip if either is empty or too short
                    if (!userInput || !assistantOutput || userInput.length < 5 || assistantOutput.length < 10) {
                        skippedCount++;
                        console.log(`Skipping line ${i + 2}: Input or output too short/empty`);
                        continue;
                    }
                    
                    // Skip obvious formatting issues
                    if (userInput.includes('Input Prompt,Paraphrased Input') || 
                        assistantOutput.includes('Input Prompt,Paraphrased Input')) {
                        skippedCount++;
                        console.log(`Skipping line ${i + 2}: Header/formatting line detected`);
                        continue;
                    }
                    
                    // Create the JSONL entry
                    const entry = {
                        messages: [
                            {
                                role: "system",
                                content: systemMessage
                            },
                            {
                                role: "user",
                                content: userInput
                            },
                            {
                                role: "assistant",
                                content: assistantOutput
                            }
                        ]
                    };
                    
                    newEntries.push(entry);
                    processedCount++;
                    
                    // Log progress every 50 entries
                    if (processedCount % 50 === 0) {
                        console.log(`Processed ${processedCount} entries...`);
                    }
                }
            }
        }
        
        // Combine existing and new entries
        const allEntries = [...newEntries];
        
        // Write to JSONL file
        console.log('Writing to JSONL file...');
        const jsonlContent = allEntries.map(entry => JSON.stringify(entry)).join('\n');
        fs.writeFileSync(outputFile, jsonlContent, 'utf8');
        
        console.log(`\n✅ Conversion completed successfully!`);
        console.log(`📊 Statistics:`);
        console.log(`   - Lines processed: ${lines.length}`);
        console.log(`   - Training examples converted: ${newEntries.length}`);
        console.log(`   - Skipped entries: ${skippedCount}`);
        console.log(`   - Total entries in file: ${allEntries.length}`);
        console.log(`📁 Output file: ${outputFile}`);
        
        // Show a few sample entries
        if (newEntries.length > 0) {
            console.log(`\n📝 Sample converted entries:`);
            for (let i = 0; i < Math.min(3, newEntries.length); i++) {
                const entry = newEntries[i];
                console.log(`\n   Entry ${i + 1}:`);
                console.log(`   User: "${entry.messages[1].content.substring(0, 60)}..."`);
                console.log(`   Assistant: "${entry.messages[2].content.substring(0, 60)}..."`);
            }
        }
        
    } catch (error) {
        console.error('❌ Error during conversion:', error.message);
        
        if (error.code === 'ENOENT') {
            if (error.path.includes('Chain Reaction')) {
                console.error('💡 The input file was not found. Please check the path:');
                console.error(`   ${inputFile}`);
            } else {
                console.error('💡 Could not access the output directory. Please check permissions.');
            }
        } else if (error instanceof SyntaxError) {
            console.error('💡 There was an issue parsing the JSON data.');
        } else {
            console.error('💡 Please check file paths and permissions.');
        }
        
        process.exit(1);
    }
}

// Run the conversion
console.log('🚀 Starting enhanced training data conversion...');
console.log(`📖 Input: Chain Reaction Model_6.4.25 v3.txt`);
console.log(`💾 Output: TrainingDataUTF8.jsonl`);
console.log('');

convertToJsonl(); 