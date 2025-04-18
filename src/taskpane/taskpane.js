import { validateCodeStrings } from './Validation.js';

import { populateCodeCollection, exportCodeCollectionToText, runCodes, processAssumptionTabs, collapseGroupingsAndNavigateToFinancials, hideColumnsAndNavigate, handleInsertWorksheetsFromBase64 } from './CodeCollection.js';
// >>> ADDED: Import the new validation function
import { validateCodeStringsForRun } from './Validation.js';
// >>> ADDED: Import the tab string generator function
import { generateTabString } from './IndexWorksheet.js';
// >>> UPDATED: Import structureDatabasequeries from the helper file
import { structureDatabasequeries } from './StructureHelper.js';
// >>> ADDED: Import setAPIKeys function from AIcalls
import { setAPIKeys } from './AIcalls.js';
// >>> ADDED: Import callOpenAI function from AIcalls
import { callOpenAI } from './AIcalls.js';
// >>> ADDED: Import conversation history functions from AIcalls
import { saveConversationHistory, loadConversationHistory } from './AIcalls.js';
// >>> ADDED: Import prompt loading functions from AIcalls
import { loadPromptFromFile, getSystemPromptFromFile } from './AIcalls.js';
// >>> ADDED: Import processPrompt function from AIcalls
import { processPrompt } from './AIcalls.js';
// >>> ADDED: Import createEmbedding function from AIcalls
import { createEmbedding } from './AIcalls.js';
// >>> ADDED: Import queryVectorDB function from AIcalls
import { queryVectorDB } from './AIcalls.js';
// >>> ADDED: Import safeJsonForPrompt function from AIcalls
import { safeJsonForPrompt } from './AIcalls.js';
// >>> ADDED: Import conversation handling and validation functions from AIcalls
// Make sure handleConversation is included here
import { handleFollowUpConversation, handleInitialConversation, handleConversation, validationCorrection } from './AIcalls.js';
// Add the codeStrings variable with the specified content
// REMOVED hardcoded codeStrings variable

import { API_KEYS as configApiKeys } from '../../config.js'; // Assuming config.js exports API_KEYS

//Debugging Toggle
const DEBUG = true;

// Variable to store loaded code strings
let loadedCodeStrings = "";

// Variable to store the parsed code database
let codeDatabase = [];

// API keys storage - initialized by initializeAPIKeys
let INTERNAL_API_KEYS = {
  OPENAI_API_KEY: "",
  PINECONE_API_KEY: ""
};

// >>> MOVED & MODIFIED: Ensure truly global scope for cursor position
var lastEditorCursorPosition = null;

// Function to load the code string database
async function loadCodeDatabase() {
  try {
    console.log("Loading code database...");
    const response = await fetch('https://localhost:3002/assets/codestringDB.txt');
    if (!response.ok) {
      throw new Error(`Failed to load codestringDB.txt: ${response.statusText}`);
    }
    const text = await response.text();
    const lines = text.split(/[\r\n]+/).filter(line => line.trim() !== ''); // Split by lines and remove empty ones

    codeDatabase = lines.map(line => {
      const parts = line.split('\t'); // Assuming tab-separated
      if (parts.length >= 2) {
        return { name: parts[0].trim(), code: parts[1].trim() };
      }
      console.warn(`Skipping malformed line in codestringDB.txt: ${line}`);
      return null;
    }).filter(item => item !== null); // Filter out null entries from malformed lines

    console.log(`Code database loaded successfully with ${codeDatabase.length} entries.`);
    if (DEBUG && codeDatabase.length > 0) {
        console.log("First few code database entries:", codeDatabase.slice(0, 5));
    }

  } catch (error) {
    console.error("Error loading code database:", error);
    showError("Failed to load code database. Search functionality will be unavailable.");
    codeDatabase = []; // Ensure it's empty on error
  }
}

// Function to load API keys from a config file
// This allows the keys to be stored in a separate file that's .gitignored
export async function initializeAPIKeys() {
  try {
    console.log("Initializing API keys from AIcalls.js...");

    // Use keys from imported config.js if available
    if (configApiKeys?.OPENAI_API_KEY) {
        INTERNAL_API_KEYS.OPENAI_API_KEY = configApiKeys.OPENAI_API_KEY;
        console.log("OpenAI API key loaded from config.js");
    } else {
         console.warn("OpenAI API key not found in config.js.");
    }

    if (configApiKeys?.PINECONE_API_KEY) {
        INTERNAL_API_KEYS.PINECONE_API_KEY = configApiKeys.PINECONE_API_KEY;
        console.log("Pinecone API key loaded from config.js");
    } else {
         console.warn("Pinecone API key not found in config.js.");
    }

    // Fallback: try fetching from the old location if config.js didn't provide them
    if (!INTERNAL_API_KEYS.OPENAI_API_KEY || !INTERNAL_API_KEYS.PINECONE_API_KEY) {
        console.log("Attempting fallback API key loading from https://localhost:3002/config.js");
        try {
            const configResponse = await fetch('https://localhost:3002/config.js');
            if (configResponse.ok) {
                const configText = await configResponse.text();
                // Extract keys from the config text using regex
                const openaiKeyMatch = configText.match(/OPENAI_API_KEY\s*=\s*["']([^"']+)["']/);
                const pineconeKeyMatch = configText.match(/PINECONE_API_KEY\s*=\s*["']([^"']+)["']/);

                if (!INTERNAL_API_KEYS.OPENAI_API_KEY && openaiKeyMatch && openaiKeyMatch[1]) {
                    INTERNAL_API_KEYS.OPENAI_API_KEY = openaiKeyMatch[1];
                    console.log("OpenAI API key loaded via fetch fallback.");
                }

                if (!INTERNAL_API_KEYS.PINECONE_API_KEY && pineconeKeyMatch && pineconeKeyMatch[1]) {
                    INTERNAL_API_KEYS.PINECONE_API_KEY = pineconeKeyMatch[1];
                    console.log("Pinecone API key loaded via fetch fallback.");
                }
            } else {
                 console.warn("Fallback fetch for config.js failed or returned non-OK status.");
            }
        } catch (error) {
            console.warn("Could not load config.js via fetch fallback:", error);
        }
    }

    // Add debug logging with secure masking of keys
    console.log("Loaded API Keys (AIcalls.js):");
    console.log("  OPENAI_API_KEY:", INTERNAL_API_KEYS.OPENAI_API_KEY ?
      `${INTERNAL_API_KEYS.OPENAI_API_KEY.substring(0, 3)}...${INTERNAL_API_KEYS.OPENAI_API_KEY.substring(INTERNAL_API_KEYS.OPENAI_API_KEY.length - 3)}` :
      "Not found");
    console.log("  PINECONE_API_KEY:", INTERNAL_API_KEYS.PINECONE_API_KEY ?
      `${INTERNAL_API_KEYS.PINECONE_API_KEY.substring(0, 3)}...${INTERNAL_API_KEYS.PINECONE_API_KEY.substring(INTERNAL_API_KEYS.PINECONE_API_KEY.length - 3)}` :
      "Not found");

    const keysFound = !!(INTERNAL_API_KEYS.OPENAI_API_KEY && INTERNAL_API_KEYS.PINECONE_API_KEY);
    console.log("API Keys Initialized:", keysFound);
    // Return a copy to prevent external modification of the internal state
    return { ...INTERNAL_API_KEYS };

  } catch (error) {
    console.error("Error initializing API keys:", error);
    // Return empty keys on error
    return { OPENAI_API_KEY: "", PINECONE_API_KEY: "" };
  }
}

// Conversation history storage
let conversationHistory = [];

// Add this function at the top level
function showMessage(message) {
    const messageDiv = document.createElement('div');
    messageDiv.style.color = 'green';
    messageDiv.style.padding = '10px';
    messageDiv.style.margin = '10px';
    messageDiv.style.border = '1px solid green';
    messageDiv.style.borderRadius = '4px';
    messageDiv.textContent = message;
    
    const appBody = document.getElementById('app-body');
    appBody.insertBefore(messageDiv, appBody.firstChild);
    
    // Remove the message after 5 seconds
    setTimeout(() => {
        messageDiv.remove();
    }, 5000);
}

// Add this function at the top level
function showError(message) {
    const errorDiv = document.createElement('div');
    errorDiv.style.color = 'red';
    errorDiv.style.padding = '10px';
    errorDiv.style.margin = '10px';
    errorDiv.style.border = '1px solid red';
    errorDiv.style.borderRadius = '4px';
    errorDiv.textContent = `Error: ${message}`;
    
    const appBody = document.getElementById('app-body');
    appBody.insertBefore(errorDiv, appBody.firstChild);
    
    // Remove the error message after 5 seconds
    setTimeout(() => {
        errorDiv.remove();
    }, 5000);
}

// Add this function at the top level
function setButtonLoading(isLoading) {
    console.log(`[setButtonLoading] Called with isLoading: ${isLoading}`);
    const sendButton = document.getElementById('send');
    const loadingAnimation = document.getElementById('loading-animation');
    
    if (sendButton) {
        sendButton.disabled = isLoading;
    } else {
        console.warn("[setButtonLoading] Could not find send button with id='send'");
    }
    
    if (loadingAnimation) {
        const newDisplay = isLoading ? 'flex' : 'none';
        console.log(`[setButtonLoading] Found loadingAnimation element. Setting display to: ${newDisplay}`);
        loadingAnimation.style.display = newDisplay;
    } else {
        console.error("[setButtonLoading] Could not find loading animation element with id='loading-animation'");
    }
}

// Add this variable to store the last response
let lastResponse = null;

// Add this variable to track if the current message is a response
let isResponse = false;

// Add this function to write to Excel
async function writeToExcel() {
    if (!lastResponse) {
        showError('No response to write to Excel');
        return;
    }

    try {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.load("rowIndex");
            range.load("columnIndex");
            await context.sync();
            
            const startRow = range.rowIndex;
            const startCol = range.columnIndex;
            
            // Split the response into individual code strings
            let codeStrings = [];
            if (Array.isArray(lastResponse)) {
                // Join the array elements and then split by brackets
                const fullText = lastResponse.join(' ');
                codeStrings = fullText.match(/<[^>]+>/g) || [];
            } else if (typeof lastResponse === 'string') {
                codeStrings = lastResponse.match(/<[^>]+>/g) || [];
            }
            
            if (codeStrings.length === 0) {
                throw new Error("No valid code strings found in response");
            }
            
            // Create a range that spans all the rows we need
            const targetRange = range.worksheet.getRangeByIndexes(
                startRow,
                startCol,
                codeStrings.length,
                1
            );
            
            // Set all values at once, with each code string in its own row
            targetRange.values = codeStrings.map(str => [str]);
            
            await context.sync();
            console.log("Response written to Excel");
        });
    } catch (error) {
        console.error("Error writing to Excel:", error);
        showError(error.message);
    }
}

// Add this function to append messages to the chat log
function appendMessage(content, isUser = false) {
    const chatLog = document.getElementById('chat-log');
    const welcomeMessage = document.getElementById('welcome-message');
    
    // Hide welcome message when first message is added
    if (welcomeMessage) {
        welcomeMessage.style.display = 'none';
    }
    
    const messageDiv = document.createElement('div');
    messageDiv.className = `chat-message ${isUser ? 'user-message' : 'assistant-message'}`;
    
    const messageContent = document.createElement('p');
    messageContent.className = 'message-content';
    messageContent.textContent = content;
    
    messageDiv.appendChild(messageContent);
    chatLog.appendChild(messageDiv);
    
    // Scroll to bottom
    chatLog.scrollTop = chatLog.scrollHeight;
}

// Modify the handleSend function
async function handleSend() {
    const userInput = document.getElementById('user-input').value.trim();
    
    if (!userInput) {
        showError('Please enter a request');
        return;
    }

    // Check if this is a response to a previous message
    isResponse = conversationHistory.length > 0;

    // Add user message to chat
    appendMessage(userInput, true);
    
    // Clear input
    document.getElementById('user-input').value = '';

    setButtonLoading(true);
    try {
        // Process the text through the main function
        console.log("Starting structureDatabasequeries");
        const dbResults = await structureDatabasequeries(userInput);
        console.log("Database queries completed");
        
        if (!dbResults || !Array.isArray(dbResults)) {
            console.error("Invalid database results:", dbResults);
            throw new Error("Failed to get valid database results");
        }
        
        // Format the database results into a string
        const plainTextResults = dbResults.map(result => {
            if (!result) return "No results found";
            
            return `Query: ${result.query || 'No query'}\n` +
                   `Training Data:\n${(result.trainingData || []).join('\n')}\n` +
                   `Code Options:\n${(result.codeOptions || []).join('\n')}\n` +
                   `Code Choosing Context:\n${(result.call1Context || []).join('\n')}\n` +
                   `Code Editing Context:\n${(result.call2Context || []).join('\n')}\n` +
                   `---\n`;
        }).join('\n');

        const enhancedPrompt = `Client Request: ${userInput}\n\nDatabase Results:\n${plainTextResults}`;
        console.log("Enhanced prompt created");
        console.log("Enhanced prompt:", enhancedPrompt);

        console.log("Starting handleConversation");
        let conversationResult = await handleConversation(enhancedPrompt, isResponse); // Store the whole result object
        console.log("Conversation completed");
        console.log("Initial Conversation Result:", conversationResult); // Log the whole object

        // Extract the response array and update history
        let responseArray = conversationResult.response;
        conversationHistory = conversationResult.history; // Update global history if needed (check AIcalls.js if it manages history internally)

        // Validate the extracted response array
        if (!responseArray || !Array.isArray(responseArray)) {
            console.error("Invalid response array extracted:", responseArray);
            throw new Error("Failed to get valid response array from conversation result");
        }

        // Run validation and correction if needed (using the extracted array)
        console.log("Starting validation");
        const validationResults = await validateCodeStrings(responseArray);
        console.log("Validation completed:", validationResults);

        if (validationResults && validationResults.length > 0) {
            console.log("Starting validation correction");
            // Pass the extracted array to validationCorrection
            responseArray = await validationCorrection(userInput, responseArray, validationResults);
            console.log("Validation correction completed");
        }

        // Store the final response array for Excel writing
        lastResponse = responseArray;

        // Add assistant message to chat (using the extracted array)
        appendMessage(responseArray.join('\n'));
        
    } catch (error) {
        console.error("Error in handleSend:", error);
        showError(error.message);
        // Add error message to chat
        appendMessage(`Error: ${error.message}`);
    } finally {
        setButtonLoading(false);
    }
}

// Add this function to reset the chat
function resetChat() {
    // Clear the chat log
    const chatLog = document.getElementById('chat-log');
    chatLog.innerHTML = '';
    
    // Restore welcome message
    const welcomeMessage = document.createElement('div');
    welcomeMessage.id = 'welcome-message';
    welcomeMessage.className = 'welcome-message';
    const welcomeTitle = document.createElement('h1');
    welcomeTitle.textContent = 'What would you like to model?';
    welcomeMessage.appendChild(welcomeTitle);
    chatLog.appendChild(welcomeMessage);
    
    // Clear the conversation history
    conversationHistory = [];
    saveConversationHistory(conversationHistory);
    
    // Reset the response flag and last response
    isResponse = false;
    lastResponse = null;
    
    // Clear the input field
    document.getElementById('user-input').value = '';
    
    console.log("Chat reset completed");
}

// *** Define Helper Function Globally (BEFORE Office.onReady) ***
function getTabBlocks(codeString) {
    if (!codeString) return [];
    const tabBlocks = [];
    const tabRegex = /(<TAB;[^>]*>)/g;
    let match;
    const indices = [];
    while ((match = tabRegex.exec(codeString)) !== null) {
        indices.push({ index: match.index, tag: match[1] });
    }
    if (indices.length === 0) {
        if (codeString.trim().length > 0) {
            console.warn("Code string provided but no <TAB;...> tags found. Processing cannot proceed based on Tabs.");
        }
        return []; 
    } 
    for (let i = 0; i < indices.length; i++) {
        const start = indices[i].index;
        const tag = indices[i].tag;
        const end = (i + 1 < indices.length) ? indices[i + 1].index : codeString.length;
        const blockText = codeString.substring(start, end).trim();
        if (blockText) {
            tabBlocks.push({ tag: tag, text: blockText });
        }
    }
    return tabBlocks;
}

// Helper function to find max driver numbers in existing text
function getMaxDriverNumbers(text) {
    const maxNumbers = {};
    // MODIFIED Regex: Allow optional spaces around =
    const regex = /row\d+\s*=\s*"([A-Z]+)(\d*)\|/g;
    let match;
    console.log("Scanning text for drivers:", text.substring(0, 200) + "..."); // Log input text

    while ((match = regex.exec(text)) !== null) {
        const prefix = match[1];
        const numberStr = match[2];
        const number = numberStr ? parseInt(numberStr, 10) : 0;
        console.log(`Found driver match: prefix='${prefix}', numberStr='${numberStr}', number=${number}`); // Log each match

        if (isNaN(number)) {
             console.warn(`Parsed NaN for number from '${numberStr}' for prefix '${prefix}'. Skipping.`);
             continue;
        }

        if (!maxNumbers[prefix] || number > maxNumbers[prefix]) {
            maxNumbers[prefix] = number;
            console.log(`Updated max for '${prefix}' to ${number}`); // Log updates
        }
    }
    if (Object.keys(maxNumbers).length === 0) {
        console.log("No existing drivers found matching the pattern.");
    }
    console.log("Final max existing driver numbers:", maxNumbers);
    return maxNumbers;
}

async function insertSheetsAndRunCodes() {

    // >>> ADDED: Get codes from textarea and save automatically
    const codesTextarea = document.getElementById('codes-textarea');
    if (!codesTextarea) {
        showError("Could not find the code input area. Cannot run codes.");
        return;
    }
    loadedCodeStrings = codesTextarea.value; // Update global variable
    try {
        localStorage.setItem('userCodeStrings', loadedCodeStrings);
        console.log("[Run Codes] Automatically saved codes from textarea to localStorage.");
    } catch (error) {
        console.error("[Run Codes] Error auto-saving codes to localStorage:", error);
        showError(`Error automatically saving codes: ${error.message}. Run may not reflect latest changes.`);
    }
    // <<< END ADDED CODE

    let codesToRun = loadedCodeStrings;
    let previousCodes = null;
    // >>> REVISED: Unified list for all content to process
    let allCodeContentToProcess = ""; // Holds full text of *all* new AND modified tabs
    let runResult = null; // To store result from runCodes if called
    // Removed: codesToProcessForRunCodes, tabsToInsertIncrementally, codeStringToValidate (use allCodeContentToProcess)

    // Main processing wrapped in try/catch/finally
    try {
        // --- Check Financials Sheet Existence ---
        let financialsSheetExists = false;
        await Excel.run(async (context) => {
            try {
                const financialsSheet = context.workbook.worksheets.getItem("Financials");
                financialsSheet.load("name");
                await context.sync();
                financialsSheetExists = true;
            } catch (error) {
                if (error instanceof OfficeExtension.Error && error.code === Excel.ErrorCodes.itemNotFound) {
                    financialsSheetExists = false;
                } else { throw error; } // Rethrow other errors
            }
        });

        // Set calculation mode to manual (do this early)
        await Excel.run(async (context) => {
            context.application.calculationMode = Excel.CalculationMode.manual;
            await context.sync();
        });

        setButtonLoading(true);
        console.log("Starting code processing...");

        // --- Pass Logic: Determine sheets to insert and codes to process ---
        if (!financialsSheetExists) {
            // *** FIRST PASS ***
            console.log("[Run Codes] FIRST PASS: Financials sheet not found.");
            allCodeContentToProcess = codesToRun; // All codes are new

            // >>> VALIDATION FOR FIRST PASS <<<
            if (allCodeContentToProcess.trim().length > 0) {
                console.log("Validating ALL codes before initial base sheet insertion...");
                const validationErrors = await validateCodeStringsForRun(allCodeContentToProcess.split(/\r?\n/).filter(line => line.trim() !== ''));
                if (validationErrors && validationErrors.length > 0) {
                    const errorMsg = "Initial validation failed. Please fix the errors before running:\n" + validationErrors.join("\n");
                    console.error("Code validation failed:", validationErrors);
                    showError("Code validation failed. See chat for details.");
                    appendMessage(errorMsg);
                    setButtonLoading(false);
                    return; // Stop execution
                }
                console.log("Initial code validation successful.");
            } else {
                console.log("[Run Codes] No codes to validate on first pass.");
                // If no codes, no need to insert base sheets? Or insert anyway? Assuming insert needed.
            }

            // --- Insert BASE sheets ---
            console.log("Inserting base sheets from Worksheets_4.3.25 v1.xlsx...");
            // ... (fetch and handleInsertWorksheetsFromBase64 for Worksheets_4.3.25 remains the same) ...
            const worksheetsResponse = await fetch('https://localhost:3002/assets/Worksheets_4.3.25 v1.xlsx');
            if (!worksheetsResponse.ok) throw new Error(`Worksheets load failed: ${worksheetsResponse.statusText}`);
            const worksheetsArrayBuffer = await worksheetsResponse.arrayBuffer();
            console.log("Converting base file to base64 (using chunks)...");
            const worksheetsUint8Array = new Uint8Array(worksheetsArrayBuffer);
            let worksheetsBinaryString = '';
            const chunkSize = 8192;
            for (let i = 0; i < worksheetsUint8Array.length; i += chunkSize) {
                const chunk = worksheetsUint8Array.slice(i, Math.min(i + chunkSize, worksheetsUint8Array.length));
                worksheetsBinaryString += String.fromCharCode.apply(null, chunk);
            }
            const worksheetsBase64String = btoa(worksheetsBinaryString);
            console.log("Base64 conversion complete.");
            await handleInsertWorksheetsFromBase64(worksheetsBase64String);
            console.log("Base sheets inserted.");
            // No need to insert codes.xlsx here, it happens after runCodes

        } else {
            // *** SECOND PASS (or later) ***
            console.log("[Run Codes] SUBSEQUENT PASS: Financials sheet found.");

            // Load previous codes for comparison
            try {
                previousCodes = localStorage.getItem('previousRunCodeStrings');
            } catch (error) {
                 console.error("[Run Codes] Error loading previous codes for comparison:", error);
                 console.warn("[Run Codes] Could not load previous codes. Processing ALL current codes as fallback.");
                 previousCodes = null; // Treat as if all codes are new if loading fails
            }

            // Check for NO changes first
            if (previousCodes !== null && previousCodes === codesToRun) {
                 console.log("[Run Codes] No change in code strings since last run. Nothing to process.");
                 // Update previous state just in case, though it's identical
                 try { localStorage.setItem('previousRunCodeStrings', codesToRun); } catch(e) { console.error("Err updating prev codes:", e); }
                 showMessage("No code changes to run.");
                 setButtonLoading(false);
                 return; // <<< EXIT EARLY
            }

            // --- Identify New and Modified Tabs & Collect ALL content ---
            const currentTabs = getTabBlocks(codesToRun);
            const previousTabs = getTabBlocks(previousCodes || ""); // Use empty string if previousCodes is null
            const previousTabMap = new Map(previousTabs.map(block => [block.tag, block.text]));

            let hasAnyChanges = false; // Flag to check if codes.xlsx needs inserting
            const codeRegex = /<[^>]+>/g; // Regex to find <...> codes

            for (const currentTab of currentTabs) {
                const currentTag = currentTab.tag; // This is the <TAB;...> tag
                const currentText = currentTab.text; // This is the full text block starting with <TAB;...>
                const previousText = previousTabMap.get(currentTag);

                if (previousText === undefined) {
                    // *** New Tab ***
                    console.log(`[Run Codes] Identified NEW tab: ${currentTag}. Adding all its codes.`);
                    const newTabCodes = currentText.match(codeRegex) || [];
                    if (newTabCodes.length > 0) {
                        allCodeContentToProcess += newTabCodes.join("\n") + "\n\n"; // newTabCodes already includes the TAB tag
                        hasAnyChanges = true;
                    }
                } else {
                    // *** Existing Tab: Compare individual codes ***
                    const currentCodes = currentText.match(codeRegex) || [];
                    const previousCodesSet = new Set((previousText || "").match(codeRegex) || []);
                    let tabHasChanges = false;
                    let codesToAddForThisTab = ""; // Collect codes for this tab

                    for (const currentCode of currentCodes) {
                        if (!previousCodesSet.has(currentCode)) {
                            console.log(`[Run Codes] Identified NEW/MODIFIED code in tab ${currentTag}: ${currentCode.substring(0, 80)}...`);
                            codesToAddForThisTab += currentCode + "\n"; // Collect only the new/modified code
                            hasAnyChanges = true;
                            tabHasChanges = true; // Mark that this specific tab had changes
                        }
                    }
                    // If codes were added from this modified tab, prefix them with the TAB tag
                    if (tabHasChanges) {
                        allCodeContentToProcess += currentTag + "\n"; // Add TAB tag
                        allCodeContentToProcess += codesToAddForThisTab + "\n"; // Add collected codes + separator
                    }
                }
            }

            // --- Validate ALL collected content BEFORE inserting codes.xlsx ---
            if (hasAnyChanges) {
                if (allCodeContentToProcess.trim().length > 0) {
                    console.log("Validating ALL content from new/modified tabs BEFORE inserting codes.xlsx...");
                    const validationErrors = await validateCodeStringsForRun(allCodeContentToProcess.split(/\r?\n/).filter(line => line.trim() !== ''));

                    if (validationErrors && validationErrors.length > 0) {
                        const errorMsg = "Validation failed. Please fix the errors before running:\n" + validationErrors.join("\n");
                        console.error("Code validation failed:", validationErrors);
                        showError("Code validation failed. See chat for details.");
                        appendMessage(errorMsg);
                        setButtonLoading(false);
                        return; // Stop execution
                    }
                    console.log("Code validation successful for new/modified tabs.");
                } else {
                    console.log("[Run Codes] Changes detected, but no code content found for validation in new/modified tabs.");
                     // This case might occur if only whitespace/comments changed within tabs
                     // Decide if codes.xlsx insertion is still needed? Assuming yes if hasAnyChanges=true
                }

                // --- Insert codes.xlsx only AFTER successful validation if changes were detected ---
                console.log(`[Run Codes] Changes detected and validated: ${currentTabs.filter(tab => !previousTabMap.has(tab.tag) || previousTabMap.get(tab.tag) !== tab.text).map(tab => tab.tag).join(', ')}. Inserting base sheets from codes.xlsx...`);
                try {
                    const codesResponse = await fetch('https://localhost:3002/assets/codes.xlsx');
                    if (!codesResponse.ok) throw new Error(`codes.xlsx load failed: ${codesResponse.statusText}`);
                    const codesArrayBuffer = await codesResponse.arrayBuffer();
                    console.log("Converting codes.xlsx file to base64 (using chunks)...");
                    // ... (base64 conversion remains the same) ...
                    const codesUint8Array = new Uint8Array(codesArrayBuffer);
                    let codesBinaryString = '';
                    const chunkSize_codes = 8192;
                    for (let i = 0; i < codesUint8Array.length; i += chunkSize_codes) {
                        const chunk = codesUint8Array.slice(i, Math.min(i + chunkSize_codes, codesUint8Array.length));
                        codesBinaryString += String.fromCharCode.apply(null, chunk);
                    }
                    const codesBase64String = btoa(codesBinaryString);

                    console.log("codes.xlsx Base64 conversion complete.");
                    await handleInsertWorksheetsFromBase64(codesBase64String);
                    console.log("codes.xlsx sheets inserted.");
                } catch (e) {
                    console.error("Failed to insert sheets from codes.xlsx:", e);
                    showError("Failed to insert necessary sheets from codes.xlsx. Aborting.");
                    setButtonLoading(false);
                    return;
                }
            } else {
                 console.log("[Run Codes] No changes identified in tabs compared to previous run. Nothing to insert or process.");
                 // Update previous state as the content is effectively the same
                 try { localStorage.setItem('previousRunCodeStrings', codesToRun); } catch(e) { console.error("Err updating prev codes:", e); }
                 showMessage("No code changes identified to run.");
                 setButtonLoading(false);
                 return; // Exit if no actionable changes
            }

        } // End of pass logic (if/else financialsSheetExists)

        // --- Execute Processing ---

        // >>> REMOVED: Incremental Insertion block <<<

        // --- Process ALL collected content (New + Modified Tabs) using runCodes ---
        if (allCodeContentToProcess.trim().length > 0) {
            console.log("[Run Codes] Processing collected content from new/modified tabs...");
            console.log("Populating collection...");
            const collection = populateCodeCollection(allCodeContentToProcess);
            console.log(`Collection populated with ${collection.length} code(s)`);

             // Check if collection is empty after population (might happen if only comments/whitespace)
            if (collection.length > 0) {
                console.log("Running codes...");
                runResult = await runCodes(collection); // Store the result
                console.log("Codes executed:", runResult);
            } else {
                 console.log("[Run Codes] Collection is empty after population, skipping runCodes execution.");
                 // Ensure runResult is initialized for post-processing
                 if (!runResult) runResult = { assumptionTabs: [] };
            }
        } else {
            console.log("[Run Codes] No code content collected to process via runCodes.");
             // Initialize runResult structure if needed by post-processing
             if (!runResult) runResult = { assumptionTabs: [] }; // Ensure runResult exists
        }

        // --- Post-processing (Runs regardless) ---
        console.log("[Run Codes] Starting post-processing steps...");
        if (runResult && runResult.assumptionTabs && runResult.assumptionTabs.length > 0) {
            console.log("Processing assumption tabs...");
            await processAssumptionTabs(runResult.assumptionTabs);
        } else {
             console.log("No assumption tabs to process.");
        }
        console.log("Hiding specific columns and navigating...");
        // Pass assumption tabs from runResult (if any), otherwise an empty array.
        await hideColumnsAndNavigate(runResult?.assumptionTabs || []);

        // Cleanup
        console.log("Deleting Codes sheet / Skipping hiding Calcs sheet...");
        await Excel.run(async (context) => {
            try {
                const codesSheet = context.workbook.worksheets.getItem("Codes");
                codesSheet.delete();
                console.log("Codes sheet deleted.");
            } catch (e) {
                if (e instanceof OfficeExtension.Error && e.code === Excel.ErrorCodes.itemNotFound) {
                     console.warn("Codes sheet not found, skipping deletion.");
                } else { console.error("Error deleting Codes sheet:", e); }
            }
            // Hiding Calcs sheet logic was removed previously, keeping it out.
            await context.sync();
        }).catch(error => { console.error("Error during sheet cleanup:", error); });

        // --- IMPORTANT: Update previous codes state AFTER successful processing ---
        try {
            localStorage.setItem('previousRunCodeStrings', codesToRun);
            console.log("[Run Codes] Updated previous run state with full current codes.");
        } catch (error) {
             console.error("[Run Codes] Failed to update previous run state:", error);
        }

        showMessage("Code processing finished successfully!");

    } catch (error) {
        console.error("An error occurred during the build process:", error);
        showError(`Operation failed: ${error.message || error.toString()}`);
    } finally {
        // Always set calculation mode back to automatic and hide loading
        try {
            await Excel.run(async (context) => {
                context.application.calculationMode = Excel.CalculationMode.automatic;
                await context.sync();
            });
        } catch (finalError) {
            console.error("Error setting calculation mode to automatic:", finalError);
        }
        setButtonLoading(false);
    }
}

// Ensure Office.onReady sets up the button click handler for the REVISED function
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // Assign the REVISED async function as the handler
    const button = document.getElementById("insert-and-run");
    if (button) {
        button.onclick = insertSheetsAndRunCodes; // Use the revised function
    } else {
        console.error("Could not find button with id='insert-and-run'");
    }

    // ... (rest of your Office.onReady remains the same) ...

    // Keep the setup for your other buttons (send-button, reset-button, etc.)
    const sendButton = document.getElementById('send');
    if (sendButton) sendButton.onclick = handleSend;

    const writeButton = document.getElementById('write-to-excel');
    if (writeButton) writeButton.onclick = writeToExcel;

    const resetButton = document.getElementById('reset-chat');
    if (resetButton) resetButton.onclick = resetChat;

    // >>> ADDED: Setup for Generate Tab String button
    const generateTabStringButton = document.getElementById('generate-tab-string-button');
    if (generateTabStringButton) {
        generateTabStringButton.onclick = generateTabString; // Assign the imported function
    } else {
        console.error("Could not find button with id='generate-tab-string-button'");
    }
    // <<< END ADDED CODE

    const codesTextarea = document.getElementById('codes-textarea');
    const editParamsButton = document.getElementById('edit-code-params-button');
    const paramsModal = document.getElementById('code-params-modal');
    const paramsModalForm = document.getElementById('code-params-modal-form');
    const closeModalButton = paramsModal.querySelector('.close-button');
    const applyParamsButton = document.getElementById('apply-code-params-button');
    const cancelParamsButton = document.getElementById('cancel-code-params-button');

    // Modal Find/Replace elements
    const modalFindInput = document.getElementById('modal-find-input');
    const modalReplaceInput = document.getElementById('modal-replace-input');
    const modalReplaceAllButton = document.getElementById('modal-replace-all-button');
    const modalSearchStatus = document.getElementById('modal-search-status');

    let currentCodeStringRange = null; // To store {start, end} of the code string being edited
    let currentCodeStringType = ''; // To store the type like 'VOL-EV'

    // State for modal find/replace (Simplified)
    let modalSearchableElements = []; // Stores {element, originalValue}
    // Removed modalSearchTerm, modalCurrentMatchIndex, modalAllMatches

    // Function to reset modal search state (Simplified)
    const resetModalSearchState = () => {
        modalSearchableElements = Array.from(paramsModalForm.querySelectorAll('input[data-param-key], textarea[data-param-key]'));
        if (modalSearchStatus) modalSearchStatus.textContent = '';
        // Clear input fields as well?
        // if (modalFindInput) modalFindInput.value = '';
        // if (modalReplaceInput) modalReplaceInput.value = '';
        console.log("Modal search state reset.");
    };

    // Function to update modal search status
    const updateModalSearchStatus = (message) => {
        if (modalSearchStatus) {
            modalSearchStatus.textContent = message;
        }
    };

    // Removed findAllMatchesInModal function

    // Function to show the modal
    const showParamsModal = () => {
        if (paramsModal) {
            paramsModal.style.display = 'block';
            resetModalSearchState(); // Reset search when modal opens
        }
    };

    // Function to hide the modal
    const hideParamsModal = () => {
        if (paramsModal) {
            paramsModal.style.display = 'none';
            paramsModalForm.innerHTML = ''; // Clear the form
            currentCodeStringRange = null; // Reset state
            currentCodeStringType = '';
            resetModalSearchState(); // Also reset search state on close
        }
    };

    // Function to find the <...> block around the cursor
    const findCodeStringAroundCursor = (text, cursorPos) => {
        const textBeforeCursor = text.substring(0, cursorPos);
        const textAfterCursor = text.substring(cursorPos);

        const lastOpenBracket = textBeforeCursor.lastIndexOf('<');
        const lastCloseBracketBefore = textBeforeCursor.lastIndexOf('>');

        // Check if cursor is potentially inside brackets
        if (lastOpenBracket > lastCloseBracketBefore) {
            const firstCloseBracketAfter = textAfterCursor.indexOf('>');
            if (firstCloseBracketAfter !== -1) {
                const start = lastOpenBracket;
                const end = cursorPos + firstCloseBracketAfter + 1; // +1 to include '>'
                const codeString = text.substring(start, end);
                console.log(`Found code string: ${codeString} at range [${start}, ${end})`);
                return { codeString, start, end };
            }
        }
        console.log("Cursor not inside a <...> block.");
        return null; // Cursor is not inside a valid <...> block
    };

    // Function to parse parameters from the code string content (inside <...>)
    const parseCodeParameters = (content) => {
        const parts = content.split(';');
        if (parts.length < 1) return { type: '', params: {} };

        const type = parts[0].trim();
        const params = {};
        // Regex to match key="value" or key=value (no quotes)
        const paramRegex = /\s*([^=\s]+)\s*=\s*(?:"([^"]*)"|([^;]*))/g;

        for (let i = 1; i < parts.length; i++) {
            const part = parts[i].trim();
            if (!part) continue;

            // Reset regex index before each exec
            paramRegex.lastIndex = 0;
            const match = paramRegex.exec(part);

            if (match) {
                const key = match[1];
                // Value could be in group 2 (quoted) or group 3 (unquoted)
                const value = match[2] !== undefined ? match[2] : match[3];
                 if (key) { // Ensure key is valid
                    params[key] = value.trim();
                }
            } else {
                console.warn(`Could not parse parameter part: '${part}'`);
            }
        }
        console.log(`Parsed type: ${type}, params:`, params);
        return { type, params };
    };

    // Function to populate the modal form (needs to update searchable elements)
    const populateParamsModal = (type, params) => {
        paramsModalForm.innerHTML = ''; // Clear previous form items
        currentCodeStringType = type; // Store the type

        Object.entries(params).forEach(([key, value]) => {
            const paramEntryDiv = document.createElement('div');
            paramEntryDiv.className = 'param-entry';

            const label = document.createElement('label');
            label.htmlFor = `param-${key}`;
            label.textContent = key;

            let inputElement;
            const isLongValue = key.toLowerCase().includes('row') || value.length > 60;
            const isLIParam = /LI\d+\|/.test(value.trim()); // Check if value starts with LI<digit>|

            if (isLongValue || isLIParam) { // Use textarea for LI params too, for consistency
                inputElement = document.createElement('textarea');
                inputElement.rows = isLIParam ? 2 : 3; // Slightly smaller for LI rows initially
            } else {
                inputElement = document.createElement('input');
                inputElement.type = 'text';
            }

            inputElement.id = `param-${key}`;
            inputElement.value = value;
            inputElement.dataset.paramKey = key;
            if (isLIParam) {
                inputElement.dataset.isOriginalLi = "true"; // Mark original LI fields
            }

            paramEntryDiv.appendChild(label);

            if (isLIParam) {
                // Create a container for the LI field and its add button
                const liContainer = document.createElement('div');
                liContainer.className = 'li-parameter-container';
                liContainer.dataset.originalLiKey = key; // Link container to original key

                liContainer.appendChild(inputElement); // Add the input field first

                // Create the Add button
                const addButton = document.createElement('button');
                addButton.type = 'button'; // Important: prevent form submission
                addButton.textContent = '+';
                addButton.className = 'ms-Button ms-Button--icon add-li-button'; // Add specific class
                addButton.title = 'Add another LI item based on this one';
                addButton.dataset.targetLiKey = key; // Link button to the input's key

                addButton.onclick = (event) => {
                    const sourceInput = document.getElementById(`param-${key}`);
                    if (!sourceInput) return;

                    const newValueContainer = document.createElement('div');
                    newValueContainer.className = 'added-li-item';

                    const newInput = sourceInput.cloneNode(true); // Clone the original input/textarea
                    // Clear ID, mark as added, remove original marker
                    newInput.id = '';
                    newInput.dataset.isAddedLi = "true";
                    delete newInput.dataset.isOriginalLi;
                    newInput.dataset.originalLiKey = key; // Link back to the original key
                    // Keep the same value as the original initially
                    newInput.value = sourceInput.value; // Duplicate the content

                    // Add a remove button for the added item
                    const removeButton = document.createElement('button');
                    removeButton.type = 'button';
                    removeButton.textContent = '-';
                    removeButton.className = 'ms-Button ms-Button--icon remove-li-button';
                    removeButton.title = 'Remove this added LI item';
                    removeButton.onclick = () => {
                        newValueContainer.remove();
                    };

                    newValueContainer.appendChild(newInput);
                    newValueContainer.appendChild(removeButton);

                    // Insert the new container after the clicked button
                    // or after the last added item within this container
                    event.target.parentNode.appendChild(newValueContainer);
                     // Maybe scroll container? paramsModalForm.scrollTop = paramsModalForm.scrollHeight;
                };

                liContainer.appendChild(addButton); // Add button after input
                paramEntryDiv.appendChild(liContainer); // Add container to entry div

            } else {
                 paramEntryDiv.appendChild(inputElement); // Non-LI params added directly
            }

            paramsModalForm.appendChild(paramEntryDiv);
        });
        // IMPORTANT: Update searchable elements after populating
        resetModalSearchState(); // Reset search state after populating form
    };

    // --- Event Listener for the Edit Parameters Button ---
    if (editParamsButton && codesTextarea && paramsModal) {
        editParamsButton.onclick = () => {
            const text = codesTextarea.value;
            const cursorPos = codesTextarea.selectionStart;

            const codeInfo = findCodeStringAroundCursor(text, cursorPos);

            if (codeInfo) {
                // Extract content within < >
                const content = codeInfo.codeString.substring(1, codeInfo.codeString.length - 1);
                const { type, params } = parseCodeParameters(content);

                if (type) {
                    currentCodeStringRange = { start: codeInfo.start, end: codeInfo.end };
                    populateParamsModal(type, params);
                    showParamsModal();
                } else {
                    showError("Could not parse the code string structure.");
                }
            } else {
                showError("Place cursor inside a <...> code block to edit parameters.");
            }
        };
    }

    // --- Event Listeners for Modal Actions ---
    if (closeModalButton) {
        closeModalButton.onclick = hideParamsModal;
    }
    if (cancelParamsButton) {
        cancelParamsButton.onclick = hideParamsModal;
    }

    // --- APPLY CHANGES LOGIC (MODIFIED) ---
    if (applyParamsButton && codesTextarea) {
        applyParamsButton.onclick = () => {
            if (!currentCodeStringRange || !currentCodeStringType) return; // Safety check

            // Use a map to reconstruct parameters, handling LI aggregation
            const paramValues = {};

            // Process all input/textarea fields in the form
            const formElements = paramsModalForm.querySelectorAll('input[data-param-key], textarea[data-param-key]');

            formElements.forEach(input => {
                const key = input.dataset.paramKey;
                const isOriginalLI = input.dataset.isOriginalLi === "true";
                const isAddedLI = input.dataset.isAddedLi === "true";
                const value = input.value;

                if (isOriginalLI) {
                    // If it's an original LI, initialize its value in the map
                    if (!paramValues[key]) {
                        paramValues[key] = value; // Start with the original value
                    }
                } else if (isAddedLI) {
                    // This case handled below by finding related elements
                    // We only need to store original keys first
                } else if (key && !isAddedLI) {
                    // Standard parameter, just store its value
                     if (!paramValues[key]) { // Check prevents overwriting if key appears twice (shouldn't happen)
                       paramValues[key] = value;
                    }
                }
            });

            // Now, aggregate added LI items
            const addedLiElements = paramsModalForm.querySelectorAll('textarea[data-is-added-li="true"]');
            addedLiElements.forEach(addedInput => {
                 const originalKey = addedInput.dataset.originalLiKey;
                 if (originalKey && paramValues[originalKey]) {
                      // Append the added value, prefixed with *
                      paramValues[originalKey] += ` *${addedInput.value}`;
                 }
            });

            // Build the final parameter string parts
            const updatedParams = Object.entries(paramValues).map(([key, finalValue]) => {
                 // Re-add quotes around the final aggregated value
                 return `${key}="${finalValue}"`;
            });

            // Reconstruct the code string
            const newCodeStringContent = `${currentCodeStringType}; ${updatedParams.join('; ')}`;
            const newCodeString = `<${newCodeStringContent}>`;

            // Update the textarea content
            const currentText = codesTextarea.value;
            const textBefore = currentText.substring(0, currentCodeStringRange.start);
            const textAfter = currentText.substring(currentCodeStringRange.end);

            codesTextarea.value = textBefore + newCodeString + textAfter;

            console.log(`Updated code string at [${currentCodeStringRange.start}, ${currentCodeStringRange.start + newCodeString.length})`);
            console.log("New string:", newCodeString);

            // Optionally, update cursor position
            const newCursorPos = currentCodeStringRange.start + newCodeString.length;
            codesTextarea.focus();
            codesTextarea.setSelectionRange(newCursorPos, newCursorPos);

            hideParamsModal(); // Close modal after applying
        };
    }

    // --- Modal Find/Replace Logic (Simplified) ---

    const modalReplaceAll = () => {
        const searchTerm = modalFindInput.value;
        const replaceTerm = modalReplaceInput.value;
        if (!searchTerm) {
            updateModalSearchStatus("Enter search term.");
            return;
        }

        // Ensure searchable elements are up-to-date
        modalSearchableElements = Array.from(paramsModalForm.querySelectorAll('input[data-param-key], textarea[data-param-key]'));

        let replacementsMade = 0;
        modalSearchableElements.forEach((element, index) => {
            let currentValue = element.value;
            // Escape regex special characters in search term
            const escapedSearchTerm = searchTerm.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
            let newValue = currentValue.replace(new RegExp(escapedSearchTerm, 'g'), () => {
                replacementsMade++;
                return replaceTerm;
            });
            if (currentValue !== newValue) {
                element.value = newValue;
                console.log(`Modal Replace All: Made replacements in element ${index}`);
            }
        });

        if (replacementsMade > 0) {
            updateModalSearchStatus(`Replaced ${replacementsMade} occurrence(s).`);
            // No need to reset search state as there's no find next
        } else {
            updateModalSearchStatus(`"${searchTerm}" not found.`);
        }
    };

    // Add event listeners for modal find/replace buttons (Simplified)
    // Removed listeners for Find Next and Replace
    if (modalReplaceAllButton) modalReplaceAllButton.onclick = modalReplaceAll;
    // Removed input listener for modalFindInput

    // --- Event Listeners for Modal Actions ---
    if (closeModalButton) {
        closeModalButton.onclick = hideParamsModal;
    }

    // ... (rest of your Office.onReady, including suggestion logic, initializations)

    // Make sure initialization runs after setting up modal logic
    Promise.all([
        initializeAPIKeys(),
        loadCodeDatabase()
    ]).then(([keys]) => {
      if (!keys) {
        showError("Failed to load API keys. Please check configuration.");
      } else {
        // >>> ADDED: Set the API keys in AIcalls module
        setAPIKeys(keys);
      }
      conversationHistory = loadConversationHistory();

      try {
          const storedCodes = localStorage.getItem('userCodeStrings');
          if (storedCodes !== null) {
              loadedCodeStrings = storedCodes;
              if (codesTextarea) {
                  codesTextarea.value = loadedCodeStrings;
              }
              console.log("Code strings loaded from localStorage into global variable.");
          } else {
              console.log("No code strings found in localStorage, initializing global variable as empty.");
              loadedCodeStrings = "";
          }
          // Also load the previous run codes if available
           const storedPreviousCodes = localStorage.getItem('previousRunCodeStrings');
           if (storedPreviousCodes) {
               console.log("Previous run code strings loaded from localStorage.");
           }

      } catch (error) {
          console.error("Error loading code strings from localStorage:", error);
          showError(`Error loading codes from storage: ${error.message}`);
          loadedCodeStrings = "";
      }

      // >>> MOVED: Assign event listeners AFTER initialization is complete
      // Setup cursor position tracking
      if (codesTextarea) {
          const updateCursorPosition = () => {
              lastEditorCursorPosition = codesTextarea.selectionStart;
              // console.log(`Cursor position updated: ${lastEditorCursorPosition}`); // Optional debug log
          };
          codesTextarea.addEventListener('keyup', updateCursorPosition); // Update on key release
          codesTextarea.addEventListener('mouseup', updateCursorPosition); // Update on mouse click release
          codesTextarea.addEventListener('focus', updateCursorPosition);   // Update when focus is gained
          // codesTextarea.addEventListener('blur', updateCursorPosition); // Maybe don't update on blur?
          console.log("[Office.onReady] Added event listeners to codesTextarea for cursor tracking."); // <<< DEBUG LOG
      }

      // Setup Insert to Editor button
      const insertToEditorButton = document.getElementById('insert-to-editor');
      if (insertToEditorButton) {
          console.log("[Office.onReady] Found insert-to-editor button."); // <<< DEBUG LOG
          insertToEditorButton.onclick = insertResponseToEditor;
          console.log("[Office.onReady] Assigned onclick for insert-to-editor button."); // <<< DEBUG LOG
      } else {
          console.error("[Office.onReady] Could not find button with id='insert-to-editor'");
      }
      // <<< END MOVED CODE

    }).catch(error => {
        console.error("Error during initialization:", error);
        showError("Error during initialization: " + error.message);
    });

    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "block";

    // ... (existing modal logic: applyParamsButton.onclick, window.onclick)

    // --- Code Suggestion Logic (Restored) ---
    let dynamicSuggestionsContainer = document.getElementById('dynamic-suggestions-container');
    if (!dynamicSuggestionsContainer) {
        dynamicSuggestionsContainer = document.createElement('div');
        dynamicSuggestionsContainer.id = 'dynamic-suggestions-container';
        dynamicSuggestionsContainer.className = 'code-suggestions'; // Reuse class if styling exists
        dynamicSuggestionsContainer.style.display = 'none';
        // Basic positioning styles (adjust in CSS for better control)
        dynamicSuggestionsContainer.style.position = 'absolute';
        dynamicSuggestionsContainer.style.border = '1px solid #ccc';
        dynamicSuggestionsContainer.style.backgroundColor = 'white';
        dynamicSuggestionsContainer.style.maxHeight = '150px';
        dynamicSuggestionsContainer.style.overflowY = 'auto';
        dynamicSuggestionsContainer.style.zIndex = '1000';

        // Insert after the textarea's container or adjust as needed
        if (codesTextarea && codesTextarea.parentNode) { // Check if codesTextarea exists
            codesTextarea.parentNode.insertBefore(dynamicSuggestionsContainer, codesTextarea.nextSibling);
        } else {
            // Fallback: Append to body, though less ideal positioning
            document.body.appendChild(dynamicSuggestionsContainer);
        }

        // Function to update position and width
        const updateSuggestionPosition = () => {
          if (dynamicSuggestionsContainer.style.display === 'block' && codesTextarea) {
              const rect = codesTextarea.getBoundingClientRect();
              dynamicSuggestionsContainer.style.width = codesTextarea.offsetWidth + 'px';
              dynamicSuggestionsContainer.style.top = (rect.bottom + window.scrollY) + 'px';
              dynamicSuggestionsContainer.style.left = (rect.left + window.scrollX) + 'px';
          }
        };

        // Update on resize and scroll
        window.addEventListener('resize', updateSuggestionPosition);
        window.addEventListener('scroll', updateSuggestionPosition, true); // Use capture phase for scroll
    }

    let highlightedSuggestionIndex = -1;
    let currentSuggestions = [];

    const updateHighlight = (newIndex) => {
      if (!dynamicSuggestionsContainer) return; // Guard against null
      const suggestionItems = dynamicSuggestionsContainer.querySelectorAll('.code-suggestion-item');
      if (highlightedSuggestionIndex >= 0 && highlightedSuggestionIndex < suggestionItems.length) {
        suggestionItems[highlightedSuggestionIndex].classList.remove('suggestion-highlight');
      }
      if (newIndex >= 0 && newIndex < suggestionItems.length) {
        suggestionItems[newIndex].classList.add('suggestion-highlight');
        suggestionItems[newIndex].scrollIntoView({ block: 'nearest' });
      }
      highlightedSuggestionIndex = newIndex;
    };

    const showSuggestionsForTerm = (searchTerm) => {
        if (!dynamicSuggestionsContainer || !codesTextarea) return; // Guard against null

        searchTerm = searchTerm.toLowerCase().trim();
        console.log(`[showSuggestionsForTerm] Search Term: '${searchTerm}'`);

        dynamicSuggestionsContainer.innerHTML = '';
        highlightedSuggestionIndex = -1;
        currentSuggestions = [];

        if (searchTerm.length < 2) {
            console.log("[showSuggestionsForTerm] Search term too short, hiding suggestions.");
            dynamicSuggestionsContainer.style.display = 'none';
            return;
        }

        console.log("[showSuggestionsForTerm] Filtering code database...");
        const suggestions = codeDatabase
            .filter(item => {
                const hasName = item && typeof item.name === 'string';
                return hasName && item.name.toLowerCase().includes(searchTerm);
            })
            .slice(0, 10);

        currentSuggestions = suggestions;
        console.log(`[showSuggestionsForTerm] Found ${currentSuggestions.length} suggestions:`, currentSuggestions);

        if (currentSuggestions.length > 0) {
            console.log("[showSuggestionsForTerm] Populating suggestions container...");
            currentSuggestions.forEach((item, i) => {
                const suggestionDiv = document.createElement('div');
                suggestionDiv.className = 'code-suggestion-item';
                suggestionDiv.textContent = item.name;
                suggestionDiv.dataset.index = i;

                suggestionDiv.onclick = () => {
                    console.log(`Suggestion clicked: '${item.name}'`);
                    const currentText = codesTextarea.value;
                    const cursorPosition = codesTextarea.selectionStart;
                    let codeToAdd = item.code;

                    let insertionPosition = cursorPosition;
                    let wasAdjusted = false;
                    const textBeforeCursor = currentText.substring(0, cursorPosition);
                    const lastOpenBracket = textBeforeCursor.lastIndexOf('<');
                    const lastCloseBracket = textBeforeCursor.lastIndexOf('>');

                    if (lastOpenBracket > lastCloseBracket) {
                        const textAfterCursor = currentText.substring(cursorPosition);
                        const nextCloseBracket = textAfterCursor.indexOf('>');
                        if (nextCloseBracket !== -1) {
                            insertionPosition = cursorPosition + nextCloseBracket + 1;
                            wasAdjusted = true;
                            console.log(`Cursor inside <>, adjusting insertion point to after > at ${insertionPosition}`);
                        }
                    }

                    const maxNumbers = getMaxDriverNumbers(currentText);
                    // Corrected regex - double escapes not needed in string literal here
                    const driverRegex = /(row\d+\s*=\s*")([A-Z]+)(\d*)(\|)/g;
                    const nextNumbers = { ...maxNumbers };

                    codeToAdd = codeToAdd.replace(driverRegex, (match, rowPart, prefix, existingNumberStr, pipePart) => {
                        nextNumbers[prefix] = (nextNumbers[prefix] || 0) + 1;
                        const newNumber = nextNumbers[prefix];
                        const replacement = `${rowPart}${prefix}${newNumber}${pipePart}`;
                        console.log(`Replacing driver: '${prefix}${existingNumberStr || ''}|' with '${prefix}${newNumber}|'`);
                        return replacement;
                    });

                    console.log("Modified code to add:", codeToAdd);

                    const textAfterInsertion = currentText.substring(insertionPosition);
                    let textBeforeFinal = "";
                    let searchStartIndex = insertionPosition;

                    if (!wasAdjusted) {
                        const textBeforeInsertion = currentText.substring(0, insertionPosition);
                        let tempSearchStart = cursorPosition - 1;
                        while (tempSearchStart >= 0) {
                            const char = textBeforeCursor[tempSearchStart];
                             // CORRECTED REGEX IN ONCLICK:
                            if (/\s|\n|>|<|;|\|/.test(char)) {
                                tempSearchStart++;
                                break;
                            }
                            tempSearchStart--;
                        }
                        if (tempSearchStart < 0) tempSearchStart = 0;

                        searchStartIndex = tempSearchStart;
                        const searchTermToRemove = textBeforeCursor.substring(searchStartIndex, cursorPosition);
                        console.log(`Attempting to replace term: '${searchTermToRemove}' starting at index ${searchStartIndex}`);
                        textBeforeFinal = currentText.substring(0, searchStartIndex);
                    } else {
                         textBeforeFinal = currentText.substring(0, insertionPosition);
                         searchStartIndex = insertionPosition;
                    }

                    const firstNewlineIndexInSuffix = textAfterInsertion.indexOf('\n');
                    let remainderOfOriginalLine = "";
                    let subsequentLines = "";

                    if (firstNewlineIndexInSuffix === -1) {
                        remainderOfOriginalLine = textAfterInsertion;
                    } else {
                        remainderOfOriginalLine = textAfterInsertion.substring(0, firstNewlineIndexInSuffix);
                        subsequentLines = textAfterInsertion.substring(firstNewlineIndexInSuffix);
                    }

                    const newText = textBeforeFinal +
                                    codeToAdd +
                                    (remainderOfOriginalLine.length > 0 ? '\n' : '') +
                                    remainderOfOriginalLine +
                                    subsequentLines;

                    codesTextarea.value = newText;

                    const newCursorPosition = (textBeforeFinal + codeToAdd).length;
                    codesTextarea.focus();
                    codesTextarea.setSelectionRange(newCursorPos, newCursorPos);

                    dynamicSuggestionsContainer.innerHTML = '';
                    dynamicSuggestionsContainer.style.display = 'none';
                    highlightedSuggestionIndex = -1;
                    currentSuggestions = [];
                };

                suggestionDiv.onmouseover = () => {
                    updateHighlight(i);
                };

                dynamicSuggestionsContainer.appendChild(suggestionDiv);
            });
            console.log("[showSuggestionsForTerm] Setting suggestions display to 'block'");

            const rect = codesTextarea.getBoundingClientRect();
            dynamicSuggestionsContainer.style.width = codesTextarea.offsetWidth + 'px';
            dynamicSuggestionsContainer.style.top = (rect.bottom + window.scrollY) + 'px';
            dynamicSuggestionsContainer.style.left = (rect.left + window.scrollX) + 'px';
            dynamicSuggestionsContainer.style.display = 'block';
        } else {
            console.log("[showSuggestionsForTerm] No suggestions found, hiding container.");
            dynamicSuggestionsContainer.style.display = 'none';
        }
    };

    if (codesTextarea && dynamicSuggestionsContainer) {
        codesTextarea.oninput = (event) => {
             if (!event.isTrusted || !dynamicSuggestionsContainer) {
                 return;
             }
            const cursorPosition = codesTextarea.selectionStart;
            const currentText = codesTextarea.value;

            const textBeforeCursor = currentText.substring(0, cursorPosition);
            const lastOpenBracket = textBeforeCursor.lastIndexOf('<');
            const lastCloseBracket = textBeforeCursor.lastIndexOf('>');
            let isInsideBrackets = false;
            if (lastOpenBracket > lastCloseBracket) {
                const textAfterCursor = currentText.substring(cursorPosition);
                const nextCloseBracket = textAfterCursor.indexOf('>');
                if (nextCloseBracket !== -1 ) {
                    isInsideBrackets = true;
                }
            }

            if (isInsideBrackets) {
                console.log("[Textarea Input] Cursor inside <>, hiding suggestions.");
                dynamicSuggestionsContainer.style.display = 'none';
                highlightedSuggestionIndex = -1;
                currentSuggestions = [];
            } else {
                let searchStart = cursorPosition - 1;
                while (searchStart >= 0) {
                    const char = textBeforeCursor[searchStart];
                    // CORRECTED REGEX: No double escapes needed
                    if (/\s|\n|>|<|;|\|/.test(char)) {
                        searchStart++;
                        break;
                    }
                    searchStart--;
                }
                if (searchStart < 0) searchStart = 0;

                console.log(`[Textarea Input Debug] cursorPosition: ${cursorPosition}, calculated searchStart: ${searchStart}, char at searchStart: '${searchStart < currentText.length ? textBeforeCursor[searchStart] : 'EOF'}'`);

                const searchTerm = textBeforeCursor.substring(searchStart, cursorPosition);
                const trimmedSearchTerm = searchTerm.trim();

                if (trimmedSearchTerm.length === 0 || !/^[a-zA-Z]/.test(trimmedSearchTerm)) {
                     if (trimmedSearchTerm.length === 0) {
                         console.log(`[Textarea Input] Hiding suggestions (empty term detected immediately after delimiter)`);
                     } else {
                         console.log(`[Textarea Input] Hiding suggestions (term does not start with letter: '${searchTerm}')`);
                     }
                    dynamicSuggestionsContainer.style.display = 'none';
                    highlightedSuggestionIndex = -1;
                    currentSuggestions = [];
                } else {
                    console.log(`[Textarea Input] Cursor outside <>, potential search term: '${searchTerm}'`);
                    showSuggestionsForTerm(searchTerm);
                }
            }
        };

        codesTextarea.onkeydown = (event) => {
            if (!dynamicSuggestionsContainer || dynamicSuggestionsContainer.style.display !== 'block' || currentSuggestions.length === 0) {
                return;
            }

            const suggestionItems = dynamicSuggestionsContainer.querySelectorAll('.code-suggestion-item');
            let newIndex = highlightedSuggestionIndex;

            switch (event.key) {
                case 'ArrowDown':
                case 'ArrowUp':
                    event.preventDefault();
                    newIndex = event.key === 'ArrowDown'
                        ? (highlightedSuggestionIndex + 1) % currentSuggestions.length
                        : (highlightedSuggestionIndex - 1 + currentSuggestions.length) % currentSuggestions.length;
                    updateHighlight(newIndex);
                    break;

                case 'Enter':
                 case 'Tab':
                    event.preventDefault();
                    if (highlightedSuggestionIndex >= 0 && highlightedSuggestionIndex < suggestionItems.length) {
                        suggestionItems[highlightedSuggestionIndex].click();
                    } else if (currentSuggestions.length > 0 && suggestionItems.length > 0) {
                         suggestionItems[0].click(); // Select first if none highlighted
                    }
                    // Suggestion click handles hiding
                    break;

                case 'Escape':
                    event.preventDefault();
                    dynamicSuggestionsContainer.style.display = 'none';
                    highlightedSuggestionIndex = -1;
                    currentSuggestions = [];
                    break;

                default:
                    if (!event.ctrlKey && !event.altKey && !event.metaKey && event.key.length === 1) {
                       updateHighlight(-1);
                    }
                    break;
            }
        };

         codesTextarea.addEventListener('blur', () => {
             if (!dynamicSuggestionsContainer) return;
             setTimeout(() => {
                 if (!dynamicSuggestionsContainer.contains(document.activeElement)) {
                      dynamicSuggestionsContainer.style.display = 'none';
                      highlightedSuggestionIndex = -1;
                 }
             }, 150);
         });
    }
    // --- End of Code Suggestion Logic ---


    // ... (rest of your Office.onReady, e.g., Promise.all)

    // Make sure initialization runs after setting up modal logic
  }
});

// >>> ADDED: Function definition moved here
async function insertResponseToEditor() {
    console.log("[insertResponseToEditor] Function called.");
    if (!lastResponse) {
        console.log("[insertResponseToEditor] Exiting: lastResponse is null or empty.");
        showError('No response to insert');
        return;
    }

    console.log("[insertResponseToEditor] lastResponse:", lastResponse);

    const codesTextarea = document.getElementById('codes-textarea');
    if (!codesTextarea) {
        console.error("[insertResponseToEditor] Exiting: Could not find codes-textarea.");
        showError('Could not find the code editor textarea');
        return;
    }

    console.log("[insertResponseToEditor] Found codesTextarea.");

    try {
        // Check for valid cursor position FIRST
        if (lastEditorCursorPosition === null) {
             console.log("[insertResponseToEditor] Exiting: lastEditorCursorPosition is null.");
             showError("Please click in the code editor first to set the insertion point.");
             return;
        }

        let responseText = "";
        // >>> MODIFIED: Extract <...> strings and ensure each is on a new line
        let codeStringsToInsert = [];
        if (Array.isArray(lastResponse)) {
            // Filter out empty strings just in case
            codeStringsToInsert = lastResponse.filter(item => typeof item === 'string' && item.trim().length > 0);
        } else if (typeof lastResponse === 'string') {
            const matches = lastResponse.match(/<[^>]+>/g); // Find all <...> patterns
            if (matches) {
                codeStringsToInsert = matches;
            } else if (lastResponse.trim().length > 0) {
                 // Fallback: If it's a string but no <...> found, maybe insert the whole string?
                 // For now, let's only insert if <...> are found based on the requirement.
                 console.log("[insertResponseToEditor] lastResponse is a string but no <...> tags found.");
            }
        } else {
            // Log if the format is unexpected
             console.warn("[insertResponseToEditor] lastResponse is not an array or string:", lastResponse);
        }

        if (codeStringsToInsert.length === 0) {
            showMessage("No code strings found in the response to insert.");
            console.log("[insertResponseToEditor] No <...> strings extracted from lastResponse.");
            return;
        }

        // Join the extracted strings, each on its own line
        responseText = codeStringsToInsert.join('\n');
        // <<< END MODIFIED BLOCK

        if (!responseText) {
            showMessage("Response is empty, nothing to insert.");
            return;
        }

        const currentText = codesTextarea.value;
        // Define insertionPoint using the validated cursor position
        const insertionPoint = lastEditorCursorPosition;

        // >>> ADDED: Check for <TAB; prefix if missing
        const tabPrefix = '<TAB; ';
        const defaultTabString = '<TAB; label1="Calcs";>';
        let addDefaultTab = false;
        if (!currentText.includes(tabPrefix) && !responseText.includes(tabPrefix)) {
            addDefaultTab = true;
            console.log("[insertResponseToEditor] Neither editor nor response contains '<TAB; '. Prepending default tab.");
        }
        // <<< END ADDED CHECK

        // Validate insertionPoint is within bounds (safety check)
        if (insertionPoint < 0 || insertionPoint > currentText.length) {
             console.error(`[insertResponseToEditor] Invalid insertionPoint: ${insertionPoint}, currentText length: ${currentText.length}`);
             showError("Invalid cursor position detected. Please click in the editor again.");
             lastEditorCursorPosition = null; // Reset invalid position
             return;
        }

        const textBefore = currentText.substring(0, insertionPoint);
        const textAfter = currentText.substring(insertionPoint);

        // Insert the response, adding a newline before if inserting mid-text and not at the start or after a newline
        let textToInsert = (addDefaultTab ? defaultTabString + '\n' : '') + responseText; // Prepend default tab if needed
        if (insertionPoint > 0 && textBefore.charAt(textBefore.length - 1) !== '\n') {
             textToInsert = '\n' + textToInsert; // Add leading newline to the combined string (tab + response)
        }
        // Add a newline after if not inserting at the very end or before an existing newline
        if (insertionPoint < currentText.length && textAfter.charAt(0) !== '\n') {
             textToInsert += '\n';
        } else if (insertionPoint === currentText.length && currentText.length > 0 && textBefore.charAt(textBefore.length - 1) !== '\n') {
             // Special case: inserting exactly at the end, ensure newline separation from previous content
             textToInsert = '\n' + textToInsert;
        }


        codesTextarea.value = textBefore + textToInsert + textAfter;

        // Update the last cursor position to be after the inserted text
        const newCursorPos = insertionPoint + textToInsert.length;
        codesTextarea.focus();
        codesTextarea.setSelectionRange(newCursorPos, newCursorPos);
        lastEditorCursorPosition = newCursorPos; // Update tracked position

        showMessage("Response inserted into editor.");
        console.log(`Response inserted at position: ${insertionPoint}`);

    } catch (error) {
        console.error("Error inserting response to editor:", error);
        showError(`Failed to insert response: ${error.message}`);
    }
}








