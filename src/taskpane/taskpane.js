/* global Office, msal */ // Global variables for Office.js and MSAL

// Import backend API integration
import backendAPI from './BackendAPI.js';
// Import user profile management
import userProfileManager, { initializeUserData, refreshUserData, canUseFeatures, getUserCredits, forceUpdateUserDisplay } from './UserProfile.js';
// Import credit system management
import { enforceFeatureAccess, useCreditForBuild, useCreditForUpdate, startSubscription } from './CreditSystem.js';
// Import chat layout debugger
import ChatLayoutDebugger from './ChatDebugger.js';

// Ensure userProfileManager is available globally
if (typeof window !== 'undefined') {
  window.userProfileManager = userProfileManager;
  
  // Debug function for testing credits display
  window.debugCredits = function() {
    console.log("🔧 Debug Credits Display");
    console.log("User data:", userProfileManager.getUserData());
    console.log("Credits:", userProfileManager.getCredits());
    console.log("Can use features:", userProfileManager.canUseFeatures());
    console.log("Backend API mock mode:", backendAPI.mockMode);
    
    // Force update the display
    forceUpdateUserDisplay();
    
    // Also show current DOM elements
    const creditsElements = document.querySelectorAll('.credits-count, #credits-count');
    console.log("Credits elements found:", creditsElements.length);
    creditsElements.forEach((el, i) => {
      console.log(`Element ${i}:`, el, "Text:", el.textContent);
    });
    
    console.log("Credits display updated!");
  };
}
// Import subscription management
import { checkSubscriptionStatus } from './SubscriptionManager.js';

import { populateCodeCollection, exportCodeCollectionToText, runCodes, processAssumptionTabs, collapseGroupingsAndNavigateToFinancials, hideColumnsAndNavigate, handleInsertWorksheetsFromBase64, parseFormulaSCustomFormula } from './CodeCollection.js';
// >>> ADDED: Import the new validation function
import { validateCodeStringsForRun } from './Validation.js';
// >>> ADDED: Import the tab string generator function
import { generateTabString } from './IndexWorksheet.js';
// >>> REMOVED: structureDatabasequeries import - now handled internally by getAICallsProcessedResponse
// import { structureDatabasequeries } from './StructureHelper.js';
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
import { handleFollowUpConversation, handleInitialConversation, handleConversation, validationCorrection, formatCodeStringsWithGPT, getAICallsProcessedResponse } from './AIcalls.js';
// >>> ADDED: Import CONFIG for URL management
import { CONFIG } from './config.js';
// >>> ADDED: Import file attachment and voice input functionality from AIModelPlanner
import { initializeFileAttachment, initializeFileAttachmentDev, initializeVoiceInput, initializeVoiceInputDev, initializeTextareaAutoResize, initializeTextareaAutoResizeDev, setAIModelPlannerOpenApiKey, currentAttachedFilesDev, formatFileDataForAIDev, removeAllAttachmentsDev } from './AIModelPlanner.js';
// >>> ADDED: Import cost tracking functionality
import { trackAPICallCost, estimateTokens } from './CostTracker.js';
// Add the codeStrings variable with the specified content
// REMOVED hardcoded codeStrings variable



//Debugging Toggle
const DEBUG = CONFIG.isDevelopment;

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
    const response = await fetch(CONFIG.getAssetUrl('assets/codestringDB.txt'));
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

// Simplified API key initialization (keys will be loaded from AIcalls.js when needed)
export async function initializeAPIKeys() {
  console.log("API keys will be loaded when needed from AIcalls.js");
    return { ...INTERNAL_API_KEYS };
}

// Conversation history storage
let conversationHistory = [];

// Add this variable to track if the current message is a response
// Removed isResponse - using isFirstMessageInSession instead

// Track if this is the first message in the current session
let isFirstMessageInSession = true;

// >>> ADDED: State for Client Chat
let conversationHistoryClient = [];
let lastResponseClient = null;
// <<< END ADDED

// Add this variable to store the last response
let lastResponse = null;

// >>> ADDED: Production Debug System
let PRODUCTION_DEBUG = false; // Toggle for production debugging - DISABLED FOR CLEAN UI
let debugContainer = null;

function productionLog(message) {
    if (!PRODUCTION_DEBUG) return;
    
    console.log(message); // Still log to console if available
    
    // Also show in UI for production debugging
    if (!debugContainer) {
        debugContainer = document.createElement('div');
        debugContainer.id = 'production-debug';
        debugContainer.style.cssText = `
            position: fixed;
            top: 10px;
            right: 10px;
            background: rgba(0,0,0,0.8);
            color: white;
            padding: 10px;
            font-size: 12px;
            font-family: monospace;
            max-width: 300px;
            max-height: 200px;
            overflow-y: auto;
            z-index: 9999;
            border-radius: 4px;
        `;
        document.body.appendChild(debugContainer);
    }
    
    const timestamp = new Date().toLocaleTimeString();
    const logEntry = document.createElement('div');
    logEntry.textContent = `[${timestamp}] ${message}`;
    debugContainer.appendChild(logEntry);
    
    // Keep only last 10 logs
    while (debugContainer.children.length > 10) {
        debugContainer.removeChild(debugContainer.firstChild);
    }
}

// Add this variable to store the last user input for training data
let lastUserInput = null;

// Add this variable to store the first user input (conversation starter) for training data
let firstUserInput = null;

// Add variables to persist training data modal values
let persistedTrainingUserInput = null;
let persistedTrainingAiResponse = null;

// Training data modal find/replace state
let trainingModalSearchableElements = [];

// Training data modal find/replace functions
function resetTrainingModalSearchState() {
    trainingModalSearchableElements = [
        document.getElementById('training-user-input'),
        document.getElementById('training-ai-response')
    ].filter(el => el !== null); // Filter out null elements
    
    const trainingSearchStatus = document.getElementById('training-search-status');
    if (trainingSearchStatus) trainingSearchStatus.textContent = '';
    
    console.log("Training modal search state reset.");
}

function updateTrainingModalSearchStatus(message) {
    const trainingSearchStatus = document.getElementById('training-search-status');
    if (trainingSearchStatus) {
        trainingSearchStatus.textContent = message;
    }
}

function trainingModalReplaceAll() {
    const trainingFindInput = document.getElementById('training-find-input');
    const trainingReplaceInput = document.getElementById('training-replace-input');
    
    if (!trainingFindInput || !trainingReplaceInput) {
        console.error("Training modal find/replace inputs not found");
        return;
    }
    
    const searchTerm = trainingFindInput.value;
    const replaceTerm = trainingReplaceInput.value;
    
    if (!searchTerm) {
        updateTrainingModalSearchStatus("Enter search term.");
        return;
    }

    // Ensure searchable elements are up-to-date
    resetTrainingModalSearchState();

    let replacementsMade = 0;
    trainingModalSearchableElements.forEach((element, index) => {
        if (!element) return;
        
        let currentValue = element.value;
        // Escape regex special characters in search term
        const escapedSearchTerm = searchTerm.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        let newValue = currentValue.replace(new RegExp(escapedSearchTerm, 'g'), () => {
            replacementsMade++;
            return replaceTerm;
        });
        if (currentValue !== newValue) {
            element.value = newValue;
            console.log(`Training Modal Replace All: Made replacements in element ${index}`);
        }
    });

    if (replacementsMade > 0) {
        updateTrainingModalSearchStatus(`Replaced ${replacementsMade} occurrence(s).`);
    } else {
        updateTrainingModalSearchStatus(`"${searchTerm}" not found.`);
    }
}

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

// >>> ADDED: Handle file attachment and conversion for Attach Actuals button
async function handleAttachActuals() {
    console.log('[handleAttachActuals] Attach Actuals button clicked');
    
    // Create file input element
    const fileInput = document.createElement('input');
    fileInput.type = 'file';
    fileInput.accept = '.xlsx,.xlsm,.csv';
    fileInput.style.display = 'none';
    
    // Add file input to document
    document.body.appendChild(fileInput);
    
    // Handle file selection
    fileInput.addEventListener('change', async function(event) {
        const file = event.target.files[0];
        if (!file) {
            console.log('[handleAttachActuals] No file selected');
            return;
        }
        
        console.log(`[handleAttachActuals] File selected: ${file.name} (${file.type}, ${file.size} bytes)`);
        
        try {
            // Show loading state
            const attachActualsButton = document.getElementById('attach-actuals-button');
            if (attachActualsButton) {
                attachActualsButton.disabled = true;
                attachActualsButton.innerHTML = '<span class="ms-Button-label">Processing...</span>';
            }
            
            // Read the file
            const arrayBuffer = await readFileAsArrayBuffer(file);
            
            let csvData = '';
            
            // Check if it's an Excel file (XLSX or XLSM) or CSV
            const fileExtension = file.name.toLowerCase().split('.').pop();
            
            if (fileExtension === 'xlsx' || fileExtension === 'xlsm') {
                console.log('[handleAttachActuals] Converting Excel file to CSV...');
                
                // Use SheetJS to read the Excel file
                const workbook = XLSX.read(arrayBuffer, { type: 'array' });
                
                // Get the first worksheet
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                
                // Convert worksheet to CSV
                csvData = XLSX.utils.sheet_to_csv(worksheet);
                
                console.log(`[handleAttachActuals] Excel conversion completed`);
                
            } else if (fileExtension === 'csv') {
                console.log('[handleAttachActuals] File is already CSV, processing...');
                
                // File is already CSV, just read as text
                csvData = new TextDecoder().decode(arrayBuffer);
                
            } else {
                throw new Error('Unsupported file type. Please select an XLSX, XLSM, or CSV file.');
            }
            
            // Process with Claude API
            console.log('[handleAttachActuals] Sending CSV data to Claude API for ACTUALS code generation...');
            await processActualsWithClaude(csvData, file.name);
            
            console.log(`[handleAttachActuals] File processing completed successfully`);
            
        } catch (error) {
            console.error('[handleAttachActuals] Error processing file:', error);
            displayActualsError(`Error processing file: ${error.message}`);
        } finally {
            // Reset button state
            const attachActualsButton = document.getElementById('attach-actuals-button');
            if (attachActualsButton) {
                attachActualsButton.disabled = false;
                attachActualsButton.innerHTML = '<span class="ms-Button-label">Attach Actuals</span>';
            }
            
            // Clean up file input
            document.body.removeChild(fileInput);
        }
    });
    
    // Trigger file dialog
    fileInput.click();
}

// Helper function to read file as ArrayBuffer
function readFileAsArrayBuffer(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = function(e) {
            resolve(e.target.result);
        };
        reader.onerror = function(e) {
            reject(new Error('Failed to read file'));
        };
        reader.readAsArrayBuffer(file);
    });
}

// Extract financial items from code editor content
function extractFinancialItems() {
    console.log('[extractFinancialItems] Extracting financial items from code editor...');
    
    try {
        // Get the codes textarea content
        const codesTextarea = document.getElementById('codes-textarea');
        if (!codesTextarea) {
            console.warn('[extractFinancialItems] Could not find codes-textarea element');
            return [];
        }
        
        const codesContent = codesTextarea.value.trim();
        if (!codesContent) {
            console.log('[extractFinancialItems] Codes textarea is empty');
            return [];
        }
        
        console.log('[extractFinancialItems] Parsing codes content for financial items...');
        
        const financialItems = [];
        
        // Split content by code boundaries (look for < at start and > at end)
        const codeMatches = codesContent.match(/<[^>]*>/g);
        
        if (!codeMatches) {
            console.log('[extractFinancialItems] No code patterns found');
            return [];
        }
        
        console.log(`[extractFinancialItems] Found ${codeMatches.length} code patterns to analyze`);
        
        codeMatches.forEach((codeString, index) => {
            try {
                // Look for row parameters that contain (F) with a value
                const rowMatches = codeString.match(/row\d+\s*=\s*"([^"]+)"/g);
                
                if (rowMatches) {
                    rowMatches.forEach(rowMatch => {
                        // Extract the row content between quotes
                        const rowContentMatch = rowMatch.match(/row\d+\s*=\s*"([^"]+)"/);
                        if (rowContentMatch && rowContentMatch[1]) {
                            const rowContent = rowContentMatch[1];
                            
                            // Split by | to get parameters
                            const parts = rowContent.split('|');
                            
                            let labelValue = '';
                            let fValue = '';
                            
                            // Extract (L) and (F) parameters
                            parts.forEach(part => {
                                part = part.trim();
                                if (part.includes('(L)')) {
                                    // Extract the label before (L)
                                    labelValue = part.replace('(L)', '').trim();
                                } else if (part.includes('(F)')) {
                                    // Extract the F value before (F)
                                    fValue = part.replace('(F)', '').trim();
                                }
                            });
                            
                            // If both label and F value exist and F is not empty, add to financial items
                            if (labelValue && fValue && fValue !== '') {
                                // Clean up the label (remove ~ symbols and other formatting)
                                const cleanLabel = labelValue.replace(/^~+/, '').replace(/~+$/, '').trim();
                                
                                if (cleanLabel && !financialItems.includes(cleanLabel)) {
                                    financialItems.push(cleanLabel);
                                    console.log(`[extractFinancialItems] Found financial item: "${cleanLabel}" with F value: "${fValue}"`);
                                }
                            }
                        }
                    });
                }
            } catch (parseError) {
                console.warn(`[extractFinancialItems] Error parsing code ${index + 1}:`, parseError);
            }
        });
        
        console.log(`[extractFinancialItems] Extracted ${financialItems.length} unique financial items:`, financialItems);
        
        // Always add fallback categories for items that don't have specific matches
        const fallbackItems = ["Other Income/(Expense)", "Other Assets", "Other Liabilities"];
        fallbackItems.forEach(item => {
            if (!financialItems.includes(item)) {
                financialItems.push(item);
            }
        });
        
        console.log(`[extractFinancialItems] Final list with fallback items (${financialItems.length} total):`, financialItems);
        return financialItems;
        
    } catch (error) {
        console.error('[extractFinancialItems] Error extracting financial items:', error);
        return [];
    }
}

// Process CSV data and generate ACTUALS code directly
async function processActualsWithClaude(csvData, filename) {
    try {
        console.log('[processActualsWithClaude] Processing CSV data and generating ACTUALS code directly...');
        
        // Extract financial items from the code editor
        const financialItems = extractFinancialItems();
        console.log('[processActualsWithClaude] Extracted financial items:', financialItems);
        
        // Prepare system prompt for ACTUALS code generation
        const systemPrompt = `You are a financial data processing specialist. Convert the provided CSV financial statement data into a single ACTUALS code.

ACTUALS CODE FORMAT: <ACTUALS; values="Description|Amount|Date|Category*Description|Amount|Date|Category*...";>

RULES:
- Column 1 (Description): Financial line item name
- Column 2 (Amount): Monetary value (positive for revenue/assets, negative for expenses)  
- Column 3 (Date): Use provided date, will be auto-converted to month-end
- Column 4 (Category): DEFAULT to same name as Column 1 (Description). Only use different categories if they exist in the Financial Items list below
- Use "|" to separate columns, "*" to separate rows
- IGNORE subtotals (Total Revenue, Total Expenses, etc.)
- FOCUS ONLY on line items
- NO explanations or additional text - ONLY output the ACTUALS code

FINANCIAL ITEMS LIST:
${financialItems.length > 0 ? financialItems.map((item, index) => `${index + 1}. ${item}`).join('\n') : 'None found - use fallback categories: Other Income/(Expense), Other Assets, Other Liabilities'}`;

        console.log('[processActualsWithClaude] Calling Claude API for ACTUALS code generation...');
        
        // Import necessary functions from AIcalls.js
        const { callClaudeAPI, initializeAPIKeys } = await import('./AIcalls.js');
        
        // Ensure API keys are initialized
        await initializeAPIKeys();
        
        // Prepare the user message with CSV data
        const userMessage = `Process this financial data from file "${filename}" and generate the ACTUALS code:\n\n${csvData}`;
        
        // Prepare messages for Claude API
        const messages = [
            { role: "system", content: systemPrompt },
            { role: "user", content: userMessage }
        ];
        
        console.log('[processActualsWithClaude] Sending request to Claude API...');
        
        // Call Claude API
        let claudeResponse = "";
        for await (const contentPart of callClaudeAPI(messages, {
            model: "claude-sonnet-4-20250514",
            temperature: 0.3,
            stream: false,
            caller: "processActualsWithClaude"
        })) {
            claudeResponse += contentPart;
        }
        
        console.log('[processActualsWithClaude] Claude API response received');
        console.log('[processActualsWithClaude] Response:', claudeResponse);
        
        // Clean up the response to ensure it's just the ACTUALS code
        const actualsCode = claudeResponse.trim();
        
        // Display result in developer mode
        displayActualsResult(actualsCode, filename);
        
    } catch (error) {
        console.error('[processActualsWithClaude] Error:', error);
        throw error; // Re-throw to be caught by the calling function
    }
}

// Display ACTUALS processing result in developer mode chat
function displayActualsResult(actualsCode, filename) {
    console.log('[displayActualsResult] Displaying ACTUALS result in developer mode');
    
    // Update lastResponse global variable so "Insert to Editor" button works
    lastResponse = actualsCode;
    console.log('[displayActualsResult] Updated lastResponse for Insert to Editor functionality');
    
    // Display the file processing message
    displayInDeveloperChat(`📄 Processed file: ${filename}`, true);
    
    // Display the ACTUALS code result
    displayInDeveloperChat(actualsCode, false);
}

// Display error in developer mode chat
function displayActualsError(errorMessage) {
    console.log('[displayActualsError] Displaying error in developer mode');
    
    // Display the error message
    displayInDeveloperChat(`❌ ${errorMessage}`, false);
}

// Display message in developer mode chat
function displayInDeveloperChat(message, isUser) {
    const chatLog = document.getElementById('chat-log');
    const welcomeMessage = document.getElementById('welcome-message');
    if (welcomeMessage) {
        welcomeMessage.style.display = 'none';
    }

    const messageElement = document.createElement('div');
    messageElement.className = `chat-message ${isUser ? 'user-message' : 'assistant-message'}`;
    
    const contentElement = document.createElement('p');
    contentElement.className = 'message-content';
    
    if (typeof message === 'string') {
        contentElement.textContent = message;
    } else if (Array.isArray(message)) {
        contentElement.textContent = message.join('\n');
    } else if (typeof message === 'object' && message !== null) {
        contentElement.textContent = JSON.stringify(message, null, 2);
        contentElement.style.whiteSpace = 'pre-wrap'; 
    } else {
        contentElement.textContent = String(message);
    }
    
    messageElement.appendChild(contentElement);
    chatLog.appendChild(messageElement);
    chatLog.scrollTop = chatLog.scrollHeight;
}
// <<< END ADDED

// >>> ADDED: setButtonLoading for Client Mode
function setButtonLoadingClient(isLoading) {
    console.log(`[setButtonLoadingClient] Called with isLoading: ${isLoading}`);
    const sendButton = document.getElementById('send-client');
    const loadingAnimation = document.getElementById('loading-animation-client');
    
    if (sendButton) {
        sendButton.disabled = isLoading;
    } else {
        console.warn("[setButtonLoadingClient] Could not find send button with id='send-client'");
    }
    
    if (loadingAnimation) {
        const newDisplay = isLoading ? 'flex' : 'none';
        console.log(`[setButtonLoadingClient] Found loadingAnimation element. Setting display to: ${newDisplay}`);
        loadingAnimation.style.display = newDisplay;
    } else {
        console.error("[setButtonLoadingClient] Could not find loading animation element with id='loading-animation-client'");
    }
}
// <<< END ADDED

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
// >>> REFACTORED: to accept chatLogId and welcomeMessageId
function appendMessage(content, isUser = false, chatLogId = 'chat-log', welcomeMessageId = 'welcome-message') {
    const chatLog = document.getElementById(chatLogId);
    const welcomeMessage = document.getElementById(welcomeMessageId);

    if (!chatLog) {
        console.error(`[appendMessage] Chat log element with ID '${chatLogId}' not found.`);
        return;
    }
    
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
    let userInput = document.getElementById('user-input').value.trim();
    
    // Check if there are attached files and include them in the prompt
    if (currentAttachedFilesDev.length > 0) {
        console.log('[handleSend] Including attached files data:', currentAttachedFilesDev.map(f => f.fileName).join(', '));
        const filesDataForAI = formatFileDataForAIDev(currentAttachedFilesDev);
        userInput = userInput + filesDataForAI;
    }
    
    if (!userInput) {
        showError('Please enter a request');
        return;
    }

    // Store the user input for training data before clearing
    lastUserInput = userInput;
    
    // Store the first user input if this is the start of the conversation
    if (isFirstMessageInSession) {
        firstUserInput = userInput;
    }

    // Check if this is the first message in the session by looking at visible chat messages BEFORE adding the new message
    const chatLog = document.getElementById('chat-log');
    const existingMessages = chatLog ? chatLog.querySelectorAll('.chat-message') : [];
    isFirstMessageInSession = existingMessages.length === 0;
    console.log(`[handleSend] Found ${existingMessages.length} existing chat messages. isFirstMessageInSession: ${isFirstMessageInSession}`);

    // Add user message to chat (include file attachment info in display)
    let displayMessage = document.getElementById('user-input').value.trim();
    if (currentAttachedFilesDev.length > 0) {
        const fileNames = currentAttachedFilesDev.map(f => f.fileName).join(', ');
        displayMessage += ` 📎 (${currentAttachedFilesDev.length} file${currentAttachedFilesDev.length !== 1 ? 's' : ''}: ${fileNames})`;
    }
    appendMessage(displayMessage, true);
    
    // Clear input and reset textarea height
    const userInputElement = document.getElementById('user-input');
    userInputElement.value = '';
    userInputElement.style.height = '24px';
    userInputElement.classList.remove('scrollable');
    
    // Clear attachments after sending
    removeAllAttachmentsDev();

    setButtonLoading(true);
    const progressMessageDiv = document.createElement('div');
    progressMessageDiv.className = 'chat-message assistant-message';
    const progressMessageContent = document.createElement('p');
    progressMessageContent.className = 'message-content';
    progressMessageContent.textContent = 'Analyzing your request...';
    progressMessageDiv.appendChild(progressMessageContent);
    chatLog.appendChild(progressMessageDiv);
    chatLog.scrollTop = chatLog.scrollHeight;
    
    try {
        let conversationResult;
        
        if (isFirstMessageInSession) {
            // First message in conversation: Use self-contained getAICallsProcessedResponse
            console.log("Starting getAICallsProcessedResponse (initial - with prompt module determination)");
            
            // Create progress callback function
            const progressCallback = (message) => {
                console.log("Progress:", message);
                progressMessageContent.textContent = message;
                chatLog.scrollTop = chatLog.scrollHeight;
            };
            
            // Set initial progress message
            progressMessageContent.textContent = 'Processing with AI (including database queries and prompt module analysis)...';
            chatLog.scrollTop = chatLog.scrollHeight;
            
            // Use the self-contained function that handles all processing internally
            const responseArray = await getAICallsProcessedResponse(userInput, progressCallback);
            
            // Create conversation result compatible with existing code
            conversationResult = {
                response: responseArray,
                history: [...conversationHistory, ['human', userInput], ['assistant', responseArray.join('\n')]]
            };
        } else {
            // Follow-up conversation: skip database queries, use simple prompt
            console.log("Starting handleConversation (follow-up - no database queries)");
            progressMessageContent.textContent = 'Processing follow-up response...';
            chatLog.scrollTop = chatLog.scrollHeight;
            
            conversationResult = await handleConversation(userInput, conversationHistory);
        }
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

        // >>> REMOVED: Redundant FormatGPT call - FormatGPT is already called within handleInitialConversation() in AIcalls.js
        // This was causing FormatGPT to be called twice for initial conversations
        console.log("FormatGPT is handled within the conversation flow in AIcalls.js");

        // Update progress message to show completion
        progressMessageContent.textContent = 'Finalizing response...';
        chatLog.scrollTop = chatLog.scrollHeight;

        // Store the final response array for Excel writing
        lastResponse = responseArray;

        // Mark that we've processed the first message in this session
        isFirstMessageInSession = false;

        // Remove the progress message and add the final assistant message
        chatLog.removeChild(progressMessageDiv);
        appendMessage(responseArray.join('\n'));
        
    } catch (error) {
        console.error("Error in handleSend:", error);
        showError(error.message);
        // Remove progress message and add error message to chat
        if (chatLog.contains(progressMessageDiv)) {
            chatLog.removeChild(progressMessageDiv);
        }
        appendMessage(`Error: ${error.message}`);
    } finally {
        setButtonLoading(false);
    }
}

// >>> ADDED: handleSend for Client Mode (Simplified)
async function handleSendClient() {
            // Disabled - CSS flexbox layout handles everything now
        // if (window.chatDebugger) {
        //     window.chatDebugger.forceFix();
        // }
    
    const userInputElement = document.getElementById('user-input-client');
    if (!userInputElement) {
        console.error("[handleSendClient] User input element 'user-input-client' not found.");
        return;
    }
    const userInput = userInputElement.value.trim();
    
    if (!userInput) {
        alert('Please enter a request (Client Mode)'); 
        return;
    }

    // Store the user input for training data before clearing
    lastUserInput = userInput;
    
    // Store the first user input if this is the start of the conversation in client mode
    if (conversationHistoryClient.length === 0) {
        firstUserInput = userInput;
    }

    // >>> ADDED: Hide welcome section and activate conversation layout on first message
    const welcomeSection = document.querySelector('.welcome-section');
    const clientChatContainer = document.getElementById('client-chat-container');
    const chatLogClient = document.getElementById('chat-log-client');
    
    // Always add conversation-active class when sending a message
    if (clientChatContainer) {
        clientChatContainer.classList.add('conversation-active');
        console.log('[sendToAPIClient] Added conversation-active class to hide welcome section');
    }
    
    // Force hide welcome section as backup
    if (welcomeSection) {
        welcomeSection.style.display = 'none';
    }
    
    // Force show chat log as backup
    if (chatLogClient) {
        chatLogClient.style.display = 'block';
    }
    // <<< END ADDED

    // Append user's message to the chat log
    appendMessage(userInput, true, 'chat-log-client', 'welcome-message-client');
    userInputElement.value = ''; // Clear the input field
    
    // Reset textarea height to default
    userInputElement.style.height = '44px';
    userInputElement.style.overflowY = 'hidden';
    userInputElement.classList.remove('scrollable');
    
    setButtonLoadingClient(true);

    // Get the chat log and welcome message elements
    const welcomeMessageClient = document.getElementById('welcome-message-client');

    if (!chatLogClient) {
        console.error("[handleSendClient] Chat log element 'chat-log-client' not found.");
        setButtonLoadingClient(false);
        return;
    }

    // Hide welcome message if it's visible
    if (welcomeMessageClient) {
        welcomeMessageClient.style.display = 'none';
    }

    // Create assistant's message container
    const assistantMessageDiv = document.createElement('div');
    assistantMessageDiv.className = 'chat-message assistant-message';
    const assistantMessageContent = document.createElement('p');
    assistantMessageContent.className = 'message-content';
    assistantMessageContent.textContent = ''; // Start with empty content
    assistantMessageDiv.appendChild(assistantMessageContent);
    chatLogClient.appendChild(assistantMessageDiv);
    chatLogClient.scrollTop = chatLogClient.scrollHeight; // Scroll to bottom

    let fullAssistantResponse = "";

    try {
        // Prepare messages for OpenAI API
        // Add current user input. For a more complete conversation, you'd include previous messages from conversationHistoryClient
        const messages = [
            // Example: Add system prompt if you have one
            // { role: "system", content: "You are a helpful assistant." },
            ...conversationHistoryClient.map(item => ({ role: "user", content: item.user })),
            ...conversationHistoryClient.map(item => ({ role: "assistant", content: item.assistant })),
            { role: "user", content: userInput }
        ];
        
        console.log("[handleSendClient] Calling OpenAI with stream enabled. Messages:", messages);

        // Call Claude API with streaming
        // Using Claude 4 via callOpenAI (configured in AIcalls.js)
        const stream = await callOpenAI(messages, { stream: true, model: "claude-sonnet-4-20250514", useClaudeAPI: true });

        for await (const chunk of stream) {
            const content = chunk.choices && chunk.choices[0]?.delta?.content;
            if (content) {
                fullAssistantResponse += content;
                assistantMessageContent.textContent += content; // Append new content
                chatLogClient.scrollTop = chatLogClient.scrollHeight; // Keep scrolling to bottom
            }
        }

        lastResponseClient = fullAssistantResponse;
        conversationHistoryClient.push({ user: userInput, assistant: fullAssistantResponse });

        console.log("[handleSendClient] Streaming finished. Full response:", fullAssistantResponse);

    } catch (error) {
        console.error("Error in handleSendClient during OpenAI call:", error);
        // Display error in the assistant's message bubble or a separate error message
        assistantMessageContent.textContent = `Error: ${error.message || 'Failed to get response'}`;
        // Also log to general error display if available
        showError(`Client mode error: ${error.message}`);
    } finally {
        setButtonLoadingClient(false);
    }
}
// <<< END ADDED

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
    
    // Reset the last response
    lastResponse = null;
    
    // Reset the session tracking
    isFirstMessageInSession = true;
    
    // Reset the stored user inputs
    firstUserInput = null;
    lastUserInput = null;
    
    // Clear persisted training data modal values
    persistedTrainingUserInput = null;
    persistedTrainingAiResponse = null;
    
    // Clear the input field and reset textarea height
    const userInputElement = document.getElementById('user-input');
    userInputElement.value = '';
    userInputElement.style.height = '24px';
    userInputElement.classList.remove('scrollable');
    
    console.log("Chat reset completed");
}

// >>> ADDED: resetChat for Client Mode
function resetChatClient() {
    const chatLogClient = document.getElementById('chat-log-client');
    const welcomeMessageClient = document.getElementById('welcome-message-client');
    const welcomeSection = document.querySelector('.welcome-section');
    const clientChatContainer = document.getElementById('client-chat-container');
    const userInputClient = document.getElementById('user-input-client');

    if (chatLogClient) {
        chatLogClient.innerHTML = '';
        if (welcomeMessageClient) {
            chatLogClient.appendChild(welcomeMessageClient);
            welcomeMessageClient.style.display = 'none';
        }
    }

    // Show the welcome section again
    if (welcomeSection) {
        welcomeSection.style.display = 'block';
    }

    if (clientChatContainer) {
        clientChatContainer.classList.remove('conversation-active');
    }

    if (userInputClient) {
        userInputClient.value = '';
        // userInputClient.style.height = '24px'; // Let CSS control height
        userInputClient.classList.remove('scrollable');
    }

    // Reset conversation history for client mode
    conversationHistoryClient = [];
    lastResponseClient = null;
    
    // Reset the stored user inputs for client mode too
    firstUserInput = null;
    lastUserInput = null;
    
    // Clear persisted training data modal values
    persistedTrainingUserInput = null;
    persistedTrainingAiResponse = null;
    
    console.log("Client mode chat reset.");
}
// <<< END ADDED



// Helper function to find max driver numbers in existing text
function getMaxDriverNumbers(text) {
    const maxNumbers = {};
    // UPDATED Regex: Handle both old format (V1|) and new format with labels (V1(D)|)
    // Pattern matches: row1="V1(D)|" or row1="V1|"
    const regex = /row\d+\s*=\s*"([A-Z]+)(\d*)(\([^)]+\))?\|/g;
    let match;
    console.log("Scanning text for drivers:", text.substring(0, 200) + "..."); // Log input text

    while ((match = regex.exec(text)) !== null) {
        const prefix = match[1];
        const numberStr = match[2];
        const label = match[3]; // This will be undefined for old format, or "(D)" etc. for new format
        const number = numberStr ? parseInt(numberStr, 10) : 0;
        console.log(`Found driver match: prefix='${prefix}', numberStr='${numberStr}', number=${number}, label='${label || 'none'}'`); // Log each match

        if (isNaN(number)) {
             console.warn(`Parsed NaN for number from '${numberStr}' for prefix '${prefix}'. Skipping.`);
             continue;
        }

        // Create a composite key for prefixes with labels
        const key = label ? `${prefix}${label}` : prefix;
        if (!maxNumbers[key] || number > maxNumbers[key]) {
            maxNumbers[key] = number;
            console.log(`Updated max for '${key}' to ${number}`); // Log updates
        }
    }
    if (Object.keys(maxNumbers).length === 0) {
        console.log("No existing drivers found matching the pattern.");
    }
    console.log("Final max existing driver numbers:", maxNumbers);
    return maxNumbers;
}



// NEW FUNCTION SPECIFICALLY FOR AI MODEL PLANNER OUTPUT (ALWAYS FIRST PASS)
export async function processModelCodesForPlanner(modelCodesString) {
    console.log(`[processModelCodesForPlanner] Called with ModelCodes.`);
    if (DEBUG) console.log("[processModelCodesForPlanner] Input (first 500 chars):", modelCodesString.substring(0,500));

    // Substitute <BR> with <BR; labelRow=""; row1 = "||||||||||||";>
    if (modelCodesString && typeof modelCodesString === 'string') {
        modelCodesString = modelCodesString.replace(/<BR>/g, '<BR; labelRow=""; row1 = "||||||||||||";>');
    }

    let runResult = null;

    try {
        // 0. Set calculation mode to manual
        await Excel.run(async (context) => {
            context.application.calculationMode = Excel.CalculationMode.manual;
            await context.sync();
            console.log("[processModelCodesForPlanner] Calculation mode set to manual.");
        });

        // 1. Validate all incoming codes
        if (modelCodesString.trim().length > 0) {
            console.log("[processModelCodesForPlanner] Validating ALL codes...");
            const validationErrors = await validateCodeStringsForRun(modelCodesString);
            if (validationErrors && validationErrors.length > 0) {
                const errorMsg = "Code validation failed for planner-generated codes:\n" + validationErrors.join("\n");
                console.error("[processModelCodesForPlanner] Code validation failed:", validationErrors);
                throw new Error(errorMsg);
            }
            console.log("[processModelCodesForPlanner] Code validation successful.");
        } else {
            console.log("[processModelCodesForPlanner] No codes provided by planner to validate or process. Exiting.");
            return; 
        }

                    // 2. Insert base sheets from Worksheets_4.3.25 v1.xlsx
        console.log("[processModelCodesForPlanner] Inserting base sheets from Worksheets_4.3.25 v1.xlsx...");
        // First, test if basic assets are accessible
        console.log(`[processModelCodesForPlanner] Testing basic asset accessibility...`);
        try {
            const testUrls = CONFIG.getAssetUrlsWithFallback('assets/codestringDB.txt');
            for (const testUrl of testUrls) {
                try {
                    const testResponse = await fetch(testUrl, { method: 'HEAD' });
                    console.log(`[processModelCodesForPlanner] Asset test - ${testUrl}: ${testResponse.status}`);
                    if (testResponse.ok) {
                        console.log(`[processModelCodesForPlanner] ✅ Basic assets are accessible via: ${testUrl}`);
                        break;
                    }
                } catch (e) {
                    console.log(`[processModelCodesForPlanner] Asset test failed for ${testUrl}: ${e.message}`);
                }
            }
        } catch (testError) {
            console.warn(`[processModelCodesForPlanner] Asset accessibility test failed:`, testError);
        }

        // Try multiple URLs with fallback in production
        const worksheetUrls = CONFIG.getAssetUrlsWithFallback('assets/Worksheets_4.3.25 v1.xlsx');
        console.log(`[processModelCodesForPlanner] Trying ${worksheetUrls.length} possible URLs:`, worksheetUrls);
        
        let worksheetsResponse = null;
        let lastError = null;
        
        // Enhanced fetch with proper binary handling and fallback
        for (let i = 0; i < worksheetUrls.length; i++) {
            const url = worksheetUrls[i];
            console.log(`[processModelCodesForPlanner] Attempt ${i + 1}/${worksheetUrls.length}: ${url}`);
            
            try {
                const response = await fetch(url, {
                    method: 'GET',
                    headers: {
                        'Accept': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/octet-stream,*/*'
                    },
                    cache: 'no-cache'
                });
                
                console.log(`[processModelCodesForPlanner] Response ${i + 1}: ${response.status} ${response.statusText}`);
                console.log(`[processModelCodesForPlanner] Content-Type: ${response.headers.get('content-type')}`);
                
                if (response.ok) {
                    // Quick check if this looks like binary data
                    const contentType = response.headers.get('content-type');
                    if (contentType && (contentType.includes('application/') || contentType.includes('octet-stream'))) {
                        console.log(`[processModelCodesForPlanner] ✅ Found valid binary response at URL ${i + 1}`);
                        worksheetsResponse = response;
                        break;
                    } else {
                        console.warn(`[processModelCodesForPlanner] ⚠️ URL ${i + 1} returned wrong content-type: ${contentType}`);
                        lastError = new Error(`Wrong content-type: ${contentType}`);
                    }
                } else {
                    console.warn(`[processModelCodesForPlanner] ⚠️ URL ${i + 1} returned ${response.status}`);
                    lastError = new Error(`HTTP ${response.status}: ${response.statusText}`);
                }
            } catch (fetchError) {
                console.warn(`[processModelCodesForPlanner] ⚠️ URL ${i + 1} failed:`, fetchError.message);
                lastError = fetchError;
            }
        }
        
        if (!worksheetsResponse) {
            console.warn(`[processModelCodesForPlanner] All URLs failed, attempting test with minimal Excel file...`);
            try {
                // Import is already done at top of file, just need to import the specific function
                const CodeCollectionModule = await import('./CodeCollection.js');
                const minimalBase64 = CodeCollectionModule.createMinimalExcelBase64();
                console.log(`[processModelCodesForPlanner] Testing with minimal Excel file (${minimalBase64.length} chars)...`);
                await handleInsertWorksheetsFromBase64(minimalBase64, ['TestSheet']);
                console.log(`[processModelCodesForPlanner] ✅ Minimal Excel test succeeded! The Excel API works.`);
                console.log(`[processModelCodesForPlanner] Issue is definitely with asset loading/corruption, not Excel API.`);
            } catch (testError) {
                console.error(`[processModelCodesForPlanner] ❌ Even minimal Excel test failed:`, testError);
                console.error(`[processModelCodesForPlanner] This suggests Excel API issue, not just asset loading issue.`);
            }
            //dfdf
            throw new Error(`All worksheet URLs failed. Last error: ${lastError?.message || 'Unknown error'}`);
        }
        
        console.log(`[processModelCodesForPlanner] Worksheets response status: ${worksheetsResponse.status} ${worksheetsResponse.statusText}`);
        console.log(`[processModelCodesForPlanner] Worksheets response headers:`, Object.fromEntries(worksheetsResponse.headers.entries()));
        console.log(`[processModelCodesForPlanner] Content-Type: ${worksheetsResponse.headers.get('content-type')}`);
        console.log(`[processModelCodesForPlanner] Content-Length: ${worksheetsResponse.headers.get('content-length')}`);
        
        if (!worksheetsResponse.ok) {
            console.error(`[processModelCodesForPlanner] Response not OK. Status: ${worksheetsResponse.status}`);
            console.error(`[processModelCodesForPlanner] Response text (first 200 chars):`, await worksheetsResponse.clone().text().then(t => t.substring(0, 200)));
            throw new Error(`[processModelCodesForPlanner] Worksheets_4.3.25 v1.xlsx load failed: ${worksheetsResponse.status} ${worksheetsResponse.statusText}`);
        }
        
        // Check if we're actually getting binary data
        const contentType = worksheetsResponse.headers.get('content-type');
        if (contentType && !contentType.includes('application/') && !contentType.includes('octet-stream')) {
            console.warn(`[processModelCodesForPlanner] ⚠️ Unexpected content-type: ${contentType}`);
            console.warn(`[processModelCodesForPlanner] This suggests we're not getting a binary file`);
            
            // Get the actual response as text to see what we're receiving
            const responseText = await worksheetsResponse.clone().text();
            console.error(`[processModelCodesForPlanner] Response content (first 500 chars): "${responseText.substring(0, 500)}"`);
            throw new Error(`Expected binary Excel file but got content-type: ${contentType}`);
        }
        
        const wsArrayBuffer = await worksheetsResponse.arrayBuffer();
        console.log(`[processModelCodesForPlanner] Worksheets ArrayBuffer size: ${wsArrayBuffer.byteLength} bytes`);
        
        // Verify the ArrayBuffer contains valid Excel data
        const uint8View = new Uint8Array(wsArrayBuffer);
        const firstFourBytes = Array.from(uint8View.slice(0, 4)).map(b => b.toString(16).padStart(2, '0')).join(' ');
        console.log(`[processModelCodesForPlanner] First 4 bytes of ArrayBuffer: ${firstFourBytes}`);
        
        // Check for Excel signature (PK = 50 4B in hex)
        if (uint8View[0] === 0x50 && uint8View[1] === 0x4B) {
            console.log(`[processModelCodesForPlanner] ✅ Valid Excel file signature in ArrayBuffer`);
        } else {
            console.error(`[processModelCodesForPlanner] ❌ Invalid Excel file signature in ArrayBuffer!`);
            console.error(`[processModelCodesForPlanner] Expected: 50 4B (PK), Got: ${uint8View[0].toString(16)} ${uint8View[1].toString(16)}`);
            throw new Error(`Excel file is corrupted - invalid file signature in source ArrayBuffer`);
        }
        
        // Robust base64 conversion using proven method
        console.log(`[processModelCodesForPlanner] Converting ArrayBuffer to base64...`);
        let wsBase64String;
        
        try {
            // Use the most reliable method: Uint8Array -> btoa
            const uint8Array = new Uint8Array(wsArrayBuffer);
            
            // Convert to binary string using reliable chunk processing
            let binaryString = '';
            const chunkSize = 8192; // Process in chunks to avoid call stack limits
            
            for (let i = 0; i < uint8Array.length; i += chunkSize) {
                const chunk = uint8Array.subarray(i, Math.min(i + chunkSize, uint8Array.length));
                // Use spread operator for better compatibility
                binaryString += String.fromCharCode(...chunk);
            }
            
            // Convert to base64
            wsBase64String = btoa(binaryString);
            
            console.log(`[processModelCodesForPlanner] ✅ Base64 conversion successful`);
            console.log(`[processModelCodesForPlanner] Worksheets base64 length: ${wsBase64String.length} characters`);
            
            // Verify the base64 is valid by testing decode
            const testDecode = atob(wsBase64String.substring(0, 100));
            console.log(`[processModelCodesForPlanner] ✅ Base64 decode test successful`);
            
        } catch (conversionError) {
            console.error(`[processModelCodesForPlanner] ❌ Base64 conversion failed:`, conversionError);
            throw new Error(`Failed to convert worksheets to base64: ${conversionError.message}`);
        }
        
        await handleInsertWorksheetsFromBase64(wsBase64String);
        console.log("[processModelCodesForPlanner] Base sheets (Worksheets_4.3.25 v1.xlsx) inserted.");

        // 3. Insert codes.xlsx (as runCodes depends on it)
        console.log("[processModelCodesForPlanner] Inserting Codes.xlsx...");
        const codesUrl = CONFIG.getAssetUrl('assets/Codes.xlsx');
        console.log(`[processModelCodesForPlanner] Loading codes from: ${codesUrl}`);
        const codesResponse = await fetch(codesUrl);
        console.log(`[processModelCodesForPlanner] Codes response status: ${codesResponse.status} ${codesResponse.statusText}`);
        if (!codesResponse.ok) throw new Error(`[processModelCodesForPlanner] Codes.xlsx load failed: ${codesResponse.status} ${codesResponse.statusText}`);
        
        const codesArrayBuffer = await codesResponse.arrayBuffer();
        console.log(`[processModelCodesForPlanner] Codes ArrayBuffer size: ${codesArrayBuffer.byteLength} bytes`);
        
        // Improved base64 conversion using modern method
        let codesBase64String;
        try {
            // Use FileReader for better binary handling
            const blob = new Blob([codesArrayBuffer]);
            codesBase64String = await new Promise((resolve, reject) => {
                const reader = new FileReader();
                reader.onload = () => {
                    const dataUrl = reader.result;
                    const base64 = dataUrl.split(',')[1]; // Remove data:application/octet-stream;base64, prefix
                    resolve(base64);
                };
                reader.onerror = reject;
                reader.readAsDataURL(blob);
            });
            console.log(`[processModelCodesForPlanner] Codes base64 length: ${codesBase64String.length} characters`);
            console.log(`[processModelCodesForPlanner] Codes base64 prefix: ${codesBase64String.substring(0, 50)}...`);
        } catch (conversionError) {
            console.error(`[processModelCodesForPlanner] Codes base64 conversion failed:`, conversionError);
            throw new Error(`Failed to convert codes to base64: ${conversionError.message}`);
        }
        
        await handleInsertWorksheetsFromBase64(codesBase64String, ["Codes"]); 
        console.log("[processModelCodesForPlanner] Codes.xlsx sheets inserted/updated.");
    
        // 4. Execute runCodes
        console.log("[processModelCodesForPlanner] Populating collection...");
        const collection = populateCodeCollection(modelCodesString);
        console.log(`[processModelCodesForPlanner] Collection populated with ${collection.length} code(s)`);

        if (collection.length > 0) {
            console.log("[processModelCodesForPlanner] Running codes...");
            runResult = await runCodes(collection);
            console.log("[processModelCodesForPlanner] runCodes executed. Result:", runResult);
        } else {
            console.log("[processModelCodesForPlanner] Collection is empty after population, skipping runCodes execution.");
            runResult = { assumptionTabs: [] };
        }

        // 5. Post-processing
        console.log("[processModelCodesForPlanner] Starting post-processing steps...");
        if (runResult && runResult.assumptionTabs && runResult.assumptionTabs.length > 0) {
            console.log("[processModelCodesForPlanner] Processing assumption tabs...");
            await processAssumptionTabs(runResult.assumptionTabs);
        } else {
            console.log("[processModelCodesForPlanner] No assumption tabs to process from runResult.");
        }

        console.log("[processModelCodesForPlanner] Hiding specific columns and navigating...");
        await hideColumnsAndNavigate(runResult?.assumptionTabs || []);

        // 6. Cleanup Codes sheet
        console.log("[processModelCodesForPlanner] Deleting Codes sheet...");
        await Excel.run(async (context) => {
            try {
                context.workbook.worksheets.getItem("Codes").delete();
                console.log("[processModelCodesForPlanner] Codes sheet deleted.");
            } catch (e) {
                if (e instanceof OfficeExtension.Error && e.code === Excel.ErrorCodes.itemNotFound) {
                    console.warn("[processModelCodesForPlanner] Codes sheet not found during cleanup, skipping deletion.");
                } else { 
                    console.error("[processModelCodesForPlanner] Error deleting Codes sheet during cleanup:", e);
                }
            }
            await context.sync();
        }).catch(error => { 
            console.error("[processModelCodesForPlanner] Error during Codes sheet cleanup sync:", error);
        });

        console.log("[processModelCodesForPlanner] Successfully completed.");

    } catch (error) {
        console.error("[processModelCodesForPlanner] Error during processing:", error);
        throw error; 
    } finally {
        try {
            await Excel.run(async (context) => {
                context.application.calculationMode = Excel.CalculationMode.automatic;
                await context.sync();
                console.log("[processModelCodesForPlanner] Calculation mode set to automatic.");
            });
        } catch (finalError) {
            console.error("[processModelCodesForPlanner] Error setting calculation mode to automatic:", finalError);
        }
    }
}

// Original insertSheetsAndRunCodes function should be here, UNCHANGED.
// Ensure it's not accidentally deleted or modified by the `// ... existing code ...` placeholder.
async function insertSheetsAndRunCodes() {
    const codesTextarea = document.getElementById('codes-textarea');
    if (!codesTextarea) {
        showError("Could not find the code input area. Cannot run codes.");
        return;
    }
    loadedCodeStrings = codesTextarea.value; // Update global variable

    // Substitute <BR> with <BR; labelRow=""; row1 = "||||||||||||";>
    if (loadedCodeStrings && typeof loadedCodeStrings === 'string') {
        loadedCodeStrings = loadedCodeStrings.replace(/<BR>/g, '<BR; labelRow=""; row1 = "||||||||||||";>');
    }

    try {
        localStorage.setItem('userCodeStrings', loadedCodeStrings);
        console.log("[Run Codes] Automatically saved codes from textarea to localStorage.");
    } catch (error) {
        console.error("[Run Codes] Error auto-saving codes to localStorage:", error);
        showError(`Error automatically saving codes: ${error.message}. Run may not reflect latest changes.`);
    }

    let codesToRun = loadedCodeStrings;
    let allCodeContentToProcess = ""; 
    let runResult = null; 

    // Substitute <BR> with <BR; labelRow=""; row1 = "||||||||||||";> in codesToRun as well, as it's derived from loadedCodeStrings after potential modification
    if (codesToRun && typeof codesToRun === 'string') {
        codesToRun = codesToRun.replace(/<BR>/g, '<BR; labelRow=""; row1 = "||||||||||||";>');
    }

    try {
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
                } else { throw error; } 
            }
        });

        await Excel.run(async (context) => {
            context.application.calculationMode = Excel.CalculationMode.manual;
            await context.sync();
        });

        setButtonLoading(true);
        console.log("Starting code processing...");

        if (!financialsSheetExists) {
            console.log("[Run Codes] FIRST PASS: Financials sheet not found.");
            allCodeContentToProcess = codesToRun;
            if (allCodeContentToProcess.trim().length > 0) {
                const validationErrors = await validateCodeStringsForRun(allCodeContentToProcess);
                if (validationErrors && validationErrors.length > 0) {
                    // Separate logic errors (LERR) from format errors (FERR)
                    const logicErrors = validationErrors.filter(err => err.includes('[LERR'));
                    const formatErrors = validationErrors.filter(err => err.includes('[FERR'));
                    
                    if (logicErrors.length > 0) {
                        // Logic errors are critical - stop the process
                        const errorMsg = "Logic validation failed. Please fix these errors before running:\n" + logicErrors.join("\n");
                        console.error("Logic validation failed:", logicErrors);
                        showError("Logic validation failed. See chat for details.");
                        appendMessage(errorMsg);
                        setButtonLoading(false);
                        return;
                    }
                    
                    if (formatErrors.length > 0) {
                        // Format errors are warnings - show them but continue
                        const warningMsg = "Format validation warnings:\n" + formatErrors.join("\n");
                        console.warn("Format validation warnings:", formatErrors);
                        showMessage("Format validation warnings detected. See chat for details.");
                        appendMessage(warningMsg);
                    }
                }
                console.log("Initial code validation completed.");
            } else {
                console.log("[Run Codes] No codes to validate on first pass.");
            }
            const worksheetsResponse = await fetch(CONFIG.getAssetUrl('assets/Worksheets_4.3.25 v1.xlsx'));
            if (!worksheetsResponse.ok) throw new Error(`Worksheets load failed: ${worksheetsResponse.statusText}`);
            const worksheetsArrayBuffer = await worksheetsResponse.arrayBuffer();
            console.log(`[insertSheetsAndRunCodes] Worksheets ArrayBuffer size: ${worksheetsArrayBuffer.byteLength} bytes`);
            
            // Improved base64 conversion using modern method
            let worksheetsBase64String;
            try {
                const blob = new Blob([worksheetsArrayBuffer]);
                worksheetsBase64String = await new Promise((resolve, reject) => {
                    const reader = new FileReader();
                    reader.onload = () => {
                        const dataUrl = reader.result;
                        const base64 = dataUrl.split(',')[1]; // Remove data:application/octet-stream;base64, prefix
                        resolve(base64);
                    };
                    reader.onerror = reject;
                    reader.readAsDataURL(blob);
                });
                console.log(`[insertSheetsAndRunCodes] Worksheets base64 length: ${worksheetsBase64String.length} characters`);
            } catch (conversionError) {
                console.error(`[insertSheetsAndRunCodes] Worksheets base64 conversion failed:`, conversionError);
                throw new Error(`Failed to convert worksheets to base64: ${conversionError.message}`);
            }
            
            await handleInsertWorksheetsFromBase64(worksheetsBase64String);
            console.log("Base sheets inserted.");
        } else {
            console.log("[Run Codes] SUBSEQUENT PASS: Financials sheet found.");
            
            // Process ALL codes (no change detection)
            allCodeContentToProcess = codesToRun;
            
            if (allCodeContentToProcess.trim().length > 0) {
                const validationErrors = await validateCodeStringsForRun(allCodeContentToProcess);
                if (validationErrors && validationErrors.length > 0) {
                    // Separate logic errors (LERR) from format errors (FERR)
                    const logicErrors = validationErrors.filter(err => err.includes('[LERR'));
                    const formatErrors = validationErrors.filter(err => err.includes('[FERR'));
                    
                    if (logicErrors.length > 0) {
                        // Logic errors are critical - stop the process
                        const errorMsg = "Logic validation failed. Please fix these errors before running:\n" + logicErrors.join("\n");
                        console.error("Logic validation failed:", logicErrors);
                        showError("Logic validation failed. See chat for details.");
                        appendMessage(errorMsg);
                        setButtonLoading(false);
                        return;
                    }
                    
                    if (formatErrors.length > 0) {
                        // Format errors are warnings - show them but continue
                        const warningMsg = "Format validation warnings:\n" + formatErrors.join("\n");
                        console.warn("Format validation warnings:", formatErrors);
                        showMessage("Format validation warnings detected. See chat for details.");
                        appendMessage(warningMsg);
                    }
                }
                console.log("Code validation completed.");
                
                try {
                    const codesResponse = await fetch(CONFIG.getAssetUrl('assets/Codes.xlsx'));
                    if (!codesResponse.ok) throw new Error(`Codes.xlsx load failed: ${codesResponse.statusText}`);
                    const codesArrayBuffer = await codesResponse.arrayBuffer();
                    console.log(`[insertSheetsAndRunCodes] Codes ArrayBuffer size: ${codesArrayBuffer.byteLength} bytes`);
                    
                    // Improved base64 conversion using modern method
                    let codesBase64String;
                    try {
                        const blob = new Blob([codesArrayBuffer]);
                        codesBase64String = await new Promise((resolve, reject) => {
                            const reader = new FileReader();
                            reader.onload = () => {
                                const dataUrl = reader.result;
                                const base64 = dataUrl.split(',')[1]; // Remove data:application/octet-stream;base64, prefix
                                resolve(base64);
                            };
                            reader.onerror = reject;
                            reader.readAsDataURL(blob);
                        });
                        console.log(`[insertSheetsAndRunCodes] Codes base64 length: ${codesBase64String.length} characters`);
                    } catch (conversionError) {
                        console.error(`[insertSheetsAndRunCodes] Codes base64 conversion failed:`, conversionError);
                        throw new Error(`Failed to convert codes to base64: ${conversionError.message}`);
                    }
                    
                    await handleInsertWorksheetsFromBase64(codesBase64String, ["Codes"]);
                    console.log("Codes.xlsx sheets inserted.");
                } catch (e) {
                    console.error("Failed to insert sheets from Codes.xlsx:", e);
                    showError("Failed to insert necessary sheets from Codes.xlsx. Aborting.");
                    setButtonLoading(false);
                    return;
                }
            }
        }

        if (allCodeContentToProcess.trim().length > 0) {
            const collection = populateCodeCollection(allCodeContentToProcess);
            if (collection.length > 0) {
                runResult = await runCodes(collection);
                console.log("Codes executed:", runResult);
            } else {
                 if (!runResult) runResult = { assumptionTabs: [] };
            }
        } else {
             if (!runResult) runResult = { assumptionTabs: [] };
        }
        if (runResult && runResult.assumptionTabs && runResult.assumptionTabs.length > 0) {
            await processAssumptionTabs(runResult.assumptionTabs);
        } else {
             console.log("No assumption tabs to process.");
        }
        await hideColumnsAndNavigate(runResult?.assumptionTabs || []);
        await Excel.run(async (context) => {
            try {
                context.workbook.worksheets.getItem("Codes").delete();
            } catch (e) {
                if (e instanceof OfficeExtension.Error && e.code === Excel.ErrorCodes.itemNotFound) {
                     console.warn("Codes sheet not found, skipping deletion.");
                } else { console.error("Error deleting Codes sheet:", e); }
            }
            await context.sync();
        }).catch(error => { console.error("Error during sheet cleanup:", error); });

        showMessage("Code processing finished successfully!");
    } catch (error) {
        console.error("An error occurred during the build process:", error);
        showError(`Operation failed: ${error.message || error.toString()}`);
    } finally {
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

Office.onReady(async (info) => {
  productionLog('Office.onReady started');
  
  // Initialize backend integration
  try {
    console.log("🚀 Initializing backend integration...");
    
    // Check backend health
    const backendHealthy = await backendAPI.healthCheck();
    console.log(`🌐 Backend health: ${backendHealthy ? 'OK' : 'DOWN'}`);
    
    // If user is already authenticated, initialize their data
    if (backendAPI.isAuthenticated()) {
      console.log("👤 User is authenticated, initializing data...");
      await initializeUserData();
    } else {
      console.log("👤 User not authenticated, using mock data for development");
      forceUpdateUserDisplay();
    }
    
  } catch (error) {
    console.warn("⚠️ Backend initialization failed:", error);
    console.log("🔧 Using fallback user display");
    forceUpdateUserDisplay();
  }
  
  // Try to set optimal task pane width for our sidebar layout
  try {
    // Method 1: Set preferred width via CSS (affects internal layout)
    document.documentElement.style.setProperty('--preferred-width', '500px');
    
    // Method 2: Request wider initial size if possible
    if (Office.context && Office.context.requirements && Office.context.requirements.isSetSupported('AddinCommands', '1.1')) {
      // Remove explicit width constraints to allow full task pane width usage
      document.body.style.minWidth = '';
      document.body.style.width = '';
      
      // Set a meta tag to suggest preferred width (but don't force it)
      const viewportMeta = document.createElement('meta');
      viewportMeta.name = 'viewport';
      viewportMeta.content = 'width=device-width, initial-scale=1.0';
      document.head.appendChild(viewportMeta);
    }
    
    // Method 3: Store user's preferred width in localStorage for consistency
    const preferredWidth = localStorage.getItem('taskpane-preferred-width') || '500';
    console.log(`Preferred width from storage: ${preferredWidth}px`);
    
  } catch (error) {
    console.log('Width control not available:', error);
  }
  
  if (info.host === Office.HostType.Excel) {
    productionLog('Excel host detected');
    
    // Get references to the new elements
    const startupMenu = document.getElementById('startup-menu');
    const developerModeButton = document.getElementById('developer-mode-button');
    const clientModeButton = document.getElementById('client-mode-button');
    // const authenticationButton = document.getElementById('authentication-button'); // Authentication removed
    const appBody = document.getElementById('app-body'); // Already exists, ensure it's captured
    const clientModeView = document.getElementById('client-mode-view');
    const authenticationView = document.getElementById('authentication-view');

    productionLog(`Elements found - startupMenu: ${!!startupMenu}, appBody: ${!!appBody}, clientModeView: ${!!clientModeView}`);

    // >>> DYNAMIC MODE DETECTION: Show startup menu on localhost, skip in production
    const isLocalDevelopment = window.location.hostname === 'localhost' || 
                              window.location.hostname === '127.0.0.1' ||
                              window.location.href.includes('localhost:3002');
    const FORCE_PRODUCTION_MODE = !isLocalDevelopment;
    
    productionLog(`URL: ${window.location.href}`);
    productionLog(`Hostname: ${window.location.hostname}`);
    productionLog(`isLocalDevelopment: ${isLocalDevelopment}`);
    productionLog(`FORCE_PRODUCTION_MODE: ${FORCE_PRODUCTION_MODE}`);
    
    if (FORCE_PRODUCTION_MODE) {
      // Skip startup menu and go directly to client mode
      productionLog('Attempting to show client mode...');
      showClientMode();
      productionLog('showClientMode() called');
    } else {
      // Development: Show startup menu with both developer and client mode options
      productionLog('Development mode - showing startup menu');
      if (developerModeButton) {
        developerModeButton.style.display = 'inline-block';
      }
      if (startupMenu) {
        startupMenu.style.display = 'flex';
      }
    }
    
    // Check authentication state on load
    updateAuthUI();
    // <<< END ADDED

    // Functions to switch views
    function showDeveloperMode() {
      if (startupMenu) startupMenu.style.display = 'none';
      if (appBody) appBody.style.display = 'flex'; // Matches .ms-welcome__main display if it's flex
      if (clientModeView) clientModeView.style.display = 'none';
      
      // Initialize developer mode file attachment functionality
      if (typeof initializeFileAttachmentDev === 'function') {
        initializeFileAttachmentDev();
      }
      
      // Initialize developer mode voice input functionality
      if (typeof initializeVoiceInputDev === 'function') {
        initializeVoiceInputDev();
      }
      
      // Initialize developer mode textarea auto-resize functionality
      if (typeof initializeTextareaAutoResizeDev === 'function') {
        initializeTextareaAutoResizeDev();
      }
      
      console.log("Developer Mode activated");
    }

    async function showClientMode() {
      productionLog('showClientMode function called');
      
      // Check if user is authenticated using the same method as isUserAuthenticated()
      // First check sessionStorage (current session)
      const googleUserSession = sessionStorage.getItem('googleUser');
      const googleCredentialSession = sessionStorage.getItem('googleCredential') || sessionStorage.getItem('googleToken');
      
      // Also check localStorage (persistent auth)
      const googleTokenLocal = localStorage.getItem('google_access_token');
      const msalTokenLocal = localStorage.getItem('msal_access_token');
      const apiKeyLocal = localStorage.getItem('user_api_key');
      
      // Check both session and local storage for authentication
      const isAuthenticated = (googleUserSession && googleCredentialSession) || 
                            !!(googleTokenLocal || msalTokenLocal || apiKeyLocal) ||
                            (currentUser !== null) || // Microsoft auth
                            isUserAuthenticated(); // Use the existing function
      
      productionLog(`Authentication check - Session: ${!!(googleUserSession && googleCredentialSession)}, Local: ${!!(googleTokenLocal || msalTokenLocal || apiKeyLocal)}, isUserAuthenticated: ${isUserAuthenticated()}`);
      
      if (!isAuthenticated) {
        productionLog('User not authenticated, showing authentication view');
        // Store that user wanted client mode after auth
        localStorage.setItem('post_auth_redirect', 'client-mode');
        showAuthentication();
        return;
      }
      
      productionLog('User is authenticated, proceeding to show client mode directly');
      
      productionLog(`Before changes - startupMenu display: ${startupMenu ? startupMenu.style.display : 'null'}`);
      productionLog(`Before changes - clientModeView display: ${clientModeView ? clientModeView.style.display : 'null'}`);
      
      if (startupMenu) {
        startupMenu.style.display = 'none';
        productionLog('Set startupMenu to display: none');
      } else {
        productionLog('startupMenu element not found!');
      }
      
      if (appBody) {
        appBody.style.display = 'none';
        productionLog('Set appBody to display: none');
      } else {
        productionLog('appBody element not found!');
      }
      
      if (clientModeView) {
        clientModeView.style.display = 'flex';
        productionLog('Set clientModeView to display: flex with row direction for sidebar layout');
        
                // Disabled - CSS flexbox layout handles everything now
        // setTimeout(() => {
        //     if (window.chatDebugger) {
        //         window.chatDebugger.debugAndFix();
        //     }
        // }, 500);
      } else {
        productionLog('clientModeView element not found!');
      }
      
      productionLog(`After changes - startupMenu display: ${startupMenu ? startupMenu.style.display : 'null'}`);
      productionLog(`After changes - clientModeView display: ${clientModeView ? clientModeView.style.display : 'null'}`);
      
      // >>> PRIORITY: Initialize authentication and update UI
      try {
        // Initialize MSAL if not already done
        if (typeof msal !== 'undefined' && !msalInstance) {
          const msalInitialized = initializeMSAL();
          productionLog(`MSAL initialization result: ${msalInitialized}`);
        }
        
        // Check authentication state for Microsoft auth
        if (typeof msal !== 'undefined' && msalInstance) {
          checkAuthState(); // Check if user is already signed in
          productionLog('Authentication state checked for client mode');
        }
        
        // User IS authenticated (we already checked above)
        // But we need to check if we have backend tokens or need to authenticate with backend
        const hasBackendToken = sessionStorage.getItem('backend_access_token') || 
                               localStorage.getItem('backend_access_token');
        
        if (!hasBackendToken) {
          productionLog('User has Google auth but no backend token - authenticating with backend');
          
          // Try to authenticate with backend using stored Google credentials
          const googleIdToken = sessionStorage.getItem('googleCredential') || 
                               localStorage.getItem('googleCredential');
          
          if (googleIdToken) {
            try {
              productionLog('Authenticating with backend using stored Google token');
              const backendAuthResult = await backendAPI.signInWithGoogle(googleIdToken);
              productionLog('Backend authentication successful');
              
              // Now show the authenticated interface
              showAuthenticatedInterface();
              productionLog('User authenticated - showing normal interface');
            } catch (error) {
              console.error('Failed to authenticate with backend:', error);
              productionLog('Backend authentication failed - showing authentication page');
              // If backend auth fails, show authentication page
              localStorage.setItem('post_auth_redirect', 'client-mode');
              showAuthentication();
              return;
            }
          } else {
            productionLog('No Google ID token found - showing authentication page');
            // No ID token to authenticate with backend, show auth page
            localStorage.setItem('post_auth_redirect', 'client-mode');
            showAuthentication();
            return;
          }
        } else {
          // We have backend tokens, show normal interface
          showAuthenticatedInterface();
          productionLog('User authenticated with backend - showing normal interface');
        }
        
      } catch (error) {
        console.error('Error in authentication initialization during client mode:', error);
        // Continue anyway since user is authenticated
        showAuthenticatedInterface();
      }
      // <<< END PRIORITY AUTHENTICATION CHECK
      
      // Initialize file attachment functionality (guarded to avoid duplicate bindings)
      if (typeof initializeFileAttachment === 'function' && !window.__clientAttachInitialized) {
        initializeFileAttachment();
        window.__clientAttachInitialized = true;
        productionLog('initializeFileAttachment called');
      }
      // Ensure attachment listeners are bound even if DOM timing changes
      initClientAttachmentWithRetry();
      
      // Initialize voice input functionality
      if (typeof initializeVoiceInput === 'function') {
        initializeVoiceInput();
        productionLog('initializeVoiceInput called');
      }
      
      // Initialize textarea auto-resize functionality
      if (typeof initializeTextareaAutoResize === 'function') {
        initializeTextareaAutoResize();
        productionLog('initializeTextareaAutoResize called');
      }
      
      // Sidebar navigation removed - no longer needed
      // initializeSidebarNavigation();
      
      productionLog("Client Mode activation completed");
    }

    // >>> ADDED: Authentication-first helper functions
    function showAuthenticationFirst() {
      productionLog('showAuthenticationFirst called - prioritizing authentication');
      
      // Keep the interface normal but hide user info since not authenticated
      const userInfo = document.getElementById('user-info');
      if (userInfo) {
        userInfo.style.display = 'none';
        productionLog('User info hidden');
      }
      
      // Show the sign-in button in sidebar
      const signInButton = document.getElementById('sign-in-button');
      if (signInButton) {
          signInButton.style.display = 'flex';
        productionLog('Sidebar sign-in button shown');
      }
      
      // Replace the chat input area with a signup/signin button
      replaceInputWithAuthButton();
      
      productionLog('Authentication-first UI setup completed');
    }
    
    function showAuthenticatedInterface() {
      productionLog('showAuthenticatedInterface called - user is authenticated');
      
      // Hide the sign-in button
      const signInButton = document.getElementById('sign-in-button');
      if (signInButton) {
        signInButton.style.display = 'none';
        productionLog('Sign-in button hidden');
      }
      
      // Show user info
      const userInfo = document.getElementById('user-info');
      if (userInfo) {
        userInfo.style.display = 'flex';
        productionLog('User info shown');
      }
      
      // Restore normal chat input interface
      restoreNormalInputInterface();

      // Enable paste-to-attach on the main input
      const clientTextArea = document.getElementById('user-input-client');
      if (clientTextArea && !clientTextArea.dataset.pasteBound) {
        clientTextArea.addEventListener('paste', async (e) => {
          try {
            if (!e.clipboardData || !e.clipboardData.items) return;
            const items = e.clipboardData.items;
            for (let i = 0; i < items.length; i += 1) {
              const it = items[i];
              if (it.kind === 'file') {
                const blob = it.getAsFile();
                if (blob && blob.type && blob.type.startsWith('image/')) {
                  console.log('[Paste-Attach] Image detected from clipboard');
                  // Name the file if clipboard lacks a name
                  const ext = blob.type.includes('png') ? 'png' : (blob.type.includes('webp') ? 'webp' : 'jpg');
                  const fileName = `pasted-${Date.now()}.${ext}`;
                  const file = new File([blob], fileName, { type: blob.type });
                  const attachFn = (typeof handleFileAttachment === 'function') ? handleFileAttachment : (window.__plannerHandleFileAttachment || null);
                  if (attachFn) {
                    // Prevent text paste insertion
                    e.preventDefault();
                    await attachFn(file);
                  }
                }
              }
            }
          } catch (err) {
            console.warn('Paste-to-attach failed:', err);
          }
        });
        clientTextArea.dataset.pasteBound = 'true';
      }
      
      // Update UI with user information
      updateSignedInStatus();
      updateAuthUI(true);
      
      // Initialize user data with credits information
      console.log('🔧 Attempting to initialize user data after Google authentication...');
      initializeUserData().then(() => {
        console.log('✅ User data initialized successfully after authentication');
        
        // Transition to client mode view
        console.log('🔄 Transitioning to client mode view...');
        
        // Hide all other views
        const authView = document.getElementById('authentication-view');
        const startupMenu = document.getElementById('startup-menu');
        const appBody = document.getElementById('app-body');
        const clientModeView = document.getElementById('client-mode-view');
        
        if (authView) authView.style.display = 'none';
        if (startupMenu) startupMenu.style.display = 'none';
        if (appBody) appBody.style.display = 'none';
        
        // Show client mode view
        if (clientModeView) {
          clientModeView.style.display = 'flex';
          console.log('✅ Client mode view displayed');
        } else {
          console.warn('⚠️ Client mode view not found');
        }
      }).catch(error => {
        console.warn('Failed to initialize user data in showAuthenticatedInterface:', error);
        console.log('🔧 Using fallback display update');
        forceUpdateUserDisplay();
      });
      
      productionLog('Authenticated interface setup completed');
    }
    
    function replaceInputWithAuthButton() {
      productionLog('Replacing input area with authentication button');
      
      // Find the input area wrapper
      const inputWrapper = document.getElementById('client-input-wrapper');
      if (!inputWrapper) {
        productionLog('Input wrapper not found');
        return;
      }
      
      // Hide the normal input elements
      const inputBar = document.getElementById('chatgpt-input-bar');
      const attachedFiles = document.getElementById('attached-files-container');
      const attachedFilesList = document.getElementById('attached-files-list');
      const loadingAnimation = document.getElementById('loading-animation-client');
      
      if (inputBar) {
        inputBar.style.display = 'none';
        productionLog('Input bar hidden');
      }
      if (attachedFiles) {
        attachedFiles.style.display = 'none';
        productionLog('Attached files hidden');
      }
      if (loadingAnimation) {
        loadingAnimation.style.display = 'none';
        productionLog('Loading animation hidden');
      }
      
      // Create or show the auth button
      let authButton = document.getElementById('main-auth-button');
      if (!authButton) {
        authButton = document.createElement('div');
        authButton.id = 'main-auth-button';
        authButton.style.cssText = `
          display: flex;
          justify-content: center;
          align-items: center;
          padding: 16px;
          margin: 20px auto;
          max-width: 400px;
        `;
        
        authButton.innerHTML = `
          <button class="signup-signin-button" style="
            background: #1f2937;
            color: white;
            border: 1px solid #374151;
            padding: 12px 24px;
            border-radius: 8px;
            font-size: 14px;
            font-weight: 500;
            cursor: pointer;
            transition: all 0.2s ease;
            box-shadow: 0 1px 2px rgba(0, 0, 0, 0.05);
            min-width: 250px;
            display: flex;
            align-items: center;
            justify-content: center;
          ">
            <svg style="width: 18px; height: 18px; margin-right: 8px;" viewBox="0 0 24 24">
              <path fill="#4285F4" d="M22.56 12.25c0-.78-.07-1.53-.2-2.25H12v4.26h5.92c-.26 1.37-1.04 2.53-2.21 3.31v2.77h3.57c2.08-1.92 3.28-4.74 3.28-8.09z"/>
              <path fill="#34A853" d="M12 23c2.97 0 5.46-.98 7.28-2.66l-3.57-2.77c-.98.66-2.23 1.06-3.71 1.06-2.86 0-5.29-1.93-6.16-4.53H2.18v2.84C3.99 20.53 7.7 23 12 23z"/>
              <path fill="#FBBC05" d="M5.84 14.09c-.22-.66-.35-1.36-.35-2.09s.13-1.43.35-2.09V7.07H2.18C1.43 8.55 1 10.22 1 12s.43 3.45 1.18 4.93l2.85-2.22.81-.62z"/>
              <path fill="#EA4335" d="M12 5.38c1.62 0 3.06.56 4.21 1.64l3.15-3.15C17.45 2.09 14.97 1 12 1 7.7 1 3.99 3.47 2.18 7.07l3.66 2.84c.87-2.6 3.3-4.53 6.16-4.53z"/>
            </svg>
            Continue with Google
          </button>
        `;
        
        inputWrapper.appendChild(authButton);
        productionLog('Auth button created and added');
      } else {
        authButton.style.display = 'flex';
        productionLog('Auth button shown');
      }
      
      // Add click handler to the button
      const button = authButton.querySelector('.signup-signin-button');
      if (button) {
        button.onclick = handleSignInClick;
        
        // Add hover effects for Google button
        button.onmouseenter = function() {
          this.style.transform = 'translateY(-1px)';
          this.style.boxShadow = '0 4px 8px rgba(0, 0, 0, 0.15)';
          this.style.backgroundColor = '#374151';
          this.style.borderColor = '#4b5563';
        };
        button.onmouseleave = function() {
          this.style.transform = 'translateY(0)';
          this.style.boxShadow = '0 1px 2px rgba(0, 0, 0, 0.05)';
          this.style.backgroundColor = '#1f2937';
          this.style.borderColor = '#374151';
        };
        
        productionLog('Auth button handlers attached');
      }
    }
    
    function restoreNormalInputInterface() {
      productionLog('Restoring normal input interface');
      
      // Hide the auth button
      const authButton = document.getElementById('main-auth-button');
      if (authButton) {
        authButton.style.display = 'none';
        productionLog('Auth button hidden');
      }
      
      // Show the normal input elements
      const inputBar = document.getElementById('chatgpt-input-bar');
      const attachedFiles = document.getElementById('attached-files-container');
      const attachedFilesList = document.getElementById('attached-files-list');
      const loadingAnimation = document.getElementById('loading-animation-client');
      
      if (inputBar) {
        inputBar.style.display = 'block';
        productionLog('Input bar restored');
      }
      if (attachedFiles) {
        const hasItems = attachedFilesList && attachedFilesList.children && attachedFilesList.children.length > 0;
        attachedFiles.style.display = hasItems ? 'block' : 'none';
        productionLog(`Attached files ${hasItems ? 'restored (visible)' : 'hidden (no items)'}`);
      }
      // Note: loadingAnimation should remain hidden until needed
      
      productionLog('Normal input interface restored');

      // Re-bind attachment listeners after restoring input UI
      initClientAttachmentWithRetry();
    }

    // Ensures the client attach-file button has working listeners,
    // retrying briefly to handle dynamic DOM timing.
    function initClientAttachmentWithRetry() {
      if (window.__clientAttachInitRunning) {
        return;
      }
      window.__clientAttachInitRunning = true;

      let attempts = 0;
      const maxAttempts = 20; // ~4s total
      const intervalId = setInterval(() => {
        attempts += 1;
        const attachBtn = document.getElementById('attach-file-client');
        const fileInput = document.getElementById('file-input-client');

        if (attachBtn && fileInput) {
          clearInterval(intervalId);
          window.__clientAttachInitRunning = false;

          // Call canonical initializer once if not already initialized
          try {
            if (typeof initializeFileAttachment === 'function' && !window.__clientAttachInitialized) {
              initializeFileAttachment();
              window.__clientAttachInitialized = true;
              productionLog('initializeFileAttachment bound via retry helper');
            }
          } catch (err) {
            console.warn('initializeFileAttachment invocation failed, considering fallback listeners', err);
          }

          // Mark as bound so we don't attach fallback listeners if the canonical initializer handled it
          attachBtn.dataset.listenerBound = 'true';
          fileInput.dataset.listenerBound = 'true';

          // Fallback: only if canonical initializer is unavailable
          if (!window.__clientAttachInitialized) {
            if (!attachBtn.dataset.fallbackListener) {
              attachBtn.addEventListener('click', () => {
                console.log('[Client-Attach] Attach file button clicked');
                fileInput.click();
              });
              attachBtn.dataset.fallbackListener = 'true';
            }
            if (!fileInput.dataset.fallbackListener) {
              fileInput.addEventListener('change', (e) => {
                const count = e && e.target && e.target.files ? e.target.files.length : 0;
                console.log('[Client-Attach] Files selected:', count);
              });
              fileInput.dataset.fallbackListener = 'true';
            }
          }
        } else if (attempts >= maxAttempts) {
          clearInterval(intervalId);
          window.__clientAttachInitRunning = false;
          productionLog('Client attach elements not found after retries');
        }
      }, 200);
    }
    
    function handleSignInClick() {
      productionLog('Google Sign-in button clicked');
      
      // Use Google authentication instead of Microsoft
      handleGoogleSignIn();
    }
    
    function simulateAuthentication() {
      // This is a temporary function for development when Azure credentials aren't set up
      productionLog('Simulating authentication for development');
      
      // Create a mock user
      currentUser = {
        name: 'Demo User',
        username: 'demo@example.com',
        localAccountId: 'demo-user-123'
      };
      
      // Update UI to show authenticated state
      showAuthenticatedInterface();
      
      showMessage('Demo authentication successful! (For production, please configure Azure authentication)');
      productionLog('Demo authentication completed');
    }
    // <<< END ADDED AUTHENTICATION HELPERS

    let sidebarInitialized = false;
    // Sidebar navigation functionality
    function initializeSidebarNavigation() {
      if (sidebarInitialized) {
        console.log("Sidebar navigation already initialized. Skipping.");
        return;
      }
      const chatTab = document.getElementById("chat-tab");
      const subscriptionTab = document.getElementById("subscription-tab");
      const chatPanel = document.getElementById('client-chat-container');
      const subscriptionPanel = document.getElementById('subscription-panel');

          function switchToTab(activeTab, activePanel) {
        // Remove active class from all tabs
        document.querySelectorAll('.sidebar-tab').forEach(tab => {
            tab.classList.remove('active');
        });
        
        // Hide both panels by setting inline display styles
        console.log('=== BEFORE HIDING PANELS ===');
        console.log('Chat panel display:', chatPanel ? chatPanel.style.display : 'null');
        console.log('Subscription panel display:', subscriptionPanel ? subscriptionPanel.style.display : 'null');
        
        if (chatPanel) {
            chatPanel.style.display = 'none';
            console.log('Chat panel hidden');
        }
        if (subscriptionPanel) {
            subscriptionPanel.style.display = 'none';
            console.log('Subscription panel hidden');
        }
        
        console.log('=== AFTER HIDING PANELS ===');
        console.log('Chat panel display:', chatPanel ? chatPanel.style.display : 'null');
        console.log('Subscription panel display:', subscriptionPanel ? subscriptionPanel.style.display : 'null');
        
        // Show the selected panel
        if (activePanel === subscriptionPanel) {
            subscriptionPanel.style.display = 'block'; // Use inline style since it's not a content-panel
            subscriptionPanel.style.position = 'absolute'; // Position absolutely to occupy same space as chat
            subscriptionPanel.style.top = '0'; // Start at top of parent
            subscriptionPanel.style.left = '0'; // Start at left of parent  
            subscriptionPanel.style.right = '0'; // Stretch to right edge
            subscriptionPanel.style.bottom = '0'; // Stretch to bottom edge
            subscriptionPanel.style.overflow = 'auto'; // Allow scrolling if needed
            subscriptionPanel.style.zIndex = '10'; // Ensure it's on top
            console.log('=== AFTER SHOWING SUBSCRIPTION PANEL ===');
            console.log('Subscription panel display:', subscriptionPanel.style.display);
            console.log('Subscription panel getBoundingClientRect():', subscriptionPanel.getBoundingClientRect());
        } else if (activePanel === chatPanel) {
            chatPanel.style.display = 'flex';
            chatPanel.style.position = 'relative'; // Ensure chat panel uses normal positioning
            console.log('=== AFTER SHOWING CHAT PANEL ===');
            console.log('Chat panel display:', chatPanel.style.display);
        }
        
        // Activate selected tab
        activeTab.classList.add('active');
    }

      // Chat tab click handler
      if (chatTab && chatPanel) {
        chatTab.addEventListener('click', () => {
          console.log('Switching to chat panel');
          switchToTab(chatTab, chatPanel);
        });
      }

      // Subscription tab click handler
      if (subscriptionTab && subscriptionPanel) {
        subscriptionTab.addEventListener('click', () => {
          console.log('Subscription tab clicked');
          console.log('subscriptionPanel element:', subscriptionPanel);
          console.log('chatPanel element:', chatPanel);
          switchToTab(subscriptionTab, subscriptionPanel);
          
          // Initialize pricing buttons when subscription panel is shown
          setTimeout(() => {
            console.log('Initializing pricing buttons after tab switch');
            initializePricingButtons();
          }, 100); // Small delay to ensure panel is fully visible
        });
      } else {
        console.error('Subscription tab or panel not found:', {
          subscriptionTab: !!subscriptionTab,
          subscriptionPanel: !!subscriptionPanel
        });
      }

      sidebarInitialized = true;
      console.log('Sidebar navigation initialized');
    }

    // >>> ADDED: Function to initialize pricing plan buttons
    function initializePricingButtons() {
        console.log('=== initializePricingButtons called ===');
        const subscriptionPanel = document.getElementById('subscription-panel');
        if (!subscriptionPanel) {
            console.error('Subscription panel not found in DOM');
            return;
        }
        
        const isVisible = subscriptionPanel.style.display !== 'none';
        const computedStyle = window.getComputedStyle(subscriptionPanel);
        const computedDisplay = computedStyle.display;
        
        console.log(`Subscription panel found:`);
        console.log(`- Style display: ${subscriptionPanel.style.display}`);
        console.log(`- Computed display: ${computedDisplay}`);
        console.log(`- Is visible: ${isVisible}`);
        console.log(`- Panel HTML (first 200 chars):`, subscriptionPanel.innerHTML.substring(0, 200));
        
        const planButtons = subscriptionPanel.querySelectorAll('.plan-button');
        console.log(`Found ${planButtons.length} plan buttons in subscription panel`);
        
        if (planButtons.length === 0) {
            console.warn('No plan buttons found - checking panel structure:');
            console.log('Plan cards:', subscriptionPanel.querySelectorAll('.plan-card').length);
            console.log('All buttons:', subscriptionPanel.querySelectorAll('button').length);
            return;
        }
        
        planButtons.forEach((button, index) => {
            console.log(`Processing button ${index}:`, button.textContent);
            
            // Check if button already has listener
            if (button.hasAttribute('data-listener-attached')) {
                console.log(`Button ${index} already has listener attached`);
                return;
            }
            
            console.log(`Adding listener to button ${index}`);
            
            // Add click event listener
            button.addEventListener('click', function(e) {
                e.preventDefault();
                console.log(`=== Button ${index} clicked! ===`);
                
                // Get the plan card parent to determine which plan was selected
                const planCard = button.closest('.plan-card');
                if (!planCard) {
                    console.error('Could not find plan card parent');
                    return;
                }
                
                const planNameElement = planCard.querySelector('.plan-name');
                if (!planNameElement) {
                    console.error('Could not find plan name element');
                    return;
                }
                
                const planName = planNameElement.textContent;
                console.log(`Pricing plan selected: ${planName}`);
                
                // Handle different plan selections
                handlePlanSelection(planName.toLowerCase(), planCard);
            });
            
            // Mark as having listener attached
            button.setAttribute('data-listener-attached', 'true');
            console.log(`Button ${index} listener attached successfully`);
        });
        
        console.log('=== Pricing plan buttons initialization complete ===');
    }

    // >>> ADDED: Function to handle plan selection
    function handlePlanSelection(planName, planCard) {
        console.log(`Handling selection for plan: ${planName}`);
        
        // Remove active class from all plan cards
        document.querySelectorAll('.plan-card').forEach(card => {
            card.classList.remove('selected');
        });
        
        // Add active class to selected plan
        planCard.classList.add('selected');
        
        // Handle specific plan logic
        switch(planName) {
            case 'starter':
                handleStarterPlan();
                break;
            case 'professional':
                handleProfessionalPlan();
                break;
            case 'unlimited':
                handleUnlimitedPlan();
                break;
            default:
                console.warn(`Unknown plan: ${planName}`);
        }
    }

    // >>> ADDED: Plan-specific handlers
    function handleStarterPlan() {
        console.log('handleStarterPlan called');
        const message = 'Starter plan selected! This would redirect to payment processing.';
        console.log(message);
        
        // Try to show message if function exists
        if (typeof showMessage === 'function') {
            showMessage(message);
        } else {
            alert(message); // Fallback
        }
    }

    function handleProfessionalPlan() {
        console.log('handleProfessionalPlan called');
        const message = 'Professional plan selected! This would redirect to payment processing.';
        console.log(message);
        
        // Try to show message if function exists
        if (typeof showMessage === 'function') {
            showMessage(message);
        } else {
            alert(message); // Fallback
        }
    }

    function handleUnlimitedPlan() {
        console.log('handleUnlimitedPlan called');
        
        // Handle the special case of the API key input in the unlimited plan
        const apiKeyInput = document.getElementById('claude-api-key');
        let message = 'Unlimited plan selected! This would redirect to payment processing.';
        
        if (apiKeyInput) {
            const apiKey = apiKeyInput.value.trim();
            if (apiKey) {
                console.log('Claude API key provided for unlimited plan');
                message = 'Unlimited plan selected with custom API key! This would redirect to payment processing.';
            }
        }
        
        console.log(message);
        
        // Try to show message if function exists
        if (typeof showMessage === 'function') {
            showMessage(message);
        } else {
            alert(message); // Fallback
        }
    }

    // >>> ADDED: Function to show startup menu
    function showStartupMenu() {
      // >>> PRODUCTION MODE CHECK: Don't show startup menu in production
      if (FORCE_PRODUCTION_MODE) {
        productionLog('Production mode - redirecting to client mode instead of startup menu');
        showClientMode();
      } else {
        productionLog('Development mode - showing startup menu');
        if (startupMenu) startupMenu.style.display = 'flex';
        if (appBody) appBody.style.display = 'none';
        if (clientModeView) clientModeView.style.display = 'none';
        if (authenticationView) authenticationView.style.display = 'none';
        
        // Reset client mode state when going back to menu
        resetChatClient();
      }
    }
    // <<< END ADDED

    function showAuthentication() {
      // Check if user is already authenticated
      if (isUserAuthenticated()) {
        const user = getCurrentUser();
        alert(`You are already signed in as ${user.name}. Redirecting to Client Mode.`);
        showClientMode();
        return;
      }
      
      if (startupMenu) startupMenu.style.display = 'none';
      if (appBody) appBody.style.display = 'none';
      if (clientModeView) clientModeView.style.display = 'none';
      if (authenticationView) authenticationView.style.display = 'flex';
      
      // Set up Google Sign-In button handler
      const googleSignInButton = document.getElementById('google-signin-button');
      if (googleSignInButton) {
        googleSignInButton.onclick = handleGoogleSignIn;
      }
      
      // Microsoft and API Key sign-in buttons removed
      
      console.log("Authentication view activated");
    }
    
    // Handle Microsoft Sign-In
    function handleMicrosoftSignIn() {
      console.log("Microsoft Sign-In clicked");
      
      // Check if user already signed in with MSAL
      if (typeof msalInstance !== 'undefined' && msalInstance) {
        handleSignInClick(); // Use existing MSAL sign-in
      } else {
        showError("Microsoft authentication is not available. Please try Google Sign-In or use an API Key.");
      }
    }
    
    // Show API Key Dialog
    function showApiKeyDialog() {
      console.log("API Key Sign-In clicked");
      
      // Create a simple dialog for API key input
      const dialog = document.createElement('div');
      dialog.style.cssText = `
        position: fixed;
        top: 50%;
        left: 50%;
        transform: translate(-50%, -50%);
        background: white;
        padding: 30px;
        border-radius: 8px;
        box-shadow: 0 4px 20px rgba(0,0,0,0.2);
        z-index: 10000;
        width: 90%;
        max-width: 400px;
      `;
      
      dialog.innerHTML = `
        <h3 style="margin-top: 0;">Enter Your OpenAI API Key</h3>
        <p style="color: #666; font-size: 14px;">Your API key will be stored locally and used for AI requests.</p>
        <input type="password" id="api-key-input" placeholder="sk-..." style="width: 100%; padding: 10px; margin: 10px 0; border: 1px solid #ddd; border-radius: 4px;">
        <div style="display: flex; gap: 10px; justify-content: flex-end; margin-top: 20px;">
          <button id="api-key-cancel" style="padding: 8px 16px; border: 1px solid #ddd; background: white; border-radius: 4px; cursor: pointer;">Cancel</button>
          <button id="api-key-submit" style="padding: 8px 16px; border: none; background: #333; color: white; border-radius: 4px; cursor: pointer;">Sign In</button>
        </div>
      `;
      
      // Add backdrop
      const backdrop = document.createElement('div');
      backdrop.style.cssText = `
        position: fixed;
        top: 0;
        left: 0;
        right: 0;
        bottom: 0;
        background: rgba(0,0,0,0.5);
        z-index: 9999;
      `;
      
      document.body.appendChild(backdrop);
      document.body.appendChild(dialog);
      
      // Focus input
      const input = document.getElementById('api-key-input');
      input.focus();
      
      // Handle cancel
      document.getElementById('api-key-cancel').onclick = () => {
        backdrop.remove();
        dialog.remove();
      };
      
      // Handle submit
      document.getElementById('api-key-submit').onclick = () => {
        const apiKey = input.value.trim();
        if (apiKey && apiKey.startsWith('sk-')) {
          // Store API key
          localStorage.setItem('user_api_key', apiKey);
          localStorage.setItem('auth_method', 'api_key');
          
          // Create mock user data
          const userData = {
            name: 'API Key User',
            email: 'api@user.local',
            picture: null,
            credits: 100,
            subscription: 'Free'
          };
          
          localStorage.setItem('user_data', JSON.stringify(userData));
          sessionStorage.setItem('googleUser', JSON.stringify(userData));
          
          // Update UI
          if (typeof window.initializeUserData === 'function') {
            window.initializeUserData();
          }
          
          // Clean up dialog
          backdrop.remove();
          dialog.remove();
          
          // Redirect based on stored preference
          const postAuthRedirect = localStorage.getItem('post_auth_redirect');
          if (postAuthRedirect === 'client-mode') {
            showClientMode();
          } else {
            showStartupMenu();
          }
          
          localStorage.removeItem('post_auth_redirect');
          showMessage('Successfully signed in with API key!');
        } else {
          showError('Please enter a valid OpenAI API key (should start with "sk-")');
        }
      };
      
      // Handle enter key
      input.addEventListener('keypress', (e) => {
        if (e.key === 'Enter') {
          document.getElementById('api-key-submit').click();
        }
      });
    }

    function handleGoogleSignIn() {
      console.log("Google Sign-In clicked");
      
      // Ensure Office.js is fully initialized before proceeding
      if (typeof Office === 'undefined') {
        console.error("Office.js is not loaded");
        showError("Office Add-in environment not ready. Please refresh and try again.");
        return;
      }
      
      // Wait for Office to be ready if it's still initializing
      Office.onReady(() => {
        console.log("Office.onReady called - proceeding with authentication");
        proceedWithGoogleAuth();
      });
    }
    
    function proceedWithGoogleAuth() {
      // Debug Office context with more detail
      console.log("=== Office Environment Debug ===");
      console.log("Office available:", typeof Office !== 'undefined');
      console.log("Office.context available:", typeof Office !== 'undefined' && Office.context);
      console.log("Office.context.ui available:", typeof Office !== 'undefined' && Office.context && Office.context.ui);
      console.log("Office.context.ui.displayDialogAsync available:", typeof Office !== 'undefined' && Office.context && Office.context.ui && typeof Office.context.ui.displayDialogAsync === 'function');
      console.log("Office host:", Office.context?.host);
      console.log("Office platform:", Office.context?.platform);
      console.log("================================");
      
      // Use Office Dialog API for authentication within Excel environment
      if (typeof Office !== 'undefined' && Office.context && Office.context.ui && typeof Office.context.ui.displayDialogAsync === 'function') {
         console.log("✅ Using Office Dialog API for authentication");
        const authUrl = buildGoogleAuthUrl();
        console.log("Auth URL:", authUrl);
         
         // Show loading message
         showMessage("Opening authentication dialog...");
        
        Office.context.ui.displayDialogAsync(authUrl, {
           height: 70,
           width: 70,
           requireHTTPS: true,
           displayInIframe: false // Ensure it opens in a proper dialog, not iframe
        }, function (result) {
           console.log("Dialog creation result:", result);
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            const dialog = result.value;
            
            // Listen for messages from the dialog
            dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
              try {
                const authResult = JSON.parse(arg.message);
                dialog.close();
                
                if (authResult.success) {
                  handleGoogleAuthSuccess(authResult);
                } else {
                  handleGoogleAuthError(authResult.error);
                }
              } catch (error) {
                console.error("Error parsing auth result:", error);
                dialog.close();
                showError("Authentication failed. Please try again.");
              }
            });
            
            // Handle dialog closed by user
            dialog.addEventHandler(Office.EventType.DialogEventReceived, function (arg) {
              if (arg.error === 12006) { // Dialog closed by user
                console.log("Authentication canceled by user");
              } else {
                console.error("Dialog error:", arg.error);
                showError("Authentication failed. Please try again.");
              }
            });
          } else {
            console.error("Failed to open authentication dialog:", result.error);
            console.error("Result details:", result);
            // Fallback to external window for development
            handleGoogleSignInFallback();
          }
        });
      } else {
        console.log("Office Dialog API not available, using fallback");
        handleGoogleSignInFallback();
      }
    }

    function buildGoogleAuthUrl() {
      // Use the current origin for better flexibility across environments
      const params = new URLSearchParams({
        client_id: GOOGLE_CONFIG.CLIENT_ID,
        redirect_uri: `${window.location.origin}/auth/google/callback.html`,
        response_type: 'token id_token',
        scope: GOOGLE_CONFIG.SCOPES,
        nonce: generateRandomState(),
        state: generateRandomState(),
        prompt: 'select_account' // Force account selection every time
      });
      
      return `https://accounts.google.com/o/oauth2/v2/auth?${params.toString()}`;
    }

    function generateRandomState() {
      return Math.random().toString(36).substring(2, 15) + Math.random().toString(36).substring(2, 15);
    }

    // Google OAuth Configuration
    const GOOGLE_CONFIG = {
      CLIENT_ID: "194102993575-6vkqrurasev2qr90nv82cbampqqq22hb.apps.googleusercontent.com",
      SCOPES: "email profile"
    };

    function handleGoogleSignInFallback() {
      // Fallback for development or when Office Dialog API is not available
      const authUrl = buildGoogleAuthUrl();
      const popup = window.open(authUrl, 'googleAuth', 'width=500,height=600,scrollbars=yes,resizable=yes');
      
      // Poll for popup closure (not ideal but works for development)
      const pollTimer = setInterval(() => {
        try {
          if (popup.closed) {
            clearInterval(pollTimer);
            console.log("Authentication popup closed");
            // You might want to check for stored tokens here
          }
        } catch (error) {
          // Cross-origin error when popup is open
        }
      }, 1000);
      
      // Clean up after 5 minutes
      setTimeout(() => {
        clearInterval(pollTimer);
        if (!popup.closed) {
          popup.close();
        }
      }, 300000);
    }

         async function handleGoogleAuthSuccess(authResult) {
       console.log("✅ Google authentication successful", authResult);
      
       try {
         // Store Google user data locally (for display purposes)
      // Store in both sessionStorage (current session) and localStorage (persistence)
      if (authResult.user) {
        sessionStorage.setItem('googleUser', JSON.stringify(authResult.user));
        localStorage.setItem('googleUser', JSON.stringify(authResult.user));
      }
      if (authResult.access_token) {
        sessionStorage.setItem('googleToken', authResult.access_token);
        localStorage.setItem('google_access_token', authResult.access_token);
      }
      if (authResult.id_token) {
        sessionStorage.setItem('googleCredential', authResult.id_token);
        localStorage.setItem('googleCredential', authResult.id_token);
      }
      
      const userName = authResult.user ? authResult.user.name : 'User';
         showMessage(`Welcome ${userName}! Connecting to backend...`);
         
         // Authenticate with backend using Google ID token
         if (authResult.id_token) {
           console.log("🔐 Authenticating with backend...");
           
           const backendAuthResult = await backendAPI.signInWithGoogle(authResult.id_token);
           console.log("✅ Backend authentication successful");
           
           // Initialize user data from backend
           await initializeUserData();
           
           showMessage(`Welcome ${userName}! You're all set!`);
         } else {
           console.warn("⚠️ No ID token received, skipping backend authentication");
           showMessage(`Welcome ${userName}! (Google only)`);
         }
         
         // Check for post-auth redirect
         const postAuthRedirect = localStorage.getItem('post_auth_redirect');
         if (postAuthRedirect === 'client-mode') {
           console.log('Post-auth redirect to client mode detected');
           localStorage.removeItem('post_auth_redirect');
           // Simply call showAuthenticatedInterface which handles the client mode display
           showAuthenticatedInterface();
         } else {
           // Update interface to show authenticated state
           showAuthenticatedInterface();
         }
         
       } catch (error) {
         console.error("❌ Backend authentication failed:", error);
         
         // Still show interface if Google auth worked, but warn about backend
         const userName = authResult.user ? authResult.user.name : 'User';
         showMessage(`Welcome ${userName}! (Limited features - backend unavailable)`);
         showAuthenticatedInterface();
         
         // Show error notification
         setTimeout(() => {
           showError("Some features may be unavailable due to backend connection issues.");
         }, 2000);
       }
    }

    function handleGoogleAuthError(error) {
      console.error("Google authentication error:", error);
      showError("Authentication failed: " + (error.message || error));
    }

    function handleGoogleCallback(response) {
      console.log("Google authentication successful", response);
      
      // Decode the JWT token to get user info
      const userInfo = parseJwt(response.credential);
      console.log("User info:", userInfo);
      
      // Store user session
      sessionStorage.setItem('googleUser', JSON.stringify(userInfo));
      sessionStorage.setItem('googleCredential', response.credential);
      
      // Show success message and update interface
      showMessage(`Welcome ${userInfo.name}! Authentication successful.`);
      
      // Update interface to show authenticated state
      showAuthenticatedInterface();
    }

    function handleGoogleTokenResponse(tokenResponse) {
      console.log("Google token response:", tokenResponse);
      
      if (tokenResponse.access_token) {
        // Fetch user info with the access token
        fetch('https://www.googleapis.com/oauth2/v2/userinfo', {
          headers: {
            'Authorization': `Bearer ${tokenResponse.access_token}`
          }
        })
        .then(response => response.json())
        .then(userInfo => {
          console.log("User info from token:", userInfo);
          
          // Store user session
          sessionStorage.setItem('googleUser', JSON.stringify(userInfo));
          sessionStorage.setItem('googleToken', tokenResponse.access_token);
          
          // Show success message and update interface
          showMessage(`Welcome ${userInfo.name}! Authentication successful.`);
          
          // Update interface to show authenticated state
          showAuthenticatedInterface();
        })
        .catch(error => {
          console.error("Error fetching user info:", error);
          showMessage("Authentication successful but failed to get user information.");
          
          // Still update interface since authentication was successful
          showAuthenticatedInterface();
        });
      }
    }

    function parseJwt(token) {
      try {
        const base64Url = token.split('.')[1];
        const base64 = base64Url.replace(/-/g, '+').replace(/_/g, '/');
        const jsonPayload = decodeURIComponent(atob(base64).split('').map(function(c) {
          return '%' + ('00' + c.charCodeAt(0).toString(16)).slice(-2);
        }).join(''));
        return JSON.parse(jsonPayload);
      } catch (error) {
        console.error("Error parsing JWT:", error);
        return null;
      }
    }

    // Function to check if user is already authenticated
    function isUserAuthenticated() {
      const user = sessionStorage.getItem('googleUser');
      const credential = sessionStorage.getItem('googleCredential') || sessionStorage.getItem('googleToken');
      return user && credential;
    }

    // Function to get current user info
    function getCurrentUser() {
      const userStr = sessionStorage.getItem('googleUser');
      return userStr ? JSON.parse(userStr) : null;
    }

    // Function to get current user display name
    function getCurrentUserDisplayName() {
      const googleUser = getCurrentUser();
      console.log('getCurrentUserDisplayName - googleUser:', googleUser);
      console.log('getCurrentUserDisplayName - currentUser (Microsoft):', currentUser);
      
      if (googleUser) {
        const displayName = googleUser.name || googleUser.email || 'Google User';
        console.log('Using Google user display name:', displayName);
        return displayName;
      }
      
      // Check for Microsoft authentication
      if (currentUser) {
        const displayName = currentUser.name || currentUser.username || 'Microsoft User';
        console.log('Using Microsoft user display name:', displayName);
        return displayName;
      }
      
      console.log('No authenticated user found');
      return null;
    }

    // Function to update signed-in status display
    function updateSignedInStatus() {
      const signedInStatus = document.getElementById('signed-in-status');
      const signedInUser = document.getElementById('signed-in-user');
      
      console.log('updateSignedInStatus called');
      console.log('signedInStatus element:', signedInStatus);
      console.log('signedInUser element:', signedInUser);
      
      if (signedInStatus && signedInUser) {
        const userDisplayName = getCurrentUserDisplayName();
        console.log('userDisplayName:', userDisplayName);
        if (userDisplayName) {
          signedInUser.textContent = userDisplayName;
          signedInStatus.style.display = 'block';
          console.log('Signed-in status shown for:', userDisplayName);
          
          // Update credits display in client mode
          const clientCreditsElement = document.getElementById('credits-count-client');
          if (clientCreditsElement && window.userProfileManager) {
            const credits = window.userProfileManager.getCredits();
            clientCreditsElement.textContent = credits.toFixed(1);
            console.log('🔧 Updated client mode credits display:', credits);
          }
          
          // Check subscription status
          checkSubscriptionStatus();
        } else {
          signedInStatus.style.display = 'none';
          console.log('Signed-in status hidden - no user found');
        }
      } else {
        console.log('Required elements not found for signed-in status');
      }
    }

    // Function to check subscription status
    async function checkSubscriptionStatus() {
      const subscriptionStatusElement = document.getElementById('subscription-status');
      if (!subscriptionStatusElement) return;

      const googleUser = getCurrentUser();
      if (!googleUser || !googleUser.email) {
        console.log('No user email found for subscription check');
        return;
      }

      try {
        console.log('Checking subscription status for:', googleUser.email);
        
        // Use BackendAPI which has mock support built-in
        const subscriptionData = await backendAPI.getSubscriptionStatus();
          updateSubscriptionDisplay(subscriptionData);
        
      } catch (error) {
        console.error('Error checking subscription status:', error);
        // Use fallback data for development
        updateSubscriptionDisplay({ 
          status: 'none', 
          plan: 'Free',
          hasActiveSubscription: false,
          credits: 15
        });
      }
    }

    // Function to update subscription status display
    function updateSubscriptionDisplay(subscriptionData) {
      const subscriptionStatusElement = document.getElementById('subscription-status');
      if (!subscriptionStatusElement) return;

      // Clear previous classes
      subscriptionStatusElement.className = 'subscription-status-text';
      
      if (subscriptionData.status === 'active' && subscriptionData.plan) {
        subscriptionStatusElement.textContent = `${subscriptionData.plan} Plan`;
        subscriptionStatusElement.classList.add('active');
        subscriptionStatusElement.style.display = 'block';
      } else if (subscriptionData.status === 'inactive' || subscriptionData.status === 'cancelled') {
        subscriptionStatusElement.textContent = 'No Active Subscription';
        subscriptionStatusElement.classList.add('inactive');
        subscriptionStatusElement.style.display = 'block';
      } else if (subscriptionData.status === 'error') {
        subscriptionStatusElement.textContent = 'Subscription status unavailable';
        subscriptionStatusElement.classList.add('error');
        subscriptionStatusElement.style.display = 'block';
      } else {
        // Unknown status, hidden, or no subscription info
        subscriptionStatusElement.style.display = 'none';
      }
      
      console.log('Subscription status updated:', subscriptionData);
    }

    // Make functions globally accessible for cross-file access
    window.updateSignedInStatus = updateSignedInStatus;
    window.getCurrentUserDisplayName = getCurrentUserDisplayName;

    // Function to sign out user
    function signOutUser() {
      sessionStorage.removeItem('googleUser');
      sessionStorage.removeItem('googleCredential');
      sessionStorage.removeItem('googleToken');
      
      // Revoke Google token if available
      if (typeof google !== 'undefined' && google.accounts) {
        google.accounts.id.disableAutoSelect();
      }
      
      console.log("User signed out");
      showStartupMenu();
    }

    // Function to update UI based on authentication state
    function updateAuthUI() {
      const isAuthenticated = isUserAuthenticated();
      const user = getCurrentUser();
      
      if (isAuthenticated && user) {
        console.log(`User is authenticated: ${user.name} (${user.email})`);
        // Could update UI to show user info, add sign-out button, etc.
      } else {
        console.log("User is not authenticated");
      }
    }

    // Assign click handlers for startup menu buttons
    if (developerModeButton) {
        developerModeButton.onclick = showDeveloperMode;
    } else {
        console.error("Could not find button with id='developer-mode-button'");
    }
    if (clientModeButton) {
        clientModeButton.onclick = showClientMode;
    } else {
        console.error("Could not find button with id='client-mode-button'");
    }
    // Authentication button removed
    // if (authenticationButton) {
    //     authenticationButton.onclick = showAuthentication;
    // } else {
    //     console.error("Could not find button with id='authentication-button'");
    // }



    // Assign the REVISED async function as the handler
    const button = document.getElementById("insert-and-run");
    if (button) {
        button.onclick = async () => {
          // Check credits before allowing model building
          await enforceFeatureAccess('build', async () => {
            // Use credit for building
            await useCreditForBuild();
            // Proceed with model building
            await insertSheetsAndRunCodes();
          });
        };
    } else {
        console.error("Could not find button with id='insert-and-run'");
    }

    // ... (rest of your Office.onReady remains the same) ...

    // Keep the setup for your other buttons (send-button, reset-button, etc.)
    const sendButton = document.getElementById('send');
    if (sendButton) {
      sendButton.onclick = async () => {
        // Check credits before allowing AI conversation
        await enforceFeatureAccess('update', async () => {
          // Use credit for update/conversation
          await useCreditForUpdate();
          // Proceed with AI conversation
          await handleSend();
        });
      };
    }

    const writeButton = document.getElementById('write-to-excel');
    if (writeButton) writeButton.onclick = writeToExcel;

    const resetButton = document.getElementById('reset-chat');
    if (resetButton) resetButton.onclick = resetChat;

    // Add to Training Data Queue button
    const addToTrainingQueueButton = document.getElementById('add-to-training-queue');
    if (addToTrainingQueueButton) {
        addToTrainingQueueButton.onclick = addToTrainingDataQueue;
    } else {
        console.error("Could not find button with id='add-to-training-queue'");
    }

    // Training Data Modal buttons
    const trainingDataModal = document.getElementById('training-data-modal');
    const saveTrainingDataButton = document.getElementById('save-training-data-button');
    const cancelTrainingDataButton = document.getElementById('cancel-training-data-button');
    
    if (saveTrainingDataButton) {
        saveTrainingDataButton.onclick = saveTrainingDataFromModal;
    } else {
        console.error("Could not find button with id='save-training-data-button'");
    }
    
    if (cancelTrainingDataButton) {
        cancelTrainingDataButton.onclick = hideTrainingDataModal;
    } else {
        console.error("Could not find button with id='cancel-training-data-button'");
    }
    
    // Training Data Modal close button
    if (trainingDataModal) {
        const trainingDataCloseButton = trainingDataModal.querySelector('.close-button');
        if (trainingDataCloseButton) {
            trainingDataCloseButton.onclick = hideTrainingDataModal;
        }
    }

    // Training Data Modal find/replace buttons
    const trainingReplaceAllButton = document.getElementById('training-replace-all-button');
    if (trainingReplaceAllButton) {
        trainingReplaceAllButton.onclick = trainingModalReplaceAll;
    } else {
        console.error("Could not find button with id='training-replace-all-button'");
    }

    // >>> ADDED: Auto-resize textarea functionality - DISABLED for single-line input
    // const userInputClient = document.getElementById('user-input-client');
    // if (userInputClient) {
    //     // Function to auto-resize textarea
    //     function autoResizeTextarea() {
    //         // Reset height to auto to get the correct scrollHeight
    //         userInputClient.style.height = 'auto';
    //         
    //         // Calculate new height
    //         const newHeight = Math.min(userInputClient.scrollHeight, 72); // Max 3 lines (~72px)
    //         userInputClient.style.height = newHeight + 'px';
    //     }
    //     
    //     // Add event listeners for auto-resize
    //     userInputClient.addEventListener('input', autoResizeTextarea);
    //     userInputClient.addEventListener('change', autoResizeTextarea);
    //     
    //     // Initial call to set correct height
    //     autoResizeTextarea();
    // }

    // >>> ADDED: Setup for Client Mode Chat Buttons
    // Don't override the send button handler - it's already set by AIcalls.js to use plannerHandleSend
    // which properly includes attachments. Just ensure credit enforcement wraps around it.
    const sendClientButton = document.getElementById('send-client');
    if (sendClientButton) {
      const originalHandler = sendClientButton.onclick;
      if (originalHandler) {
        sendClientButton.onclick = async () => {
          // Check credits before allowing AI conversation
          await enforceFeatureAccess('update', async () => {
            // Use credit for update/conversation
            await useCreditForUpdate();
            // Call the original handler (plannerHandleSend)
            await originalHandler();
          });
        };
      }
    }
    
    // >>> ADDED: Voice Recording Setup for Client Mode
    initializeVoiceRecordingClient();

    // Add click handler for EBITDAI header to open website
    const headerTitleLink = document.getElementById('header-title-link');
    if (headerTitleLink) {
        headerTitleLink.onclick = () => {
            window.open('https://ebitdai.co', '_blank');
        };
    }

    const resetChatClientButton = document.getElementById('reset-chat-client');
    if (resetChatClientButton) resetChatClientButton.onclick = resetChatClient;

    // Setup subscription upgrade buttons
    const upgradeButtons = document.querySelectorAll('.upgrade-button');
    upgradeButtons.forEach(button => {
      button.onclick = async () => {
        console.log("💰 Upgrade button clicked");
        await startSubscription();
      };
    });

    // Add event listener for the new icon button
    const resetChatIconButton = document.getElementById('reset-chat-icon-button');
    if (resetChatIconButton) {
        resetChatIconButton.onclick = resetChatClient;
    } else {
        console.error("Could not find button with id='reset-chat-icon-button'");
    }

    const writeToExcelClientButton = document.getElementById('write-to-excel-client');
    if (writeToExcelClientButton) {
        // writeToExcelClientButton.onclick = () => alert('Client Mode "Write to Excel" is not yet implemented.');
    }
    const insertToEditorClientButton = document.getElementById('insert-to-editor-client');
    if (insertToEditorClientButton) {
        // insertToEditorClientButton.onclick = () => alert('Client Mode "Insert to Editor" is not yet implemented.');
    }
    // <<< END ADDED

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

    // Update the welcome title and steps based on the selected mode
    function updateWelcomeByMode(modeValue) {
        try {
            const titleEl = document.querySelector('.welcome-section .welcome-title');
            const stepTextEls = document.querySelectorAll('.welcome-section .welcome-steps .step-text');

            if (!titleEl || stepTextEls.length < 2) {
                return; // Welcome screen not visible in current context
            }

            if (modeValue === 'one-shot') {
                titleEl.textContent = 'One-Shot Mode';
                stepTextEls[0].textContent = "Provide your key business assumptions, including revenue, pricing, expenses, financing, and other critical details.";
                stepTextEls[1].textContent = "Your model will be generated instantly based on the information you provide.";
            } else if (modeValue === 'limited-guidance') {
                titleEl.textContent = 'Limited-Guidance Mode';
                stepTextEls[0].textContent = "Share your business assumptions covering revenue, pricing, expenses, financing, and other essential inputs.";
                stepTextEls[1].textContent = "We’ll ask one to two clarifying questions before delivering your completed model.";
            } else if (modeValue === 'full-guidance') {
                titleEl.textContent = 'Full-Guidance Mode';
                stepTextEls[0].textContent = 'Start with a concise overview of your business.';
                stepTextEls[1].textContent = 'We’ll work with you step-by-step to define every aspect of your model.';
            }
        } catch (err) {
            console.warn('updateWelcomeByMode failed:', err);
        }
    }

    // Initialize system prompt dropdown functionality
    function initializeSystemPromptDropdown() {
        // Handle both the old dropdown and the new header dropdown
        const oldDropdown = document.getElementById('system-prompt-dropdown');
        const newDropdown = document.getElementById('mode-dropdown');
        
        const handleDropdownChange = function() {
                console.log(`System prompt mode changed to: ${this.value}`);
            // Sync both dropdowns if they exist
            if (oldDropdown && this === newDropdown) oldDropdown.value = this.value;
            if (newDropdown && this === oldDropdown) newDropdown.value = this.value;
                // Optionally reset conversation when mode changes
                // resetChatClient();

            // Update welcome screen content according to selected mode
            const currentMode = this.value;
            updateWelcomeByMode(currentMode);
        };
        
        if (oldDropdown) {
            oldDropdown.addEventListener('change', handleDropdownChange);
        }
        if (newDropdown) {
            newDropdown.addEventListener('change', handleDropdownChange);
            // Initialize the welcome screen once with current value
            updateWelcomeByMode(newDropdown.value);
        }
    }

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
        
        // >>> ADDED: Set the API key for AI Model Planner voice input
        if (keys.OPENAI_API_KEY) {
          setAIModelPlannerOpenApiKey(keys.OPENAI_API_KEY);
        }
      }
      
      // >>> ADDED: Initialize system prompt dropdown
      initializeSystemPromptDropdown();
      
      // >>> ADDED: Clear conversation history on startup to ensure fresh start
      console.log("Clearing conversation history on startup...");
      conversationHistory = [];
      saveConversationHistory(conversationHistory);
      lastResponse = null;
      isFirstMessageInSession = true;
      firstUserInput = null;
      lastUserInput = null;
      persistedTrainingUserInput = null;
      persistedTrainingAiResponse = null;
      console.log("Conversation history cleared on startup");
      // <<< END ADDED

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

      // Startup Menu Logic - Placed before Promise.all to ensure elements are handled
      const startupMenu = document.getElementById('startup-menu');
      const developerModeButton = document.getElementById('developer-mode-button');
      const clientModeButton = document.getElementById('client-mode-button');
      // appBody and clientModeView will be fetched inside showDeveloperMode/showClientMode
      // or assume they are accessible if defined earlier in Office.onReady

      if (developerModeButton) {
          developerModeButton.onclick = showDeveloperMode; // Assumes showDeveloperMode is globally accessible
      } else {
          console.error("[Office.onReady] Could not find button with id='developer-mode-button'");
      }
      if (clientModeButton) {
          clientModeButton.onclick = showClientMode; // Assumes showClientMode is globally accessible
      } else {
          console.error("[Office.onReady] Could not find button with id='client-mode-button'");
      }

      // Get references and assign handlers for Back to Menu buttons
      const backToMenuDevButton = document.getElementById('back-to-menu-dev-button');
      const backToMenuClientButton = document.getElementById('back-to-menu-client-button');
      const backToMenuAuthButton = document.getElementById('back-to-menu-auth-button');

      if (backToMenuDevButton) {
          backToMenuDevButton.onclick = showStartupMenu; // Assumes showStartupMenu is globally accessible
      } else {
          console.error("[Office.onReady] Could not find button with id='back-to-menu-dev-button'");
      }
      if (backToMenuClientButton) {
          backToMenuClientButton.onclick = showStartupMenu; // Assumes showStartupMenu is globally accessible
      } else {
          console.error("[Office.onReady] Could not find button with id='back-to-menu-client-button'");
      }
      if (backToMenuAuthButton) {
          backToMenuAuthButton.onclick = showStartupMenu; // Assumes showStartupMenu is globally accessible
      } else {
          console.error("[Office.onReady] Could not find button with id='back-to-menu-auth-button'");
      }

      document.getElementById("sideload-msg").style.display = "none";
      const appBody = document.getElementById('app-body');
      const clientModeView = document.getElementById('client-mode-view');

      // >>> DYNAMIC MODE CHECK: Use same detection as main logic
      const isLocalDev = window.location.hostname === 'localhost' || 
                        window.location.hostname === '127.0.0.1' ||
                        window.location.href.includes('localhost:3002');
      const forceProductionMode = !isLocalDev;
      
      if (forceProductionMode) {
        productionLog('Initialization complete - maintaining client mode (production)');
        // Keep client mode active, don't show startup menu
        if (startupMenu) startupMenu.style.display = "none";
        if (appBody) appBody.style.display = "none";
        if (clientModeView) clientModeView.style.display = "flex";
      } else {
        productionLog('Initialization complete - showing startup menu (development)');
        if (startupMenu) startupMenu.style.display = "flex";
        if (appBody) appBody.style.display = "none";
        if (clientModeView) clientModeView.style.display = "none";
      }
      // End Startup Menu Logic

    }).catch(error => {
        console.error("Error during initialization:", error);
        showError("Error during initialization: " + error.message);
    });

    document.getElementById("sideload-msg").style.display = "none";
    // document.getElementById("app-body").style.display = "block"; // Keep app-body hidden initially
    
    // >>> DYNAMIC MODE CHECK: Use same detection as main logic
    const isLocalDev2 = window.location.hostname === 'localhost' || 
                       window.location.hostname === '127.0.0.1' ||
                       window.location.href.includes('localhost:3002');
    const forceProductionMode2 = !isLocalDev2;
    
    if (forceProductionMode2) {
      productionLog('Post-initialization - maintaining client mode (production)');
      if (startupMenu) startupMenu.style.display = "none"; // Keep startup menu hidden
      if (appBody) appBody.style.display = "none";
      if (clientModeView) clientModeView.style.display = "flex"; // Keep client mode active
    } else {
      productionLog('Post-initialization - showing startup menu (development)');
      if (startupMenu) startupMenu.style.display = "flex"; // Show startup menu instead
      if (appBody) appBody.style.display = "none";
      if (clientModeView) clientModeView.style.display = "none";
    }

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
                    // UPDATED regex to handle both old format (V|) and new format with labels (V(D)|)
                    const driverRegex = /(row\d+\s*=\s*")([A-Z]+)(\d*)(\([^)]+\))?(\|)/g;
                    const nextNumbers = { ...maxNumbers };

                    codeToAdd = codeToAdd.replace(driverRegex, (match, rowPart, prefix, existingNumberStr, label, pipePart) => {
                        // Create composite key for tracking numbers (same as in getMaxDriverNumbers)
                        const key = label ? `${prefix}${label}` : prefix;
                        nextNumbers[key] = (nextNumbers[key] || 0) + 1;
                        const newNumber = nextNumbers[key];
                        
                        // Build replacement: prefix + number + label (if exists) + pipe
                        const replacement = `${rowPart}${prefix}${newNumber}${label || ''}${pipePart}`;
                        console.log(`Replacing driver: '${prefix}${existingNumberStr || ''}${label || ''}|' with '${prefix}${newNumber}${label || ''}|'`);
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

                    const newCursorPos = (textBeforeFinal + codeToAdd).length;
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

    // >>> ADDED: Setup for the Clear Training Data Queue button
    const clearTrainingDataQueButtonHTML = document.getElementById('clear-training-data-que-button');
    if (clearTrainingDataQueButtonHTML) {
        clearTrainingDataQueButtonHTML.onclick = clearTrainingDataQueue; // Use the local function
    } else {
        console.error("Could not find button with id='clear-training-data-que-button' for developer mode.");
    }
    // <<< END ADDED

    // >>> ADDED: Setup for the Attach Actuals button
    const attachActualsButton = document.getElementById('attach-actuals-button');
    if (attachActualsButton) {
        attachActualsButton.onclick = handleAttachActuals;
    } else {
        console.error("Could not find button with id='attach-actuals-button' for developer mode.");
    }
    // <<< END ADDED

    // >>> ADDED: Microsoft Authentication Setup
    // Initialize authentication system
    try {
        console.log('Setting up Microsoft authentication...');
        
        if (typeof msal !== 'undefined') {
            const initSuccess = initializeMSAL();
            if (initSuccess) {
                console.log('Authentication system initialized successfully');
            } else {
                console.warn('Authentication system failed to initialize');
            }
        } else {
            console.warn('MSAL library not available - authentication features disabled');
            showError('Microsoft authentication library not loaded. Authentication features are disabled.');
        }
    } catch (error) {
        console.error('Error initializing authentication:', error);
        showError('Failed to set up authentication: ' + error.message);
    }

    // Setup authentication event handlers
    const signInButton = document.getElementById('sign-in-button');
    if (signInButton) {
        signInButton.onclick = async () => {
            try {
                await signIn();
            } catch (error) {
                console.error('Sign-in error:', error);
                showError('Sign-in failed. Please try again.');
            }
        };
        
        // Always show the sign-in button initially (it will be managed by auth state)
        signInButton.style.display = 'flex';
        console.log('Sign-in button event handler attached and button shown');
    } else {
        console.log('Sign-in button not found (may be in different mode)');
    }

    const signOutButton = document.getElementById('sign-out-button');
    if (signOutButton) {
        signOutButton.onclick = async () => {
            try {
                await signOut();
            } catch (error) {
                console.error('Sign-out error:', error);
                showError('Sign-out failed. Please try again.');
            }
        };
        console.log('Sign-out button event handler attached');
    } else {
        console.log('Sign-out button not found (may be in different mode)');
    }

    // Add event listener for the mini sign-out button in client mode
    const signOutMiniButton = document.getElementById('sign-out-mini-button');
    if (signOutMiniButton) {
        signOutMiniButton.onclick = async () => {
            try {
                await signOut();
            } catch (error) {
                console.error('Sign-out error:', error);
                showError('Sign-out failed. Please try again.');
            }
        };
        console.log('Sign-out mini button event handler attached');
    } else {
        console.log('Sign-out mini button not found');
    }
    // <<< END ADDED

    // >>> ADDED: Fullscreen toggle functionality for code editor
    const fullscreenToggle = document.getElementById('fullscreen-toggle');
    if (fullscreenToggle) {
        let isFullscreen = false;
        
        fullscreenToggle.onclick = () => {
            const appBody = document.getElementById('app-body');
            if (!appBody) return;
            
            if (!isFullscreen) {
                // Enter fullscreen mode
                appBody.classList.add('code-fullscreen');
                fullscreenToggle.title = 'Exit fullscreen';
                
                // Update icon to exit fullscreen icon
                fullscreenToggle.innerHTML = `
                    <svg viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
                        <path d="M8 3v3a2 2 0 0 1-2 2H3m18 0h-3a2 2 0 0 1-2-2V3m0 18v-3a2 2 0 0 1 2-2h3M3 16h3a2 2 0 0 1 2 2v3" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" fill="none"/>
                    </svg>
                `;
                isFullscreen = true;
                console.log('Code editor entered fullscreen mode');
            } else {
                // Exit fullscreen mode
                appBody.classList.remove('code-fullscreen');
                fullscreenToggle.title = 'Toggle fullscreen';
                
                // Update icon back to expand icon
                fullscreenToggle.innerHTML = `
                    <svg viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
                        <path d="M8 3H5a2 2 0 0 0-2 2v3m18 0V5a2 2 0 0 0-2-2h-3m0 18h3a2 2 0 0 0 2-2v-3M3 16v3a2 2 0 0 0 2 2h3" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" fill="none"/>
                    </svg>
                `;
                isFullscreen = false;
                console.log('Code editor exited fullscreen mode');
            }
        };
        
        // ESC key to exit fullscreen
        document.addEventListener('keydown', (event) => {
            if (event.key === 'Escape' && isFullscreen) {
                fullscreenToggle.click();
            }
        });
        
        console.log('Fullscreen toggle button initialized');
    } else {
        console.error("Could not find fullscreen toggle button with id='fullscreen-toggle'");
    }
    // <<< END ADDED

    // Window click handler to close modals when clicking outside
    window.onclick = function(event) {
        const trainingDataModal = document.getElementById('training-data-modal');
        const codeParamsModal = document.getElementById('code-params-modal');
        
        // Close training data modal if clicking outside
        if (trainingDataModal && event.target === trainingDataModal) {
            hideTrainingDataModal();
        }
        
        // Close code params modal if clicking outside (existing functionality)
        if (codeParamsModal && event.target === codeParamsModal) {
            hideParamsModal();
        }
    };

    // Initialize new header menu functionality
    function initializeHeaderMenu() {
        const hamburgerButton = document.getElementById('hamburger-menu');
        const slideMenu = document.getElementById('slide-menu');
        const closeMenuButton = document.getElementById('close-menu');
        const profileModal = document.getElementById('profile-modal');
        const modalCloseButton = profileModal?.querySelector('.modal-close-button');
        const accountModal = document.getElementById('account-modal');
        const accountModalCloseButton = accountModal?.querySelector('.modal-close-button');
        
        // Hamburger menu toggle
        if (hamburgerButton && slideMenu) {
            hamburgerButton.addEventListener('click', () => {
                slideMenu.classList.add('open');
            });
            
            closeMenuButton?.addEventListener('click', () => {
                slideMenu.classList.remove('open');
            });
        }
        
        // Menu item actions (robust: delegate on nav to capture clicks on inner SVG/spans)
        const slideMenuNav = document.querySelector('.slide-menu-nav');
        if (slideMenuNav) {
            slideMenuNav.addEventListener('click', (event) => {
                const item = event.target.closest('button.menu-item');
                if (!item) return;
                const action = item.dataset.action;
                
                switch (action) {
                    case 'chat':
                        // Show chat panel, hide subscription panel
                        document.getElementById('client-chat-container').style.display = 'flex';
                        document.getElementById('subscription-panel').style.display = 'none';
                        slideMenu.classList.remove('open');
                        break;
                        
                    case 'subscription':
                        // Show subscription panel, hide chat
                        document.getElementById('client-chat-container').style.display = 'none';
                        document.getElementById('subscription-panel').style.display = 'block';
                        slideMenu.classList.remove('open');
                        break;
                        
                    case 'plans':
                        // Open pricing/credits page in default browser
                        try {
                            window.open('https://ebitdai.co/credits', '_blank');
                        } catch (e) {
                            console.warn('Failed to open plans link, falling back');
                            location.href = 'https://ebitdai.co/credits';
                        }
                        slideMenu.classList.remove('open');
                        break;

                    case 'developer':
                        // Switch to developer mode
                        if (typeof showDeveloperMode === 'function') {
                            showDeveloperMode();
                            slideMenu.classList.remove('open');
                        }
                        break;
                        
                    case 'account':
                        // Show account modal
                        const accountModal = document.getElementById('account-modal');
                        if (accountModal) {
                            accountModal.style.display = 'flex';
                            slideMenu.classList.remove('open');
                            updateAccountModal();
                        }
                        break;
                        
                    case 'profile':
                        // Show profile modal
                        if (profileModal) {
                            profileModal.style.display = 'flex';
                            slideMenu.classList.remove('open');
                            updateProfileModal();
                        }
                        break;
                        
                    case 'home':
                    case 'back-to-menu':
                        // Go back to startup menu
                        if (typeof showStartupMenu === 'function') {
                            showStartupMenu();
                            slideMenu.classList.remove('open');
                        }
                        break;
                }
            });
        }
        
        // Profile modal close
        modalCloseButton?.addEventListener('click', () => {
            profileModal.style.display = 'none';
        });
        
        // Close modal when clicking outside
        window.addEventListener('click', (event) => {
            if (event.target === profileModal) {
                profileModal.style.display = 'none';
            }
            if (event.target === accountModal) {
                accountModal.style.display = 'none';
            }
        });
        
        // Account modal close button
        accountModalCloseButton?.addEventListener('click', () => {
            accountModal.style.display = 'none';
        });
        
        // Logout button in account modal
        const logoutButtonModal = document.getElementById('logout-button-modal');
        logoutButtonModal?.addEventListener('click', () => {
            // Call existing logout functionality
            handleSignOut();
            accountModal.style.display = 'none';
        });
    }
    
    // Update account modal with user data
    function updateAccountModal() {
        const userData = userProfileManager.getUserData();
        const subscriptionData = userProfileManager.getSubscriptionData();

        if (userData && userData.email) {
            // Update avatar text
            const avatarText = document.getElementById('user-avatar-text-modal');
            if (avatarText) {
                const initials = userData.name ? userData.name.split(' ').map(n => n[0]).join('').toUpperCase().slice(0, 2) : 'U';
                avatarText.textContent = initials;
            }
            
            // Update email
            const emailElement = document.getElementById('user-email-modal');
            if (emailElement) {
                emailElement.textContent = userData.email;
            }

            // Update credits
            const creditsEl = document.getElementById('user-credits-modal');
            if (creditsEl) {
                const credits = typeof userData.credits === 'number' ? userData.credits : getUserCredits();
                creditsEl.textContent = credits != null ? credits.toFixed(1) : '--';
            }

            // Update subscription type
            const subEl = document.getElementById('user-subscription-modal');
            if (subEl) {
                const status = subscriptionData?.status || (userProfileManager.hasActiveSubscription() ? 'active' : 'free');
                subEl.textContent = status;
            }
        }
    }
    
    // Update profile modal with user data
    function updateProfileModal() {
        const userData = userProfileManager.getUserData();
        
        if (userData && userData.email) {
            // Update avatar
            const profileAvatar = document.getElementById('profile-user-avatar');
            const profileInitials = document.getElementById('profile-user-initials');
            
            if (userData.picture) {
                profileAvatar.src = userData.picture;
                profileAvatar.style.display = 'block';
                profileInitials.style.display = 'none';
            } else {
                profileAvatar.style.display = 'none';
                profileInitials.style.display = 'flex';
                profileInitials.textContent = userData.name.split(' ').map(n => n[0]).join('').toUpperCase();
            }
            
            // Update user details
            document.getElementById('profile-user-name').textContent = userData.name;
            document.getElementById('profile-user-email').textContent = userData.email;
        }
        
        // Handle sign out button
        const profileSignOutButton = document.getElementById('profile-sign-out-button');
        if (profileSignOutButton) {
            profileSignOutButton.onclick = () => {
                handleSignOut();
                document.getElementById('profile-modal').style.display = 'none';
            };
        }
    }
    
    // Update footer with credits and subscription info
    function updateFooterDisplay() {
        const credits = getUserCredits();
        const userData = userProfileManager.getUserData();
        
        // Update credits count
        const footerCreditsCount = document.getElementById('footer-credits-count');
        if (footerCreditsCount) {
            footerCreditsCount.textContent = credits !== null ? credits.toFixed(1) : '--';
        }
        
        // Update subscription plan
        const footerPlanName = document.getElementById('footer-plan-name');
        if (footerPlanName && userData) {
            const subscriptionType = userData.subscription?.type || 'Free';
            footerPlanName.textContent = subscriptionType.charAt(0).toUpperCase() + subscriptionType.slice(1);
        }
        
        // Show/hide sign in button based on auth state
        const footerSignInButton = document.getElementById('footer-sign-in');
        if (footerSignInButton) {
            footerSignInButton.style.display = userData && userData.email ? 'none' : 'block';
        }
    }
    
    // Make updateFooterDisplay globally available
    window.updateFooterDisplay = updateFooterDisplay;
    
    // Initialize header menu and footer
    initializeHeaderMenu();
    updateFooterDisplay();

    // ... (rest of your Office.onReady, e.g., Promise.all)
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

// Function to get selected text from the code editor
function getSelectedTextFromEditor() {
    const codesTextarea = document.getElementById('codes-textarea');
    if (!codesTextarea) {
        return '';
    }
    
    const start = codesTextarea.selectionStart;
    const end = codesTextarea.selectionEnd;
    
    if (start === end) {
        // No text selected, return empty string
        return '';
    }
    
    return codesTextarea.value.substring(start, end);
}

// Function to remove TAB codes from text
function removeTABCodes(text) {
    if (!text || typeof text !== 'string') {
        return text;
    }
    
    // Remove any <> brackets that contain the word "TAB"
    const tabCodeRegex = /<[^>]*TAB[^>]*>/g;
    const cleanedText = text.replace(tabCodeRegex, '');
    
    console.log("[removeTABCodes] Original text length:", text.length);
    console.log("[removeTABCodes] Cleaned text length:", cleanedText.length);
    console.log("[removeTABCodes] TAB codes removed:", text.length - cleanedText.length > 0);
    
    return cleanedText;
}

// Function to remove commas from text
function removeCommas(text) {
    if (!text || typeof text !== 'string') {
        return text;
    }
    
    const cleanedText = text.replace(/,/g, '');
    
    console.log("[removeCommas] Original text length:", text.length);
    console.log("[removeCommas] Cleaned text length:", cleanedText.length);
    console.log("[removeCommas] Commas removed:", text.length - cleanedText.length > 0);
    
    return cleanedText;
}

// Function to show training data modal
async function addToTrainingDataQueue() {
    try {
        // Use the first user input (conversation starter) for training data
        const userPrompt = firstUserInput || '';
        
        // Get AI response - convert array to string if needed
        let aiResponse = '';
        if (Array.isArray(lastResponse)) {
            aiResponse = lastResponse.join('\n');
        } else if (typeof lastResponse === 'string') {
            aiResponse = lastResponse;
        }
        
        // Show the training data modal
        showTrainingDataModal(userPrompt, aiResponse);
        
    } catch (error) {
        console.error("Error opening training data modal:", error);
        showError(`Error opening training data dialog: ${error.message}`);
    }
}

// Function to show the training data modal
function showTrainingDataModal(userPrompt = '', aiResponse = '') {
    const modal = document.getElementById('training-data-modal');
    const userInputField = document.getElementById('training-user-input');
    const aiResponseField = document.getElementById('training-ai-response');
    
    if (!modal || !userInputField || !aiResponseField) {
        showError('Training data modal elements not found');
        return;
    }
    
    // Use persisted values if they exist, otherwise use provided parameters
    const finalUserPrompt = persistedTrainingUserInput !== null ? persistedTrainingUserInput : userPrompt;
    const finalAiResponse = persistedTrainingAiResponse !== null ? persistedTrainingAiResponse : aiResponse;
    
    // Populate the fields
    userInputField.value = finalUserPrompt;
    aiResponseField.value = finalAiResponse;
    
    console.log("[showTrainingDataModal] Using persisted values:", {
        userInputPersisted: persistedTrainingUserInput !== null,
        aiResponsePersisted: persistedTrainingAiResponse !== null
    });
    
    // Initialize find/replace functionality for training modal
    resetTrainingModalSearchState();
    
    // Ensure paste functionality works by adding explicit event handlers
    const enablePasteForElement = (element) => {
        // Remove any existing paste handlers to avoid duplicates
        element.removeEventListener('paste', handlePasteEvent);
        
        // Add paste event handler
        element.addEventListener('paste', handlePasteEvent);
        
        // Also ensure the element is properly focusable
        element.setAttribute('tabindex', '0');
        
        // Add context menu support (right-click paste)
        element.addEventListener('contextmenu', (e) => {
            // Allow default context menu behavior
            e.stopPropagation();
        });
        
        // Add keyboard shortcut support
        element.addEventListener('keydown', (e) => {
            if ((e.ctrlKey || e.metaKey) && e.key === 'v') {
                // Allow default paste behavior
                e.stopPropagation();
            }
        });
    };
    
    // Enable paste for both textarea elements
    enablePasteForElement(userInputField);
    enablePasteForElement(aiResponseField);
    
    // Show the modal
    modal.style.display = 'block';
    
    // Focus the first field to ensure proper initialization
    setTimeout(() => {
        userInputField.focus();
        
        // Debug logging to help troubleshoot paste issues
        console.log('Training data modal opened successfully');
        console.log('User input field can focus:', document.activeElement === userInputField);
        console.log('User input field readonly:', userInputField.readOnly);
        console.log('User input field disabled:', userInputField.disabled);
        console.log('AI response field readonly:', aiResponseField.readOnly);
        console.log('AI response field disabled:', aiResponseField.disabled);
        
        // Test if paste events can be triggered
        try {
            const testEvent = new ClipboardEvent('paste', {
                bubbles: true,
                cancelable: true,
                clipboardData: new DataTransfer()
            });
            console.log('Clipboard event creation test passed');
        } catch (e) {
            console.warn('Clipboard event creation test failed:', e.message);
        }
    }, 100);
}

// Helper function to handle paste events
function handlePasteEvent(e) {
    try {
        // Allow the default paste behavior
        console.log('Paste event detected in training data field');
        
        // For debugging - log what's being pasted (if accessible)
        if (e.clipboardData && e.clipboardData.getData) {
            try {
                const pastedText = e.clipboardData.getData('text');
                console.log('Pasted text length:', pastedText.length);
            } catch (clipboardError) {
                console.log('Could not access clipboard data for logging');
            }
        }
        
        // Don't prevent default - let the browser handle the paste
        return true;
    } catch (error) {
        console.error('Error in paste handler:', error);
        // Still allow the paste to proceed
        return true;
    }
}

// Function to hide the training data modal
function hideTrainingDataModal() {
    const modal = document.getElementById('training-data-modal');
    const userInputField = document.getElementById('training-user-input');
    const aiResponseField = document.getElementById('training-ai-response');
    
    // Save current field values before hiding
    if (userInputField && aiResponseField) {
        persistedTrainingUserInput = userInputField.value;
        persistedTrainingAiResponse = aiResponseField.value;
        console.log("[hideTrainingDataModal] Saved field values for persistence");
    }
    
    if (modal) {
        modal.style.display = 'none';
    }
}

// Function to save training data from modal
async function saveTrainingDataFromModal() {
    try {
        const userInputField = document.getElementById('training-user-input');
        const aiResponseField = document.getElementById('training-ai-response');
        
        if (!userInputField || !aiResponseField) {
            showError('Training data input fields not found');
            return;
        }
        
        const userPrompt = userInputField.value.trim();
        const aiResponse = aiResponseField.value.trim();
        
        // Validate inputs - require at least one field to be filled
        if (!userPrompt && !aiResponse) {
            showError("Please enter either a client message or AI response before saving.");
            return;
        }
        
        const trainingEntry = {
            prompt: userPrompt,
            selectedCode: aiResponse
        };
        
        console.log("[saveTrainingDataFromModal] Created training entry:", trainingEntry);
        
        // Load existing training queue and add new entry
        let trainingQueue = [];
        try {
            const existingQueue = localStorage.getItem('trainingDataQueue');
            if (existingQueue) {
                const parsed = JSON.parse(existingQueue);
                // Only use existing data if it's in the new format (no timestamp, hasPrompt, etc.)
                if (parsed.length > 0 && !parsed[0].timestamp && parsed[0].hasPrompt === undefined) {
                    trainingQueue = parsed;
                    console.log("[saveTrainingDataFromModal] Loaded existing queue with", trainingQueue.length, "entries");
                } else {
                    console.log("[saveTrainingDataFromModal] Found old format data, starting fresh");
                    trainingQueue = [];
                }
            } else {
                console.log("[saveTrainingDataFromModal] No existing queue found, starting fresh");
            }
        } catch (error) {
            console.warn("Error loading existing training queue:", error);
            trainingQueue = [];
        }
        
        // Add new entry to queue
        trainingQueue.push(trainingEntry);
        console.log("[saveTrainingDataFromModal] Queue now has", trainingQueue.length, "entries");
        
        // Save back to localStorage
        localStorage.setItem('trainingDataQueue', JSON.stringify(trainingQueue));
        
        // Create TXT content
        const txtContent = convertTrainingQueueToTXT(trainingQueue);
        console.log("[saveTrainingDataFromModal] Generated TXT content:", txtContent);
        
        // Save TXT file (using browser download)
        downloadTXTFile(txtContent, `training_data_queue.txt`);
        
        // Show success message
        showMessage(`Training data added to queue! Entry ${trainingQueue.length} saved. TXT file downloaded.`);
        
        // Clear persisted values since data was successfully saved
        persistedTrainingUserInput = null;
        persistedTrainingAiResponse = null;
        
        // Hide the modal
        hideTrainingDataModal();
        
        console.log("Training data entry added:", trainingEntry);
        
    } catch (error) {
        console.error("Error saving training data:", error);
        showError(`Error saving training data: ${error.message}`);
    }
}

// Function to convert training queue to TXT format
function convertTrainingQueueToTXT(trainingQueue) {
    if (!trainingQueue || trainingQueue.length === 0) {
        return ''; // Return empty string if no data
    }
    
    let txt = '';
    
    // Add each entry - only prompt and selected code, no quotes
    trainingQueue.forEach(entry => {
        // Don't escape or quote the fields, just use them as-is
        const prompt = entry.prompt || '';
        // Replace newlines in code with spaces to keep it on one line
        const code = (entry.selectedCode || '').replace(/\n/g, ' ').replace(/\r/g, ' ');
        
        txt += `${prompt}^^^${code}@\n`; // Added @ symbol before newline
    });
    
    return txt;
}



// Function to download TXT file
function downloadTXTFile(txtContent, filename) {
    try {
        // Create blob with TXT content
        const blob = new Blob([txtContent], { type: 'text/plain;charset=utf-8;' });
        
        // Create download link
        const link = document.createElement('a');
        const url = URL.createObjectURL(blob);
        
        link.setAttribute('href', url);
        link.setAttribute('download', filename);
        link.style.visibility = 'hidden';
        
        // Add to document, click, and remove
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        
        // Clean up the URL object
        URL.revokeObjectURL(url);
        
        console.log(`TXT file "${filename}" download initiated`);
        
    } catch (error) {
        console.error("Error downloading TXT file:", error);
        // Fallback: log the TXT content to console
        console.log("TXT Content (download failed):", txtContent);
        showError("Could not download TXT file, but data was saved to localStorage. Check console for TXT content.");
    }
}

// Function to clear old training data (for debugging)
function clearTrainingDataQueue() {
    try {
        localStorage.removeItem('trainingDataQueue');
        console.log("Training data queue cleared from localStorage");
        showMessage("Training data queue cleared successfully!");
    } catch (error) {
        console.error("Error clearing training data queue:", error);
        showError(`Error clearing training data: ${error.message}`);
    }
}

// Make it globally accessible for debugging
window.clearTrainingDataQueue = clearTrainingDataQueue;

// >>> ADDED: Microsoft Authentication System
let msalInstance;
let currentUser = null;

/**
 * Initialize Microsoft Authentication Library (MSAL)
 */
function initializeMSAL() {
    try {
        console.log('Initializing MSAL...');
        
        // Check if MSAL library is loaded
        if (typeof msal === 'undefined') {
            console.error('MSAL library not loaded - authentication will not work');
            showError('Authentication library not loaded. Please refresh the page.');
            return false;
        }
        
        console.log('MSAL library found, checking configuration...');
        
        // Validate configuration
        if (!CONFIG || !CONFIG.authentication || !CONFIG.authentication.msalConfig) {
            console.error('MSAL configuration not found');
            showError('Authentication configuration missing');
            return false;
        }
        
        const config = CONFIG.authentication.msalConfig;
        console.log('MSAL config:', {
            clientId: config.auth.clientId,
            authority: config.auth.authority,
            redirectUri: config.auth.redirectUri
        });
        
        // Check for placeholder values
        if (CONFIG.authentication.isDevelopmentPlaceholder()) {
            console.warn('MSAL using placeholder client ID - authentication will not work');
            
            const isDev = process.env.NODE_ENV === 'development';
            const setupMessage = isDev 
                ? 'Microsoft authentication is not configured. Please follow the setup guide in MICROSOFT_AUTH_SETUP.md to create an Azure app registration.'
                : 'Microsoft authentication is not configured for production. Please update the client ID in config.js';
            
            showError(setupMessage);
            return false;
        }
        
        // Create MSAL instance
        console.log('Creating MSAL instance...');
        msalInstance = new msal.PublicClientApplication(config);
        
        // Handle redirect response
        msalInstance.handleRedirectPromise()
            .then(handleAuthResponse)
            .catch(error => {
                console.error('Error handling redirect response:', error);
                showError('Authentication redirect error: ' + error.message);
            });

        // Check if user is already signed in
        checkAuthState();
        
        console.log('MSAL initialized successfully');
        return true;
        
    } catch (error) {
        console.error('Failed to initialize MSAL:', error);
        showError('Authentication system initialization failed: ' + error.message);
        msalInstance = null; // Ensure it's null on failure
        return false;
    }
}

/**
 * Handle authentication response
 */
function handleAuthResponse(response) {
    if (response && response.account) {
        console.log('Authentication successful:', response);
        currentUser = response.account;
        updateAuthUI(true);
        getUserProfile();
    } else {
        console.log('No authentication response or account');
        updateAuthUI(false);
    }
}

/**
 * Check current authentication state
 */
function checkAuthState() {
    try {
        if (!msalInstance) {
            console.log('MSAL not initialized - cannot check auth state');
            updateAuthUI(false);
            return;
        }
        
        const currentAccounts = msalInstance.getAllAccounts();
        if (currentAccounts.length > 0) {
            currentUser = currentAccounts[0];
            updateAuthUI(true);
            getUserProfile();
            console.log('User already authenticated:', currentUser);
        } else {
            updateAuthUI(false);
            console.log('No authenticated user found');
        }
    } catch (error) {
        console.error('Error checking auth state:', error);
        updateAuthUI(false);
    }
}

/**
 * Sign in user
 */
async function signIn() {
    try {
        console.log('Initiating sign-in...');
        
        // Check if Google authentication is already active
        const googleUser = sessionStorage.getItem('googleUser');
        if (googleUser) {
            console.log('User already authenticated with Google - skipping MSAL sign-in');
            showMessage('You are already signed in with Google!');
            return;
        }
        
        // Check if MSAL is initialized
        if (!msalInstance) {
            console.error('MSAL not initialized - cannot sign in');
            showError('Microsoft authentication not available. Please use Google Sign-In from the Authentication menu.');
            return;
        }
        
        // Check if configuration is valid
        if (!CONFIG.authentication.loginRequest) {
            console.error('Login request configuration missing');
            showError('Authentication configuration incomplete');
            return;
        }
        
        console.log('Starting login popup...');
        const response = await msalInstance.loginPopup(CONFIG.authentication.loginRequest);
        console.log('Sign-in successful:', response);
        
        if (response && response.account) {
            currentUser = response.account;
            updateAuthUI(true);
            getUserProfile();
            updateSignedInStatus();
            showMessage('Successfully signed in to Microsoft!');
        } else {
            console.error('No account in sign-in response');
            showError('Sign-in completed but no account information received');
        }
        
    } catch (error) {
        console.error('Sign-in failed:', error);
        
        if (error.errorCode === 'popup_window_error' || error.errorCode === 'user_cancelled') {
            showError('Sign-in was cancelled or blocked. Please allow popups and try again.');
        } else if (error.errorCode === 'interaction_in_progress') {
            showError('Another sign-in is already in progress. Please wait and try again.');
        } else if (error.errorCode === 'invalid_request') {
            showError('Invalid authentication request. Please check the configuration.');
        } else {
            showError(`Sign-in failed: ${error.errorMessage || error.message || 'Unknown error'}`);
        }
    }
}

/**
 * Sign out user
 */
async function signOut() {
    try {
        console.log('Signing out...');
        
        // Check if we're using Google authentication and call the proper logout function
        if (typeof window.handleGoogleSignOutView === 'function') {
            console.log('Using Google logout function');
            const result = window.handleGoogleSignOutView();
            if (result && result.success) {
                console.log('Google logout completed successfully');
            }
        } else {
            console.log('Google logout function not available, clearing local storage');
            // Fallback: Clear Google authentication state from sessionStorage
            sessionStorage.removeItem('googleUser');
            sessionStorage.removeItem('googleToken');
            sessionStorage.removeItem('googleCredential');
            
            // Clear any other auth-related session data
            sessionStorage.removeItem('authToken');
            sessionStorage.removeItem('userInfo');
            
            console.log('Google authentication state cleared from sessionStorage');
        }
        
        // Only try MSAL logout if MSAL was actually initialized and used
        try {
            if (typeof msalInstance !== 'undefined' && msalInstance && currentUser) {
                console.log('Attempting MSAL logout');
                await msalInstance.logoutPopup();
                console.log('MSAL logout completed');
            }
        } catch (msalError) {
            console.log('MSAL logout not applicable or failed:', msalError.message);
        }
        
        currentUser = null;
        updateAuthUI(false);
        updateSignedInStatus();
        
        // Redirect back to authentication view
        const authView = document.getElementById('authentication-view');
        const clientView = document.getElementById('client-mode-view');
        if (authView && clientView) {
            authView.style.display = 'block';
            clientView.style.display = 'none';
            console.log('Redirected to authentication view');
        }
        
        console.log('Sign-out successful');
        showMessage('Successfully signed out');
    } catch (error) {
        console.error('Sign-out failed:', error);
        // Even if logout fails, clear local state
        currentUser = null;
        updateAuthUI(false);
        
        // Force clear sessionStorage even on error
        sessionStorage.removeItem('googleUser');
        sessionStorage.removeItem('googleToken');
        sessionStorage.removeItem('googleCredential');
        sessionStorage.removeItem('authToken');
        sessionStorage.removeItem('userInfo');
        
        showError('Sign-out encountered an error but local session was cleared');
    }
}

/**
 * Get user profile from Microsoft Graph
 */
async function getUserProfile() {
    try {
        if (!currentUser) {
            console.log('No current user for profile fetch');
            return;
        }

        // Get access token
        const tokenResponse = await msalInstance.acquireTokenSilent({
            ...CONFIG.authentication.tokenRequest,
            account: currentUser
        });

        // Fetch user profile
        const profileResponse = await fetch(CONFIG.authentication.graphConfig.graphMeUrl, {
            headers: {
                'Authorization': `Bearer ${tokenResponse.accessToken}`
            }
        });

        if (profileResponse.ok) {
            const profile = await profileResponse.json();
            console.log('User profile:', profile);
            updateUserProfile(profile);
            
            // Initialize user data from backend after successful sign-in
            try {
                await initializeUserData();
            } catch (error) {
                console.warn("Failed to initialize user data after sign-in:", error);
                forceUpdateUserDisplay();
            }
            
            // Try to get user photo
            getUserPhoto(tokenResponse.accessToken);
        } else {
            console.error('Failed to fetch user profile:', profileResponse.status);
        }
    } catch (error) {
        console.error('Error getting user profile:', error);
        // If we can't get extended profile, just use basic account info
        updateUserProfile({
            displayName: currentUser.name || 'User',
            mail: currentUser.username || '',
            userPrincipalName: currentUser.username || ''
        });
    }
}

/**
 * Get user photo from Microsoft Graph
 */
async function getUserPhoto(accessToken) {
    try {
        const photoResponse = await fetch(CONFIG.authentication.graphConfig.graphPhotoUrl, {
            headers: {
                'Authorization': `Bearer ${accessToken}`
            }
        });

        if (photoResponse.ok) {
            const photoBlob = await photoResponse.blob();
            const photoUrl = URL.createObjectURL(photoBlob);
            
            const avatarImg = document.getElementById('user-avatar');
            const initialsDiv = document.getElementById('user-initials');
            
            if (avatarImg) {
                avatarImg.src = photoUrl;
                avatarImg.style.display = 'block';
                if (initialsDiv) {
                    initialsDiv.style.display = 'none';
                }
            }
        } else {
            console.log('No user photo available or failed to fetch');
        }
    } catch (error) {
        console.log('Error getting user photo:', error);
        // Photo is optional, so we just log the error
    }
}

/**
 * Update user profile display
 */
function updateUserProfile(profile) {
    const userNameElement = document.getElementById('user-name');
    const userEmailElement = document.getElementById('user-email');
    const userInitialsElement = document.getElementById('user-initials');

    if (userNameElement) {
        userNameElement.textContent = profile.displayName || profile.name || 'User';
    }

    if (userEmailElement) {
        userEmailElement.textContent = profile.mail || profile.userPrincipalName || '';
    }

    if (userInitialsElement) {
        const name = profile.displayName || profile.name || 'User';
        const initials = name.split(' ')
            .map(n => n.charAt(0))
            .join('')
            .substring(0, 2)
            .toUpperCase();
        userInitialsElement.textContent = initials;
    }
}

/**
 * Update authentication UI based on state
 */
function updateAuthUI(isAuthenticated) {
    const signInButton = document.getElementById('sign-in-button');
    const userInfo = document.getElementById('user-info');

    console.log('Updating auth UI - authenticated:', isAuthenticated);

    if (isAuthenticated) {
        // User is signed in - show user info, hide sign-in button
        if (signInButton) {
            signInButton.style.display = 'none';
            console.log('Sign-in button hidden - user authenticated');
        }
        if (userInfo) {
            userInfo.style.display = 'flex';
            console.log('User info shown');
        }
    } else {
        // User is not signed in - show sign-in button, hide user info
        if (signInButton) {
            signInButton.style.display = 'flex';
            console.log('Sign-in button shown - user not authenticated');
        }
        if (userInfo) {
            userInfo.style.display = 'none';
            console.log('User info hidden');
        }
        
        // Clear user display
        const userAvatar = document.getElementById('user-avatar');
        const userInitials = document.getElementById('user-initials');
        if (userAvatar) {
            userAvatar.style.display = 'none';
            userAvatar.src = '';
        }
        if (userInitials) {
            userInitials.style.display = 'flex';
            userInitials.textContent = '';
        }
    }
}

/**
 * Get current authenticated user (prioritizes Google, then Microsoft)
 */
function getCurrentUser() {
    // Check Google authentication first
    const googleUserStr = sessionStorage.getItem('googleUser');
    if (googleUserStr) {
        try {
            return JSON.parse(googleUserStr);
        } catch (e) {
            console.error('Error parsing Google user data:', e);
        }
    }
    
    // Fallback to Microsoft authentication
    return currentUser;
}

/**
 * Check if user is authenticated (checks Google first, then Microsoft)
 */
function isUserAuthenticated() {
    // Check Google authentication first - check both session and local storage
    const googleUserSession = sessionStorage.getItem('googleUser');
    const googleCredentialSession = sessionStorage.getItem('googleCredential') || sessionStorage.getItem('googleToken');
    
    const googleUserLocal = localStorage.getItem('googleUser');
    const googleCredentialLocal = localStorage.getItem('googleCredential') || localStorage.getItem('google_access_token');
    
    if ((googleUserSession && googleCredentialSession) || (googleUserLocal && googleCredentialLocal)) {
        return true;
    }
    
    // Check other auth methods in localStorage
    const msalToken = localStorage.getItem('msal_access_token');
    const apiKey = localStorage.getItem('user_api_key');
    
    if (msalToken || apiKey) {
        return true;
    }
    
    // Fallback to Microsoft authentication
    return currentUser !== null;
}

/**
 * Test authentication setup - useful for debugging
 */
function testAuthSetup() {
    console.log('=== Microsoft Authentication Setup Test ===');
    
    // Check MSAL library
    console.log('1. MSAL Library:', typeof msal !== 'undefined' ? '✅ Loaded' : '❌ Not loaded');
    
    // Check configuration
    const hasConfig = CONFIG && CONFIG.authentication && CONFIG.authentication.msalConfig;
    console.log('2. Configuration:', hasConfig ? '✅ Found' : '❌ Missing');
    
    if (hasConfig) {
        const config = CONFIG.authentication.msalConfig;
        console.log('   - Client ID:', config.auth.clientId);
        console.log('   - Authority:', config.auth.authority);
        console.log('   - Redirect URI:', config.auth.redirectUri);
        console.log('   - Is Placeholder:', CONFIG.authentication.isDevelopmentPlaceholder() ? '⚠️ Yes' : '✅ No');
    }
    
    // Check MSAL instance
    console.log('3. MSAL Instance:', msalInstance ? '✅ Initialized' : '❌ Not initialized');
    
    // Check authentication state
    console.log('4. User Authentication:', isUserAuthenticated() ? '✅ Signed in' : '❌ Not signed in');
    
    if (isUserAuthenticated() && currentUser) {
        console.log('   - User:', currentUser.name || currentUser.username);
    }
    
    console.log('=== End Test ===');
    
    return {
        msalLoaded: typeof msal !== 'undefined',
        configExists: hasConfig,
        isPlaceholder: hasConfig ? CONFIG.authentication.isDevelopmentPlaceholder() : true,
        msalInitialized: !!msalInstance,
        userAuthenticated: isUserAuthenticated()
    };
}

// Make test function globally available for debugging
window.testAuthSetup = testAuthSetup;
// <<< END ADDED

// >>> ADDED: Voice Recording Implementation for Client Mode
// Voice recording variables
let mediaRecorderClient = null;
let audioChunksClient = [];
let recordingStartTimeClient = null;
let recordingTimerIntervalClient = null;
let audioContextClient = null;
let analyserClient = null;
let waveformCanvasClient = null;
let waveformAnimationIdClient = null;
let cursorPositionBeforeRecording = 0;

// Initialize voice recording for client mode
function initializeVoiceRecordingClient() {
    // Check if getUserMedia is supported
    if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) {
        console.warn('[Voice-Client] getUserMedia API not supported in this browser');
        const voiceButton = document.getElementById('voice-input-client');
        if (voiceButton) {
            voiceButton.style.display = 'none';
        }
        return;
    }
    
    if (!window.MediaRecorder) {
        console.warn('[Voice-Client] MediaRecorder API not supported in this browser');
        const voiceButton = document.getElementById('voice-input-client');
        if (voiceButton) {
            voiceButton.style.display = 'none';
        }
        return;
    }
    
    // Get voice button
    const voiceButton = document.getElementById('voice-input-client');
    
    // Voice button click - start recording immediately
    if (voiceButton) {
        voiceButton.addEventListener('click', () => {
            console.log('[Voice-Client] Voice button clicked - switching to voice mode');
            switchToVoiceModeClient();
        });
    }
}

// Switch to voice recording mode
function switchToVoiceModeClient() {
    const inputBar = document.querySelector('.chatgpt-input-bar');
    const textArea = document.getElementById('user-input-client');
    
    if (!inputBar || !textArea) {
        console.error('[Voice-Client] Required elements not found');
        return;
    }
    
    // Reset debug flag
    window.waveformDebugLogged = false;
    
    // Store cursor position
    cursorPositionBeforeRecording = textArea.selectionStart;
    
    // Create voice recording UI
    const voiceRecordingUI = createVoiceRecordingUI();
    
    // Hide the input bar and show voice recording UI
    inputBar.style.display = 'none';
    inputBar.parentNode.insertBefore(voiceRecordingUI, inputBar);
    
    // Initialize waveform after a slight delay to ensure DOM is ready
    setTimeout(() => {
        initializeWaveformClient();
        // Start recording immediately
        startVoiceRecordingClient();
    }, 50);
}

// Create voice recording UI
function createVoiceRecordingUI() {
    const voiceRecordingMode = document.createElement('div');
    voiceRecordingMode.id = 'voice-recording-mode-client';
    voiceRecordingMode.className = 'voice-recording-mode';
    voiceRecordingMode.style.cssText = `
        display: flex;
        align-items: center;
        justify-content: space-between;
        padding: 12px 16px;
        background: #f3f4f6;
        border-radius: 8px;
        margin: 0 16px 16px 16px;
        gap: 12px;
    `;
    
    voiceRecordingMode.innerHTML = `
        <div style="flex: 1; display: flex; align-items: center; gap: 12px;">
            <canvas id="voice-waveform-client" style="flex: 1; height: 40px; background: transparent;"></canvas>
            <div id="voice-recording-timer-client" style="font-size: 14px; color: #6b7280; font-weight: 500;">00:00</div>
        </div>
        <div style="display: flex; gap: 4px; align-items: center;">
            <button id="cancel-voice-recording-client" style="
                padding: 4px;
                background: transparent;
                color: #6b7280;
                border: none;
                cursor: pointer;
                display: flex;
                align-items: center;
                justify-content: center;
                transition: color 0.2s;
            " title="Cancel recording" 
               onmouseover="this.style.color='#374151'" 
               onmouseout="this.style.color='#6b7280'">
                <svg viewBox="0 0 16 16" style="width: 14px; height: 14px;">
                    <path d="M12 4L4 12M4 4l8 8" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/>
                </svg>
            </button>
            <button id="accept-voice-recording-client" style="
                padding: 4px;
                background: transparent;
                color: #6b7280;
                border: none;
                cursor: pointer;
                display: flex;
                align-items: center;
                justify-content: center;
                transition: color 0.2s;
            " title="Done recording"
               onmouseover="this.style.color='#374151'" 
               onmouseout="this.style.color='#6b7280'">
                <svg viewBox="0 0 16 16" style="width: 14px; height: 14px;">
                    <path d="M3 8l3 3L13 4" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" fill="none"/>
                </svg>
            </button>
        </div>
    `;
    
    // Add event listeners
    setTimeout(() => {
        const cancelBtn = document.getElementById('cancel-voice-recording-client');
        const acceptBtn = document.getElementById('accept-voice-recording-client');
        
        if (cancelBtn) {
            cancelBtn.onclick = () => {
                console.log('[Voice-Client] Cancel recording clicked');
                switchToNormalModeClient();
            };
        }
        
        if (acceptBtn) {
            acceptBtn.onclick = () => {
                console.log('[Voice-Client] Accept recording clicked');
                stopVoiceRecordingClient();
            };
        }
    }, 0);
    
    return voiceRecordingMode;
}

// Initialize waveform canvas
function initializeWaveformClient() {
    waveformCanvasClient = document.getElementById('voice-waveform-client');
    if (waveformCanvasClient) {
        const ctx = waveformCanvasClient.getContext('2d');
        
        // Force a layout refresh to ensure dimensions are calculated
        waveformCanvasClient.offsetHeight; // Force reflow
        
        // Set canvas size
        const rect = waveformCanvasClient.getBoundingClientRect();
        const width = rect.width || 300; // Fallback width
        const height = rect.height || 40; // Fallback height
        
        waveformCanvasClient.width = width;
        waveformCanvasClient.height = height;
        
        console.log('[Voice-Client] Canvas dimensions:', { width, height, rect });
        
        // Clear canvas
        ctx.clearRect(0, 0, waveformCanvasClient.width, waveformCanvasClient.height);
        
        console.log('[Voice-Client] Waveform canvas initialized');
    } else {
        console.error('[Voice-Client] Waveform canvas element not found');
    }
}

// Draw waveform animation (using the same bar-style animation as dev side)
function drawWaveformClient() {
    if (!waveformCanvasClient || !analyserClient) {
        console.warn('[Voice-Client] Canvas or analyser not ready');
        return;
    }
    
    const ctx = waveformCanvasClient.getContext('2d');
    const rect = waveformCanvasClient.getBoundingClientRect();
    const width = rect.width || waveformCanvasClient.width;
    const height = rect.height || waveformCanvasClient.height;
    
    // Debug log on first frame
    if (!window.waveformDebugLogged) {
        console.log('[Voice-Client] Drawing waveform:', { width, height, canvas: waveformCanvasClient });
        window.waveformDebugLogged = true;
    }
    
    // Get time domain data for real-time audio visualization
    const bufferLength = analyserClient.fftSize;
    const dataArray = new Uint8Array(bufferLength);
    analyserClient.getByteTimeDomainData(dataArray);
    
    // Clear canvas with transparent background
    ctx.clearRect(0, 0, width, height);
    
    // Calculate how many bars we want to show
    const numBars = Math.min(60, Math.floor(width / 4)); // Adjust bar count based on width
    const barWidth = (width - (numBars - 1)) / numBars; // Account for spacing
    
    // Process audio data to get average values for each bar
    const samplesPerBar = Math.floor(bufferLength / numBars);
    
    for (let i = 0; i < numBars; i++) {
        let sum = 0;
        let max = 0;
        
        // Calculate average amplitude for this bar's samples
        for (let j = 0; j < samplesPerBar; j++) {
            const sampleIndex = i * samplesPerBar + j;
            if (sampleIndex < dataArray.length) {
                const amplitude = Math.abs(dataArray[sampleIndex] - 128);
                sum += amplitude;
                max = Math.max(max, amplitude);
            }
        }
        
        // Use a combination of average and max for more dynamic visualization
        const average = sum / samplesPerBar;
        const value = (average * 0.7 + max * 0.3) / 128; // Weighted combination
        
        // Calculate bar height with minimum and maximum constraints
        const minHeight = 3; // Minimum bar height
        const maxHeight = height * 0.8; // Maximum bar height
        let barHeight = value * height * 2; // Scale up for visibility
        
        // Apply logarithmic scaling for better visual dynamics
        barHeight = Math.log10(1 + barHeight * 9) * (height / 2);
        
        // Ensure bars are within constraints
        barHeight = Math.max(barHeight, minHeight);
        barHeight = Math.min(barHeight, maxHeight);
        
        // Calculate x position with spacing
        const x = i * (barWidth + 1);
        
        // Draw the bar centered vertically
        const y = (height - barHeight) / 2;
        
        ctx.fillStyle = '#000000';
        ctx.fillRect(x, y, barWidth, barHeight);
    }
    
    // Continue animation - this ensures continuous updates
    if (mediaRecorderClient && mediaRecorderClient.state === 'recording') {
        waveformAnimationIdClient = requestAnimationFrame(drawWaveformClient);
    }
}

// Start voice recording
async function startVoiceRecordingClient() {
    try {
        console.log('[Voice-Client] Requesting microphone access...');
        
        const stream = await navigator.mediaDevices.getUserMedia({ 
            audio: {
                echoCancellation: true,
                noiseSuppression: true,
                sampleRate: 44100
            } 
        });
        
        console.log('[Voice-Client] Microphone access granted');
        
        // Setup audio context for waveform visualization
        audioContextClient = new (window.AudioContext || window.webkitAudioContext)();
        analyserClient = audioContextClient.createAnalyser();
        const source = audioContextClient.createMediaStreamSource(stream);
        source.connect(analyserClient);
        
        // Configure analyser
        analyserClient.fftSize = 1024;
        analyserClient.smoothingTimeConstant = 0.3;
        analyserClient.minDecibels = -90;
        analyserClient.maxDecibels = -10;
        
        // Start waveform animation
        drawWaveformClient();
        
        // Backup timer to ensure continuous animation (in case requestAnimationFrame fails)
        const backupTimer = setInterval(() => {
            if (mediaRecorderClient && mediaRecorderClient.state === 'recording' && !waveformAnimationIdClient) {
                console.log('[Voice-Client] Restarting waveform animation via backup timer');
                drawWaveformClient();
            } else if (!mediaRecorderClient || mediaRecorderClient.state !== 'recording') {
                clearInterval(backupTimer);
            }
        }, 100); // Check every 100ms
        
        // Create MediaRecorder
        const options = { mimeType: 'audio/webm;codecs=opus' };
        if (!MediaRecorder.isTypeSupported(options.mimeType)) {
            options.mimeType = 'audio/webm';
            if (!MediaRecorder.isTypeSupported(options.mimeType)) {
                options.mimeType = 'audio/mp4';
            }
        }
        
        mediaRecorderClient = new MediaRecorder(stream, options);
        audioChunksClient = [];
        
        mediaRecorderClient.ondataavailable = (event) => {
            if (event.data.size > 0) {
                audioChunksClient.push(event.data);
            }
        };
        
        mediaRecorderClient.onstop = async () => {
            console.log('[Voice-Client] Recording stopped');
            
            // Stop all tracks
            stream.getTracks().forEach(track => track.stop());
            
            // Create audio blob
            const audioBlob = new Blob(audioChunksClient, { type: 'audio/webm' });
            console.log('[Voice-Client] Audio blob created, size:', audioBlob.size);
            
            // Show loading state
            switchToLoadingModeClient();
            
            // Transcribe audio
            await transcribeAudioClient(audioBlob);
        };
        
        // Start recording
        mediaRecorderClient.start();
        recordingStartTimeClient = Date.now();
        
        // Start timer
        updateRecordingTimerClient();
        recordingTimerIntervalClient = setInterval(updateRecordingTimerClient, 100);
        
        console.log('[Voice-Client] Recording started');
        
    } catch (error) {
        console.error('[Voice-Client] Error starting recording:', error);
        showError('Failed to access microphone. Please check your permissions.');
        switchToNormalModeClient();
    }
}

// Update recording timer
function updateRecordingTimerClient() {
    if (!recordingStartTimeClient) return;
    
    const elapsed = Math.floor((Date.now() - recordingStartTimeClient) / 1000);
    const minutes = Math.floor(elapsed / 60);
    const seconds = elapsed % 60;
    
    const timerElement = document.getElementById('voice-recording-timer-client');
    if (timerElement) {
        timerElement.textContent = `${minutes.toString().padStart(2, '0')}:${seconds.toString().padStart(2, '0')}`;
    }
}

// Stop voice recording
function stopVoiceRecordingClient() {
    if (mediaRecorderClient && mediaRecorderClient.state === 'recording') {
        mediaRecorderClient.stop();
        
        // Stop timer
        if (recordingTimerIntervalClient) {
            clearInterval(recordingTimerIntervalClient);
            recordingTimerIntervalClient = null;
        }
        
        // Stop waveform animation
        if (waveformAnimationIdClient) {
            cancelAnimationFrame(waveformAnimationIdClient);
            waveformAnimationIdClient = null;
        }
    }
}

// Switch to loading mode
function switchToLoadingModeClient() {
    const voiceRecordingMode = document.getElementById('voice-recording-mode-client');
    
    if (voiceRecordingMode) {
        voiceRecordingMode.innerHTML = `
            <div style="display: flex; align-items: center; justify-content: center; gap: 12px; width: 100%;">
                <div style="
                    width: 24px;
                    height: 24px;
                    border: 3px solid #e5e7eb;
                    border-top-color: #3b82f6;
                    border-radius: 50%;
                    animation: spin 0.8s linear infinite;
                "></div>
                <span style="color: #6b7280; font-size: 14px;">Transcribing...</span>
            </div>
        `;
        
        // Add spinning animation if not already present
        if (!document.getElementById('voice-spin-style')) {
            const style = document.createElement('style');
            style.id = 'voice-spin-style';
            style.textContent = `
                @keyframes spin {
                    to { transform: rotate(360deg); }
                }
            `;
            document.head.appendChild(style);
        }
    }
}

// Switch back to normal mode
function switchToNormalModeClient() {
    // Remove voice recording UI
    const voiceRecordingMode = document.getElementById('voice-recording-mode-client');
    if (voiceRecordingMode) {
        voiceRecordingMode.remove();
    }
    
    // Show input bar again
    const inputBar = document.querySelector('.chatgpt-input-bar');
    if (inputBar) {
        inputBar.style.display = 'block';
    }
    
    // Stop any ongoing recording
    if (mediaRecorderClient && mediaRecorderClient.state === 'recording') {
        mediaRecorderClient.stop();
    }
    
    // Cleanup
    cleanupVoiceRecordingClient();
}

// Cleanup voice recording resources
function cleanupVoiceRecordingClient() {
    if (recordingTimerIntervalClient) {
        clearInterval(recordingTimerIntervalClient);
        recordingTimerIntervalClient = null;
    }
    
    if (waveformAnimationIdClient) {
        cancelAnimationFrame(waveformAnimationIdClient);
        waveformAnimationIdClient = null;
    }
    
    if (audioContextClient) {
        audioContextClient.close();
        audioContextClient = null;
    }
    
    analyserClient = null;
    mediaRecorderClient = null;
    audioChunksClient = [];
    recordingStartTimeClient = null;
}

// Transcribe audio
async function transcribeAudioClient(audioBlob) {
    try {
        console.log('[Voice-Client] Starting transcription...');
        
        // Get API key from AIcalls.js
        const { initializeAPIKeys } = await import('./AIcalls.js');
        const apiKeys = await initializeAPIKeys();
        const apiKey = apiKeys.OPENAI_API_KEY;
        
        if (!apiKey) {
            throw new Error('No API key available for transcription');
        }
        
        // Create form data
        const formData = new FormData();
        formData.append('file', audioBlob, 'recording.webm');
        formData.append('model', 'whisper-1');
        
        // Make API request
        const response = await fetch('https://api.openai.com/v1/audio/transcriptions', {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${apiKey}`
            },
            body: formData
        });
        
        if (!response.ok) {
            throw new Error(`Transcription failed: ${response.status}`);
        }
        
        const result = await response.json();
        console.log('[Voice-Client] Transcription result:', result);
        
        if (result.text) {
            // Insert transcribed text at cursor position
            insertTranscribedTextClient(result.text);
        }
        
    } catch (error) {
        console.error('[Voice-Client] Transcription error:', error);
        showError('Failed to transcribe audio. Please try again.');
    } finally {
        switchToNormalModeClient();
    }
}

// Insert transcribed text at cursor position
function insertTranscribedTextClient(text) {
    const textArea = document.getElementById('user-input-client');
    if (!textArea) return;
    
    // Get current text
    const currentText = textArea.value;
    const beforeCursor = currentText.substring(0, cursorPositionBeforeRecording);
    const afterCursor = currentText.substring(cursorPositionBeforeRecording);
    
    // Insert text at cursor position
    textArea.value = beforeCursor + text + afterCursor;
    
    // Set cursor position after inserted text
    const newCursorPosition = cursorPositionBeforeRecording + text.length;
    textArea.setSelectionRange(newCursorPosition, newCursorPosition);
    
    // Focus the textarea
    textArea.focus();
    
    // Trigger input event for any listeners
    textArea.dispatchEvent(new Event('input', { bubbles: true }));
    
    console.log('[Voice-Client] Transcribed text inserted');
}

// Add keyboard shortcuts for voice recording
document.addEventListener('keydown', (e) => {
    const voiceMode = document.getElementById('voice-recording-mode-client');
    
    if (voiceMode && voiceMode.parentNode) {
        // ESC to cancel recording
        if (e.key === 'Escape') {
            e.preventDefault();
            switchToNormalModeClient();
        }
        // Enter to accept recording
        else if (e.key === 'Enter') {
            e.preventDefault();
            stopVoiceRecordingClient();
        }
    }
});
// <<< END ADDED









