import { populateCodeCollection, exportCodeCollectionToText, runCodes, processAssumptionTabs, collapseGroupingsAndNavigateToFinancials, hideColumnsAndNavigate, handleInsertWorksheetsFromBase64, parseFormulaSCustomFormula } from './CodeCollection.js';
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
import { handleFollowUpConversation, handleInitialConversation, handleConversation, validationCorrection, formatCodeStringsWithGPT, consolidateAndDeduplicateTrainingData, getAICallsProcessedResponse } from './AIcalls.js';
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
    const userInput = document.getElementById('user-input').value.trim();
    
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

    // Add user message to chat
    appendMessage(userInput, true);
    
    // Clear input
    document.getElementById('user-input').value = '';

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
            // First message in conversation: query database and create enhanced prompt
            console.log("Database queries starting...");
            
            // Set progress message to indicate database querying
            progressMessageContent.textContent = 'Searching training database...';
            chatLog.scrollTop = chatLog.scrollHeight;

            // Create progress callback function
            const progressCallback = (message) => {
                console.log("Progress:", message);
                progressMessageContent.textContent = message;
                chatLog.scrollTop = chatLog.scrollHeight;
            };

            const dbResults = await structureDatabasequeries(userInput, progressCallback);
            console.log("Database queries completed");

            // Format database results into an enhanced prompt with consolidated training data and context
            const consolidatedData = consolidateAndDeduplicateTrainingData(dbResults);

            const enhancedPrompt = `Client Request: ${userInput}\n\n${consolidatedData}`;
            console.log("Enhanced prompt created");
            console.log("Enhanced prompt:", enhancedPrompt);

            console.log("Starting getAICallsProcessedResponse (initial - with prompt module determination)");
            
            // Update progress message
            progressMessageContent.textContent = 'Processing with AI (including prompt module analysis)...';
            chatLog.scrollTop = chatLog.scrollHeight;
            
            // Use the new function that includes prompt module determination
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

    // >>> ADDED: Hide header and activate conversation layout on first message
    const clientHeader = document.getElementById('client-mode-header');
    const clientChatContainer = document.getElementById('client-chat-container');
    const chatLogClient = document.getElementById('chat-log-client');
    
    if (clientHeader && !clientHeader.classList.contains('hidden')) {
        clientHeader.classList.add('hidden');
        // After transition, completely hide it
        setTimeout(() => {
            clientHeader.style.display = 'none';
        }, 300);
    }
    
    if (clientChatContainer && !clientChatContainer.classList.contains('conversation-active')) {
        clientChatContainer.classList.add('conversation-active');
    }
    
    if (chatLogClient) {
        chatLogClient.style.display = 'block';
    }
    // <<< END ADDED

    // Append user's message to the chat log
    appendMessage(userInput, true, 'chat-log-client', 'welcome-message-client');
    userInputElement.value = ''; // Clear the input field
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

        // Call OpenAI API with streaming
        // Assuming callOpenAI is available and handles API key internally,
        // and returns an async iterable for stream.
        // Adjust the model as needed, e.g., "gpt-3.5-turbo" or "gpt-4"
        const stream = await callOpenAI(messages, { stream: true, model: "gpt-3.5-turbo" });

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
    
    // Clear the input field
    document.getElementById('user-input').value = '';
    
    console.log("Chat reset completed");
}

// >>> ADDED: resetChat for Client Mode
function resetChatClient() {
    const chatLogClient = document.getElementById('chat-log-client');
    const welcomeMessageClient = document.getElementById('welcome-message-client');
    const clientModeHeader = document.getElementById('client-mode-header');
    const clientChatContainer = document.getElementById('client-chat-container');
    const userInputClient = document.getElementById('user-input-client');

    if (chatLogClient) {
        chatLogClient.innerHTML = '';
        if (welcomeMessageClient) {
            chatLogClient.appendChild(welcomeMessageClient);
            welcomeMessageClient.style.display = 'none';
        }
    }

    if (clientModeHeader) {
        clientModeHeader.classList.remove('hidden');
    }

    if (clientChatContainer) {
        clientChatContainer.classList.remove('conversation-active');
    }

    if (userInputClient) {
        userInputClient.value = '';
    }

    // Reset conversation history for client mode
    // Note: This might need to be adjusted based on how client mode history is managed
    
    // Reset the stored user inputs for client mode too
    firstUserInput = null;
    lastUserInput = null;
    
    // Clear persisted training data modal values
    persistedTrainingUserInput = null;
    persistedTrainingAiResponse = null;
    
    console.log("Client mode chat reset.");
}
// <<< END ADDED

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
        const worksheetsResponse = await fetch('https://localhost:3002/assets/Worksheets_4.3.25 v1.xlsx');
        if (!worksheetsResponse.ok) throw new Error(`[processModelCodesForPlanner] Worksheets_4.3.25 v1.xlsx load failed: ${worksheetsResponse.statusText}`);
        const wsArrayBuffer = await worksheetsResponse.arrayBuffer();
        const wsUint8Array = new Uint8Array(wsArrayBuffer);
        let wsBinaryString = '';
        for (let i = 0; i < wsUint8Array.length; i += 8192) {
            wsBinaryString += String.fromCharCode.apply(null, wsUint8Array.slice(i, Math.min(i + 8192, wsUint8Array.length)));
        }
        await handleInsertWorksheetsFromBase64(btoa(wsBinaryString));
        console.log("[processModelCodesForPlanner] Base sheets (Worksheets_4.3.25 v1.xlsx) inserted.");

        // 3. Insert codes.xlsx (as runCodes depends on it)
        console.log("[processModelCodesForPlanner] Inserting codes.xlsx...");
        const codesResponse = await fetch('https://localhost:3002/assets/codes.xlsx');
        if (!codesResponse.ok) throw new Error(`[processModelCodesForPlanner] codes.xlsx load failed: ${codesResponse.statusText}`);
        const codesArrayBuffer = await codesResponse.arrayBuffer();
        const codesUint8Array = new Uint8Array(codesArrayBuffer);
        let codesBinaryString = '';
        for (let i = 0; i < codesUint8Array.length; i += 8192) {
            codesBinaryString += String.fromCharCode.apply(null, codesUint8Array.slice(i, Math.min(i + 8192, codesUint8Array.length)));
        }
        await handleInsertWorksheetsFromBase64(btoa(codesBinaryString), ["Codes"]); 
        console.log("[processModelCodesForPlanner] codes.xlsx sheets inserted/updated.");
    
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
    let previousCodes = null;
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
            const worksheetsResponse = await fetch('https://localhost:3002/assets/Worksheets_4.3.25 v1.xlsx');
            if (!worksheetsResponse.ok) throw new Error(`Worksheets load failed: ${worksheetsResponse.statusText}`);
            const worksheetsArrayBuffer = await worksheetsResponse.arrayBuffer();
            const worksheetsUint8Array = new Uint8Array(worksheetsArrayBuffer);
            let worksheetsBinaryString = '';
            const chunkSize = 8192;
            for (let i = 0; i < worksheetsUint8Array.length; i += chunkSize) {
                const chunk = worksheetsUint8Array.slice(i, Math.min(i + chunkSize, worksheetsUint8Array.length));
                worksheetsBinaryString += String.fromCharCode.apply(null, chunk);
            }
            const worksheetsBase64String = btoa(worksheetsBinaryString);
            await handleInsertWorksheetsFromBase64(worksheetsBase64String);
            console.log("Base sheets inserted.");
        } else {
            console.log("[Run Codes] SUBSEQUENT PASS: Financials sheet found.");
            try {
                previousCodes = localStorage.getItem('previousRunCodeStrings');
            } catch (error) {
                 console.error("[Run Codes] Error loading previous codes for comparison:", error);
                 console.warn("[Run Codes] Could not load previous codes. Processing ALL current codes as fallback.");
                 previousCodes = null;
            }
            if (previousCodes !== null && previousCodes === codesToRun) {
                 console.log("[Run Codes] No change in code strings since last run. Nothing to process.");
                 try { localStorage.setItem('previousRunCodeStrings', codesToRun); } catch(e) { console.error("Err updating prev codes:", e); }
                 showMessage("No code changes to run.");
                 setButtonLoading(false);
                 return;
            }
            const currentTabs = getTabBlocks(codesToRun);
            const previousTabs = getTabBlocks(previousCodes || "");
            const previousTabMap = new Map(previousTabs.map(block => [block.tag, block.text]));
            let hasAnyChanges = false;
            const codeRegex = /<[^>]+>/g;
            for (const currentTab of currentTabs) {
                const currentTag = currentTab.tag;
                const currentText = currentTab.text;
                const previousText = previousTabMap.get(currentTag);
                if (previousText === undefined) {
                    const newTabCodes = currentText.match(codeRegex) || [];
                    if (newTabCodes.length > 0) {
                        allCodeContentToProcess += newTabCodes.join("\n") + "\n\n";
                        hasAnyChanges = true;
                    }
                } else {
                    const currentCodes = currentText.match(codeRegex) || [];
                    const previousCodesSet = new Set((previousText || "").match(codeRegex) || []);
                    let tabHasChanges = false;
                    let codesToAddForThisTab = "";
                    for (const currentCode of currentCodes) {
                        if (!previousCodesSet.has(currentCode)) {
                            codesToAddForThisTab += currentCode + "\n";
                            hasAnyChanges = true;
                            tabHasChanges = true;
                        }
                    }
                    if (tabHasChanges) {
                        allCodeContentToProcess += currentTag + "\n" + codesToAddForThisTab + "\n";
                    }
                }
            }
            if (hasAnyChanges) {
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
                    console.log("Code validation completed for new/modified tabs.");
                } else {
                    console.log("[Run Codes] Changes detected, but no code content found for validation in new/modified tabs.");
                }
                try {
                    const codesResponse = await fetch('https://localhost:3002/assets/codes.xlsx');
                    if (!codesResponse.ok) throw new Error(`codes.xlsx load failed: ${codesResponse.statusText}`);
                    const codesArrayBuffer = await codesResponse.arrayBuffer();
                    const codesUint8Array = new Uint8Array(codesArrayBuffer);
                    let codesBinaryString = '';
                    for (let i = 0; i < codesUint8Array.length; i += 8192) {
                        codesBinaryString += String.fromCharCode.apply(null, codesUint8Array.slice(i, Math.min(i + 8192, codesUint8Array.length)));
                    }
                    await handleInsertWorksheetsFromBase64(btoa(codesBinaryString), ["Codes"]);
                    console.log("codes.xlsx sheets inserted.");
                } catch (e) {
                    console.error("Failed to insert sheets from codes.xlsx:", e);
                    showError("Failed to insert necessary sheets from codes.xlsx. Aborting.");
                    setButtonLoading(false);
                    return;
                }
            } else {
                 console.log("[Run Codes] No changes identified in tabs compared to previous run. Nothing to insert or process.");
                 try { localStorage.setItem('previousRunCodeStrings', codesToRun); } catch(e) { console.error("Err updating prev codes:", e); }
                 showMessage("No code changes identified to run.");
                 setButtonLoading(false);
                 return;
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
        try {
            localStorage.setItem('previousRunCodeStrings', codesToRun);
        } catch (error) {
             console.error("[Run Codes] Failed to update previous run state:", error);
        }
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

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // Get references to the new elements
    const startupMenu = document.getElementById('startup-menu');
    const developerModeButton = document.getElementById('developer-mode-button');
    const clientModeButton = document.getElementById('client-mode-button');
    const appBody = document.getElementById('app-body'); // Already exists, ensure it's captured
    const clientModeView = document.getElementById('client-mode-view');

    // Functions to switch views
    function showDeveloperMode() {
      if (startupMenu) startupMenu.style.display = 'none';
      if (appBody) appBody.style.display = 'flex'; // Matches .ms-welcome__main display if it's flex
      if (clientModeView) clientModeView.style.display = 'none';
      console.log("Developer Mode activated");
    }

    function showClientMode() {
      if (startupMenu) startupMenu.style.display = 'none';
      if (appBody) appBody.style.display = 'none';
      if (clientModeView) clientModeView.style.display = 'flex'; // Matches .ms-welcome__main display
      console.log("Client Mode activated");
    }

    // >>> ADDED: Function to show startup menu
    function showStartupMenu() {
      if (startupMenu) startupMenu.style.display = 'flex';
      if (appBody) appBody.style.display = 'none';
      if (clientModeView) clientModeView.style.display = 'none';
      
      // Reset client mode state when going back to menu
      resetChatClient();
      
      console.log("Returned to startup menu");
    }
    // <<< END ADDED

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

    // >>> ADDED: Auto-resize for client textarea
    const userInputClientTextarea = document.getElementById('user-input-client');
    if (userInputClientTextarea) {
        userInputClientTextarea.addEventListener('input', function() {
            this.style.height = 'auto'; // Reset height to recalculate
            let newHeight = this.scrollHeight;
            const maxHeight = parseInt(window.getComputedStyle(this).maxHeight, 10);
            if (maxHeight && newHeight > maxHeight) {
                newHeight = maxHeight;
            }
            this.style.height = newHeight + 'px';
        });
    }
    // <<< END ADDED

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

    // >>> ADDED: Setup for Client Mode Chat Buttons
    const sendClientButton = document.getElementById('send-client');
    if (sendClientButton) sendClientButton.onclick = handleSendClient;

    const resetChatClientButton = document.getElementById('reset-chat-client');
    if (resetChatClientButton) resetChatClientButton.onclick = resetChatClient;

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

      document.getElementById("sideload-msg").style.display = "none";
      const appBody = document.getElementById('app-body');
      const clientModeView = document.getElementById('client-mode-view');

      if (startupMenu) startupMenu.style.display = "flex";
      if (appBody) appBody.style.display = "none";
      if (clientModeView) clientModeView.style.display = "none";
      // End Startup Menu Logic

    }).catch(error => {
        console.error("Error during initialization:", error);
        showError("Error during initialization: " + error.message);
    });

    document.getElementById("sideload-msg").style.display = "none";
    // document.getElementById("app-body").style.display = "block"; // Keep app-body hidden initially
    if (startupMenu) startupMenu.style.display = "flex"; // Show startup menu instead
    if (appBody) appBody.style.display = "none";
    if (clientModeView) clientModeView.style.display = "none";

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









