/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */
// Remove imports from Langchain to avoid ESM module issues
// Using direct fetch calls instead
// Add this test function
import { validateCodeStrings } from './Validation.js';
// >>> ADDED: Import CONFIG for URL management
import { CONFIG } from './config.js';
// Import the spreadsheet utilities
// import { handleInsertWorksheetsFromBase64 } from './SpreadsheetUtils.js';
// Import code collection functions
import { populateCodeCollection, exportCodeCollectionToText, runCodes, processAssumptionTabs, collapseGroupingsAndNavigateToFinancials, hideColumnsAndNavigate, handleInsertWorksheetsFromBase64 } from './CodeCollection.js';
// >>> ADDED: Import the new validation function
import { validateCodeStringsForRun, validateLogicWithRetry, validateFormatWithRetry } from './Validation.js';
// >>> ADDED: Import the tab string generator function
import { generateTabString } from './IndexWorksheet.js';
// >>> ADDED: Import AIModelPlanner functions
import { handleAIModelPlannerConversation, resetAIModelPlannerConversation, setAIModelPlannerOpenApiKey, plannerHandleSend, plannerHandleReset, plannerHandleWriteToExcel, plannerHandleInsertToEditor } from './AIModelPlanner.js';
// >>> ADDED: Import logic validation functions  
import { getLogicErrorsForPrompt, getLogicErrors, hasLogicErrors } from './CodeCollection.js';
// >>> ADDED: Import format validation functions  
import { getFormatErrorsForPrompt, getFormatErrors, hasFormatErrors } from './CodeCollection.js';
// >>> ADDED: Import cost tracking functions
import { trackAPICallCost, estimateTokens } from './CostTracker.js';
// Add the codeStrings variable with the specified content
// REMOVED hardcoded codeStrings variable




//*********Setup*********
// Start the timer
const startTime = performance.now();

//Debugging Toggle
const DEBUG = false;

// Variable to store loaded code strings
let loadedCodeStrings = "";

// Office-compatible notification functions
function showMessage(message) {
    const messageDiv = document.createElement('div');
    messageDiv.style.color = 'green';
    messageDiv.style.padding = '10px';
    messageDiv.style.margin = '10px';
    messageDiv.style.border = '1px solid green';
    messageDiv.style.borderRadius = '4px';
    messageDiv.textContent = message;
    
    const appBody = document.getElementById('app-body');
    if (appBody) {
        appBody.insertBefore(messageDiv, appBody.firstChild);
        
        // Remove the message after 5 seconds
        setTimeout(() => {
            messageDiv.remove();
        }, 5000);
    }
}

function showError(message) {
    const errorDiv = document.createElement('div');
    errorDiv.style.color = 'red';
    errorDiv.style.padding = '10px';
    errorDiv.style.margin = '10px';
    errorDiv.style.border = '1px solid red';
    errorDiv.style.borderRadius = '4px';
    errorDiv.textContent = `Error: ${message}`;
    
    const appBody = document.getElementById('app-body');
    if (appBody) {
        appBody.insertBefore(errorDiv, appBody.firstChild);
        
        // Remove the error message after 5 seconds
        setTimeout(() => {
            errorDiv.remove();
        }, 5000);
    }
}

// Variable to store the parsed code database
let codeDatabase = [];

// >>> ADDED: Variable to track validation pass number
let validationPassCounter = 0;

// >>> ADDED: Variable to store the original clean client prompt for LogicCheckerGPT
let originalClientPrompt = "";

// >>> ADDED: Function to reset validation pass counter (call at start of new processing)
export function resetValidationPassCounter() {
    validationPassCounter = 0;
}

// >>> ADDED: Variables for search/replace state <<<
// >>> REMOVED: Main search/replace state variables <<<
// let lastSearchTerm = '';
// let lastSearchIndex = -1; // Tracks the starting index of the last found match
// let searchResultIndices = []; // Stores indices of all matches for Replace All
// let currentHighlightIndex = -1; // Index within searchResultIndices for Find Next

/*
 * OPENAI API ARCHITECTURE:
 * 
 * This module supports two OpenAI API endpoints:
 * 
 * 1. Chat Completions API (default): /v1/chat/completions
 *    - Uses message arrays with roles (system, user, assistant)
 *    - Supports all OpenAI models including GPT-4, GPT-3.5, etc.
 *    - Function: callOpenAIChatCompletions()
 * 
 * 2. Responses API (optional): /v1/responses  
 *    - Uses single input string for single-turn prompts
 *    - Specifically designed for o3 model
 *    - Supports reasoning effort control and built-in tools
 *    - Function: callOpenAIResponses()
 * 
 * 3. Claude API (Anthropic): /v1/messages
 *    - Uses message arrays with separate system prompt and user messages
 *    - Supports Claude models (claude-opus-4-1-20250805, etc.)
 *    - Function: callClaudeAPI()
 *
 * 4. Unified Interface: callOpenAI()
 *    - Automatically chooses between the three APIs based on configuration
 *    - For main encoder calls: controlled by ENCODER_API_TYPE setting
 *    - For other calls: uses Chat Completions API by default
 *    - Can be overridden per-call with useResponsesAPI or useClaudeAPI options
 * 
 * Configuration:
 *    - setEncoderAPIType(ENCODER_API_TYPES.CHAT_COMPLETIONS|RESPONSES|CLAUDE): Set encoder API type
 *    - getEncoderAPIType(): Get current encoder API type
 *    - Backward compatibility functions also available
 */

// API keys storage - initialized by initializeAPIKeys
let INTERNAL_API_KEYS = {
  OPENAI_API_KEY: "",
  PINECONE_API_KEY: "",
  CLAUDE_API_KEY: ""
};

// Function to set API keys from outside this module
export function setAPIKeys(keys) {
  if (keys && typeof keys === 'object') {
    if (keys.OPENAI_API_KEY) {
      INTERNAL_API_KEYS.OPENAI_API_KEY = keys.OPENAI_API_KEY;
      console.log("AIcalls.js: OpenAI API key set externally");
    }
    if (keys.PINECONE_API_KEY) {
      INTERNAL_API_KEYS.PINECONE_API_KEY = keys.PINECONE_API_KEY;
      console.log("AIcalls.js: Pinecone API key set externally");
    }
    if (keys.CLAUDE_API_KEY) {
      INTERNAL_API_KEYS.CLAUDE_API_KEY = keys.CLAUDE_API_KEY;
      console.log("AIcalls.js: Claude API key set externally");
    }
  }
}

const srcPaths = [
  CONFIG.getAssetUrl('src/prompts/Encoder_System.txt'),
  CONFIG.getAssetUrl('src/prompts/Encoder_Main.txt'),
  CONFIG.getAssetUrl('src/prompts/Followup_System.txt'),
  CONFIG.getAssetUrl('src/prompts/Structure_System.txt'),
  CONFIG.getAssetUrl('src/prompts/Validation_System.txt'),
  CONFIG.getAssetUrl('src/prompts/Validation_Main.txt')
];

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

// Function to load API keys from environment variables
// Keys are loaded from .env file via webpack DefinePlugin at build time
export async function initializeAPIKeys() {
  try {
    console.log("Initializing API keys from AIcalls.js...");
    console.log("Environment check - NODE_ENV:", process.env.NODE_ENV);
    console.log("Environment check - typeof process.env.OPENAI_API_KEY:", typeof process.env.OPENAI_API_KEY);
//fdf
    // Use keys from environment variables via webpack DefinePlugin
    if (process.env.OPENAI_API_KEY && process.env.OPENAI_API_KEY !== 'undefined') {
        INTERNAL_API_KEYS.OPENAI_API_KEY = process.env.OPENAI_API_KEY;
        setAIModelPlannerOpenApiKey(process.env.OPENAI_API_KEY);
        console.log("OpenAI API key loaded from environment variables and set for AI Model Planner");
    } else {
         console.warn("OpenAI API key not found in environment variables.");
    }

    if (process.env.PINECONE_API_KEY && process.env.PINECONE_API_KEY !== 'undefined') {
        INTERNAL_API_KEYS.PINECONE_API_KEY = process.env.PINECONE_API_KEY;
        console.log("Pinecone API key loaded from environment variables");
    } else {
         console.warn("Pinecone API key not found in environment variables.");
    }

    if (process.env.CLAUDE_API_KEY && process.env.CLAUDE_API_KEY !== 'undefined') {
        INTERNAL_API_KEYS.CLAUDE_API_KEY = process.env.CLAUDE_API_KEY;
        console.log("Claude API key loaded from environment variables");
    } else {
         console.warn("Claude API key not found in environment variables.");
    }
    
    // Add debug logging with secure masking of keys
    console.log("Loaded API Keys (AIcalls.js):");
    console.log("  OPENAI_API_KEY:", INTERNAL_API_KEYS.OPENAI_API_KEY ?
      `${INTERNAL_API_KEYS.OPENAI_API_KEY.substring(0, 3)}...${INTERNAL_API_KEYS.OPENAI_API_KEY.substring(INTERNAL_API_KEYS.OPENAI_API_KEY.length - 3)}` :
      "Not found");
    console.log("  PINECONE_API_KEY:", INTERNAL_API_KEYS.PINECONE_API_KEY ?
      `${INTERNAL_API_KEYS.PINECONE_API_KEY.substring(0, 3)}...${INTERNAL_API_KEYS.PINECONE_API_KEY.substring(INTERNAL_API_KEYS.PINECONE_API_KEY.length - 3)}` :
      "Not found");
    console.log("  CLAUDE_API_KEY:", INTERNAL_API_KEYS.CLAUDE_API_KEY ?
      `${INTERNAL_API_KEYS.CLAUDE_API_KEY.substring(0, 3)}...${INTERNAL_API_KEYS.CLAUDE_API_KEY.substring(INTERNAL_API_KEYS.CLAUDE_API_KEY.length - 3)}` :
      "Not found");

    const keysFound = !!(INTERNAL_API_KEYS.OPENAI_API_KEY && INTERNAL_API_KEYS.PINECONE_API_KEY && INTERNAL_API_KEYS.CLAUDE_API_KEY);
    console.log("API Keys Initialized:", keysFound);
    // Return a copy to prevent external modification of the internal state
    return { ...INTERNAL_API_KEYS };

  } catch (error) {
    console.error("Error initializing API keys:", error);
    // Return empty keys on error
    return { OPENAI_API_KEY: "", PINECONE_API_KEY: "", CLAUDE_API_KEY: "" };
  }
}

// Update Pinecone configuration to handle multiple indexes
const PINECONE_ENVIRONMENT = "gcp-starter";

// Define configurations for each index
const PINECONE_INDEXES = {
    // codes: {
    //     name: "codes",
    //     apiEndpoint: "https://codes-hz34tmv.svc.aped-4627-b74a.pinecone.io"
    // },
    call2trainingdata: {
        name: "call2trainingdata",
        apiEndpoint: "https://call2trainingdata-zmg9zog.svc.aped-4627-b74a.pinecone.io"
    },
    // >>> COMMENTED OUT: Context database indexes
    // call2context: {
    //     name: "call2context",
    //     apiEndpoint: "https://call2context-zmg9zog.svc.aped-4627-b74a.pinecone.io"
    // },
    // call1context: {
    //     name: "call1context",
    //     apiEndpoint: "https://call1context-zmg9zog.svc.aped-4627-b74a.pinecone.io"
    // },


};

//Models
const GPT4O_MINI = "gpt-4o-mini"
const GPT4O = "gpt-4o"
const GPT41 = "gpt-4.1"
const GPT45_TURBO = "gpt-4.5-turbo"
const GPT35_TURBO = "gpt-3.5-turbo"
const GPT_O3mini = "o3-mini"
const GPT4_TURBO = "gpt-4-turbo"
const GPTO3 = "gpt-o3"
const GPT_O3 = "o3"  // Added for responses API
const GPTFT1 =  "ft:gpt-4.1-2025-04-14:personal:jun25gpt4-1:BeyDTNt1"

// API Configuration - Encoder API Type Selection
const ENCODER_API_TYPES = {
  CHAT_COMPLETIONS: 'chat_completions',
  RESPONSES: 'responses',
  CLAUDE: 'claude'
};

let ENCODER_API_TYPE = ENCODER_API_TYPES.CLAUDE; // Default to chat completions API

// Functions to control API configuration
export function setEncoderAPIType(apiType) {
  if (Object.values(ENCODER_API_TYPES).includes(apiType)) {
    ENCODER_API_TYPE = apiType;
    console.log(`[setEncoderAPIType] Encoder API type set to: ${apiType}`);
  } else {
    console.error(`[setEncoderAPIType] Invalid API type: ${apiType}. Valid options: ${Object.values(ENCODER_API_TYPES).join(', ')}`);
  }
}

export function getEncoderAPIType() {
  return ENCODER_API_TYPE;
}

// Backward compatibility functions
export function setUseResponsesAPIForEncoder(useResponsesAPI) {
  ENCODER_API_TYPE = useResponsesAPI ? ENCODER_API_TYPES.RESPONSES : ENCODER_API_TYPES.CHAT_COMPLETIONS;
  console.log(`[setUseResponsesAPIForEncoder] Encoder API type set to: ${ENCODER_API_TYPE} (backward compatibility)`);
}

export function getUseResponsesAPIForEncoder() {
  return ENCODER_API_TYPE === ENCODER_API_TYPES.RESPONSES;
}

export function setUseClaudeAPIForEncoder(useClaudeAPI) {
  ENCODER_API_TYPE = useClaudeAPI ? ENCODER_API_TYPES.CLAUDE : ENCODER_API_TYPES.CHAT_COMPLETIONS;
  console.log(`[setUseClaudeAPIForEncoder] Encoder API type set to: ${ENCODER_API_TYPE}`);
}

export function getUseClaudeAPIForEncoder() {
  return ENCODER_API_TYPE === ENCODER_API_TYPES.CLAUDE;
}

// Export the constants for external use
export { ENCODER_API_TYPES };

// Helper function to test responses API with a simple call
export async function testResponsesAPI(input, options = {}) {
  console.log("[testResponsesAPI] Testing OpenAI Responses API...");
  
  const testOptions = {
    model: GPT_O3,
    reasoning: { effort: "medium" },
    stream: false,
    caller: "testResponsesAPI",
    ...options
  };
  
  try {
    let response = "";
    for await (const contentPart of callOpenAIResponses(input, testOptions)) {
      response += contentPart;
    }
    
    console.log("[testResponsesAPI] Test completed successfully");
    console.log("[testResponsesAPI] Response length:", response.length);
    
    return response;
  } catch (error) {
    console.error("[testResponsesAPI] Test failed:", error);
    throw error;
  }
}

// Helper function to test Claude API with a simple call
export async function testClaudeAPI(input, options = {}) {
  console.log("[testClaudeAPI] Testing Claude API...");
  
  const testMessages = [
    { role: "system", content: "You are a helpful assistant." },
    { role: "user", content: input }
  ];
  
  const testOptions = {
    model: "claude-opus-4-1-20250805",
    temperature: 1,
    stream: false,
    caller: "testClaudeAPI",
    ...options
  };
  
  try {
    let response = "";
    for await (const contentPart of callClaudeAPI(testMessages, testOptions)) {
      response += contentPart;
    }
    
    console.log("[testClaudeAPI] Test completed successfully");
    console.log("[testClaudeAPI] Response length:", response.length);
    
    return response;
  } catch (error) {
    console.error("[testClaudeAPI] Test failed:", error);
    throw error;
  }
}

// Conversation history storage
let conversationHistory = [];

// Functions to save and load conversation history (use localStorage directly)
export function saveConversationHistory(history) {
    try {
        localStorage.setItem('conversationHistory', JSON.stringify(history));
        if (DEBUG) console.log('Conversation history saved to localStorage');
    } catch (error) {
        console.error("Error saving conversation history:", error);
    }
}

export function loadConversationHistory() {
    try {
        const history = localStorage.getItem('conversationHistory');
        if (history) {
            if (DEBUG) console.log('Loaded conversation history from localStorage');
            const parsedHistory = JSON.parse(history);

            if (!Array.isArray(parsedHistory)) {
                console.error("Invalid history format, expected array");
                return [];
            }

            return parsedHistory;
        }
        if (DEBUG) console.log("No conversation history found in localStorage");
        return [];
    } catch (error) {
        console.error("Error loading conversation history:", error);
        return [];
    }
}

// Direct OpenAI Chat Completions API call function
export async function* callOpenAIChatCompletions(messages, options = {}) {
  const { model = GPT_O3, temperature = 0.7, stream = false, caller = "Unknown" } = options;

  // >>> ADDED: Cost tracking variables
  const startTime = Date.now();
  let inputTokens = 0;
  let outputTokens = 0;
  let fullResponse = "";

  try {
    console.log(`Calling OpenAI API with model: ${model}, stream: ${stream}`);
    
    // >>> ADDED: Calculate input tokens for cost tracking
    inputTokens = messages.reduce((total, message) => {
      return total + estimateTokens(message.content || "");
    }, 0);
    
    // >>> ADDED: Comprehensive logging of all messages sent to OpenAI with descriptive names
    // Map system prompts to descriptive call names
    let callName = "Unknown";
    let validationErrors = "";
    
    if (caller.includes("Structure_System")) {
      callName = "Prompt Breakup";
    } else if (caller.includes("Encoder_System") || caller.includes("Followup_System")) {
      callName = "Main Encoder";
    } else if (caller.includes("LogicCheckerGPT")) {
      callName = "Logic Checker GPT";
    } else if (caller.includes("FormatGPT")) {
      callName = "Format GPT";
    } else if (caller.includes("Validation_System")) {
      // Parse validation pass number and errors from caller
      const passMatch = caller.match(/PASS_(\d+)/);
      const errorsMatch = caller.match(/ERRORS:(.+)$/);
      const passNumber = passMatch ? passMatch[1] : "1";
      callName = `Validation Pass ${passNumber}`;
      validationErrors = errorsMatch ? errorsMatch[1] : "";
    }
    
    console.log("\n╔════════════════════════════════════════════════════════════════╗");
    console.log(`║              OPENAI API CALL - ${callName.padEnd(25)} ║`);
    console.log("╠════════════════════════════════════════════════════════════════╣");
    console.log(`║ CALLER: ${caller.padEnd(54)} ║`);
    console.log(`║ FILE: AIcalls.js                                               ║`);
    console.log(`║ FUNCTION: callOpenAI()                                         ║`);
    
    // Display validation errors if this is a validation call
    if (validationErrors) {
      console.log("╠════════════════════════════════════════════════════════════════╣");
      console.log("║ VALIDATION ERRORS BEING CHECKED:                               ║");
      console.log("╠════════════════════════════════════════════════════════════════╣");
      const errorLines = validationErrors.split('\n').slice(0, 5); // Show first 5 lines
      errorLines.forEach(line => {
        const truncatedLine = line.substring(0, 62);
        console.log(`║ ${truncatedLine.padEnd(62)} ║`);
      });
      if (validationErrors.split('\n').length > 5) {
        console.log("║ ... (additional errors truncated)                              ║");
      }
    }
    
    console.log("╚════════════════════════════════════════════════════════════════╝");
    console.log(`Model: ${model}`);
    console.log(`Temperature: ${temperature}`);
    console.log(`Stream: ${stream}`);
    console.log(`Total Messages: ${messages.length}`);
    console.log("────────────────────────────────────────────────────────────────");
    
    messages.forEach((message, index) => {
      console.log(`\n[Message ${index + 1}] Role: ${message.role.toUpperCase()}`);
      console.log("────────────────────────────────────────────────────────────────");
      console.log(message.content);
      console.log("────────────────────────────────────────────────────────────────");
    });
    
    console.log("\n╔════════════════════════════════════════════════════════════════╗");
    console.log(`║              END OF ${callName.toUpperCase()} CALL${' '.repeat(Math.max(0, 25 - callName.length))}║`);
    console.log("╚════════════════════════════════════════════════════════════════╝\n");
    // <<< END ADDED

    if (!INTERNAL_API_KEYS.OPENAI_API_KEY) {
      throw new Error("OpenAI API key not found. Please check your API keys.");
    }

    // Filter out messages with empty content
    const validMessages = messages.filter(message => {
      if (!message.content || message.content.trim() === "") {
        console.log(`⚠️ Skipping empty ${message.role} message in OpenAI API call`);
        return false;
      }
      return true;
    });

    // Validate that we have at least one valid message
    if (validMessages.length === 0) {
      throw new Error("No valid messages found for OpenAI API call. All messages were empty or invalid.");
    }

    const body = {
      model: model,
      messages: validMessages
    };

    // Only include temperature for models that support it
    // o3-mini and gpt-o3 models don't support temperature parameter
    const modelsWithoutTemperature = ['o3-mini', 'gpt-o3'];
    if (!modelsWithoutTemperature.includes(model)) {
      body.temperature = temperature;
    }

    if (stream) {
      body.stream = true;
    }

    const response = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${INTERNAL_API_KEYS.OPENAI_API_KEY}`
      },
      body: JSON.stringify(body)
    });

    if (!response.ok) {
      const errorData = await response.json().catch(() => ({ message: "Failed to parse error JSON." }));
      console.error("OpenAI API error response:", errorData);
      throw new Error(`OpenAI API error: ${response.status} ${response.statusText} - ${errorData.message || JSON.stringify(errorData)}`);
    }

    if (stream) {
      console.log("OpenAI API response received (stream)");
      const reader = response.body.getReader();
      const decoder = new TextDecoder("utf-8");

      while (true) {
        const { done, value } = await reader.read();
        if (done) {
          console.log("Stream finished.");
          
          // >>> ADDED: Track cost for streaming response
          outputTokens = estimateTokens(fullResponse);
          const duration = Date.now() - startTime;
          
          trackAPICallCost({
            provider: 'openai',
            model: model,
            inputTokens: inputTokens,
            outputTokens: outputTokens,
            caller: caller,
            timestamp: startTime,
            duration: duration,
            success: true
          });
          
          break;
        }
        const chunk = decoder.decode(value);
        const lines = chunk.split("\n");
        const parsedLines = lines
          .map((line) => line.replace(/^data: /, "").trim()) // Remove SSE "data: " prefix
          .filter((line) => line !== "" && line !== "[DONE]") // Filter empty lines and [DONE] message
          .map((line) => {
            try {
              return JSON.parse(line);
            } catch (e) {
              console.warn("Could not parse JSON line from stream:", line, e);
              return null; // Or handle error appropriately
            }
          })
          .filter(line => line !== null);

        for (const parsedLine of parsedLines) {
          // >>> ADDED: Accumulate response text for cost tracking
          if (parsedLine.choices && parsedLine.choices[0]?.delta?.content) {
            fullResponse += parsedLine.choices[0].delta.content;
          }
          
          yield parsedLine;
        }
      }
    } else {
      const data = await response.json();
      console.log("OpenAI API response received (non-stream)");
      
      // >>> ADDED: Track cost for non-streaming response
      const responseContent = data.choices[0].message.content;
      outputTokens = estimateTokens(responseContent);
      const duration = Date.now() - startTime;
      
      // Use actual token counts from API response if available
      if (data.usage) {
        inputTokens = data.usage.prompt_tokens || inputTokens;
        outputTokens = data.usage.completion_tokens || outputTokens;
      }
      
      trackAPICallCost({
        provider: 'openai',
        model: model,
        inputTokens: inputTokens,
        outputTokens: outputTokens,
        caller: caller,
        timestamp: startTime,
        duration: duration,
        success: true
      });
      
      // For non-streaming, to maintain compatibility with handleSendClient's expectation of an iterable, 
      // we yield a single object that mimics the structure of a stream chunk if needed, 
      // or simply return the content if the caller adapts.
      // For now, let's assume the non-streaming path is not used by handleSendClient directly.
      // Returning the content directly as before for other potential callers.
      // If callOpenAI is *only* called by handleSendClient, this else block might need to yield as well.
      // However, processPrompt calls callOpenAI without expecting a stream.
      // If callOpenAI is *only* called by handleSendClient, this else block might need to yield as well.
      // However, processPrompt calls callOpenAI without expecting a stream.
      yield responseContent; // Yield the content string once
      return; // Explicitly end the generator
    }

  } catch (error) {
    console.error("Error calling OpenAI API:", error);
    
    // >>> ADDED: Track failed API call for cost analysis
    const duration = Date.now() - startTime;
    trackAPICallCost({
      provider: 'openai',
      model: model,
      inputTokens: inputTokens,
      outputTokens: 0, // No output on error
      caller: caller,
      timestamp: startTime,
      duration: duration,
      success: false
    });
    
    // If it's a stream, we can't return, but the error will propagate.
    // If not a stream, re-throw as before.
    if (!stream) throw error;
    // For a stream, the error breaks the generator. Consider yielding an error object if preferred.
    // For now, just log and let the generator terminate.
    // yield { error: error.message }; // Optional: yield an error object
  }
}

// Direct OpenAI Responses API call function (for o3 model)
export async function* callOpenAIResponses(input, options = {}) {
  const { model = GPT_O3, reasoning = { effort: "medium" }, stream = false, caller = "Unknown", tools = [] } = options;

  // >>> ADDED: Cost tracking variables
  const startTime = Date.now();
  let inputTokens = estimateTokens(input);
  let outputTokens = 0;
  let fullResponse = "";

  try {
    console.log(`Calling OpenAI Responses API with model: ${model}, stream: ${stream}`);
    
    // >>> ADDED: Comprehensive logging for Responses API calls
    let callName = "Unknown";
    
    if (caller.includes("Structure_System")) {
      callName = "Prompt Breakup";
    } else if (caller.includes("Encoder_System") || caller.includes("Followup_System")) {
      callName = "Main Encoder";
    } else if (caller.includes("LogicCheckerGPT")) {
      callName = "Logic Checker GPT";
    } else if (caller.includes("FormatGPT")) {
      callName = "Format GPT";
    } else if (caller.includes("Validation_System")) {
      const passMatch = caller.match(/PASS_(\d+)/);
      const passNumber = passMatch ? passMatch[1] : "1";
      callName = `Validation Pass ${passNumber}`;
    }
    
    console.log("\n╔════════════════════════════════════════════════════════════════╗");
    console.log(`║           OPENAI RESPONSES API CALL - ${callName.padEnd(25)} ║`);
    console.log("╠════════════════════════════════════════════════════════════════╣");
    console.log(`║ CALLER: ${caller.padEnd(54)} ║`);
    console.log(`║ FILE: AIcalls.js                                               ║`);
    console.log(`║ FUNCTION: callOpenAIResponses()                                ║`);
    console.log("╚════════════════════════════════════════════════════════════════╝");
    console.log(`Model: ${model}`);
    console.log(`Reasoning Effort: ${reasoning.effort}`);
    console.log(`Stream: ${stream}`);
    console.log(`Tools: ${tools.length > 0 ? tools.length : 'None'}`);
    console.log("────────────────────────────────────────────────────────────────");
    
    console.log("\n[Input]");
    console.log("────────────────────────────────────────────────────────────────");
    console.log(input);
    console.log("────────────────────────────────────────────────────────────────");
    
    console.log("\n╔════════════════════════════════════════════════════════════════╗");
    console.log(`║              END OF ${callName.toUpperCase()} CALL${' '.repeat(Math.max(0, 25 - callName.length))}║`);
    console.log("╚════════════════════════════════════════════════════════════════╝\n");

    if (!INTERNAL_API_KEYS.OPENAI_API_KEY) {
      throw new Error("OpenAI API key not found. Please check your API keys.");
    }

    // Validate input is not empty
    if (!input || input.trim() === "") {
      throw new Error("No valid input found for OpenAI Responses API call. Input was empty or invalid.");
    }

    const body = {
      model: model,
      input: input,
      reasoning: reasoning
    };

    if (tools.length > 0) {
      body.tools = tools;
    }

    if (stream) {
      body.stream = true;
    }

    const response = await fetch('https://api.openai.com/v1/responses', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${INTERNAL_API_KEYS.OPENAI_API_KEY}`
      },
      body: JSON.stringify(body)
    });

    if (!response.ok) {
      const errorData = await response.json().catch(() => ({ message: "Failed to parse error JSON." }));
      console.error("OpenAI Responses API error response:", errorData);
      throw new Error(`OpenAI Responses API error: ${response.status} ${response.statusText} - ${errorData.message || JSON.stringify(errorData)}`);
    }

    if (stream) {
      console.log("OpenAI Responses API response received (stream)");
      const reader = response.body.getReader();
      const decoder = new TextDecoder("utf-8");

      while (true) {
        const { done, value } = await reader.read();
        if (done) {
          console.log("Stream finished.");
          
          // >>> ADDED: Track cost for streaming response
          outputTokens = estimateTokens(fullResponse);
          const duration = Date.now() - startTime;
          
          trackAPICallCost({
            provider: 'openai',
            model: model,
            inputTokens: inputTokens,
            outputTokens: outputTokens,
            caller: caller,
            timestamp: startTime,
            duration: duration,
            success: true
          });
          
          break;
        }
        const chunk = decoder.decode(value);
        const lines = chunk.split("\n");
        const parsedLines = lines
          .map((line) => line.replace(/^data: /, "").trim())
          .filter((line) => line !== "" && line !== "[DONE]")
          .map((line) => {
            try {
              return JSON.parse(line);
            } catch (e) {
              console.warn("Could not parse JSON line from stream:", line, e);
              return null;
            }
          })
          .filter(line => line !== null);

        for (const parsedLine of parsedLines) {
          // >>> ADDED: Accumulate response text for cost tracking
          if (parsedLine.choices && parsedLine.choices[0]?.delta?.content) {
            fullResponse += parsedLine.choices[0].delta.content;
          } else if (parsedLine.output) {
            fullResponse += parsedLine.output;
          }
          
          yield parsedLine;
        }
      }
    } else {
      const data = await response.json();
      console.log("OpenAI Responses API response received (non-stream)");
      
      // >>> ADDED: Track cost for non-streaming response
      let responseContent = "";
      
      // For responses API, the response structure is different
      // Extract content from the response
      if (data.choices && data.choices[0] && data.choices[0].message) {
        responseContent = data.choices[0].message.content;
      } else if (data.output) {
        responseContent = data.output;
      } else {
        responseContent = JSON.stringify(data); // Fallback to full response
      }
      
      outputTokens = estimateTokens(responseContent);
      const duration = Date.now() - startTime;
      
      // Use actual token counts from API response if available
      if (data.usage) {
        inputTokens = data.usage.prompt_tokens || inputTokens;
        outputTokens = data.usage.completion_tokens || outputTokens;
      }
      
      trackAPICallCost({
        provider: 'openai',
        model: model,
        inputTokens: inputTokens,
        outputTokens: outputTokens,
        caller: caller,
        timestamp: startTime,
        duration: duration,
        success: true
      });
      
      yield responseContent;
      return;
    }

  } catch (error) {
    console.error("Error calling OpenAI Responses API:", error);
    
    // >>> ADDED: Track failed API call for cost analysis
    const duration = Date.now() - startTime;
    trackAPICallCost({
      provider: 'openai',
      model: model,
      inputTokens: inputTokens,
      outputTokens: 0, // No output on error
      caller: caller,
      timestamp: startTime,
      duration: duration,
      success: false
    });
    
    if (!stream) throw error;
  }
}

// Direct Claude API call function (for Anthropic models)
export async function* callClaudeAPI(messages, options = {}) {
  const { model = "claude-opus-4-1-20250805", temperature = 1, stream = false, caller = "Unknown" } = options;

  // >>> ADDED: Cost tracking variables
  const startTime = Date.now();
  let inputTokens = 0;
  let outputTokens = 0;
  let fullResponse = "";

  try {
    console.log(`Calling Claude API with model: ${model}, stream: ${stream}`);
    
    // >>> ADDED: Calculate input tokens for cost tracking
    inputTokens = messages.reduce((total, message) => {
      return total + estimateTokens(message.content || "");
    }, 0);
    
    // >>> ADDED: Comprehensive logging for Claude API calls
    let callName = "Unknown";
    
    if (caller.includes("Structure_System")) {
      callName = "Prompt Breakup";
    } else if (caller.includes("Encoder_System") || caller.includes("Encoder_Main")) {
      callName = "Main Encoder";
    } else if (caller.includes("LogicCheckerGPT")) {
      callName = "Logic Checker GPT";
    } else if (caller.includes("FormatGPT")) {
      callName = "Format GPT";
    } else if (caller.includes("Validation_System")) {
      const passMatch = caller.match(/PASS_(\d+)/);
      const passNumber = passMatch ? passMatch[1] : "1";
      callName = `Validation Pass ${passNumber}`;
    }
    
    console.log("\n╔════════════════════════════════════════════════════════════════╗");
    console.log(`║              CLAUDE API CALL - ${callName.padEnd(25)} ║`);
    console.log("╠════════════════════════════════════════════════════════════════╣");
    console.log(`║ CALLER: ${caller.padEnd(54)} ║`);
    console.log(`║ FILE: AIcalls.js                                               ║`);
    console.log(`║ FUNCTION: callClaudeAPI()                                      ║`);
    console.log("╚════════════════════════════════════════════════════════════════╝");
    console.log(`Model: ${model}`);
    console.log(`Temperature: ${temperature}`);
    console.log(`Stream: ${stream}`);
    console.log(`Total Messages: ${messages.length}`);
    console.log("────────────────────────────────────────────────────────────────");
    
    messages.forEach((message, index) => {
      console.log(`\n[Message ${index + 1}] Role: ${message.role.toUpperCase()}`);
      console.log("────────────────────────────────────────────────────────────────");
      console.log(message.content);
      console.log("────────────────────────────────────────────────────────────────");
    });
    
    console.log("\n╔════════════════════════════════════════════════════════════════╗");
    console.log(`║              END OF ${callName.toUpperCase()} CALL${' '.repeat(Math.max(0, 25 - callName.length))}║`);
    console.log("╚════════════════════════════════════════════════════════════════╝\n");

    if (!INTERNAL_API_KEYS.CLAUDE_API_KEY) {
      throw new Error("Claude API key not found. Please check your API keys.");
    }

    // Extract system message and user messages for Claude API format
    let systemMessage = "";
    const userMessages = [];
    
    for (const message of messages) {
      if (message.role === "system") {
        systemMessage = message.content;
      } else if (message.role === "user" || message.role === "assistant") {
        // Only add messages with non-empty content
        if (message.content && message.content.trim() !== "") {
          userMessages.push({
            role: message.role,
            content: [
              {
                type: "text",
                text: message.content
              }
            ]
          });
        } else {
          console.log(`⚠️ Skipping empty ${message.role} message in Claude API call`);
        }
      }
    }

    // Validate that we have at least one valid message
    if (userMessages.length === 0) {
      throw new Error("No valid messages found for Claude API call. All messages were empty or invalid.");
    }

    const body = {
      model: model,
      max_tokens: 20000,
      temperature: temperature,
      system: systemMessage,
      messages: userMessages
    };

    if (stream) {
      body.stream = true;
    }

    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': INTERNAL_API_KEYS.CLAUDE_API_KEY,
        'anthropic-version': '2023-06-01',
        'anthropic-dangerous-direct-browser-access': 'true'
      },
      body: JSON.stringify(body)
    });

    if (!response.ok) {
      const errorData = await response.json().catch(() => ({ message: "Failed to parse error JSON." }));
      console.error("Claude API error response:", errorData);
      throw new Error(`Claude API error: ${response.status} ${response.statusText} - ${errorData.message || JSON.stringify(errorData)}`);
    }

    if (stream) {
      console.log("Claude API response received (stream)");
      const reader = response.body.getReader();
      const decoder = new TextDecoder("utf-8");

      while (true) {
        const { done, value } = await reader.read();
        if (done) {
          console.log("Stream finished.");
          
          // >>> ADDED: Track cost for Claude streaming response
          outputTokens = estimateTokens(fullResponse);
          const duration = Date.now() - startTime;
          
          trackAPICallCost({
            provider: 'claude',
            model: model,
            inputTokens: inputTokens,
            outputTokens: outputTokens,
            caller: caller,
            timestamp: startTime,
            duration: duration,
            success: true
          });
          
          break;
        }
        const chunk = decoder.decode(value);
        const lines = chunk.split("\n");
        
        // Parse SSE format correctly - only process 'data:' lines as JSON
        const parsedLines = [];
        for (const line of lines) {
          const trimmedLine = line.trim();
          
          // Skip empty lines, comments, and non-data SSE lines
          if (trimmedLine === "" || trimmedLine.startsWith(":")) {
            continue;
          }
          
          // Handle SSE data lines
          if (trimmedLine.startsWith("data: ")) {
            const jsonData = trimmedLine.substring(6); // Remove "data: " prefix
            
            // Skip [DONE] markers
            if (jsonData === "[DONE]") {
              continue;
            }
            
            try {
              const parsed = JSON.parse(jsonData);
              parsedLines.push(parsed);
            } catch (e) {
              console.warn("Could not parse JSON data from SSE stream:", jsonData, e);
            }
          }
          // Handle other SSE lines (event:, id:, retry:) - log for debugging but don't process
          else if (trimmedLine.startsWith("event: ") || trimmedLine.startsWith("id: ") || trimmedLine.startsWith("retry: ")) {
            // These are valid SSE lines, just not data we need to process
            continue;
          }
          // Log unexpected line format for debugging
          else if (trimmedLine !== "") {
            console.debug("Unexpected SSE line format:", trimmedLine);
          }
        }

        for (const parsedLine of parsedLines) {
          // >>> ADDED: Accumulate response text for cost tracking
          // Handle Claude SSE format: content_block_delta events with text_delta
          if (parsedLine.type === "content_block_delta" && 
              parsedLine.delta && 
              parsedLine.delta.type === "text_delta" && 
              parsedLine.delta.text) {
            fullResponse += parsedLine.delta.text;
          }
          // Also handle non-streaming content blocks
          else if (parsedLine.content_block && parsedLine.content_block.text) {
            fullResponse += parsedLine.content_block.text;
          }
          
          yield parsedLine;
        }
      }
    } else {
      const data = await response.json();
      console.log("Claude API response received (non-stream)");
      
      // Extract content from Claude API response
      let responseContent = "";
      if (data.content && Array.isArray(data.content) && data.content.length > 0) {
        responseContent = data.content
          .filter(item => item.type === "text")
          .map(item => item.text)
          .join("");
      } else {
        responseContent = JSON.stringify(data); // Fallback to full response
      }
      
      // >>> ADDED: Track cost for Claude non-streaming response
      outputTokens = estimateTokens(responseContent);
      const duration = Date.now() - startTime;
      
      // Use actual token counts from API response if available
      if (data.usage) {
        inputTokens = data.usage.input_tokens || inputTokens;
        outputTokens = data.usage.output_tokens || outputTokens;
      }
      
      trackAPICallCost({
        provider: 'claude',
        model: model,
        inputTokens: inputTokens,
        outputTokens: outputTokens,
        caller: caller,
        timestamp: startTime,
        duration: duration,
        success: true
      });
      
      yield responseContent;
      return;
    }

  } catch (error) {
    console.error("Error calling Claude API:", error);
    
    // >>> ADDED: Track failed Claude API call for cost analysis
    const duration = Date.now() - startTime;
    trackAPICallCost({
      provider: 'claude',
      model: model,
      inputTokens: inputTokens,
      outputTokens: 0, // No output on error
      caller: caller,
      timestamp: startTime,
      duration: duration,
      success: false
    });
    
    if (!stream) throw error;
  }
}

// Unified API call function that chooses between Chat Completions, Responses, and Claude APIs
export async function* callOpenAI(messagesOrInput, options = {}) {
  const { useResponsesAPI = false, useClaudeAPI = false, model = GPT_O3, caller = "Unknown" } = options;
  
  // Determine if this is a main encoder call and should use the configured encoder API
  const isMainEncoderCall = caller.includes("Encoder_System") || caller.includes("Encoder_Main") || 
                           caller.includes("processPrompt() - Encoder_System") ||
                           caller.includes("processPrompt() - Encoder_Main");
  
  let apiTypeToUse = null;
  
  if (isMainEncoderCall) {
    // Use the configured encoder API type for main encoder calls
    apiTypeToUse = ENCODER_API_TYPE;
  } else {
    // For non-encoder calls, use explicit options or default to chat completions
    if (useClaudeAPI) {
      apiTypeToUse = ENCODER_API_TYPES.CLAUDE;
    } else if (useResponsesAPI) {
      apiTypeToUse = ENCODER_API_TYPES.RESPONSES;
    } else {
      apiTypeToUse = ENCODER_API_TYPES.CHAT_COMPLETIONS;
    }
  }
  
  // Ensure we have messages array format for all APIs
  let messages = messagesOrInput;
  if (typeof messagesOrInput === 'string') {
    // Convert string to messages array
    messages = [{ role: "user", content: messagesOrInput }];
  }
  
  if (apiTypeToUse === ENCODER_API_TYPES.CLAUDE) {
    // Use Claude API
    console.log(`[callOpenAI] Using Claude API for ${caller} with model ${model}`);
    yield* callClaudeAPI(messages, options);
  } else if (apiTypeToUse === ENCODER_API_TYPES.RESPONSES && (model === GPT_O3 || model === GPTO3)) {
    // Use Responses API - convert messages to single input string
    let input = "";
    
    if (Array.isArray(messages)) {
      // Convert messages array to single input string
      for (const message of messages) {
        if (message.role === "system") {
          input += `System: ${message.content}\n\n`;
        } else if (message.role === "user") {
          input += `User: ${message.content}\n\n`;
        } else if (message.role === "assistant") {
          input += `Assistant: ${message.content}\n\n`;
        }
      }
      input = input.trim();
    } else {
      // Already a string input
      input = messagesOrInput;
    }
    
    console.log(`[callOpenAI] Using Responses API for ${caller} with model ${model}`);
    yield* callOpenAIResponses(input, options);
  } else {
    // Use Chat Completions API (default)
    console.log(`[callOpenAI] Using Chat Completions API for ${caller} with model ${model}`);
    yield* callOpenAIChatCompletions(messages, options);
  }
}

// OpenAI embeddings function
export async function createEmbedding(text) {
  try {
    console.log("Creating embedding for text");
    
    // >>> ADDED: Log embedding creation details
    console.log("\n╔════════════════════════════════════════════════════════════════╗");
    console.log("║                createEmbedding() CALLED                        ║");
    console.log("╠════════════════════════════════════════════════════════════════╣");
    console.log("║ FILE: AIcalls.js                                               ║");
    console.log("║ FUNCTION: createEmbedding()                                    ║");
    console.log("║ PURPOSE: Creating text embeddings for vector DB search         ║");
    console.log("╚════════════════════════════════════════════════════════════════╝");
    console.log("Full text being embedded:");
    console.log("────────────────────────────────────────────────────────────────");
    console.log(text);
    console.log("────────────────────────────────────────────────────────────────\n");
    // <<< END ADDED

    if (!INTERNAL_API_KEYS.OPENAI_API_KEY) {
      throw new Error("OpenAI API key not found. Please check your API keys.");
    }

    const response = await fetch('https://api.openai.com/v1/embeddings', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${INTERNAL_API_KEYS.OPENAI_API_KEY}`
      },
      body: JSON.stringify({
        model: "text-embedding-3-large",
        input: text
      })
    });

    if (!response.ok) {
      const errorData = await response.json().catch(() => null);
      console.error("OpenAI Embeddings API error response:", errorData);
      throw new Error(`OpenAI Embeddings API error: ${response.status} ${response.statusText}`);
    }

    const data = await response.json();
    console.log("OpenAI Embeddings API response received");
    
    // Add debugging for embedding dimensions
    const embedding = data.data[0].embedding;
    console.log(`Embedding generated: dimension=${embedding.length}, model=text-embedding-3-large`);
    
    return embedding;
  } catch (error) {
    console.error("Error creating embedding:", error);
    throw error;
  }
}

// Function to load prompts from files
export async function loadPromptFromFile(promptKey) {
  try {
    const paths = [
      CONFIG.getPromptUrl(`${promptKey}.txt`),
      ...srcPaths // Add fallback paths if needed
    ];

    let response = null;
    for (const path of paths) {
      if (DEBUG) console.log(`Attempting to load prompt from: ${path}`);
      try {
        response = await fetch(path);
        if (response.ok) {
          if (DEBUG) console.log(`Successfully loaded prompt from: ${path}`);
          break;
        }
      } catch (err) {
        if (DEBUG) console.log(`Path ${path} failed: ${err.message}`);
      }
    }

    if (!response || !response.ok) {
      throw new Error(`Failed to load prompt: ${promptKey} (Could not find file in any location)`);
    }

    return await response.text();
  } catch (error) {
    console.error(`Error loading prompt ${promptKey}:`, error);
    throw error;
  }
}

// Function to get system prompt from file
export async function getSystemPromptFromFile(promptKey) {
  try {
    const prompt = await loadPromptFromFile(promptKey);
    if (!prompt) {
      throw new Error(`Prompt key "${promptKey}" not found`);
    }
    return prompt;
  } catch (error) {
    console.error(`Error getting prompt for key ${promptKey}:`, error);
    return null; // Return null or handle error as appropriate
  }
};

// Function: OpenAI Call with conversation history support
export async function processPrompt({ userInput, systemPrompt, model, temperature, history = [], promptFiles = {}, useClaudeAPI = false }) {
    if (DEBUG) console.log("API Key being used for processPrompt:", INTERNAL_API_KEYS.OPENAI_API_KEY ? `${INTERNAL_API_KEYS.OPENAI_API_KEY.substring(0, 3)}...` : "None");

    // >>> ADDED: Log the function call details with prompt file info
    console.log("\n╔══════════════════════════════════════╗");
    console.log("║    processPrompt() CALLED            ║");
    console.log("╚══════════════════════════════════════╝");
    console.log(`Model: ${model}`);
    console.log(`Temperature: ${temperature}`);
    console.log(`History items: ${history.length}`);
    if (promptFiles.system) console.log(`System Prompt File: ${promptFiles.system}`);
    if (promptFiles.main) console.log(`Main Prompt File: ${promptFiles.main}`);
    console.log("────────────────────────────────────────\n");
    // <<< END ADDED

    const messages = [
        { role: "system", content: systemPrompt }
    ];

    if (history.length > 0) {
        history.forEach(message => {
            // Ensure message is in the correct format [role, content]
             if (Array.isArray(message) && message.length === 2) {
                 messages.push({
                     role: message[0] === "human" ? "user" : "assistant",
                     content: message[1]
                 });
             } else {
                 console.warn("Skipping malformed history message:", message);
             }
        });
    }

    messages.push({ role: "user", content: userInput });

    try {
        // Correctly call callOpenAI with an options object including caller info
        const openaiCallOptions = { 
            model: model, 
            temperature: temperature, 
            stream: false,
            useClaudeAPI: useClaudeAPI,
            caller: `processPrompt() - ${promptFiles.system || 'Unknown System Prompt'}`
        };
        let responseContent = "";

        // Consume the async iterator from callOpenAI
        // For non-streaming, this loop will run once, getting the single yielded string content.
        for await (const contentPart of callOpenAI(messages, openaiCallOptions)) {
            responseContent += contentPart;
        }

        // Try to parse JSON response if applicable, otherwise split lines
        try {
            const parsed = JSON.parse(responseContent);
            // Expecting an array of strings based on original code
            if (Array.isArray(parsed)) {
                return parsed;
            }
            // If not array, maybe single string JSON? Unlikely based on usage.
             console.warn("Parsed JSON response, but it was not an array:", parsed);
             // Fallback to splitting the original string
             return responseContent.split('\n').filter(line => line.trim());
        } catch (e) {
            return responseContent.split('\n').filter(line => line.trim());
        }
    } catch (error) {
        console.error("Error in processPrompt:", error);
        throw error; // Re-throw to be caught by caller
    }
}

// Function: Structure database queries
export async function structureDatabasequeries(clientprompt, progressCallback = null) {
  if (DEBUG) console.log("Processing structured database queries:", clientprompt);

  try {
      if (DEBUG) console.log("Getting structure system prompt");
      const systemStructurePrompt = await getSystemPromptFromFile('Structure_System');

      if (!systemStructurePrompt) {
          throw new Error("Failed to load structure system prompt");
      }

      if (DEBUG) console.log("Got system prompt, processing query strings");
      // processPrompt expects history, pass empty array if none applicable here
              const queryStrings = await processPrompt({
            userInput: clientprompt,
            systemPrompt: systemStructurePrompt,
            model: "claude-opus-4-1-20250805",
            temperature: 1,
            history: [], // Explicitly empty
            promptFiles: { system: 'Structure_System' },
            useClaudeAPI: true
        });

      // >>> ADDED: Console log the full response array from Prompt Breakup call
      console.log("\n╔════════════════════════════════════════════════════════════════╗");
      console.log("║                    PROMPT BREAKUP RESPONSE                     ║");
      console.log("╠════════════════════════════════════════════════════════════════╣");
      console.log("║ Full Response Array from Structure_System call:                ║");
      console.log("╚════════════════════════════════════════════════════════════════╝");
      console.log("Response Array Type:", Array.isArray(queryStrings) ? "Array" : typeof queryStrings);
      console.log("Response Array Length:", queryStrings?.length || "N/A");
      console.log("Full Response Array Contents:");
      console.log(queryStrings);
      console.log("────────────────────────────────────────────────────────────────");
      if (Array.isArray(queryStrings)) {
          queryStrings.forEach((item, index) => {
              console.log(`[${index}]:`, item);
          });
      }
      console.log("────────────────────────────────────────────────────────────────\n");
      // <<< END ADDED

      if (!queryStrings || !Array.isArray(queryStrings)) {
          console.error("Invalid query strings received:", queryStrings);
          throw new Error("Failed to get valid query strings from structuring prompt");
      }

      if (DEBUG) {
          console.log("=== PROMPT PARSING RESULTS ===");
          console.log(`Original prompt: "${clientprompt}"`);
          console.log(`Parsed into ${queryStrings.length} chunks:`);
          queryStrings.forEach((chunk, index) => {
              console.log(`  Chunk ${index + 1}: "${chunk}"`);
          });
          console.log("=== END PARSING RESULTS ===");
      }

      const results = [];

      // >>> ADDED: Direct query using entire client prompt to training data database
      if (progressCallback) {
          progressCallback("Querying training data with full prompt...");
      }
      
      if (DEBUG) {
          console.log("=== PROCESSING FULL PROMPT QUERY ===");
          console.log(`Full prompt query: "${clientprompt}"`);
          console.log("Querying call2trainingdata database with full prompt...");
      }
      
      try {
          const fullPromptTrainingData = await queryVectorDB({
              queryPrompt: clientprompt,
              similarityThreshold: .4,
              indexName: 'call2trainingdata',
              numResults: 30 // Slightly higher number since this is the full prompt
          });

          // Add the full prompt query as a separate result
          const fullPromptResult = {
              query: clientprompt + " (FULL PROMPT)",
              trainingData: fullPromptTrainingData,
              call2Context: [], // Empty for full prompt query

            //   codeOptions: [] // Empty for full prompt query
          };

          results.push(fullPromptResult);
          
          if (DEBUG) {
              console.log("Full prompt query results summary:");
              console.log(`  - Training data: ${fullPromptTrainingData.length} items`);
              console.log("=== END FULL PROMPT QUERY ===");
          }
          
          if (progressCallback) {
              progressCallback("Completed full prompt training data query");
          }
      } catch (error) {
          console.error(`Error processing full prompt query "${clientprompt}":`, error);
          if (progressCallback) {
              progressCallback(`Error in full prompt query: ${error.message}`);
          }
      }

      // Create async function to process each chunk
      const processChunk = async (queryString, chunkNumber) => {
          if (progressCallback) {
              progressCallback(`Processing chunk ${chunkNumber} of ${queryStrings.length}: "${queryString.substring(0, 50)}${queryString.length > 50 ? '...' : ''}"`);
          }
          
          if (DEBUG) {
              console.log(`=== PROCESSING CHUNK ${chunkNumber}/${queryStrings.length} ===`);
              console.log(`Query: "${queryString}"`);
              console.log("Querying vector databases...");
          }
          
          try {
              // Process both database queries in parallel for each chunk
              // >>> COMMENTED OUT: Context database queries
              // const [trainingData, call2Context] = await Promise.all([
              //     queryVectorDB({
              //         queryPrompt: queryString,
              //         similarityThreshold: .4,
              //         indexName: 'call2trainingdata',
              //         numResults: 30
              //     }),
              //     queryVectorDB({
              //         queryPrompt: queryString,
              //         similarityThreshold: .2,
              //         indexName: 'call2context',
              //         numResults: 20
              //     })
              // ]);

              // Only query training data, skip context queries
              const trainingData = await queryVectorDB({
                  queryPrompt: queryString,
                  similarityThreshold: .4,
                  indexName: 'call2trainingdata',
                  numResults: 30
              });

              const queryResults = {
                  query: queryString,
                  trainingData,
                  call2Context: [] // Empty array since context queries are commented out
              };
              
              if (DEBUG) {
                  console.log(`Chunk ${chunkNumber} results summary:`);
                  console.log(`  - Training data: ${queryResults.trainingData.length} items`);
                  console.log(`  - Context: ${queryResults.call2Context.length} items (DISABLED - context queries commented out)`);
                  console.log(`=== END CHUNK ${chunkNumber} ===`);
              }
              
              if (progressCallback) {
                  progressCallback(`Completed chunk ${chunkNumber} of ${queryStrings.length}`);
              }

              return queryResults;
          } catch (error) {
              console.error(`Error processing query "${queryString}":`, error);
              if (progressCallback) {
                  progressCallback(`Error in chunk ${chunkNumber}: ${error.message}`);
              }
              // Return null for failed chunks, we'll filter them out later
              return null;
          }
      };

      // Process all chunks in parallel
      if (progressCallback) {
          progressCallback(`Starting parallel processing of ${queryStrings.length} chunks...`);
      }
      
      if (DEBUG) {
          console.log(`=== PROCESSING ${queryStrings.length} CHUNKS IN PARALLEL ===`);
      }

      const chunkPromises = queryStrings.map((queryString, index) => 
          processChunk(queryString, index + 1)
      );

      // Wait for all chunks to complete (using allSettled to handle partial failures)
      const chunkResults = await Promise.allSettled(chunkPromises);
      
      // Process results and handle any failures
      chunkResults.forEach((result, index) => {
          if (result.status === 'fulfilled' && result.value !== null) {
              results.push(result.value);
          } else if (result.status === 'rejected') {
              console.error(`Chunk ${index + 1} failed with error:`, result.reason);
              if (progressCallback) {
                  progressCallback(`Chunk ${index + 1} failed: ${result.reason.message}`);
              }
          }
          // Note: fulfilled but null results (handled errors) are already logged in processChunk
      });

      if (DEBUG) {
          console.log(`=== PARALLEL PROCESSING COMPLETED ===`);
          console.log(`Successfully processed ${results.length - 1} out of ${queryStrings.length} chunks`); // -1 to exclude full prompt query
      }

      if (results.length === 0 && queryStrings.length > 0) {
           console.warn("All structured queries failed to produce results.");
           // Decide whether to throw an error or return empty results
           // Throwing error based on original logic
           throw new Error("No valid results were obtained from any structured queries");
      } else if (queryStrings.length === 0) {
           console.warn("Structuring prompt returned no query strings.");
           // Throwing error as subsequent steps likely depend on results
           throw new Error("Structuring prompt did not return any queries to process.");
      }

      if (DEBUG) {
          console.log("=== FINAL SUMMARY ===");
          console.log(`Successfully processed ${results.length} total queries (including full prompt query)`);
          console.log("=== END SUMMARY ===");
      }

      // >>> ADDED: Deduplicate training data to remove redundant code strings based on output codestrings
      const deduplicatedResults = deduplicateTrainingDataByOutput(results);

      return deduplicatedResults;
  } catch (error) {
      console.error("Error in structureDatabasequeries:", error);
      throw error; // Re-throw
  }
}

// >>> ADDED: Function to deduplicate training data entries
export function deduplicateTrainingData(results) {
    const seenInputs = new Map(); // Maps input portion to first occurrence details
    let totalOriginalCount = 0;
    let totalDeduplicatedCount = 0;

    // Process each result set
    results.forEach((result, resultIndex) => {
        if (!result.trainingData || !Array.isArray(result.trainingData)) {
            return;
        }

        totalOriginalCount += result.trainingData.length;

        // Process each training data entry in this result
        result.trainingData = result.trainingData.map((trainingEntry, entryIndex) => {
            // Extract the input portion (before code strings)
            const inputPortion = extractInputPortion(trainingEntry);
            
            if (!inputPortion) {
                // If we can't extract input portion, keep the original
                return trainingEntry;
            }

            // Check if we've seen this input before
            if (seenInputs.has(inputPortion)) {
                // This is a duplicate, replace with reference
                const firstOccurrence = seenInputs.get(inputPortion);
                return `${inputPortion} Output: (Duplicate: See codes in earlier example above)`;
            } else {
                // First occurrence, store it and keep the original
                seenInputs.set(inputPortion, { resultIndex, entryIndex });
                return trainingEntry;
            }
        });

        totalDeduplicatedCount += result.trainingData.length;
    });

    return results;
}

// >>> ADDED: Function to deduplicate training data entries based on output codestrings
export function deduplicateTrainingDataByOutput(results) {
    const seenOutputs = new Set(); // Track unique output codestrings
    let totalOriginalCount = 0;
    let totalDeduplicatedCount = 0;
    let duplicatesRemoved = 0;

    if (DEBUG) console.log("[deduplicateTrainingDataByOutput] Starting output-based deduplication...");

    // Process each result set
    results.forEach((result, resultIndex) => {
        if (!result.trainingData || !Array.isArray(result.trainingData)) {
            return;
        }

        const originalLength = result.trainingData.length;
        totalOriginalCount += originalLength;

        if (DEBUG) console.log(`[deduplicateTrainingDataByOutput] Processing result ${resultIndex + 1} with ${originalLength} training entries`);

        // Filter training data to remove entries with duplicate outputs
        result.trainingData = result.trainingData.filter((trainingEntry, entryIndex) => {
            // Handle both old string format and new object format
            const trainingText = typeof trainingEntry === 'object' ? trainingEntry.text : trainingEntry;
            
            // Extract the output portion (codestrings after "Output:")
            const outputCodestrings = extractOutputCodestrings(trainingText);
            
            if (!outputCodestrings) {
                // If we can't extract output codestrings, keep the entry
                return true;
            }

            // Check if we've seen this output before
            if (seenOutputs.has(outputCodestrings)) {
                // This is a duplicate output, remove it
                duplicatesRemoved++;
                if (DEBUG) console.log(`[deduplicateTrainingDataByOutput] Duplicate output found, removing entry ${entryIndex}`);
                if (DEBUG) console.log(`[deduplicateTrainingDataByOutput] Duplicate output: ${outputCodestrings.substring(0, 100)}...`);
                return false;
            } else {
                // First occurrence of this output, keep it and mark as seen
                seenOutputs.add(outputCodestrings);
                if (DEBUG) console.log(`[deduplicateTrainingDataByOutput] New unique output found, keeping entry ${entryIndex}`);
                return true;
            }
        });

        const newLength = result.trainingData.length;
        totalDeduplicatedCount += newLength;
        
        if (DEBUG) console.log(`[deduplicateTrainingDataByOutput] Result ${resultIndex + 1}: ${originalLength} -> ${newLength} (removed ${originalLength - newLength} duplicates)`);
    });

    if (DEBUG) {
        console.log(`[deduplicateTrainingDataByOutput] Deduplication complete:`);
        console.log(`  - Total original entries: ${totalOriginalCount}`);
        console.log(`  - Total after deduplication: ${totalDeduplicatedCount}`);
        console.log(`  - Duplicates removed: ${duplicatesRemoved}`);
        console.log(`  - Unique outputs retained: ${seenOutputs.size}`);
        console.log(`  - Similarity scores preserved: TRUE`);
    }

    return results;
}

// >>> ADDED: Helper function to extract output codestrings from training data entry
function extractOutputCodestrings(trainingEntry) {
    if (!trainingEntry || typeof trainingEntry !== 'string') {
        return null;
    }

    // Find the "Output:" portion which typically precedes code strings in training data
    const outputMatch = trainingEntry.search(/\bOutput:\s*/i);
    if (outputMatch === -1) {
        // No "Output:" found, look for code patterns directly
        const codestrings = extractCodestrings(trainingEntry);
        return codestrings.length > 0 ? codestrings.join('\n') : null;
    }

    // Extract everything after "Output:" which should contain the codestrings
    const outputPortion = trainingEntry.substring(outputMatch + trainingEntry.match(/\bOutput:\s*/i)[0].length).trim();
    
    // Extract codestrings from the output portion
    const codestrings = extractCodestrings(outputPortion);
    
    // If no codestrings found in output portion, use the raw output portion
    if (codestrings.length === 0) {
        return outputPortion || null;
    }
    
    // Return normalized codestrings joined consistently
    // Sort to ensure order doesn't matter for comparison
    const sortedCodestrings = codestrings.sort();
    return sortedCodestrings.join('\n');
}

// >>> ADDED: Helper function to format code strings in training entry output
function formatCodeStringsInTrainingEntry(trainingEntry) {
    if (!trainingEntry || typeof trainingEntry !== 'string') {
        return trainingEntry;
    }

    // Find the "Output:" portion
    const outputMatch = trainingEntry.search(/\bOutput:\s*/i);
    if (outputMatch === -1) {
        // No "Output:" found, return as is
        return trainingEntry;
    }

    // Split into input and output portions
    const inputPortion = trainingEntry.substring(0, outputMatch + trainingEntry.match(/\bOutput:\s*/i)[0].length);
    const outputPortion = trainingEntry.substring(outputMatch + trainingEntry.match(/\bOutput:\s*/i)[0].length);

    // Add line breaks between code strings (< and >) in the output portion
    // This regex finds code strings like <BR;...>, <LABELH1;...>, etc.
    const formattedOutput = outputPortion.replace(/></g, '>\n<');

    return inputPortion + formattedOutput;
}

// >>> ADDED: Function to consolidate and deduplicate training data and context across ALL queries
export function consolidateAndDeduplicateTrainingData(results) {
    const allTrainingData = [];
    const allContextData = [];
    const seenTrainingOutputs = new Set(); // Track unique training output codestrings
    const seenContextData = new Set(); // Track unique context entries
    let totalOriginalTrainingCount = 0;
    let totalOriginalContextCount = 0;
    let trainingDuplicatesRemoved = 0;
    let contextDuplicatesRemoved = 0;

    // Collect all training data and context data from all results
    results.forEach((result, resultIndex) => {
        // Process training data
        if (result.trainingData && Array.isArray(result.trainingData)) {
            totalOriginalTrainingCount += result.trainingData.length;

            result.trainingData.forEach((trainingEntry, entryIndex) => {
                // Handle new object format with separate input/output fields
                let trainingText = "";
                let trainingScore = null;
                let outputCodestrings = "";

                                 if (typeof trainingEntry === 'object') {
                     // New format: separate text (input) and output fields
                     if (trainingEntry.text && trainingEntry.output) {
                         // Combine input and output into expected training format
                         trainingText = `Client input: ${trainingEntry.text} Output: ${trainingEntry.output}`;
                         outputCodestrings = trainingEntry.output; // Use output directly for deduplication
                         if (DEBUG) {
                             console.log(`[consolidateAndDeduplicateTrainingData] Combined input/output - Input: ${trainingEntry.text.substring(0, 50)}..., Output: ${trainingEntry.output.substring(0, 50)}...`);
                         }
                     } else if (trainingEntry.text) {
                         // Only input text available (backwards compatibility)
                         trainingText = trainingEntry.text;
                         outputCodestrings = extractOutputCodestrings(trainingText);
                         if (DEBUG) {
                             console.log(`[consolidateAndDeduplicateTrainingData] Using text field only (backwards compatibility): ${trainingText.substring(0, 50)}...`);
                         }
                     } else {
                         // Handle unexpected object structure
                         trainingText = JSON.stringify(trainingEntry);
                         outputCodestrings = extractOutputCodestrings(trainingText);
                         if (DEBUG) {
                             console.log(`[consolidateAndDeduplicateTrainingData] Unexpected object structure, stringified: ${trainingText.substring(0, 50)}...`);
                         }
                     }
                     trainingScore = trainingEntry.score;
                 } else {
                     // Old format: single string with input and output combined
                     trainingText = trainingEntry;
                     outputCodestrings = extractOutputCodestrings(trainingText);
                     if (DEBUG) {
                         console.log(`[consolidateAndDeduplicateTrainingData] Using old string format: ${trainingText.substring(0, 50)}...`);
                     }
                 }
                
                if (!outputCodestrings) {
                    // If we can't extract output codestrings, keep the original but mark it as processed
                    allTrainingData.push({ text: trainingText, score: trainingScore });
                    return;
                }

                // Check if we've seen this output before
                if (seenTrainingOutputs.has(outputCodestrings)) {
                    // This is a duplicate output, skip it
                    trainingDuplicatesRemoved++;
                } else {
                    // First occurrence of this output, add it and mark as seen
                    seenTrainingOutputs.add(outputCodestrings);
                    allTrainingData.push({ text: trainingText, score: trainingScore });
                }
            });
        }

        // Process context data (NOTE: Context queries are currently DISABLED/commented out, so this will process empty arrays)
        if (result.call2Context && Array.isArray(result.call2Context)) {
            totalOriginalContextCount += result.call2Context.length;

            result.call2Context.forEach((contextEntry, entryIndex) => {
                // Handle both old string format and new object format
                let contextText = "";
                let contextScore = null;

                if (typeof contextEntry === 'object') {
                    // For context data, we typically just use the text field
                    // (context data might not have separate input/output structure)
                    contextText = contextEntry.text || JSON.stringify(contextEntry);
                    contextScore = contextEntry.score;
                } else {
                    // Old format: single string
                    contextText = contextEntry;
                }
                
                // For context, use the full entry as the unique identifier
                const contextKey = contextText.trim();
                
                if (contextKey && !seenContextData.has(contextKey)) {
                    seenContextData.add(contextKey);
                    allContextData.push({ text: contextText, score: contextScore });
                } else if (contextKey) {
                    contextDuplicatesRemoved++;
                }
            });
        }
    });

    // >>> ADDED: Apply training data filtering based on similarity thresholds
    console.log(`\n🔍 === TRAINING DATA FILTERING ===`);
    console.log(`📊 Training data before filtering: ${allTrainingData.length} entries`);
    
    // Sort training data by similarity score (highest first)
    const sortedTrainingData = allTrainingData.sort((a, b) => {
        // Handle null scores - put them at the end
        if (a.score === null && b.score === null) return 0;
        if (a.score === null) return 1;
        if (b.score === null) return -1;
        return b.score - a.score; // Descending order (highest similarity first)
    });
    
    // Apply filtering logic:
    // 1) Include up to 10 sets if similarity > 0.6
    // 2) Include at least 4 sets regardless of similarity
    const highQualityEntries = sortedTrainingData.filter(entry => entry.score !== null && entry.score > 0.6);
    const filteredTrainingData = [];
    
    if (highQualityEntries.length === 0) {
        // No matches above 0.6 - include top 4 (or all available if fewer than 4)
        const entriesToInclude = Math.min(4, sortedTrainingData.length);
        filteredTrainingData.push(...sortedTrainingData.slice(0, entriesToInclude));
        console.log(`📊 No entries above 0.6 similarity - including top ${entriesToInclude} entries`);
    } else if (highQualityEntries.length <= 10) {
        // 1-10 matches above 0.6 - include all of them, but ensure minimum of 4 (or all available if fewer than 4)
        const entriesToInclude = Math.max(Math.min(4, sortedTrainingData.length), highQualityEntries.length);
        filteredTrainingData.push(...sortedTrainingData.slice(0, entriesToInclude));
        console.log(`📊 ${highQualityEntries.length} entries above 0.6 similarity - including ${entriesToInclude} entries (minimum 4 or all available)`);
    } else {
        // 10+ matches above 0.6 - include only the top 10
        filteredTrainingData.push(...highQualityEntries.slice(0, 10));
        console.log(`📊 ${highQualityEntries.length} entries above 0.6 similarity - including top 10`);
    }
    
    // Log the filtering results
    const entriesAbove06 = filteredTrainingData.filter(entry => entry.score !== null && entry.score > 0.6).length;
    const minScore = filteredTrainingData.length > 0 ? Math.min(...filteredTrainingData.map(e => e.score || 0)) : 0;
    const maxScore = filteredTrainingData.length > 0 ? Math.max(...filteredTrainingData.map(e => e.score || 0)) : 0;
    
    console.log(`📊 Filtering results:`);
    console.log(`  - Final count: ${filteredTrainingData.length} entries`);
    console.log(`  - Entries above 0.6: ${entriesAbove06}`);
    console.log(`  - Score range: ${minScore.toFixed(4)} - ${maxScore.toFixed(4)}`);
    console.log(`🔍 === END TRAINING DATA FILTERING ===\n`);

    // Format the consolidated context data first
    let formattedOutput = "";
    if (allContextData.length > 0) {
        formattedOutput += "Client request-specific Context:\n****\n";
        
        allContextData.forEach((contextEntry, index) => {
            const scorePrefix = contextEntry.score !== null ? ` - similarity threshold: ${contextEntry.score.toFixed(4)}:` : ')';
            formattedOutput += `Context item ${index + 1}${scorePrefix} ${contextEntry.text}\n***\n`;
        });
        formattedOutput += "\n";
    }

    // Format the filtered training data
    formattedOutput += "Training Data:\n****\n";
    
    filteredTrainingData.forEach((trainingEntry, index) => {
        // Process the training entry to add line breaks between code strings in output
        let processedEntry = formatCodeStringsInTrainingEntry(trainingEntry.text);
        
        // Ensure training data starts with "Client input:" if it doesn't already
        if (!processedEntry.toLowerCase().startsWith('client input:')) {
            processedEntry = `Client input: ${processedEntry}`;
        }
        
        const scorePrefix = trainingEntry.score !== null ? ` - similarity threshold: ${trainingEntry.score.toFixed(4)}:` : ')';
        formattedOutput += `Input/output set ${index + 1}${scorePrefix} ${processedEntry}\n***\n`;
    });

    if (DEBUG) {
        console.log(`[consolidateAndDeduplicateTrainingData] Enhanced prompt generated with similarity scores:`);
        console.log(`  - Training data entries (after filtering): ${filteredTrainingData.length} (before filtering: ${allTrainingData.length}, total original: ${totalOriginalTrainingCount})`);
        console.log(`  - Context entries: ${allContextData.length} (total original: ${totalOriginalContextCount}) [DISABLED - context queries commented out]`);
        console.log(`  - Training duplicates removed: ${trainingDuplicatesRemoved}`);
        console.log(`  - Context duplicates removed: ${contextDuplicatesRemoved}`);
        console.log(`  - Training entries with scores: ${filteredTrainingData.filter(item => item.score !== null).length}`);
        console.log(`  - Context entries with scores: ${allContextData.filter(item => item.score !== null).length}`);
        console.log(`  - Training entries filtered out: ${allTrainingData.length - filteredTrainingData.length}`);
        console.log(`  - Using new database structure with separate input/output fields`);
    }

    return formattedOutput;
}

// >>> ADDED: Helper function to extract codestrings (items between <> brackets) from text
function extractCodestrings(text) {
    if (!text || typeof text !== 'string') {
        return [];
    }
    
    // Find all text between < and > brackets
    const codestringMatches = text.match(/<[^>]*>/g);
    return codestringMatches || [];
}

// >>> ADDED: Helper function to compare and log codestring changes
function logCodestringComparison(beforeText, afterText, stepName) {
    // Convert arrays to strings if needed
    const beforeStr = Array.isArray(beforeText) ? beforeText.join('\n') : String(beforeText);
    const afterStr = Array.isArray(afterText) ? afterText.join('\n') : String(afterText);
    
    const beforeCodestrings = extractCodestrings(beforeStr);
    const afterCodestrings = extractCodestrings(afterStr);
    
    // Create sets for easier comparison
    const beforeSet = new Set(beforeCodestrings);
    const afterSet = new Set(afterCodestrings);
    
    // Find codestrings that are different
    const removedCodestrings = beforeCodestrings.filter(code => !afterSet.has(code));
    const addedCodestrings = afterCodestrings.filter(code => !beforeSet.has(code));
    
    // Only log if there are differences
    if (removedCodestrings.length > 0 || addedCodestrings.length > 0) {
        console.log(`\n📊 ${stepName} - Changes detected:`);
        
        // Helper function to extract key identifiers for matching
        const getCodeKey = (codestring) => {
            const codeTypeMatch = codestring.match(/<([^;]+);/);
            const codeType = codeTypeMatch ? codeTypeMatch[1] : '';
            
            // Extract row1 parameter to identify the same logical codestring
            const row1Match = codestring.match(/row1\s*=\s*"([^"]*)"/) || codestring.match(/row1\s*=\s*"([^"]*)/);
            const row1 = row1Match ? row1Match[1].split('|')[0] : ''; // Get driver ID (first part before |)
            
            return `${codeType}:${row1}`;
        };
        
        // Create maps for matching
        const removedByKey = new Map();
        const addedByKey = new Map();
        
        removedCodestrings.forEach(code => {
            const key = getCodeKey(code);
            if (!removedByKey.has(key)) removedByKey.set(key, []);
            removedByKey.get(key).push(code);
        });
        
        addedCodestrings.forEach(code => {
            const key = getCodeKey(code);
            if (!addedByKey.has(key)) addedByKey.set(key, []);
            addedByKey.get(key).push(code);
        });
        
        // Find matches and show before/after pairs
        const processedKeys = new Set();
        let changeCount = 1;
        
        // Process matched pairs first
        removedByKey.forEach((removedList, key) => {
            if (addedByKey.has(key)) {
                const addedList = addedByKey.get(key);
                
                // Pair them up (assuming 1:1 mapping for simplicity)
                for (let i = 0; i < Math.max(removedList.length, addedList.length); i++) {
                    const before = removedList[i];
                    const after = addedList[i];
                    
                    if (before && after) {
                        console.log(`   ${changeCount++}. MODIFIED:`);
                        console.log(`      BEFORE: ${before}`);
                        console.log(`      AFTER:  ${after}`);
                    } else if (before) {
                        console.log(`   ${changeCount++}. REMOVED: ${before}`);
                    } else if (after) {
                        console.log(`   ${changeCount++}. ADDED: ${after}`);
                    }
                }
                processedKeys.add(key);
            }
        });
        
        // Process remaining removed items (no match found)
        removedByKey.forEach((removedList, key) => {
            if (!processedKeys.has(key)) {
                removedList.forEach(code => {
                    console.log(`   ${changeCount++}. REMOVED: ${code}`);
                });
            }
        });
        
        // Process remaining added items (no match found)
        addedByKey.forEach((addedList, key) => {
            if (!processedKeys.has(key)) {
                addedList.forEach(code => {
                    console.log(`   ${changeCount++}. ADDED: ${code}`);
                });
            }
        });
        
        console.log(''); // Add spacing
    } else {
        console.log(`✨ ${stepName}: No changes detected`);
    }
}

// >>> ADDED: Helper function to extract input portion from training data entry
function extractInputPortion(trainingEntry) {
    if (!trainingEntry || typeof trainingEntry !== 'string') {
        return null;
    }

    // First try to find "Output:" which typically precedes code strings in training data
    const outputMatch = trainingEntry.search(/\bOutput:\s*/i);
    if (outputMatch !== -1) {
        // Extract everything before "Output:" and preserve full input including trailing punctuation
        const inputPortion = trainingEntry.substring(0, outputMatch).trim();
        return inputPortion;
    }

    // Fallback: Look for common code string patterns in your training data
    // Based on the actual patterns I see: <BR;, <LABELH1;, <SPREAD-E;, etc.
    const codePatterns = [
        /<BR;/i,                 // <BR; pattern
        /<LABELH\d*;/i,          // <LABELH1;, <LABELH3;, etc.
        /<SPREAD-E;/i,           // <SPREAD-E; pattern
        /<CONST-E;/i,            // <CONST-E; pattern
        /<OFFSETCOLUMN-S;/i,     // <OFFSETCOLUMN-S; pattern
        /<DIRECT-S;/i,           // <DIRECT-S; pattern
        /<MULT2-S;/i,            // <MULT2-S; pattern
        /<SUM2-S;/i,             // <SUM2-S; pattern
        /<AVGMULT3-S;/i,         // <AVGMULT3-S; pattern
        /<DEANNUALIZE-S;/i,      // <DEANNUALIZE-S; pattern
        /<SUBTRACT2-S;/i,        // <SUBTRACT2-S; pattern
        /<ENDPOINT-E;/i,         // <ENDPOINT-E; pattern
        /<const-/i,              // Original <const-e; pattern (keep for compatibility)
        /<code\d*/i,             // Original <code2>, <code3>, etc. (keep for compatibility)
        /<[A-Z][A-Z0-9_]*-[A-Z];/i  // Generic pattern for code markers like PATTERN-X;
    ];

    let splitIndex = -1;
    
    // Find the earliest occurrence of any code pattern
    for (const pattern of codePatterns) {
        const match = trainingEntry.search(pattern);
        if (match !== -1 && (splitIndex === -1 || match < splitIndex)) {
            splitIndex = match;
        }
    }

    if (splitIndex === -1) {
        // No code patterns found, assume the entire entry is input
        // This might happen with short entries or different formats
        return trainingEntry.trim();
    }

    // Extract everything before the first code pattern and preserve full input
    const inputPortion = trainingEntry.substring(0, splitIndex).trim();
    
    // Return the full input portion without removing trailing punctuation
    return inputPortion;
}

// Function: Query Vector Database using Pinecone REST API
export async function queryVectorDB({ queryPrompt, indexName = 'codes', numResults = 10, similarityThreshold = null }) {
    try {
        if (DEBUG) console.log("Generating embeddings for query:", queryPrompt);

        // Ensure API key exists before proceeding
        if (!INTERNAL_API_KEYS.PINECONE_API_KEY) {
            throw new Error("Pinecone API key not found. Please check your API keys.");
        }

        const embedding = await createEmbedding(queryPrompt); // Uses OpenAI key internally
        if (DEBUG) console.log("Embeddings generated successfully");

        const indexConfig = PINECONE_INDEXES[indexName];
        if (!indexConfig) {
            throw new Error(`Invalid index name provided: ${indexName}`);
        }

        const url = `${indexConfig.apiEndpoint}/query`;
        if (DEBUG) console.log("Making Pinecone API request to:", url);

        // Add detailed debugging for Pinecone request
        const requestBody = {
            vector: embedding,
            topK: numResults,
            includeMetadata: true,
            namespace: "ns1" // Assuming namespace is constant
        };
        
        console.log("Pinecone Request Debug Info:", {
            url: url,
            indexName: indexName,
            vectorLength: embedding ? embedding.length : 'undefined',
            vectorSample: embedding ? `[${embedding.slice(0, 3).join(', ')}...${embedding.slice(-3).join(', ')}]` : 'no vector',
            topK: numResults,
            namespace: "ns1",
            hasApiKey: !!INTERNAL_API_KEYS.PINECONE_API_KEY,
            apiKeyPrefix: INTERNAL_API_KEYS.PINECONE_API_KEY ? INTERNAL_API_KEYS.PINECONE_API_KEY.substring(0, 10) + '...' : 'missing'
        });

        const response = await fetch(url, {
            method: 'POST',
            headers: {
                'api-key': INTERNAL_API_KEYS.PINECONE_API_KEY,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(requestBody)
        });

        if (!response.ok) {
            const errorText = await response.text().catch(() => "Could not read error response body");
            console.error("Pinecone API error details:", {
                status: response.status,
                statusText: response.statusText,
                error: errorText
            });
            throw new Error(`Pinecone API error: ${response.status} ${response.statusText} - ${errorText}`);
        }

        const data = await response.json();
        if (DEBUG) console.log("Pinecone API response received");

        let matches = data.matches || [];

        if (similarityThreshold !== null) {
            matches = matches.filter(match => match.score >= similarityThreshold);
        }

        // Apply numResults limit *after* threshold filtering
        matches = matches.slice(0, numResults);

        // Extract text and preserve similarity scores
        const enrichedMatches = matches.map(match => extractTextAndScoreFromJson(match)).filter(item => item.text !== "");

        if (DEBUG) {
            console.log(`Found ${enrichedMatches.length} matches (after threshold/limit/extraction):`);
            enrichedMatches.forEach((item, i) => console.log(`  ${i + 1}: Score ${item.score.toFixed(4)} - ${item.text.substring(0, 100)}...`));
        }

        return enrichedMatches;

    } catch (error) {
        console.error(`Error during vector database query for index "${indexName}":`, error);
        throw error; // Re-throw
    }
}


// Diagnostic function to test Pinecone connection and configuration
export async function testPineconeConnection(indexName = 'call2trainingdata') {
    console.log("\n========== PINECONE DIAGNOSTIC TEST ==========");
    
    try {
        // 1. Check API key
        console.log("1. Checking API key...");
        if (!INTERNAL_API_KEYS.PINECONE_API_KEY) {
            console.error("❌ No Pinecone API key found!");
            return;
        }
        console.log("✅ API key found:", INTERNAL_API_KEYS.PINECONE_API_KEY.substring(0, 10) + "...");
        
        // 2. Check index configuration
        console.log("\n2. Checking index configuration...");
        const indexConfig = PINECONE_INDEXES[indexName];
        if (!indexConfig) {
            console.error(`❌ Index '${indexName}' not found in configuration!`);
            console.log("Available indexes:", Object.keys(PINECONE_INDEXES));
            return;
        }
        console.log(`✅ Index configuration found:`, indexConfig);
        
        // 3. Test embedding generation
        console.log("\n3. Testing embedding generation...");
        const testText = "This is a test query";
        const embedding = await createEmbedding(testText);
        console.log(`✅ Embedding generated successfully: ${embedding.length} dimensions`);
        console.log(`   Using model: text-embedding-3-large`);
        
        // 4. Test Pinecone query with minimal vector
        console.log("\n4. Testing Pinecone query...");
        const url = `${indexConfig.apiEndpoint}/query`;
        
        // Create a test vector with correct dimensions
        const testVector = new Array(embedding.length).fill(0.1);
        
        const testBody = {
            vector: testVector,
            topK: 1,
            includeMetadata: true,
            namespace: "ns1"
        };
        
        console.log("Test request details:", {
            url: url,
            vectorDimensions: testVector.length,
            namespace: "ns1"
        });
        
        const response = await fetch(url, {
            method: 'POST',
            headers: {
                'api-key': INTERNAL_API_KEYS.PINECONE_API_KEY,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(testBody)
        });
        
        if (!response.ok) {
            const errorText = await response.text();
            console.error("❌ Pinecone query failed!");
            console.error("Status:", response.status);
            console.error("Error:", errorText);
            
            // Try to parse error for more details
            try {
                const errorJson = JSON.parse(errorText);
                console.error("Parsed error:", errorJson);
            } catch (e) {
                // Not JSON, that's ok
            }
        } else {
            const data = await response.json();
            console.log("✅ Pinecone query successful!");
            console.log("Response:", data);
        }
        
    } catch (error) {
        console.error("❌ Diagnostic test failed:", error);
    }
    
    console.log("\n========== END DIAGNOSTIC TEST ==========\n");
}

// Helper function to extract text from Pinecone match JSON
function extractTextFromJson(jsonInput) {
   try {
       // Input might already be an object if response was parsed
       const jsonData = typeof jsonInput === 'string' ? JSON.parse(jsonInput) : jsonInput;

       // Check common structures
       if (jsonData?.metadata?.text) {
           return jsonData.metadata.text;
       }
       // Fallback for older structures or direct text? (Less likely based on usage)
       if (typeof jsonData?.text === 'string') {
           return jsonData.text;
       }

       // Handle array case (though query response is usually object with 'matches')
       if (Array.isArray(jsonData)) {
           for (const item of jsonData) {
               if (item?.metadata?.text) {
                   return item.metadata.text; // Return first found
               }
           }
       }

       // If no text found
       console.warn("Could not find 'text' field in metadata for match:", JSON.stringify(jsonInput).substring(0, 100));
       return ""; // Return empty string if text cannot be extracted

   } catch (error) {
       console.error(`Error processing JSON for text extraction: ${error.message}`);
       // Log the problematic input for debugging
       console.error("Input causing error:", jsonInput);
       return ""; // Return empty string on error
   }
}

// >>> ADDED: Helper function to extract both text and similarity score from Pinecone match JSON
function extractTextAndScoreFromJson(jsonInput) {
   try {
       // Input might already be an object if response was parsed
       const jsonData = typeof jsonInput === 'string' ? JSON.parse(jsonInput) : jsonInput;

       let text = "";
       let output = "";
       let score = 0;

       // Extract input text using same logic as extractTextFromJson
       if (jsonData?.metadata?.text) {
           text = jsonData.metadata.text;
       } else if (typeof jsonData?.text === 'string') {
           text = jsonData.text;
       } else if (Array.isArray(jsonData)) {
           for (const item of jsonData) {
               if (item?.metadata?.text) {
                   text = item.metadata.text; // Use first found
                   break;
               }
           }
       }

       // Extract output from metadata (new field for call2trainingdata)
       if (jsonData?.metadata?.output) {
           output = jsonData.metadata.output;
       }

       // Extract similarity score
       if (typeof jsonData?.score === 'number') {
           score = jsonData.score;
       }

       // Debug logging for new database structure
       if (DEBUG && text && output) {
           console.log(`[extractTextAndScoreFromJson] Found separate input/output fields - Input: ${text.substring(0, 50)}..., Output: ${output.substring(0, 50)}...`);
       }

       // Return object with text (input), output, and score
       return {
           text: text,
           output: output,
           score: score
       };

   } catch (error) {
       console.error(`Error processing JSON for text and score extraction: ${error.message}`);
       // Log the problematic input for debugging
       console.error("Input causing error:", jsonInput);
       return { text: "", output: "", score: 0 }; // Return empty values on error
   }
}


// Helper function to format JSON for prompts (handle potential errors)
export function safeJsonForPrompt(obj, readable = true) {
    try {
        if (!readable) {
            // Simple stringify, remove potential noise, escape braces
            let jsonString = JSON.stringify(obj);
            // Remove empty values/metadata pairs if they exist and are noise
            // Be cautious with overly broad replaces
            // jsonString = jsonString.replace(/"values":\s*\[\],\s*"metadata":/g, '"metadata":'); // Example, adjust if needed
            return jsonString
                // .replace(/{/g, '\\u007B') // Escaping might not be needed depending on LLM
                // .replace(/}/g, '\\u007D');
        }

        // Readable format (extract text, add score)
        if (Array.isArray(obj)) {
            return obj.map(item => {
                let result = "";
                if (item?.metadata?.text) {
                    // Basic cleaning: replace newlines, trim
                    const text = item.metadata.text.replace(/[\r\n]+/g, ' ').trim();
                    result = text; // Use cleaned text directly
                    // Example of splitting if structure was known:
                    // const parts = text.split(';');
                    // if (parts.length >= 1) result += parts[0].trim();
                    // ... etc ...
                } else {
                    // Fallback if no text found
                    result = JSON.stringify(item); // Stringify the whole item as fallback
                }

                if (item?.score) {
                    result += `\nSimilarity Score: ${item.score.toFixed(4)}`;
                }
                return result;
            }).join('\n\n'); // Separate items clearly
        }

        // Fallback for non-array objects (less likely for lists of results)
        return JSON.stringify(obj, null, 2); // Pretty print as fallback

    } catch (error) {
        console.error("Error in safeJsonForPrompt:", error);
        // Return a safe representation of the error or the input
        return `[Error formatting JSON: ${error.message}]`;
    }
}


// Function: Handle Follow-Up Conversation - Standalone Simple Process
export async function handleFollowUpConversation(clientprompt, currentHistory) {
    if (DEBUG) console.log("Processing follow-up question (standalone):", clientprompt);
    if (DEBUG) console.log("Using conversation history length:", currentHistory.length);

    try {
        // Ensure API keys are available
        if (!INTERNAL_API_KEYS.OPENAI_API_KEY) {
            throw new Error("OpenAI API key not initialized for follow-up conversation.");
        }

        // Determine which followup system prompt to use based on current mode
        let dropdown = document.getElementById('mode-dropdown');
        if (!dropdown) {
            dropdown = document.getElementById('system-prompt-dropdown');
        }
        const currentMode = dropdown ? dropdown.value : 'one-shot';
        const followupPromptName = (currentMode === 'one-shot') ? 'oneshotfollowup_system' : 'Followup_System';
        
        // Load the appropriate followup system prompt
        const followupSystemPrompt = await getSystemPromptFromFile(followupPromptName);
        if (!followupSystemPrompt) {
            throw new Error(`Failed to load ${followupPromptName} prompt.`);
        }

        // Extract current codestrings from the last assistant response in history
        let currentCodestrings = "";
        if (currentHistory.length >= 2) {
            const lastAssistantResponse = currentHistory[currentHistory.length - 1];
            if (lastAssistantResponse[0] === "assistant") {
                const assistantResponse = lastAssistantResponse[1];
                const codestrings = extractCodestrings(assistantResponse);
                currentCodestrings = codestrings.length > 0 ? codestrings.join('\n') : assistantResponse;
            }
        }
        
        // Also try to get existing model codes from localStorage as backup
        let existingModelCodes = "";
        try {
            // Try to use the enhanced storage system first
            try {
                const { getStoredModelCodes } = await import('./ModelCodeStorage.js');
                const storedCodes = getStoredModelCodes(false); // Get raw codes without timestamp
                if (storedCodes && storedCodes.trim()) {
                    existingModelCodes = storedCodes;
                    console.log("[handleFollowUpConversation] Retrieved existing model codes from enhanced storage");
                }
            } catch (importError) {
                // Fallback to direct localStorage access
                const savedModelCodes = localStorage.getItem('rawModelCodes') || localStorage.getItem('lastModelCodes');
                if (savedModelCodes && savedModelCodes.trim()) {
                    existingModelCodes = savedModelCodes;
                    console.log("[handleFollowUpConversation] Retrieved existing model codes from localStorage (fallback)");
                }
            }
        } catch (error) {
            console.warn("[handleFollowUpConversation] Could not retrieve model codes from localStorage:", error);
        }

        // Build conversation history string for the prompt
        let conversationHistoryString = "";
        for (let i = 0; i < currentHistory.length; i += 2) {
            if (i + 1 < currentHistory.length) {
                const humanMessage = currentHistory[i];
                const assistantMessage = currentHistory[i + 1];
                
                if (humanMessage[0] === "human" && assistantMessage[0] === "assistant") {
                    // Clean the human message (remove training data if any)
                    let cleanHumanPrompt = humanMessage[1];
                    cleanHumanPrompt = cleanHumanPrompt.replace(/\n\n(Client request-specific Context|Training Data):[\s\S]*$/i, '');
                    cleanHumanPrompt = cleanHumanPrompt.replace(/^Client [Rr]equest:\s*/, '');
                    
                    // Extract codestrings from assistant response
                    const assistantResponse = assistantMessage[1];
                    const codestrings = extractCodestrings(assistantResponse);
                    const codestringsOnly = codestrings.length > 0 ? codestrings.join('\n') : assistantResponse;
                    
                    conversationHistoryString += `Human: ${cleanHumanPrompt.trim()}\nAssistant: ${codestringsOnly}\n\n`;
                }
            }
        }

        // Create the complete prompt for the followup with existing encoding context
        const existingEncodingSection = existingModelCodes ? `\n\nEXISTING ENCODING (preserve unless specifically told to change):\n${existingModelCodes}` : "";
        const followupPrompt = `CONVERSATION HISTORY:\n${conversationHistoryString}CURRENT CODESTRINGS:\n${currentCodestrings}${existingEncodingSection}\n\nNEW USER MESSAGE:\n${clientprompt}`;

        if (DEBUG) {
            console.log("[handleFollowUpConversation] Followup prompt created:");
            console.log("System prompt:", followupSystemPrompt.substring(0, 100) + "...");
            console.log("User prompt preview:", followupPrompt.substring(0, 300) + "...");
        }

        // Make direct Claude API call (not using processPrompt to avoid any history manipulation)
        const messages = [
            { role: "system", content: followupSystemPrompt },
            { role: "user", content: followupPrompt }
        ];

        const claudeCallOptions = { 
            model: "claude-opus-4-1-20250805", 
            temperature: 1, 
            stream: false,
            useClaudeAPI: true,
            caller: "handleFollowUpConversation - Standalone"
        };

        let responseContent = "";
        for await (const contentPart of callOpenAI(messages, claudeCallOptions)) {
            responseContent += contentPart;
        }

        // Convert response to array format for consistency
        let responseArray;
        try {
            const parsed = JSON.parse(responseContent);
            if (Array.isArray(parsed)) {
                responseArray = parsed;
            } else {
                responseArray = responseContent.split('\n').filter(line => line.trim());
            }
        } catch (e) {
            responseArray = responseContent.split('\n').filter(line => line.trim());
        }

        // >>> ADDED: Console log the followup output
        console.log("\n╔════════════════════════════════════════════════════════════════╗");
        console.log("║                FOLLOWUP OUTPUT (STANDALONE)                    ║");
        console.log("╚════════════════════════════════════════════════════════════════╝");
        console.log("Followup Response Array:");
        console.log(responseArray);
        console.log("────────────────────────────────────────────────────────────────\n");

        // Update history (create new array, don't modify inplace)
        const updatedHistory = [
            ...currentHistory,
            ["human", clientprompt],
            ["assistant", responseArray.join("\n")] // Store response as single string
        ];

        // Persist updated history and analysis data (using localStorage helpers)
        saveConversationHistory(updatedHistory); // Save the new history state
        saveTrainingData(clientprompt, responseArray);

        if (DEBUG) console.log("Follow-up conversation processed (standalone). History length:", updatedHistory.length);

        // Return the response and the updated history
        return { response: responseArray, history: updatedHistory };

    } catch (error) {
        console.error("Error in standalone follow-up conversation:", error);
        // Return error message and unchanged history
        return {
            response: ["Error processing follow-up request: " + error.message],
            history: currentHistory || []
        };
    }
}


// Function: Handle Initial Conversation
export async function handleInitialConversation(clientprompt) {
    if (DEBUG) console.log("Processing initial question:", clientprompt);

    // >>> ADDED: Store the original clean client prompt at the very start
    originalClientPrompt = clientprompt;
    if (DEBUG) console.log("[handleInitialConversation] Stored original client prompt:", originalClientPrompt.substring(0, 100) + "...");

     // Ensure API keys are available
    if (!INTERNAL_API_KEYS.OPENAI_API_KEY) {
        throw new Error("OpenAI API key not initialized for initial conversation.");
    }

    // Load necessary prompts
    const systemPrompt = await getSystemPromptFromFile('Encoder_System');
    const mainPromptText = await getSystemPromptFromFile('Encoder_Main');

     if (!systemPrompt || !mainPromptText) {
         throw new Error("Failed to load required prompts for initial conversation.");
     }

    if (DEBUG) console.log("SYSTEM PROMPT: ", systemPrompt ? systemPrompt.substring(0,100) + "..." : "Not loaded");
    if (DEBUG) console.log("MAIN PROMPT: ", mainPromptText ? mainPromptText.substring(0,100) + "..." : "Not loaded");

    // Construct the prompt for the first call (no vector DB context yet)
    const initialCallPrompt = `Client request: ${clientprompt}\n` +
                           `Main Prompt: ${mainPromptText}`;

    // Call the LLM (processPrompt uses OpenAI key internally)
    // Use appropriate model and API based on encoder configuration
    let encoderModel = GPT_O3mini;
    let apiType = "Chat Completions API";
    
    if (ENCODER_API_TYPE === ENCODER_API_TYPES.RESPONSES) {
        encoderModel = GPT_O3;
        apiType = "Responses API";
    } else if (ENCODER_API_TYPE === ENCODER_API_TYPES.CLAUDE) {
        encoderModel = "claude-opus-4-1-20250805";
        apiType = "Claude API";
    }
    
    console.log(`🔧 Initial conversation using ${apiType} with model: ${encoderModel}`);
    console.log(`📊 Current encoder API type: ${ENCODER_API_TYPE}`);
    
    let outputArray = await processPrompt({
        userInput: initialCallPrompt,
        systemPrompt: systemPrompt,
        model: encoderModel,
        temperature: 1,
        history: [], // No history for initial call
        promptFiles: { system: 'Encoder_System', main: 'Encoder_Main' }
    });

    // >>> ADDED: Console log the main encoder output
    console.log("\n╔════════════════════════════════════════════════════════════════╗");
    console.log("║                    MAIN ENCODER OUTPUT                         ║");
    console.log("╚════════════════════════════════════════════════════════════════╝");
    console.log("Main Encoder Response Array:");
    console.log(outputArray);
    console.log("────────────────────────────────────────────────────────────────\n");

    // >>> ADDED: Run pipe correction before validation - COMMENTED OUT
    /*
    if (DEBUG) console.log("[handleInitialConversation] Running automatic pipe correction...");
    const { autoCorrectPipeCounts } = await import('./PipeValidation.js');
    const pipeResult = autoCorrectPipeCounts(outputArray);
    if (pipeResult.changesMade > 0) {
        console.log(`✅ Pipe correction: ${pipeResult.changesMade} codestrings corrected`);
        outputArray = pipeResult.correctedText;
    } else {
        console.log("✅ Pipe correction: No changes needed");
    }
    */

    // >>> ADDED: Run logic validation and correction mechanism (up to 3 passes total)
    if (DEBUG) console.log("[handleInitialConversation] Running logic validation and correction mechanism...");
    let currentPassNumber = 1;
    let maxPasses = 3;
    let validationComplete = false;
    
    while (currentPassNumber <= maxPasses && !validationComplete) {
        const codestringsForValidation = Array.isArray(outputArray) ? outputArray.join("\n") : String(outputArray);
        const retryResult = await validateLogicWithRetry(codestringsForValidation, currentPassNumber);
        
        if (retryResult.logicErrors.length === 0) {
            // No logic errors - validation passed
            validationComplete = true;
            if (DEBUG) console.log(`[handleInitialConversation] ✅ Logic validation passed on pass ${currentPassNumber}`);
        } else if (currentPassNumber >= maxPasses) {
            // Maximum passes reached - stop retrying
            validationComplete = true;
            if (DEBUG) console.log(`[handleInitialConversation] ⚠️ Logic validation completed after ${currentPassNumber} passes with ${retryResult.logicErrors.length} remaining errors`);
        } else {
            // Logic errors found and haven't reached max passes - retry with LogicCorrectorGPT
            if (DEBUG) console.log(`[handleInitialConversation] 🔄 Pass ${currentPassNumber} found ${retryResult.logicErrors.length} logic errors - retrying with LogicCorrectorGPT...`);
            
            // Run LogicCorrectorGPT to fix the remaining errors
            outputArray = await checkCodeStringsWithLogicCorrector(outputArray, retryResult.logicErrors);
            if (DEBUG) console.log(`[handleInitialConversation] LogicCorrectorGPT pass ${currentPassNumber} completed`);
            
            currentPassNumber++;
        }
    }
    
    if (DEBUG) console.log(`[handleInitialConversation] Logic validation and correction mechanism completed after ${currentPassNumber} pass(es)`);

    // >>> ADDED: Check for format errors and only call FormatGPT if errors exist
    if (DEBUG) console.log("[handleInitialConversation] Checking for format validation errors...");
    const codestringsForInitialFormatCheck = Array.isArray(outputArray) ? outputArray.join("\n") : String(outputArray);
    const initialFormatErrors = await getFormatErrorsForPrompt(codestringsForInitialFormatCheck);
    
    if (initialFormatErrors && initialFormatErrors.trim() !== "") {
        if (DEBUG) console.log("[handleInitialConversation] Format errors detected - calling FormatGPT...");
        outputArray = await formatCodeStringsWithGPT(outputArray);
        if (DEBUG) console.log("[handleInitialConversation] FormatGPT formatting completed");
        
        // >>> ADDED: Run format validation retry mechanism (up to 2 passes total) only after FormatGPT was called
        if (DEBUG) console.log("[handleInitialConversation] Running post-FormatGPT validation retry...");
        let formatCurrentPassNumber = 1;
        let formatMaxPasses = 2;
        let formatValidationComplete = false;
        
        while (formatCurrentPassNumber <= formatMaxPasses && !formatValidationComplete) {
            const codestringsForFormatValidation = Array.isArray(outputArray) ? outputArray.join("\n") : String(outputArray);
            const formatRetryResult = await validateFormatWithRetry(codestringsForFormatValidation, formatCurrentPassNumber);
            
            if (formatRetryResult.formatErrors.length === 0) {
                // No format errors - validation passed
                formatValidationComplete = true;
                if (DEBUG) console.log(`[handleInitialConversation] ✅ Format validation passed on pass ${formatCurrentPassNumber}`);
            } else if (formatCurrentPassNumber >= formatMaxPasses) {
                // Maximum passes reached - stop retrying
                formatValidationComplete = true;
                if (DEBUG) console.log(`[handleInitialConversation] ⚠️ Format validation completed after ${formatCurrentPassNumber} passes with ${formatRetryResult.formatErrors.length} remaining errors`);
            } else {
                // Format errors found and haven't reached max passes - retry with FormatGPT
                if (DEBUG) console.log(`[handleInitialConversation] 🔄 Pass ${formatCurrentPassNumber} found ${formatRetryResult.formatErrors.length} format errors - retrying with FormatGPT...`);
                
                // Run FormatGPT again to fix the remaining errors
                outputArray = await formatCodeStringsWithGPT(outputArray);
                if (DEBUG) console.log(`[handleInitialConversation] FormatGPT retry pass ${formatCurrentPassNumber + 1} completed`);
                
                formatCurrentPassNumber++;
            }
        }
        
        if (DEBUG) console.log(`[handleInitialConversation] Format validation retry mechanism completed after ${formatCurrentPassNumber} pass(es)`);
    } else {
        if (DEBUG) console.log("[handleInitialConversation] ✅ No format errors detected - skipping FormatGPT");
    }

    // >>> ADDED: Check labels using LabelCheckerGPT
    // if (DEBUG) console.log("[handleInitialConversation] Checking labels with LabelCheckerGPT...");
    // outputArray = await checkLabelsWithGPT(outputArray);
    // if (DEBUG) console.log("[handleInitialConversation] LabelCheckerGPT checking completed");

    // Create the initial history
    const initialHistory = [
        ["human", clientprompt],
        ["assistant", outputArray.join("\n")] // Store response as single string
    ];

    // Persist history and analysis data
    saveConversationHistory(initialHistory);
   
    saveTrainingData(clientprompt, outputArray);

    if (DEBUG) console.log("Initial conversation processed. History length:", initialHistory.length);
    if (DEBUG) console.log("Initial Response:", outputArray);

    // Return the response and the new history
    return { response: outputArray, history: initialHistory };
}

// Main conversation handler - decides between initial and follow-up
// Takes current history and returns { response, history }
export async function handleConversation(clientprompt, currentHistory) {
    // Reset validation pass counter for new conversation
    resetValidationPassCounter();
    
    try {
        const isFollowUp = currentHistory && currentHistory.length > 0;
        if (isFollowUp) {
            return await handleFollowUpConversation(clientprompt, currentHistory);
        } else {
            return await handleInitialConversation(clientprompt);
        }
    } catch (error) {
        console.error("Error in conversation handling:", error);
        // Return error message and unchanged history
        return {
            response: ["Error processing your request: " + error.message],
            history: currentHistory || [] // Return existing or empty history
        };
    }
}




// Function: Save training data pair to localStorage
function saveTrainingData(clientprompt, outputArray) {
    try {
        // Helper to clean text for storage
        function cleanText(text) {
            if (!text) return '';
            // Convert non-strings (like arrays) to string first
            const str = Array.isArray(text) ? JSON.stringify(text) : String(text);
            return str.replace(/[\r\n\t]+/g, ' ').trim(); // Replace newlines/tabs with space
        }

        const trainingData = {
            prompt: cleanText(clientprompt),
            // Ensure outputArray is stringified if it's an array
            response: cleanText(outputArray)
        };

        localStorage.setItem('trainingData', JSON.stringify(trainingData));
        if (DEBUG) console.log('Training data saved to localStorage');
    } catch (error) {
        console.error("Error saving training data:", error);
    }
}


// Function: Perform validation correction using LLM
// Note: Assumes localStorage contains relevant context from previous calls
export async function validationCorrection(clientprompt, initialResponse, validationResults) {
    try {
         // Ensure API keys are available
         if (!INTERNAL_API_KEYS.OPENAI_API_KEY) {
            throw new Error("OpenAI API key not initialized for validation correction.");
        }

        // Load validation prompts
        const validationSystemPrompt = await getSystemPromptFromFile('Validation_System');
        const validationMainPrompt = await getSystemPromptFromFile('Validation_Main');

        if (!validationSystemPrompt || !validationMainPrompt) {
            throw new Error("Failed to load validation system or main prompt");
        }

        // Format the initial response and validation results as strings
        const responseString = Array.isArray(initialResponse) ? initialResponse.join("\n") : String(initialResponse);
        const validationResultsString = Array.isArray(validationResults) ? validationResults.join("\n") : String(validationResults);

        // Construct the correction prompt using only essential information
        const correctionPrompt =
            `Main Prompt: ${validationMainPrompt}\n\n` +
            `Original User Input: ${clientprompt}\n\n` +
            `Initial Response (to be corrected): ${responseString}\n\n` +
            `Validation Errors Found: ${validationResultsString}`;


        if (DEBUG) {
            console.log("====== VALIDATION CORRECTION INPUT ======");
            console.log("System Prompt:", validationSystemPrompt.substring(0,100) + "...");
            console.log("User Input Prompt (truncated):", correctionPrompt.substring(0, 500) + "...");
            console.log("=========================================");
        }

        // Increment validation pass counter
        validationPassCounter++;
        
        // Call LLM for correction (processPrompt uses Claude API)
        // Pass an empty history, as correction likely doesn't need chat context
        const correctedResponseArray = await processPrompt({
            userInput: correctionPrompt,
            systemPrompt: validationSystemPrompt,
            model: "claude-opus-4-1-20250805",
            temperature: 0.7, // Lower temperature for correction
            history: [],
            promptFiles: { system: `Validation_System|PASS_${validationPassCounter}|ERRORS:${validationResultsString}` },
            useClaudeAPI: true
        });

        if (DEBUG) console.log("Validation correction completed successfully");
        if (DEBUG) console.log("Corrected Response:", correctedResponseArray);

        // >>> ADDED: Compare before and after codestrings for Validation Correction
        logCodestringComparison(initialResponse, correctedResponseArray, "Validation Correction");

        return correctedResponseArray; // Return the array format expected by caller

    } catch (error) {
        console.error("Error in validation correction:", error);
        // console.error(error.stack); // Keep stack trace for detailed debugging
        // Return an error message array, consistent with other function returns
        return ["Error during validation correction: " + error.message];
    }
}

// >>> ADDED: Global variable for AI Model Planner responses
let lastAIModelPlannerResponse = null; 

// >>> ADDED: Functions for Client Mode Chat
function displayInClientChat(message, isUser) {
    const chatLog = document.getElementById('chat-log-client');
    const welcomeMessage = document.getElementById('welcome-message-client');
    if (welcomeMessage) {
        welcomeMessage.style.display = 'none';
    }

    const messageElement = document.createElement('div');
    messageElement.className = `chat-message ${isUser ? 'user-message' : 'assistant-message'}`;
    
    const contentElement = document.createElement('p');
    contentElement.className = 'message-content';
    
    if (typeof message === 'string') {
        contentElement.textContent = message;
    } else if (Array.isArray(message)) { // Assuming array of strings for text
        contentElement.textContent = message.join('\n');
    } else if (typeof message === 'object' && message !== null) { // For JSON objects
        contentElement.textContent = JSON.stringify(message, null, 2);
        // Optionally, add a class or style for preformatted JSON
        contentElement.style.whiteSpace = 'pre-wrap'; 
    } else {
        contentElement.textContent = String(message); // Fallback
    }
    
    messageElement.appendChild(contentElement);
    chatLog.appendChild(messageElement);
    chatLog.scrollTop = chatLog.scrollHeight;
}

function setClientLoadingButton(isLoading) {
    const sendButton = document.getElementById('send-client');
    const loadingAnimation = document.getElementById('loading-animation-client');

    if (sendButton) {
        sendButton.disabled = isLoading;
    }
    if (loadingAnimation) {
        loadingAnimation.style.display = isLoading ? 'flex' : 'none';
    }
}

async function handleClientModeSend() {
    const userInputElement = document.getElementById('user-input-client');
    const userInput = userInputElement.value.trim();

    if (!userInput) {
        // Potentially show a message to the user, but for now, just log and return
        console.warn("Client mode: User input is empty.");
        return;
    }

    displayInClientChat(userInput, true);
    userInputElement.value = '';
    userInputElement.style.height = 'auto';
    userInputElement.classList.remove('scrollable');
    // Remove expanded class from input bar
    const inputBar = userInputElement.closest('.chatgpt-input-bar');
    inputBar?.classList.remove('expanded');
    // Trigger resize to ensure proper height
    const event = new Event('input', { bubbles: true });
    userInputElement.dispatchEvent(event);
    setClientLoadingButton(true);

    try {
        const result = await handleAIModelPlannerConversation(userInput);
        lastAIModelPlannerResponse = result.response; // Store the raw response

        // Display logic handles string, array, or object responses
        displayInClientChat(result.response, false);

    } catch (error) {
        console.error("Error in client mode conversation:", error);
        displayInClientChat(`Error: ${error.message}`, false);
    } finally {
        setClientLoadingButton(false);
    }
}

function handleClientModeResetChat() {
    const chatLog = document.getElementById('chat-log-client');
    const welcomeMessage = document.getElementById('welcome-message-client');
    
    chatLog.innerHTML = ''; // Clear existing messages
    if (welcomeMessage) {
        // Re-add welcome message or set its display to block
        const newWelcome = document.createElement('div');
        newWelcome.id = 'welcome-message-client';
        newWelcome.className = 'welcome-message';
        newWelcome.innerHTML = '<h1>Ask me anything (Client Mode)</h1>';
        chatLog.appendChild(newWelcome);
    }
    
    resetAIModelPlannerConversation(); // Reset the conversation history in the planner
    lastAIModelPlannerResponse = null;
    const userInputClient = document.getElementById('user-input-client');
    if (userInputClient) {
        userInputClient.value = '';
        userInputClient.style.height = 'auto'; // Reset height
        userInputClient.classList.remove('scrollable');
        // Remove expanded class from input bar
        const inputBar = userInputClient.closest('.chatgpt-input-bar');
        inputBar?.classList.remove('expanded');
        // Trigger resize to ensure proper height
        const event = new Event('input', { bubbles: true });
        userInputClient.dispatchEvent(event);
    }
    console.log("Client mode chat reset.");
}

// These functions are placeholders for client mode; actual Excel/editor interaction might differ or be disabled.
function handleClientModeWriteToExcel() {
    if (!lastAIModelPlannerResponse) {
        displayInClientChat("No response to write to Excel.", false);
        return;
    }
    // For now, just log it or display a message. 
    // Actual Excel writing might be complex if it's JSON.
    let contentToWrite = "";
    if (typeof lastAIModelPlannerResponse === 'object') {
        contentToWrite = JSON.stringify(lastAIModelPlannerResponse, null, 2);
    } else if (Array.isArray(lastAIModelPlannerResponse)) {
        contentToWrite = lastAIModelPlannerResponse.join("\n");
    } else {
        contentToWrite = String(lastAIModelPlannerResponse);
    }
    console.log("Client Mode - Write to Excel (Placeholder):\n", contentToWrite);
    displayInClientChat("Write to Excel (Placeholder): Response logged to console. Actual Excel writing depends on format.", false);
}

function handleClientModeInsertToEditor() {
    if (!lastAIModelPlannerResponse) {
        displayInClientChat("No response to insert into editor.", false);
        return;
    }
    let contentToInsert = "";
     if (typeof lastAIModelPlannerResponse === 'object') {
        contentToInsert = JSON.stringify(lastAIModelPlannerResponse, null, 2);
    } else if (Array.isArray(lastAIModelPlannerResponse)) {
        contentToInsert = lastAIModelPlannerResponse.join("\n");
    } else {
        contentToInsert = String(lastAIModelPlannerResponse);
    }
    console.log("Client Mode - Insert to Editor (Placeholder):\n", contentToInsert);
    displayInClientChat("Insert to Editor (Placeholder): Response logged to console. Actual editor insertion depends on context.", false);
}


Office.onReady(async (info) => {
  if (info.host === Office.HostType.Excel) {
    // Check if this is being loaded as a module by taskpane.js
    // If window.isTaskpaneMain is set, skip view management to prevent duplicate auth
    if (window.isTaskpaneMain) {
      console.log('[AIcalls.js] Loaded as module - skipping view management');
      return;
    }
    
    // ... existing setup for developer mode buttons, API keys, etc. ...
    
    // Initialize API Keys (calls setAIModelPlannerOpenApiKey inside)
    try {
        await initializeAPIKeys();
    } catch (error) {
        console.error("Failed to initialize API keys on startup:", error);
        // Potentially show error to user
    }

    // --- SETUP FOR CLIENT MODE UI ---
    // Note: send-client button handler is set in taskpane.js with credit enforcement wrapper

    const resetClientChatButton = document.getElementById('reset-chat-client');
    if (resetClientChatButton) resetClientChatButton.onclick = plannerHandleReset;
    
    const writeToExcelClientButton = document.getElementById('write-to-excel-client');
    if (writeToExcelClientButton) writeToExcelClientButton.onclick = plannerHandleWriteToExcel;

    const insertToEditorClientButton = document.getElementById('insert-to-editor-client');
    if (insertToEditorClientButton) insertToEditorClientButton.onclick = plannerHandleInsertToEditor;
    
    const userInputClient = document.getElementById('user-input-client');
    if (userInputClient) {
        userInputClient.addEventListener('keypress', function(event) {
            if (event.key === 'Enter' && !event.shiftKey) {
                event.preventDefault(); 
                if (plannerHandleSend) plannerHandleSend(); // Call the imported handler
            }
        });
    }
    // --- END CLIENT MODE UI SETUP ---

    // Load developer chat history
    // conversationHistory = loadConversationHistory(); // Assuming loadConversationHistory is for dev chat
    // ... display developer chat history ...

    // View references
    const appBody = document.getElementById('app-body');
    const clientModeView = document.getElementById('client-mode-view');
    const authenticationView = document.getElementById('authentication-view');

    function showDeveloperModeView() { // Renamed to avoid conflict if global
      if (appBody) appBody.style.display = 'flex'; 
      if (clientModeView) clientModeView.style.display = 'none';
      console.log("Developer Mode view activated");
    }

    async function showClientModeView() { // Renamed
      console.log("showClientModeView called");
      
      // Check if user is authenticated using multiple sources
      const googleUserSession = sessionStorage.getItem('googleUser');
      const googleCredentialSession = sessionStorage.getItem('googleCredential') || sessionStorage.getItem('googleToken');
      const googleTokenLocal = localStorage.getItem('google_access_token');
      const msalTokenLocal = localStorage.getItem('msal_access_token');
      const apiKeyLocal = localStorage.getItem('user_api_key');
      
      const isAuthenticated = (googleUserSession && googleCredentialSession) || 
                            !!(googleTokenLocal || msalTokenLocal || apiKeyLocal);
      
      if (DEBUG) {
        console.log(`Authentication check - Session: ${!!(googleUserSession && googleCredentialSession)}, Local: ${!!(googleTokenLocal || msalTokenLocal || apiKeyLocal)}`);
      }
      
      if (!isAuthenticated) {
        if (DEBUG) {
          console.log('User not authenticated, showing authentication view');
        }
        // Store that user wanted client mode after auth
        localStorage.setItem('post_auth_redirect', 'client-mode');
        showAuthenticationView();
        return;
      }
      
      if (DEBUG) {
        console.log('User is authenticated, proceeding to show client mode directly');
      }
      
      // Check if we need to authenticate with backend
      const hasBackendToken = sessionStorage.getItem('backend_access_token') || 
                             localStorage.getItem('backend_access_token');
      
      if (!hasBackendToken) {
        console.log('User has Google auth but no backend token - need to authenticate with backend');
        
        // Try to authenticate with backend using stored Google credentials
        const googleIdToken = sessionStorage.getItem('googleCredential') || 
                             localStorage.getItem('googleCredential');
        
        if (googleIdToken && window.backendAPI) {
          try {
            console.log('Authenticating with backend using stored Google token');
            const backendAuthResult = await window.backendAPI.signInWithGoogle(googleIdToken);
            console.log('Backend authentication successful');
          } catch (error) {
            console.error('Failed to authenticate with backend:', error);
            // If backend auth fails, show authentication page
            localStorage.setItem('post_auth_redirect', 'client-mode');
            showAuthenticationView();
            return;
          }
        } else {
          console.log('No Google ID token or backendAPI found - showing authentication page');
          // No ID token to authenticate with backend, show auth page
          localStorage.setItem('post_auth_redirect', 'client-mode');
          showAuthenticationView();
          return;
        }
      }
      
      if (appBody) appBody.style.display = 'none';
      if (clientModeView) clientModeView.style.display = 'flex';
      if (authenticationView) authenticationView.style.display = 'none';
      console.log("Client Mode view activated");
      
      // Update signed-in status when showing client mode
      if (typeof window.updateSignedInStatus === 'function') {
        window.updateSignedInStatus();
      }
    }

    function showAuthenticationView() { // Added authentication view function
      if (appBody) appBody.style.display = 'none';
      if (clientModeView) clientModeView.style.display = 'none';
      if (authenticationView) authenticationView.style.display = 'flex';
      
      // >>> REMOVED: Google Sign-In button handler to prevent duplicate authentication
      // The button handler is already set up in taskpane.js
      // const googleSignInButton = document.getElementById('google-signin-button');
      // if (googleSignInButton) {
      //   googleSignInButton.onclick = handleGoogleSignInView;
      // }
      
      if (DEBUG) {
        console.log("Authentication view activated");
      }
    }

    function handleGoogleSignInView() {
      if (DEBUG) {
        console.log("Google Sign-In clicked (AIcalls.js)");
      }
      
      // Use Office Dialog API for authentication within Excel environment
      if (typeof Office !== 'undefined' && Office.context && Office.context.ui) {
        const authUrl = buildGoogleAuthUrlView();
        
        Office.context.ui.displayDialogAsync(authUrl, {
          height: 60,
          width: 60,
          requireHTTPS: true
        }, function (result) {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            const dialog = result.value;
            
            // Listen for messages from the dialog (with error handling to prevent debug popup)
            try {
              dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
                try {
                  const authResult = JSON.parse(arg.message);
                  dialog.close();
                  
                  if (authResult.success) {
                    handleGoogleAuthSuccessView(authResult);
                  } else {
                    handleGoogleAuthErrorView(authResult.error);
                  }
                } catch (error) {
                  console.error("Error parsing auth result:", error);
                  dialog.close();
                  showError("Authentication failed. Please try again.");
                }
              });
              
              // Handle dialog closed by user (with error handling to prevent debug popup)  
              dialog.addEventHandler(Office.EventType.DialogEventReceived, function (arg) {
                if (arg.error === 12006) { // Dialog closed by user
                  console.log("Authentication canceled by user");
                } else {
                  console.error("Dialog error:", arg.error);
                  showError("Authentication failed. Please try again.");
                }
              });
            } catch (eventHandlerError) {
              console.warn("Could not register dialog event handlers:", eventHandlerError);
            }
          } else {
            console.error("Failed to open authentication dialog:", result.error);
            // Fallback to external window for development
            handleGoogleSignInFallbackView();
          }
        });
      } else {
        console.log("Office Dialog API not available, using fallback");
        handleGoogleSignInFallbackView();
      }
    }

    function buildGoogleAuthUrlView() {
      const GOOGLE_CLIENT_ID = "194102993575-6vkqrurasev2qr90nv82cbampqqq22hb.apps.googleusercontent.com";
      const params = new URLSearchParams({
        client_id: GOOGLE_CLIENT_ID,
        redirect_uri: `${window.location.origin}/auth/google/callback.html`,
        response_type: 'token id_token',
        scope: 'email profile',
        nonce: generateRandomStateView(),
        state: generateRandomStateView(),
        prompt: 'select_account' // Force account selection every time
      });
      
      return `https://accounts.google.com/o/oauth2/v2/auth?${params.toString()}`;
    }

    function generateRandomStateView() {
      return Math.random().toString(36).substring(2, 15) + Math.random().toString(36).substring(2, 15);
    }

    function handleGoogleSignInFallbackView() {
      // Fallback for development or when Office Dialog API is not available
      const authUrl = buildGoogleAuthUrlView();
      const popup = window.open(authUrl, 'googleAuth', 'width=500,height=600,scrollbars=yes,resizable=yes');
      
      // Poll for popup closure (not ideal but works for development)
      const pollTimer = setInterval(() => {
        try {
          if (popup.closed) {
            clearInterval(pollTimer);
            console.log("Authentication popup closed");
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

    function handleGoogleAuthSuccessView(authResult) {
      console.log("Google authentication successful (AIcalls.js)", authResult);
      
      // Store user session in both sessionStorage and localStorage for persistence
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
      
      // Show success message and redirect
      const userName = authResult.user ? authResult.user.name : 'User';
      showMessage(`Welcome ${userName}! Authentication successful.`);
      
      // Check for post-auth redirect
      const postAuthRedirect = localStorage.getItem('post_auth_redirect');
      if (postAuthRedirect === 'client-mode') {
        console.log('Post-auth redirect to client mode detected');
        localStorage.removeItem('post_auth_redirect');
      }
      
      showClientModeView();
    }

    function handleGoogleAuthErrorView(error) {
      console.error("Google authentication error (AIcalls.js):", error);
      console.error("Error details:", JSON.stringify(error, null, 2));
      
      let errorMessage = "Authentication failed: ";
      if (error && typeof error === 'object') {
        errorMessage += error.description || error.message || error.error || JSON.stringify(error);
      } else {
        errorMessage += error || 'Unknown error';
      }
      
      showError(errorMessage);
    }

    function handleGoogleCallbackView(response) {
      console.log("Google authentication successful (AIcalls.js)", response);
      
      // Decode the JWT token to get user info
      const userInfo = parseJwtView(response.credential);
      console.log("User info:", userInfo);
      
      // Store user session
      sessionStorage.setItem('googleUser', JSON.stringify(userInfo));
      sessionStorage.setItem('googleCredential', response.credential);
      
      // Show success message and redirect
      showMessage(`Welcome ${userInfo.name}! Authentication successful.`);
      showClientModeView();
    }

    function handleGoogleTokenResponseView(tokenResponse) {
      console.log("Google token response (AIcalls.js):", tokenResponse);
      
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
          
          // Show success message and redirect
          showMessage(`Welcome ${userInfo.name}! Authentication successful.`);
          showClientModeView();
        })
        .catch(error => {
          console.error("Error fetching user info:", error);
          showMessage("Authentication successful but failed to get user information.");
          showClientModeView();
        });
      }
    }

    function parseJwtView(token) {
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
    
    function showStartupMenuView() { // Redirects based on auth status
        // Check if user is authenticated
        const isAuthenticated = window.isUserAuthenticated && window.isUserAuthenticated();
        
        if (isAuthenticated) {
            console.log('[AIcalls.js] User authenticated - redirecting to client mode');
            showClientModeView();
        } else {
            console.log('[AIcalls.js] User not authenticated - redirecting to authentication view');
            showAuthenticationView();
        }
    }

    // Button handlers removed - startup menu and back buttons no longer exist
    
    document.getElementById("sideload-msg").style.display = "none";
    
    // Initial view is handled by taskpane.js based on authentication status
    console.log('[AIcalls.js] View initialization handled by main taskpane.js');

    // ... any other existing Office.onReady logic for developer mode ...
  }
});



// // >>> ADDED: Function to save the complete prompt sent to AI to lastprompt.txt
// async function saveEnhancedPrompt(fullPrompt) {
//     try {
//         console.log("Saving complete AI prompt to lastprompt.txt...");
        
//         // Try to save to server via POST request
//         const response = await fetch('https://localhost:3003/save-prompt', {
//             method: 'POST',
//             headers: {
//                 'Content-Type': 'application/json'
//             },
//             body: JSON.stringify({
//                 filename: 'src/prompts/lastprompt.txt',
//                 content: fullPrompt
//             })
//         });

//         if (response.ok) {
//             console.log("Complete AI prompt saved successfully to lastprompt.tx3
//         } else {
//             console.warn("Failed to save prompt to server:", response.statusText);
//             // Fallback: save to localStorage
//             localStorage.setItem('lastPrompt', fullPrompt);
//             console.log("Complete AI prompt saved to localStorage as fallback");
//         }
//     } catch (error) {
//         console.error("Error saving prompt:", error);
//         // Fallback: save to localStorage
//         try {
//             localStorage.setItem('lastPrompt', fullPrompt);
//             console.log("Complete AI prompt saved to localStorage as fallback");
//         } catch (storageError) {
//             console.error("Failed to save to localStorage:", storageError);
//         }
//     }
// }

// >>> ADDED: Helper function to check if "margin" is found in codestrings (case insensitive)
function containsMargin(codestrings) {
    if (!codestrings || typeof codestrings !== 'string') {
        if (DEBUG) console.log("[containsMargin] Invalid input:", typeof codestrings);
        return false;
    }
    
    // Convert to lowercase for case-insensitive search
    const lowercaseCodestrings = codestrings.toLowerCase();
    
    // Check for "margin" anywhere in the codestrings
    const hasMargin = lowercaseCodestrings.includes('margin');
    
    if (DEBUG) {
        console.log("[containsMargin] Has 'margin':", hasMargin);
        
        // Show examples of matches if found
        if (hasMargin) {
            const marginMatches = codestrings.match(/[^|]*margin[^|]*/gi);
            console.log("[containsMargin] 'margin' matches found:", marginMatches);
        }
    }
    
    return hasMargin;
}

// >>> ADDED: Helper function to check if "beg" or "end" words are found in codestrings (case insensitive)
function containsBeginningOrEnding(codestrings) {
    if (!codestrings || typeof codestrings !== 'string') {
        if (DEBUG) console.log("[containsBeginningOrEnding] Invalid input:", typeof codestrings);
        return false;
    }
    
    // Convert to lowercase for case-insensitive search
    const lowercaseCodestrings = codestrings.toLowerCase();
    
    // Check for "beg" or "end" anywhere in the codestrings
    const hasBeg = lowercaseCodestrings.includes('beg');
    const hasEnd = lowercaseCodestrings.includes('end');
    const result = hasBeg || hasEnd;
    
    if (DEBUG) {
        console.log("[containsBeginningOrEnding] Has 'beg':", hasBeg);
        console.log("[containsBeginningOrEnding] Has 'end':", hasEnd);
        console.log("[containsBeginningOrEnding] Final result:", result);
        
        // Show examples of matches if found
        if (hasBeg) {
            const begMatches = codestrings.match(/[^|]*beg[^|]*/gi);
            console.log("[containsBeginningOrEnding] 'beg' matches found:", begMatches);
        }
        if (hasEnd) {
            const endMatches = codestrings.match(/[^|]*end[^|]*/gi);
            console.log("[containsBeginningOrEnding] 'end' matches found:", endMatches);
        }
    }
    
    return result;
}

// >>> ADDED: Function to load MarginCOGS.txt content
async function loadMarginCOGSContent() {
    try {
        if (DEBUG) console.log("[loadMarginCOGSContent] Loading MarginCOGS.txt content...");
        
        const response = await fetch(CONFIG.getPromptUrl('MarginCOGS.txt'));
        if (DEBUG) console.log("[loadMarginCOGSContent] Fetch response status:", response.status, response.statusText);
        
        if (!response.ok) {
            throw new Error(`Failed to load MarginCOGS.txt: ${response.statusText}`);
        }
        
        const marginCOGSContent = await response.text();
        if (DEBUG) {
            console.log("[loadMarginCOGSContent] MarginCOGS.txt content loaded successfully, length:", marginCOGSContent.length);
            console.log("[loadMarginCOGSContent] Content preview:", marginCOGSContent.substring(0, 200) + "...");
        }
        
        return marginCOGSContent;
    } catch (error) {
        console.error("[loadMarginCOGSContent] Error loading MarginCOGS.txt:", error);
        // Return empty string on error to avoid breaking the flow
        return "";
    }
}

// >>> ADDED: Function to load ReconTable.txt content
async function loadReconTableContent() {
    try {
        if (DEBUG) console.log("[loadReconTableContent] Loading ReconTable.txt content...");
        
        const response = await fetch(CONFIG.getPromptUrl('ReconTable.txt'));
        if (DEBUG) console.log("[loadReconTableContent] Fetch response status:", response.status, response.statusText);
        
        if (!response.ok) {
            throw new Error(`Failed to load ReconTable.txt: ${response.statusText}`);
        }
        
        const reconTableContent = await response.text();
        if (DEBUG) {
            console.log("[loadReconTableContent] ReconTable.txt content loaded successfully, length:", reconTableContent.length);
            console.log("[loadReconTableContent] Content preview:", reconTableContent.substring(0, 200) + "...");
        }
        
        return reconTableContent;
    } catch (error) {
        console.error("[loadReconTableContent] Error loading ReconTable.txt:", error);
        // Return empty string on error to avoid breaking the flow
        return "";
    }
}

// >>> ADDED: LogicCorrectorGPT function to correct codestrings when logic errors are found
export async function checkCodeStringsWithLogicCorrector(responseArray, logicErrors) {
    if (DEBUG) console.log("[checkCodeStringsWithLogicCorrector] Processing codestrings for logic correction...");

    try {
        // Ensure API keys are available
        if (!INTERNAL_API_KEYS.OPENAI_API_KEY) {
            throw new Error("OpenAI API key not initialized for LogicCorrectorGPT correction.");
        }

        // Extract clean client request from the original prompt
        let cleanClientRequest = originalClientPrompt;
        
        // Remove "Client Request:" prefix if it exists
        cleanClientRequest = cleanClientRequest.replace(/^Client [Rr]equest:\s*/, '');
        
        // Find where training data or context starts and cut it off
        const contextStart = cleanClientRequest.search(/\n\n(Client request-specific Context|Training Data)/i);
        if (contextStart !== -1) {
            cleanClientRequest = cleanClientRequest.substring(0, contextStart).trim();
        }

        // Load the LogicCorrectorGPT system prompt
        const logicCorrectorSystemPrompt = await getSystemPromptFromFile('LogicCorrectorGPT');
        if (!logicCorrectorSystemPrompt) {
            throw new Error("Failed to load LogicCorrectorGPT system prompt");
        }

        // Convert responseArray to string for processing
        let codestringsInput = "";
        if (Array.isArray(responseArray)) {
            codestringsInput = responseArray.join("\n");
        } else {
            codestringsInput = String(responseArray);
        }

        // Format logic errors for the prompt
        let logicErrorsFormatted = "";
        if (logicErrors && logicErrors.length > 0) {
            logicErrorsFormatted = "\n\nValidation Errors Found:\n" + logicErrors.join("\n");
        }

        // Create the main message with the clean client request + codestrings + logic errors
        const logicCorrectorInput = `Original Client Request: ${cleanClientRequest}\n\nCodestrings to Correct:\n${codestringsInput}${logicErrorsFormatted}`;

        if (DEBUG) {
            console.log("[checkCodeStringsWithLogicCorrector] Final clean client request being sent:", cleanClientRequest.substring(0, 100) + "...");
            console.log("[checkCodeStringsWithLogicCorrector] Input codestrings:", codestringsInput.substring(0, 200) + "...");
            console.log("[checkCodeStringsWithLogicCorrector] Logic errors included:", logicErrors ? logicErrors.length : 0);
            console.log("[checkCodeStringsWithLogicCorrector] Using LogicCorrectorGPT system prompt");
        }

        // Call processPrompt with LogicCorrectorGPT system prompt
        const correctedResponseArray = await processPrompt({
            userInput: logicCorrectorInput,
            systemPrompt: logicCorrectorSystemPrompt,
            model: "claude-opus-4-1-20250805", // Using Claude model
            temperature: 0.3, // Lower temperature for logic correction consistency
            history: [], // No history needed for logic correction
            promptFiles: { system: 'LogicCorrectorGPT' },
            useClaudeAPI: true
        });

        if (DEBUG) {
            console.log("[checkCodeStringsWithLogicCorrector] LogicCorrectorGPT processing completed");
            console.log("[checkCodeStringsWithLogicCorrector] Corrected output:", correctedResponseArray);
        }

        // >>> ADDED: Compare before and after codestrings for LogicCorrectorGPT
        logCodestringComparison(responseArray, correctedResponseArray, "LogicCorrectorGPT");

        return correctedResponseArray;

    } catch (error) {
        console.error("[checkCodeStringsWithLogicCorrector] Error during LogicCorrectorGPT processing:", error);
        // Return original response on error to avoid breaking the flow
        console.warn("[checkCodeStringsWithLogicCorrector] Returning original response due to logic correction error");
        return responseArray;
    }
}

// >>> ADDED: LogicCheckerGPT function to check codestrings after encoder main step
export async function checkCodeStringsWithLogicChecker(responseArray) {
    if (DEBUG) console.log("[checkCodeStringsWithLogicChecker] Processing codestrings for logic checking...");

    try {
        // Ensure API keys are available
        if (!INTERNAL_API_KEYS.OPENAI_API_KEY) {
            throw new Error("OpenAI API key not initialized for LogicCheckerGPT checking.");
        }

        // >>> ADDED: Run logic validation BEFORE LogicGPT call
        // Convert responseArray to string for logic validation
        let codestringsForValidation = "";
        if (Array.isArray(responseArray)) {
            codestringsForValidation = responseArray.join("\n");
        } else {
            codestringsForValidation = String(responseArray);
        }
        
        console.log("Logic validation beginning for:", codestringsForValidation);
        
        // Get logic errors formatted for prompt inclusion
        const logicErrors = await getLogicErrorsForPrompt(codestringsForValidation);
        
        if (logicErrors && logicErrors.trim() !== "") {
            console.log("Validation errors:");
            // Extract just the error messages from the formatted output
            const errorLines = logicErrors.split('\n').filter(line => 
                line.trim() && 
                !line.includes('LOGIC VALIDATION ERRORS DETECTED') && 
                !line.includes('Please address the following') &&
                !line.includes('Please ensure your corrected')
            );
            errorLines.forEach(line => console.log(line));
        } else {
            console.log("No validation errors found");
        }

        // >>> SIMPLE FIX: Extract clean client request from whatever we have
        // Simple extraction: get everything after the first "Client Request:" or "Client request:" and before any training data
        let cleanClientRequest = originalClientPrompt;
        
        // Remove "Client Request:" prefix if it exists
        cleanClientRequest = cleanClientRequest.replace(/^Client [Rr]equest:\s*/, '');
        
        // Find where training data or context starts and cut it off
        const contextStart = cleanClientRequest.search(/\n\n(Client request-specific Context|Training Data)/i);
        if (contextStart !== -1) {
            cleanClientRequest = cleanClientRequest.substring(0, contextStart).trim();
        }

        // Load the LogicCheckerGPT system prompt
        const logicCheckerSystemPrompt = await getSystemPromptFromFile('LogicCheckerGPT');
        if (!logicCheckerSystemPrompt) {
            throw new Error("Failed to load LogicCheckerGPT system prompt");
        }

        // Convert responseArray to string for processing
        let codestringsInput = "";
        if (Array.isArray(responseArray)) {
            codestringsInput = responseArray.join("\n");
        } else {
            codestringsInput = String(responseArray);
        }

        // Create the main message with ONLY the clean client request + completed codestrings + logic errors (if any)
        let logicCheckerInput = `Original Client Request: ${cleanClientRequest}\n\nCompleted Codestrings to Check:\n${codestringsInput}`;
        
        // >>> ADDED: Check if ReconTable content should be appended
        if (DEBUG) {
            console.log("[checkCodeStringsWithLogicChecker] Starting ReconTable detection...");
            console.log("[checkCodeStringsWithLogicChecker] Codestrings sample:", codestringsInput.substring(0, 500) + "...");
        }
        
        const needsReconTableInfo = containsBeginningOrEnding(codestringsInput);
        
        if (DEBUG) {
            console.log("[checkCodeStringsWithLogicChecker] ReconTable detection result:", needsReconTableInfo);
        }
        
        if (needsReconTableInfo) {
            if (DEBUG) console.log("[checkCodeStringsWithLogicChecker] Detected 'beg' or 'end' - loading ReconTable content...");
            
            try {
                const reconTableContent = await loadReconTableContent();
                if (DEBUG) console.log("[checkCodeStringsWithLogicChecker] ReconTable content loaded, length:", reconTableContent ? reconTableContent.length : 0);
                
                if (reconTableContent && reconTableContent.trim() !== "") {
                    logicCheckerInput += `\n\nRECONCILIATION TABLE VALIDATION RULES:\n${reconTableContent}`;
                    if (DEBUG) console.log("[checkCodeStringsWithLogicChecker] ReconTable content appended to LogicCheckerGPT prompt");
                } else {
                    if (DEBUG) console.log("[checkCodeStringsWithLogicChecker] ReconTable content was empty or failed to load");
                }
            } catch (error) {
                console.error("[checkCodeStringsWithLogicChecker] Error loading ReconTable content:", error);
                // Continue without ReconTable content rather than failing
            }
        } else {
            if (DEBUG) console.log("[checkCodeStringsWithLogicChecker] No 'beg' or 'end' detected - ReconTable content not needed");
        }
        

        
        // >>> ADDED: Append logic errors to LogicGPT prompt if any were found
        if (logicErrors && logicErrors.trim() !== "") {
            logicCheckerInput += logicErrors; // logicErrors already includes proper formatting and line breaks
        }

        if (DEBUG) {
            console.log("[checkCodeStringsWithLogicChecker] Final clean client request being sent:", cleanClientRequest.substring(0, 100) + "...");
            console.log("[checkCodeStringsWithLogicChecker] Input codestrings:", codestringsInput.substring(0, 200) + "...");
            console.log("[checkCodeStringsWithLogicChecker] Logic errors included in prompt:", logicErrors ? "YES" : "NO");
            console.log("[checkCodeStringsWithLogicChecker] ReconTable content included in prompt:", needsReconTableInfo ? "YES" : "NO");
            console.log("[checkCodeStringsWithLogicChecker] Using LogicCheckerGPT system prompt");
        }

        // Call processPrompt with LogicCheckerGPT system prompt
        const checkedResponseArray = await processPrompt({
            userInput: logicCheckerInput,
            systemPrompt: logicCheckerSystemPrompt,
            model: "claude-opus-4-1-20250805", // Using Claude model
            temperature: 0.3, // Lower temperature for logic checking consistency
            history: [], // No history needed for logic checking
            promptFiles: { system: 'LogicCheckerGPT' },
            useClaudeAPI: true
        });

        if (DEBUG) {
            console.log("[checkCodeStringsWithLogicChecker] LogicCheckerGPT processing completed");
            console.log("[checkCodeStringsWithLogicChecker] Checked output:", checkedResponseArray);
        }

        // >>> ADDED: Compare before and after codestrings for LogicCheckerGPT
        logCodestringComparison(responseArray, checkedResponseArray, "LogicCheckerGPT");

        return checkedResponseArray;

    } catch (error) {
        console.error("[checkCodeStringsWithLogicChecker] Error during LogicCheckerGPT processing:", error);
        // Return original response on error to avoid breaking the flow
        console.warn("[checkCodeStringsWithLogicChecker] Returning original response due to logic checking error");
        return responseArray;
    }
}

// >>> ADDED: FormatGPT function to format codestrings after initial encoder call
export async function formatCodeStringsWithGPT(responseArray) {
    if (DEBUG) console.log("[formatCodeStringsWithGPT] Processing codestrings for formatting...");

    try {
        // Ensure API keys are available
        if (!INTERNAL_API_KEYS.OPENAI_API_KEY) {
            throw new Error("OpenAI API key not initialized for FormatGPT formatting.");
        }

        // >>> ADDED: Run format validation BEFORE FormatGPT call
        // Convert responseArray to string for format validation
        let codestringsForValidation = "";
        if (Array.isArray(responseArray)) {
            codestringsForValidation = responseArray.join("\n");
        } else {
            codestringsForValidation = String(responseArray);
        }
        
        console.log("Format validation beginning for:", codestringsForValidation);
        
        // Get format errors formatted for prompt inclusion
        const formatErrors = await getFormatErrorsForPrompt(codestringsForValidation);
        
        if (formatErrors && formatErrors.trim() !== "") {
            console.log("Format validation errors:");
            // Extract just the error messages from the formatted output
            const errorLines = formatErrors.split('\n').filter(line => 
                line.trim() && 
                !line.includes('FORMAT VALIDATION ERRORS DETECTED') && 
                !line.includes('Please address the following') &&
                !line.includes('Please ensure your corrected')
            );
            errorLines.forEach(line => console.log(line));
        } else {
            console.log("No format validation errors found");
        }

        // Load the FormatGPT system prompt
        const formatSystemPrompt = await getSystemPromptFromFile('FormatGPT');
        if (!formatSystemPrompt) {
            throw new Error("Failed to load FormatGPT system prompt");
        }

        // Convert responseArray to string for processing
        let codestringsInput = "";
        if (Array.isArray(responseArray)) {
            codestringsInput = responseArray.join("\n");
        } else {
            codestringsInput = String(responseArray);
        }

        // Create the main message with codestrings + format errors (if any)
        let formatInput = `Codestrings to Format:\n${codestringsInput}`;
        
        // >>> ADDED: Check if MarginCOGS content should be appended
        if (DEBUG) {
            console.log("[formatCodeStringsWithGPT] Starting MarginCOGS detection...");
        }
        
        const needsMarginCOGSInfo = containsMargin(codestringsInput);
        
        if (DEBUG) {
            console.log("[formatCodeStringsWithGPT] MarginCOGS detection result:", needsMarginCOGSInfo);
        }
        
        if (needsMarginCOGSInfo) {
            if (DEBUG) console.log("[formatCodeStringsWithGPT] Detected 'margin' - loading MarginCOGS content...");
            
            try {
                const marginCOGSContent = await loadMarginCOGSContent();
                if (DEBUG) console.log("[formatCodeStringsWithGPT] MarginCOGS content loaded, length:", marginCOGSContent ? marginCOGSContent.length : 0);
                
                if (marginCOGSContent && marginCOGSContent.trim() !== "") {
                    formatInput += `\n\nMARGIN & COGS VALIDATION RULES:\n${marginCOGSContent}`;
                    if (DEBUG) console.log("[formatCodeStringsWithGPT] MarginCOGS content appended to FormatGPT prompt");
                } else {
                    if (DEBUG) console.log("[formatCodeStringsWithGPT] MarginCOGS content was empty or failed to load");
                }
            } catch (error) {
                console.error("[formatCodeStringsWithGPT] Error loading MarginCOGS content:", error);
                // Continue without MarginCOGS content rather than failing
            }
        } else {
            if (DEBUG) console.log("[formatCodeStringsWithGPT] No 'margin' detected - MarginCOGS content not needed");
        }
        
        // >>> ADDED: Append format errors to FormatGPT prompt if any were found
        if (formatErrors && formatErrors.trim() !== "") {
            formatInput += formatErrors; // formatErrors already includes proper formatting and line breaks
        }

        if (DEBUG) {
            console.log("[formatCodeStringsWithGPT] Input codestrings:", codestringsInput.substring(0, 200) + "...");
            console.log("[formatCodeStringsWithGPT] Format errors included in prompt:", formatErrors ? "YES" : "NO");
            console.log("[formatCodeStringsWithGPT] MarginCOGS content included in prompt:", needsMarginCOGSInfo ? "YES" : "NO");
            console.log("[formatCodeStringsWithGPT] Using FormatGPT system prompt");
        }

        // Call processPrompt with FormatGPT system prompt and enhanced input
        const formattedResponseArray = await processPrompt({
            userInput: formatInput,
            systemPrompt: formatSystemPrompt,
            model: "claude-opus-4-1-20250805", // Using Claude model
            temperature: 0.3, // Lower temperature for formatting consistency
            history: [], // No history needed for formatting
            promptFiles: { system: 'FormatGPT' },
            useClaudeAPI: true
        });

        if (DEBUG) {
            console.log("[formatCodeStringsWithGPT] FormatGPT processing completed");
            console.log("[formatCodeStringsWithGPT] Formatted output:", formattedResponseArray);
        }

        // >>> ADDED: Compare before and after codestrings for FormatGPT
        logCodestringComparison(responseArray, formattedResponseArray, "FormatGPT");

        return formattedResponseArray;

    } catch (error) {
        console.error("[formatCodeStringsWithGPT] Error during FormatGPT processing:", error);
        // Return original response on error to avoid breaking the flow
        console.warn("[formatCodeStringsWithGPT] Returning original response due to formatting error");
        return responseArray;
    }
}

// >>> ADDED: LabelCheckerGPT function to check and correct labels after FormatGPT
export async function checkLabelsWithGPT(responseArray) {
    if (DEBUG) console.log("[checkLabelsWithGPT] Processing codestrings for label checking...");

    try {
        // Ensure API keys are available
        if (!INTERNAL_API_KEYS.OPENAI_API_KEY) {
            throw new Error("OpenAI API key not initialized for LabelCheckerGPT checking.");
        }

        // Extract clean client request from the original prompt
        let cleanClientRequest = originalClientPrompt;
        
        // Remove "Client Request:" prefix if it exists
        cleanClientRequest = cleanClientRequest.replace(/^Client [Rr]equest:\s*/, '');
        
        // Find where training data or context starts and cut it off
        const contextStart = cleanClientRequest.search(/\n\n(Client request-specific Context|Training Data)/i);
        if (contextStart !== -1) {
            cleanClientRequest = cleanClientRequest.substring(0, contextStart).trim();
        }

        // Load the LabelCheckerGPT system prompt
        const labelCheckerSystemPrompt = await getSystemPromptFromFile('LabelCheckerGPT');
        if (!labelCheckerSystemPrompt) {
            throw new Error("Failed to load LabelCheckerGPT system prompt");
        }

        // Convert responseArray to string for processing
        let codestringsInput = "";
        if (Array.isArray(responseArray)) {
            codestringsInput = responseArray.join("\n");
        } else {
            codestringsInput = String(responseArray);
        }

        // Create the message with ONLY the clean client request + codestrings (no training data or context)
        const labelCheckerInput = `Original Client Request: ${cleanClientRequest}\n\nCodestrings to Check Labels:\n${codestringsInput}`;

        if (DEBUG) {
            console.log("[checkLabelsWithGPT] Final clean client request being sent:", cleanClientRequest.substring(0, 100) + "...");
            console.log("[checkLabelsWithGPT] Input codestrings:", codestringsInput.substring(0, 200) + "...");
            console.log("[checkLabelsWithGPT] Using LabelCheckerGPT system prompt");
        }

        // Call processPrompt with LabelCheckerGPT system prompt
        const checkedResponseArray = await processPrompt({
            userInput: labelCheckerInput,
            systemPrompt: labelCheckerSystemPrompt,
            model: "claude-opus-4-1-20250805", // Using Claude model
            temperature: 0.3, // Lower temperature for label checking consistency
            history: [], // No history needed for label checking
            promptFiles: { system: 'LabelCheckerGPT' },
            useClaudeAPI: true
        });

        if (DEBUG) {
            console.log("[checkLabelsWithGPT] LabelCheckerGPT processing completed");
            console.log("[checkLabelsWithGPT] Label-checked output:", checkedResponseArray);
        }

        // >>> ADDED: Compare before and after codestrings for LabelCheckerGPT
        logCodestringComparison(responseArray, checkedResponseArray, "LabelCheckerGPT");

        return checkedResponseArray;

    } catch (error) {
        console.error("[checkLabelsWithGPT] Error during LabelCheckerGPT processing:", error);
        // Return original response on error to avoid breaking the flow
        console.warn("[checkLabelsWithGPT] Returning original response due to label checking error");
        return responseArray;
    }
}

// >>> ADDED: Function to determine which prompt modules to include using PromptModulesGPT
export async function determinePromptModules(clientRequest) {
    const startTime = performance.now();
    console.log("\n🔍 === PROMPT MODULE DETERMINATION STARTED ===");
    console.log(`📝 Client request length: ${clientRequest.length} characters`);
    console.log(`📝 Client request preview: ${clientRequest.substring(0, 150)}...`);
    console.log(`⏰ Start time: ${new Date().toISOString()}`);
    
    if (DEBUG) console.log("[determinePromptModules] Starting prompt module determination...");
    
    try {
        console.log("📋 Loading PromptModulesGPT system prompt...");
        
        // Load the PromptModulesGPT system prompt
        const promptModulesSystemPrompt = await getSystemPromptFromFile('PromptModulesGPT');
        if (!promptModulesSystemPrompt) {
            console.warn("⚠️ [determinePromptModules] Failed to load PromptModulesGPT system prompt, returning empty modules");
            console.log("❌ === PROMPT MODULE DETERMINATION FAILED ===\n");
            return [];
        }
        
        console.log(`✅ PromptModulesGPT system prompt loaded (${promptModulesSystemPrompt.length} chars)`);
        console.log("🚀 Calling Claude for prompt module analysis...");
        
        if (DEBUG) console.log("[determinePromptModules] Calling Claude to determine required modules...");
        
        const gptStartTime = performance.now();
        
        // Call Claude with the PromptModulesGPT system prompt
        const moduleResponse = await processPrompt({
            userInput: clientRequest,
            systemPrompt: promptModulesSystemPrompt,
            model: "claude-opus-4-1-20250805", // Using Claude model
            temperature: 0.3, // Lower temperature for more consistent module selection
            history: [],
            promptFiles: { system: 'PromptModulesGPT' },
            useClaudeAPI: true
        });
        
        const gptEndTime = performance.now();
        console.log(`⏱️ Claude call completed in ${(gptEndTime - gptStartTime).toFixed(2)}ms`);
        
        if (DEBUG) {
            console.log("[determinePromptModules] Raw response from PromptModulesGPT:");
            console.log(moduleResponse);
        }
        
        console.log("📊 Raw Claude response type:", Array.isArray(moduleResponse) ? "Array" : typeof moduleResponse);
        console.log("📊 Raw Claude response:", moduleResponse);
        console.log("📊 Input keywords analysis:", {
            hasLoan: clientRequest.toLowerCase().includes('loan'),
            hasDebt: clientRequest.toLowerCase().includes('debt'),
            hasBorrow: clientRequest.toLowerCase().includes('borrow'),
            hasLend: clientRequest.toLowerCase().includes('lend'),
            hasInterest: clientRequest.toLowerCase().includes('interest'),
            hasBalloon: clientRequest.toLowerCase().includes('balloon'),
            hasPayment: clientRequest.toLowerCase().includes('payment'),
            hasOriginate: clientRequest.toLowerCase().includes('originate')
        });
        
        // Parse the response - expecting an array of section labels
        let selectedModules = [];
        console.log("🔍 Parsing Claude response for module selection...");
        
        if (Array.isArray(moduleResponse)) {
            // If response is already an array, use it
            selectedModules = moduleResponse.filter(item => item && item.trim());
            console.log("✅ Response was already an array, filtered to:", selectedModules);
        } else if (typeof moduleResponse === 'string') {
            console.log("🔧 Response is string, attempting to parse...");
            try {
                // Try to parse as JSON array
                const parsed = JSON.parse(moduleResponse);
                if (Array.isArray(parsed)) {
                    selectedModules = parsed.filter(item => item && typeof item === 'string' && item.trim());
                    console.log("✅ Successfully parsed JSON array:", selectedModules);
                } else {
                    console.log("⚠️ Parsed JSON but not an array, falling back to string splitting");
                    // If not JSON, split by lines/commas and clean up
                    selectedModules = moduleResponse
                        .split(/[\n,]/)
                        .map(item => item.trim().replace(/["\[\]]/g, ''))
                        .filter(item => item && item.length > 0);
                    console.log("🔧 String split result:", selectedModules);
                }
            } catch (parseError) {
                console.log("⚠️ JSON parsing failed, using string splitting fallback");
                console.log("❌ Parse error:", parseError.message);
                // If JSON parsing fails, treat as comma/line separated string
                selectedModules = moduleResponse
                    .split(/[\n,]/)
                    .map(item => item.trim().replace(/["\[\]]/g, ''))
                    .filter(item => item && item.length > 0);
                console.log("🔧 Fallback string split result:", selectedModules);
            }
        }
        
        const endTime = performance.now();
        const totalTime = endTime - startTime;
        
        console.log("\n🎯 === PROMPT MODULE DETERMINATION COMPLETED ===");
        console.log(`📋 Selected modules: [${selectedModules.join(', ')}]`);
        console.log(`📊 Number of modules selected: ${selectedModules.length}`);
        console.log(`⏱️ Total execution time: ${totalTime.toFixed(2)}ms`);
        console.log(`✅ End time: ${new Date().toISOString()}`);
        
        // Log each selected module with details
        if (selectedModules.length > 0) {
            console.log("📝 Module selection details:");
            selectedModules.forEach((module, index) => {
                console.log(`  ${index + 1}. "${module}"`);
            });
        } else {
            console.log("📝 No modules were selected by Claude");
        }
        
        console.log("🏁 === PROMPT MODULE DETERMINATION END ===\n");
        
        if (DEBUG) {
            console.log("[determinePromptModules] Selected modules:", selectedModules);
        }
        
        return selectedModules;
        
    } catch (error) {
        const endTime = performance.now();
        const totalTime = endTime - startTime;
        
        console.error("❌ === PROMPT MODULE DETERMINATION ERROR ===");
        console.error(`💥 Error after ${totalTime.toFixed(2)}ms:`, error);
        console.error(`🔍 Error stack:`, error.stack);
        console.error("[determinePromptModules] Error determining prompt modules:", error);
        console.log("🔄 Returning empty array to continue main flow");
        console.log("❌ === PROMPT MODULE DETERMINATION END (ERROR) ===\n");
        
        // Return empty array on error to avoid breaking the main flow
        return [];
    }
}

// >>> ADDED: Function to load prompt module content based on selected modules
export async function loadSelectedPromptModules(selectedModules) {
    const startTime = performance.now();
    console.log("\n📚 === PROMPT MODULE LOADING STARTED ===");
    console.log(`📋 Modules to load: [${selectedModules.join(', ')}]`);
    console.log(`📊 Total modules: ${selectedModules.length}`);
    console.log(`⏰ Start time: ${new Date().toISOString()}`);
    
    if (DEBUG) console.log("[loadSelectedPromptModules] Loading content for selected modules:", selectedModules);
    
    const moduleContent = [];
    let loadedCount = 0;
    let failedCount = 0;
    let totalContentLength = 0;
    
    console.log("🔍 Using same loading pattern as working prompt functions");
    
    for (const module of selectedModules) {
        const moduleStartTime = performance.now();
        console.log(`\n📁 Loading module: "${module}"`);
        
        // Generate possible filename variants for the module
        const possibleFilenames = generateFilenameVariants(module);
        console.log(`🔗 Trying filename variants: [${possibleFilenames.join(', ')}]`);
        
        let fileLoaded = false;
        
        for (const fileName of possibleFilenames) {
            if (fileLoaded) break; // Skip remaining variants if we already found the file
            
            // Use the same path patterns as the working functions
                    const paths = [
            CONFIG.getPromptUrl(fileName), // Primary path (works for files copied to src/prompts/)
            CONFIG.getAssetUrl(`src/prompts/${fileName}`), // Direct src path (rarely works due to server config)
            CONFIG.getAssetUrl(`src/prompts/Prompt_Modules/${fileName}`), // Specific module subfolder (usually doesn't work)
            CONFIG.getPromptUrl(`Prompt_Modules/${fileName}`) // Alternative module subfolder (usually doesn't work)
        ];
            
            let response = null;
            for (const path of paths) {
                try {
                    if (DEBUG) console.log(`[loadSelectedPromptModules] Trying ${fileName} at ${path}...`);
                    console.log(`🌐 Fetching: ${path}`);
                    const fetchStartTime = performance.now();
                    
                    response = await fetch(path);
                    
                    const fetchEndTime = performance.now();
                    console.log(`📡 Fetch completed in ${(fetchEndTime - fetchStartTime).toFixed(2)}ms`);
                    console.log(`📊 Response status: ${response.status} ${response.statusText}`);
                    
                    if (response.ok) {
                        if (DEBUG) console.log(`[loadSelectedPromptModules] Successfully loaded from: ${path}`);
                        console.log(`✅ Successfully found ${fileName} at: ${path}`);
                        break;
                    } else if (response.status === 404) {
                        console.log(`⚠️ File not found at: ${path} (trying next path...)`);
                    } else {
                        console.warn(`❌ Failed to load from ${path}: ${response.status} ${response.statusText}`);
                    }
                } catch (err) {
                    console.log(`💥 Error fetching from ${path}: ${err.message} (trying next path...)`);
                    if (DEBUG) console.log(`[loadSelectedPromptModules] Error trying ${path}:`, err);
                }
            }
            
            if (response && response.ok) {
                try {
                    const content = await response.text();
                    const contentLength = content.length;
                    totalContentLength += contentLength;
                    
                    const formattedContent = `\n\n=== ${module.toUpperCase()} MODULE ===\n${content}`;
                    moduleContent.push(formattedContent);
                    loadedCount++;
                    fileLoaded = true;
                    
                    const moduleEndTime = performance.now();
                    console.log(`✅ Successfully loaded ${fileName}`);
                    console.log(`📏 Content length: ${contentLength} characters`);
                    console.log(`⏱️ Module load time: ${(moduleEndTime - moduleStartTime).toFixed(2)}ms`);
                    console.log(`📝 Content preview: ${content.substring(0, 100)}...`);
                    
                    if (DEBUG) console.log(`[loadSelectedPromptModules] Successfully loaded ${fileName} (${content.length} chars)`);
                    break; // Exit the filename variant loop since we found the file
                } catch (contentError) {
                    console.error(`❌ Error reading content from ${fileName}:`, contentError);
                    if (DEBUG) console.log(`[loadSelectedPromptModules] Content reading error:`, contentError);
                }
            }
        }
        
        if (!fileLoaded) {
            failedCount++;
            const moduleEndTime = performance.now();
            console.warn(`❌ Could not load module "${module}" - none of the filename variants worked`);
            console.warn(`⏱️ Failed after ${(moduleEndTime - moduleStartTime).toFixed(2)}ms`);
            console.warn(`📋 Tried variants: [${possibleFilenames.join(', ')}]`);
            console.warn(`[loadSelectedPromptModules] Unknown module: ${module}`);
        }
    }
    
    const endTime = performance.now();
    const totalTime = endTime - startTime;
    const combinedContent = moduleContent.join('');
    
    console.log("\n📊 === PROMPT MODULE LOADING SUMMARY ===");
    console.log(`✅ Successfully loaded: ${loadedCount} modules`);
    console.log(`❌ Failed to load: ${failedCount} modules`);
    console.log(`📏 Total content length: ${totalContentLength} characters`);
    console.log(`📄 Combined content length: ${combinedContent.length} characters`);
    console.log(`⏱️ Total loading time: ${totalTime.toFixed(2)}ms`);
    console.log(`📈 Average time per module: ${selectedModules.length > 0 ? (totalTime / selectedModules.length).toFixed(2) : 0}ms`);
    console.log(`✅ End time: ${new Date().toISOString()}`);
    
    if (combinedContent.length > 0) {
        console.log("📝 Combined content preview:");
        console.log(combinedContent.substring(0, 200) + "...");
    }
    
    console.log("🏁 === PROMPT MODULE LOADING END ===\n");
    
    if (DEBUG) {
        console.log(`[loadSelectedPromptModules] Loaded ${moduleContent.length} modules with total length: ${moduleContent.join('').length} chars`);
    }
    
    return combinedContent;
}

// >>> ADDED: Helper function to generate possible filename variants for a module
function generateFilenameVariants(moduleName) {
    const variants = [];
    
    // Convert module name to different filename formats
    const baseName = moduleName.trim();
    
    // 1. Exact match with .txt extension
    variants.push(`${baseName}.txt`);
    
    // 2. Special mapping for common module name variations
    if (baseName.toLowerCase() === 'debt') {
        variants.push('Debt.txt');
        variants.push('LendingBorrowing.txt');
    } else if (baseName.toLowerCase() === 'lendingborrowing') {
        variants.push('LendingBorrowing.txt');
        variants.push('Debt.txt');
    } else if (baseName.toLowerCase() === 'fixedassets') {
        variants.push('FixedAssets.txt');
    } else if (baseName.toLowerCase() === 'fixed assets') {
        variants.push('FixedAssets.txt');
    }
    
    // 3. Replace spaces with underscores
    if (baseName.includes(' ')) {
        variants.push(`${baseName.replace(/\s+/g, '_')}.txt`);
    }
    
    // 4. Remove spaces entirely (camelCase-like)
    if (baseName.includes(' ')) {
        const noSpaces = baseName.replace(/\s+/g, '');
        variants.push(`${noSpaces}.txt`);
    }
    
    // 5. PascalCase (capitalize first letter of each word, remove spaces)
    if (baseName.includes(' ')) {
        const pascalCase = baseName
            .split(/\s+/)
            .map(word => word.charAt(0).toUpperCase() + word.slice(1).toLowerCase())
            .join('');
        variants.push(`${pascalCase}.txt`);
    }
    
    // 6. Snake_case (lowercase with underscores)
    if (baseName.includes(' ')) {
        const snakeCase = baseName.toLowerCase().replace(/\s+/g, '_');
        variants.push(`${snakeCase}.txt`);
    }
    
    // 7. kebab-case (lowercase with hyphens)
    if (baseName.includes(' ')) {
        const kebabCase = baseName.toLowerCase().replace(/\s+/g, '-');
        variants.push(`${kebabCase}.txt`);
    }
    
    // 8. All lowercase
    const lowercase = baseName.toLowerCase();
    if (lowercase !== baseName) {
        variants.push(`${lowercase}.txt`);
    }
    
    // 9. All uppercase
    const uppercase = baseName.toUpperCase();
    if (uppercase !== baseName) {
        variants.push(`${uppercase}.txt`);
    }
    
    // Remove duplicates while preserving order
    return [...new Set(variants)];
}

export async function getAICallsProcessedResponse(userInputString, progressCallback = null) {
    const overallStartTime = performance.now();
    console.log("\n🚀 ============ AI PROCESSING PIPELINE STARTED ============");
    console.log(`📝 Input length: ${userInputString.length} characters`);
    console.log(`📝 Input preview: ${userInputString.substring(0, 100)}...`);
    console.log(`⏰ Pipeline start time: ${new Date().toISOString()}`);
    console.log("🔄 Initializing parallel processing...");
    
    if (DEBUG) console.log("[getAICallsProcessedResponse] Processing input:", userInputString.substring(0, 100) + "...");

    // >>> ADDED: Store the original clean client prompt at the very start (global variable)
    originalClientPrompt = userInputString;
    if (DEBUG) console.log("[getAICallsProcessedResponse] Stored original client prompt:", originalClientPrompt.substring(0, 100) + "...");

    // Reset validation pass counter for new processing
    resetValidationPassCounter();

    try {
        // >>> ADDED: Start async prompt module determination in parallel with structure processing
        console.log("\n⚡ === PARALLEL PROCESSING PHASE ===");
        console.log("🎯 Starting two parallel processes:");
        console.log("  1️⃣ Prompt Module Determination (GPT-4.1)");
        console.log("  2️⃣ Structure Database Queries");
        
        if (DEBUG) console.log("[getAICallsProcessedResponse] Starting parallel prompt module determination...");
        const promptModulesStartTime = performance.now();
        const promptModulesPromise = determinePromptModules(userInputString);
        console.log(`🚀 Prompt module determination started at ${new Date().toISOString()}`);

        // 1. Structure database queries (runs in parallel with prompt module determination)
        console.log("\n📊 Starting structure database queries...");
        if (DEBUG) console.log("[getAICallsProcessedResponse] Calling structureDatabasequeries...");
        const structureStartTime = performance.now();
        const dbResults = await structureDatabasequeries(userInputString, progressCallback);
        const structureEndTime = performance.now();
        
        console.log(`✅ Structure database queries completed in ${(structureEndTime - structureStartTime).toFixed(2)}ms`);
        if (DEBUG) console.log("[getAICallsProcessedResponse] structureDatabasequeries completed. Results:", dbResults);

        if (!dbResults || !Array.isArray(dbResults)) {
            console.error("❌ [getAICallsProcessedResponse] Invalid database results:", dbResults);
            throw new Error("Failed to get valid database results from structureDatabasequeries");
        }

        console.log(`📈 Database results: ${dbResults.length} result sets obtained`);

        // 2. Format database results into an enhanced prompt with consolidated training data and context
        console.log("\n🔄 Consolidating and deduplicating training data...");
        const consolidationStartTime = performance.now();
        const consolidatedData = consolidateAndDeduplicateTrainingData(dbResults);
        const consolidationEndTime = performance.now();

        console.log(`✅ Data consolidation completed in ${(consolidationEndTime - consolidationStartTime).toFixed(2)}ms`);
        console.log(`📏 Consolidated data length: ${consolidatedData.length} characters`);

        const enhancedPrompt = `Client request: ${userInputString}\n\n${consolidatedData}`;
        if (DEBUG) console.log("[getAICallsProcessedResponse] Enhanced prompt created:", enhancedPrompt.substring(0, 200) + "...");

        // 3. Load system and main prompts for AI call
        console.log("\n📋 Loading base prompts...");
        if (DEBUG) console.log("[getAICallsProcessedResponse] Loading system and main prompts for AI call...");
        const promptLoadStartTime = performance.now();
        
        const systemPrompt = await getSystemPromptFromFile('Encoder_System');
        const mainPromptText = await getSystemPromptFromFile('Encoder_Main');
        
        const promptLoadEndTime = performance.now();
        console.log(`✅ Base prompts loaded in ${(promptLoadEndTime - promptLoadStartTime).toFixed(2)}ms`);

        if (!systemPrompt || !mainPromptText) {
            throw new Error("[getAICallsProcessedResponse] Failed to load 'Encoder_System' or 'Encoder_Main' prompt.");
        }

        console.log(`📏 System prompt length: ${systemPrompt.length} characters`);
        console.log(`📏 Main prompt length: ${mainPromptText.length} characters`);

        // >>> ADDED: Wait for prompt module determination and load selected modules
        console.log("\n⏳ === WAITING FOR PROMPT MODULE DETERMINATION ===");
        if (DEBUG) console.log("[getAICallsProcessedResponse] Waiting for prompt module determination to complete...");
        
        const moduleWaitStartTime = performance.now();
        const selectedModules = await promptModulesPromise;
        const moduleWaitEndTime = performance.now();
        const promptModulesEndTime = performance.now();
        
        console.log(`✅ Prompt module determination completed in ${(promptModulesEndTime - promptModulesStartTime).toFixed(2)}ms`);
        console.log(`⏱️ Wait time for parallel process: ${(moduleWaitEndTime - moduleWaitStartTime).toFixed(2)}ms`);
        console.log(`📋 Selected modules: [${selectedModules.join(', ')}]`);
        
        let additionalModuleContent = "";
        if (selectedModules && selectedModules.length > 0) {
            console.log("\n📚 Loading selected prompt modules...");
            if (DEBUG) console.log("[getAICallsProcessedResponse] Loading selected prompt modules:", selectedModules);
            
            const moduleLoadStartTime = performance.now();
            additionalModuleContent = await loadSelectedPromptModules(selectedModules);
            const moduleLoadEndTime = performance.now();
            
            console.log(`✅ Module loading completed in ${(moduleLoadEndTime - moduleLoadStartTime).toFixed(2)}ms`);
            console.log(`📏 Additional module content length: ${additionalModuleContent.length} characters`);
        } else {
            console.log("📝 No additional prompt modules selected - using base prompt only");
            if (DEBUG) console.log("[getAICallsProcessedResponse] No additional prompt modules selected");
        }

        // >>> MODIFIED: Append selected prompt modules to the main prompt
        console.log("\n🔧 === PROMPT ENHANCEMENT PHASE ===");
        const enhancedMainPrompt = mainPromptText + additionalModuleContent;
        
        console.log(`📏 Base main prompt: ${mainPromptText.length} characters`);
        console.log(`📏 Additional modules: ${additionalModuleContent.length} characters`);
        console.log(`📏 Enhanced main prompt: ${enhancedMainPrompt.length} characters`);
        console.log(`📈 Prompt enhancement: +${additionalModuleContent.length} characters (${additionalModuleContent.length > 0 ? '+' : ''}${((additionalModuleContent.length / mainPromptText.length) * 100).toFixed(1)}%)`);
        
        if (DEBUG && additionalModuleContent) {
            console.log("[getAICallsProcessedResponse] Enhanced main prompt with modules, total length:", enhancedMainPrompt.length);
            console.log("[getAICallsProcessedResponse] Added modules content preview:", additionalModuleContent.substring(0, 300) + "...");
        }

        // Create final combined input
        const combinedInputForAI = `Client request: ${enhancedPrompt}\nMain Prompt: ${enhancedMainPrompt}`; // Using enhanced main prompt
        console.log(`📏 Final combined input length: ${combinedInputForAI.length} characters`);
        
        if (DEBUG) console.log("[getAICallsProcessedResponse] Calling processPrompt...");

        // Main AI processing call
        console.log("\n🤖 === MAIN AI PROCESSING CALL ===");
        console.log("🚀 Calling main encoder with enhanced prompt...");
        
        // Use appropriate model and API based on encoder configuration
        let encoderModel = GPT_O3mini;
        let apiType = "Chat Completions API";
        
        if (ENCODER_API_TYPE === ENCODER_API_TYPES.RESPONSES) {
            encoderModel = GPT_O3;
            apiType = "Responses API";
        } else if (ENCODER_API_TYPE === ENCODER_API_TYPES.CLAUDE) {
            encoderModel = "claude-opus-4-1-20250805";
            apiType = "Claude API";
        }
        
        console.log(`🔧 Using ${apiType} with model: ${encoderModel}`);
        console.log(`📊 Current encoder API type: ${ENCODER_API_TYPE}`);
        
        const mainAIStartTime = performance.now();
        
        let responseArray = await processPrompt({
            userInput: combinedInputForAI,
            systemPrompt: systemPrompt,
            model: encoderModel,
            temperature: 1, // Consistent temperature
            history: [], // Treat each call as independent for this processing
            promptFiles: { system: 'Encoder_System', main: 'Encoder_Main' }
        });
        
        const mainAIEndTime = performance.now();
        console.log(`✅ Main AI processing completed in ${(mainAIEndTime - mainAIStartTime).toFixed(2)}ms`);
        
        if (DEBUG) console.log("[getAICallsProcessedResponse] processPrompt completed. Response:", responseArray);

        // >>> ADDED: Console log the main encoder output
        console.log("\n╔════════════════════════════════════════════════════════════════╗");
        console.log("║              MAIN ENCODER OUTPUT (getAICallsProcessedResponse)  ║");
        console.log("╚════════════════════════════════════════════════════════════════╝");
        console.log("📊 Response type:", Array.isArray(responseArray) ? "Array" : typeof responseArray);
        console.log("📊 Response length:", Array.isArray(responseArray) ? responseArray.length : "N/A");
        console.log("Main Encoder Response Array:");
        console.log(responseArray);
        console.log("────────────────────────────────────────────────────────────────\n");

        // 4. COLUMN LABEL POST-PROCESSING PHASE (runs after main encoder, before validation)
        console.log("🏷️ === COLUMN LABEL POST-PROCESSING PHASE ===");
        if (DEBUG) console.log("[getAICallsProcessedResponse] Running column label post-processing...");
        const columnLabelStartTime = performance.now();
        
        try {
            // Import the post-processing function
            const { postProcessColumnLabels } = await import('./CodeCollection.js');
            
            // Convert responseArray to string for processing
            const responseString = Array.isArray(responseArray) ? responseArray.join("\n") : String(responseArray);
            
            // Apply column label post-processing
            const correctedResponseString = postProcessColumnLabels(responseString);
            
            // Convert back to array format to match expected format
            if (Array.isArray(responseArray)) {
                // Split back into array, handling both single and multiple code strings
                const codeStringPattern = /<[^>]+>/g;
                const correctedCodeStrings = correctedResponseString.match(codeStringPattern) || [];
                
                if (correctedCodeStrings.length > 0) {
                    responseArray = correctedCodeStrings;
                    console.log("✅ Column label post-processing completed - updated responseArray");
                } else {
                    console.log("⚠️ Column label post-processing: No code strings found in corrected output, keeping original");
                }
            } else {
                responseArray = correctedResponseString;
                console.log("✅ Column label post-processing completed - updated response string");
            }
            
            // Log the corrected output for comparison
            console.log("📊 Post-processed response length:", Array.isArray(responseArray) ? responseArray.length : "1 string");
            if (DEBUG) {
                console.log("Post-processed Response Array:");
                console.log(responseArray);
            }
            
        } catch (columnLabelError) {
            console.error("❌ Error during column label post-processing:", columnLabelError);
            console.log("⚠️ Continuing with original response due to post-processing error");
            // Continue with original responseArray if post-processing fails
        }
        
        const columnLabelEndTime = performance.now();
        console.log(`🏁 Column label post-processing completed in ${(columnLabelEndTime - columnLabelStartTime).toFixed(2)}ms`);

        // 6. PIPE CORRECTION PHASE (runs before validation) - COMMENTED OUT
        /*
        console.log("🔧 === PIPE CORRECTION PHASE ===");
        if (DEBUG) console.log("[getAICallsProcessedResponse] Running automatic pipe correction...");
        const pipeStartTime = performance.now();
        
        // Import pipe correction function
        const { autoCorrectPipeCounts } = await import('./PipeValidation.js');
        
        // Apply pipe corrections to the response array
        const pipeResult = autoCorrectPipeCounts(responseArray);
        if (pipeResult.changesMade > 0) {
            console.log(`✅ Pipe correction completed: ${pipeResult.changesMade} codestrings corrected`);
            responseArray = pipeResult.correctedText; // Update responseArray with corrected version
        } else {
            console.log("✅ Pipe correction completed: No changes needed");
        }
        
        const pipeEndTime = performance.now();
        console.log(`🏁 Pipe correction phase completed in ${(pipeEndTime - pipeStartTime).toFixed(2)}ms`);
        */

        // 6. Run logic validation and correction mechanism (up to 3 passes total)
        console.log("🔍 === LOGIC VALIDATION & CORRECTION PHASE ===");
        if (DEBUG) console.log("[getAICallsProcessedResponse] Running logic validation and correction mechanism...");
        let currentPassNumber = 1;
        let maxPasses = 3;
        let validationComplete = false;
        const validationStartTime = performance.now();
        
        while (currentPassNumber <= maxPasses && !validationComplete) {
            console.log(`\n🔍 Logic validation pass ${currentPassNumber}/${maxPasses}...`);
            const codestringsForValidation = Array.isArray(responseArray) ? responseArray.join("\n") : String(responseArray);
            const retryResult = await validateLogicWithRetry(codestringsForValidation, currentPassNumber);
            
            if (retryResult.logicErrors.length === 0) {
                // No logic errors - validation passed
                validationComplete = true;
                console.log(`✅ Logic validation passed on pass ${currentPassNumber}!`);
                if (DEBUG) console.log(`[getAICallsProcessedResponse] ✅ Logic validation passed on pass ${currentPassNumber}`);
            } else if (currentPassNumber >= maxPasses) {
                // Maximum passes reached - stop retrying
                validationComplete = true;
                console.log(`⚠️ Logic validation completed after ${currentPassNumber} passes with ${retryResult.logicErrors.length} remaining errors`);
                if (DEBUG) console.log(`[getAICallsProcessedResponse] ⚠️ Logic validation completed after ${currentPassNumber} passes with ${retryResult.logicErrors.length} remaining errors`);
            } else {
                // Logic errors found and haven't reached max passes - retry with LogicCorrectorGPT
                console.log(`🔄 Pass ${currentPassNumber} found ${retryResult.logicErrors.length} logic errors - retrying with LogicCorrectorGPT...`);
                if (DEBUG) console.log(`[getAICallsProcessedResponse] 🔄 Pass ${currentPassNumber} found ${retryResult.logicErrors.length} logic errors - retrying with LogicCorrectorGPT...`);
                
                // Run LogicCorrectorGPT to fix the remaining errors
                responseArray = await checkCodeStringsWithLogicCorrector(responseArray, retryResult.logicErrors);
                if (DEBUG) console.log(`[getAICallsProcessedResponse] LogicCorrectorGPT pass ${currentPassNumber} completed`);
                
                currentPassNumber++;
            }
        }
        
        const validationEndTime = performance.now();
        console.log(`🏁 Logic validation phase completed in ${(validationEndTime - validationStartTime).toFixed(2)}ms`);
        if (DEBUG) console.log(`[getAICallsProcessedResponse] Logic validation and correction mechanism completed after ${currentPassNumber} pass(es)`);

        // 7. Check for format errors and only call FormatGPT if errors exist
        console.log("\n✨ === FORMAT VALIDATION & CORRECTION PHASE ===");
        if (DEBUG) console.log("[getAICallsProcessedResponse] Checking for format validation errors...");
        const formatStartTime = performance.now();
        const codestringsForInitialFormatCheck = Array.isArray(responseArray) ? responseArray.join("\n") : String(responseArray);
        const initialFormatErrors = await getFormatErrorsForPrompt(codestringsForInitialFormatCheck);
        
        if (initialFormatErrors && initialFormatErrors.trim() !== "") {
            console.log("🔧 Format errors detected - calling FormatGPT...");
            if (DEBUG) console.log("[getAICallsProcessedResponse] Format errors detected - calling FormatGPT...");
            responseArray = await formatCodeStringsWithGPT(responseArray);
            if (DEBUG) console.log("[getAICallsProcessedResponse] FormatGPT formatting completed");

            // >>> ADDED: Run format validation retry mechanism (up to 2 passes total) only after FormatGPT was called
            if (DEBUG) console.log("[getAICallsProcessedResponse] Running post-FormatGPT validation retry...");
            let formatCurrentPassNumber = 1;
            let formatMaxPasses = 2;
            let formatValidationComplete = false;
            
            while (formatCurrentPassNumber <= formatMaxPasses && !formatValidationComplete) {
                console.log(`\n✨ Format validation retry pass ${formatCurrentPassNumber}/${formatMaxPasses}...`);
                const codestringsForFormatValidation = Array.isArray(responseArray) ? responseArray.join("\n") : String(responseArray);
                const formatRetryResult = await validateFormatWithRetry(codestringsForFormatValidation, formatCurrentPassNumber);
                
                if (formatRetryResult.formatErrors.length === 0) {
                    // No format errors - validation passed
                    formatValidationComplete = true;
                    console.log(`✅ Format validation passed on retry pass ${formatCurrentPassNumber}!`);
                    if (DEBUG) console.log(`[getAICallsProcessedResponse] ✅ Format validation passed on pass ${formatCurrentPassNumber}`);
                } else if (formatCurrentPassNumber >= formatMaxPasses) {
                    // Maximum passes reached - stop retrying
                    formatValidationComplete = true;
                    console.log(`⚠️ Format validation completed after ${formatCurrentPassNumber} retry passes with ${formatRetryResult.formatErrors.length} remaining errors`);
                    if (DEBUG) console.log(`[getAICallsProcessedResponse] ⚠️ Format validation completed after ${formatCurrentPassNumber} passes with ${formatRetryResult.formatErrors.length} remaining errors`);
                } else {
                    // Format errors found and haven't reached max passes - retry with FormatGPT
                    console.log(`🔄 Retry pass ${formatCurrentPassNumber} found ${formatRetryResult.formatErrors.length} format errors - retrying with FormatGPT...`);
                    if (DEBUG) console.log(`[getAICallsProcessedResponse] 🔄 Pass ${formatCurrentPassNumber} found ${formatRetryResult.formatErrors.length} format errors - retrying with FormatGPT...`);
                    
                    // Run FormatGPT again to fix the remaining errors
                    responseArray = await formatCodeStringsWithGPT(responseArray);
                    if (DEBUG) console.log(`[getAICallsProcessedResponse] FormatGPT retry pass ${formatCurrentPassNumber + 1} completed`);
                    
                    formatCurrentPassNumber++;
                }
            }
            
            if (DEBUG) console.log(`[getAICallsProcessedResponse] Format validation retry mechanism completed after ${formatCurrentPassNumber} pass(es)`);
        } else {
            console.log("✅ No format errors detected - skipping FormatGPT");
            if (DEBUG) console.log("[getAICallsProcessedResponse] ✅ No format errors detected - skipping FormatGPT");
        }
        
        const formatEndTime = performance.now();
        console.log(`🏁 Format validation phase completed in ${(formatEndTime - formatStartTime).toFixed(2)}ms`);

        // >>> ADDED: Check labels using LabelCheckerGPT
        // console.log("\n🏷️ === LABEL CHECKING PHASE ===");
        // if (DEBUG) console.log("[getAICallsProcessedResponse] Checking labels with LabelCheckerGPT...");
        // const labelStartTime = performance.now();
        // responseArray = await checkLabelsWithGPT(responseArray);
        // const labelEndTime = performance.now();
        // console.log(`✅ Label checking completed in ${(labelEndTime - labelStartTime).toFixed(2)}ms`);
        // if (DEBUG) console.log("[getAICallsProcessedResponse] LabelCheckerGPT checking completed");

        // >>> ADDED: Save the complete enhanced prompt that was sent to AI to lastprompt.txt
        // await saveEnhancedPrompt(combinedInputForAI);

        // Final summary
        const overallEndTime = performance.now();
        const totalProcessingTime = overallEndTime - overallStartTime;
        
        console.log("\n🎉 ============ AI PROCESSING PIPELINE COMPLETED ============");
        console.log(`⏱️ Total processing time: ${totalProcessingTime.toFixed(2)}ms`);
        console.log(`📊 Final response length: ${Array.isArray(responseArray) ? responseArray.length : 1} items`);
        console.log(`✅ Pipeline end time: ${new Date().toISOString()}`);
        console.log(`🎯 Modules used: ${selectedModules.length > 0 ? selectedModules.join(', ') : 'None'}`);
        console.log("🎉 ============ PROCESSING COMPLETE ============\n");

        return responseArray;

    } catch (error) {
        const overallEndTime = performance.now();
        const totalProcessingTime = overallEndTime - overallStartTime;
        
        console.error("💥 ============ AI PROCESSING PIPELINE ERROR ============");
        console.error(`❌ Error after ${totalProcessingTime.toFixed(2)}ms:`, error);
        console.error(`🔍 Error stack:`, error.stack);
        console.error("[getAICallsProcessedResponse] Error during processing:", error);
        console.log("🔄 Returning error message array");
        console.error("💥 ============ PROCESSING FAILED ============\n");
        
        // Return an error message array, consistent with other function returns
        return [`Error processing tab description: ${error.message}`];
    }
}

