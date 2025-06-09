/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

// Remove imports from Langchain to avoid ESM module issues
// Using direct fetch calls instead
// Add this test function
import { validateCodeStrings } from './Validation.js';
// Import the spreadsheet utilities
// import { handleInsertWorksheetsFromBase64 } from './SpreadsheetUtils.js';
// Import code collection functions
import { populateCodeCollection, exportCodeCollectionToText, runCodes, processAssumptionTabs, collapseGroupingsAndNavigateToFinancials, hideColumnsAndNavigate, handleInsertWorksheetsFromBase64 } from './CodeCollection.js';
// >>> ADDED: Import the new validation function
import { validateCodeStringsForRun, validateLogicWithRetry } from './Validation.js';
// >>> ADDED: Import the tab string generator function
import { generateTabString } from './IndexWorksheet.js';
// >>> ADDED: Import AIModelPlanner functions
import { handleAIModelPlannerConversation, resetAIModelPlannerConversation, setAIModelPlannerOpenApiKey, plannerHandleSend, plannerHandleReset, plannerHandleWriteToExcel, plannerHandleInsertToEditor } from './AIModelPlanner.js';
// >>> ADDED: Import logic validation functions  
import { getLogicErrorsForPrompt, getLogicErrors, hasLogicErrors } from './CodeCollection.js';
// >>> ADDED: Import format validation functions  
import { getFormatErrorsForPrompt, getFormatErrors, hasFormatErrors } from './CodeCollection.js';
// Add the codeStrings variable with the specified content
// REMOVED hardcoded codeStrings variable

import { API_KEYS as configApiKeys } from '../../config.js'; // Assuming config.js exports API_KEYS

// Mock fs module for browser environment (if needed within AIcalls)
const fs = {
    writeFileSync: (path, content) => {
        // console.log(`Mock writeFileSync called with path: ${path}`);
        // // In browser, we'll just log the content instead of writing to file
        // console.log(`Content would be written to ${path}:`, content.substring(0, 100) + '...');
    }
};

//*********Setup*********
// Start the timer
const startTime = performance.now();

//Debugging Toggle
const DEBUG = true;

// Variable to store loaded code strings
let loadedCodeStrings = "";

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

// API keys storage - initialized by initializeAPIKeys
let INTERNAL_API_KEYS = {
  OPENAI_API_KEY: "",
  PINECONE_API_KEY: ""
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
  }
}

const srcPaths = [
  'https://localhost:3002/src/prompts/Encoder_System.txt',
  'https://localhost:3002/src/prompts/Encoder_Main.txt',
  'https://localhost:3002/src/prompts/Followup_System.txt',
  'https://localhost:3002/src/prompts/Structure_System.txt',
  'https://localhost:3002/src/prompts/Validation_System.txt',
  'https://localhost:3002/src/prompts/Validation_Main.txt'
];

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
        setAIModelPlannerOpenApiKey(configApiKeys.OPENAI_API_KEY);
        console.log("OpenAI API key loaded from config.js and set for AI Model Planner");
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
                    setAIModelPlannerOpenApiKey(openaiKeyMatch[1]);
                    console.log("OpenAI API key loaded via fetch fallback and set for AI Model Planner.");
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
    call2context: {
        name: "call2context",
        apiEndpoint: "https://call2context-zmg9zog.svc.aped-4627-b74a.pinecone.io"
    },

};

//Models
const GPT4O_MINI = "gpt-4o-mini"
const GPT4O = "gpt-4o"
const GPT41 = "gpt-4.1"
const GPT45_TURBO = "gpt-4.5-turbo"
const GPT35_TURBO = "gpt-3.5-turbo"
const GPT4_TURBO = "gpt-4-turbo"
const GPTO3 = "gpt-o3"
const GPTFT1 =  "ft:gpt-4.1-2025-04-14:personal:jun25gpt4-1:BeyDTNt1"

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

// Direct OpenAI API call function
export async function* callOpenAI(messages, options = {}) {
  const { model = GPT41, temperature = 0.7, stream = false, caller = "Unknown" } = options;

  try {
    console.log(`Calling OpenAI API with model: ${model}, stream: ${stream}`);
    
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

    const body = {
      model: model,
      messages: messages,
      temperature: temperature
    };

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
          yield parsedLine;
        }
      }
    } else {
      const data = await response.json();
      console.log("OpenAI API response received (non-stream)");
      // For non-streaming, to maintain compatibility with handleSendClient's expectation of an iterable, 
      // we yield a single object that mimics the structure of a stream chunk if needed, 
      // or simply return the content if the caller adapts.
      // For now, let's assume the non-streaming path is not used by handleSendClient directly.
      // Returning the content directly as before for other potential callers.
      // If callOpenAI is *only* called by handleSendClient, this else block might need to yield as well.
      // However, processPrompt calls callOpenAI without expecting a stream.
      // If callOpenAI is *only* called by handleSendClient, this else block might need to yield as well.
      // However, processPrompt calls callOpenAI without expecting a stream.
      yield data.choices[0].message.content; // Yield the content string once
      return; // Explicitly end the generator
    }

  } catch (error) {
    console.error("Error calling OpenAI API:", error);
    // If it's a stream, we can't return, but the error will propagate.
    // If not a stream, re-throw as before.
    if (!stream) throw error;
    // For a stream, the error breaks the generator. Consider yielding an error object if preferred.
    // For now, just log and let the generator terminate.
    // yield { error: error.message }; // Optional: yield an error object
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

    return data.data[0].embedding;
  } catch (error) {
    console.error("Error creating embedding:", error);
    throw error;
  }
}

// Function to load prompts from files
export async function loadPromptFromFile(promptKey) {
  try {
    const paths = [
      `https://localhost:3002/prompts/${promptKey}.txt`,
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
export async function processPrompt({ userInput, systemPrompt, model, temperature, history = [], promptFiles = {} }) {
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
            model: GPT41,
            temperature: 1,
            history: [], // Explicitly empty
            promptFiles: { system: 'Structure_System' }
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

      for (let i = 0; i < queryStrings.length; i++) {
          const queryString = queryStrings[i];
          const chunkNumber = i + 1;
          
          // Call progress callback if provided
          if (progressCallback) {
              progressCallback(`Processing chunk ${chunkNumber} of ${queryStrings.length}: "${queryString.substring(0, 50)}${queryString.length > 50 ? '...' : ''}"`);
          }
          
          if (DEBUG) {
              console.log(`=== PROCESSING CHUNK ${chunkNumber}/${queryStrings.length} ===`);
              console.log(`Query: "${queryString}"`);
              console.log("Querying vector databases...");
          }
          
          try {
              // Make sure queryVectorDB uses the internal API keys
              const queryResults = {
                  query: queryString,
                  trainingData: await queryVectorDB({
                      queryPrompt: queryString,
                      similarityThreshold: .31,
                      indexName: 'call2trainingdata',
                      numResults: 30
                  }),
                  call2Context: await queryVectorDB({
                      queryPrompt: queryString,
                      similarityThreshold: .31,
                      indexName: 'call2context',
                      numResults: 20
                  }),

                //   codeOptions: await queryVectorDB({
                //       queryPrompt: queryString,
                //       indexName: 'codes',
                //       numResults: 3,
                //       similarityThreshold: .1
                //   })
              };

              results.push(queryResults);
              
              if (DEBUG) {
                  console.log(`Chunk ${chunkNumber} results summary:`);
                  console.log(`  - Training data: ${queryResults.trainingData.length} items`);
                  console.log(`  - Context: ${queryResults.call2Context.length} items`);
                //   console.log(`  - Call1 context: ${queryResults.call1Context.length} items`);
                //   console.log(`  - Code options: ${queryResults.codeOptions.length} items`);
                  console.log(`=== END CHUNK ${chunkNumber} ===`);
              }
              
              // Call progress callback for completion of this chunk
              if (progressCallback) {
                  progressCallback(`Completed chunk ${chunkNumber} of ${queryStrings.length}`);
              }
          } catch (error) {
              console.error(`Error processing query "${queryString}":`, error);
              // Continue with next query instead of failing completely
              if (progressCallback) {
                  progressCallback(`Error in chunk ${chunkNumber}: ${error.message}`);
              }
          }
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
                if (DEBUG) console.log(`[deduplicateTrainingDataByOutput] Could not extract output from entry ${entryIndex}, keeping it`);
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

        // Process context data
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

    // Format the consolidated training data
    formattedOutput += "Training Data:\n****\n";
    
    allTrainingData.forEach((trainingEntry, index) => {
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
        console.log(`  - Training data entries: ${allTrainingData.length} (total original: ${totalOriginalTrainingCount})`);
        console.log(`  - Context entries: ${allContextData.length} (total original: ${totalOriginalContextCount})`);
        console.log(`  - Training duplicates removed: ${trainingDuplicatesRemoved}`);
        console.log(`  - Context duplicates removed: ${contextDuplicatesRemoved}`);
        console.log(`  - Training entries with scores: ${allTrainingData.filter(item => item.score !== null).length}`);
        console.log(`  - Context entries with scores: ${allContextData.filter(item => item.score !== null).length}`);
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

        const response = await fetch(url, {
            method: 'POST',
            headers: {
                'api-key': INTERNAL_API_KEYS.PINECONE_API_KEY,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                vector: embedding,
                topK: numResults,
                includeMetadata: true,
                namespace: "ns1" // Assuming namespace is constant
            })
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

        // Load the followup system prompt
        const followupSystemPrompt = await getSystemPromptFromFile('Followup_System');
        if (!followupSystemPrompt) {
            throw new Error("Failed to load Followup_System prompt.");
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

        // Create the complete prompt for the followup
        const followupPrompt = `CONVERSATION HISTORY:\n${conversationHistoryString}CURRENT CODESTRINGS:\n${currentCodestrings}\n\nNEW USER MESSAGE:\n${clientprompt}`;

        if (DEBUG) {
            console.log("[handleFollowUpConversation] Followup prompt created:");
            console.log("System prompt:", followupSystemPrompt.substring(0, 100) + "...");
            console.log("User prompt preview:", followupPrompt.substring(0, 300) + "...");
        }

        // Make direct OpenAI call (not using processPrompt to avoid any history manipulation)
        const messages = [
            { role: "system", content: followupSystemPrompt },
            { role: "user", content: followupPrompt }
        ];

        const openaiCallOptions = { 
            model: GPT41, 
            temperature: 1, 
            stream: false,
            caller: "handleFollowUpConversation - Standalone"
        };

        let responseContent = "";
        for await (const contentPart of callOpenAI(messages, openaiCallOptions)) {
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
            let outputArray = await processPrompt({
            userInput: initialCallPrompt,
            systemPrompt: systemPrompt,
            model: GPT41,
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

    // >>> ADDED: Check the response using LogicCheckerGPT (uses global originalClientPrompt variable)
    if (DEBUG) console.log("[handleInitialConversation] Checking response with LogicCheckerGPT...");
    if (DEBUG) console.log("[handleInitialConversation] Using stored original client prompt for LogicCheckerGPT:", originalClientPrompt.substring(0, 100) + "...");
    outputArray = await checkCodeStringsWithLogicChecker(outputArray);
    if (DEBUG) console.log("[handleInitialConversation] LogicCheckerGPT checking completed");

    // >>> ADDED: Run logic validation retry mechanism (up to 2 passes total)
    if (DEBUG) console.log("[handleInitialConversation] Running post-LogicGPT validation retry...");
    let currentPassNumber = 1;
    let maxPasses = 2;
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
            // Logic errors found and haven't reached max passes - retry with LogicGPT
            if (DEBUG) console.log(`[handleInitialConversation] 🔄 Pass ${currentPassNumber} found ${retryResult.logicErrors.length} logic errors - retrying with LogicGPT...`);
            
            // Run LogicCheckerGPT again to fix the remaining errors
            outputArray = await checkCodeStringsWithLogicChecker(outputArray);
            if (DEBUG) console.log(`[handleInitialConversation] LogicCheckerGPT retry pass ${currentPassNumber + 1} completed`);
            
            currentPassNumber++;
        }
    }
    
    if (DEBUG) console.log(`[handleInitialConversation] Logic validation retry mechanism completed after ${currentPassNumber} pass(es)`);

    // >>> ADDED: Format the response using FormatGPT
    if (DEBUG) console.log("[handleInitialConversation] Formatting response with FormatGPT...");
    outputArray = await formatCodeStringsWithGPT(outputArray);
    if (DEBUG) console.log("[handleInitialConversation] FormatGPT formatting completed");

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
        
        // Call LLM for correction (processPrompt uses OpenAI key)
        // Pass an empty history, as correction likely doesn't need chat context
        const correctedResponseArray = await processPrompt({
            userInput: correctionPrompt,
            systemPrompt: validationSystemPrompt,
            model: GPT41,
            temperature: 0.7, // Lower temperature for correction
            history: [],
            promptFiles: { system: `Validation_System|PASS_${validationPassCounter}|ERRORS:${validationResultsString}` }
        });

        // Save the output using the mock fs (as per original logic)
        const correctionOutputPath = "C:\\Users\\joeor\\Dropbox\\B - Freelance\\C_Projectify\\VanPC\\Training Data\\Main Script Training and Context Data\\validation_correction_output.txt";
        const correctedResponseString = Array.isArray(correctedResponseArray) ? correctedResponseArray.join("\n") : correctedResponseArray;
        fs.writeFileSync(correctionOutputPath, correctedResponseString);

        if (DEBUG) console.log(`Validation correction output saved via mock fs to ${correctionOutputPath}`);
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
    document.getElementById('user-input-client').value = '';
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
    // ... existing setup for developer mode buttons, API keys, etc. ...
    
    // Initialize API Keys (calls setAIModelPlannerOpenApiKey inside)
    try {
        await initializeAPIKeys();
    } catch (error) {
        console.error("Failed to initialize API keys on startup:", error);
        // Potentially show error to user
    }

    // --- SETUP FOR CLIENT MODE UI ---
    // Assign event handlers from AIModelPlanner.js to client mode buttons
    const sendClientButton = document.getElementById('send-client');
    if (sendClientButton) sendClientButton.onclick = plannerHandleSend;

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

    // Startup Menu Logic (assuming this is still part of AIcalls.js)
    const startupMenu = document.getElementById('startup-menu');
    const developerModeButton = document.getElementById('developer-mode-button');
    const clientModeButton = document.getElementById('client-mode-button');
    const appBody = document.getElementById('app-body');
    const clientModeView = document.getElementById('client-mode-view');

    function showDeveloperModeView() { // Renamed to avoid conflict if global
      if (startupMenu) startupMenu.style.display = 'none';
      if (appBody) appBody.style.display = 'flex'; 
      if (clientModeView) clientModeView.style.display = 'none';
      console.log("Developer Mode view activated");
    }

    function showClientModeView() { // Renamed
      if (startupMenu) startupMenu.style.display = 'none';
      if (appBody) appBody.style.display = 'none';
      if (clientModeView) clientModeView.style.display = 'flex';
      console.log("Client Mode view activated");
    }
    
    function showStartupMenuView() { // Renamed
        if (startupMenu) startupMenu.style.display = 'flex';
        if (appBody) appBody.style.display = 'none';
        if (clientModeView) clientModeView.style.display = 'none';
        console.log("Startup Menu view activated");
    }

    if (developerModeButton) developerModeButton.onclick = showDeveloperModeView;
    if (clientModeButton) clientModeButton.onclick = showClientModeView;

    const backToMenuDevButton = document.getElementById('back-to-menu-dev-button');
    if (backToMenuDevButton) backToMenuDevButton.onclick = showStartupMenuView;
    const backToMenuClientButton = document.getElementById('back-to-menu-client-button');
    if (backToMenuClientButton) backToMenuClientButton.onclick = showStartupMenuView;
    
    document.getElementById("sideload-msg").style.display = "none";
    if (startupMenu) startupMenu.style.display = "flex"; // Show startup menu first
    if (appBody) appBody.style.display = "none";
    if (clientModeView) clientModeView.style.display = "none";

    // ... any other existing Office.onReady logic for developer mode ...
  }
});

// NEW FUNCTION to process text input like handleSend but without UI and main history side effects
export async function getAICallsProcessedResponse(userInputString, progressCallback = null) {
    if (DEBUG) console.log("[getAICallsProcessedResponse] Processing input:", userInputString.substring(0, 100) + "...");

    // >>> ADDED: Store the original clean client prompt at the very start (global variable)
    originalClientPrompt = userInputString;
    if (DEBUG) console.log("[getAICallsProcessedResponse] Stored original client prompt:", originalClientPrompt.substring(0, 100) + "...");

    // Reset validation pass counter for new processing
    resetValidationPassCounter();

    try {
        // 1. Structure database queries
        if (DEBUG) console.log("[getAICallsProcessedResponse] Calling structureDatabasequeries...");
        const dbResults = await structureDatabasequeries(userInputString, progressCallback);
        if (DEBUG) console.log("[getAICallsProcessedResponse] structureDatabasequeries completed. Results:", dbResults);

        if (!dbResults || !Array.isArray(dbResults)) {
            console.error("[getAICallsProcessedResponse] Invalid database results:", dbResults);
            throw new Error("Failed to get valid database results from structureDatabasequeries");
        }

        // 2. Format database results into an enhanced prompt with consolidated training data and context
        const consolidatedData = consolidateAndDeduplicateTrainingData(dbResults);

        const enhancedPrompt = `Client request: ${userInputString}\n\n${consolidatedData}`;
        if (DEBUG) console.log("[getAICallsProcessedResponse] Enhanced prompt created:", enhancedPrompt.substring(0, 200) + "...");

        // 3. Call the AI using processPrompt (to avoid main history side effects of handleConversation)
        if (DEBUG) console.log("[getAICallsProcessedResponse] Loading system and main prompts for AI call...");
        const systemPrompt = await getSystemPromptFromFile('Encoder_System');
        const mainPromptText = await getSystemPromptFromFile('Encoder_Main');

        if (!systemPrompt || !mainPromptText) {
            throw new Error("[getAICallsProcessedResponse] Failed to load 'Encoder_System' or 'Encoder_Main' prompt.");
        }
        
        const combinedInputForAI = `Client request: ${enhancedPrompt}\nMain Prompt: ${mainPromptText}`; // This matches how handleInitialConversation constructs it
        if (DEBUG) console.log("[getAICallsProcessedResponse] Calling processPrompt...");

        let responseArray = await processPrompt({
            userInput: combinedInputForAI,
            systemPrompt: systemPrompt,
            model: GPT41, // Using the same model as in other parts
            temperature: 1, // Consistent temperature
            history: [], // Treat each call as independent for this processing
            promptFiles: { system: 'Encoder_System', main: 'Encoder_Main' }
        });
        if (DEBUG) console.log("[getAICallsProcessedResponse] processPrompt completed. Response:", responseArray);

        // >>> ADDED: Console log the main encoder output
        console.log("\n╔════════════════════════════════════════════════════════════════╗");
        console.log("║              MAIN ENCODER OUTPUT (getAICallsProcessedResponse)  ║");
        console.log("╚════════════════════════════════════════════════════════════════╝");
        console.log("Main Encoder Response Array:");
        console.log(responseArray);
        console.log("────────────────────────────────────────────────────────────────\n");

        // 4. Check the response using LogicCheckerGPT (uses global originalClientPrompt variable)
        if (DEBUG) console.log("[getAICallsProcessedResponse] Checking response with LogicCheckerGPT...");
        if (DEBUG) console.log("[getAICallsProcessedResponse] Using stored original client prompt for LogicCheckerGPT:", originalClientPrompt.substring(0, 100) + "...");
        responseArray = await checkCodeStringsWithLogicChecker(responseArray);
        if (DEBUG) console.log("[getAICallsProcessedResponse] LogicCheckerGPT checking completed");

        // >>> ADDED: Run logic validation retry mechanism (up to 2 passes total)
        if (DEBUG) console.log("[getAICallsProcessedResponse] Running post-LogicGPT validation retry...");
        let currentPassNumber = 1;
        let maxPasses = 2;
        let validationComplete = false;
        
        while (currentPassNumber <= maxPasses && !validationComplete) {
            const codestringsForValidation = Array.isArray(responseArray) ? responseArray.join("\n") : String(responseArray);
            const retryResult = await validateLogicWithRetry(codestringsForValidation, currentPassNumber);
            
            if (retryResult.logicErrors.length === 0) {
                // No logic errors - validation passed
                validationComplete = true;
                if (DEBUG) console.log(`[getAICallsProcessedResponse] ✅ Logic validation passed on pass ${currentPassNumber}`);
            } else if (currentPassNumber >= maxPasses) {
                // Maximum passes reached - stop retrying
                validationComplete = true;
                if (DEBUG) console.log(`[getAICallsProcessedResponse] ⚠️ Logic validation completed after ${currentPassNumber} passes with ${retryResult.logicErrors.length} remaining errors`);
            } else {
                // Logic errors found and haven't reached max passes - retry with LogicGPT
                if (DEBUG) console.log(`[getAICallsProcessedResponse] 🔄 Pass ${currentPassNumber} found ${retryResult.logicErrors.length} logic errors - retrying with LogicGPT...`);
                
                // Run LogicCheckerGPT again to fix the remaining errors
                responseArray = await checkCodeStringsWithLogicChecker(responseArray);
                if (DEBUG) console.log(`[getAICallsProcessedResponse] LogicCheckerGPT retry pass ${currentPassNumber + 1} completed`);
                
                currentPassNumber++;
            }
        }
        
        if (DEBUG) console.log(`[getAICallsProcessedResponse] Logic validation retry mechanism completed after ${currentPassNumber} pass(es)`);

        // 5. Format the response using FormatGPT
        if (DEBUG) console.log("[getAICallsProcessedResponse] Formatting response with FormatGPT...");
        responseArray = await formatCodeStringsWithGPT(responseArray);
        if (DEBUG) console.log("[getAICallsProcessedResponse] FormatGPT formatting completed");

        // >>> ADDED: Save the complete enhanced prompt that was sent to AI to lastprompt.txt
        // await saveEnhancedPrompt(combinedInputForAI);

        return responseArray;

    } catch (error) {
        console.error("[getAICallsProcessedResponse] Error during processing:", error);
        // Return an error message array, consistent with other function returns
        return [`Error processing tab description: ${error.message}`];
    }
}

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
//             console.log("Complete AI prompt saved successfully to lastprompt.txt");
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
        
        const response = await fetch('https://localhost:3002/prompts/MarginCOGS.txt');
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
        
        const response = await fetch('https://localhost:3002/prompts/ReconTable.txt');
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
            model: GPT41, // Using same model as other calls
            temperature: 0.3, // Lower temperature for logic checking consistency
            history: [], // No history needed for logic checking
            promptFiles: { system: 'LogicCheckerGPT' }
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
            model: GPT41, // Using same model as other calls
            temperature: 0.3, // Lower temperature for formatting consistency
            history: [], // No history needed for formatting
            promptFiles: { system: 'FormatGPT' }
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

