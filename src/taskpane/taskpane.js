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
import { validateCodeStringsForRun } from './Validation.js';
// >>> ADDED: Import the tab string generator function
import { generateTabString } from './IndexWorksheet.js';
// >>> UPDATED: Import structureDatabasequeries from the helper file
import { structureDatabasequeries } from './StructureHelper.js';
// >>> ADDED: Import setAPIKeys function from AIcalls
import { setAPIKeys } from './AIcalls.js';
// Add the codeStrings variable with the specified content
// REMOVED hardcoded codeStrings variable

import { API_KEYS as configApiKeys } from '../../config.js'; // Assuming config.js exports API_KEYS

// Mock fs module for browser environment (if needed within AIcalls)
const fs = {
    writeFileSync: (path, content) => {
        console.log(`Mock writeFileSync called with path: ${path}`);
        // In browser, we'll just log the content instead of writing to file
        console.log(`Content would be written to ${path}:`, content.substring(0, 100) + '...');
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

// Update Pinecone configuration to handle multiple indexes
const PINECONE_ENVIRONMENT = "gcp-starter";

// Define configurations for each index
const PINECONE_INDEXES = {
    codes: {
        name: "codes",
        apiEndpoint: "https://codes-zmg9zog.svc.aped-4627-b74a.pinecone.io"
    },
    call2trainingdata: {
        name: "call2trainingdata",
        apiEndpoint: "https://call2trainingdata-zmg9zog.svc.aped-4627-b74a.pinecone.io"
    },
    call2context: {
        name: "call2context",
        apiEndpoint: "https://call2context-zmg9zog.svc.aped-4627-b74a.pinecone.io"
    },
    call1context: {
        name: "call1context",
        apiEndpoint: "https://call1context-zmg9zog.svc.aped-4627-b74a.pinecone.io"
    }
};

//Models
const GPT4O_MINI = "gpt-4o-mini"
const GPT4O = "gpt-4o"
const GPT41 = "gpt-4.1"
const GPT45_TURBO = "gpt-4.5-turbo"
const GPT35_TURBO = "gpt-3.5-turbo"
const GPT4_TURBO = "gpt-4-turbo"
const GPTFT1 =  "ft:gpt-3.5-turbo-1106:orsi-advisors:cohcolumnsgpt35:B6Wlrql1"

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
export async function callOpenAI(messages, model = GPT41, temperature = 0.7) {
  try {
    console.log(`Calling OpenAI API with model: ${model}`);

    if (!INTERNAL_API_KEYS.OPENAI_API_KEY) {
      throw new Error("OpenAI API key not found. Please check your API keys.");
    }

    const response = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${INTERNAL_API_KEYS.OPENAI_API_KEY}`
      },
      body: JSON.stringify({
        model: model,
        messages: messages,
        temperature: temperature
      })
    });

    if (!response.ok) {
      const errorData = await response.json().catch(() => null);
      console.error("OpenAI API error response:", errorData);
      throw new Error(`OpenAI API error: ${response.status} ${response.statusText}`);
    }

    const data = await response.json();
    console.log("OpenAI API response received");

    return data.choices[0].message.content;
  } catch (error) {
    console.error("Error calling OpenAI API:", error);
    throw error;
  }
}

// OpenAI embeddings function
export async function createEmbedding(text) {
  try {
    console.log("Creating embedding for text");

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
async function loadPromptFromFile(promptKey) {
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
async function getSystemPromptFromFile(promptKey) {
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
async function processPrompt({ userInput, systemPrompt, model, temperature, history = [] }) {
    if (DEBUG) console.log("API Key being used for processPrompt:", INTERNAL_API_KEYS.OPENAI_API_KEY ? `${INTERNAL_API_KEYS.OPENAI_API_KEY.substring(0, 3)}...` : "None");

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
        const responseContent = await callOpenAI(messages, model, temperature);

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
// >>> REMOVED: structureDatabasequeries function definition <<<
// export async function structureDatabasequeries(clientprompt) {
//   ... function body ...
// }

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

        // Extract text using the helper function
        const cleanMatches = matches.map(match => extractTextFromJson(match)).filter(text => text !== "");

        if (DEBUG) {
            console.log(`Found ${cleanMatches.length} matches (after threshold/limit/extraction):`);
            cleanMatches.forEach((text, i) => console.log(`  ${i + 1}: ${text.substring(0, 100)}...`));
        }

        return cleanMatches;

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


// Helper function to format JSON for prompts (handle potential errors)
function safeJsonForPrompt(obj, readable = true) {
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


// Function: Handle Follow-Up Conversation
async function handleFollowUpConversation(clientprompt, currentHistory) {
    if (DEBUG) console.log("Processing follow-up question:", clientprompt);
    if (DEBUG) console.log("Using conversation history length:", currentHistory.length);

    // Ensure API keys are available
    if (!INTERNAL_API_KEYS.OPENAI_API_KEY || !INTERNAL_API_KEYS.PINECONE_API_KEY) {
        throw new Error("API keys not initialized for follow-up conversation.");
    }

    // Load necessary prompts
    const systemPrompt = await getSystemPromptFromFile('Followup_System');
    const mainPromptText = await getSystemPromptFromFile('Encoder_Main'); // Assuming this is the 'MainPrompt' context needed

     if (!systemPrompt || !mainPromptText) {
         throw new Error("Failed to load required prompts for follow-up.");
     }

    // Fetch context using vector DB queries
    // These calls internally use createEmbedding (OpenAI key) and query (Pinecone key)
    const trainingdataCall2 = await queryVectorDB({
        queryPrompt: clientprompt,
        similarityThreshold: .4,
        indexName: 'call2trainingdata',
        numResults: 3
    });

    const call2context = await queryVectorDB({
        queryPrompt: clientprompt + safeJsonForPrompt(trainingdataCall2, false), // Append context for better query
        similarityThreshold: .3,
        indexName: 'call2context',
        numResults: 5
    });

    const call1context = await queryVectorDB({
        queryPrompt: clientprompt + safeJsonForPrompt(trainingdataCall2, false),
        similarityThreshold: .3,
        indexName: 'call1context',
        numResults: 5
    });

    const codeOptions = await queryVectorDB({
        queryPrompt: clientprompt + safeJsonForPrompt(trainingdataCall2, false) + safeJsonForPrompt(call1context, false),
        indexName: 'codes',
        numResults: 10,
        similarityThreshold: .1
    });

    // Construct the prompt for the LLM
    const followUpPrompt = `Client request: ${clientprompt}\n` +
                   `Main Prompt Context: ${mainPromptText}\n` + // Use loaded main prompt text
                   `Training Data Context: ${safeJsonForPrompt(trainingdataCall2, true)}\n` + // Use readable format for prompt
                   `Code Choosing Context: ${safeJsonForPrompt(call1context, true)}\n` +
                   `Code Editing Context: ${safeJsonForPrompt(call2context, true)}\n` +
                   `Relevant Code Options: ${safeJsonForPrompt(codeOptions, true)}`;

    // Call the LLM (processPrompt uses OpenAI key internally)
    const responseArray = await processPrompt({
        userInput: followUpPrompt,
        systemPrompt: systemPrompt,
        model: GPT41,
        temperature: 1,
        history: currentHistory // Pass the existing history
    });

    // Update history (create new array, don't modify inplace)
    const updatedHistory = [
        ...currentHistory,
        ["human", clientprompt],
        ["assistant", responseArray.join("\n")] // Store response as single string
    ];

    // Persist updated history and analysis data (using localStorage helpers)
    saveConversationHistory(updatedHistory); // Save the new history state

   

    if (DEBUG) console.log("Follow-up conversation processed. History length:", updatedHistory.length);

    // Return the response and the updated history
    return { response: responseArray, history: updatedHistory };
}


// Function: Handle Initial Conversation
async function handleInitialConversation(clientprompt) {
    if (DEBUG) console.log("Processing initial question:", clientprompt);

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
    const outputArray = await processPrompt({
        userInput: initialCallPrompt,
        systemPrompt: systemPrompt,
        model: GPT41,
        temperature: 1,
        history: [] // No history for initial call
    });

    // Create the initial history
    const initialHistory = [
        ["human", clientprompt],
        ["assistant", outputArray.join("\n")] // Store response as single string
    ];

    // Persist history and analysis data
    saveConversationHistory(initialHistory);



    if (DEBUG) console.log("Initial conversation processed. History length:", initialHistory.length);
    if (DEBUG) console.log("Initial Response:", outputArray);

    // Return the response and the new history
    return { response: outputArray, history: initialHistory };
}

// Main conversation handler - decides between initial and follow-up
// Takes current history and returns { response, history }
export async function handleConversation(clientprompt, currentHistory) {
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




// Function: Perform validation correction using LLM
// Note: Assumes localStorage contains relevant context from previous calls
export async function validationCorrection(clientprompt, initialResponse, validationResults) {
    try {
         // Ensure API keys are available
         if (!INTERNAL_API_KEYS.OPENAI_API_KEY) {
            throw new Error("OpenAI API key not initialized for validation correction.");
        }

        // Load context from localStorage (as per original logic)
        // Consider passing these as arguments if localStorage access becomes problematic
        const trainingData = localStorage.getItem('trainingData') || '{"prompt":"","response":""}'; // Provide default structure
        const promptAnalysisData = JSON.parse(localStorage.getItem('promptAnalysis') || '{}');

        // Load validation prompts
        const validationSystemPrompt = await getSystemPromptFromFile('Validation_System');
        const validationMainPrompt = await getSystemPromptFromFile('Validation_Main');

        if (!validationSystemPrompt || !validationMainPrompt) {
            throw new Error("Failed to load validation system or main prompt");
        }

        // Format the initial response and validation results as strings
        const responseString = Array.isArray(initialResponse) ? initialResponse.join("\n") : String(initialResponse);
        const validationResultsString = Array.isArray(validationResults) ? validationResults.join("\n") : String(validationResults);

        // Construct the correction prompt using loaded context
        const correctionPrompt =
            `Main Prompt: ${validationMainPrompt}\n\n` +
            `Original User Input: ${clientprompt}\n\n` +
            `Initial Response (to be corrected): ${responseString}\n\n` +
            `Validation Errors Found: ${validationResultsString}\n\n` +
            // Include context from the last analysis if available
            `Training Data Example: ${trainingData}\n\n` + // Use loaded training data string
            `Code Options Context: ${promptAnalysisData.codeOptions || "Not available"}\n\n` +
            `Code Choosing Context: ${promptAnalysisData.call1context || "Not available"}\n\n` +
            `Code Editing Context: ${promptAnalysisData.call2context || "Not available"}`;


        if (DEBUG) {
            console.log("====== VALIDATION CORRECTION INPUT ======");
            console.log("System Prompt:", validationSystemPrompt.substring(0,100) + "...");
            console.log("User Input Prompt (truncated):", correctionPrompt.substring(0, 500) + "...");
            console.log("=========================================");
        }

        // Call LLM for correction (processPrompt uses OpenAI key)
        // Pass an empty history, as correction likely doesn't need chat context
        const correctedResponseArray = await processPrompt({
            userInput: correctionPrompt,
            systemPrompt: validationSystemPrompt,
            model: GPT41,
            temperature: 0.7, // Lower temperature for correction
            history: []
        });

        // Save the output using the mock fs (as per original logic)
        const correctionOutputPath = "C:\\Users\\joeor\\Dropbox\\B - Freelance\\C_Projectify\\VanPC\\Training Data\\Main Script Training and Context Data\\validation_correction_output.txt";
        const correctedResponseString = Array.isArray(correctedResponseArray) ? correctedResponseArray.join("\n") : correctedResponseArray;
        fs.writeFileSync(correctionOutputPath, correctedResponseString);

        if (DEBUG) console.log(`Validation correction output saved via mock fs to ${correctionOutputPath}`);
        if (DEBUG) console.log("Corrected Response:", correctedResponseArray);

        return correctedResponseArray; // Return the array format expected by caller

    } catch (error) {
        console.error("Error in validation correction:", error);
        // console.error(error.stack); // Keep stack trace for detailed debugging
        // Return an error message array, consistent with other function returns
        return ["Error during validation correction: " + error.message];
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
    const sendButton = document.getElementById('send');
    const loadingAnimation = document.getElementById('loading-animation');
    
    if (sendButton) {
        sendButton.disabled = isLoading;
    }
    
    if (loadingAnimation) {
        loadingAnimation.style.display = isLoading ? 'flex' : 'none';
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
        let response = await handleConversation(enhancedPrompt, isResponse);
        console.log("Conversation completed");
        console.log("Initial Response:", response);

        if (!response || !Array.isArray(response)) {
            console.error("Invalid response:", response);
            throw new Error("Failed to get valid response from conversation");
        }

        // Run validation and correction if needed
        console.log("Starting validation");
        const validationResults = await validateCodeStrings(response);
        console.log("Validation completed:", validationResults);

        if (validationResults && validationResults.length > 0) {
            console.log("Starting validation correction");
            response = await validationCorrection(userInput, response, validationResults);
            console.log("Validation correction completed");
        }
        
        // Store the response for Excel writing
        lastResponse = response;
        
        // Add assistant message to chat
        appendMessage(response.join('\n'));
        
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

async function findNewCodes(previousText, currentText) {
    const codeRegex = /<[^>]+>/g; // Regex to find <...> codes
    const prevCodes = new Set(previousText.match(codeRegex) || []);
    const currentCodes = currentText.match(codeRegex) || [];
    
    // Filter current codes to find those not present in the previous set
    const newCodes = currentCodes.filter(code => !prevCodes.has(code));
    console.log(`[findNewCodes] Found ${newCodes.length} new codes by comparing content.`);
    if (newCodes.length > 0 && DEBUG) {
        console.log("[findNewCodes] New codes:", newCodes);
    }
    return newCodes;
}

// Helper function to extract the sheet name from the label1 parameter within a TAB tag
function getSheetNameFromTabTag(tabTag) {
    // Matches label1="Sheet Name" inside the tag, capturing "Sheet Name"
    // Allows for optional spaces around =
    const match = tabTag.match(/label1\s*=\s*"([^"]+)"/);
    if (match && match[1]) {
        const sheetName = match[1].trim();
        console.log(`[getSheetNameFromTabTag] Extracted sheet name '${sheetName}' from tag: ${tabTag}`);
        return sheetName;
    }
    // Fallback or error handling if label1 is not found
    console.warn(`[getSheetNameFromTabTag] Could not extract sheet name (label1) from tab tag: ${tabTag}`);
    return null; // Return null if name cannot be extracted
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
                    codesTextarea.setSelectionRange(newCursorPosition, newCursorPosition);

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








