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
export async function processPrompt({ userInput, systemPrompt, model, temperature, history = [] }) {
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
export async function structureDatabasequeries(clientprompt) {
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
          history: [] // Explicitly empty
      });

      if (!queryStrings || !Array.isArray(queryStrings)) {
          console.error("Invalid query strings received:", queryStrings);
          throw new Error("Failed to get valid query strings from structuring prompt");
      }

      if (DEBUG) console.log("Got query strings:", queryStrings);
      const results = [];

      for (const queryString of queryStrings) {
          if (DEBUG) console.log("Processing query:", queryString);
          try {
              // Make sure queryVectorDB uses the internal API keys
              const queryResults = {
                  query: queryString,
                  trainingData: await queryVectorDB({
                      queryPrompt: queryString,
                      similarityThreshold: .2,
                      indexName: 'call2trainingdata',
                      numResults: 3
                  }),
                  call2Context: await queryVectorDB({
                      queryPrompt: queryString,
                      similarityThreshold: .2,
                      indexName: 'call2context',
                      numResults: 5
                  }),
                  call1Context: await queryVectorDB({
                      queryPrompt: queryString,
                      similarityThreshold: .2,
                      indexName: 'call1context',
                      numResults: 5
                  }),
                  codeOptions: await queryVectorDB({
                      queryPrompt: queryString,
                      indexName: 'codes',
                      numResults: 3,
                      similarityThreshold: .1
                  })
              };

              results.push(queryResults);
              if (DEBUG) console.log("Successfully processed query:", queryString);
          } catch (error) {
              console.error(`Error processing query "${queryString}":`, error);
              // Continue with next query instead of failing completely
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

      return results;
  } catch (error) {
      console.error("Error in structureDatabasequeries:", error);
      throw error; // Re-throw
  }
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


// Function: Handle Follow-Up Conversation
export async function handleFollowUpConversation(clientprompt, currentHistory) {
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
    savePromptAnalysis(
        clientprompt,
        systemPrompt,
        mainPromptText,
        null, // No validation prompt info available here
        null, // No validation prompt info available here
        null, // No validation results available here
        safeJsonForPrompt(call2context, false), // Save non-readable for potential re-use
        safeJsonForPrompt(call1context, false),
        safeJsonForPrompt(trainingdataCall2, false),
        safeJsonForPrompt(codeOptions, false),
        responseArray
    );
    saveTrainingData(clientprompt, responseArray);

    if (DEBUG) console.log("Follow-up conversation processed. History length:", updatedHistory.length);

    // Return the response and the updated history
    return { response: responseArray, history: updatedHistory };
}


// Function: Handle Initial Conversation
export async function handleInitialConversation(clientprompt) {
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
    savePromptAnalysis(
        clientprompt,
        systemPrompt,
        mainPromptText,
        null, null, null, // No validation info
        "", "", "", "", // No vector DB context yet
        outputArray
    );
    saveTrainingData(clientprompt, outputArray);

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


// Function: Save prompt analysis data to localStorage
function savePromptAnalysis(clientprompt, systemPrompt, mainPrompt, validationSystemPrompt, validationMainPrompt, validationResults, call2context, call1context, trainingdataCall2, codeOptions, outputArray) {
    try {
        const analysisData = {
            clientRequest: clientprompt || "",
            systemPrompt: systemPrompt || "",
            mainPrompt: mainPrompt || "",
            validationSystemPrompt: validationSystemPrompt || "",
            validationMainPrompt: validationMainPrompt || "",
            validationResults: validationResults || [],
            call2context: call2context || "", // Store the potentially non-readable string used
            call1context: call1context || "",
            trainingdataCall2: trainingdataCall2 || "",
            codeOptions: codeOptions || "",
            outputArray: outputArray || []
        };

        localStorage.setItem('promptAnalysis', JSON.stringify(analysisData));
        if (DEBUG) console.log('Prompt analysis saved to localStorage');
    } catch (error) {
        console.error("Error saving prompt analysis:", error);
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

