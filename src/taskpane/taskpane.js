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
// Add the codeStrings variable with the specified content
// REMOVED hardcoded codeStrings variable

// Mock fs module for browser environment
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
let lastSearchTerm = '';
let lastSearchIndex = -1; // Tracks the starting index of the last found match
let searchResultIndices = []; // Stores indices of all matches for Replace All
let currentHighlightIndex = -1; // Index within searchResultIndices for Find Next

// API keys storage
let API_KEYS = {
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
async function initializeAPIKeys() {
  try {
    console.log("Initializing API keys...");
    
    // Try to load config.js which is .gitignored
    try {
      const configResponse = await fetch('https://localhost:3002/config.js');
      if (configResponse.ok) {
        const configText = await configResponse.text();
        // Extract keys from the config text using regex
        const openaiKeyMatch = configText.match(/OPENAI_API_KEY\s*=\s*["']([^"']+)["']/);
        const pineconeKeyMatch = configText.match(/PINECONE_API_KEY\s*=\s*["']([^"']+)["']/);
        
        if (openaiKeyMatch && openaiKeyMatch[1]) {
          API_KEYS.OPENAI_API_KEY = openaiKeyMatch[1];
          console.log("OpenAI API key loaded from config.js");
        }
        
        if (pineconeKeyMatch && pineconeKeyMatch[1]) {
          API_KEYS.PINECONE_API_KEY = pineconeKeyMatch[1];
          console.log("Pinecone API key loaded from config.js");
        }
      }
    } catch (error) {
      console.warn("Could not load config.js, will use empty API keys:", error);
    }
    
    // Add debug logging with secure masking of keys
    console.log("OPENAI_API_KEY:", API_KEYS.OPENAI_API_KEY ? 
      `${API_KEYS.OPENAI_API_KEY.substring(0, 3)}...${API_KEYS.OPENAI_API_KEY.substring(API_KEYS.OPENAI_API_KEY.length - 3)}` : 
      "Not found");
    console.log("PINECONE_API_KEY:", API_KEYS.PINECONE_API_KEY ? 
      `${API_KEYS.PINECONE_API_KEY.substring(0, 3)}...${API_KEYS.PINECONE_API_KEY.substring(API_KEYS.PINECONE_API_KEY.length - 3)}` : 
      "Not found");
    
    return API_KEYS.OPENAI_API_KEY && API_KEYS.PINECONE_API_KEY;
  } catch (error) {
    console.error("Error initializing API keys:", error);
    return false;
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
const GPT45_TURBO = "gpt-4.5-turbo"
const GPT35_TURBO = "gpt-3.5-turbo"
const GPT4_TURBO = "gpt-4-turbo"
const GPTFT1 =  "ft:gpt-3.5-turbo-1106:orsi-advisors:cohcolumnsgpt35:B6Wlrql1"

// Conversation history storage
let conversationHistory = [];

// Functions to save and load conversation history
function saveConversationHistory(history) {
    try {
        localStorage.setItem('conversationHistory', JSON.stringify(history));
        if (DEBUG) console.log('Conversation history saved to localStorage');
    } catch (error) {
        console.error("Error saving conversation history:", error);
    }
}

function loadConversationHistory() {
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

// Direct OpenAI API call function (replaces LangChain)
async function callOpenAI(messages, model = GPT4O, temperature = 0.7) {
  try {
    console.log(`Calling OpenAI API with model: ${model}`);
    
    // Check for API key
    if (!API_KEYS.OPENAI_API_KEY) {
      throw new Error("OpenAI API key not found. Please check your API keys.");
    }
    
    const response = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${API_KEYS.OPENAI_API_KEY}`
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

// OpenAI embeddings function (replaces LangChain)
async function createEmbedding(text) {
  try {
    console.log("Creating embedding for text");
    
    // Check for API key
    if (!API_KEYS.OPENAI_API_KEY) {
      throw new Error("OpenAI API key not found. Please check your API keys.");
    }
    
    const response = await fetch('https://api.openai.com/v1/embeddings', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${API_KEYS.OPENAI_API_KEY}`
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

// Remove the PROMPTS object and add a function to load prompts
async function loadPromptFromFile(promptKey) {
  try {
    // Use a simplified path approach that works with dev server with correct port
    const paths = [
      `https://localhost:3002/prompts/${promptKey}.txt`,
    ];
    
    // Combine all paths to try
    paths.push(...srcPaths);
 
    // Try each path until one works
    let response = null;
    for (const path of paths) {
      console.log(`Attempting to load prompt from: ${path}`);
      try {
        response = await fetch(path);
        if (response.ok) {
          console.log(`Successfully loaded prompt from: ${path}`);
          break;
        }
      } catch (err) {
        console.log(`Path ${path} failed: ${err.message}`);
      }
    }
    
    if (!response || !response.ok) {
      throw new Error(`Failed to load prompt: ${promptKey} (Could not find file in any location)`);
    }
    
    return await response.text();
  } catch (error) {
    console.error(`Error loading prompt ${promptKey}:`, error);
    throw error; // Re-throw the error to be handled by the caller
  }
}

// Update the getSystemPromptFromFile function
const getSystemPromptFromFile = async (promptKey) => {
  try {
    const prompt = await loadPromptFromFile(promptKey);
    if (!prompt) {
      throw new Error(`Prompt key "${promptKey}" not found`);
    }
    return prompt;
  } catch (error) {
    console.error(`Error getting prompt for key ${promptKey}:`, error);
    return null;
  }
};

//************Functions************
// Function 1: OpenAI Call with conversation history support
async function processPrompt({ userInput, systemPrompt, model, temperature, history = [] }) {
    console.log("API Key being used:", API_KEYS.OPENAI_API_KEY ? `${API_KEYS.OPENAI_API_KEY.substring(0, 3)}...` : "None");
    
    // Format messages in the way OpenAI expects
    const messages = [
        { role: "system", content: systemPrompt }
    ];
    
    // Add conversation history
    if (history.length > 0) {
        history.forEach(message => {
            messages.push({ 
                role: message[0] === "human" ? "user" : "assistant", 
                content: message[1] 
            });
        });
    }
    
    // Add current user input
    messages.push({ role: "user", content: userInput });
    
    try {
        // Call OpenAI API directly
        const responseContent = await callOpenAI(messages, model, temperature);
        
        // Try to parse JSON response if applicable
        try {
            const parsed = JSON.parse(responseContent);
            if (Array.isArray(parsed)) {
                return parsed;
            }
            return responseContent.split('\n').filter(line => line.trim());
        } catch (e) {
            // If not JSON, treat as text and split by lines
            return responseContent.split('\n').filter(line => line.trim());
        }
    } catch (error) {
        console.error("Error in processPrompt:", error);
        throw error;
    }
}

async function structureDatabasequeries(clientprompt) {
  if (DEBUG) console.log("Processing structured database queries:", clientprompt);

  try {
      console.log("Getting structure system prompt");
      const systemStructurePrompt = await getSystemPromptFromFile('Structure_System');
      
      if (!systemStructurePrompt) {
          throw new Error("Failed to load structure system prompt");
      }

      console.log("Got system prompt, processing query strings");
      const queryStrings = await processPrompt({
          userInput: clientprompt,
          systemPrompt: systemStructurePrompt,
          model: GPT4O,
          temperature: 1
      });

      if (!queryStrings || !Array.isArray(queryStrings)) {
          throw new Error("Failed to get valid query strings");
      }

      console.log("Got query strings:", queryStrings);
      const results = [];

      for (const queryString of queryStrings) {
          console.log("Processing query:", queryString);
          try {
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
              console.log("Successfully processed query:", queryString);
          } catch (error) {
              console.error(`Error processing query "${queryString}":`, error);
              // Continue with next query instead of failing completely
              continue;
          }
      }

      if (results.length === 0) {
          throw new Error("No valid results were obtained from any queries");
      }

      return results;
  } catch (error) {
      console.error("Error in structureDatabasequeries:", error);
      throw error;
  }
}

// Function 3: Query Vector Database using Pinecone REST API
async function queryVectorDB({ queryPrompt, indexName = 'codes', numResults = 10, similarityThreshold = null }) {
    try {
        console.log("Generating embeddings for query:", queryPrompt);
        
        // Generate embeddings using our direct API call
        const embedding = await createEmbedding(queryPrompt);
        console.log("Embeddings generated successfully");
        
        // Get the correct endpoint for the specified index
        const indexConfig = PINECONE_INDEXES[indexName];
        if (!indexConfig) {
            throw new Error(`Invalid index name: ${indexName}`);
        }
        
        const url = `${indexConfig.apiEndpoint}/query`;
        console.log("Making Pinecone API request to:", url);
        
        const response = await fetch(url, {
            method: 'POST',
            headers: {
                'api-key': API_KEYS.PINECONE_API_KEY,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                vector: embedding,
                topK: numResults,
                includeMetadata: true,
                namespace: "ns1"
            })
        });

        if (!response.ok) {
            const errorText = await response.text();
            console.error("Pinecone API error details:", {
                status: response.status,
                statusText: response.statusText,
                error: errorText
            });
            throw new Error(`Pinecone API error: ${response.status} ${response.statusText} - ${errorText}`);
        }

        const data = await response.json();
        console.log("Pinecone API response received");
        
        let matches = data.matches || [];

        if (similarityThreshold !== null) {
            matches = matches.filter(match => match.score >= similarityThreshold);
        }

        matches = matches.slice(0, numResults);

        matches = matches.map(match => {
            try {
                if (match.metadata && match.metadata.text) {
                    return {
                        ...match,
                        text: match.metadata.text
                    };
                }
                return match;
            } catch (error) {
                console.error("Error processing match:", error);
                return match;
            }
        });

        if (DEBUG) {
            const matchesDescription = matches
                .map((match, i) => `Match ${i + 1} (score: ${match.score.toFixed(4)}): ${match.text || JSON.stringify(match.metadata)}`)
                .join('\n');
            console.log(matchesDescription);
        }

        const cleanMatches = matches.map(match => extractTextFromJson(match));
        return cleanMatches.filter(text => text !== "");

    } catch (error) {
        console.error("Error during vector database query:", error);
        throw error;
    }
}

function extractTextFromJson(jsonInput) {
   try {
       const jsonData = typeof jsonInput === 'string' ? JSON.parse(jsonInput) : jsonInput;
       
       if (Array.isArray(jsonData)) {
           for (const item of jsonData) {
               if (item.metadata && item.metadata.text) {
                   return item.metadata.text;
               }
           }
           throw new Error("No text field found in the JSON array");
       } 
       else if (jsonData.metadata && jsonData.metadata.text) {
           return jsonData.metadata.text;
       } 
       else {
           throw new Error("Invalid JSON structure: missing metadata.text field");
       }
   } catch (error) {
       console.error(`Error processing JSON: ${error.message}`);
       return "";
   }
}

function safeJsonForPrompt(obj, readable = true) {
    if (!readable) {
        let jsonString = JSON.stringify(obj);
        jsonString = jsonString.replace(/"values":\s*\[\],\s*"metadata":/g, '');
        return jsonString
            .replace(/{/g, '\\u007B')
            .replace(/}/g, '\\u007D');
    }
    
    if (Array.isArray(obj)) {
        return obj.map(item => {
            if (item.metadata && item.metadata.text) {
                const text = item.metadata.text.replace(/~/g, ',');
                const parts = text.split(';');
                
                let result = '';
                if (parts.length >= 1) result += parts[0].trim();
                if (parts.length >= 2) result += '\n' + parts[1].trim();
                if (parts.length >= 3) result += '\n' + parts[2].trim();
                
                if (item.score) {
                    result += `\nSimilarity Score: ${item.score.toFixed(4)}`;
                }
                
                return result;
            }
            return JSON.stringify(item).replace(/~/g, ',');
        }).join('\n\n');
    }
    
    const jsonString = JSON.stringify(obj, null, 2).replace(/~/g, ',');
    return jsonString
        .replace(/{/g, '\\u007B')
        .replace(/}/g, '\\u007D');
}

async function handleFollowUpConversation(clientprompt) {
    if (DEBUG) console.log("Processing follow-up question. Loading conversation history...");
    conversationHistory = loadConversationHistory();
    
    if (conversationHistory.length > 0) {
        if (DEBUG) console.log("Processing follow-up question:", clientprompt);
        if (DEBUG) console.log("Loaded conversation history:", JSON.stringify(conversationHistory, null, 2));
        
        const systemPrompt = await getSystemPromptFromFile('Followup_System');
        // const MainPrompt = await getSystemPromptFromFile('main');
        
        const trainingdataCall2 = await queryVectorDB({
            queryPrompt: clientprompt,
            similarityThreshold: .4,
            indexName: 'call2trainingdata',
            numResults: 3
        });

        const call2context = await queryVectorDB({
            queryPrompt: clientprompt + trainingdataCall2,
            similarityThreshold: .3,
            indexName: 'call2context',
            numResults: 5
        });

        const call1context = await queryVectorDB({
            queryPrompt: clientprompt + trainingdataCall2,
            similarityThreshold: .3,
            indexName: 'call1context',
            numResults: 5
        });

        const codeOptions = await queryVectorDB({
            queryPrompt: clientprompt + trainingdataCall2 + call1context,
            indexName: 'codes',
            numResults: 10,
            similarityThreshold: .1
        });
        
        const followUpPrompt = "Client request: " + clientprompt + "\n" +
                       "Main Prompt: " + MainPrompt + "\n" +
                       "Training Data: " + safeJsonForPrompt(trainingdataCall2).replace(/~/g, ',') + "\n" +
                       "Code choosing context: " + safeJsonForPrompt(call1context) + "\n" +
                       "Code editing Context: " + safeJsonForPrompt(call2context) + "\n" +
                       "Code descriptions: " + safeJsonForPrompt(codeOptions);
        
        const response = await processPrompt({
            userInput: followUpPrompt,
            systemPrompt: systemPrompt,
            model: GPT4O,
            temperature: 1,
            history: conversationHistory
        });
        
        conversationHistory.push(["human", clientprompt]);
        conversationHistory.push(["assistant", response.join("\n")]);
        
        saveConversationHistory(conversationHistory);
        
        if (DEBUG) console.log("Updated conversation history:", JSON.stringify(conversationHistory, null, 2));
        
        savePromptAnalysis(clientprompt, systemPrompt, MainPrompt, call2context, call1context, trainingdataCall2, codeOptions, response);
        saveTrainingData(clientprompt, response);
        
        return response;
    } else {
        if (DEBUG) console.log("No conversation history found. Treating as initial question.");
        return handleInitialConversation(clientprompt);
    }
}

async function handleConversation(clientprompt, isFollowUp = false) {
    try {
        if (isFollowUp) {
            return await handleFollowUpConversation(clientprompt);
        } else {
            return await handleInitialConversation(clientprompt);
        }
    } catch (error) {
        console.error("Error in conversation handling:", error);
        return ["Error processing your request: " + error.message];
    }
}

async function handleInitialConversation(clientprompt) {
    if (DEBUG) console.log("Processing initial question:", clientprompt);
    
    const systemPrompt = await getSystemPromptFromFile('Encoder_System');
    console.log("SYSTEM PROMPT: ", systemPrompt);
    const MainPrompt = await getSystemPromptFromFile('Encoder_Main');
    console.log("MAIN PROMPT: ", MainPrompt);


    const Call2prompt = "Client request: " + clientprompt + "\n" +
                       "Main Prompt: " + MainPrompt;
    
    const outputArray2 = await processPrompt({
        userInput: Call2prompt,
        systemPrompt: systemPrompt,
        model: GPT4O,
        temperature: 1 
    });
    
    conversationHistory = [
        ["human", clientprompt],
        ["assistant", outputArray2.join("\n")]
    ];
    
    saveConversationHistory(conversationHistory);
    
    savePromptAnalysis(clientprompt, systemPrompt, MainPrompt, [], [], [], [], outputArray2);
    saveTrainingData(clientprompt, outputArray2);
    
    console.log("Initial Response - in the function:", outputArray2);
    return outputArray2;

}

function savePromptAnalysis(clientprompt, systemPrompt, MainPrompt, validationSystemPrompt, validationMainPrompt, validationResults, call2context, call1context, trainingdataCall2, codeOptions, outputArray2) {
    try {
        const analysisData = {
            clientRequest: clientprompt,
            systemPrompt,
            mainPrompt: MainPrompt,
            validationSystemPrompt,
            validationMainPrompt,
            validationResults,
            call2context,
            call1context,
            trainingdataCall2,
            codeOptions,
            outputArray2
        };
        
        localStorage.setItem('promptAnalysis', JSON.stringify(analysisData));
        if (DEBUG) console.log('Prompt analysis saved to localStorage');
    } catch (error) {
        console.error("Error saving prompt analysis:", error);
    }
}

function saveTrainingData(clientprompt, outputArray2) {
    try {
        function cleanText(text) {
            if (!text) return '';
            return text.toString()
                .replace(/\r?\n|\r/g, ' ')
                .trim();
        }
        
        const trainingData = {
            prompt: cleanText(clientprompt),
            response: cleanText(JSON.stringify(outputArray2))
        };
        
        localStorage.setItem('trainingData', JSON.stringify(trainingData));
        if (DEBUG) console.log('Training data saved to localStorage');
    } catch (error) {
        console.error("Error saving training data:", error);
    }
}

async function validationCorrection(clientprompt, initialResponse, validationResults) {
    try {
        const conversationHistory = loadConversationHistory();
        
        const trainingData = localStorage.getItem('trainingData') || "";
        const codeDescriptions = localStorage.getItem('codeDescriptions') || "";
        const lastCallContext = localStorage.getItem('lastCallContext') || "";
        
        const validationSystemPrompt = await getSystemPromptFromFile('Validation_System');
        const validationMainPrompt = await getSystemPromptFromFile('Validation_Main');
        
        if (!validationSystemPrompt) {
            throw new Error("Failed to load validation system prompt");
        }
        
        const responseString = Array.isArray(initialResponse) ? initialResponse.join("\n") : String(initialResponse);
        
        const correctionPrompt = 
            "Main Prompt: " + validationMainPrompt + "\n\n" +
            "Original User Input: " + clientprompt + "\n\n" +
            "Initial Response: " + responseString + "\n\n" +
            "Validation Results: " + validationResults + "\n\n" +
            "Training Data: " + trainingData + "\n\n" +
            "Code Descriptions: " + codeDescriptions + "\n\n" +
            "Context from Last Call: " + lastCallContext;
        
        if (DEBUG) {
            console.log("====== VALIDATION CORRECTION INPUT ======");
            console.log(correctionPrompt.substring(0, 500) + "...(truncated)");
            console.log("=========================================");
        }
        
        const correctedResponse = await processPrompt({
            userInput: correctionPrompt,
            systemPrompt: validationSystemPrompt,
            model: GPT4O,
            temperature: 0.7
        });
        
        const correctionOutputPath = "C:\\Users\\joeor\\Dropbox\\B - Freelance\\C_Projectify\\VanPC\\Training Data\\Main Script Training and Context Data\\validation_correction_output.txt";
        fs.writeFileSync(correctionOutputPath, Array.isArray(correctedResponse) ? correctedResponse.join("\n") : correctedResponse);
        
        if (DEBUG) console.log(`Validation correction saved to ${correctionOutputPath}`);
        
        return correctedResponse;
    } catch (error) {
        console.error("Error in validation correction:", error);
        console.error(error.stack);
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

/**
 * Inserts worksheets from a base64-encoded Excel file
 */
async function insertSheetsFromBase64() {
    try {
        // Fetch the Excel file
        const response = await fetch('https://localhost:3002/assets/Worksheets_4.3.25 v1.xlsx');
        if (!response.ok) {
            throw new Error('Failed to load Excel file');
        }
        
        // Convert the response to an ArrayBuffer
        const arrayBuffer = await response.arrayBuffer();
        
        // Convert ArrayBuffer to base64 string in chunks
        const uint8Array = new Uint8Array(arrayBuffer);
        let binaryString = '';
        const chunkSize = 8192; // Process in 8KB chunks
        
        for (let i = 0; i < uint8Array.length; i += chunkSize) {
            const chunk = uint8Array.slice(i, Math.min(i + chunkSize, uint8Array.length));
            binaryString += String.fromCharCode.apply(null, chunk);
        }
        
        const base64String = btoa(binaryString);
        
        // Call the function to insert worksheets
        await handleInsertWorksheetsFromBase64(base64String);
        console.log("Worksheets inserted successfully");
    } catch (error) {
        console.error("Error inserting worksheets:", error);
        showError(error.message);
    }
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

            for (const currentTab of currentTabs) {
                const currentTag = currentTab.tag;
                const currentText = currentTab.text;
                const previousText = previousTabMap.get(currentTag);

                if (previousText === undefined) {
                    // *** New Tab ***
                    console.log(`[Run Codes] Identified NEW tab: ${currentTag}`);
                    allCodeContentToProcess += currentText + "\n\n"; // Add full content
                    hasAnyChanges = true;
                } else if (previousText !== currentText) {
                    // *** Modified Existing Tab ***
                    console.log(`[Run Codes] Identified MODIFIED tab: ${currentTag}.`);
                    allCodeContentToProcess += currentText + "\n\n"; // Add full content
                    hasAnyChanges = true; // Mark change because the text differs
                }
                // else: Tab exists and text is identical - do nothing.
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
                console.log("[Run Codes] Changes detected and validated. Inserting base sheets from codes.xlsx...");
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

    // >>> ADDED: Helper function findNewCodes (defined above, ensure it's accessible)
    // No need to add it here again if defined globally or within taskpane.js scope before Office.onReady


    // ... (rest of your Office.onReady remains the same) ...

    // Keep the setup for your other buttons (send-button, reset-button, etc.)
    const sendButton = document.getElementById('send');
    if (sendButton) sendButton.onclick = handleSend;

    const writeButton = document.getElementById('write-to-excel');
    if (writeButton) writeButton.onclick = writeToExcel;

    const resetButton = document.getElementById('reset-chat');
    if (resetButton) resetButton.onclick = resetChat;

    const codesTextarea = document.getElementById('codes-textarea');
    const editParamsButton = document.getElementById('edit-code-params-button');
    const paramsModal = document.getElementById('code-params-modal');
    const paramsModalForm = document.getElementById('code-params-modal-form');
    const closeModalButton = paramsModal.querySelector('.close-button');
    const applyParamsButton = document.getElementById('apply-code-params-button');
    const cancelParamsButton = document.getElementById('cancel-code-params-button');

    let currentCodeStringRange = null; // To store {start, end} of the code string being edited
    let currentCodeStringType = ''; // To store the type like 'VOL-EV'

    // Function to show the modal
    const showParamsModal = () => {
        if (paramsModal) paramsModal.style.display = 'block';
    };

    // Function to hide the modal
    const hideParamsModal = () => {
        if (paramsModal) paramsModal.style.display = 'none';
        paramsModalForm.innerHTML = ''; // Clear the form
        currentCodeStringRange = null; // Reset state
        currentCodeStringType = '';
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

    // Function to populate the modal form
    const populateParamsModal = (type, params) => {
        paramsModalForm.innerHTML = ''; // Clear previous form items
        currentCodeStringType = type; // Store the type

        Object.entries(params).forEach(([key, value]) => {
            const paramEntryDiv = document.createElement('div');
            paramEntryDiv.className = 'param-entry';

            const label = document.createElement('label');
            label.htmlFor = `param-${key}`;
            label.textContent = key;

            // Use textarea for potentially long values like 'row1', input otherwise
            let inputElement;
            if (key.toLowerCase().includes('row') || value.length > 60) { // Heuristic for textarea
                inputElement = document.createElement('textarea');
                inputElement.rows = 3; // Adjust as needed
            } else {
                inputElement = document.createElement('input');
                inputElement.type = 'text';
            }

            inputElement.id = `param-${key}`;
            inputElement.value = value;
            inputElement.dataset.paramKey = key; // Store the key

            paramEntryDiv.appendChild(label);
            paramEntryDiv.appendChild(inputElement);
            paramsModalForm.appendChild(paramEntryDiv);
        });
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

    if (applyParamsButton && codesTextarea) {
        applyParamsButton.onclick = () => {
            if (!currentCodeStringRange || !currentCodeStringType) return; // Safety check

            const formElements = paramsModalForm.querySelectorAll('input[data-param-key], textarea[data-param-key]');
            const updatedParams = [];

            formElements.forEach(input => {
                const key = input.dataset.paramKey;
                const value = input.value;
                // Re-add quotes around the value
                updatedParams.push(`${key}="${value}"`);
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

            // Optionally, update cursor position
            const newCursorPos = currentCodeStringRange.start + newCodeString.length;
            codesTextarea.focus();
            codesTextarea.setSelectionRange(newCursorPos, newCursorPos);

            hideParamsModal(); // Close modal after applying
        };
    }

    // Hide modal if clicked outside the content
    window.onclick = (event) => {
        if (event.target == paramsModal) {
            hideParamsModal();
        }
    };


    // ... (rest of your Office.onReady, including suggestion logic, initializations)

    // Make sure initialization runs after setting up modal logic
    Promise.all([
        initializeAPIKeys(),
        loadCodeDatabase()
    ]).then(([keysLoaded]) => {
      if (!keysLoaded) {
        showError("Failed to load API keys. Please check configuration.");
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

// >>> ADDED: Search and Replace Functions <<<

function clearSearchHighlight() {
    const textarea = document.getElementById('codes-textarea');
    if (textarea && lastSearchIndex !== -1) {
        // Simple way: just reset selection to the start
        textarea.setSelectionRange(0, 0);
        textarea.blur(); // Remove focus to clear visible selection highlight
        textarea.focus();
        console.log("Cleared search highlight.");
    }
    lastSearchIndex = -1;
    searchResultIndices = [];
    currentHighlightIndex = -1;
    updateSearchStatus('');
}

function updateSearchStatus(message) {
    const statusElement = document.getElementById('search-status');
    if (statusElement) {
        statusElement.textContent = message;
    }
}

function findNext() {
    const textarea = document.getElementById('codes-textarea');
    const searchTerm = document.getElementById('search-input').value;
    const statusElement = document.getElementById('search-status');
    const selectionOnlyCheckbox = document.getElementById('search-selection-only');

    if (!textarea || !searchTerm) {
        updateSearchStatus("Enter a search term.");
        return;
    }

    const isSelectionOnly = selectionOnlyCheckbox?.checked;
    let currentText = textarea.value;
    let scopeStartIndex = 0;
    let selectionEndIndex = currentText.length; // Use current length

    if (isSelectionOnly) {
        scopeStartIndex = textarea.selectionStart;
        selectionEndIndex = textarea.selectionEnd;
        if (scopeStartIndex === selectionEndIndex) {
            updateSearchStatus("Select text first for 'Search Selection Only'.");
            return;
        }
        let searchScopeText = currentText.substring(scopeStartIndex, selectionEndIndex);
        console.log(`Searching within selection (${scopeStartIndex}-${selectionEndIndex}): "${searchScopeText}"`);

        // Determine if a re-scan is needed
        const storedSelStart = textarea.dataset.lastSelectionScanStart;
        const storedSelEnd = textarea.dataset.lastSelectionScanEnd;
        const needsRescan = lastSearchTerm !== searchTerm || 
                            !storedSelStart || 
                            storedSelStart != scopeStartIndex || 
                            storedSelEnd != selectionEndIndex;

        if (needsRescan) {
            console.log("Re-scanning selection.");
            lastSearchTerm = searchTerm;
            textarea.dataset.lastSelectionScanStart = scopeStartIndex; // Store bounds used for THIS scan
            textarea.dataset.lastSelectionScanEnd = selectionEndIndex;
            lastSearchIndex = -1; // Absolute index reset
            currentHighlightIndex = -1;
            searchResultIndices = []; // Stores relative indices

            let relativeIndex = searchScopeText.indexOf(searchTerm);
            while (relativeIndex !== -1) {
                searchResultIndices.push(relativeIndex);
                relativeIndex = searchScopeText.indexOf(searchTerm, relativeIndex + 1);
            }
            console.log(`Found ${searchResultIndices.length} occurrences within selection. Relative Indices:`, searchResultIndices);
        }
        // If no re-scan needed, we continue with existing searchResultIndices and currentHighlightIndex

        if (searchResultIndices.length === 0) {
            updateSearchStatus(`"${searchTerm}" not found in selection.`);
            return;
        }
    }
    else { // Full text search
        const storedSelStart = textarea.dataset.lastSelectionScanStart;
        // Reset if term changed or switching FROM selection mode
        if (searchTerm !== lastSearchTerm || storedSelStart) {
            console.log("Scanning full text.");
            lastSearchTerm = searchTerm;
            lastSearchIndex = -1; // Absolute index
            currentHighlightIndex = -1;
            searchResultIndices = []; // Stores absolute indices
            textarea.dataset.lastSelectionScanStart = ''; // Clear selection memory
            textarea.dataset.lastSelectionScanEnd = '';

            let index = currentText.indexOf(searchTerm);
            while (index !== -1) {
                searchResultIndices.push(index);
                index = currentText.indexOf(searchTerm, index + 1);
            }
            console.log(`Found ${searchResultIndices.length} occurrences of "${searchTerm}". Absolute Indices:`, searchResultIndices);
        }

        if (searchResultIndices.length === 0) {
            updateSearchStatus(`"${searchTerm}" not found.`);
            return;
        }
    }

    // Cycle through the found indices
    currentHighlightIndex = (currentHighlightIndex + 1) % searchResultIndices.length;
    const foundIndex = searchResultIndices[currentHighlightIndex]; // Could be relative or absolute

    // Highlight the found text (calculate absolute index)
    const highlightStartIndex = isSelectionOnly ? scopeStartIndex + foundIndex : foundIndex;
    const highlightEndIndex = highlightStartIndex + searchTerm.length;

    // Store the absolute index of the current highlight for replace validation
    lastSearchIndex = highlightStartIndex;

    textarea.focus();
    textarea.setSelectionRange(highlightStartIndex, highlightEndIndex);
    textarea.scrollTop = textarea.scrollHeight * (highlightStartIndex / currentText.length) - 50; // Estimate scroll position

    updateSearchStatus(`Found at index ${highlightStartIndex} (${currentHighlightIndex + 1}/${searchResultIndices.length})${isSelectionOnly ? ' (in selection)' : ''}`);
    console.log(`Highlighting ${isSelectionOnly ? 'relative index ' + foundIndex : 'absolute index'} (Absolute: ${highlightStartIndex})`);
}

function replace() {
    const textarea = document.getElementById('codes-textarea');
    const searchTerm = document.getElementById('search-input').value;
    const replaceTerm = document.getElementById('replace-input').value;
    const selectionOnlyCheckbox = document.getElementById('search-selection-only');

    if (!textarea || !searchTerm) {
        updateSearchStatus("Enter a search term.");
        return;
    }

    // Must have a valid highlighted match from findNext
    if (lastSearchIndex === -1 || textarea.selectionStart !== lastSearchIndex || textarea.selectionEnd !== lastSearchIndex + searchTerm.length) {
         updateSearchStatus("Find match first.");
         // Attempt to find the first match if none is selected
         findNext();
         return;
    }

    const isSelectionOnly = selectionOnlyCheckbox?.checked;

    // If selection only, double-check the match is within the bounds used for the last scan
    if (isSelectionOnly) {
        const storedSelStart = parseInt(textarea.dataset.lastSelectionScanStart || '-1', 10);
        const storedSelEnd = parseInt(textarea.dataset.lastSelectionScanEnd || '-1', 10);
        if (storedSelStart === -1 || lastSearchIndex < storedSelStart || (lastSearchIndex + searchTerm.length) > storedSelEnd) {
            updateSearchStatus("Match is outside selection bounds. Find Next?");
            console.log("Replace cancelled: Highlight outside selection scan bounds.");
            // Reset highlight and let user find again
            clearSearchHighlight();
            lastSearchTerm = searchTerm; // Keep term
            return;
        }
    }

    // Perform the replacement
    const currentText = textarea.value;
    const before = currentText.substring(0, lastSearchIndex);
    const after = currentText.substring(lastSearchIndex + searchTerm.length);
    const lengthDifference = replaceTerm.length - searchTerm.length;

    textarea.value = before + replaceTerm + after;
    console.log(`Replaced "${searchTerm}" with "${replaceTerm}" at absolute index ${lastSearchIndex}`);

    // --- State Update ---
    if (isSelectionOnly) {
        // Update the END boundary used for the scan to reflect the change
        const storedSelEnd = parseInt(textarea.dataset.lastSelectionScanEnd || '-1', 10);
        if (storedSelEnd !== -1) {
            textarea.dataset.lastSelectionScanEnd = storedSelEnd + lengthDifference;
            console.log(`Updated selection scan end boundary to: ${textarea.dataset.lastSelectionScanEnd}`);
        }
        // Clear the specific match state, forcing findNext to re-scan the selection
        lastSearchIndex = -1;
        searchResultIndices = [];
        currentHighlightIndex = -1;
    } else {
        // Full text replace: Just clear everything, simplest approach
        clearSearchHighlight();
        lastSearchTerm = searchTerm; // Keep term
    }

    // Set cursor position after the replaced text
    textarea.focus();
    const newCursorPos = lastSearchIndex + replaceTerm.length; // lastSearchIndex is the START of the replaced section
    textarea.setSelectionRange(newCursorPos, newCursorPos);

    updateSearchStatus(`Replaced at ${lastSearchIndex}. Find Next?`);

    // DO NOT automatically call findNext(). Let the user do it.
}

function replaceAll() {
    const textarea = document.getElementById('codes-textarea');
    const searchTerm = document.getElementById('search-input').value;
    const replaceTerm = document.getElementById('replace-input').value;
    const selectionOnlyCheckbox = document.getElementById('search-selection-only'); // <<< ADDED

    if (!textarea || !searchTerm) {
        updateSearchStatus("Enter search term.");
        return;
    }

    const isSelectionOnly = selectionOnlyCheckbox?.checked; // <<< ADDED
    let replacementsMade = 0;
    const originalText = textarea.value;
    let newText = originalText;

    // Escaping regex special characters in search term for safety
    const escapedSearchTerm = searchTerm.replace(/[.*+?^${}()|[\\]\\\\]/g, '\\\\$&');
    const regex = new RegExp(escapedSearchTerm, 'g'); // Global flag

    // <<< ADDED: Logic for Selection Only >>>
    if (isSelectionOnly) {
        const startIndex = textarea.selectionStart;
        const endIndex = textarea.selectionEnd;

        if (startIndex === endIndex) {
            updateSearchStatus("Select text first for 'Replace All in Selection'.");
            return;
        }

        const selectedText = originalText.substring(startIndex, endIndex);
        let replacedSelectedText = selectedText.replace(regex, () => {
            replacementsMade++;
            return replaceTerm;
        });

        if (replacementsMade > 0) {
            newText = originalText.substring(0, startIndex) + replacedSelectedText + originalText.substring(endIndex);
             textarea.value = newText;
             // Optionally re-select the modified text
             textarea.focus();
             textarea.setSelectionRange(startIndex, startIndex + replacedSelectedText.length);

             console.log(`Replaced ${replacementsMade} occurrences within selection.`);
             updateSearchStatus(`Replaced ${replacementsMade} in selection.`);
             // Reset search state
             lastSearchTerm = '';
             clearSearchHighlight();
        } else {
             updateSearchStatus(`"${searchTerm}" not found in selection.`);
             console.log(`"${searchTerm}" not found for Replace All within selection.`);
        }

    }
    // <<< END ADDED >>>
    else {
        // Full text replace (existing logic)
        newText = originalText.replace(regex, () => {
            replacementsMade++;
            return replaceTerm;
        });

        if (replacementsMade > 0) {
            textarea.value = newText;
            console.log(`Replaced ${replacementsMade} occurrences of "${searchTerm}" with "${replaceTerm}".`);
            updateSearchStatus(`Replaced ${replacementsMade} occurrences.`);
            // Reset search state as content has significantly changed
            lastSearchTerm = '';
            clearSearchHighlight();
        } else {
            updateSearchStatus(`"${searchTerm}" not found.`);
            console.log(`"${searchTerm}" not found for Replace All.`);
        }
    }
}

