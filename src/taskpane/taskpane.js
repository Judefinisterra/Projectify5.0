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

// Function to insert sheets and then run code collection
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
        // Optionally show a brief success message, though might be too noisy
        // showMessage("Codes saved automatically."); 
    } catch (error) {
        console.error("[Run Codes] Error auto-saving codes to localStorage:", error);
        showError(`Error automatically saving codes: ${error.message}. Run may not reflect latest changes.`);
        // Decide if we should stop here or continue with potentially old codes
        // For now, we'll continue but the user has been warned.
    }
    // <<< END ADDED CODE

    let codesToRun = loadedCodeStrings; // Use the updated global variable
    let previousCodes = null;
    let codeStringToProcess = ""; // Reset this here

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
            console.log("Inserting base sheets from Worksheets_4.3.25 v1.xlsx...");
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
            
            // Process ALL codes on first pass
            codeStringToProcess = codesToRun;

        } else {
            // *** SECOND PASS (or later) ***
            console.log("[Run Codes] SUBSEQUENT PASS: Financials sheet found.");

            // Load previous codes for comparison
            try {
                previousCodes = localStorage.getItem('previousRunCodeStrings');
            } catch (error) {
                 console.error("[Run Codes] Error loading previous codes for comparison:", error);
                 console.warn("[Run Codes] Could not load previous codes. Processing ALL current codes as fallback.");
                 previousCodes = null;
            }

            // Check for NO changes first
            if (previousCodes !== null && previousCodes === codesToRun) {
                 console.log("[Run Codes] No change in code strings since last run. Nothing to process.");
                 // Update previous codes timestamp anyway, then exit
                 try { localStorage.setItem('previousRunCodeStrings', codesToRun); } catch(e) { console.error("Err updating prev codes:", e); }
                 showMessage("No code changes to run."); // <<< Show message to user
                 setButtonLoading(false); // Ensure button is re-enabled
                 return; // <<< EXIT EARLY
            }

            // If changed or first subsequent run, insert codes.xlsx
            console.log("Inserting base sheets from codes.xlsx...");
            try {
                const codesResponse = await fetch('https://localhost:3002/assets/codes.xlsx');
                if (!codesResponse.ok) throw new Error(`codes.xlsx load failed: ${codesResponse.statusText}`);
                const codesArrayBuffer = await codesResponse.arrayBuffer();
                console.log("Converting codes.xlsx file to base64 (using chunks)...");
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

            // Now determine which codes to process based on comparison
            if (previousCodes === null) {
                 console.log("[Run Codes] First subsequent run or previous codes missing. Processing all current codes.");
                 codeStringToProcess = codesToRun;
            } else {
                // Find new tabs (since previousCodes !== codesToRun here)
                const currentTabs = getTabBlocks(codesToRun);
                const previousTabs = getTabBlocks(previousCodes);
                const previousTabTags = new Set(previousTabs.map(block => block.tag));
                const newTabs = currentTabs.filter(block => !previousTabTags.has(block.tag));
                
                if (newTabs.length > 0) {
                    console.log(`[Run Codes] Found ${newTabs.length} new TAB block(s) to process:`, newTabs.map(t => t.tag));
                    codeStringToProcess = newTabs.map(block => block.text).join('\n\n'); // Join new blocks
                    console.log("[Run Codes] Filtered codeStringToProcess for new tabs (truncated):", codeStringToProcess.substring(0, 200) + (codeStringToProcess.length > 200 ? '...' : ''));
                } else {
                    console.log("[Run Codes] Codes changed, but no new TAB blocks detected. Nothing specific to process incrementally.");
                    codeStringToProcess = ""; // Ensure nothing is processed if only old tabs changed
                }
            }
        } // End of pass logic (if/else financialsSheetExists)
        
        // --- Execute Processing --- 
        if (codeStringToProcess.trim().length === 0) {
             console.log("[Run Codes] After filtering/pass logic, there are no codes to process.");
             // Only show message if it wasn't already shown by the 'no change' check
             if (!(previousCodes !== null && previousCodes === codesToRun)) {
                showMessage("No new codes identified to run.");
             }
             // Still update previous state even if nothing ran this time
             try { localStorage.setItem('previousRunCodeStrings', codesToRun); } catch(e) { console.error("Err updating prev codes:", e); }
             setButtonLoading(false);
             return; // Exit if nothing to process
        }
         
        // If we have codes to process, continue...
        console.log("Populating collection for codes to process...");
        
        // >>> ADDED: Validate the codes before running
        console.log("Validating codes before execution...");
        // Use the new validation function
        const validationErrors = await validateCodeStringsForRun(codeStringToProcess.split(/\r?\n/).filter(line => line.trim() !== '')); 
        
        if (validationErrors && validationErrors.length > 0) { // Check if the array has errors
            // Revert to simpler error message construction assuming an array
            const errorMsg = "Validation failed. Please fix the errors before running:\n" + validationErrors.join("\n");
            
            /* // Remove the complex type checking
            // Check if the result is an array before joining
            if (Array.isArray(validationErrors)) {
                errorMsg += validationErrors.join("\n");
            } else {
                // If not an array, treat as a single message or convert object to string
                errorMsg += String(validationErrors); 
            }
            */
            
            console.error("Code validation failed:", validationErrors);
            showError("Code validation failed. See chat for details.");
            appendMessage(errorMsg); // Show errors in chat
            setButtonLoading(false); // Re-enable button
            return; // Stop execution
        }
        console.log("Code validation successful.");
        // <<< END VALIDATION
        
        const collection = populateCodeCollection(codeStringToProcess);
        console.log(`Collection populated with ${collection.length} code(s)`);

        console.log("Running codes...");
        const runResult = await runCodes(collection);
        console.log("Codes executed:", runResult);

        // Post-processing
        if (runResult && runResult.assumptionTabs && runResult.assumptionTabs.length > 0) {
            console.log("Processing assumption tabs...");
            await processAssumptionTabs(runResult.assumptionTabs);
        }
        console.log("Hiding specific columns and navigating...");
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

// Ensure Office.onReady sets up the button click handler for the REVERTED function
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // ***** ADAPT THIS ID TO YOUR ACTUAL BUTTON *****
    const button = document.getElementById("insert-and-run"); 
    if (button) {
        // Assign the REVERTED async function as the handler
        button.onclick = insertSheetsAndRunCodes; 
    } else {
        // Update the error message to reflect the ID we were looking for
        console.error("Could not find button with id='insert-and-run'");
    }

    // Keep the setup for your other buttons (send-button, reset-button, etc.)
    // Assign event listeners to the buttons
    const sendButton = document.getElementById('send');
    if (sendButton) sendButton.onclick = handleSend;

    const writeButton = document.getElementById('write-to-excel');
    if (writeButton) writeButton.onclick = writeToExcel;
    
    const resetButton = document.getElementById('reset-chat');
    if (resetButton) resetButton.onclick = resetChat;
    
    // Get modal elements
    const codesTextarea = document.getElementById('codes-textarea');
    
    // >>>>>>>>> ADDED: Get references to new modal elements (assuming IDs)
    const codeSearchInput = document.getElementById('code-search-input');
    const codeSuggestionsContainer = document.getElementById('code-suggestions');
    
    // >>>>>>>>> ADDED: Variables for keyboard navigation state
    let highlightedSuggestionIndex = -1;
    let currentSuggestions = []; 

    // >>>>>>>>> ADDED: Helper function to update highlighting
    const updateHighlight = (newIndex) => {
      const suggestionItems = codeSuggestionsContainer.querySelectorAll('.code-suggestion-item');
      if (highlightedSuggestionIndex >= 0 && highlightedSuggestionIndex < suggestionItems.length) {
        suggestionItems[highlightedSuggestionIndex].classList.remove('suggestion-highlight');
      }
      if (newIndex >= 0 && newIndex < suggestionItems.length) {
        suggestionItems[newIndex].classList.add('suggestion-highlight');
        // Optional: Scroll the highlighted item into view
        suggestionItems[newIndex].scrollIntoView({ block: 'nearest' });
      }
      highlightedSuggestionIndex = newIndex;
    };

    // >>>>>>>>> REFACTORED: Function to show suggestions based on a search term
    const showSuggestionsForTerm = (searchTerm) => {
        searchTerm = searchTerm.toLowerCase().trim();
        console.log(`[showSuggestionsForTerm] Search Term: '${searchTerm}'`);

        codeSuggestionsContainer.innerHTML = ''; // Clear previous suggestions
        highlightedSuggestionIndex = -1; 
        currentSuggestions = []; 

        if (searchTerm.length < 2) { 
            console.log("[showSuggestionsForTerm] Search term too short, hiding suggestions.");
            codeSuggestionsContainer.style.display = 'none';
            return;
        }

        console.log("[showSuggestionsForTerm] Filtering code database...");
        const suggestions = codeDatabase
            .filter(item => {
                const hasName = item && typeof item.name === 'string';
                // if (!hasName) console.warn("Skipping item with missing/invalid name:", item);
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
                    let wasAdjusted = false; // Flag to track if insertion point moved
                    const textBeforeCursor = currentText.substring(0, cursorPosition);
                    const lastOpenBracket = textBeforeCursor.lastIndexOf('<');
                    const lastCloseBracket = textBeforeCursor.lastIndexOf('>');
                    
                    if (lastOpenBracket > lastCloseBracket) { 
                        const textAfterCursor = currentText.substring(cursorPosition);
                        const nextCloseBracket = textAfterCursor.indexOf('>');
                        if (nextCloseBracket !== -1) { 
                            insertionPosition = cursorPosition + nextCloseBracket + 1;
                            wasAdjusted = true; // Mark as adjusted
                            console.log(`Cursor inside <>, adjusting insertion point to after > at ${insertionPosition}`);
                        }
                    }

                    const maxNumbers = getMaxDriverNumbers(currentText); 
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

                    // 3. Insert the modified code string, potentially replacing the search term
                    const textAfterInsertion = currentText.substring(insertionPosition);
                    let textBeforeFinal = "";
                    let searchStartIndex = insertionPosition; // Default to insertion point

                    if (!wasAdjusted) {
                        // If not adjusted, try to find and remove the search term
                        const textBeforeInsertion = currentText.substring(0, insertionPosition);
                        const match = textBeforeInsertion.match(/[^<>\s\n]*$/);
                        const searchTermToRemove = match ? match[0] : '';
                        if (searchTermToRemove.length > 0) {
                            searchStartIndex = insertionPosition - searchTermToRemove.length;
                            console.log(`Replacing typed term: '${searchTermToRemove}'`);
                            textBeforeFinal = currentText.substring(0, searchStartIndex);
                        } else {
                            // No term found to remove, behave like simple insertion
                            textBeforeFinal = textBeforeInsertion;
                        }
                    } else {
                        // If adjusted (was inside <>), just insert, don't remove text
                         textBeforeFinal = currentText.substring(0, insertionPosition); 
                    }

                    // Add newlines for better formatting
                    let prefixNewline = '';
                    // Check char before the *final* start position (after potential removal)
                    if (searchStartIndex > 0 && textBeforeFinal.length > 0 && textBeforeFinal[textBeforeFinal.length - 1] !== '\n') {
                        prefixNewline = '\n';
                    }
                    let suffixNewline = '\n';
                    if (insertionPosition === currentText.length || (textAfterInsertion.length > 0 && textAfterInsertion[0] === '\n') ) {
                        suffixNewline = '';
                    }
                    
                    // Construct the new text using textBeforeFinal
                    const newText = textBeforeFinal + prefixNewline + codeToAdd + suffixNewline + textAfterInsertion;
                    codesTextarea.value = newText;
                    
                    // Set cursor position after the inserted text (relative to the final text structure)
                    const newCursorPosition = textBeforeFinal.length + (prefixNewline + codeToAdd + suffixNewline).length;
                    codesTextarea.setSelectionRange(newCursorPosition, newCursorPosition);

                    // Clear search and hide suggestions
                    if (codeSearchInput) codeSearchInput.value = ''; // Clear search bar too
                    codeSuggestionsContainer.innerHTML = '';
                    codeSuggestionsContainer.style.display = 'none';
                    highlightedSuggestionIndex = -1; 
                    currentSuggestions = []; 
                };
                
                suggestionDiv.onmouseover = () => {
                    updateHighlight(i);
                };

                codeSuggestionsContainer.appendChild(suggestionDiv);
            });
            console.log("[showSuggestionsForTerm] Setting suggestions display to 'block'");
            codeSuggestionsContainer.style.display = 'block'; 
        } else {
            console.log("[showSuggestionsForTerm] No suggestions found, hiding container.");
            codeSuggestionsContainer.style.display = 'none'; 
        }
    };
    // >>>>>>>>> END REFACTORED FUNCTION

    // Event listener for search input (This check remains correct)
    if (codeSearchInput && codeSuggestionsContainer) {
        codeSearchInput.oninput = () => {
            // Simply call the refactored function
            showSuggestionsForTerm(codeSearchInput.value);
        };

      // >>>>>>>>> MODIFIED: Keydown listener for navigation (applies to search input)
      codeSearchInput.onkeydown = (event) => {
        if (codeSuggestionsContainer.style.display !== 'block' || currentSuggestions.length === 0) {
          return;
        }

        const suggestionItems = codeSuggestionsContainer.querySelectorAll('.code-suggestion-item');
        let newIndex = highlightedSuggestionIndex;

        switch (event.key) {
          case 'ArrowDown':
            event.preventDefault(); // Prevent cursor moving in input
            newIndex = (highlightedSuggestionIndex + 1) % currentSuggestions.length;
            updateHighlight(newIndex);
            break;

          case 'ArrowUp':
            event.preventDefault(); // Prevent cursor moving in input
            newIndex = (highlightedSuggestionIndex - 1 + currentSuggestions.length) % currentSuggestions.length;
            updateHighlight(newIndex);
            break;

          case 'Enter':
            event.preventDefault(); // Prevent form submission/newline
            // Check if suggestions are visible and there's at least one
            if (codeSuggestionsContainer.style.display === 'block' && currentSuggestions.length > 0) {
                // Find the first suggestion item in the container
                const firstSuggestionItem = codeSuggestionsContainer.querySelector('.code-suggestion-item[data-index="0"]');
                if (firstSuggestionItem) {
                    firstSuggestionItem.click(); // Trigger the click handler for the first item
                }
            }
            break;
          
          case 'Escape': // Optional: Hide suggestions on Escape
             event.preventDefault();
             codeSuggestionsContainer.style.display = 'none';
             highlightedSuggestionIndex = -1;
             currentSuggestions = [];
             break;
        }
      };
    }
    // >>>>>>>>> END ADDED

    // >>>>>>>>> ADDED: Listener for the main codes textarea
    if (codesTextarea && codeSuggestionsContainer) {
        codesTextarea.oninput = () => {
            const cursorPosition = codesTextarea.selectionStart;
            const currentText = codesTextarea.value;

            // Check if cursor is inside <...>
            const textBeforeCursor = currentText.substring(0, cursorPosition);
            const lastOpenBracket = textBeforeCursor.lastIndexOf('<');
            const lastCloseBracket = textBeforeCursor.lastIndexOf('>');
            let isInsideBrackets = false;
            if (lastOpenBracket > lastCloseBracket) { 
                const textAfterCursor = currentText.substring(cursorPosition);
                const nextCloseBracket = textAfterCursor.indexOf('>');
                if (nextCloseBracket !== -1) {
                    isInsideBrackets = true;
                }
            }

            if (isInsideBrackets) {
                // If inside brackets, hide suggestions
                console.log("[Textarea Input] Cursor inside <>, hiding suggestions.");
                codeSuggestionsContainer.style.display = 'none';
                highlightedSuggestionIndex = -1;
                currentSuggestions = [];
            } else {
                // If outside brackets, find the word before cursor
                // Find the start of the current word/segment
                let searchStart = cursorPosition - 1;
                while (searchStart >= 0) {
                    const char = textBeforeCursor[searchStart];
                    // Break on space, newline, or closing bracket
                    if (char === ' ' || char === '\n' || char === '>') {
                        searchStart++; // Start search after the delimiter
                        break;
                    }
                    searchStart--;
                }
                if (searchStart < 0) searchStart = 0; // Handle start of text

                const searchTerm = textBeforeCursor.substring(searchStart, cursorPosition);
                console.log(`[Textarea Input] Cursor outside <>, potential search term: '${searchTerm}'`);
                
                // Show suggestions for this term
                showSuggestionsForTerm(searchTerm);
            }
        };

        // >>>>>>>>> ADDED: Keydown listener for navigation (applies to TEXTAREA)
        codesTextarea.onkeydown = (event) => {
            if (codeSuggestionsContainer.style.display !== 'block' || currentSuggestions.length === 0) {
                return; // Only act if suggestions are visible
            }

            const suggestionItems = codeSuggestionsContainer.querySelectorAll('.code-suggestion-item');
            let newIndex = highlightedSuggestionIndex;

            switch (event.key) {
                case 'ArrowDown':
                case 'ArrowUp':
                    event.preventDefault(); // Prevent cursor moving in textarea
                    if (event.key === 'ArrowDown') {
                        newIndex = (highlightedSuggestionIndex + 1) % currentSuggestions.length;
                    } else {
                        newIndex = (highlightedSuggestionIndex - 1 + currentSuggestions.length) % currentSuggestions.length;
                    }
                    updateHighlight(newIndex);
                    break;

                case 'Enter':
                    event.preventDefault(); // Prevent newline in textarea when selecting
                    if (highlightedSuggestionIndex >= 0 && highlightedSuggestionIndex < suggestionItems.length) {
                        suggestionItems[highlightedSuggestionIndex].click(); // Trigger the click handler
                    }
                    break;
                
                case 'Escape':
                    event.preventDefault();
                    codeSuggestionsContainer.style.display = 'none';
                    highlightedSuggestionIndex = -1;
                    currentSuggestions = [];
                    break;
                
                // Allow other keys (like Tab, Shift, Ctrl, letters, numbers, backspace, delete etc.) to function normally
                default:
                    // If a suggestion is highlighted, typing might clear it
                    // Optionally, you could choose to hide suggestions on typing other chars
                    // codeSuggestionsContainer.style.display = 'none';
                    // highlightedSuggestionIndex = -1;
                    // currentSuggestions = [];
                    break; 
            }
        };
    }
    // >>>>>>>>> END ADDED

    // Test Buttons (Add similar checks if they are essential)
    // const testGreenCellButton = document.getElementById('test-green-cell');

    // Initialize API keys, load history etc.
    Promise.all([ // <<<<< UPDATED: Use Promise.all for parallel loading
        initializeAPIKeys(),
        loadCodeDatabase() // <<<<< ADDED: Load code database during init
    ]).then(([keysLoaded]) => {
      if (!keysLoaded) {
        showError("Failed to load API keys. Please check configuration.");
      }
      // Load conversation history after keys are potentially loaded
      conversationHistory = loadConversationHistory();

      // Load code strings from localStorage into the global variable
      try {
          const storedCodes = localStorage.getItem('userCodeStrings');
          if (storedCodes !== null) {
              // Initialize the global variable
              loadedCodeStrings = storedCodes;
              if (codesTextarea) { // Check if textarea exists
                  codesTextarea.value = loadedCodeStrings;
              }
              console.log("Code strings loaded from localStorage into global variable.");
              if (DEBUG) console.log("Initial loaded codes:", loadedCodeStrings.substring(0, 100) + '...');
          } else {
              console.log("No code strings found in localStorage, initializing global variable as empty.");
              // Initialize the global variable
              loadedCodeStrings = "";
          }
      } catch (error) {
          console.error("Error loading code strings from localStorage:", error);
          showError(`Error loading codes from storage: ${error.message}`);
          // Initialize the global variable
          loadedCodeStrings = "";
      }

    }).catch(error => {
        console.error("Error during initialization:", error);
        showError("Error during initialization: " + error.message);
    });
    
    // Hide sideload message and show app body
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "block";
  }
});

