// AIModelPlanner.js

// Assuming API keys are managed and accessible similarly to AIcalls.js
// Or that callOpenAI is available globally/imported if AIcalls.js exports it.
// For now, let's assume a local way to call OpenAI or that it's handled by the main taskpane script.

let modelPlannerConversationHistory = [];
let AI_MODEL_PLANNER_OPENAI_API_KEY = "";
let lastPlannerResponseForClient = null; // To store the last response for client mode buttons

const DEBUG_PLANNER = true; // For planner-specific debugging

import { getAICallsProcessedResponse } from './AIcalls.js'; // <<< ADDED IMPORT

// Helper function to get API keys (placeholder, adapt as needed based on your structure)
// This might need to be coordinated with how API keys are managed in your main taskpane.js
// For now, we'll assume AIcalls.js might make setAPIKeys and INTERNAL_API_KEYS available or similar.
// As per the prompt, we are not changing AIcalls.js, so we'll define what's needed here.

export function setAIModelPlannerOpenApiKey(key) {
    if (key) {
        AI_MODEL_PLANNER_OPENAI_API_KEY = key;
        if (DEBUG_PLANNER) console.log("AIModelPlanner.js: OpenAI API key set.");
    } else {
        if (DEBUG_PLANNER) console.warn("AIModelPlanner.js: Attempted to set an empty OpenAI API key.");
    }
}


// Updated to use the more robust fetching approach
async function getAIModelPlanningSystemPrompt() {
  const promptKey = "AIModelPlanning_System"; // Key for this specific prompt file
  const paths = [
    // Try path relative to root if /src/ is not working, assuming 'prompts' is then at root level of served dir
    // THIS IS A GUESS - The original path `https://localhost:3002/src/prompts/...` should work if server is configured for it.
    `https://localhost:3002/prompts/${promptKey}.txt`, 
    `https://localhost:3002/src/prompts/${promptKey}.txt` // Original path as a fallback
  ];

  if (DEBUG_PLANNER) console.log(`AIModelPlanner: Attempting to load prompt file: ${promptKey}.txt`);

  let response = null;
  for (const path of paths) {
    if (DEBUG_PLANNER) console.log(`AIModelPlanner: Attempting to load prompt from: ${path}`);
    try {
      response = await fetch(path);
      if (response.ok) {
        if (DEBUG_PLANNER) console.log(`AIModelPlanner: Successfully loaded prompt from: ${path}`);
        const text = await response.text();
        return text;
      } else {
        if (DEBUG_PLANNER) console.warn(`AIModelPlanner: Failed to load from ${path} - Status: ${response.status} ${response.statusText}`);
      }
    } catch (err) {
      if (DEBUG_PLANNER) console.error(`AIModelPlanner: Error fetching from ${path}: ${err.message}`);
    }
  }

  // If all paths fail
  console.error(`AIModelPlanner: Failed to load prompt ${promptKey}.txt from all attempted paths.`);
  // Fallback prompt
  return "You are a helpful assistant for financial model planning. [Error: System prompt AIModelPlanning_System.txt could not be loaded]"; 
}

// Direct OpenAI API call function (simplified version, adapt if AIcalls.js exports its own)
async function callOpenAIForModelPlanner(messages, model = "gpt-4.1", temperature = 0.7) {
  if (!AI_MODEL_PLANNER_OPENAI_API_KEY) {
    console.error("AIModelPlanner: OpenAI API key not set.");
    throw new Error("OpenAI API key not set for AIModelPlanner.");
  }

  if (DEBUG_PLANNER) {
    if (messages && messages.length > 0) {
      const systemMessage = messages.find(msg => msg.role === 'system');
      const userMessages = messages.filter(msg => msg.role === 'user');
      const lastUserMessage = userMessages.length > 0 ? userMessages[userMessages.length - 1] : null;

      if (systemMessage) {
        console.log("AIModelPlanner API Call: System Prompt:", systemMessage.content);
      } else {
        console.warn("AIModelPlanner API Call: No system prompt found in messages.");
      }
      if (lastUserMessage) {
        console.log("AIModelPlanner API Call: Main User Prompt:", lastUserMessage.content);
      } else {
        console.warn("AIModelPlanner API Call: No user prompt found in messages.");
      }
    } else {
      console.warn("AIModelPlanner API Call: Messages array is empty or undefined.");
    }
  }

  try {
    const response = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${AI_MODEL_PLANNER_OPENAI_API_KEY}`
      },
      body: JSON.stringify({
        model: model,
        messages: messages,
        temperature: temperature
      })
    });

    if (!response.ok) {
      const errorData = await response.json().catch(() => null);
      console.error("AIModelPlanner - OpenAI API error response:", errorData);
      throw new Error(`OpenAI API error: ${response.status} ${response.statusText}`);
    }

    const data = await response.json();
    return data.choices[0].message.content;
  } catch (error) {
    console.error("Error calling OpenAI API in AIModelPlanner:", error);
    throw error;
  }
}


// Function to process a prompt for the AI Model Planner
async function processAIModelPlannerPromptInternal({ userInput, systemPrompt, model, temperature, history = [] }) {
    const messages = [
        { role: "system", content: systemPrompt }
    ];

    if (history.length > 0) {
        history.forEach(message => {
             if (Array.isArray(message) && message.length === 2) {
                 messages.push({
                     role: message[0] === "human" ? "user" : "assistant",
                     content: message[1]
                 });
             } else {
                 console.warn("AIModelPlanner: Skipping malformed history message:", message);
             }
        });
    }

    messages.push({ role: "user", content: userInput });

    try {
        const responseContent = await callOpenAIForModelPlanner(messages, model, temperature);
        // The prompt asks for JSON output in the final step.
        // For intermediate steps, it might be text. We need to handle both.
        try {
            // Try to parse as JSON. If it works, and it's the final step format, return as object.
            const parsedJson = JSON.parse(responseContent);
            // If it's an object (likely the final JSON output), return it directly
            if (typeof parsedJson === 'object' && parsedJson !== null) {
                return parsedJson; 
            }
            // If it parsed to something else (e.g. a string that happened to be valid JSON), treat as text.
            return responseContent.split('\n').filter(line => line.trim());

        } catch (e) {
            // If JSON parsing fails, it's likely a text response.
            return responseContent.split('\n').filter(line => line.trim());
        }
    } catch (error) {
        console.error("Error in processAIModelPlannerPrompt:", error);
        throw error; 
    }
}

// Handle initial conversation for AI Model Planner
async function handleInitialAIModelPlannerConversation(userInput) {
    console.log("AIModelPlanner: Processing initial question:", userInput);
    
    const systemPrompt = await getAIModelPlanningSystemPrompt();
    if (!systemPrompt) {
        throw new Error("Failed to load the AI Model Planning system prompt.");
    }

    const model = "gpt-4.1"; // As specified
    const temperature = 0.7; // A reasonable default

    const response = await processAIModelPlannerPromptInternal({
        userInput: userInput,
        systemPrompt: systemPrompt,
        model: model,
        temperature: temperature,
        history: [] 
    });
    
    // Update history
    // The prompt implies the final output is JSON, but intermediate steps are text.
    let assistantResponseContent = "";
    if (typeof response === 'object') {
        assistantResponseContent = JSON.stringify(response); // Store JSON as string in history
    } else if (Array.isArray(response)) {
        assistantResponseContent = response.join("\n");
    } else {
        assistantResponseContent = String(response);
    }
    
    modelPlannerConversationHistory = [
        ["human", userInput],
        ["assistant", assistantResponseContent]
    ];
    
    console.log("AIModelPlanner: Initial conversation processed.");
    return { response: response, history: modelPlannerConversationHistory };
}

// Handle follow-up conversation for AI Model Planner
async function handleFollowUpAIModelPlannerConversation(userInput, currentHistory) {
    console.log("AIModelPlanner: Processing follow-up question:", userInput);
    
    const systemPrompt = await getAIModelPlanningSystemPrompt();
    if (!systemPrompt) {
        throw new Error("Failed to load the AI Model Planning system prompt.");
    }

    const model = "gpt-4.1"; // As specified
    const temperature = 0.7; // A reasonable default

    const response = await processAIModelPlannerPromptInternal({
        userInput: userInput,
        systemPrompt: systemPrompt,
        model: model,
        temperature: temperature,
        history: currentHistory
    });

    let assistantResponseContent = "";
    if (typeof response === 'object') {
        assistantResponseContent = JSON.stringify(response);
    } else if (Array.isArray(response)) {
        assistantResponseContent = response.join("\n");
    } else {
        assistantResponseContent = String(response);
    }

    const updatedHistory = [
        ...currentHistory,
        ["human", userInput],
        ["assistant", assistantResponseContent]
    ];
    modelPlannerConversationHistory = updatedHistory;
    
    console.log("AIModelPlanner: Follow-up conversation processed.");
    return { response: response, history: updatedHistory };
}

// Main conversation handler for AI Model Planner
export async function handleAIModelPlannerConversation(userInput) {
    try {
        const isFollowUp = modelPlannerConversationHistory.length > 0;
        if (isFollowUp) {
            return await handleFollowUpAIModelPlannerConversation(userInput, modelPlannerConversationHistory);
        } else {
            return await handleInitialAIModelPlannerConversation(userInput);
        }
    } catch (error) {
        console.error("Error in AI Model Planner conversation handling:", error);
        return {
            // Return response as an array of strings for text, or object for JSON
            response: (typeof error.message === 'string') ? [error.message] : error.message,
            history: modelPlannerConversationHistory
        };
    }
}

// Function to clear conversation history for the model planner
export function resetAIModelPlannerConversation() {
    modelPlannerConversationHistory = [];
    console.log("AIModelPlanner: Conversation history reset.");
}

// UI Helper functions specific to AIModelPlanner controlling client chat
function displayInClientChatLogPlanner(message, isUser) {
    const chatLog = document.getElementById('chat-log-client');
    const welcomeMessage = document.getElementById('welcome-message-client');
    if (!chatLog) { console.error("AIModelPlanner: Client chat log element not found."); return; }
    if (welcomeMessage) welcomeMessage.style.display = 'none';

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

function setClientLoadingStatePlanner(isLoading) {
    const sendButton = document.getElementById('send-client');
    const loadingAnimation = document.getElementById('loading-animation-client');
    if (sendButton) sendButton.disabled = isLoading;
    if (loadingAnimation) loadingAnimation.style.display = isLoading ? 'flex' : 'none';
}

// Core conversation logic, now private to this module
async function _handleAIModelPlannerConversation(userInput) {
    const systemPrompt = await getAIModelPlanningSystemPrompt();
    if (!systemPrompt) throw new Error("Failed to load AI Model Planning system prompt.");

    const isFollowUp = modelPlannerConversationHistory.length > 0;
    const model = "gpt-4.1";
    const temperature = 0.7;

    const response = await processAIModelPlannerPromptInternal({
        userInput: userInput,
        systemPrompt: systemPrompt,
        model: model,
        temperature: temperature,
        history: isFollowUp ? modelPlannerConversationHistory : []
    });

    let assistantResponseContent = "";
    if (typeof response === 'object') assistantResponseContent = JSON.stringify(response);
    else if (Array.isArray(response)) assistantResponseContent = response.join("\n");
    else assistantResponseContent = String(response);

    if (isFollowUp) {
        modelPlannerConversationHistory.push(["human", userInput], ["assistant", assistantResponseContent]);
    } else {
        modelPlannerConversationHistory = [["human", userInput], ["assistant", assistantResponseContent]];
    }
    return response; // Return the direct response (object or array)
}

// Exported Event Handlers for Client Mode UI
export async function plannerHandleSend() {
    const userInputElement = document.getElementById('user-input-client');
    if (!userInputElement) { console.error("AIModelPlanner: Client user input element not found."); return; }
    const userInput = userInputElement.value.trim();

    if (!userInput) {
        alert('Please enter a request (Client Mode)');
        return;
    }

    displayInClientChatLogPlanner(userInput, true);
    userInputElement.value = '';
    setClientLoadingStatePlanner(true);

    try {
        const resultResponse = await _handleAIModelPlannerConversation(userInput);
        lastPlannerResponseForClient = resultResponse; // Store for other buttons

        let jsonObjectToProcess = null;

        if (typeof resultResponse === 'string') {
            try {
                const parsedResponse = JSON.parse(resultResponse);
                if (typeof parsedResponse === 'object' && parsedResponse !== null && !Array.isArray(parsedResponse)) {
                    jsonObjectToProcess = parsedResponse;
                    console.log("test passed (JSON string successfully parsed to an object for tab processing)");
                }
            } catch (e) {
                // Not a JSON string, or parsing did not result in a suitable object
            }
        } else if (typeof resultResponse === 'object' && resultResponse !== null && !Array.isArray(resultResponse)) {
            jsonObjectToProcess = resultResponse;
            console.log("test passed (response is already a suitable JSON object for tab processing)");
        }

        if (jsonObjectToProcess) {
            let ModelCodes = ""; 
            console.log("AIModelPlanner: Starting to process JSON object for ModelCodes generation.");

            for (const tabLabel in jsonObjectToProcess) {
                if (Object.prototype.hasOwnProperty.call(jsonObjectToProcess, tabLabel)) {
                    ModelCodes += `<TAB; label1="${tabLabel}";>\n`; 

                    const tabDescription = jsonObjectToProcess[tabLabel];
                    let tabDescriptionString = "";

                    if (typeof tabDescription === 'string') {
                        tabDescriptionString = tabDescription;
                    } else if (typeof tabDescription === 'object' && tabDescription !== null) {
                        tabDescriptionString = JSON.stringify(tabDescription);
                    } else {
                        tabDescriptionString = String(tabDescription);
                    }
                    
                    if (tabDescriptionString.trim() !== "") {
                        console.log(`AIModelPlanner: Submitting description for tab "${tabLabel}" to getAICallsProcessedResponse...`);
                        try {
                            const aiResponseForTabArray = await getAICallsProcessedResponse(tabDescriptionString);
                            
                            let formattedAiResponse = "";
                            if (typeof aiResponseForTabArray === 'object' && aiResponseForTabArray !== null && !Array.isArray(aiResponseForTabArray)) {
                                formattedAiResponse = JSON.stringify(aiResponseForTabArray, null, 2); 
                            } else if (Array.isArray(aiResponseForTabArray)) {
                                formattedAiResponse = aiResponseForTabArray.join('\n');
                            } else {
                                formattedAiResponse = String(aiResponseForTabArray);
                            }
                            
                            ModelCodes += formattedAiResponse + "\n\n"; 
                            console.log(`AIModelPlanner: Received and appended AI response for tab "${tabLabel}"`);
                        } catch (tabError) {
                            console.error(`AIModelPlanner: Error processing description for tab "${tabLabel}" via getAICallsProcessedResponse:`, tabError);
                            ModelCodes += `// Error processing tab ${tabLabel}: ${tabError.message}\n\n`;
                        }
                    } else {
                        ModelCodes += `// No description provided for tab ${tabLabel}\n\n`;
                    }
                }
            }
            console.log("Generated ModelCodes (final):\n" + ModelCodes); 
        }

        displayInClientChatLogPlanner(resultResponse, false);
    } catch (error) {
        console.error("Error in AIModelPlanner conversation:", error);
        displayInClientChatLogPlanner(`Error: ${error.message}`, false);
    } finally {
        setClientLoadingStatePlanner(false);
    }
}

export function plannerHandleReset() {
    const chatLog = document.getElementById('chat-log-client');
    const welcomeMessage = document.getElementById('welcome-message-client');
    if (chatLog) {
        chatLog.innerHTML = '';
        if (welcomeMessage) {
            const newWelcome = document.createElement('div');
            newWelcome.id = 'welcome-message-client';
            newWelcome.className = 'welcome-message';
            newWelcome.innerHTML = '<h1>Ask me anything (Client Mode)</h1>';
            chatLog.appendChild(newWelcome);
        }
    }
    modelPlannerConversationHistory = [];
    lastPlannerResponseForClient = null;
    const userInputElement = document.getElementById('user-input-client');
    if (userInputElement) userInputElement.value = '';
    console.log("AIModelPlanner: Client chat and history reset.");
}

export function plannerHandleWriteToExcel() {
    if (!lastPlannerResponseForClient) {
        displayInClientChatLogPlanner("No response to write to Excel.", false);
        return;
    }
    let contentToWrite = "";
    if (typeof lastPlannerResponseForClient === 'object') contentToWrite = JSON.stringify(lastPlannerResponseForClient, null, 2);
    else if (Array.isArray(lastPlannerResponseForClient)) contentToWrite = lastPlannerResponseForClient.join("\n");
    else contentToWrite = String(lastPlannerResponseForClient);
    
    console.log("AIModelPlanner Client Mode - Write to Excel (Placeholder):\n", contentToWrite);
    displayInClientChatLogPlanner("Write to Excel (Placeholder): Response logged to console. Actual Excel writing depends on format.", false);
    // Actual Excel.run call would be complex here, especially for JSON, and is out of scope of this file's direct responsibility as per prompt.
}

export function plannerHandleInsertToEditor() {
    if (!lastPlannerResponseForClient) {
        displayInClientChatLogPlanner("No response to insert into editor.", false);
        return;
    }
    let contentToInsert = "";
    if (typeof lastPlannerResponseForClient === 'object') contentToInsert = JSON.stringify(lastPlannerResponseForClient, null, 2);
    else if (Array.isArray(lastPlannerResponseForClient)) contentToInsert = lastPlannerResponseForClient.join("\n");
    else contentToInsert = String(lastPlannerResponseForClient);

    console.log("AIModelPlanner Client Mode - Insert to Editor (Placeholder):\n", contentToInsert);
    displayInClientChatLogPlanner("Insert to Editor (Placeholder): Response logged to console. Actual editor insertion depends on editor availability in client mode.", false);
    // Actual editor insertion logic is out of scope here.
}


