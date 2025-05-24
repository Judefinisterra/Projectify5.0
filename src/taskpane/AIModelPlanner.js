// AIModelPlanner.js

// Assuming API keys are managed and accessible similarly to AIcalls.js
// Or that callOpenAI is available globally/imported if AIcalls.js exports it.
// For now, let's assume a local way to call OpenAI or that it's handled by the main taskpane script.

// Imports needed for _executePlannerCodes
import { validateCodeStringsForRun } from './Validation.js';
import { 
    populateCodeCollection, 
    runCodes, 
    processAssumptionTabs, 
    hideColumnsAndNavigate, 
    handleInsertWorksheetsFromBase64 
} from './CodeCollection.js';
import { getAICallsProcessedResponse } from './AIcalls.js';
// We don't import from taskpane.js to avoid cycles

let modelPlannerConversationHistory = [];
let AI_MODEL_PLANNER_OPENAI_API_KEY = "";
let lastPlannerResponseForClient = null; // To store the last response for client mode buttons

const DEBUG_PLANNER = true; // For planner-specific debugging

import { processModelCodesForPlanner } from './taskpane.js'; // <<< UPDATED IMPORT

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
async function* callOpenAIForModelPlanner(messages, options = {}) {
  const { model = "gpt-4.1", temperature = 0.7, stream = false } = options;

  if (!AI_MODEL_PLANNER_OPENAI_API_KEY) {
    console.error("AIModelPlanner: OpenAI API key not set.");
    throw new Error("OpenAI API key not set for AIModelPlanner.");
  }

  if (DEBUG_PLANNER) {
    // Condensed logging
    const systemMessageContent = messages.find(msg => msg.role === 'system')?.content?.substring(0, 100) + "...";
    const lastUserMessageContent = messages.filter(msg => msg.role === 'user').pop()?.content?.substring(0, 100) + "...";
    console.log(`AIModelPlanner API Call: Model: ${model}, Stream: ${stream}`);
    console.log("AIModelPlanner API Call: System Prompt (start):", systemMessageContent || "N/A");
    console.log("AIModelPlanner API Call: Last User Prompt (start):", lastUserMessageContent || "N/A");
  }

  try {
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
        'Authorization': `Bearer ${AI_MODEL_PLANNER_OPENAI_API_KEY}`
      },
      body: JSON.stringify(body)
    });

    if (!response.ok) {
      const errorData = await response.json().catch(() => ({ message: "Failed to parse error JSON." }));
      console.error("AIModelPlanner - OpenAI API error response:", errorData);
      throw new Error(`OpenAI API error: ${response.status} ${response.statusText} - ${errorData.message || JSON.stringify(errorData)}`);
    }

    if (stream) {
      console.log("AIModelPlanner - OpenAI API response received (stream)");
      const reader = response.body.getReader();
      const decoder = new TextDecoder("utf-8");

      while (true) {
        const { done, value } = await reader.read();
        if (done) {
          console.log("AIModelPlanner - Stream finished.");
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
              console.warn("AIModelPlanner - Could not parse JSON line from stream:", line, e);
              return null;
            }
          })
          .filter(line => line !== null);

        for (const parsedLine of parsedLines) {
          yield parsedLine;
        }
      }
    } else {
      // Non-streaming path (kept for potential compatibility, though planner chat will use stream)
      const data = await response.json();
      console.log("AIModelPlanner - OpenAI API response received (non-stream)");
      if (data.choices && data.choices[0] && data.choices[0].message) {
        return data.choices[0].message.content;
      } else {
        console.error("AIModelPlanner - Invalid non-stream response structure:", data);
        throw new Error("Invalid response structure from OpenAI (non-stream).");
      }
    }
  } catch (error) {
    console.error("Error calling OpenAI API in AIModelPlanner:", error);
    if (!stream) throw error; // Re-throw for non-streaming errors
    // For streams, error breaks the generator.
  }
}


// Function to process a prompt for the AI Model Planner
async function* processAIModelPlannerPromptInternal({ userInput, systemPrompt, model, temperature, history = [], stream = false }) {
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
        // Pass the stream option to callOpenAIForModelPlanner
        const streamResponse = callOpenAIForModelPlanner(messages, { model, temperature, stream });
        
        if (stream) {
            // If streaming, yield all chunks from the response stream
            for await (const chunk of streamResponse) {
                yield chunk;
            }
        } else {
            // If not streaming, get the full response content (original behavior for non-stream callers)
            // This part assumes callOpenAIForModelPlanner returns content directly when not streaming.
            const responseContent = await streamResponse; // This will await the promise for non-streaming path
            // The prompt asks for JSON output in the final step.
            // For intermediate steps, it might be text. We need to handle both.
            try {
                const parsedJson = JSON.parse(responseContent);
                if (typeof parsedJson === 'object' && parsedJson !== null) {
                    yield parsedJson; // Yield the single parsed object
                    return;
                }
                yield responseContent.split('\n').filter(line => line.trim());
                return;

            } catch (e) {
                yield responseContent.split('\n').filter(line => line.trim());
                return;
            }
        }
    } catch (error) {
        console.error("Error in processAIModelPlannerPromptInternal:", error);
        // For streams, the error should propagate from callOpenAIForModelPlanner
        // For non-streams, rethrow or yield an error structure
        if (stream) {
             // Error already logged by callOpenAIForModelPlanner, generator will break
        } else {
            throw error; 
        }
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
    console.log("[displayInClientChatLogPlanner] Called with message:", message.substring(0, 50) + "...");
    
    let chatLog = document.getElementById('chat-log-client');
    const welcomeMessage = document.getElementById('welcome-message-client');
    
    // If chat log doesn't exist, try to find the container and create it
    if (!chatLog) {
        console.error("AIModelPlanner: Client chat log element not found. Attempting to create it...");
        const container = document.getElementById('client-chat-container');
        if (container) {
            chatLog = document.createElement('div');
            chatLog.id = 'chat-log-client';
            chatLog.className = 'chat-log';
            chatLog.style.display = 'block';
            chatLog.style.flexGrow = '1';
            chatLog.style.overflowY = 'auto';
            container.appendChild(chatLog);
            console.log("AIModelPlanner: Created chat log element");
        } else {
            console.error("AIModelPlanner: Could not find container to create chat log");
            return;
        }
    }
    
    if (welcomeMessage) welcomeMessage.style.display = 'none';

    // >>> ADDED: Ensure chat log is visible
    const chatLogDisplay = window.getComputedStyle(chatLog).display;
    if (chatLogDisplay === 'none') {
        console.warn("[displayInClientChatLogPlanner] Chat log was hidden, forcing visibility");
        chatLog.style.display = 'block';
        chatLog.style.visibility = 'visible';
        chatLog.style.opacity = '1';
        
        // Also ensure parent container is visible
        const container = chatLog.parentElement;
        if (container && container.classList.contains('container')) {
            container.classList.add('conversation-active');
        }
    }
    // <<< END ADDED

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
    console.log(`[displayInClientChatLogPlanner] Added ${isUser ? 'user' : 'assistant'} message to chat log. Total messages: ${chatLog.children.length}`);
}

function setClientLoadingStatePlanner(isLoading) {
    const sendButton = document.getElementById('send-client');
    const loadingAnimation = document.getElementById('loading-animation-client');
    if (sendButton) sendButton.disabled = isLoading;
    if (loadingAnimation) loadingAnimation.style.display = isLoading ? 'flex' : 'none';
}

// Core conversation logic, now private to this module
async function* _handleAIModelPlannerConversation(userInput, options = {}) {
    const { stream = false } = options;
    const systemPrompt = await getAIModelPlanningSystemPrompt();
    if (!systemPrompt) throw new Error("Failed to load AI Model Planning system prompt.");

    const isFollowUp = modelPlannerConversationHistory.length > 0;
    const model = "gpt-4.1"; // Or get from options if you want to make it configurable here
    const temperature = 0.7;   // Or get from options

    // Pass the stream option down to processAIModelPlannerPromptInternal
    const streamResponse = processAIModelPlannerPromptInternal({
        userInput: userInput,
        systemPrompt: systemPrompt,
        model: model,
        temperature: temperature,
        history: isFollowUp ? modelPlannerConversationHistory : [],
        stream: stream // Pass the stream flag
    });

    if (stream) {
        let fullAssistantResponseContent = "";
        for await (const chunk of streamResponse) {
            yield chunk; // Yield the raw chunk for UI streaming
            if (chunk.choices && chunk.choices[0]?.delta?.content) {
                fullAssistantResponseContent += chunk.choices[0].delta.content;
            }
        }
        // After stream, update history with the fully accumulated content
        // Note: The original _handleAIModelPlannerConversation returned the response object/array directly.
        // For streaming, the primary output is the stream itself. The full content is for history.
        // The caller (plannerHandleSend) will now be responsible for final parsing if it was JSON.
        if (isFollowUp) {
            modelPlannerConversationHistory.push(["human", userInput], ["assistant", fullAssistantResponseContent]);
        } else {
            modelPlannerConversationHistory = [["human", userInput], ["assistant", fullAssistantResponseContent]];
        }
        // We don't explicitly return fullAssistantResponseContent here because the generator yields chunks.
        // The caller will accumulate it if needed (which plannerHandleSend will do).
    } else {
        // Non-streaming: Get the single yielded item (which is the full response)
        let fullResponse;
        for await (const item of streamResponse) { // Will iterate once for non-streaming
            fullResponse = item;
            break;
        }
        
        let assistantResponseContent = "";
        if (typeof fullResponse === 'object') assistantResponseContent = JSON.stringify(fullResponse);
        else if (Array.isArray(fullResponse)) assistantResponseContent = fullResponse.join("\n");
        else assistantResponseContent = String(fullResponse);

        if (isFollowUp) {
            modelPlannerConversationHistory.push(["human", userInput], ["assistant", assistantResponseContent]);
        } else {
            modelPlannerConversationHistory = [["human", userInput], ["assistant", assistantResponseContent]];
        }
        yield fullResponse; // Yield the full response once for non-streaming callers
    }
}

async function _executePlannerCodes(modelCodesString) {
    console.log(`[AIModelPlanner._executePlannerCodes] Called.`);

    // Substitute <BR> with <BR; labelRow=""; row1 = "||||||||||||";>
    if (modelCodesString && typeof modelCodesString === 'string') {
        modelCodesString = modelCodesString.replace(/<BR>/g, '<BR; labelRow=""; row1 = "||||||||||||";>');
    }

    if (!modelCodesString || modelCodesString.trim().length === 0) {
        console.log("[AIModelPlanner._executePlannerCodes] No model codes to process.");
        displayInClientChatLogPlanner("No code content generated to apply to workbook.", false);
        return;
    }

    let runResult = null;

    try {
        await Excel.run(async (context) => {
            context.application.calculationMode = Excel.CalculationMode.manual;
            await context.sync();
            console.log("[AIModelPlanner._executePlannerCodes] Calculation mode set to manual.");
        });

        console.log("[AIModelPlanner._executePlannerCodes] Validating ALL codes...");
        const validationErrors = await validateCodeStringsForRun(modelCodesString.split(/\r?\n/).filter(line => line.trim() !== ''));
        if (validationErrors && validationErrors.length > 0) {
            const errorMsg = "Code validation failed for planner-generated codes:\n" + validationErrors.join("\n");
            console.error("[AIModelPlanner._executePlannerCodes] Code validation failed:", validationErrors);
            throw new Error(errorMsg);
        }
        console.log("[AIModelPlanner._executePlannerCodes] Code validation successful.");

        console.log("[AIModelPlanner._executePlannerCodes] Inserting base sheets from Worksheets_4.3.25 v1.xlsx...");
        const worksheetsResponse = await fetch('https://localhost:3002/assets/Worksheets_4.3.25 v1.xlsx');
        if (!worksheetsResponse.ok) throw new Error(`[AIModelPlanner._executePlannerCodes] Worksheets_4.3.25 v1.xlsx load failed: ${worksheetsResponse.statusText}`);
        const wsArrayBuffer = await worksheetsResponse.arrayBuffer();
        const wsUint8Array = new Uint8Array(wsArrayBuffer);
        let wsBinaryString = '';
        for (let i = 0; i < wsUint8Array.length; i += 8192) {
            wsBinaryString += String.fromCharCode.apply(null, wsUint8Array.slice(i, Math.min(i + 8192, wsUint8Array.length)));
        }
        await handleInsertWorksheetsFromBase64(btoa(wsBinaryString));
        console.log("[AIModelPlanner._executePlannerCodes] Base sheets (Worksheets_4.3.25 v1.xlsx) inserted.");

        console.log("[AIModelPlanner._executePlannerCodes] Inserting codes.xlsx...");
        const codesResponse = await fetch('https://localhost:3002/assets/codes.xlsx');
        if (!codesResponse.ok) throw new Error(`[AIModelPlanner._executePlannerCodes] codes.xlsx load failed: ${codesResponse.statusText}`);
        const codesArrayBuffer = await codesResponse.arrayBuffer();
        const codesUint8Array = new Uint8Array(codesArrayBuffer);
        let codesBinaryString = '';
        for (let i = 0; i < codesUint8Array.length; i += 8192) {
            codesBinaryString += String.fromCharCode.apply(null, codesUint8Array.slice(i, Math.min(i + 8192, codesUint8Array.length)));
        }
        await handleInsertWorksheetsFromBase64(btoa(codesBinaryString), ["Codes"]); 
        console.log("[AIModelPlanner._executePlannerCodes] codes.xlsx sheets inserted/updated.");
    
        console.log("[AIModelPlanner._executePlannerCodes] Populating collection...");
        const collection = populateCodeCollection(modelCodesString);
        console.log(`[AIModelPlanner._executePlannerCodes] Collection populated with ${collection.length} code(s)`);

        if (collection.length > 0) {
            console.log("[AIModelPlanner._executePlannerCodes] Running codes...");
            runResult = await runCodes(collection);
            console.log("[AIModelPlanner._executePlannerCodes] runCodes executed. Result:", runResult);
        } else {
            console.log("[AIModelPlanner._executePlannerCodes] Collection is empty after population, skipping runCodes execution.");
            runResult = { assumptionTabs: [] };
        }

        console.log("[AIModelPlanner._executePlannerCodes] Starting post-processing steps...");
        if (runResult && runResult.assumptionTabs && runResult.assumptionTabs.length > 0) {
            console.log("[AIModelPlanner._executePlannerCodes] Processing assumption tabs...");
            await processAssumptionTabs(runResult.assumptionTabs);
        } else {
            console.log("[AIModelPlanner._executePlannerCodes] No assumption tabs to process from runResult.");
        }

        console.log("[AIModelPlanner._executePlannerCodes] Hiding specific columns and navigating...");
        await hideColumnsAndNavigate(runResult?.assumptionTabs || []);

        console.log("[AIModelPlanner._executePlannerCodes] Deleting Codes sheet...");
        await Excel.run(async (context) => {
            try {
                context.workbook.worksheets.getItem("Codes").delete();
                console.log("[AIModelPlanner._executePlannerCodes] Codes sheet deleted.");
            } catch (e) {
                if (e instanceof OfficeExtension.Error && e.code === Excel.ErrorCodes.itemNotFound) {
                    console.warn("[AIModelPlanner._executePlannerCodes] Codes sheet not found during cleanup, skipping deletion.");
                } else { 
                    console.error("[AIModelPlanner._executePlannerCodes] Error deleting Codes sheet during cleanup:", e);
                }
            }
            await context.sync();
        }).catch(error => { 
            console.error("[AIModelPlanner._executePlannerCodes] Error during Codes sheet cleanup sync:", error);
        });

        displayInClientChatLogPlanner("Workbook updated with the generated model structure.", false);
        console.log("[AIModelPlanner._executePlannerCodes] Successfully completed.");

    } catch (error) {
        console.error("[AIModelPlanner._executePlannerCodes] Error during processing:", error);
        displayInClientChatLogPlanner(`Error applying model structure: ${error.message}`, false);
        // No re-throw here, error is displayed in chat by this function's caller (plannerHandleSend)
    } finally {
        try {
            await Excel.run(async (context) => {
                context.application.calculationMode = Excel.CalculationMode.automatic;
                await context.sync();
                console.log("[AIModelPlanner._executePlannerCodes] Calculation mode set to automatic.");
            });
        } catch (finalError) {
            console.error("[AIModelPlanner._executePlannerCodes] Error setting calculation mode to automatic:", finalError);
        }
    }
}

export async function plannerHandleSend() {
    console.log("[plannerHandleSend] Function called - VERSION 2"); // Debug to ensure new code is running
    
    const userInputElement = document.getElementById('user-input-client');
    if (!userInputElement) { console.error("AIModelPlanner: Client user input element not found."); return; }
    const userInput = userInputElement.value.trim();

    if (!userInput) {
        alert('Please enter a request (Client Mode)');
        return;
    }

    // >>> ADDED: Hide header and activate conversation layout on first message
    const clientHeader = document.getElementById('client-mode-header');
    const clientChatContainer = document.getElementById('client-chat-container');
    const chatLogClient = document.getElementById('chat-log-client');
    
    console.log("[plannerHandleSend] Initial state - chatLogClient display:", chatLogClient ? window.getComputedStyle(chatLogClient).display : "Element not found");
    
    // Debug parent elements
    if (clientChatContainer) {
        console.log("[plannerHandleSend] clientChatContainer display:", window.getComputedStyle(clientChatContainer).display);
    }
    const clientModeView = document.getElementById('client-mode-view');
    if (clientModeView) {
        console.log("[plannerHandleSend] client-mode-view display:", window.getComputedStyle(clientModeView).display);
    }
    
    // Force immediate visibility of chat container
    if (clientChatContainer) {
        clientChatContainer.classList.add('conversation-active');
        console.log("[plannerHandleSend] Added conversation-active class to container");
        
        // Force a reflow to ensure the class is applied
        void clientChatContainer.offsetHeight;
    }
    
    if (clientHeader && !clientHeader.classList.contains('hidden')) {
        clientHeader.classList.add('hidden');
        // After transition, completely hide it
        setTimeout(() => {
            clientHeader.style.display = 'none';
        }, 300);
    }
    
    if (chatLogClient) {
        // Remove all inline styles that might be hiding it
        chatLogClient.style.cssText = '';
        // Then explicitly set the needed styles
        chatLogClient.style.flexGrow = '1';
        chatLogClient.style.overflowY = 'auto';
        chatLogClient.style.borderBottom = '1px solid #eee';
        chatLogClient.style.marginBottom = '10px';
        chatLogClient.style.display = 'block';
        chatLogClient.style.visibility = 'visible';
        chatLogClient.style.opacity = '1';
        
        console.log("[plannerHandleSend] Forcefully set all chat log styles");
        
        // Force a reflow
        void chatLogClient.offsetHeight;
        
        // Double check the computed style
        const computedDisplay = window.getComputedStyle(chatLogClient).display;
        console.log("[plannerHandleSend] After forcing - chatLogClient computed display:", computedDisplay);
    }
    // <<< END ADDED

    displayInClientChatLogPlanner(userInput, true);
    userInputElement.value = '';
    setClientLoadingStatePlanner(true);

    // Create assistant message elements for streaming
    const welcomeMessageClient = document.getElementById('welcome-message-client');
    if (welcomeMessageClient) welcomeMessageClient.style.display = 'none';

    const assistantMessageDiv = document.createElement('div');
    assistantMessageDiv.className = 'chat-message assistant-message';
    const assistantMessageContent = document.createElement('p');
    assistantMessageContent.className = 'message-content';
    assistantMessageContent.textContent = ''; // Start empty
    assistantMessageDiv.appendChild(assistantMessageContent);
    if (chatLogClient) {
        chatLogClient.appendChild(assistantMessageDiv);
        chatLogClient.scrollTop = chatLogClient.scrollHeight;
    } else {
        console.error("AIModelPlanner: Client chat log element not found for streaming message.");
    }

    let fullAssistantTextResponse = "";

    try {
        // Call _handleAIModelPlannerConversation with stream option
        const stream = _handleAIModelPlannerConversation(userInput, { stream: true });

        for await (const chunk of stream) {
            if (chunk.choices && chunk.choices[0]?.delta?.content) {
                const content = chunk.choices[0].delta.content;
                fullAssistantTextResponse += content;
                assistantMessageContent.textContent += content;
                if (chatLogClient) chatLogClient.scrollTop = chatLogClient.scrollHeight;
            }
        }
        
        // At this point, fullAssistantTextResponse contains the complete text from the LLM.
        // The conversation history has been updated inside _handleAIModelPlannerConversation's streaming path.

        lastPlannerResponseForClient = fullAssistantTextResponse; // Store the raw text or try to parse if always JSON

        // Now, attempt to parse the fullAssistantTextResponse as JSON for further processing
        // This mirrors the previous logic but operates on the accumulated streamed text.
        let jsonObjectToProcess = null;
        if (fullAssistantTextResponse) {
            try {
                const parsedResponse = JSON.parse(fullAssistantTextResponse);
                if (typeof parsedResponse === 'object' && parsedResponse !== null && !Array.isArray(parsedResponse)) {
                    jsonObjectToProcess = parsedResponse;
                    lastPlannerResponseForClient = parsedResponse; // Update to store parsed object
                    console.log("AIModelPlanner: Streamed text successfully parsed to an object for tab processing.");
                }
            } catch (e) {
                // Not a JSON string, or parsing did not result in a suitable object.
                // The UI already has the full text. lastPlannerResponseForClient remains the text.
                console.log("AIModelPlanner: Streamed text was not a parsable JSON object. Displaying as text.");
            }
        }
        
        // If UI was updated with text and it turned out to be JSON object for processing, 
        // we might want to clear/replace the text content with a message like "Processing JSON..."
        // or just let the text remain. For now, text remains.
        // If it wasn't JSON, the text is already correctly displayed.

        if (jsonObjectToProcess) {
            let ModelCodes = ""; 
            console.log("AIModelPlanner: Starting to process JSON object for ModelCodes generation.");
            // Update UI to indicate processing of JSON (optional)
            // assistantMessageContent.textContent = "Processing received plan..."; 
            // displayInClientChatLogPlanner("Generating model structure from AI response...", false); //This would add a new bubble

            for (const tabLabel in jsonObjectToProcess) {
                if (Object.prototype.hasOwnProperty.call(jsonObjectToProcess, tabLabel)) {
                    const lowerCaseTabLabel = tabLabel.toLowerCase();
                    if (lowerCaseTabLabel === "financials" || lowerCaseTabLabel === "financials tab") {
                        console.log(`AIModelPlanner: Skipping excluded tab - "${tabLabel}"`);
                        continue; 
                    }
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
                        // Update UI for this sub-task (optional)
                        // assistantMessageContent.textContent = `Processing details for tab: ${tabLabel}...`;
                        displayInClientChatLogPlanner(`Processing details for tab: ${tabLabel}...`, false); // Adds new bubble
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
                            displayInClientChatLogPlanner(`Completed details for tab: ${tabLabel}.`, false); // Adds new bubble
                        } catch (tabError) {
                            console.error(`AIModelPlanner: Error processing description for tab "${tabLabel}" via getAICallsProcessedResponse:`, tabError);
                            ModelCodes += `// Error processing tab ${tabLabel}: ${tabError.message}\n\n`;
                            displayInClientChatLogPlanner(`Error processing details for tab ${tabLabel}: ${tabError.message}`, false);
                        }
                    } else {
                        ModelCodes += `// No description provided for tab ${tabLabel}\n\n`;
                    }
                }
            }
            console.log("Generated ModelCodes (final):\n" + ModelCodes); 
            if (ModelCodes.trim().length > 0) {
                displayInClientChatLogPlanner("Model structure generated. Now applying to workbook...", false);
                await _executePlannerCodes(ModelCodes); 
            } else {
                console.log("AIModelPlanner: ModelCodes string is empty. Skipping _executePlannerCodes call.");
                displayInClientChatLogPlanner("No code content generated to apply to workbook.", false);
            }
        } else {
            // If it wasn't a JSON object for processing, the text is already displayed via streaming.
            // No further action needed here for UI unless fullAssistantTextResponse was empty.
            if (!fullAssistantTextResponse) {
                 assistantMessageContent.textContent = "Received an empty response.";
            }
        }

    } catch (error) {
        console.error("Error in AIModelPlanner conversation:", error);
        
        // Check if this is a validation error
        if (error.message && error.message.includes("Code validation failed")) {
            console.log("AIModelPlanner: Validation error detected. Attempting automatic resubmission...");
            
            // Display validation error to user
            displayInClientChatLogPlanner(`Validation error detected: ${error.message}\n\nAutomatically resubmitting for correction...`, false);
            
            // Check if we haven't exceeded retry attempts (prevent infinite loops)
            const maxRetries = 2;
            const currentRetryCount = error.retryCount || 0;
            
            if (currentRetryCount >= maxRetries) {
                console.log("AIModelPlanner: Maximum retry attempts reached. Stopping automatic correction.");
                displayInClientChatLogPlanner(`Maximum retry attempts (${maxRetries}) reached. Please manually adjust your request.`, false);
                assistantMessageContent.textContent = `Error: ${error.message}`;
                return;
            }
            
            // If we have a JSON object that failed validation, resubmit it
            if (lastPlannerResponseForClient && typeof lastPlannerResponseForClient === 'object') {
                try {
                    // Create a new prompt that includes the validation error
                    const retryPrompt = `The previous model structure had validation errors:\n${error.message}\n\nPlease regenerate the model structure fixing these validation issues. Keep the same tabs but ensure all code strings are valid.`;
                    
                    // Clear the assistant message for retry
                    assistantMessageContent.textContent = "Regenerating with corrections...";
                    
                    // Retry the conversation with the error context
                    const retryStream = _handleAIModelPlannerConversation(retryPrompt, { stream: true });
                    let retryFullResponse = "";
                    
                    for await (const chunk of retryStream) {
                        if (chunk.choices && chunk.choices[0]?.delta?.content) {
                            const content = chunk.choices[0].delta.content;
                            retryFullResponse += content;
                            assistantMessageContent.textContent = content; // Replace, don't append for cleaner retry
                            if (chatLogClient) chatLogClient.scrollTop = chatLogClient.scrollHeight;
                        }
                    }
                    
                    // Try to parse and process the retry response
                    let retryJsonObject = null;
                    try {
                        const parsedRetry = JSON.parse(retryFullResponse);
                        if (typeof parsedRetry === 'object' && parsedRetry !== null && !Array.isArray(parsedRetry)) {
                            retryJsonObject = parsedRetry;
                            lastPlannerResponseForClient = parsedRetry;
                        }
                    } catch (e) {
                        console.log("AIModelPlanner: Retry response was not parsable JSON.");
                    }
                    
                    if (retryJsonObject) {
                        // Process the corrected JSON
                        let ModelCodes = "";
                        for (const tabLabel in retryJsonObject) {
                            if (Object.prototype.hasOwnProperty.call(retryJsonObject, tabLabel)) {
                                const lowerCaseTabLabel = tabLabel.toLowerCase();
                                if (lowerCaseTabLabel === "financials" || lowerCaseTabLabel === "financials tab") {
                                    continue;
                                }
                                ModelCodes += `<TAB; label1="${tabLabel}";>\n`;
                                const tabDescription = retryJsonObject[tabLabel];
                                let tabDescriptionString = "";
                                if (typeof tabDescription === 'string') {
                                    tabDescriptionString = tabDescription;
                                } else if (typeof tabDescription === 'object' && tabDescription !== null) {
                                    tabDescriptionString = JSON.stringify(tabDescription);
                                } else {
                                    tabDescriptionString = String(tabDescription);
                                }
                                if (tabDescriptionString.trim() !== "") {
                                    displayInClientChatLogPlanner(`Processing corrected details for tab: ${tabLabel}...`, false);
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
                                        displayInClientChatLogPlanner(`Completed corrected details for tab: ${tabLabel}.`, false);
                                    } catch (tabError) {
                                        console.error(`AIModelPlanner: Error processing corrected tab "${tabLabel}":`, tabError);
                                        ModelCodes += `// Error processing tab ${tabLabel}: ${tabError.message}\n\n`;
                                        displayInClientChatLogPlanner(`Error processing corrected tab ${tabLabel}: ${tabError.message}`, false);
                                    }
                                } else {
                                    ModelCodes += `// No description provided for tab ${tabLabel}\n\n`;
                                }
                            }
                        }
                        
                        if (ModelCodes.trim().length > 0) {
                            displayInClientChatLogPlanner("Corrected model structure generated. Now applying to workbook...", false);
                            // Pass retry count in case there's another validation error
                            try {
                                await _executePlannerCodes(ModelCodes);
                            } catch (retryValidationError) {
                                // If it's another validation error, increment retry count
                                if (retryValidationError.message && retryValidationError.message.includes("Code validation failed")) {
                                    retryValidationError.retryCount = (currentRetryCount || 0) + 1;
                                }
                                throw retryValidationError;
                            }
                        } else {
                            displayInClientChatLogPlanner("No corrected code content generated.", false);
                        }
                    } else {
                        displayInClientChatLogPlanner("Could not generate corrected model structure. Please try rephrasing your request.", false);
                    }
                    
                } catch (retryError) {
                    console.error("Error during automatic retry:", retryError);
                    displayInClientChatLogPlanner(`Failed to automatically correct validation errors: ${retryError.message}`, false);
                }
            } else {
                // No JSON object to retry with, just show the error
                assistantMessageContent.textContent = `Error: ${error.message}`;
            }
        } else {
            // Not a validation error, show as normal
            assistantMessageContent.textContent = `Error: ${error.message}`;
        }

    } finally {
        setClientLoadingStatePlanner(false);
    }
}

export function plannerHandleReset() {
    const chatLog = document.getElementById('chat-log-client');
    if (chatLog) {
        chatLog.innerHTML = '';
        // >>> ADDED: Hide chat log when resetting
        chatLog.style.display = 'none';
        // <<< END ADDED
    }

    // >>> ADDED: Restore header and initial layout
    const clientHeader = document.getElementById('client-mode-header');
    const clientChatContainer = document.getElementById('client-chat-container');
    
    if (clientHeader) {
        clientHeader.style.display = ''; // Show it first
        clientHeader.classList.remove('hidden'); // Remove hidden class to trigger transition
    }
    
    if (clientChatContainer) {
        clientChatContainer.classList.remove('conversation-active');
    }
    // <<< END ADDED

    const welcomeMessage = document.createElement('div');
    welcomeMessage.id = 'welcome-message-client';
    welcomeMessage.className = 'welcome-message';
    welcomeMessage.style.display = 'none'; // Keep hidden since we're showing the header
    const welcomeTitle = document.createElement('h1');
    welcomeTitle.textContent = 'Ask me anything (Client Mode)';
    welcomeMessage.appendChild(welcomeTitle);
    if (chatLog) chatLog.appendChild(welcomeMessage);

    modelPlannerConversationHistory = [];
    lastPlannerResponseForClient = null;

    const userInput = document.getElementById('user-input-client');
    if (userInput) userInput.value = '';

    console.log("AIModelPlanner: Client chat reset completed.");
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


