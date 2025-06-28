// AIModelPlanner.js

// Assuming API keys are managed and accessible similarly to AIcalls.js
// Or that callOpenAI is available globally/imported if AIcalls.js exports it.
// For now, let's assume a local way to call OpenAI or that it's handled by the main taskpane script.

// >>> ADDED: Import CONFIG for URL management
import { CONFIG } from './config.js';
// Imports needed for _executePlannerCodes
import { validateCodeStringsForRun } from './Validation.js';
import { 
    populateCodeCollection, 
    runCodes, 
    processAssumptionTabs, 
    hideColumnsAndNavigate, 
    handleInsertWorksheetsFromBase64 
} from './CodeCollection.js';
import { getAICallsProcessedResponse, validationCorrection } from './AIcalls.js';
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
    CONFIG.getPromptUrl(`${promptKey}.txt`), 
    CONFIG.getAssetUrl(`src/prompts/${promptKey}.txt`) // Original path as a fallback
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

// NEW HELPER FUNCTION FOR PER-TAB VALIDATION
async function validateAndCorrectTabCode(tabLabel, tabCode, maxRetries = 2) {
    console.log(`[validateAndCorrectTabCode] Validating code for tab: ${tabLabel}`);
    
    for (let retryCount = 0; retryCount <= maxRetries; retryCount++) {
        // Add the <BR> substitution for validation
        let processedTabCode = tabCode;
        if (processedTabCode && typeof processedTabCode === 'string') {
            processedTabCode = processedTabCode.replace(/<BR>/g, '<BR; labelRow=""; row1 = "||||||||||||";>');
        }
        
        // Validate the tab code
        const codeLines = processedTabCode.split(/\r?\n/).filter(line => line.trim() !== '');
        const validationErrors = await validateCodeStringsForRun(codeLines);
        
        if (!validationErrors || validationErrors.length === 0) {
            console.log(`[validateAndCorrectTabCode] Tab "${tabLabel}" validation successful${retryCount > 0 ? ` after ${retryCount} attempt(s)` : ''}`);
            return { success: true, code: tabCode, retryCount };
        }
        
        // If we have validation errors and haven't exceeded max retries
        if (retryCount < maxRetries) {
            console.log(`[validateAndCorrectTabCode] Tab "${tabLabel}" validation failed (attempt ${retryCount + 1}/${maxRetries + 1}). Attempting correction...`);
            displayInClientChatLogPlanner(`Validation errors detected for tab ${tabLabel} (attempt ${retryCount + 1}/${maxRetries + 1}). Attempting automatic correction...`, false);
            
            try {
                // Use validationCorrection for consistency with main validation flow
                const tabContext = `Tab: ${tabLabel}\nCode:\n<TAB; label1="${tabLabel}";>\n${tabCode}`;
                const correctedResponse = await validationCorrection(
                    tabContext, // clientprompt - context about what this tab code is for
                    [tabCode], // initialResponse - the code as an array (validationCorrection expects array format)
                    validationErrors // validationResults - the errors found
                );
                
                let correctedCode = "";
                if (Array.isArray(correctedResponse)) {
                    correctedCode = correctedResponse.join('\n');
                } else if (typeof correctedResponse === 'object' && correctedResponse !== null && !Array.isArray(correctedResponse)) {
                    correctedCode = JSON.stringify(correctedResponse, null, 2);
                } else {
                    correctedCode = String(correctedResponse);
                }
                
                // Remove the tab header if the AI included it in the response
                correctedCode = correctedCode.replace(/^<TAB;[^>]+>\s*/i, '');
                
                console.log(`[validateAndCorrectTabCode] Received corrected code for tab "${tabLabel}", will retry validation`);
                tabCode = correctedCode; // Update for next iteration
                
            } catch (correctionError) {
                console.error(`[validateAndCorrectTabCode] Error during tab validation correction for "${tabLabel}":`, correctionError);
                displayInClientChatLogPlanner(`Failed to automatically correct validation errors for tab ${tabLabel}: ${correctionError.message}`, false);
                // Continue to next retry or fail
            }
        } else {
            // Max retries reached
            console.error(`[validateAndCorrectTabCode] Tab "${tabLabel}" validation failed after ${maxRetries + 1} attempts`);
            displayInClientChatLogPlanner(`Tab ${tabLabel}: Maximum validation correction attempts (${maxRetries + 1}) reached. Errors: ${validationErrors.join(", ")}`, false);
            return { 
                success: false, 
                code: tabCode, 
                errors: validationErrors,
                retryCount: maxRetries
            };
        }
    }
    
    // Should never reach here, but just in case
    return { success: false, code: tabCode, errors: ['Unknown error'], retryCount: maxRetries };
}

// Direct OpenAI API call function (simplified version, adapt if AIcalls.js exports its own)
async function* callOpenAIForModelPlanner(messages, options = {}) {
  const { model = "gpt-4.1", temperature = 0.7, stream = false } = options;

  if (!AI_MODEL_PLANNER_OPENAI_API_KEY) {
    console.error("AIModelPlanner: OpenAI API key not set.");
    throw new Error("OpenAI API key not set for AIModelPlanner.");
  }

  try {
    console.log(`AIModelPlanner - Calling OpenAI API with model: ${model}, stream: ${stream}`);
    
    // >>> ADDED: Comprehensive logging of all messages sent to OpenAI
    console.log("\n╔════════════════════════════════════════════════════════════════╗");
    console.log("║               AI MODEL PLANNER - OPENAI API CALL               ║");
    console.log("╠════════════════════════════════════════════════════════════════╣");
    console.log("║ FILE: AIModelPlanner.js                                        ║");
    console.log("║ FUNCTION: callOpenAIForModelPlanner()                          ║");
    console.log("║ PURPOSE: AI Model Planning / Client Mode Chat                  ║");
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
    console.log("║           END OF AI MODEL PLANNER OPENAI API CALL             ║");
    console.log("╚════════════════════════════════════════════════════════════════╝\n");
    // <<< END ADDED

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
      let buffer = ""; // Buffer to accumulate incomplete chunks

      while (true) {
        const { done, value } = await reader.read();
        if (done) {
          console.log("AIModelPlanner - Stream finished.");
          // Process any remaining buffer content
          if (buffer.trim()) {
            const finalLines = buffer.split("\n");
            for (const line of finalLines) {
              const cleanedLine = line.replace(/^data: /, "").trim();
              if (cleanedLine && cleanedLine !== "[DONE]") {
                try {
                  const parsedLine = JSON.parse(cleanedLine);
                  yield parsedLine;
                } catch (e) {
                  console.warn("AIModelPlanner - Could not parse final JSON line from stream:", cleanedLine, e);
                }
              }
            }
          }
          break;
        }
        
        const chunk = decoder.decode(value, { stream: true }); // Use stream: true for proper decoding
        buffer += chunk;
        
        // Split by double newlines to get complete SSE events
        const events = buffer.split("\n\n");
        
        // Keep the last potentially incomplete event in the buffer
        buffer = events.pop() || "";
        
        // Process complete events
        for (const event of events) {
          if (!event.trim()) continue;
          
          const lines = event.split("\n");
          for (const line of lines) {
            const cleanedLine = line.replace(/^data: /, "").trim();
            if (cleanedLine && cleanedLine !== "[DONE]") {
              try {
                const parsedLine = JSON.parse(cleanedLine);
                yield parsedLine;
              } catch (e) {
                console.warn("AIModelPlanner - Could not parse JSON line from stream:", cleanedLine, e);
              }
            }
          }
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
    // >>> ADDED: Log the function call details
    console.log("\n╔══════════════════════════════════════════════════════════════╗");
    console.log("║        processAIModelPlannerPromptInternal CALLED              ║");
    console.log("╠════════════════════════════════════════════════════════════════╣");
    console.log("║ FILE: AIModelPlanner.js                                        ║");
    console.log("║ PROMPT FILE: AIModelPlanning_System.txt                        ║");
    console.log("╚══════════════════════════════════════════════════════════════╝");
    console.log(`Model: ${model}`);
    console.log(`Temperature: ${temperature}`);
    console.log(`Stream: ${stream}`);
    console.log(`History items: ${history.length}`);
    console.log("────────────────────────────────────────────────────────────────\n");
    // <<< END ADDED

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
    console.log("[displayInClientChatLogPlanner] Called with message:", message);
    
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

async function _executePlannerCodes(modelCodesString, retryCount = 0) {
    console.log(`[AIModelPlanner._executePlannerCodes] Called. Retry count: ${retryCount}`);
    const MAX_RETRIES = 3;

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
        const validationErrors = await validateCodeStringsForRun(modelCodesString);
        if (validationErrors && validationErrors.length > 0) {
            const errorMsg = "Code validation failed for planner-generated codes:\n" + validationErrors.join("\n");
            console.error("[AIModelPlanner._executePlannerCodes] Code validation failed:", validationErrors);
            
            // Check if we should retry
            if (retryCount < MAX_RETRIES) {
                console.log(`[AIModelPlanner._executePlannerCodes] Attempting validation correction (attempt ${retryCount + 1}/${MAX_RETRIES})...`);
                displayInClientChatLogPlanner(`Validation errors detected (attempt ${retryCount + 1}/${MAX_RETRIES}). Attempting automatic correction...`, false);
                
                try {
                    // Call the AI to fix validation errors using validationCorrection
                    const modelContext = "Complete model code strings for all tabs";
                    const correctedCodes = await validationCorrection(
                        modelContext, // clientprompt - context about what this code is for
                        modelCodesString.split(/\r?\n/).filter(line => line.trim() !== ''), // initialResponse as array
                        validationErrors // validationResults - the errors found
                    );
                    
                    let correctedModelCodes = "";
                    if (Array.isArray(correctedCodes)) {
                        correctedModelCodes = correctedCodes.join('\n');
                    } else if (typeof correctedCodes === 'object' && correctedCodes !== null && !Array.isArray(correctedCodes)) {
                        correctedModelCodes = JSON.stringify(correctedCodes, null, 2);
                    } else {
                        correctedModelCodes = String(correctedCodes);
                    }
                    
                    console.log(`[AIModelPlanner._executePlannerCodes] Received corrected codes, retrying...`);
                    
                    // Recursive call with incremented retry count
                    return await _executePlannerCodes(correctedModelCodes, retryCount + 1);
                    
                } catch (correctionError) {
                    console.error(`[AIModelPlanner._executePlannerCodes] Error during validation correction:`, correctionError);
                    displayInClientChatLogPlanner(`Failed to automatically correct validation errors: ${correctionError.message}`, false);
                    // Fall through to throw the original error
                }
            } else {
                displayInClientChatLogPlanner(`Maximum validation correction attempts (${MAX_RETRIES}) reached. Please review and adjust your model structure.`, false);
            }
            
            // Throw the error to be caught by caller if max retries reached or correction failed
            const error = new Error(errorMsg);
            error.isValidationError = true;
            error.retryCount = retryCount;
            throw error;
        }
        console.log("[AIModelPlanner._executePlannerCodes] Code validation successful.");

        // LOG THE ENTIRE VALIDATED CODESTRINGS BEFORE BUILDING THE MODEL
        console.log("[AIModelPlanner._executePlannerCodes] === COMPLETE VALIDATED CODESTRINGS ===");
        console.log(modelCodesString);
        console.log("[AIModelPlanner._executePlannerCodes] === END OF CODESTRINGS ===");

        console.log("[AIModelPlanner._executePlannerCodes] Inserting base sheets from Worksheets_4.3.25 v1.xlsx...");
        const worksheetsResponse = await fetch(CONFIG.getAssetUrl('assets/Worksheets_4.3.25 v1.xlsx'));
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
        const codesResponse = await fetch(CONFIG.getAssetUrl('assets/codes.xlsx'));
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
        
        // Re-throw validation errors so they can be handled by the caller
        if (error.isValidationError) {
            throw error;
        }
        // For non-validation errors, don't re-throw (maintain existing behavior)
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
            console.log("AIModelPlanner: Starting to process JSON object for ModelCodes generation.");
            
            // Prepare all tabs for parallel processing
            const tabsToProcess = [];
            for (const tabLabel in jsonObjectToProcess) {
                if (Object.prototype.hasOwnProperty.call(jsonObjectToProcess, tabLabel)) {
                    const lowerCaseTabLabel = tabLabel.toLowerCase();
                    if (lowerCaseTabLabel === "financials" || lowerCaseTabLabel === "financials tab") {
                        console.log(`AIModelPlanner: Skipping excluded tab - "${tabLabel}"`);
                        continue; 
                    }
                    
                    const tabDescription = jsonObjectToProcess[tabLabel];
                    let tabDescriptionString = "";
                    if (typeof tabDescription === 'string') {
                        tabDescriptionString = tabDescription;
                    } else if (typeof tabDescription === 'object' && tabDescription !== null) {
                        tabDescriptionString = JSON.stringify(tabDescription);
                    } else {
                        tabDescriptionString = String(tabDescription);
                    }
                    
                    tabsToProcess.push({
                        label: tabLabel,
                        description: tabDescriptionString,
                        order: tabsToProcess.length // Preserve original order
                    });
                }
            }
            
            if (tabsToProcess.length === 0) {
                console.log("AIModelPlanner: No tabs to process after filtering.");
                displayInClientChatLogPlanner("No valid tabs found to process.", false);
                return;
            }
            
            // Display initial processing message
            displayInClientChatLogPlanner(`Processing ${tabsToProcess.length} tabs in parallel...`, false);
            
            // Process all tabs in parallel
            const tabProcessingPromises = tabsToProcess.map(async (tab) => {
                const { label: tabLabel, description: tabDescriptionString, order } = tab;
                
                console.log(`AIModelPlanner: Starting parallel processing for tab "${tabLabel}"`);
                displayInClientChatLogPlanner(`Processing details for tab: ${tabLabel}...`, false);
                
                try {
                    if (tabDescriptionString.trim() === "") {
                        return {
                            order,
                            tabLabel,
                            content: `// No description provided for tab ${tabLabel}\n\n`,
                            success: true
                        };
                    }
                    
                    // Call AI to generate tab content
                    const aiResponseForTabArray = await getAICallsProcessedResponse(tabDescriptionString);
                    let formattedAiResponse = "";
                    if (typeof aiResponseForTabArray === 'object' && aiResponseForTabArray !== null && !Array.isArray(aiResponseForTabArray)) {
                        formattedAiResponse = JSON.stringify(aiResponseForTabArray, null, 2); 
                    } else if (Array.isArray(aiResponseForTabArray)) {
                        formattedAiResponse = aiResponseForTabArray.join('\n');
                    } else {
                        formattedAiResponse = String(aiResponseForTabArray);
                    }
                    
                    // Validate and correct tab code
                    console.log(`AIModelPlanner: Validating generated code for tab "${tabLabel}"...`);
                    const validationResult = await validateAndCorrectTabCode(tabLabel, formattedAiResponse, 2);
                    
                    if (validationResult.success) {
                        console.log(`AIModelPlanner: Tab "${tabLabel}" code validated successfully`);
                        displayInClientChatLogPlanner(`✓ Completed and validated details for tab: ${tabLabel}`, false);
                        return {
                            order,
                            tabLabel,
                            content: validationResult.code + "\n\n",
                            success: true
                        };
                    } else {
                        // Tab validation failed after all retries
                        console.error(`AIModelPlanner: Tab "${tabLabel}" validation failed`);
                        displayInClientChatLogPlanner(`⚠ Warning: Tab ${tabLabel} has validation errors. Code included as comments.`, false);
                        const errorContent = `// WARNING: Tab ${tabLabel} has validation errors after ${validationResult.retryCount + 1} correction attempts\n` +
                                           `// Errors: ${validationResult.errors.join("; ")}\n` +
                                           `// Original code included below for reference:\n` +
                                           validationResult.code.split('\n').map(line => `// ${line}`).join('\n') + "\n\n";
                        return {
                            order,
                            tabLabel,
                            content: errorContent,
                            success: false
                        };
                    }
                } catch (tabError) {
                    console.error(`AIModelPlanner: Error processing tab "${tabLabel}":`, tabError);
                    displayInClientChatLogPlanner(`✗ Error processing details for tab ${tabLabel}: ${tabError.message}`, false);
                    return {
                        order,
                        tabLabel,
                        content: `// Error processing tab ${tabLabel}: ${tabError.message}\n\n`,
                        success: false
                    };
                }
            });
            
            // Wait for all tabs to complete processing
            console.log(`AIModelPlanner: Waiting for ${tabProcessingPromises.length} tabs to complete processing...`);
            const tabResults = await Promise.all(tabProcessingPromises);
            
            // Sort results by original order and assemble ModelCodes
            tabResults.sort((a, b) => a.order - b.order);
            
            let ModelCodes = "";
            let successCount = 0;
            let errorCount = 0;
            
            for (const result of tabResults) {
                ModelCodes += `<TAB; label1="${result.tabLabel}";>\n`;
                ModelCodes += result.content;
                
                if (result.success) {
                    successCount++;
                } else {
                    errorCount++;
                }
            }
            
            // Display processing summary
            const summaryMessage = `Parallel processing completed: ${successCount} tabs processed successfully${errorCount > 0 ? `, ${errorCount} tabs had errors` : ''}`;
            console.log(`AIModelPlanner: ${summaryMessage}`);
            displayInClientChatLogPlanner(summaryMessage, false);
            console.log("Generated ModelCodes (final):\n" + ModelCodes); 
            if (ModelCodes.trim().length > 0) {
                displayInClientChatLogPlanner("Model structure generated. Now applying to workbook...", false);
                await _executePlannerCodes(ModelCodes);
                // Save the ModelCodes to Last Model.txt file after successful execution
                await saveModelCodesToLastModelFile(ModelCodes);
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
        
        // The retry logic is now handled inside _executePlannerCodes
        // We just need to display the final error message
        assistantMessageContent.textContent = `Error: ${error.message}`;
        
        // If it was a validation error that exceeded max retries, provide additional guidance
        if (error.isValidationError && error.retryCount >= 3) {
            displayInClientChatLogPlanner("The model structure could not be automatically corrected after 3 attempts. Please review the validation errors and adjust your request.", false);
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

// Function to save ModelCodes to Last Model.txt file
async function saveModelCodesToLastModelFile(modelCodesString) {
    try {
        console.log("[saveModelCodesToLastModelFile] Saving ModelCodes to Last Model.txt...");
        
        // Create the content to save
        const timestamp = new Date().toISOString();
        const contentToSave = `// Model codes generated on ${timestamp}\n// Generated by AI Model Planner Client Mode\n\n${modelCodesString}`;
        
        // Store in localStorage as a backup
        try {
            localStorage.setItem('lastModelCodes', contentToSave);
            console.log("[saveModelCodesToLastModelFile] ModelCodes stored in localStorage as backup");
        } catch (storageError) {
            console.warn("[saveModelCodesToLastModelFile] Failed to store in localStorage:", storageError);
        }
        
        // For now, we'll log the content and provide instructions to the user
        // In a production environment, this would need a proper file save endpoint
        console.log("[saveModelCodesToLastModelFile] ModelCodes content for Last Model.txt:\n", contentToSave);
        
        // Create a downloadable file for the user
        try {
            const blob = new Blob([contentToSave], { type: 'text/plain' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'Last Model.txt';
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
            
            console.log("[saveModelCodesToLastModelFile] ModelCodes file download initiated");
            displayInClientChatLogPlanner("Model codes saved! A download should have started for 'Last Model.txt'. Please save this file to your src/prompts/ directory.", false);
        } catch (downloadError) {
            console.warn("[saveModelCodesToLastModelFile] Download creation failed:", downloadError);
            displayInClientChatLogPlanner("Model codes generated. Please check the console and copy the content to manually update src/prompts/Last Model.txt", false);
        }
        
    } catch (error) {
        console.error("[saveModelCodesToLastModelFile] Error saving ModelCodes:", error);
        displayInClientChatLogPlanner(`Error saving model codes: ${error.message}`, false);
    }
}

// Function to retrieve the last saved ModelCodes
export function getLastModelCodes() {
    try {
        const lastModelCodes = localStorage.getItem('lastModelCodes');
        if (lastModelCodes) {
            console.log("[getLastModelCodes] Retrieved ModelCodes from localStorage");
            return lastModelCodes;
        } else {
            console.log("[getLastModelCodes] No ModelCodes found in localStorage");
            return null;
        }
    } catch (error) {
        console.error("[getLastModelCodes] Error retrieving ModelCodes:", error);
        return null;
    }
}

// Function to clear the last saved ModelCodes
export function clearLastModelCodes() {
    try {
        localStorage.removeItem('lastModelCodes');
        console.log("[clearLastModelCodes] ModelCodes cleared from localStorage");
        return true;
    } catch (error) {
        console.error("[clearLastModelCodes] Error clearing ModelCodes:", error);
        return false;
    }
}


