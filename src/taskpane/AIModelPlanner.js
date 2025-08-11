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
import { ModelUpdateHandler, stripCodeBlockMarkers } from './ModelUpdateHandler.js';
// We don't import from taskpane.js to avoid cycles

let modelPlannerConversationHistory = [];
let AI_MODEL_PLANNER_OPENAI_API_KEY = "";
let lastPlannerResponseForClient = null; // To store the last response for client mode buttons
let currentAttachedFiles = []; // To store multiple attached files data

import { processModelCodesForPlanner } from './taskpane.js'; // <<< UPDATED IMPORT

const DEBUG_PLANNER = CONFIG.isDevelopment; // For planner-specific debugging

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


// Function to get the currently selected system prompt mode
function getSelectedSystemPromptMode() {
  // Check both possible dropdown IDs
  let dropdown = document.getElementById('mode-dropdown');
  if (!dropdown) {
    dropdown = document.getElementById('system-prompt-dropdown');
  }
  return dropdown ? dropdown.value : 'one-shot'; // Default to one-shot
}

// Function to load the Encoder Summary system prompt
async function getEncoderSummarySystemPrompt() {
  const promptKey = "Encoder_Summary_System";
  const paths = [
    CONFIG.getPromptUrl(`${promptKey}.txt`), 
    CONFIG.getAssetUrl(`src/prompts/${promptKey}.txt`)
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

  console.error(`AIModelPlanner: Failed to load prompt ${promptKey}.txt from all attempted paths.`);
  return "You are an expert financial analyst who translates technical financial model codes into clear, plain English explanations for business clients. [Error: Encoder Summary system prompt could not be loaded]"; 
}

// Function to generate model summary after successful model execution with streaming
async function generateModelSummary(modelCodesString) {
  if (!modelCodesString || modelCodesString.trim().length === 0) {
    console.log("[generateModelSummary] No model codes to summarize.");
    return;
  }

  if (!AI_MODEL_PLANNER_OPENAI_API_KEY) {
    console.error("[generateModelSummary] OpenAI API key not set.");
    displayInClientChatLogPlanner("Cannot generate model explanation - API key not configured.", false);
    return;
  }

  try {
    console.log("[generateModelSummary] Loading Encoder Summary system prompt...");
    const systemPrompt = await getEncoderSummarySystemPrompt();
    
    console.log("[generateModelSummary] Calling OpenAI to generate model summary with streaming...");
    const userPrompt = `Please analyze these model codes and provide a clear, business-friendly explanation of how the model works:\n\n${modelCodesString}`;
    
    const messages = [
      { role: "system", content: systemPrompt },
      { role: "user", content: userPrompt }
    ];

    // >>> ADDED: Create DOM elements for streaming response
    let chatLogClient = document.getElementById('chat-log-client');
    
    // If chat log doesn't exist, try to create it (consistent with displayInClientChatLogPlanner)
    if (!chatLogClient) {
        console.error("[generateModelSummary] Client chat log element not found. Attempting to create it...");
        const container = document.getElementById('client-chat-container');
        if (container) {
            chatLogClient = document.createElement('div');
            chatLogClient.id = 'chat-log-client';
            chatLogClient.className = 'chat-log';
            chatLogClient.style.display = 'block';
            chatLogClient.style.flexGrow = '1';
            chatLogClient.style.overflowY = 'auto';
            container.appendChild(chatLogClient);
            console.log("[generateModelSummary] Created chat log element");
        } else {
            console.error("[generateModelSummary] Could not find container to create chat log");
            return;
        }
    }

    // Hide welcome message if it exists
    const welcomeMessage = document.getElementById('welcome-message-client');
    if (welcomeMessage) welcomeMessage.style.display = 'none';

    const assistantMessageDiv = document.createElement('div');
    assistantMessageDiv.className = 'chat-message assistant-message';
    const assistantMessageContent = document.createElement('p');
    assistantMessageContent.className = 'message-content';
    assistantMessageContent.textContent = '## 📊 Your Model Explanation\n\n'; // Start with header
    assistantMessageDiv.appendChild(assistantMessageContent);
    
    chatLogClient.appendChild(assistantMessageDiv);
    chatLogClient.scrollTop = chatLogClient.scrollHeight;
    // <<< END ADDED

    const response = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${AI_MODEL_PLANNER_OPENAI_API_KEY}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        model: "gpt-4o",
        messages: messages,
        temperature: 0.3,
        max_tokens: 4000,
        stream: true // >>> ADDED: Enable streaming
      })
    });

    if (!response.ok) {
      throw new Error(`OpenAI API call failed: ${response.status} ${response.statusText}`);
    }

    // >>> ADDED: Handle streaming response
    const reader = response.body.getReader();
    const decoder = new TextDecoder();
    let buffer = "";
    let fullSummaryContent = "";

    while (true) {
      const { done, value } = await reader.read();
      
      if (done) {
        // Process any remaining buffer content
        if (buffer.trim()) {
          const finalLines = buffer.split("\n");
          for (const line of finalLines) {
            const cleanedLine = line.replace(/^data: /, "").trim();
            if (cleanedLine && cleanedLine !== "[DONE]") {
              try {
                const parsedLine = JSON.parse(cleanedLine);
                if (parsedLine.choices && parsedLine.choices[0]?.delta?.content) {
                  const content = parsedLine.choices[0].delta.content;
                  fullSummaryContent += content;
                  assistantMessageContent.textContent += content;
                  if (chatLogClient) chatLogClient.scrollTop = chatLogClient.scrollHeight;
                }
              } catch (e) {
                console.warn("[generateModelSummary] Could not parse final JSON line from stream:", cleanedLine, e);
              }
            }
          }
        }
        break;
      }
      
      const chunk = decoder.decode(value, { stream: true });
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
              if (parsedLine.choices && parsedLine.choices[0]?.delta?.content) {
                const content = parsedLine.choices[0].delta.content;
                fullSummaryContent += content;
                assistantMessageContent.textContent += content;
                if (chatLogClient) chatLogClient.scrollTop = chatLogClient.scrollHeight;
              }
            } catch (e) {
              console.warn("[generateModelSummary] Could not parse JSON line from stream:", cleanedLine, e);
            }
          }
        }
      }
    }
    // <<< END ADDED

    if (fullSummaryContent) {
      console.log("[generateModelSummary] Model summary generated successfully with streaming");
      // Update conversation history with the complete summary
      modelPlannerConversationHistory.push(['assistant', "## 📊 Your Model Explanation\n\n" + fullSummaryContent]);
    } else {
      throw new Error("No summary content received from OpenAI streaming response");
    }

  } catch (error) {
    console.error("[generateModelSummary] Error generating model summary:", error);
    displayInClientChatLogPlanner("Model built successfully, but encountered an error generating the explanation. You can still use your model in Excel.", false);
  }
}

// Function to detect if a model has been built (JSON output in conversation history)
function hasModelBeenBuilt() {
  // Check if there's any assistant message in the history that contains JSON with model tabs
  for (const [role, content] of modelPlannerConversationHistory) {
    if (role === 'assistant' && typeof content === 'string') {
      try {
        // >>> ADDED: Strip code block markers if present (for stored conversation history)
        const cleanedContent = stripCodeBlockMarkers(content);
        // <<< END ADDED
        
        // Try to parse as JSON
        const parsed = JSON.parse(cleanedContent);
        // Check if it looks like a model structure (has tab names as keys)
        if (typeof parsed === 'object' && parsed !== null && !Array.isArray(parsed)) {
          // Check if any key contains typical model tab indicators
          const keys = Object.keys(parsed);
          const modelTabIndicators = ['Revenue', 'Direct Costs', 'Corporate Overhead', 'Working Capital', 'Debt', 'Equity', 'PP&E', 'Fixed Assets'];
          const hasModelTabs = keys.some(key => 
            modelTabIndicators.some(indicator => key.includes(indicator))
          );
          if (hasModelTabs) {
            console.log('AIModelPlanner: Model detected in conversation history');
            return true;
          }
        }
      } catch (e) {
        // Not JSON, continue checking
      }
    }
  }
  return false;
}

// Updated to use the more robust fetching approach with dynamic prompt selection
async function getAIModelPlanningSystemPrompt() {
  const selectedMode = getSelectedSystemPromptMode();
  
  // Map the dropdown values to prompt file names
  const promptMap = {
    'one-shot': 'OneShot_Planner_System',
    'limited-guidance': 'AIModelPlanning_System',
    'full-guidance': 'ModelPLannerGuided_System',
    'updater': 'Model Updater_System',
    'one-shot-updater': 'OneShot_Model_Updater_System'
  };
  
  // Auto-switch to updater mode if model has been built and we're in follow-up conversation
  let effectiveMode = selectedMode;
  if (modelPlannerConversationHistory.length > 0 && hasModelBeenBuilt()) {
    // If original mode was one-shot, use one-shot updater; otherwise use regular updater
    effectiveMode = (selectedMode === 'one-shot') ? 'one-shot-updater' : 'updater';
    console.log(`AIModelPlanner: Auto-switching to ${effectiveMode} mode - model detected in conversation history`);
  }
  
  const promptKey = promptMap[effectiveMode] || 'OneShot_Planner_System'; // Fallback to one-shot
  
  const paths = [
    // Try path relative to root if /src/ is not working, assuming 'prompts' is then at root level of served dir
    // THIS IS A GUESS - The original path `https://localhost:3002/src/prompts/...` should work if server is configured for it.
    CONFIG.getPromptUrl(`${promptKey}.txt`), 
    CONFIG.getAssetUrl(`src/prompts/${promptKey}.txt`) // Original path as a fallback
  ];

  if (DEBUG_PLANNER) console.log(`AIModelPlanner: Attempting to load prompt file: ${promptKey}.txt (Mode: ${effectiveMode} ${effectiveMode !== selectedMode ? `[auto-switched from ${selectedMode}]` : ''})`);

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
  return "You are a helpful assistant for financial model planning. [Error: System prompt could not be loaded]"; 
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
    console.log("║ PROMPT FILE: [Dynamically determined based on conversation state] ║");
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
                // >>> ADDED: Strip code block markers if present
                const cleanedContent = stripCodeBlockMarkers(responseContent);
                // <<< END ADDED
                
                const parsedJson = JSON.parse(cleanedContent);
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
            // Chat log now exists in HTML; reuse it instead of creating
            chatLog = document.getElementById('chat-log-client');
            if (!chatLog) {
                console.error('Chat log element missing in DOM');
                return;
            }
            chatLog.style.display = 'block';
            console.log("AIModelPlanner: Reusing chat log element");
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
    
    // >>> IMMEDIATE DEBUG: Check window.currentUpdateHandler status
    console.log("🔍 [FUNCTION START] _executePlannerCodes - window.currentUpdateHandler exists:", !!window.currentUpdateHandler);
    if (window.currentUpdateHandler) {
        console.log("🔍 [FUNCTION START] _executePlannerCodes - isNewBuild:", window.currentUpdateHandler.isNewBuild);
        console.log("🔍 [FUNCTION START] _executePlannerCodes - updateData keys:", window.currentUpdateHandler.updateData ? Object.keys(window.currentUpdateHandler.updateData) : 'null');
    }
    // <<< END IMMEDIATE DEBUG
    
    const MAX_RETRIES = 3;

    // >>> ADDED: Apply ModelUpdateHandler transformations if this is an update
    if (window.currentUpdateHandler && !window.currentUpdateHandler.isNewBuild) {
        console.log("[AIModelPlanner._executePlannerCodes] Applying update transformations...");
        const modelCodesArray = modelCodesString.split('\n');
        const transformedCodes = window.currentUpdateHandler.processModelCodes(modelCodesArray);
        modelCodesString = transformedCodes.join('\n');
        console.log("[AIModelPlanner._executePlannerCodes] Model codes transformed for update");
        
        // Handle calcs deletion if required
        if (window.currentUpdateHandler.shouldDeleteCalcs()) {
            console.log("[AIModelPlanner._executePlannerCodes] Calcs deletion required for this update type");
            // Add logic to delete calcs sheet here if needed
        }
    }
    // <<< END ADDED
    
    console.log("🔍 [FLOW TRACE] After update transformations, continuing to validation section...");

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

    console.log("🔍 [PRE-VALIDATION] About to enter validation try block...");
    console.log("🔍 [PRE-VALIDATION] modelCodesString length:", modelCodesString.length);
    console.log("🔍 [PRE-VALIDATION] modelCodesString preview:", modelCodesString.substring(0, 200));

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

        // >>> ADDED: Clean up Financials references for replaced tabs BEFORE processing codes
        console.log("🔍 [CLEANUP TRACE] Code validation complete - starting cleanup check...");
        console.log("🔍 [CLEANUP TRACE] window.currentUpdateHandler exists:", !!window.currentUpdateHandler);
        console.log("🔍 [CLEANUP TRACE] window object keys:", Object.keys(window).filter(k => k.includes('Update')));
        if (window.currentUpdateHandler) {
            console.log("[AIModelPlanner._executePlannerCodes] CLEANUP DEBUG - isNewBuild:", window.currentUpdateHandler.isNewBuild);
            console.log("[AIModelPlanner._executePlannerCodes] CLEANUP DEBUG - updateData:", window.currentUpdateHandler.updateData);
        }
        
        if (window.currentUpdateHandler && !window.currentUpdateHandler.isNewBuild) {
            console.log("🟢 [CLEANUP EXECUTING] Cleaning up Financials references for replaced tabs...");
            try {
                await window.currentUpdateHandler.cleanupReplacedTabs();
                console.log("✅ [CLEANUP SUCCESS] Financials cleanup completed");
            } catch (cleanupError) {
                console.error("❌ [CLEANUP ERROR] Error during Financials cleanup:", cleanupError);
                // Continue execution - cleanup failure shouldn't stop model building
            }
        } else {
            const reason = !window.currentUpdateHandler ? "No updateHandler" : `isNewBuild=${window.currentUpdateHandler.isNewBuild}`;
            console.log("⏭️ [CLEANUP SKIPPED] Reason:", reason);
        }
        // <<< END ADDED

        // LOG THE ENTIRE VALIDATED CODESTRINGS BEFORE BUILDING THE MODEL
        console.log("[AIModelPlanner._executePlannerCodes] === COMPLETE VALIDATED CODESTRINGS ===");
        console.log(modelCodesString);
        console.log("[AIModelPlanner._executePlannerCodes] === END OF CODESTRINGS ===");

        // >>> ADDED: Only insert base worksheets for new builds
        if (!window.currentUpdateHandler || window.currentUpdateHandler.isNewBuild) {
            console.log("[AIModelPlanner._executePlannerCodes] NEW BUILD: Inserting base sheets from Worksheets_4.3.25 v1.xlsx...");
            const worksheetsResponse = await fetch(CONFIG.getAssetUrl('assets/Worksheets_4.3.25 v1.xlsx'));
            if (!worksheetsResponse.ok) throw new Error(`[AIModelPlanner._executePlannerCodes] Worksheets_4.3.25 v1.xlsx load failed: ${worksheetsResponse.statusText}`);
            const wsArrayBuffer = await worksheetsResponse.arrayBuffer();
            console.log(`[AIModelPlanner._executePlannerCodes] Worksheets ArrayBuffer size: ${wsArrayBuffer.byteLength} bytes`);
            
            // Improved base64 conversion using modern method
            let wsBase64String;
            try {
                const blob = new Blob([wsArrayBuffer]);
                wsBase64String = await new Promise((resolve, reject) => {
                    const reader = new FileReader();
                    reader.onload = () => {
                        const dataUrl = reader.result;
                        const base64 = dataUrl.split(',')[1]; // Remove data:application/octet-stream;base64, prefix
                        resolve(base64);
                    };
                    reader.onerror = reject;
                    reader.readAsDataURL(blob);
                });
                console.log(`[AIModelPlanner._executePlannerCodes] Worksheets base64 length: ${wsBase64String.length} characters`);
            } catch (conversionError) {
                console.error(`[AIModelPlanner._executePlannerCodes] Worksheets base64 conversion failed:`, conversionError);
                throw new Error(`Failed to convert worksheets to base64: ${conversionError.message}`);
            }
            
            await handleInsertWorksheetsFromBase64(wsBase64String);
            console.log("[AIModelPlanner._executePlannerCodes] Base sheets (Worksheets_4.3.25 v1.xlsx) inserted.");
        } else {
            console.log("[AIModelPlanner._executePlannerCodes] UPDATE MODE: Skipping base worksheets insertion");
        }
        // <<< END ADDED

        // >>> ADDED: Determine which sheets to insert based on update type
        const sheetsToInsert = window.currentUpdateHandler ? window.currentUpdateHandler.getSheetsToMove() : ['Codes'];
        console.log(`[AIModelPlanner._executePlannerCodes] Inserting Codes.xlsx sheets: ${sheetsToInsert.join(', ')}...`);
        
        const codesResponse = await fetch(CONFIG.getAssetUrl('assets/Codes.xlsx'));
        if (!codesResponse.ok) throw new Error(`[AIModelPlanner._executePlannerCodes] Codes.xlsx load failed: ${codesResponse.statusText}`);
        const codesArrayBuffer = await codesResponse.arrayBuffer();
        console.log(`[AIModelPlanner._executePlannerCodes] Codes ArrayBuffer size: ${codesArrayBuffer.byteLength} bytes`);
        
        // Improved base64 conversion using modern method
        let codesBase64String;
        try {
            const blob = new Blob([codesArrayBuffer]);
            codesBase64String = await new Promise((resolve, reject) => {
                const reader = new FileReader();
                reader.onload = () => {
                    const dataUrl = reader.result;
                    const base64 = dataUrl.split(',')[1]; // Remove data:application/octet-stream;base64, prefix
                    resolve(base64);
                };
                reader.onerror = reject;
                reader.readAsDataURL(blob);
            });
            console.log(`[AIModelPlanner._executePlannerCodes] Codes base64 length: ${codesBase64String.length} characters`);
        } catch (conversionError) {
            console.error(`[AIModelPlanner._executePlannerCodes] Codes base64 conversion failed:`, conversionError);
            throw new Error(`Failed to convert codes to base64: ${conversionError.message}`);
        }
        
        await handleInsertWorksheetsFromBase64(codesBase64String, sheetsToInsert); 
        console.log(`[AIModelPlanner._executePlannerCodes] Codes.xlsx sheets (${sheetsToInsert.join(', ')}) inserted/updated.`);
        // <<< END ADDED
    
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

        // >>> ADDED: Set calculation mode to automatic BEFORE model summary generation
        displayInClientChatLogPlanner("🔄 Activating Excel calculations - your model is now live!", false);
        console.log("[AIModelPlanner._executePlannerCodes] Setting calculation mode to automatic...");
        try {
            await Excel.run(async (context) => {
                context.application.calculationMode = Excel.CalculationMode.automatic;
                await context.sync();
                console.log("[AIModelPlanner._executePlannerCodes] Calculation mode set to automatic - Excel will now recalculate.");
            });
        } catch (calcError) {
            console.error("[AIModelPlanner._executePlannerCodes] Error setting calculation mode to automatic:", calcError);
        }
        // <<< END ADDED

        // >>> ADDED: Generate model summary after recalculation
        console.log("[AIModelPlanner._executePlannerCodes] Generating model summary...");
        try {
            await generateModelSummary(modelCodesString);
        } catch (summaryError) {
            console.error("[AIModelPlanner._executePlannerCodes] Error generating model summary:", summaryError);
            displayInClientChatLogPlanner("Model built successfully, but could not generate summary explanation.", false);
        }
        // <<< END ADDED

    } catch (error) {
        console.error("[AIModelPlanner._executePlannerCodes] Error during processing:", error);
        displayInClientChatLogPlanner(`Error applying model structure: ${error.message}`, false);
        
        // Re-throw validation errors so they can be handled by the caller
        if (error.isValidationError) {
            throw error;
        }
        // For non-validation errors, don't re-throw (maintain existing behavior)
    } finally {
        // >>> MODIFIED: Ensure calculation mode is always set to automatic as final safety measure
        try {
            await Excel.run(async (context) => {
                context.application.calculationMode = Excel.CalculationMode.automatic;
                await context.sync();
                console.log("[AIModelPlanner._executePlannerCodes] Calculation mode ensured automatic in finally block.");
            });
        } catch (finalError) {
            console.error("[AIModelPlanner._executePlannerCodes] Error ensuring calculation mode in finally block:", finalError);
        }
        // <<< END MODIFIED
    }
}

export async function plannerHandleSend() {
    console.log("[plannerHandleSend] Function called - VERSION 3"); // Debug to ensure new code is running
    
    const userInputElement = document.getElementById('user-input-client');
    if (!userInputElement) { console.error("AIModelPlanner: Client user input element not found."); return; }
    let userInput = userInputElement.value.trim();

    // Check if there are attached files and include them in the prompt
    if (currentAttachedFiles.length > 0) {
        console.log("[plannerHandleSend] Including attached files data:", currentAttachedFiles.map(f => f.fileName).join(', '));
        const filesDataForAI = formatFileDataForAI(currentAttachedFiles);
        userInput = userInput + filesDataForAI;
    }

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
        // Minimal inline styles - let CSS handle the layout
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

    // Display only the original user input (without file data) in chat log
    const originalUserInput = userInputElement.value.trim();
    let displayMessage = originalUserInput;
    if (currentAttachedFiles.length > 0) {
        const fileNames = currentAttachedFiles.map(f => f.fileName).join(', ');
        displayMessage += ` 📎 (${currentAttachedFiles.length} file${currentAttachedFiles.length !== 1 ? 's' : ''}: ${fileNames})`;
    }
    displayInClientChatLogPlanner(displayMessage, true);
    
    userInputElement.value = '';
    
    // Clear attachments after sending
    removeAllAttachments();
    
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
                // >>> ADDED: Strip code block markers if present
                const cleanedResponse = stripCodeBlockMarkers(fullAssistantTextResponse);
                if (cleanedResponse !== fullAssistantTextResponse.trim()) {
                    console.log("AIModelPlanner: Stripped code block markers from response");
                }
                // <<< END ADDED
                
                const parsedResponse = JSON.parse(cleanedResponse);
                if (typeof parsedResponse === 'object' && parsedResponse !== null && !Array.isArray(parsedResponse)) {
                    jsonObjectToProcess = parsedResponse;
                    lastPlannerResponseForClient = parsedResponse; // Update to store parsed object
                    console.log("AIModelPlanner: Streamed text successfully parsed to an object for tab processing.");
                    
                    // >>> ADDED: Replace JSON content with user-friendly message
                    assistantMessageContent.textContent = "✅ Model plan confirmed! Building your financial model now...";
                    if (chatLogClient) chatLogClient.scrollTop = chatLogClient.scrollHeight;
                    console.log("AIModelPlanner: Replaced JSON content with user-friendly message");
                    // <<< END ADDED
                }
            } catch (e) {
                // Not a JSON string, or parsing did not result in a suitable object.
                // The UI already has the full text. lastPlannerResponseForClient remains the text.
                console.log("AIModelPlanner: Streamed text was not a parsable JSON object. Displaying as text.");
            }
        }

        if (jsonObjectToProcess) {
            console.log("AIModelPlanner: Starting to process JSON object for ModelCodes generation.");
            
            // >>> ADDED: Initialize ModelUpdateHandler to determine update type
            const updateHandler = new ModelUpdateHandler();
            const isNewBuild = updateHandler.determineUpdateType(jsonObjectToProcess);
            const updateSummary = updateHandler.getUpdateSummary();
            
            console.log("AIModelPlanner: Update analysis:", updateSummary);
            console.log("AIModelPlanner: JSON Object keys:", Object.keys(jsonObjectToProcess));
            console.log("AIModelPlanner: isNewBuild determined as:", isNewBuild);
            console.log("AIModelPlanner: updateHandler.isNewBuild:", updateHandler.isNewBuild);
            console.log("AIModelPlanner: updateHandler.updateData:", updateHandler.updateData);
            
            // Store update handler for later use in _executePlannerCodes
            window.currentUpdateHandler = updateHandler;
            console.log("🔍 [UPDATE HANDLER] Stored window.currentUpdateHandler with isNewBuild:", window.currentUpdateHandler.isNewBuild);
            console.log("🔍 [UPDATE HANDLER] updateData:", window.currentUpdateHandler.updateData);
            console.log("🔍 [UPDATE HANDLER] Verification - window.currentUpdateHandler exists:", !!window.currentUpdateHandler);
            // <<< END ADDED
            
            // Prepare all tabs for parallel processing
            const tabsToProcess = [];
            
            // >>> ADDED: Handle update vs new build JSON structures differently
            let tabsToIterate;
            if (!isNewBuild && jsonObjectToProcess.tab_updates) {
                // For updates, iterate over the tabs inside tab_updates
                tabsToIterate = jsonObjectToProcess.tab_updates;
                console.log("AIModelPlanner: Processing UPDATE structure - iterating over tab_updates content");
            } else {
                // For new builds, iterate over top-level keys
                tabsToIterate = jsonObjectToProcess;
                console.log("AIModelPlanner: Processing NEW BUILD structure - iterating over top-level keys");
            }
            // <<< END ADDED
            
            for (const tabLabel in tabsToIterate) {
                if (Object.prototype.hasOwnProperty.call(tabsToIterate, tabLabel)) {
                    const lowerCaseTabLabel = tabLabel.toLowerCase();
                    if (lowerCaseTabLabel === "financials" || lowerCaseTabLabel === "financials tab") {
                        console.log(`AIModelPlanner: Skipping excluded tab - "${tabLabel}"`);
                        continue; 
                    }
                    
                    const tabData = tabsToIterate[tabLabel];
                    let tabDescriptionString = "";
                    
                    // >>> ADDED: Handle update structure (object with update_type and description)
                    if (!isNewBuild && typeof tabData === 'object' && tabData !== null && tabData.description) {
                        tabDescriptionString = tabData.description;
                        console.log(`AIModelPlanner: Processing update for tab "${tabLabel}" with type "${tabData.update_type}"`);
                    } else if (typeof tabData === 'string') {
                        tabDescriptionString = tabData;
                    } else if (typeof tabData === 'object' && tabData !== null) {
                        tabDescriptionString = JSON.stringify(tabData);
                    } else {
                        tabDescriptionString = String(tabData);
                    }
                    // <<< END ADDED
                    
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
                
                // Also use enhanced storage system for better tab management
                try {
                    const { storeModelCodes } = await import('./ModelCodeStorage.js');
                    storeModelCodes(ModelCodes);
                    console.log("[plannerHandleSend] Model codes stored in enhanced storage system");
                } catch (importError) {
                    console.warn("[plannerHandleSend] Could not import ModelCodeStorage:", importError);
                }
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

    // >>> ADDED: Clear attached files
    removeAllAttachments();
    // <<< END ADDED

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
            // Also store just the raw codes without timestamp for easier processing
            localStorage.setItem('rawModelCodes', modelCodesString);
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

// ========== FILE ATTACHMENT FUNCTIONALITY ==========

// Function to format file size
function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

// Function to process XLSX file using SheetJS
function processXLSXFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                
                // Determine the actual file type based on extension
                const fileExtension = file.name.toLowerCase().split('.').pop();
                let fileType = 'XLSX'; // default
                if (fileExtension === 'xls') {
                    fileType = 'XLS';
                } else if (fileExtension === 'xlsm') {
                    fileType = 'XLSM';
                }
                
                const result = {
                    fileName: file.name,
                    fileSize: file.size,
                    fileType: fileType,
                    sheets: {},
                    sheetNames: workbook.SheetNames
                };
                
                // Process each sheet
                workbook.SheetNames.forEach(sheetName => {
                    const worksheet = workbook.Sheets[sheetName];
                    
                    // Convert to JSON format (array of arrays)
                    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                    
                    // Also get CSV format for text processing
                    const csvData = XLSX.utils.sheet_to_csv(worksheet);
                    
                    result.sheets[sheetName] = {
                        json: jsonData,
                        csv: csvData,
                        range: worksheet['!ref'] || 'A1:A1'
                    };
                });
                
                console.log(`[processXLSXFile] Successfully processed ${fileType} file:`, result.fileName);
                resolve(result);
            } catch (error) {
                console.error('[processXLSXFile] Error processing XLSX file:', error);
                reject(error);
            }
        };
        reader.onerror = () => reject(new Error('Failed to read file'));
        reader.readAsArrayBuffer(file);
    });
}

// Function to process CSV file
function processCSVFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const csvData = e.target.result;
                
                // Parse CSV data using SheetJS
                const workbook = XLSX.read(csvData, { type: 'string' });
                const worksheet = workbook.Sheets[workbook.SheetNames[0]];
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                
                const result = {
                    fileName: file.name,
                    fileSize: file.size,
                    fileType: 'CSV',
                    sheets: {
                        'Sheet1': {
                            json: jsonData,
                            csv: csvData,
                            range: worksheet['!ref'] || 'A1:A1'
                        }
                    },
                    sheetNames: ['Sheet1']
                };
                
                console.log('[processCSVFile] Successfully processed CSV file:', result.fileName);
                resolve(result);
            } catch (error) {
                console.error('[processCSVFile] Error processing CSV file:', error);
                reject(error);
            }
        };
        reader.onerror = () => reject(new Error('Failed to read file'));
        reader.readAsText(file);
    });
}

// Function to process Word document files (.docx, .doc)
function processWordFile(file) {
    return new Promise((resolve, reject) => {
        // Check if mammoth.js is available
        if (typeof mammoth === 'undefined') {
            reject(new Error('Mammoth.js library not loaded. Please refresh the page and try again.'));
            return;
        }
        
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const arrayBuffer = e.target.result;
                
                console.log('[processWordFile] Using mammoth.js to extract text from Word document');
                
                // Use mammoth.js to extract text from Word document
                mammoth.extractRawText({ arrayBuffer: arrayBuffer })
                    .then(function(mammothResult) {
                        const textContent = mammothResult.value; // The raw text
                        const messages = mammothResult.messages; // Any warnings
                        
                        // Log any warnings from mammoth
                        if (messages.length > 0) {
                            console.warn('[processWordFile] Mammoth warnings:', messages);
                        }
                        
                        // Split text into paragraphs for better formatting
                        const paragraphs = textContent.split(/\n\s*\n/).filter(p => p.trim().length > 0);
                        
                        const fileData = {
                            fileName: file.name,
                            fileSize: file.size,
                            fileType: 'WORD',
                            content: {
                                rawText: textContent,
                                paragraphs: paragraphs,
                                wordCount: textContent.split(/\s+/).filter(word => word.length > 0).length,
                                characterCount: textContent.length
                            }
                        };
                        
                        console.log('[processWordFile] Successfully processed Word file:', fileData.fileName);
                        console.log(`[processWordFile] Extracted ${fileData.content.wordCount} words, ${fileData.content.paragraphs.length} paragraphs`);
                        resolve(fileData);
                    })
                    .catch(function(error) {
                        console.error('[processWordFile] Error extracting text from Word document:', error);
                        reject(new Error(`Failed to process Word document: ${error.message}`));
                    });
                    
            } catch (error) {
                console.error('[processWordFile] Error processing Word file:', error);
                reject(error);
            }
        };
        reader.onerror = () => reject(new Error('Failed to read Word file'));
        reader.readAsArrayBuffer(file);
    });
}

// Function to process PDF files (.pdf)
function processPDFFile(file) {
    return new Promise((resolve, reject) => {
        // Check if PDF.js is available
        if (typeof pdfjsLib === 'undefined') {
            reject(new Error('PDF.js library not loaded. Please refresh the page and try again.'));
            return;
        }

        // Set PDF.js worker path
        pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js';

        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const arrayBuffer = e.target.result;
                
                console.log('[processPDFFile] Using PDF.js to extract text from PDF document');
                
                // Load PDF document
                pdfjsLib.getDocument({ data: arrayBuffer }).promise.then(function(pdf) {
                    console.log('[processPDFFile] PDF loaded with', pdf.numPages, 'pages');
                    
                    const promises = [];
                    const pages = [];
                    
                    // Create promises for all pages
                    for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
                        promises.push(
                            pdf.getPage(pageNum).then(function(page) {
                                return page.getTextContent().then(function(textContent) {
                                    const pageText = textContent.items.map(item => item.str).join(' ');
                                    return {
                                        pageNumber: pageNum,
                                        text: pageText,
                                        wordCount: pageText.split(/\s+/).filter(word => word.length > 0).length
                                    };
                                });
                            })
                        );
                    }
                    
                    // Wait for all pages to be processed
                    Promise.all(promises).then(function(pageResults) {
                        // Sort pages by page number (just in case)
                        pageResults.sort((a, b) => a.pageNumber - b.pageNumber);
                        
                        // Combine all text
                        const fullText = pageResults.map(page => page.text).join('\n\n');
                        const totalWordCount = pageResults.reduce((sum, page) => sum + page.wordCount, 0);
                        
                        // Split into paragraphs for better formatting
                        const paragraphs = fullText.split(/\n\s*\n/).filter(p => p.trim().length > 0);
                        
                        const fileData = {
                            fileName: file.name,
                            fileSize: file.size,
                            fileType: 'PDF',
                            content: {
                                rawText: fullText,
                                paragraphs: paragraphs,
                                pages: pageResults,
                                pageCount: pdf.numPages,
                                wordCount: totalWordCount,
                                characterCount: fullText.length
                            }
                        };
                        
                        console.log('[processPDFFile] Successfully processed PDF file:', fileData.fileName);
                        console.log(`[processPDFFile] Extracted ${fileData.content.wordCount} words from ${fileData.content.pageCount} pages`);
                        resolve(fileData);
                        
                    }).catch(function(error) {
                        console.error('[processPDFFile] Error extracting text from PDF pages:', error);
                        reject(new Error(`Failed to extract text from PDF pages: ${error.message}`));
                    });
                    
                }).catch(function(error) {
                    console.error('[processPDFFile] Error loading PDF document:', error);
                    reject(new Error(`Failed to load PDF document: ${error.message}`));
                });
                
            } catch (error) {
                console.error('[processPDFFile] Error processing PDF file:', error);
                reject(error);
            }
        };
        reader.onerror = () => reject(new Error('Failed to read PDF file'));
        reader.readAsArrayBuffer(file);
    });
}

// Function to process text files (.txt)
function processTXTFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const textContent = e.target.result;
                
                console.log('[processTXTFile] Processing plain text file');
                
                // Split text into paragraphs for better formatting
                const paragraphs = textContent.split(/\n\s*\n/).filter(p => p.trim().length > 0);
                
                // Split into lines for line count
                const lines = textContent.split(/\n/);
                
                const fileData = {
                    fileName: file.name,
                    fileSize: file.size,
                    fileType: 'TXT',
                    content: {
                        rawText: textContent,
                        paragraphs: paragraphs,
                        lines: lines,
                        lineCount: lines.length,
                        wordCount: textContent.split(/\s+/).filter(word => word.length > 0).length,
                        characterCount: textContent.length
                    }
                };
                
                console.log('[processTXTFile] Successfully processed text file:', fileData.fileName);
                console.log(`[processTXTFile] Extracted ${fileData.content.wordCount} words, ${fileData.content.lineCount} lines, ${fileData.content.paragraphs.length} paragraphs`);
                resolve(fileData);
                
            } catch (error) {
                console.error('[processTXTFile] Error processing text file:', error);
                reject(error);
            }
        };
        reader.onerror = () => reject(new Error('Failed to read text file'));
        reader.readAsText(file, 'UTF-8'); // Specify UTF-8 encoding
    });
}

// Function to format file data for AI consumption (supports multiple files)
function formatFileDataForAI(filesData) {
    if (!Array.isArray(filesData)) {
        filesData = [filesData]; // Convert single file to array for compatibility
    }
    
    if (filesData.length === 0) {
        return '';
    }
    
    let formattedData = `\n\n📎 **Attached Files (${filesData.length}):**\n\n`;
    
    filesData.forEach((fileData, fileIndex) => {
        formattedData += `**File ${fileIndex + 1}: ${fileData.fileName}** (${formatFileSize(fileData.fileSize)})\n`;
        formattedData += `File Type: ${fileData.fileType}\n`;
        
        if (fileData.fileType === 'WORD') {
            // Format Word document content
            formattedData += `Word Count: ${fileData.content.wordCount}\n`;
            formattedData += `Character Count: ${fileData.content.characterCount}\n`;
            formattedData += `Paragraphs: ${fileData.content.paragraphs.length}\n\n`;
            
            formattedData += `**Document Content:**\n`;
            formattedData += '```\n';
            
            // Show first 2000 characters with paragraph breaks
            const previewText = fileData.content.rawText.length > 2000 
                ? fileData.content.rawText.substring(0, 2000) + '...' 
                : fileData.content.rawText;
            
            formattedData += previewText;
            formattedData += '\n```\n\n';
            
            if (fileData.content.rawText.length > 2000) {
                formattedData += `*Note: Showing first 2000 characters of ${fileData.content.characterCount} total characters.*\n\n`;
            }
        } else if (fileData.fileType === 'TXT') {
            // Format text file content
            formattedData += `Line Count: ${fileData.content.lineCount}\n`;
            formattedData += `Word Count: ${fileData.content.wordCount}\n`;
            formattedData += `Character Count: ${fileData.content.characterCount}\n`;
            formattedData += `Paragraphs: ${fileData.content.paragraphs.length}\n\n`;
            
            formattedData += `**Text Content:**\n`;
            formattedData += '```\n';
            
            // Show first 2000 characters with line breaks preserved
            const previewText = fileData.content.rawText.length > 2000 
                ? fileData.content.rawText.substring(0, 2000) + '...' 
                : fileData.content.rawText;
            
            formattedData += previewText;
            formattedData += '\n```\n\n';
            
            if (fileData.content.rawText.length > 2000) {
                formattedData += `*Note: Showing first 2000 characters of ${fileData.content.characterCount} total characters from ${fileData.content.lineCount} lines.*\n\n`;
            }
        } else if (fileData.fileType === 'PDF') {
            // Format PDF document content
            formattedData += `Page Count: ${fileData.content.pageCount}\n`;
            formattedData += `Word Count: ${fileData.content.wordCount}\n`;
            formattedData += `Character Count: ${fileData.content.characterCount}\n`;
            formattedData += `Paragraphs: ${fileData.content.paragraphs.length}\n\n`;
            
            formattedData += `**Document Content:**\n`;
            formattedData += '```\n';
            
            // Show first 2000 characters with paragraph breaks
            const previewText = fileData.content.rawText.length > 2000 
                ? fileData.content.rawText.substring(0, 2000) + '...' 
                : fileData.content.rawText;
            
            formattedData += previewText;
            formattedData += '\n```\n\n';
            
            if (fileData.content.rawText.length > 2000) {
                formattedData += `*Note: Showing first 2000 characters of ${fileData.content.characterCount} total characters from ${fileData.content.pageCount} pages.*\n\n`;
            }
        } else {
            // Format Excel/CSV data
            formattedData += `Number of Sheets: ${fileData.sheetNames.length}\n\n`;
            
            // Process each sheet
            fileData.sheetNames.forEach((sheetName, index) => {
                const sheet = fileData.sheets[sheetName];
                formattedData += `**Sheet ${index + 1}: ${sheetName}**\n`;
                formattedData += `Range: ${sheet.range}\n`;
                
                if (sheet.json && sheet.json.length > 0) {
                    formattedData += `Data Preview (first 10 rows):\n`;
                    formattedData += '```\n';
                    
                    // Show first 10 rows of data
                    const previewRows = sheet.json.slice(0, 10);
                    previewRows.forEach((row, rowIndex) => {
                        const rowData = row.map(cell => cell !== null && cell !== undefined ? String(cell) : '').join(' | ');
                        formattedData += `Row ${rowIndex + 1}: ${rowData}\n`;
                    });
                    
                    if (sheet.json.length > 10) {
                        formattedData += `... (${sheet.json.length - 10} more rows)\n`;
                    }
                    formattedData += '```\n\n';
                } else {
                    formattedData += 'No data found in this sheet.\n\n';
                }
            });
        }
        
        if (fileIndex < filesData.length - 1) {
            formattedData += `\n${'─'.repeat(50)}\n\n`;
        }
    });
    
    formattedData += `Please analyze ${filesData.length === 1 ? 'this file' : 'these files'} and help me build a financial model based on the information provided.`;
    
    return formattedData;
}

// Function to handle file attachment
async function handleFileAttachment(file) {
    try {
        console.log('[handleFileAttachment] Processing file:', file.name);
        console.log('[handleFileAttachment] File type (MIME):', file.type);
        console.log('[handleFileAttachment] File size:', file.size);
        
        // Validate file type
        const allowedTypes = [
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', // .xlsx
            'application/vnd.ms-excel', // .xls
            'application/vnd.ms-excel.sheet.macroEnabled.12', // .xlsm
            'text/csv', // .csv
            'application/vnd.openxmlformats-officedocument.wordprocessingml.document', // .docx
            'application/msword', // .doc
            'application/pdf', // .pdf
            'text/plain' // .txt
        ];
        
        const fileExtension = file.name.toLowerCase().split('.').pop();
        console.log('[handleFileAttachment] File extension:', fileExtension);
        
        const mimeTypeAllowed = allowedTypes.includes(file.type);
        const extensionAllowed = ['xlsx', 'xls', 'xlsm', 'csv', 'doc', 'docx', 'pdf', 'txt'].includes(fileExtension);
        
        console.log('[handleFileAttachment] MIME type allowed:', mimeTypeAllowed);
        console.log('[handleFileAttachment] Extension allowed:', extensionAllowed);
        
        if (!mimeTypeAllowed && !extensionAllowed) {
            console.log('[handleFileAttachment] File validation failed - neither MIME type nor extension is allowed');
            throw new Error('Please upload an Excel file (.xlsx, .xls, .xlsm), CSV file (.csv), Word document (.doc, .docx), PDF file (.pdf), or Text file (.txt)');
        }
        
        // Validate file size (max 10MB)
        const maxSize = 10 * 1024 * 1024; // 10MB
        if (file.size > maxSize) {
            throw new Error('File size must be less than 10MB');
        }
        
        // Process file based on type
        let fileData;
        if (fileExtension === 'csv' || file.type === 'text/csv') {
            console.log('[handleFileAttachment] Processing as CSV file');
            fileData = await processCSVFile(file);
        } else if (fileExtension === 'doc' || fileExtension === 'docx' || 
                   file.type === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' ||
                   file.type === 'application/msword') {
            console.log('[handleFileAttachment] Processing as Word document');
            fileData = await processWordFile(file);
        } else if (fileExtension === 'pdf' || file.type === 'application/pdf') {
            console.log('[handleFileAttachment] Processing as PDF file');
            fileData = await processPDFFile(file);
        } else if (fileExtension === 'txt' || file.type === 'text/plain') {
            console.log('[handleFileAttachment] Processing as Text file');
            fileData = await processTXTFile(file);
        } else {
            console.log('[handleFileAttachment] Processing as Excel file (.xlsx/.xls/.xlsm)');
            fileData = await processXLSXFile(file);
        }
        
        // Add the processed file data to the array
        currentAttachedFiles.push(fileData);
        
        // Update UI to show attached files
        showAttachedFiles(currentAttachedFiles);
        
        console.log('[handleFileAttachment] File processed successfully:', fileData.fileName);
        console.log('[handleFileAttachment] Total attached files:', currentAttachedFiles.length);
        return fileData;
        
    } catch (error) {
        console.error('[handleFileAttachment] Error processing file:', error);
        // Show error to user
        displayInClientChatLogPlanner(`Error processing file: ${error.message}`, false);
        throw error;
    }
}

// Function to show attached files in UI
function showAttachedFiles(filesData) {
    const attachedFilesContainer = document.getElementById('attached-files-container');
    const attachedFilesList = document.getElementById('attached-files-list');
    const attachedFilesCount = document.querySelector('.attached-files-count');
    
    if (!attachedFilesContainer || !attachedFilesList || !attachedFilesCount) {
        console.warn('[showAttachedFiles] Required elements not found');
        return;
    }
    
    // Update count
    attachedFilesCount.textContent = `${filesData.length} file${filesData.length !== 1 ? 's' : ''} attached`;
    
    // Clear existing file items
    attachedFilesList.innerHTML = '';
    
    // Create file items
    filesData.forEach((fileData, index) => {
        const fileItem = document.createElement('div');
        fileItem.className = 'attached-file-item';
        fileItem.setAttribute('data-file-index', index);
        
        // Get file type badge color based on type
        let badgeClass = 'file-type-badge';
        let badgeText = fileData.fileType;
        
        fileItem.innerHTML = `
            <div class="attached-file-info">
                <span class="file-name" title="${fileData.fileName}">${fileData.fileName}</span>
                <span class="file-size">${formatFileSize(fileData.fileSize)}</span>
                <span class="${badgeClass}">${badgeText}</span>
            </div>
            <button class="remove-file-btn" data-file-index="${index}" title="Remove ${fileData.fileName}">×</button>
        `;
        
        // Add remove file event listener
        const removeBtn = fileItem.querySelector('.remove-file-btn');
        removeBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            removeFileByIndex(parseInt(e.target.getAttribute('data-file-index')));
        });
        
        attachedFilesList.appendChild(fileItem);
    });
    
    // Show the container
    attachedFilesContainer.style.display = 'block';
    
    console.log('[showAttachedFiles] Showing', filesData.length, 'attached files in UI');
}

// Function to hide attached files in UI
function hideAttachedFiles() {
    const attachedFilesContainer = document.getElementById('attached-files-container');
    
    if (attachedFilesContainer) {
        attachedFilesContainer.style.display = 'none';
        console.log('[hideAttachedFiles] Hidden attached files from UI');
    }
}

// Function to remove a specific file by index
function removeFileByIndex(index) {
    if (index >= 0 && index < currentAttachedFiles.length) {
        const removedFile = currentAttachedFiles.splice(index, 1)[0];
        console.log('[removeFileByIndex] Removed file:', removedFile.fileName);
        
        // Update UI
        if (currentAttachedFiles.length > 0) {
            showAttachedFiles(currentAttachedFiles);
        } else {
            hideAttachedFiles();
        }
    }
}

// Function to remove all attachments
function removeAllAttachments() {
    currentAttachedFiles = [];
    hideAttachedFiles();
    console.log('[removeAllAttachments] All attachments removed');
}

// Function to initialize file attachment event listeners
export function initializeFileAttachment() {
    console.log('[initializeFileAttachment] Setting up multiple file attachment listeners with full Excel support (including .xlsm) - VERSION 5.1');
    
    // Get elements
    const attachFileButton = document.getElementById('attach-file-client');
    const fileInput = document.getElementById('file-input-client');
    const clearAllButton = document.getElementById('clear-all-attachments');
    
    if (!attachFileButton || !fileInput) {
        console.warn('[initializeFileAttachment] Required attachment elements not found');
        return;
    }
    
    // Attach file button click
    attachFileButton.addEventListener('click', () => {
        console.log('[initializeFileAttachment] Attach file button clicked');
        fileInput.click();
    });
    
    // File input change - handle multiple files
    fileInput.addEventListener('change', async (e) => {
        const files = Array.from(e.target.files);
        if (files.length > 0) {
            console.log('[initializeFileAttachment] Files selected:', files.length);
            
            // Process each file
            for (const file of files) {
                console.log('[initializeFileAttachment] Processing file:', file.name);
                try {
                    await handleFileAttachment(file);
                } catch (error) {
                    // Error already handled in handleFileAttachment
                    console.warn('[initializeFileAttachment] Failed to process file:', file.name);
                }
            }
        }
        // Clear the input so the same files can be selected again
        fileInput.value = '';
    });
    
    // Clear all attachments button click
    if (clearAllButton) {
        clearAllButton.addEventListener('click', () => {
            console.log('[initializeFileAttachment] Clear all attachments button clicked');
            removeAllAttachments();
        });
    }
    
    console.log('[initializeFileAttachment] Multiple file attachment listeners set up successfully');
}

// Developer Mode File Attachment Variables
let currentAttachedFilesDev = [];

// Export developer mode file attachment functions and variables
export { currentAttachedFilesDev, formatFileDataForAIDev, removeAllAttachmentsDev };

// Function to initialize file attachment event listeners for Developer Mode
export function initializeFileAttachmentDev() {
    console.log('[initializeFileAttachmentDev] Setting up developer mode file attachment listeners - VERSION 1.0');
    
    // Get elements
    const attachFileButton = document.getElementById('attach-file-dev');
    const fileInput = document.getElementById('file-input-dev');
    const clearAllButton = document.getElementById('clear-all-attachments-dev');
    
    if (!attachFileButton || !fileInput) {
        console.warn('[initializeFileAttachmentDev] Required attachment elements not found');
        return;
    }
    
    // Attach file button click
    attachFileButton.addEventListener('click', () => {
        console.log('[initializeFileAttachmentDev] Attach file button clicked');
        fileInput.click();
    });
    
    // File input change - handle multiple files
    fileInput.addEventListener('change', async (e) => {
        const files = Array.from(e.target.files);
        if (files.length > 0) {
            console.log('[initializeFileAttachmentDev] Files selected:', files.length);
            
            // Process each file
            for (const file of files) {
                console.log('[initializeFileAttachmentDev] Processing file:', file.name);
                try {
                    await handleFileAttachmentDev(file);
                } catch (error) {
                    // Error already handled in handleFileAttachmentDev
                    console.warn('[initializeFileAttachmentDev] Failed to process file:', file.name);
                }
            }
        }
        // Clear the input so the same files can be selected again
        fileInput.value = '';
    });
    
    // Clear all attachments button
    if (clearAllButton) {
        clearAllButton.addEventListener('click', () => {
            console.log('[initializeFileAttachmentDev] Clear all attachments clicked');
            removeAllAttachmentsDev();
        });
    }
    
    console.log('[initializeFileAttachmentDev] Developer mode file attachment functionality initialized successfully');
}

// Function to handle file attachment for Developer Mode
async function handleFileAttachmentDev(file) {
    try {
        console.log('[handleFileAttachmentDev] Processing file:', file.name);
        console.log('[handleFileAttachmentDev] File type (MIME):', file.type);
        console.log('[handleFileAttachmentDev] File size:', file.size);
        
        // Validate file type
        const allowedTypes = [
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', // .xlsx
            'application/vnd.ms-excel', // .xls
            'application/vnd.ms-excel.sheet.macroEnabled.12', // .xlsm
            'text/csv', // .csv
            'application/vnd.openxmlformats-officedocument.wordprocessingml.document', // .docx
            'application/msword', // .doc
            'application/pdf', // .pdf
            'text/plain' // .txt
        ];
        
        const fileExtension = file.name.toLowerCase().split('.').pop();
        console.log('[handleFileAttachmentDev] File extension:', fileExtension);
        
        const mimeTypeAllowed = allowedTypes.includes(file.type);
        const extensionAllowed = ['xlsx', 'xls', 'xlsm', 'csv', 'doc', 'docx', 'pdf', 'txt'].includes(fileExtension);
        
        console.log('[handleFileAttachmentDev] MIME type allowed:', mimeTypeAllowed);
        console.log('[handleFileAttachmentDev] Extension allowed:', extensionAllowed);
        
        if (!mimeTypeAllowed && !extensionAllowed) {
            console.log('[handleFileAttachmentDev] File validation failed - neither MIME type nor extension is allowed');
            throw new Error('Please upload an Excel file (.xlsx, .xls, .xlsm), CSV file (.csv), Word document (.doc, .docx), PDF file (.pdf), or Text file (.txt)');
        }
        
        // Validate file size (max 10MB)
        const maxSize = 10 * 1024 * 1024; // 10MB
        if (file.size > maxSize) {
            throw new Error('File size must be less than 10MB');
        }
        
        // Process file based on type (reuse existing functions)
        let fileData;
        if (fileExtension === 'csv' || file.type === 'text/csv') {
            console.log('[handleFileAttachmentDev] Processing as CSV file');
            fileData = await processCSVFile(file);
        } else if (fileExtension === 'doc' || fileExtension === 'docx' || 
                   file.type === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' ||
                   file.type === 'application/msword') {
            console.log('[handleFileAttachmentDev] Processing as Word document');
            fileData = await processWordFile(file);
        } else if (fileExtension === 'pdf' || file.type === 'application/pdf') {
            console.log('[handleFileAttachmentDev] Processing as PDF file');
            fileData = await processPDFFile(file);
        } else if (fileExtension === 'txt' || file.type === 'text/plain') {
            console.log('[handleFileAttachmentDev] Processing as Text file');
            fileData = await processTXTFile(file);
        } else {
            console.log('[handleFileAttachmentDev] Processing as Excel file (.xlsx/.xls/.xlsm)');
            fileData = await processXLSXFile(file);
        }
        
        // Add the processed file data to the developer array
        currentAttachedFilesDev.push(fileData);
        
        // Update UI to show attached files
        showAttachedFilesDev(currentAttachedFilesDev);
        
        console.log('[handleFileAttachmentDev] File processed successfully:', fileData.fileName);
        console.log('[handleFileAttachmentDev] Total attached files:', currentAttachedFilesDev.length);
        return fileData;
        
    } catch (error) {
        console.error('[handleFileAttachmentDev] Error processing file:', error);
        // Show error to user in developer mode
        displayInDeveloperChat(`Error processing file: ${error.message}`, false);
        throw error;
    }
}

// Function to show attached files in Developer Mode UI
function showAttachedFilesDev(filesData) {
    const attachedFilesContainer = document.getElementById('attached-files-container-dev');
    const attachedFilesList = document.getElementById('attached-files-list-dev');
    const attachedFilesCount = document.querySelector('.attached-files-count-dev');
    
    if (!attachedFilesContainer || !attachedFilesList || !attachedFilesCount) {
        console.warn('[showAttachedFilesDev] Required elements not found');
        return;
    }
    
    // Show container if files are attached
    if (filesData.length > 0) {
        attachedFilesContainer.style.display = 'block';
    } else {
        attachedFilesContainer.style.display = 'none';
        return;
    }
    
    // Update count
    attachedFilesCount.textContent = `${filesData.length} file${filesData.length !== 1 ? 's' : ''} attached`;
    
    // Clear existing file items
    attachedFilesList.innerHTML = '';
    
    // Create file items
    filesData.forEach((fileData, index) => {
        const fileItem = document.createElement('div');
        fileItem.className = 'attached-file-item';
        fileItem.setAttribute('data-file-index', index);
        
        // Get file type badge color based on type
        let badgeClass = 'file-type-badge';
        let badgeText = fileData.fileType;
        
        fileItem.innerHTML = `
            <div class="attached-file-info">
                <span class="file-name" title="${fileData.fileName}">${fileData.fileName}</span>
                <span class="file-size">${formatFileSize(fileData.fileSize)}</span>
                <span class="${badgeClass}">${badgeText}</span>
            </div>
            <button class="remove-file-btn" data-file-index="${index}" title="Remove ${fileData.fileName}">×</button>
        `;
        
        // Add remove file event listener
        const removeBtn = fileItem.querySelector('.remove-file-btn');
        removeBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            removeFileByIndexDev(parseInt(e.target.getAttribute('data-file-index')));
        });
        
        attachedFilesList.appendChild(fileItem);
    });
}

// Function to remove file by index for Developer Mode
function removeFileByIndexDev(index) {
    if (index >= 0 && index < currentAttachedFilesDev.length) {
        const removedFile = currentAttachedFilesDev.splice(index, 1)[0];
        console.log('[removeFileByIndexDev] Removed file:', removedFile.fileName);
        
        // Update UI
        showAttachedFilesDev(currentAttachedFilesDev);
    }
}

// Function to remove all attachments for Developer Mode
function removeAllAttachmentsDev() {
    console.log('[removeAllAttachmentsDev] Removing all attached files');
    currentAttachedFilesDev = [];
    
    // Hide the container
    const attachedFilesContainer = document.getElementById('attached-files-container-dev');
    if (attachedFilesContainer) {
        attachedFilesContainer.style.display = 'none';
    }
    
    // Clear the list
    const attachedFilesList = document.getElementById('attached-files-list-dev');
    if (attachedFilesList) {
        attachedFilesList.innerHTML = '';
    }
    
    console.log('[removeAllAttachmentsDev] All files removed');
}

// Function to format file data for AI consumption (Developer Mode)
function formatFileDataForAIDev(filesData) {
    if (!Array.isArray(filesData)) {
        filesData = [filesData]; // Convert single file to array for compatibility
    }
    
    if (filesData.length === 0) {
        return '';
    }
    
    let formattedData = `\n\n📎 **Attached Files (${filesData.length}):**\n\n`;
    
    filesData.forEach((fileData, fileIndex) => {
        formattedData += `**File ${fileIndex + 1}: ${fileData.fileName}** (${formatFileSize(fileData.fileSize)})\n`;
        formattedData += `File Type: ${fileData.fileType}\n`;
        
        if (fileData.fileType === 'WORD') {
            // Format Word document content
            formattedData += `Word Count: ${fileData.content.wordCount}\n`;
            formattedData += `Character Count: ${fileData.content.characterCount}\n`;
            formattedData += `Paragraphs: ${fileData.content.paragraphs.length}\n\n`;
            
            formattedData += `**Document Content:**\n`;
            formattedData += '```\n';
            
            // Show first 2000 characters with paragraph breaks
            const previewText = fileData.content.rawText.length > 2000 
                ? fileData.content.rawText.substring(0, 2000) + '...' 
                : fileData.content.rawText;
            
            formattedData += previewText;
            formattedData += '\n```\n\n';
            
            if (fileData.content.rawText.length > 2000) {
                formattedData += `*Note: Showing first 2000 characters of ${fileData.content.characterCount} total characters.*\n\n`;
            }
        } else if (fileData.fileType === 'EXCEL' || fileData.fileType === 'CSV') {
            // Format spreadsheet content
            formattedData += `Sheets: ${fileData.content.sheetNames.join(', ')}\n`;
            formattedData += `Total Rows: ${fileData.content.totalRows}\n`;
            formattedData += `Total Columns: ${fileData.content.totalColumns}\n\n`;
            
            // Show preview of each sheet
            Object.keys(fileData.content.sheets).forEach(sheetName => {
                const sheet = fileData.content.sheets[sheetName];
                formattedData += `**Sheet: ${sheetName}**\n`;
                formattedData += '```\n';
                formattedData += sheet.preview;
                formattedData += '\n```\n\n';
            });
        } else if (fileData.fileType === 'PDF') {
            // Format PDF content
            formattedData += `Pages: ${fileData.content.pageCount}\n`;
            formattedData += `Character Count: ${fileData.content.characterCount}\n\n`;
            
            formattedData += `**Document Content:**\n`;
            formattedData += '```\n';
            
            // Show first 2000 characters
            const previewText = fileData.content.text.length > 2000 
                ? fileData.content.text.substring(0, 2000) + '...' 
                : fileData.content.text;
            
            formattedData += previewText;
            formattedData += '\n```\n\n';
            
            if (fileData.content.text.length > 2000) {
                formattedData += `*Note: Showing first 2000 characters of ${fileData.content.characterCount} total characters.*\n\n`;
            }
        } else if (fileData.fileType === 'TXT') {
            // Format text content
            formattedData += `Character Count: ${fileData.content.characterCount}\n`;
            formattedData += `Line Count: ${fileData.content.lineCount}\n\n`;
            
            formattedData += `**File Content:**\n`;
            formattedData += '```\n';
            
            // Show first 2000 characters
            const previewText = fileData.content.text.length > 2000 
                ? fileData.content.text.substring(0, 2000) + '...' 
                : fileData.content.text;
            
            formattedData += previewText;
            formattedData += '\n```\n\n';
            
            if (fileData.content.text.length > 2000) {
                formattedData += `*Note: Showing first 2000 characters of ${fileData.content.characterCount} total characters.*\n\n`;
            }
        }
        
        if (fileIndex < filesData.length - 1) {
            formattedData += '---\n\n';
        }
    });
    
    return formattedData;
}

// ========== VOICE INPUT FUNCTIONALITY ==========

let mediaRecorder = null;
let audioChunks = [];
let recordingTimer = null;
let recordingStartTime = null;
let audioContext = null;
let analyser = null;
let waveformCanvas = null;
let waveformAnimationId = null;

// Function to format recording time
function formatRecordingTime(seconds) {
    const mins = Math.floor(seconds / 60);
    const secs = seconds % 60;
    return `${mins.toString().padStart(2, '0')}:${secs.toString().padStart(2, '0')}`;
}

// Function to switch to voice recording mode
function switchToVoiceMode() {
    const normalMode = document.getElementById('normal-input-mode');
    const voiceMode = document.getElementById('voice-recording-mode');
    
    if (normalMode && voiceMode) {
        normalMode.style.display = 'none';
        voiceMode.style.display = 'flex';
        
        // Initialize waveform canvas
        initializeWaveform();
        
        // Start recording immediately
        startVoiceRecording();
    }
}

// Function to switch back to normal input mode
function switchToNormalMode() {
    const normalMode = document.getElementById('normal-input-mode');
    const voiceMode = document.getElementById('voice-recording-mode');
    const loadingMode = document.getElementById('transcription-loading-mode');
    
    if (normalMode && voiceMode && loadingMode) {
        normalMode.style.display = 'flex';
        voiceMode.style.display = 'none';
        loadingMode.style.display = 'none';
        
        // Stop any ongoing recording
        if (mediaRecorder && mediaRecorder.state === 'recording') {
            stopVoiceRecording();
        }
        
        // Cleanup waveform
        cleanupWaveform();
        resetVoiceState();
    }
}

// Function to switch to transcription loading mode
function switchToLoadingMode() {
    const normalMode = document.getElementById('normal-input-mode');
    const voiceMode = document.getElementById('voice-recording-mode');
    const loadingMode = document.getElementById('transcription-loading-mode');
    
    if (normalMode && voiceMode && loadingMode) {
        normalMode.style.display = 'none';
        voiceMode.style.display = 'none';
        loadingMode.style.display = 'flex';
    }
}

// Function to reset voice state
function resetVoiceState() {
    const timer = document.getElementById('voice-recording-timer');
    
    if (timer) {
        timer.textContent = '00:00';
    }
    
    // Clear timer if running
    if (recordingTimer) {
        clearInterval(recordingTimer);
        recordingTimer = null;
    }
    
    // Reset audio data
    audioChunks = [];
    recordingStartTime = null;
}

// Function to initialize waveform visualization
function initializeWaveform() {
    waveformCanvas = document.getElementById('voice-waveform');
    if (!waveformCanvas) return;
    
    const ctx = waveformCanvas.getContext('2d');
    const rect = waveformCanvas.getBoundingClientRect();
    
    // Set canvas size to match display size
    waveformCanvas.width = rect.width * window.devicePixelRatio;
    waveformCanvas.height = rect.height * window.devicePixelRatio;
    ctx.scale(window.devicePixelRatio, window.devicePixelRatio);
    
    // Style the canvas
    waveformCanvas.style.width = rect.width + 'px';
    waveformCanvas.style.height = rect.height + 'px';
}

// Function to cleanup waveform
function cleanupWaveform() {
    console.log('[Voice] Cleaning up waveform visualization');
    
    if (waveformAnimationId) {
        cancelAnimationFrame(waveformAnimationId);
        waveformAnimationId = null;
    }
    
    if (audioContext && audioContext.state !== 'closed') {
        audioContext.close().catch(err => {
            console.warn('[Voice] Error closing audio context:', err);
        });
        audioContext = null;
    }
    
    analyser = null;
    
    // Clear the canvas
    if (waveformCanvas) {
        const ctx = waveformCanvas.getContext('2d');
        const rect = waveformCanvas.getBoundingClientRect();
        ctx.clearRect(0, 0, rect.width, rect.height);
    }
}

// Function to draw waveform
function drawWaveform() {
    if (!waveformCanvas || !analyser) {
        return;
    }
    
    const ctx = waveformCanvas.getContext('2d');
    const rect = waveformCanvas.getBoundingClientRect();
    const width = rect.width;
    const height = rect.height;
    
    // Get time domain data for real-time audio visualization
    const bufferLength = analyser.fftSize;
    const dataArray = new Uint8Array(bufferLength);
    analyser.getByteTimeDomainData(dataArray);
    
    // Clear canvas with transparent background
    ctx.clearRect(0, 0, width, height);
    
    // Calculate how many bars we want to show
    const numBars = Math.min(60, Math.floor(width / 4)); // Adjust bar count based on width
    const barWidth = (width - (numBars - 1)) / numBars; // Account for spacing
    
    // Process audio data to create frequency-style visualization
    const samplesPerBar = Math.floor(bufferLength / numBars);
    
    for (let i = 0; i < numBars; i++) {
        let sum = 0;
        let amplitude = 0;
        
        // Calculate RMS (root mean square) for this bar's audio data
        for (let j = 0; j < samplesPerBar; j++) {
            const sample = (dataArray[i * samplesPerBar + j] - 128) / 128; // Normalize to -1 to 1
            sum += sample * sample;
        }
        
        amplitude = Math.sqrt(sum / samplesPerBar);
        
        // Add some smoothing and make it more responsive
        const minHeight = 3; // Minimum bar height for better visibility
        const maxHeight = height * 0.8;
        
        // Amplify the signal for better visual feedback
        const amplifiedAmplitude = Math.pow(amplitude * 2, 0.8); // Power curve for better visual response
        let barHeight = minHeight + (amplifiedAmplitude * maxHeight);
        
        // Add subtle animation for visual interest
        const timeOffset = Date.now() * 0.005; // Slow time-based animation
        const animationFactor = Math.sin(timeOffset + i * 0.2) * 0.1; // Slight wave effect
        barHeight += animationFactor * 3;
        
        // Add some randomness for visual interest when there's audio
        if (amplitude > 0.005) { // Lower threshold for better responsiveness
            barHeight += Math.random() * 3;
        }
        
        // Ensure minimum visibility and cap maximum
        barHeight = Math.max(barHeight, minHeight);
        barHeight = Math.min(barHeight, maxHeight);
        
        // Calculate x position with spacing
        const x = i * (barWidth + 1);
        
        // Draw the bar centered vertically
        const y = (height - barHeight) / 2;
        
        ctx.fillStyle = '#000000';
        ctx.fillRect(x, y, barWidth, barHeight);
    }
    
    // Continue animation - this ensures continuous updates
    if (mediaRecorder && mediaRecorder.state === 'recording') {
        waveformAnimationId = requestAnimationFrame(drawWaveform);
    }
}

// Function to start audio recording with waveform
async function startVoiceRecording() {
    try {
        console.log('[Voice] Requesting microphone access...');
        
        const stream = await navigator.mediaDevices.getUserMedia({ 
            audio: {
                echoCancellation: true,
                noiseSuppression: true,
                sampleRate: 44100
            } 
        });
        
        console.log('[Voice] Microphone access granted');
        
        // Setup audio context for waveform visualization
        audioContext = new (window.AudioContext || window.webkitAudioContext)();
        analyser = audioContext.createAnalyser();
        const source = audioContext.createMediaStreamSource(stream);
        source.connect(analyser);
        
        // Configure analyser for real-time visualization
        analyser.fftSize = 1024; // Higher resolution for better visualization
        analyser.smoothingTimeConstant = 0.3; // Smooth but responsive
        analyser.minDecibels = -90;
        analyser.maxDecibels = -10;
        
        // Start waveform animation
        drawWaveform();
        
        // Backup timer to ensure continuous animation (in case requestAnimationFrame fails)
        const backupTimer = setInterval(() => {
            if (mediaRecorder && mediaRecorder.state === 'recording' && !waveformAnimationId) {
                console.log('[Voice] Restarting waveform animation via backup timer');
                drawWaveform();
            } else if (!mediaRecorder || mediaRecorder.state !== 'recording') {
                clearInterval(backupTimer);
            }
        }, 100); // Check every 100ms
        
        // Create MediaRecorder
        const options = { mimeType: 'audio/webm;codecs=opus' };
        if (!MediaRecorder.isTypeSupported(options.mimeType)) {
            options.mimeType = 'audio/webm';
            if (!MediaRecorder.isTypeSupported(options.mimeType)) {
                options.mimeType = 'audio/mp4';
            }
        }
        
        mediaRecorder = new MediaRecorder(stream, options);
        audioChunks = [];
        
        mediaRecorder.ondataavailable = (event) => {
            if (event.data.size > 0) {
                audioChunks.push(event.data);
            }
        };
        
        mediaRecorder.onstop = async () => {
            console.log('[Voice] Recording stopped, processing audio...');
            
            // Stop all tracks to release microphone
            stream.getTracks().forEach(track => track.stop());
            
            // Switch to loading mode
            switchToLoadingMode();
            
            if (audioChunks.length > 0) {
                await processRecordedAudio();
            }
        };
        
        // Start recording
        mediaRecorder.start();
        recordingStartTime = Date.now();
        
        console.log('[Voice] Recording started with continuous waveform visualization');
        
        // Start timer
        startRecordingTimer();
        
    } catch (error) {
        console.error('[Voice] Error accessing microphone:', error);
        alert('Unable to access microphone. Please check your permissions and try again.');
        switchToNormalMode();
    }
}

// Function to stop audio recording
function stopVoiceRecording() {
    if (mediaRecorder && mediaRecorder.state === 'recording') {
        console.log('[Voice] Stopping recording...');
        mediaRecorder.stop();
        
        // Clear timer
        if (recordingTimer) {
            clearInterval(recordingTimer);
            recordingTimer = null;
        }
        
        // Stop waveform animation
        cleanupWaveform();
    }
}

// Function to start recording timer
function startRecordingTimer() {
    const timer = document.getElementById('voice-recording-timer');
    if (timer) {
        recordingTimer = setInterval(() => {
            if (recordingStartTime) {
                const elapsed = Math.floor((Date.now() - recordingStartTime) / 1000);
                timer.textContent = formatRecordingTime(elapsed);
            }
        }, 1000);
    }
}

// Function to process recorded audio and send to OpenAI
async function processRecordedAudio() {
    try {
        console.log('[Voice] Creating audio blob from chunks...');
        
        // Create blob from audio chunks
        const audioBlob = new Blob(audioChunks, { type: 'audio/webm' });
        console.log('[Voice] Audio blob created, size:', audioBlob.size, 'bytes');
        
        if (audioBlob.size === 0) {
            throw new Error('No audio data recorded');
        }
        
        // Send to OpenAI transcription API
        const transcription = await transcribeAudio(audioBlob);
        
        // Use the transcribed text immediately
        useTranscribedText(transcription);
        
    } catch (error) {
        console.error('[Voice] Error processing audio:', error);
        alert(`Error processing audio: ${error.message}`);
        switchToNormalMode();
    }
}

// Function to transcribe audio using OpenAI API
async function transcribeAudio(audioBlob) {
    try {
        console.log('[Voice] Sending audio to OpenAI transcription API...');
        
        // Check if API key is available
        if (!AI_MODEL_PLANNER_OPENAI_API_KEY) {
            throw new Error('OpenAI API key not available. Please set up your API key first.');
        }
        
        // Create form data
        const formData = new FormData();
        formData.append('file', audioBlob, 'audio.webm');
        formData.append('model', 'whisper-1');
        formData.append('language', 'en'); // Can be made configurable
        
        // Make API call
        const response = await fetch('https://api.openai.com/v1/audio/transcriptions', {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${AI_MODEL_PLANNER_OPENAI_API_KEY}`
            },
            body: formData
        });
        
        if (!response.ok) {
            const errorData = await response.json().catch(() => ({ message: "Failed to parse error response" }));
            throw new Error(`OpenAI API error: ${response.status} ${response.statusText} - ${errorData.message || JSON.stringify(errorData)}`);
        }
        
        const result = await response.json();
        console.log('[Voice] Transcription successful:', result.text);
        
        return result.text;
        
    } catch (error) {
        console.error('[Voice] Transcription error:', error);
        throw error;
    }
}

// Function to use transcribed text
function useTranscribedText(text) {
    const userInput = document.getElementById('user-input-client');
    
    if (userInput && text) {
        // Add text to input, preserving existing content
        const currentText = userInput.value;
        const newText = currentText ? `${currentText} ${text}` : text;
        userInput.value = newText;
        
        // Trigger auto-resize with a small delay to ensure proper alignment
        setTimeout(() => {
            autoResizeTextarea(userInput);
            // Force a style recalculation to maintain proper alignment
            userInput.style.display = 'none';
            userInput.offsetHeight; // Trigger reflow
            userInput.style.display = '';
        }, 10);
        
        console.log('[Voice] Transcribed text added to input:', text);
    }
    
    // Switch back to normal mode
    switchToNormalMode();
    
    // Focus input and move cursor to end with proper timing
    setTimeout(() => {
        if (userInput) {
            userInput.focus();
            userInput.setSelectionRange(userInput.value.length, userInput.value.length);
        }
    }, 50);
}

// Function to initialize voice input functionality
export function initializeVoiceInput() {
    console.log('[Voice] Initializing ChatGPT-style voice input functionality');
    
    // Check for browser support
    if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) {
        console.warn('[Voice] MediaDevices API not supported in this browser');
        // Hide voice button if not supported
        const voiceButton = document.getElementById('voice-input-client');
        if (voiceButton) {
            voiceButton.style.display = 'none';
        }
        return;
    }
    
    if (!window.MediaRecorder) {
        console.warn('[Voice] MediaRecorder API not supported in this browser');
        // Hide voice button if not supported
        const voiceButton = document.getElementById('voice-input-client');
        if (voiceButton) {
            voiceButton.style.display = 'none';
        }
        return;
    }
    
    // Get elements
    const voiceButton = document.getElementById('voice-input-client');
    const cancelVoiceBtn = document.getElementById('cancel-voice-recording');
    const acceptVoiceBtn = document.getElementById('accept-voice-recording');
    
    // Voice button click - start recording immediately
    if (voiceButton) {
        voiceButton.addEventListener('click', () => {
            console.log('[Voice] Voice button clicked - switching to voice mode');
            switchToVoiceMode();
        });
    }
    
    // Cancel recording button
    if (cancelVoiceBtn) {
        cancelVoiceBtn.addEventListener('click', () => {
            console.log('[Voice] Cancel recording clicked');
            switchToNormalMode();
        });
    }
    
    // Accept recording button (checkmark)
    if (acceptVoiceBtn) {
        acceptVoiceBtn.addEventListener('click', () => {
            console.log('[Voice] Accept recording clicked');
            stopVoiceRecording();
        });
    }
    
    // Keyboard shortcuts
    document.addEventListener('keydown', (e) => {
        const voiceMode = document.getElementById('voice-recording-mode');
        const loadingMode = document.getElementById('transcription-loading-mode');
        
        // ESC to cancel recording
        if (e.key === 'Escape') {
            if (voiceMode && voiceMode.style.display !== 'none') {
                switchToNormalMode();
            } else if (loadingMode && loadingMode.style.display !== 'none') {
                switchToNormalMode();
            }
        }
        
        // Enter to accept recording
        if (e.key === 'Enter' && voiceMode && voiceMode.style.display !== 'none') {
            e.preventDefault();
            stopVoiceRecording();
        }
    });
    
    console.log('[Voice] ChatGPT-style voice input functionality initialized successfully');
}

// ========== DEVELOPER MODE VOICE INPUT FUNCTIONALITY ==========

// Developer mode specific variables
let mediaRecorderDev = null;
let audioChunksDev = [];
let recordingTimerDev = null;
let recordingStartTimeDev = null;
let audioContextDev = null;
let analyserDev = null;
let waveformCanvasDev = null;
let waveformAnimationIdDev = null;

// Function to switch to voice recording mode for developer mode
function switchToVoiceModeDev() {
    const normalMode = document.getElementById('normal-input-mode-dev');
    const voiceMode = document.getElementById('voice-recording-mode-dev');
    
    if (normalMode && voiceMode) {
        normalMode.style.display = 'none';
        voiceMode.style.display = 'flex';
        
        // Initialize waveform canvas
        initializeWaveformDev();
        
        // Start recording immediately
        startVoiceRecordingDev();
    }
}

// Function to switch back to normal input mode for developer mode
function switchToNormalModeDev() {
    const normalMode = document.getElementById('normal-input-mode-dev');
    const voiceMode = document.getElementById('voice-recording-mode-dev');
    const loadingMode = document.getElementById('transcription-loading-mode-dev');
    
    if (normalMode && voiceMode && loadingMode) {
        normalMode.style.display = 'flex';
        voiceMode.style.display = 'none';
        loadingMode.style.display = 'none';
        
        // Stop any ongoing recording
        if (mediaRecorderDev && mediaRecorderDev.state === 'recording') {
            stopVoiceRecordingDev();
        }
        
        // Cleanup waveform
        cleanupWaveformDev();
        resetVoiceStateDev();
    }
}

// Function to switch to transcription loading mode for developer mode
function switchToLoadingModeDev() {
    const normalMode = document.getElementById('normal-input-mode-dev');
    const voiceMode = document.getElementById('voice-recording-mode-dev');
    const loadingMode = document.getElementById('transcription-loading-mode-dev');
    
    if (normalMode && voiceMode && loadingMode) {
        normalMode.style.display = 'none';
        voiceMode.style.display = 'none';
        loadingMode.style.display = 'flex';
    }
}

// Function to reset voice state for developer mode
function resetVoiceStateDev() {
    const timer = document.getElementById('voice-recording-timer-dev');
    
    if (timer) {
        timer.textContent = '00:00';
    }
    
    // Clear timer if running
    if (recordingTimerDev) {
        clearInterval(recordingTimerDev);
        recordingTimerDev = null;
    }
    
    // Reset audio data
    audioChunksDev = [];
    recordingStartTimeDev = null;
}

// Function to initialize waveform visualization for developer mode
function initializeWaveformDev() {
    waveformCanvasDev = document.getElementById('voice-waveform-dev');
    if (!waveformCanvasDev) return;
    
    const ctx = waveformCanvasDev.getContext('2d');
    const rect = waveformCanvasDev.getBoundingClientRect();
    
    // Set canvas size to match display size
    waveformCanvasDev.width = rect.width * window.devicePixelRatio;
    waveformCanvasDev.height = rect.height * window.devicePixelRatio;
    ctx.scale(window.devicePixelRatio, window.devicePixelRatio);
    
    // Style the canvas
    waveformCanvasDev.style.width = rect.width + 'px';
    waveformCanvasDev.style.height = rect.height + 'px';
}

// Function to cleanup waveform for developer mode
function cleanupWaveformDev() {
    console.log('[Voice-Dev] Cleaning up waveform visualization');
    
    if (waveformAnimationIdDev) {
        cancelAnimationFrame(waveformAnimationIdDev);
        waveformAnimationIdDev = null;
    }
    
    if (audioContextDev && audioContextDev.state !== 'closed') {
        audioContextDev.close().catch(err => {
            console.warn('[Voice-Dev] Error closing audio context:', err);
        });
        audioContextDev = null;
    }
    
    analyserDev = null;
    
    // Clear the canvas
    if (waveformCanvasDev) {
        const ctx = waveformCanvasDev.getContext('2d');
        const rect = waveformCanvasDev.getBoundingClientRect();
        ctx.clearRect(0, 0, rect.width, rect.height);
    }
}

// Function to draw waveform for developer mode
function drawWaveformDev() {
    if (!waveformCanvasDev || !analyserDev) {
        return;
    }
    
    const ctx = waveformCanvasDev.getContext('2d');
    const rect = waveformCanvasDev.getBoundingClientRect();
    const width = rect.width;
    const height = rect.height;
    
    // Get time domain data for real-time audio visualization
    const bufferLength = analyserDev.fftSize;
    const dataArray = new Uint8Array(bufferLength);
    analyserDev.getByteTimeDomainData(dataArray);
    
    // Clear canvas with transparent background
    ctx.clearRect(0, 0, width, height);
    
    // Calculate how many bars we want to show
    const numBars = Math.min(60, Math.floor(width / 4)); // Adjust bar count based on width
    const barWidth = (width - (numBars - 1)) / numBars; // Account for spacing
    
    // Process audio data to create frequency-style visualization
    const samplesPerBar = Math.floor(bufferLength / numBars);
    
    for (let i = 0; i < numBars; i++) {
        let sum = 0;
        let amplitude = 0;
        
        // Calculate RMS (root mean square) for this bar's audio data
        for (let j = 0; j < samplesPerBar; j++) {
            const sample = (dataArray[i * samplesPerBar + j] - 128) / 128; // Normalize to -1 to 1
            sum += sample * sample;
        }
        
        amplitude = Math.sqrt(sum / samplesPerBar);
        
        // Add some smoothing and make it more responsive
        const minHeight = 3; // Minimum bar height for better visibility
        const maxHeight = height * 0.8;
        
        // Amplify the signal for better visual feedback
        const amplifiedAmplitude = Math.pow(amplitude * 2, 0.8); // Power curve for better visual response
        let barHeight = minHeight + (amplifiedAmplitude * maxHeight);
        
        // Add subtle animation for visual interest
        const timeOffset = Date.now() * 0.005; // Slow time-based animation
        const animationFactor = Math.sin(timeOffset + i * 0.2) * 0.1; // Slight wave effect
        barHeight += animationFactor * 3;
        
        // Add some randomness for visual interest when there's audio
        if (amplitude > 0.005) { // Lower threshold for better responsiveness
            barHeight += Math.random() * 3;
        }
        
        // Ensure minimum visibility and cap maximum
        barHeight = Math.max(barHeight, minHeight);
        barHeight = Math.min(barHeight, maxHeight);
        
        // Calculate x position with spacing
        const x = i * (barWidth + 1);
        
        // Draw the bar centered vertically
        const y = (height - barHeight) / 2;
        
        ctx.fillStyle = '#000000';
        ctx.fillRect(x, y, barWidth, barHeight);
    }
    
    // Continue animation - this ensures continuous updates
    if (mediaRecorderDev && mediaRecorderDev.state === 'recording') {
        waveformAnimationIdDev = requestAnimationFrame(drawWaveformDev);
    }
}

// Function to start audio recording with waveform for developer mode
async function startVoiceRecordingDev() {
    try {
        console.log('[Voice-Dev] Requesting microphone access...');
        
        const stream = await navigator.mediaDevices.getUserMedia({ 
            audio: {
                echoCancellation: true,
                noiseSuppression: true,
                sampleRate: 44100
            } 
        });
        
        console.log('[Voice-Dev] Microphone access granted');
        
        // Setup audio context for waveform visualization
        audioContextDev = new (window.AudioContext || window.webkitAudioContext)();
        analyserDev = audioContextDev.createAnalyser();
        const source = audioContextDev.createMediaStreamSource(stream);
        source.connect(analyserDev);
        
        // Configure analyser for real-time visualization
        analyserDev.fftSize = 1024; // Higher resolution for better visualization
        analyserDev.smoothingTimeConstant = 0.3; // Smooth but responsive
        analyserDev.minDecibels = -90;
        analyserDev.maxDecibels = -10;
        
        // Start waveform animation
        drawWaveformDev();
        
        // Backup timer to ensure continuous animation (in case requestAnimationFrame fails)
        const backupTimer = setInterval(() => {
            if (mediaRecorderDev && mediaRecorderDev.state === 'recording' && !waveformAnimationIdDev) {
                console.log('[Voice-Dev] Restarting waveform animation via backup timer');
                drawWaveformDev();
            } else if (!mediaRecorderDev || mediaRecorderDev.state !== 'recording') {
                clearInterval(backupTimer);
            }
        }, 100); // Check every 100ms
        
        // Create MediaRecorder
        const options = { mimeType: 'audio/webm;codecs=opus' };
        if (!MediaRecorder.isTypeSupported(options.mimeType)) {
            options.mimeType = 'audio/webm';
            if (!MediaRecorder.isTypeSupported(options.mimeType)) {
                options.mimeType = 'audio/mp4';
            }
        }
        
        mediaRecorderDev = new MediaRecorder(stream, options);
        audioChunksDev = [];
        
        mediaRecorderDev.ondataavailable = (event) => {
            if (event.data.size > 0) {
                audioChunksDev.push(event.data);
            }
        };
        
        mediaRecorderDev.onstop = async () => {
            console.log('[Voice-Dev] Recording stopped, processing audio...');
            
            // Stop all tracks to release microphone
            stream.getTracks().forEach(track => track.stop());
            
            // Switch to loading mode
            switchToLoadingModeDev();
            
            if (audioChunksDev.length > 0) {
                await processRecordedAudioDev();
            }
        };
        
        // Start recording
        mediaRecorderDev.start();
        recordingStartTimeDev = Date.now();
        
        console.log('[Voice-Dev] Recording started with continuous waveform visualization');
        
        // Start timer
        startRecordingTimerDev();
        
    } catch (error) {
        console.error('[Voice-Dev] Error accessing microphone:', error);
        alert('Unable to access microphone. Please check your permissions and try again.');
        switchToNormalModeDev();
    }
}

// Function to stop audio recording for developer mode
function stopVoiceRecordingDev() {
    if (mediaRecorderDev && mediaRecorderDev.state === 'recording') {
        console.log('[Voice-Dev] Stopping recording...');
        mediaRecorderDev.stop();
        
        // Clear timer
        if (recordingTimerDev) {
            clearInterval(recordingTimerDev);
            recordingTimerDev = null;
        }
        
        // Stop waveform animation
        cleanupWaveformDev();
    }
}

// Function to start recording timer for developer mode
function startRecordingTimerDev() {
    const timer = document.getElementById('voice-recording-timer-dev');
    if (timer) {
        recordingTimerDev = setInterval(() => {
            if (recordingStartTimeDev) {
                const elapsed = Math.floor((Date.now() - recordingStartTimeDev) / 1000);
                timer.textContent = formatRecordingTime(elapsed);
            }
        }, 1000);
    }
}

// Function to process recorded audio and send to OpenAI for developer mode
async function processRecordedAudioDev() {
    try {
        console.log('[Voice-Dev] Creating audio blob from chunks...');
        
        // Create blob from audio chunks
        const audioBlob = new Blob(audioChunksDev, { type: 'audio/webm' });
        console.log('[Voice-Dev] Audio blob created, size:', audioBlob.size, 'bytes');
        
        if (audioBlob.size === 0) {
            throw new Error('No audio data recorded');
        }
        
        // Send to OpenAI transcription API
        const transcription = await transcribeAudio(audioBlob);
        
        // Use the transcribed text immediately
        useTranscribedTextDev(transcription);
        
    } catch (error) {
        console.error('[Voice-Dev] Error processing audio:', error);
        alert(`Error processing audio: ${error.message}`);
        switchToNormalModeDev();
    }
}

// Function to use transcribed text for developer mode
function useTranscribedTextDev(text) {
    const userInput = document.getElementById('user-input');
    
    if (userInput && text) {
        // Add text to input, preserving existing content
        const currentText = userInput.value;
        const newText = currentText ? `${currentText} ${text}` : text;
        userInput.value = newText;
        
        // Trigger auto-resize with a small delay to ensure proper alignment
        setTimeout(() => {
            autoResizeTextareaDev(userInput);
            // Force a style recalculation to maintain proper alignment
            userInput.style.display = 'none';
            userInput.offsetHeight; // Trigger reflow
            userInput.style.display = '';
        }, 10);
        
        console.log('[Voice-Dev] Transcribed text added to input:', text);
    }
    
    // Switch back to normal mode
    switchToNormalModeDev();
    
    // Focus input and move cursor to end with proper timing
    setTimeout(() => {
        if (userInput) {
            userInput.focus();
            userInput.setSelectionRange(userInput.value.length, userInput.value.length);
        }
    }, 50);
}

// Function to initialize developer mode voice input functionality
export function initializeVoiceInputDev() {
    console.log('[Voice-Dev] Initializing developer mode voice input functionality');
    
    // Check for browser support
    if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) {
        console.warn('[Voice-Dev] MediaDevices API not supported in this browser');
        // Hide voice button if not supported
        const voiceButton = document.getElementById('voice-input-dev');
        if (voiceButton) {
            voiceButton.style.display = 'none';
        }
        return;
    }
    
    if (!window.MediaRecorder) {
        console.warn('[Voice-Dev] MediaRecorder API not supported in this browser');
        // Hide voice button if not supported
        const voiceButton = document.getElementById('voice-input-dev');
        if (voiceButton) {
            voiceButton.style.display = 'none';
        }
        return;
    }
    
    // Get elements
    const voiceButton = document.getElementById('voice-input-dev');
    const cancelVoiceBtn = document.getElementById('cancel-voice-recording-dev');
    const acceptVoiceBtn = document.getElementById('accept-voice-recording-dev');
    
    // Voice button click - start recording immediately
    if (voiceButton) {
        voiceButton.addEventListener('click', () => {
            console.log('[Voice-Dev] Voice button clicked - switching to voice mode');
            switchToVoiceModeDev();
        });
    }
    
    // Cancel recording button
    if (cancelVoiceBtn) {
        cancelVoiceBtn.addEventListener('click', () => {
            console.log('[Voice-Dev] Cancel recording clicked');
            switchToNormalModeDev();
        });
    }
    
    // Accept recording button (checkmark)
    if (acceptVoiceBtn) {
        acceptVoiceBtn.addEventListener('click', () => {
            console.log('[Voice-Dev] Accept recording clicked');
            stopVoiceRecordingDev();
        });
    }
    
    // Keyboard shortcuts
    document.addEventListener('keydown', (e) => {
        const voiceMode = document.getElementById('voice-recording-mode-dev');
        const loadingMode = document.getElementById('transcription-loading-mode-dev');
        
        // ESC to cancel recording
        if (e.key === 'Escape') {
            if (voiceMode && voiceMode.style.display !== 'none') {
                switchToNormalModeDev();
            } else if (loadingMode && loadingMode.style.display !== 'none') {
                switchToNormalModeDev();
            }
        }
        
        // Enter to accept recording
        if (e.key === 'Enter' && voiceMode && voiceMode.style.display !== 'none') {
            e.preventDefault();
            stopVoiceRecordingDev();
        }
    });
    
    console.log('[Voice-Dev] Developer mode voice input functionality initialized successfully');
}

// ========== TEXTAREA AUTO-RESIZE FUNCTIONALITY ==========

// Function to auto-resize textarea based on content
function autoResizeTextarea(textarea) {
    if (!textarea) return;
    
    const baseHeight = 24; // Base single-line height (matches CSS min-height)
    const maxHeight = 120; // Maximum height before scrolling (matches CSS max-height)
    const lineHeight = 24; // Approximate line height based on font size and line-height
    
    // Reset height to base to get accurate scrollHeight
    textarea.style.height = baseHeight + 'px';
    
    // Calculate required height based on content
    const scrollHeight = textarea.scrollHeight;
    const requiredHeight = Math.max(baseHeight, scrollHeight);
    
    // Set height up to maximum, then enable scrolling
    if (requiredHeight <= maxHeight) {
        textarea.style.height = requiredHeight + 'px';
        textarea.classList.remove('scrollable');
    } else {
        textarea.style.height = maxHeight + 'px';
        textarea.classList.add('scrollable');
    }
    
    console.log(`[TextArea] Auto-resize: required=${requiredHeight}px, set=${textarea.style.height}, scrollable=${textarea.classList.contains('scrollable')}`);
}

// Function to auto-resize textarea for developer mode
function autoResizeTextareaDev(textarea) {
    if (!textarea) return;
    
    const baseHeight = 24; // Base single-line height (matches CSS min-height)
    const maxHeight = 120; // Maximum height before scrolling (matches CSS max-height)
    
    // Reset height to base to get accurate scrollHeight
    textarea.style.height = baseHeight + 'px';
    
    // Calculate required height based on content
    const scrollHeight = textarea.scrollHeight;
    const requiredHeight = Math.max(baseHeight, scrollHeight);
    
    // Set height up to maximum, then enable scrolling
    if (requiredHeight <= maxHeight) {
        textarea.style.height = requiredHeight + 'px';
        textarea.classList.remove('scrollable');
    } else {
        textarea.style.height = maxHeight + 'px';
        textarea.classList.add('scrollable');
    }
    
    console.log(`[TextArea-Dev] Auto-resize: required=${requiredHeight}px, set=${textarea.style.height}, scrollable=${textarea.classList.contains('scrollable')}`);
}

// Function to initialize textarea auto-resize functionality for client mode
export function initializeTextareaAutoResize() {
    console.log('[TextArea] Initializing client mode auto-resize functionality');
    
    const textarea = document.getElementById('user-input-client');
    if (!textarea) {
        console.warn('[TextArea] Client textarea not found');
        return;
    }
    
    // Auto-resize on input
    textarea.addEventListener('input', () => {
        autoResizeTextarea(textarea);
    });
    
    // Auto-resize on paste
    textarea.addEventListener('paste', () => {
        // Use setTimeout to wait for paste content to be inserted
        setTimeout(() => {
            autoResizeTextarea(textarea);
        }, 10);
    });
    
    // Auto-resize when text is set programmatically (like from voice input)
    const originalValueDescriptor = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value');
    Object.defineProperty(textarea, 'value', {
        set: function(newValue) {
            originalValueDescriptor.set.call(this, newValue);
            autoResizeTextarea(this);
        },
        get: function() {
            return originalValueDescriptor.get.call(this);
        },
        configurable: true
    });
    
    // Initial resize to set proper height
    autoResizeTextarea(textarea);
    
    console.log('[TextArea] Client mode auto-resize functionality initialized successfully');
}

// Function to initialize textarea auto-resize functionality for developer mode
export function initializeTextareaAutoResizeDev() {
    console.log('[TextArea-Dev] Initializing developer mode auto-resize functionality');
    
    const textarea = document.getElementById('user-input');
    if (!textarea) {
        console.warn('[TextArea-Dev] Developer textarea not found');
        return;
    }
    
    // Auto-resize on input
    textarea.addEventListener('input', () => {
        autoResizeTextareaDev(textarea);
    });
    
    // Auto-resize on paste
    textarea.addEventListener('paste', () => {
        // Use setTimeout to wait for paste content to be inserted
        setTimeout(() => {
            autoResizeTextareaDev(textarea);
        }, 10);
    });
    
    // Initial resize to set proper height
    autoResizeTextareaDev(textarea);
    
    console.log('[TextArea-Dev] Developer mode auto-resize functionality initialized successfully');
}


