# Developer API Reference

## Overview

This document provides comprehensive API documentation for developers working with the Projectify5.0 system. It covers all major modules, functions, and integration points.

## Table of Contents

1. [Core Modules](#core-modules)
2. [AI Integration](#ai-integration)
3. [Excel Integration](#excel-integration)
4. [Training Data System](#training-data-system)
5. [Validation System](#validation-system)
6. [Authentication](#authentication)
7. [Configuration](#configuration)
8. [Error Handling](#error-handling)
9. [Development Tools](#development-tools)

## Core Modules

### taskpane.js

The main application controller that orchestrates all system components.

#### Key Functions

##### `initializeAPIKeys()`
Initializes API keys from environment variables.

```javascript
export async function initializeAPIKeys()
```

**Returns:**
- `Promise<Object>`: Object containing API keys
  - `OPENAI_API_KEY`: OpenAI API key
  - `PINECONE_API_KEY`: Pinecone API key

**Example:**
```javascript
const keys = await initializeAPIKeys();
console.log("OpenAI key available:", !!keys.OPENAI_API_KEY);
```

##### `loadCodeDatabase()`
Loads the codestring database from assets.

```javascript
async function loadCodeDatabase()
```

**Returns:**
- `Promise<void>`: Resolves when database is loaded

**Side Effects:**
- Populates global `codeDatabase` array
- Logs loading progress and errors

##### `handleSend()`
Processes user input in developer mode.

```javascript
async function handleSend()
```

**Process:**
1. Validates user input
2. Calls AI processing pipeline
3. Updates UI with results
4. Handles errors gracefully

##### `handleSendClient()`
Processes user input in client mode.

```javascript
async function handleSendClient()
```

**Process:**
1. Simplified input validation
2. Client-focused AI processing
3. Streamlined UI updates

##### `processModelCodesForPlanner(modelCodesString)`
Processes model codes for the AI planner.

```javascript
export async function processModelCodesForPlanner(modelCodesString)
```

**Parameters:**
- `modelCodesString` (string): Raw model codes to process

**Returns:**
- `Promise<Object>`: Processed model structure
  - `tabs`: Array of tab objects
  - `metadata`: Processing metadata
  - `errors`: Array of processing errors

**Example:**
```javascript
const codes = `<CONST-E; row1="V1(D)|Revenue(L)|...">`;
const result = await processModelCodesForPlanner(codes);
console.log("Generated tabs:", result.tabs);
```

### CodeCollection.js

Manages codestring execution and Excel integration.

#### Key Functions

##### `populateCodeCollection(codes)`
Populates the code collection for execution.

```javascript
export function populateCodeCollection(codes)
```

**Parameters:**
- `codes` (Array): Array of codestring objects

**Returns:**
- `Object`: Populated code collection ready for execution

##### `runCodes(codeStrings)`
Executes codestrings in Excel.

```javascript
export async function runCodes(codeStrings)
```

**Parameters:**
- `codeStrings` (Array): Validated codestring objects

**Returns:**
- `Promise<Object>`: Execution results
  - `success`: Boolean indicating success
  - `errors`: Array of execution errors
  - `results`: Execution results data

##### `processAssumptionTabs()`
Processes assumption tabs for financial statements.

```javascript
export async function processAssumptionTabs()
```

**Returns:**
- `Promise<Object>`: Processing results
  - `processed`: Number of tabs processed
  - `errors`: Array of processing errors

##### `exportCodeCollectionToText()`
Exports code collection to text format.

```javascript
export function exportCodeCollectionToText()
```

**Returns:**
- `string`: Text representation of code collection

## AI Integration

### AIcalls.js

Handles all AI-related functionality including OpenAI and Pinecone integration.

#### Key Functions

##### `setAPIKeys(keys)`
Sets API keys for AI services.

```javascript
export function setAPIKeys(keys)
```

**Parameters:**
- `keys` (Object): API key configuration
  - `OPENAI_API_KEY`: OpenAI API key
  - `PINECONE_API_KEY`: Pinecone API key

##### `callOpenAI(prompt, systemPrompt, options)`
Makes calls to OpenAI API.

```javascript
export async function callOpenAI(prompt, systemPrompt, options = {})
```

**Parameters:**
- `prompt` (string): User prompt
- `systemPrompt` (string): System prompt
- `options` (Object): Optional configuration
  - `model`: Model name (default: "gpt-4")
  - `temperature`: Response creativity (default: 0.7)
  - `max_tokens`: Maximum response length

**Returns:**
- `Promise<Object>`: OpenAI response
  - `choices`: Array of response choices
  - `usage`: Token usage information

##### `createEmbedding(text)`
Creates embeddings for vector search.

```javascript
export async function createEmbedding(text)
```

**Parameters:**
- `text` (string): Text to embed

**Returns:**
- `Promise<Array>`: Vector embedding array

##### `queryVectorDB(embedding, topK)`
Queries Pinecone vector database.

```javascript
export async function queryVectorDB(embedding, topK = 10)
```

**Parameters:**
- `embedding` (Array): Query vector
- `topK` (number): Number of results to return

**Returns:**
- `Promise<Array>`: Array of similar vectors with metadata

##### `handleConversation(userInput, conversationHistory)`
Handles conversation flow with AI.

```javascript
export async function handleConversation(userInput, conversationHistory)
```

**Parameters:**
- `userInput` (string): User's message
- `conversationHistory` (Array): Previous conversation messages

**Returns:**
- `Promise<Object>`: Conversation result
  - `response`: AI response text
  - `codestrings`: Generated codestrings (if any)
  - `metadata`: Response metadata

##### `getAICallsProcessedResponse(userInput, conversationHistory)`
Gets processed AI response with structured data.

```javascript
export async function getAICallsProcessedResponse(userInput, conversationHistory)
```

**Parameters:**
- `userInput` (string): User input
- `conversationHistory` (Array): Conversation history

**Returns:**
- `Promise<Object>`: Processed response
  - `response`: Formatted response text
  - `codestrings`: Parsed codestrings
  - `validation`: Validation results
  - `suggestions`: Improvement suggestions

### AIModelPlanner.js

Manages AI model planning workflow.

#### Key Functions

##### `setAIModelPlannerOpenApiKey(apiKey)`
Sets OpenAI API key for model planner.

```javascript
export function setAIModelPlannerOpenApiKey(apiKey)
```

**Parameters:**
- `apiKey` (string): OpenAI API key

##### `initializeFileAttachment()`
Initializes file attachment functionality.

```javascript
export function initializeFileAttachment()
```

**Sets up:**
- File input handling
- File type validation
- Progress tracking

##### `initializeVoiceInput()`
Initializes voice input for client mode.

```javascript
export function initializeVoiceInput()
```

**Features:**
- Speech recognition
- Voice commands
- Audio processing

##### `initializeVoiceInputDev()`
Initializes voice input for developer mode.

```javascript
export function initializeVoiceInputDev()
```

**Features:**
- Advanced voice commands
- Code dictation
- Developer-specific voice features

## Excel Integration

### Functions.js

Custom Excel functions for financial calculations.

#### Custom Functions

##### `ADD(first, second)`
Adds two numbers.

```javascript
/**
 * Adds two numbers
 * @customfunction
 * @param {number} first First number
 * @param {number} second Second number
 * @returns {number} Sum of the two numbers
 */
function ADD(first, second)
```

##### `PROJECTIFY_CALCULATE(formula, parameters)`
Executes Projectify calculations.

```javascript
/**
 * Executes Projectify calculations
 * @customfunction
 * @param {string} formula Formula to execute
 * @param {any} parameters Calculation parameters
 * @returns {any} Calculation result
 */
function PROJECTIFY_CALCULATE(formula, parameters)
```

### Commands.js

Ribbon command handlers.

#### Key Functions

##### `showTaskpane()`
Shows the main taskpane.

```javascript
function showTaskpane(event)
```

**Parameters:**
- `event` (Object): Office.js event object

##### `insertModel()`
Inserts a model template.

```javascript
function insertModel(event)
```

**Parameters:**
- `event` (Object): Office.js event object

##### `validateModel()`
Validates current model.

```javascript
function validateModel(event)
```

**Parameters:**
- `event` (Object): Office.js event object

## Training Data System

### Training Data Queue

#### Key Functions

##### `addToTrainingDataQueue()`
Adds entry to training data queue.

```javascript
async function addToTrainingDataQueue()
```

**Process:**
1. Captures current prompt
2. Captures selected code
3. Saves to localStorage
4. Downloads CSV file

##### `showTrainingDataModal(userPrompt, aiResponse)`
Shows training data modal for editing.

```javascript
function showTrainingDataModal(userPrompt = '', aiResponse = '')
```

**Parameters:**
- `userPrompt` (string): User's prompt
- `aiResponse` (string): AI's response

##### `saveTrainingDataFromModal()`
Saves training data from modal.

```javascript
async function saveTrainingDataFromModal()
```

**Process:**
1. Validates modal input
2. Saves to training queue
3. Downloads updated CSV
4. Closes modal

##### `clearTrainingDataQueue()`
Clears the training data queue.

```javascript
function clearTrainingDataQueue()
```

**Effects:**
- Clears localStorage
- Resets UI indicators
- Logs clearing action

## Validation System

### Validation.js

Comprehensive validation for codestrings and models.

#### Key Functions

##### `validateCodeStringsForRun(codeStrings)`
Validates codestrings before execution.

```javascript
export function validateCodeStringsForRun(codeStrings)
```

**Parameters:**
- `codeStrings` (Array): Array of codestring objects

**Returns:**
- `Object`: Validation results
  - `valid`: Boolean indicating overall validity
  - `errors`: Array of error objects
  - `warnings`: Array of warning objects
  - `suggestions`: Array of improvement suggestions

**Validation Checks:**
- Syntax validation
- Driver reference validation
- Column structure validation
- Financial logic validation

##### `validateSingleCodeString(codeString)`
Validates individual codestring.

```javascript
export function validateSingleCodeString(codeString)
```

**Parameters:**
- `codeString` (Object): Single codestring object

**Returns:**
- `Object`: Validation result for single codestring

##### `validateDriverReferences(codeStrings)`
Validates driver references across codestrings.

```javascript
export function validateDriverReferences(codeStrings)
```

**Parameters:**
- `codeStrings` (Array): Array of codestring objects

**Returns:**
- `Object`: Driver validation results
  - `valid`: Boolean
  - `missingDrivers`: Array of missing driver references
  - `duplicateDrivers`: Array of duplicate driver codes

### PipeValidation.js

Specialized validation for pipe structure.

#### Key Functions

##### `autoCorrectPipeCounts(inputText)`
Automatically corrects pipe counts in codestrings.

```javascript
export function autoCorrectPipeCounts(inputText)
```

**Parameters:**
- `inputText` (string|Array): Input containing codestrings

**Returns:**
- `Object`: Correction results
  - `correctedText`: Corrected input
  - `changesMade`: Number of corrections made
  - `details`: Array of correction details

##### `validateWithPipeCorrection(inputText)`
Validates input with automatic pipe correction.

```javascript
export function validateWithPipeCorrection(inputText)
```

**Parameters:**
- `inputText` (string|Array): Input to validate and correct

**Returns:**
- `Object`: Combined validation and correction results

## Authentication

### Microsoft Azure AD Integration

#### Key Functions

##### `initializeMSAL()`
Initializes Microsoft Authentication Library.

```javascript
function initializeMSAL()
```

**Configuration:**
- Client ID setup
- Redirect URI configuration
- Scope definitions

##### `signIn()`
Signs in user with Microsoft account.

```javascript
async function signIn()
```

**Returns:**
- `Promise<Object>`: Authentication result
  - `accessToken`: Access token
  - `account`: User account information

##### `signOut()`
Signs out current user.

```javascript
async function signOut()
```

**Effects:**
- Clears authentication tokens
- Resets user interface
- Redirects to login

##### `getUserProfile()`
Gets current user profile.

```javascript
async function getUserProfile()
```

**Returns:**
- `Promise<Object>`: User profile data
  - `displayName`: User's display name
  - `mail`: User's email
  - `jobTitle`: User's job title

##### `isUserAuthenticated()`
Checks if user is authenticated.

```javascript
function isUserAuthenticated()
```

**Returns:**
- `boolean`: True if user is authenticated

## Configuration

### config.js

Central configuration management.

#### CONFIG Object

##### `isDevelopment`
Indicates if running in development mode.

```javascript
const isDevelopment = CONFIG.isDevelopment;
```

##### `getAssetUrl(path)`
Gets URL for asset files.

```javascript
const imageUrl = CONFIG.getAssetUrl('assets/logo.png');
```

**Parameters:**
- `path` (string): Asset path

**Returns:**
- `string`: Complete asset URL

##### `API_ENDPOINTS`
API endpoint configurations.

```javascript
const endpoints = CONFIG.API_ENDPOINTS;
// endpoints.OPENAI_API
// endpoints.PINECONE_API
```

## Error Handling

### Error Types

#### `ValidationError`
Thrown when validation fails.

```javascript
class ValidationError extends Error {
  constructor(message, details = {}) {
    super(message);
    this.name = 'ValidationError';
    this.details = details;
  }
}
```

#### `APIError`
Thrown when API calls fail.

```javascript
class APIError extends Error {
  constructor(message, statusCode, response) {
    super(message);
    this.name = 'APIError';
    this.statusCode = statusCode;
    this.response = response;
  }
}
```

#### `ExcelError`
Thrown when Excel operations fail.

```javascript
class ExcelError extends Error {
  constructor(message, context) {
    super(message);
    this.name = 'ExcelError';
    this.context = context;
  }
}
```

### Error Handling Patterns

#### Try-Catch with Context

```javascript
try {
  const result = await someOperation();
  return result;
} catch (error) {
  console.error('Operation failed:', error);
  
  // Add context to error
  const contextualError = new Error(`Operation failed in ${context}: ${error.message}`);
  contextualError.originalError = error;
  
  throw contextualError;
}
```

#### Graceful Degradation

```javascript
async function robustOperation() {
  try {
    return await primaryMethod();
  } catch (error) {
    console.warn('Primary method failed, trying fallback:', error);
    
    try {
      return await fallbackMethod();
    } catch (fallbackError) {
      console.error('Both methods failed:', { error, fallbackError });
      return defaultValue;
    }
  }
}
```

## Development Tools

### Debug Mode

#### Enable Debug Mode

```javascript
const DEBUG = true; // Set in taskpane.js
```

#### Debug Logging

```javascript
function debugLog(message, data = {}) {
  if (DEBUG) {
    console.log(`[DEBUG] ${message}`, data);
  }
}
```

#### Performance Monitoring

```javascript
function performanceMonitor(operationName) {
  const start = performance.now();
  
  return {
    end: () => {
      const duration = performance.now() - start;
      debugLog(`${operationName} took ${duration.toFixed(2)}ms`);
    }
  };
}

// Usage
const monitor = performanceMonitor('AI Processing');
await processWithAI();
monitor.end();
```

### Testing Utilities

#### Mock API Responses

```javascript
const mockAIResponse = {
  choices: [{
    message: {
      content: 'Mock AI response'
    }
  }],
  usage: {
    prompt_tokens: 10,
    completion_tokens: 20
  }
};

// Use in tests
if (CONFIG.isTest) {
  return mockAIResponse;
}
```

#### Validation Testing

```javascript
function testValidation() {
  const testCases = [
    {
      name: 'Valid codestring',
      input: '<CONST-E; row1="V1(D)|Test(L)|(F)|(C1)|(C2)|(C3)|(C4)|(C5)|(C6)|(Y1)|(Y2)|(Y3)|(Y4)|(Y5)|(Y6)|">',
      expected: { valid: true, errors: [] }
    },
    {
      name: 'Invalid pipe count',
      input: '<CONST-E; row1="V1(D)|Test(L)|(F)|(C1)|(C2)|(C3)|(C4)|(C5)|(C6)|(Y1)|(Y2)|(Y3)|(Y4)|(Y5)|">',
      expected: { valid: false, errors: [{ type: 'PIPE_COUNT' }] }
    }
  ];

  testCases.forEach(testCase => {
    const result = validateSingleCodeString(testCase.input);
    console.assert(
      result.valid === testCase.expected.valid,
      `Test failed: ${testCase.name}`
    );
  });
}
```

## Integration Examples

### Complete Model Generation Flow

```javascript
async function generateCompleteModel(userPrompt) {
  try {
    // 1. Initialize system
    const keys = await initializeAPIKeys();
    setAPIKeys(keys);
    
    // 2. Process user input
    const conversationHistory = [];
    const aiResponse = await handleConversation(userPrompt, conversationHistory);
    
    // 3. Validate generated codestrings
    const validationResult = validateCodeStringsForRun(aiResponse.codestrings);
    
    if (!validationResult.valid) {
      throw new ValidationError('Invalid codestrings generated', validationResult);
    }
    
    // 4. Execute in Excel
    const executionResult = await runCodes(aiResponse.codestrings);
    
    // 5. Process financial statements
    await processAssumptionTabs();
    
    // 6. Return success
    return {
      success: true,
      model: executionResult,
      metadata: {
        prompt: userPrompt,
        codestrings: aiResponse.codestrings.length,
        validation: validationResult
      }
    };
    
  } catch (error) {
    console.error('Model generation failed:', error);
    
    return {
      success: false,
      error: error.message,
      details: error.details || {}
    };
  }
}
```

### Custom Integration

```javascript
// Example: Custom financial calculation
async function customFinancialCalculation(parameters) {
  // Build custom codestring
  const codestring = `<FORMULA-S; customformula="${parameters.formula}"; row1="V1(D)|${parameters.label}(L)|${parameters.fincode}(F)|(C1)|(C2)|(C3)|(C4)|(C5)|(C6)|$F(Y1)|$F(Y2)|$F(Y3)|$F(Y4)|$F(Y5)|$F(Y6)|">`;
  
  // Validate
  const validation = validateSingleCodeString(codestring);
  if (!validation.valid) {
    throw new ValidationError('Invalid custom calculation', validation);
  }
  
  // Execute
  return await runCodes([codestring]);
}
```

This API reference provides comprehensive coverage of the Projectify5.0 system for developers. For additional details on specific functions, consult the inline documentation in the source code. 