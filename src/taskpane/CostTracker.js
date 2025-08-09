/**
 * API Cost Tracking System
 * Console-based cost tracking for OpenAI and Claude API calls with backend credit deduction
 */

import backendAPI from './BackendAPI.js';
import { refreshUserData } from './UserProfile.js';

// ============================================================================
// PRICING CONSTANTS (Updated as of January 2025)
// ============================================================================

const API_PRICING = {
  // OpenAI Models (per 1M tokens)
  OPENAI: {
    'gpt-4o': { input: 2.50, output: 10.00 },
    'gpt-4o-mini': { input: 0.15, output: 0.60 },
    'gpt-4.1': { input: 2.00, output: 8.00 },
    'gpt-4.1-mini': { input: 0.40, output: 1.60 },
    'gpt-4.1-nano': { input: 0.10, output: 0.40 },
    'gpt-4-turbo': { input: 10.00, output: 30.00 },
    'gpt-4': { input: 30.00, output: 60.00 },
    'gpt-3.5-turbo': { input: 0.50, output: 1.50 },
    'o3': { input: 1.00, output: 4.00, cached_input: 0.25 },
    'o1': { input: 15.00, output: 60.00 },
    'o1-mini': { input: 3.00, output: 12.00 },
    'o1-preview': { input: 15.00, output: 60.00 },
    'gpt-o3': { input: 1.00, output: 4.00 }, // Alternative naming
    'o3-mini': { input: 0.30, output: 1.20 }
  },
  
  // Claude Models (per 1M tokens) - Updated January 2025
  CLAUDE: {
    'claude-4-opus': { input: 15.00, output: 75.00 },
    'claude-4.1-opus': { input: 15.00, output: 75.00 },
    'claude-4-sonnet': { input: 3.00, output: 15.00 },
    'claude-sonnet-4-20250514': { input: 3.00, output: 15.00 }, // Current model used in your codebase
    'claude-3.7-sonnet': { input: 3.00, output: 15.00 },
    'claude-3.5-sonnet': { input: 3.00, output: 15.00 },
    'claude-3.5-haiku': { input: 0.80, output: 4.00 },
    'claude-3-opus': { input: 15.00, output: 75.00 },
    'claude-3-sonnet': { input: 3.00, output: 15.00 },
    'claude-3-haiku': { input: 0.25, output: 1.25 }
  }
};

// ============================================================================
// SIMPLE COST TRACKING
// ============================================================================

let totalSessionCost = 0;
let totalSessionCalls = 0;

/**
 * Calculate cost for an API call
 */
function calculateCost(provider, model, inputTokens, outputTokens) {
  const providerKey = provider.toUpperCase();
  const pricing = API_PRICING[providerKey];
  
  if (!pricing || !pricing[model]) {
    return {
      inputCost: 0,
      outputCost: 0,
      totalCost: 0,
      error: `Unknown model: ${provider}:${model}`
    };
  }

  const modelPricing = pricing[model];
  
  // Calculate costs (convert per million tokens to actual cost)
  const inputCost = (inputTokens / 1000000) * modelPricing.input;
  const outputCost = (outputTokens / 1000000) * modelPricing.output;
  const totalCost = inputCost + outputCost;

  return {
    inputCost,
    outputCost,
    totalCost
  };
}

/**
 * Track and log an API call cost
 */
function trackAPICallCost(callData) {
  const {
    provider,
    model,
    inputTokens,
    outputTokens,
    caller,
    duration = 0,
    success = true
  } = callData;

  const costInfo = calculateCost(provider, model, inputTokens, outputTokens);
  
  if (costInfo.error) {
    console.warn(`💰 Cost Tracking Warning: ${costInfo.error}`);
    return;
  }

  // Update session totals
  if (success) {
    totalSessionCost += costInfo.totalCost;
    totalSessionCalls++;
  }

  // Log cost information to console
  const statusIcon = success ? '✅' : '❌';
  const providerDisplay = provider.toUpperCase();
  
  console.log(`\n💰 ═════════════════ API COST TRACKING ═════════════════`);
  console.log(`${statusIcon} ${providerDisplay} API Call - ${caller || 'Unknown'}`);
  console.log(`🤖 Model: ${model}`);
  console.log(`📊 Tokens: ${inputTokens.toLocaleString()} input + ${outputTokens.toLocaleString()} output = ${(inputTokens + outputTokens).toLocaleString()} total`);
  console.log(`💵 Cost Breakdown:`);
  console.log(`   • Input:  $${costInfo.inputCost.toFixed(6)} (${inputTokens.toLocaleString()} × $${(API_PRICING[provider.toUpperCase()][model].input / 1000000).toFixed(8)}/token)`);
  console.log(`   • Output: $${costInfo.outputCost.toFixed(6)} (${outputTokens.toLocaleString()} × $${(API_PRICING[provider.toUpperCase()][model].output / 1000000).toFixed(8)}/token)`);
  console.log(`   • Total:  $${costInfo.totalCost.toFixed(6)}`);
  
  if (duration > 0) {
    console.log(`⏱️ Duration: ${(duration / 1000).toFixed(2)}s`);
  }
  
  console.log(`📈 Session Summary: ${totalSessionCalls} calls, $${totalSessionCost.toFixed(6)} total`);
  
  // Deduct credits from backend (async, don't block)
  if (success && costInfo.totalCost > 0) {
    deductCreditsAsync({
      cost: costInfo.totalCost,
      provider,
      model,
      tokens: { input: inputTokens, output: outputTokens },
      caller
    });
  }
  
  console.log(`═══════════════════════════════════════════════════════\n`);

  return costInfo;
}

/**
 * Deduct credits asynchronously (don't block the UI)
 */
async function deductCreditsAsync(costData) {
  try {
    const creditCost = costData.cost / 0.05; // Convert to credits
    console.log(`💳 Deducting ${creditCost.toFixed(4)} credits for $${costData.cost.toFixed(6)} API call...`);
    
    try {
      // Try the new precise credit deduction endpoint
      console.log(`🔍 Calling deductCredits with:`, costData);
      const result = await backendAPI.deductCredits(costData);
      console.log(`🔍 Backend response:`, result);
      
      if (result.success) {
        console.log(`✅ Credits deducted successfully: ${result.creditCost.toFixed(4)} credits. Remaining: ${result.remainingCredits}`);
        
        // Refresh user data to update UI
        setTimeout(() => {
          refreshUserData().catch(err => 
            console.warn('Failed to refresh user data after credit deduction:', err)
          );
        }, 500);
        return;
      }
    } catch (endpointError) {
      console.warn(`⚠️ New credit deduction endpoint not available, falling back to legacy method...`);
      
      // Fallback to legacy single-credit deduction if cost is significant enough
      if (costData.cost >= 0.025) { // If cost is at least half a credit
        try {
          console.log(`💳 Using legacy credit deduction (1 credit) for $${costData.cost.toFixed(6)} call`);
          const fallbackResult = await backendAPI.useCredit('api_call');
          
          if (fallbackResult.success) {
            console.log(`✅ Legacy credit deducted. Remaining: ${fallbackResult.remainingCredits}`);
            
            // Refresh user data to update UI
            setTimeout(() => {
              refreshUserData().catch(err => 
                console.warn('Failed to refresh user data after credit deduction:', err)
              );
            }, 500);
          }
        } catch (fallbackError) {
          console.error(`❌ Legacy credit deduction also failed:`, fallbackError);
        }
      } else {
        console.log(`💰 API call cost ($${costData.cost.toFixed(6)}) too small for legacy 1-credit deduction, skipping...`);
      }
    }
  } catch (error) {
    console.error('❌ Error deducting credits:', error);
    // Don't throw - we don't want to break the API call flow
  }
}

/**
 * Estimate tokens in text (rough approximation)
 */
function estimateTokens(text) {
  if (!text || typeof text !== 'string') return 0;
  
  // Rough estimation: 1 token ≈ 0.75 words ≈ 4 characters
  // This is approximate and varies by language and content
  const words = text.trim().split(/\s+/).length;
  const chars = text.length;
  
  // Use the higher of word-based or character-based estimate
  const wordBasedTokens = Math.ceil(words / 0.75);
  const charBasedTokens = Math.ceil(chars / 4);
  
  return Math.max(wordBasedTokens, charBasedTokens);
}

/**
 * Get session cost summary
 */
function getSessionSummary() {
  return {
    totalCalls: totalSessionCalls,
    totalCost: totalSessionCost
  };
}

/**
 * Reset session tracking
 */
function resetSessionTracking() {
  totalSessionCost = 0;
  totalSessionCalls = 0;
  console.log('💰 Session cost tracking reset');
}

// ============================================================================
// EXPORTS
// ============================================================================

export {
  trackAPICallCost,
  estimateTokens,
  calculateCost,
  getSessionSummary,
  resetSessionTracking,
  API_PRICING
}; 