/**
 * Credit System Management
 * Handles credit consumption, checking, and feature gating
 */

import backendAPI, { InsufficientCreditsError } from './BackendAPI.js';
import { getUserCredits, refreshUserData } from './UserProfile.js';

class CreditSystemManager {
  constructor() {
    this.creditsCache = null;
    this.lastCreditCheck = null;
    this.checkInterval = 30000; // 30 seconds
  }

  // ============================================================================
  // CREDIT CHECKING
  // ============================================================================

  /**
   * Check if user has sufficient credits for an action
   */
  async checkCreditsForAction(action) {
    try {
      // Development mode bypass - if no authentication is set up, allow unlimited usage
      if (typeof window !== 'undefined' && !window.userProfileManager) {
        console.log("🔧 Development mode: No user profile manager - allowing unlimited usage");
        return { canProceed: true, reason: 'development' };
      }

      const userCredits = getUserCredits();
      const hasSubscription = this.hasActiveSubscription();

      // If user has active subscription, allow unlimited usage
      if (hasSubscription) {
        console.log("✅ User has active subscription - unlimited usage");
        return { canProceed: true, reason: 'subscription' };
      }

      // Check credit balance
      if (userCredits > 0) {
        console.log(`✅ User has ${userCredits} credits available`);
        return { canProceed: true, reason: 'credits', remaining: userCredits };
      }

      // If user profile manager exists but userData is not loaded, allow usage temporarily
      if (window.userProfileManager && !window.userProfileManager.getUserData()) {
        console.log("🔧 User data not loaded - allowing usage temporarily");
        return { canProceed: true, reason: 'user_data_not_loaded' };
      }

      console.log("❌ Insufficient credits and no active subscription");
      return { 
        canProceed: false, 
        reason: 'insufficient_credits', 
        remaining: userCredits 
      };

    } catch (error) {
      console.error("❌ Error checking credits:", error);
      // In case of error, allow usage to prevent blocking legitimate use
      console.log("🔧 Credit check failed - allowing usage to prevent blocking");
      return { canProceed: true, reason: 'error_bypass' };
    }
  }

  /**
   * Check if user has active subscription
   */
  hasActiveSubscription() {
    // Access userProfileManager through global or import
    if (typeof window !== 'undefined' && window.userProfileManager) {
      return window.userProfileManager.hasActiveSubscription();
    }
    return false;
  }

  // ============================================================================
  // CREDIT CONSUMPTION
  // ============================================================================

  /**
   * Use a credit for model building
   */
  async useCreditForBuild() {
    return await this.useCredit('build');
  }

  /**
   * Use a credit for model updating
   */
  async useCreditForUpdate() {
    return await this.useCredit('update');
  }

  /**
   * Use a credit for a specific action
   */
  async useCredit(action) {
    try {
      console.log(`💳 Attempting to use credit for: ${action}`);

      // First check if user can proceed
      const creditCheck = await this.checkCreditsForAction(action);
      if (!creditCheck.canProceed) {
        if (creditCheck.reason === 'insufficient_credits') {
          this.showInsufficientCreditsModal(creditCheck.remaining);
          throw new InsufficientCreditsError(creditCheck.remaining);
        } else {
          throw new Error(creditCheck.error || 'Cannot use credit');
        }
      }

      // If subscription, don't actually consume credits
      if (creditCheck.reason === 'subscription') {
        console.log("✅ Action allowed via subscription (no credit consumed)");
        return {
          success: true,
          message: 'Action completed (subscription)',
          remainingCredits: null,
          viaSubscription: true
        };
      }

      // If development mode or user data not loaded, skip backend call
      if (creditCheck.reason === 'development' || creditCheck.reason === 'user_data_not_loaded' || creditCheck.reason === 'error_bypass') {
        console.log(`🔧 ${creditCheck.reason} - skipping backend credit consumption`);
        return {
          success: true,
          message: `Action completed (${creditCheck.reason})`,
          remainingCredits: null,
          viaDevelopmentMode: true
        };
      }

      // Consume credit via backend
      const result = await backendAPI.useCredit(action);
      
      if (result.success) {
        console.log(`✅ Credit consumed. Remaining: ${result.remainingCredits}`);
        
        // Update local credit display
        await refreshUserData();
        
        return result;
      } else {
        throw new Error(result.error || 'Failed to use credit');
      }

    } catch (error) {
      console.error(`❌ Failed to use credit for ${action}:`, error);
      
      if (error instanceof InsufficientCreditsError) {
        this.showInsufficientCreditsModal(error.credits);
      } else {
        this.showCreditError(error.message);
      }
      
      throw error;
    }
  }

  // ============================================================================
  // FEATURE GATING
  // ============================================================================

  /**
   * Check if user can use a feature and gate access
   */
  async enforceFeatureAccess(featureType, callback) {
    try {
      const creditCheck = await this.checkCreditsForAction(featureType);
      
      if (creditCheck.canProceed) {
        // Execute the callback if access is allowed
        if (typeof callback === 'function') {
          return await callback();
        }
        return true;
      } else {
        // Show subscription prompt
        this.showSubscriptionPrompt(creditCheck.remaining);
        return false;
      }
    } catch (error) {
      console.error("❌ Feature access check failed:", error);
      if (typeof showError === 'function') {
        showError("Unable to verify feature access. Please try again.");
      }
      return false;
    }
  }

  /**
   * Check if user can use features (utility function)
   */
  canUseFeatures() {
    const userCredits = getUserCredits();
    const hasSubscription = this.hasActiveSubscription();
    return hasSubscription || userCredits > 0;
  }

  // ============================================================================
  // UI INTERACTIONS
  // ============================================================================

  /**
   * Show insufficient credits modal
   */
  showInsufficientCreditsModal(credits) {
    console.log("🚫 Showing insufficient credits modal");

    // Show modal or notification
    const message = credits === 0 
      ? "You've run out of credits! Upgrade to Pro for unlimited usage."
      : `You only have ${credits} credit${credits === 1 ? '' : 's'} left. Upgrade to Pro for unlimited usage.`;

    if (typeof showError === 'function') {
      showError(message);
    }

    // Show upgrade prompt
    this.showSubscriptionPrompt(credits);
  }

  /**
   * Show subscription prompt
   */
  showSubscriptionPrompt(credits = 0) {
    console.log("💰 Showing subscription prompt");

    // Update low credits warning
    const warningElement = document.getElementById('low-credits-warning');
    if (warningElement) {
      warningElement.style.display = 'block';
    }

    // Show upgrade buttons
    const upgradeButtons = document.querySelectorAll('.upgrade-button');
    upgradeButtons.forEach(button => {
      button.style.display = 'inline-block';
      button.onclick = () => this.startSubscription();
    });

    // Show message
    const message = credits === 0 
      ? "No credits remaining - upgrade to Pro for unlimited usage!"
      : `Only ${credits} credit${credits === 1 ? '' : 's'} remaining. Upgrade to Pro!`;

    if (typeof showMessage === 'function') {
      showMessage(message, 'warning');
    }
  }

  /**
   * Show credit-related error
   */
  showCreditError(message) {
    console.error("❌ Credit error:", message);
    if (typeof showError === 'function') {
      showError(`Credit system error: ${message}`);
    }
  }

  /**
   * Start subscription process
   */
  async startSubscription() {
    try {
      console.log("💰 Starting subscription process...");
      
      if (typeof showMessage === 'function') {
        showMessage("Redirecting to checkout...");
      }

      const checkoutData = await backendAPI.createCheckoutSession();
      
      if (checkoutData.url) {
        console.log("✅ Checkout session created, redirecting...");
        window.open(checkoutData.url, '_blank');
      } else {
        throw new Error('No checkout URL received');
      }

    } catch (error) {
      console.error("❌ Failed to start subscription:", error);
      this.showCreditError(`Failed to start subscription: ${error.message}`);
    }
  }

  // ============================================================================
  // UTILITY FUNCTIONS
  // ============================================================================

  /**
   * Format credits display
   */
  formatCreditsDisplay(credits) {
    if (credits === 0) return "No credits";
    if (credits === 1) return "1.0 credit";
    return `${credits.toFixed(1)} credits`;
  }

  /**
   * Show low credits warning if needed
   */
  checkAndShowLowCreditsWarning() {
    const credits = getUserCredits();
    const hasSubscription = this.hasActiveSubscription();

    if (!hasSubscription && credits < 3) {
      this.showSubscriptionPrompt(credits);
    }
  }
}

// ============================================================================
// SINGLETON INSTANCE
// ============================================================================

const creditSystemManager = new CreditSystemManager();

// ============================================================================
// GLOBAL FUNCTIONS FOR COMPATIBILITY
// ============================================================================

/**
 * Check if user can use features
 */
function canUseFeature(featureType) {
  return creditSystemManager.canUseFeatures();
}

/**
 * Enforce feature access
 */
async function enforceFeatureAccess(featureType, callback) {
  return await creditSystemManager.enforceFeatureAccess(featureType, callback);
}

/**
 * Use credit for build action
 */
async function useCreditForBuild() {
  return await creditSystemManager.useCreditForBuild();
}

/**
 * Use credit for update action
 */
async function useCreditForUpdate() {
  return await creditSystemManager.useCreditForUpdate();
}

/**
 * Start subscription
 */
async function startSubscription() {
  return await creditSystemManager.startSubscription();
}

// ============================================================================
// EXPORTS
// ============================================================================

export {
  creditSystemManager as default,
  CreditSystemManager,
  canUseFeature,
  enforceFeatureAccess,
  useCreditForBuild,
  useCreditForUpdate,
  startSubscription
};