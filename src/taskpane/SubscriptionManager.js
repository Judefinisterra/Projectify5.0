/**
 * Subscription Management
 * Handles Stripe subscriptions, cancellation, and status tracking
 */

import backendAPI from './BackendAPI.js';
import { refreshUserData } from './UserProfile.js';

class SubscriptionManager {
  constructor() {
    this.subscriptionData = null;
    this.checkoutInProgress = false;
  }

  // ============================================================================
  // SUBSCRIPTION OPERATIONS
  // ============================================================================

  /**
   * Start subscription checkout process
   */
  async startCheckout(successUrl = null, cancelUrl = null) {
    if (this.checkoutInProgress) {
      console.log("⚠️ Checkout already in progress");
      return;
    }

    try {
      this.checkoutInProgress = true;
      console.log("💰 Starting subscription checkout...");

      if (typeof showMessage === 'function') {
        showMessage("Redirecting to checkout...");
      }

      // Use current URL as base for redirect URLs
      const baseUrl = window.location.origin + window.location.pathname.replace(/\/[^/]*$/, '');
      
      const checkoutData = await backendAPI.createCheckoutSession(
        successUrl || `${baseUrl}?checkout=success`,
        cancelUrl || `${baseUrl}?checkout=cancelled`
      );

      if (checkoutData.url) {
        console.log("✅ Checkout session created");
        console.log("🔗 Checkout URL:", checkoutData.url);

        // Open checkout in new window/tab
        const checkoutWindow = window.open(
          checkoutData.url, 
          'stripe-checkout', 
          'width=800,height=600,scrollbars=yes,resizable=yes'
        );

        if (!checkoutWindow) {
          throw new Error('Popup blocked. Please allow popups and try again.');
        }

        // Monitor checkout window
        this.monitorCheckoutWindow(checkoutWindow);

      } else {
        throw new Error('No checkout URL received from backend');
      }

    } catch (error) {
      console.error("❌ Failed to start checkout:", error);
      this.handleCheckoutError(error);
    } finally {
      this.checkoutInProgress = false;
    }
  }

  /**
   * Monitor checkout window for completion
   */
  monitorCheckoutWindow(checkoutWindow) {
    const checkInterval = setInterval(() => {
      try {
        if (checkoutWindow.closed) {
          clearInterval(checkInterval);
          console.log("🔄 Checkout window closed, refreshing user data...");
          
          // Refresh user data to get updated subscription status
          setTimeout(async () => {
            try {
              await refreshUserData();
              await this.checkSubscriptionStatus();
              
              if (typeof showMessage === 'function') {
                showMessage("Subscription status updated!");
              }
            } catch (error) {
              console.error("❌ Failed to refresh after checkout:", error);
            }
          }, 2000); // Wait 2 seconds for backend to process
          
          return;
        }

        // Check if we can access the URL (same origin after redirect)
        try {
          const currentUrl = checkoutWindow.location.href;
          
          if (currentUrl.includes('checkout=success')) {
            clearInterval(checkInterval);
            checkoutWindow.close();
            this.handleCheckoutSuccess();
          } else if (currentUrl.includes('checkout=cancelled')) {
            clearInterval(checkInterval);
            checkoutWindow.close();
            this.handleCheckoutCancelled();
          }
        } catch (e) {
          // Cross-origin error is expected while on Stripe domain
        }

      } catch (error) {
        // Ignore errors from closed window
      }
    }, 1000);

    // Stop monitoring after 15 minutes
    setTimeout(() => {
      clearInterval(checkInterval);
      if (!checkoutWindow.closed) {
        console.log("⏰ Checkout monitoring timed out");
      }
    }, 15 * 60 * 1000);
  }

  /**
   * Handle successful checkout
   */
  async handleCheckoutSuccess() {
    console.log("✅ Checkout completed successfully!");
    
    try {
      // Refresh user data
      await refreshUserData();
      await this.checkSubscriptionStatus();
      
      if (typeof showMessage === 'function') {
        showMessage("🎉 Welcome to Pro! Your subscription is now active.");
      }

      // Hide upgrade buttons and warnings
      this.hideUpgradePrompts();
      
    } catch (error) {
      console.error("❌ Error handling checkout success:", error);
      if (typeof showMessage === 'function') {
        showMessage("Subscription activated! Please refresh to see changes.");
      }
    }
  }

  /**
   * Handle cancelled checkout
   */
  handleCheckoutCancelled() {
    console.log("❌ Checkout was cancelled");
    
    if (typeof showMessage === 'function') {
      showMessage("Checkout cancelled. You can upgrade anytime!");
    }
  }

  /**
   * Handle checkout errors
   */
  handleCheckoutError(error) {
    console.error("❌ Checkout error:", error);
    
    const message = error.message.includes('Popup blocked') 
      ? "Popup blocked! Please allow popups and try again."
      : `Checkout failed: ${error.message}`;
      
    if (typeof showError === 'function') {
      showError(message);
    }
  }

  // ============================================================================
  // SUBSCRIPTION STATUS
  // ============================================================================

  /**
   * Check current subscription status
   */
  async checkSubscriptionStatus() {
    try {
      console.log("📋 Checking subscription status...");
      
      this.subscriptionData = await backendAPI.getSubscriptionStatus();
      
      console.log("✅ Subscription status:", this.subscriptionData.status);
      console.log("📊 Has active subscription:", this.subscriptionData.hasActiveSubscription);
      
      this.updateSubscriptionUI();
      
      return this.subscriptionData;
      
    } catch (error) {
      console.error("❌ Failed to check subscription status:", error);
      return null;
    }
  }

  /**
   * Cancel current subscription
   */
  async cancelSubscription() {
    try {
      const confirmed = confirm(
        "Are you sure you want to cancel your subscription? " +
        "You'll continue to have access until the end of your current billing period."
      );
      
      if (!confirmed) {
        console.log("❌ Subscription cancellation cancelled by user");
        return false;
      }

      console.log("❌ Cancelling subscription...");
      
      if (typeof showMessage === 'function') {
        showMessage("Cancelling subscription...");
      }

      const result = await backendAPI.cancelSubscription();
      
      if (result.success) {
        console.log("✅ Subscription cancelled successfully");
        
        const endDate = result.cancelAt ? new Date(result.cancelAt).toLocaleDateString() : 'the end of the billing period';
        
        if (typeof showMessage === 'function') {
          showMessage(`Subscription cancelled. Access continues until ${endDate}.`);
        }
        
        // Refresh subscription status
        await this.checkSubscriptionStatus();
        
        return true;
      } else {
        throw new Error(result.error || 'Cancellation failed');
      }
      
    } catch (error) {
      console.error("❌ Failed to cancel subscription:", error);
      
      if (typeof showError === 'function') {
        showError(`Failed to cancel subscription: ${error.message}`);
      }
      
      return false;
    }
  }

  // ============================================================================
  // UI UPDATES
  // ============================================================================

  /**
   * Update subscription-related UI elements
   */
  updateSubscriptionUI() {
    if (!this.subscriptionData) return;

    const hasActive = this.subscriptionData.hasActiveSubscription;
    const status = this.subscriptionData.status || 'none';

    // Update subscription badges
    const badges = document.querySelectorAll('.subscription-badge');
    badges.forEach(badge => {
      badge.className = `subscription-badge subscription-${status}`;
      
      if (hasActive) {
        badge.textContent = 'Pro Active';
      } else if (status === 'trialing') {
        badge.textContent = 'Trial';
      } else if (status === 'cancelled') {
        badge.textContent = 'Cancelled';
      } else {
        badge.textContent = 'Free';
      }
    });

    // Update upgrade buttons visibility
    const upgradeButtons = document.querySelectorAll('.upgrade-button');
    upgradeButtons.forEach(button => {
      button.style.display = hasActive ? 'none' : 'inline-block';
    });

    // Update cancel buttons visibility
    const cancelButtons = document.querySelectorAll('.cancel-subscription-button');
    cancelButtons.forEach(button => {
      button.style.display = hasActive ? 'inline-block' : 'none';
      button.onclick = () => this.cancelSubscription();
    });

    // Hide low credits warning if subscription is active
    if (hasActive) {
      this.hideUpgradePrompts();
    }

    console.log(`📋 Subscription UI updated: ${status} (active: ${hasActive})`);
  }

  /**
   * Hide upgrade prompts and warnings
   */
  hideUpgradePrompts() {
    const warningElement = document.getElementById('low-credits-warning');
    if (warningElement) {
      warningElement.style.display = 'none';
    }

    const upgradeButtons = document.querySelectorAll('.upgrade-button');
    upgradeButtons.forEach(button => {
      button.style.display = 'none';
    });
  }

  /**
   * Show upgrade prompts
   */
  showUpgradePrompts() {
    const warningElement = document.getElementById('low-credits-warning');
    if (warningElement) {
      warningElement.style.display = 'block';
    }

    const upgradeButtons = document.querySelectorAll('.upgrade-button');
    upgradeButtons.forEach(button => {
      button.style.display = 'inline-block';
    });
  }

  // ============================================================================
  // UTILITY METHODS
  // ============================================================================

  /**
   * Get subscription data
   */
  getSubscriptionData() {
    return this.subscriptionData;
  }

  /**
   * Check if user has active subscription
   */
  hasActiveSubscription() {
    return this.subscriptionData?.hasActiveSubscription || false;
  }

  /**
   * Get subscription status
   */
  getSubscriptionStatus() {
    return this.subscriptionData?.status || 'none';
  }
}

// ============================================================================
// SINGLETON INSTANCE
// ============================================================================

const subscriptionManager = new SubscriptionManager();

// ============================================================================
// GLOBAL FUNCTIONS
// ============================================================================

/**
 * Start subscription checkout
 */
async function startSubscriptionCheckout() {
  return await subscriptionManager.startCheckout();
}

/**
 * Check subscription status
 */
async function checkSubscriptionStatus() {
  return await subscriptionManager.checkSubscriptionStatus();
}

/**
 * Cancel subscription
 */
async function cancelSubscription() {
  return await subscriptionManager.cancelSubscription();
}

// ============================================================================
// EXPORTS
// ============================================================================

export {
  subscriptionManager as default,
  SubscriptionManager,
  startSubscriptionCheckout,
  checkSubscriptionStatus,
  cancelSubscription
};