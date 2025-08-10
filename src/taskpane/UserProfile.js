/**
 * User Profile Management
 * Handles user data display, credits, and subscription status
 */

import backendAPI from './BackendAPI.js';

class UserProfileManager {
  constructor() {
    this.userData = null;
    this.subscriptionData = null;
    this.usageData = null;
  }

  // ============================================================================
  // DATA MANAGEMENT
  // ============================================================================

  /**
   * Initialize user data from backend
   */
  async initializeUserData() {
    try {
      console.log("👤 Initializing user data...");
      
      // Fetch user profile
      this.userData = await backendAPI.getUserProfile();
      
      // Fetch subscription status
      this.subscriptionData = await backendAPI.getSubscriptionStatus();
      
      // Update UI with loaded data
      this.updateUserInterface();
      
      console.log("✅ User data initialized successfully");
      return true;
      
    } catch (error) {
      console.error("❌ Failed to initialize user data:", error);
      this.handleInitializationError(error);
      return false;
    }
  }

  /**
   * Refresh user data from backend
   */
  async refreshUserData() {
    return await this.initializeUserData();
  }

  /**
   * Force update the UI display (useful for troubleshooting)
   */
  forceUpdateDisplay() {
    console.log("🔧 Force updating user interface display");
    
    // If no user data exists, show loading/unavailable state
    if (!this.userData) {
      this.userData = {
        name: "Loading...",
        email: "Backend unavailable", 
        credits: 0
      };
    }
    
    if (!this.subscriptionData) {
      this.subscriptionData = {
        hasActiveSubscription: false,
        status: 'none'
      };
    }
    
    this.updateUserInterface();
  }

  /**
   * Get current user data
   */
  getUserData() {
    return this.userData;
  }

  /**
   * Get current subscription data
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
   * Check if user can use features (has credits or subscription)
   */
  canUseFeatures() {
    if (this.hasActiveSubscription()) return true;
    return (this.userData?.credits || 0) > 0;
  }

  /**
   * Get user's credit count
   */
  getCredits() {
    return this.userData?.credits || 0;
  }

  // ============================================================================
  // UI UPDATES
  // ============================================================================

  /**
   * Update user interface with current data
   */
  updateUserInterface() {
    if (!this.userData) return;

    this.updateUserInfo();
    this.updateCreditsDisplay();
    this.updateSubscriptionDisplay();
    this.checkLowCreditsWarning();
    
    // Also update footer if the function exists
    if (typeof window.updateFooterDisplay === 'function') {
      window.updateFooterDisplay();
    }
  }

  /**
   * Update user information display
   */
  updateUserInfo() {
    const userInfo = document.getElementById('user-info');
    if (!userInfo || !this.userData) return;

    const userName = this.userData.name || 'User';
    const userEmail = this.userData.email || '';

    // Show the user info section
    userInfo.style.display = 'block';

    // Update user name
    const nameElement = userInfo.querySelector('.user-name');
    if (nameElement) {
      nameElement.textContent = userName;
    }

    // Update user email
    const emailElement = userInfo.querySelector('.user-email');
    if (emailElement) {
      emailElement.textContent = userEmail;
    }

    console.log(`👤 Updated user info: ${userName} (${userEmail}) - user-info section shown`);
  }

  /**
   * Update credits display
   */
  updateCreditsDisplay() {
    const credits = this.getCredits();
    
    // Update credits counter (both sidebar and client mode)
    const creditsElements = document.querySelectorAll('.credits-count, #credits-count, #credits-count-client');
    creditsElements.forEach(element => {
      element.textContent = credits.toFixed(1);
    });

    // Update credits status
    const creditsStatus = document.querySelectorAll('.credits-status');
    creditsStatus.forEach(element => {
      if (credits === 0) {
        element.textContent = 'No credits remaining';
        element.className = 'credits-status credits-empty';
      } else if (credits < 3) {
        element.textContent = `${credits.toFixed(1)} credits remaining`;
        element.className = 'credits-status credits-low';
      } else {
        element.textContent = `${credits.toFixed(1)} credits available`;
        element.className = 'credits-status credits-good';
      }
    });

    // Check if low credits warning should be shown
    this.checkLowCreditsWarning();

    console.log(`💳 Updated credits display: ${credits} (elements found: ${creditsElements.length})`);
  }

  /**
   * Update subscription display
   */
  updateSubscriptionDisplay() {
    if (!this.subscriptionData) return;

    const hasActive = this.hasActiveSubscription();
    const status = this.subscriptionData.status || 'none';

    // Update subscription badge
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

    // Update subscription buttons
    const upgradeButtons = document.querySelectorAll('.upgrade-button');
    upgradeButtons.forEach(button => {
      button.style.display = hasActive ? 'none' : 'inline-block';
    });

    console.log(`📋 Updated subscription display: ${status} (active: ${hasActive})`);
  }

  /**
   * Check and show low credits warning
   */
  checkLowCreditsWarning() {
    const credits = this.getCredits();
    const hasSubscription = this.hasActiveSubscription();

    console.log(`🔔 Credits warning check: ${credits} credits, hasSubscription: ${hasSubscription}`);

    if (credits < 2 && !hasSubscription) {
      console.log("🔔 Showing low credits warning");
      this.showLowCreditsWarning();
    } else {
      console.log("🔔 Hiding low credits warning");
      this.hideLowCreditsWarning();
    }
  }

  /**
   * Show low credits warning
   */
  showLowCreditsWarning() {
    const warning = document.getElementById('low-credits-warning');
    if (warning) {
      warning.style.display = 'block';
    }

    // Show notification
    if (typeof showMessage === 'function') {
      showMessage("You're running low on credits. Consider upgrading to Pro for unlimited usage!", 'warning');
    }
  }

  /**
   * Hide low credits warning
   */
  hideLowCreditsWarning() {
    const warning = document.getElementById('low-credits-warning');
    if (warning) {
      warning.style.display = 'none';
    }
  }

  // ============================================================================
  // ERROR HANDLING
  // ============================================================================

  /**
   * Handle initialization errors
   */
  handleInitializationError(error) {
    console.error("User data initialization failed:", error);

    // For development mode without backend, show user as not loaded
    if (!this.userData) {
      console.log("⚠️ Backend unavailable - user data not loaded");
      // Don't set fake data, let the app handle missing data gracefully
    }

    if (error.message.includes('Network error')) {
      this.showErrorMessage("Unable to connect to our servers. Some features may be limited.");
    } else if (error.message.includes('Authentication failed')) {
      this.showErrorMessage("Session expired. Please sign in again.");
    } else {
      this.showErrorMessage("Unable to load your profile. Some features may be limited.");
    }
  }

  /**
   * Show error message to user
   */
  showErrorMessage(message) {
    if (typeof showError === 'function') {
      showError(message);
    } else {
      console.error("Error message:", message);
    }
  }

  // ============================================================================
  // USAGE ANALYTICS
  // ============================================================================

  /**
   * Load and display usage analytics
   */
  async loadUsageAnalytics() {
    try {
      console.log("📊 Loading usage analytics...");
      
      this.usageData = await backendAPI.getUsageAnalytics();
      this.displayUsageStats();
      
      console.log("✅ Usage analytics loaded");
      
    } catch (error) {
      console.error("❌ Failed to load usage analytics:", error);
    }
  }

  /**
   * Display usage statistics
   */
  displayUsageStats() {
    if (!this.usageData) return;

    const stats = {
      totalBuilds: this.usageData.summary?.build || 0,
      totalUpdates: this.usageData.summary?.update || 0,
      totalActions: this.usageData.totalActions || 0,
      recentActivity: this.usageData.recentActivity || []
    };

    // Update stats displays
    const buildsElement = document.getElementById('total-builds');
    if (buildsElement) buildsElement.textContent = stats.totalBuilds;

    const updatesElement = document.getElementById('total-updates');
    if (updatesElement) updatesElement.textContent = stats.totalUpdates;

    const actionsElement = document.getElementById('total-actions');
    if (actionsElement) actionsElement.textContent = stats.totalActions;

    // Update recent activity list
    const activityList = document.getElementById('recent-activity-list');
    if (activityList && stats.recentActivity.length > 0) {
      activityList.innerHTML = stats.recentActivity
        .slice(0, 5) // Show only last 5 activities
        .map(activity => `
          <div class="activity-item">
            <span class="activity-action">${activity.action}</span>
            <span class="activity-date">${new Date(activity.createdAt).toLocaleDateString()}</span>
          </div>
        `).join('');
    }

    console.log("📊 Usage stats displayed:", stats);
  }
}

// ============================================================================
// SINGLETON INSTANCE
// ============================================================================

const userProfileManager = new UserProfileManager();

// Expose globally for other modules to access
if (typeof window !== 'undefined') {
  window.userProfileManager = userProfileManager;
}

// ============================================================================
// GLOBAL FUNCTIONS
// ============================================================================

/**
 * Initialize user data (global function for compatibility)
 */
async function initializeUserData() {
  return await userProfileManager.initializeUserData();
}

/**
 * Refresh user data (global function for compatibility)
 */
async function refreshUserData() {
  return await userProfileManager.refreshUserData();
}

/**
 * Check if user can use features (global function for compatibility)
 */
function canUseFeatures() {
  return userProfileManager.canUseFeatures();
}

/**
 * Get user credits (global function for compatibility)
 */
function getUserCredits() {
  return userProfileManager.getCredits();
}

/**
 * Force update user display (global function for troubleshooting)
 */
function forceUpdateUserDisplay() {
  return userProfileManager.forceUpdateDisplay();
}

// ============================================================================
// EXPORTS
// ============================================================================

export {
  userProfileManager as default,
  UserProfileManager,
  initializeUserData,
  refreshUserData,
  canUseFeatures,
  getUserCredits,
  forceUpdateUserDisplay
};