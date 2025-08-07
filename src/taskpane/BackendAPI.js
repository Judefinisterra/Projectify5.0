/**
 * Backend API Integration for Excel Add-in
 * Handles authentication, credits, subscriptions, and user management
 */

// Import configuration
import { CONFIG } from './config.js';

class BackendAPI {
  constructor() {
    this.baseUrl = CONFIG.backend.baseUrl;
    this.endpoints = CONFIG.backend.endpoints;
    this.timeout = CONFIG.backend.timeout;
  }

  // ============================================================================
  // AUTHENTICATION HELPERS
  // ============================================================================

  /**
   * Get authentication headers for API calls
   */
  getAuthHeaders() {
    const token = sessionStorage.getItem('backend_access_token') || sessionStorage.getItem('access_token');
    return {
      'Content-Type': 'application/json',
      ...(token && { 'Authorization': `Bearer ${token}` })
    };
  }

  /**
   * Store authentication tokens
   */
  storeTokens(accessToken, refreshToken) {
    if (accessToken) {
      sessionStorage.setItem('backend_access_token', accessToken);
      sessionStorage.setItem('access_token', accessToken); // Backward compatibility
    }
    if (refreshToken) {
      sessionStorage.setItem('backend_refresh_token', refreshToken);
      sessionStorage.setItem('refresh_token', refreshToken); // Backward compatibility
    }
  }

  /**
   * Clear stored tokens
   */
  clearTokens() {
    sessionStorage.removeItem('backend_access_token');
    sessionStorage.removeItem('backend_refresh_token');
    sessionStorage.removeItem('access_token');
    sessionStorage.removeItem('refresh_token');
  }

  /**
   * Check if user is authenticated with backend
   */
  isAuthenticated() {
    return !!sessionStorage.getItem('backend_access_token');
  }

  // ============================================================================
  // CORE API METHODS
  // ============================================================================

  /**
   * Make authenticated API call with error handling and token refresh
   */
  async makeAPICall(endpoint, options = {}) {
    const url = `${this.baseUrl}${endpoint}`;
    const requestOptions = {
      timeout: this.timeout,
      ...options,
      headers: { 
        ...this.getAuthHeaders(), 
        ...options.headers 
      }
    };

    try {
      console.log(`🌐 API Call: ${options.method || 'GET'} ${url}`);
      
      const response = await fetch(url, requestOptions);
      
      // Handle authentication errors
      if (response.status === 401) {
        console.log('🔄 Token expired, attempting refresh...');
        const refreshed = await this.refreshToken();
        if (refreshed) {
          // Retry with new token
          const retryOptions = {
            ...requestOptions,
            headers: { 
              ...this.getAuthHeaders(), 
              ...options.headers 
            }
          };
          return fetch(url, retryOptions);
        } else {
          throw new Error('Authentication failed');
        }
      }

      // Handle insufficient credits
      if (response.status === 402) {
        const error = await response.json();
        throw new InsufficientCreditsError(error.credits || 0, error.message);
      }

      // Handle other errors
      if (!response.ok) {
        const error = await response.json().catch(() => ({ error: 'Unknown error' }));
        throw new Error(error.error || error.message || `HTTP ${response.status}`);
      }

      return response;
      
    } catch (error) {
      console.error('❌ API Error:', error);
      if (error.name === 'TypeError' && error.message.includes('fetch')) {
        throw new Error('Network error - please check your connection');
      }
      throw error;
    }
  }

  // ============================================================================
  // AUTHENTICATION API
  // ============================================================================

  /**
   * Authenticate with Google OAuth token
   */
  async signInWithGoogle(googleIdToken) {
    console.log('🔐 Signing in with Google via backend...');
    
    const response = await this.makeAPICall(this.endpoints.auth.google, {
      method: 'POST',
      body: JSON.stringify({ token: googleIdToken })
    });

    const data = await response.json();
    
    if (data.access_token) {
      this.storeTokens(data.access_token, data.session?.refresh_token);
      console.log('✅ Backend authentication successful');
      return data;
    }
    
    throw new Error('Backend authentication failed');
  }

  /**
   * Refresh authentication token
   */
  async refreshToken() {
    const refreshToken = sessionStorage.getItem('backend_refresh_token');
    if (!refreshToken) {
      console.log('❌ No refresh token available');
      return false;
    }

    try {
      const response = await fetch(`${this.baseUrl}${this.endpoints.auth.refresh}`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ refresh_token: refreshToken })
      });

      const data = await response.json();
      if (data.access_token) {
        this.storeTokens(data.access_token);
        console.log('✅ Token refreshed successfully');
        return true;
      }
    } catch (error) {
      console.error('❌ Token refresh failed:', error);
    }

    this.clearTokens();
    return false;
  }

  // ============================================================================
  // USER PROFILE API
  // ============================================================================

  /**
   * Get user profile and credits
   */
  async getUserProfile() {
    console.log('👤 Fetching user profile...');
    
    const response = await this.makeAPICall(this.endpoints.user.profile);
    const user = await response.json();
    
    console.log('✅ User profile loaded:', user.name, `(${user.credits} credits)`);
    return user;
  }

  /**
   * Get usage analytics
   */
  async getUsageAnalytics() {
    console.log('📊 Fetching usage analytics...');
    
    const response = await this.makeAPICall(this.endpoints.user.usage);
    const usage = await response.json();
    
    console.log('✅ Usage analytics loaded');
    return usage;
  }

  // ============================================================================
  // CREDITS API
  // ============================================================================

  /**
   * Use a credit for an action
   */
  async useCredit(action) {
    console.log(`💳 Using credit for action: ${action}`);
    
    const response = await this.makeAPICall(this.endpoints.credits.use, {
      method: 'POST',
      body: JSON.stringify({ action })
    });

    const result = await response.json();
    
    if (result.success) {
      console.log(`✅ Credit used. Remaining: ${result.remainingCredits}`);
    }
    
    return result;
  }

  // ============================================================================
  // SUBSCRIPTION API
  // ============================================================================

  /**
   * Create Stripe checkout session
   */
  async createCheckoutSession(successUrl, cancelUrl) {
    console.log('💰 Creating Stripe checkout session...');
    
    const response = await this.makeAPICall(this.endpoints.subscription.create, {
      method: 'POST',
      body: JSON.stringify({ 
        successUrl: successUrl || `${window.location.origin}/success`,
        cancelUrl: cancelUrl || `${window.location.origin}/cancel`
      })
    });

    const data = await response.json();
    console.log('✅ Checkout session created');
    return data;
  }

  /**
   * Get subscription status
   */
  async getSubscriptionStatus() {
    console.log('📋 Fetching subscription status...');
    
    const response = await this.makeAPICall(this.endpoints.subscription.status);
    const subscription = await response.json();
    
    console.log('✅ Subscription status loaded:', subscription.status);
    return subscription;
  }

  /**
   * Cancel subscription
   */
  async cancelSubscription() {
    console.log('❌ Cancelling subscription...');
    
    const response = await this.makeAPICall(this.endpoints.subscription.cancel, {
      method: 'POST'
    });

    const result = await response.json();
    
    if (result.success) {
      console.log('✅ Subscription cancelled');
    }
    
    return result;
  }

  // ============================================================================
  // HEALTH CHECK
  // ============================================================================

  /**
   * Check backend health
   */
  async healthCheck() {
    try {
      const response = await fetch(`${this.baseUrl}${this.endpoints.health}`, {
        timeout: 5000
      });
      return response.ok;
    } catch (error) {
      console.error('❌ Backend health check failed:', error);
      return false;
    }
  }
}

// ============================================================================
// CUSTOM ERROR CLASSES
// ============================================================================

class InsufficientCreditsError extends Error {
  constructor(credits, message) {
    super(message || 'Insufficient credits');
    this.name = 'InsufficientCreditsError';
    this.credits = credits;
  }
}

// ============================================================================
// SINGLETON INSTANCE
// ============================================================================

const backendAPI = new BackendAPI();

// ============================================================================
// EXPORTS
// ============================================================================

export { 
  backendAPI as default, 
  BackendAPI, 
  InsufficientCreditsError 
};