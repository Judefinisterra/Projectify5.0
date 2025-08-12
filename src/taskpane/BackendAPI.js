/**
 * Backend API Integration for Excel Add-in
 * Handles authentication, credits, subscriptions, and user management
 */

// Import configuration
import { CONFIG } from './config.js';

class BackendAPI {
  constructor() {
    // Check if backend configuration exists
    if (!CONFIG.backend) {
      console.warn('⚠️ Backend configuration not found in CONFIG. Using mock mode.');
      this.mockMode = true;
    } else {
      this.baseUrl = CONFIG.backend.baseUrl || 'https://your-backend-api.com';
      this.endpoints = CONFIG.backend.endpoints || {};
      this.timeout = CONFIG.backend.timeout || 10000;
      
      // Only use mock mode if backend URL is a placeholder
      if (this.baseUrl.includes('your-backend-api.com') || this.baseUrl.includes('your-dev-backend-api.com')) {
        console.warn('⚠️ Backend URL appears to be a placeholder. Enabling mock mode.');
        this.mockMode = true;
      } else {
        console.log('✅ Using real backend at:', this.baseUrl);
        this.mockMode = false;
      }
    }
    
    console.log('🔧 BackendAPI initialized:', this.mockMode ? 'Mock Mode' : `Real API at ${this.baseUrl}`);
  }

  // ============================================================================
  // AUTHENTICATION HELPERS
  // ============================================================================

  /**
   * Get authentication headers for API calls
   */
  getAuthHeaders() {
    // Try sessionStorage first, then localStorage for persistence
    const token = sessionStorage.getItem('backend_access_token') || 
                  sessionStorage.getItem('access_token') ||
                  localStorage.getItem('backend_access_token') ||
                  localStorage.getItem('access_token');
    console.log('🔧 BackendAPI getAuthHeaders - token:', token ? `${token.substring(0, 20)}...` : 'No token found');
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
      // Store in both session and local storage for persistence
      sessionStorage.setItem('backend_access_token', accessToken);
      localStorage.setItem('backend_access_token', accessToken);
      sessionStorage.setItem('access_token', accessToken); // Backward compatibility
      localStorage.setItem('access_token', accessToken); // Backward compatibility
    }
    if (refreshToken) {
      // Store in both session and local storage for persistence
      sessionStorage.setItem('backend_refresh_token', refreshToken);
      localStorage.setItem('backend_refresh_token', refreshToken);
      sessionStorage.setItem('refresh_token', refreshToken); // Backward compatibility
      localStorage.setItem('refresh_token', refreshToken); // Backward compatibility
    }
  }

  /**
   * Clear stored tokens
   */
  clearTokens() {
    // Clear from both session and local storage
    sessionStorage.removeItem('backend_access_token');
    sessionStorage.removeItem('backend_refresh_token');
    sessionStorage.removeItem('access_token');
    sessionStorage.removeItem('refresh_token');
    
    localStorage.removeItem('backend_access_token');
    localStorage.removeItem('backend_refresh_token');
    localStorage.removeItem('access_token');
    localStorage.removeItem('refresh_token');
  }

  /**
   * Check if user is authenticated with backend
   */
  isAuthenticated() {
    // Check both session and local storage
    const hasToken = !!sessionStorage.getItem('backend_access_token') || 
                    !!localStorage.getItem('backend_access_token');
    
    // In mock mode, also consider authenticated if Google auth was successful
    if (this.mockMode) {
      return hasToken || !!window.googleUser;
    }
    return hasToken;
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
    
    // Return mock authentication in mock mode
    if (this.mockMode) {
      console.log('🔧 Mock mode: Returning mock authentication');
      const mockAuthData = {
        access_token: 'mock-access-token-123',
        refresh_token: 'mock-refresh-token-123',
        user: {
          id: 'mock-user-123',
          name: 'Development User',
          email: 'dev@example.com'
        },
        session: {
          refresh_token: 'mock-refresh-token-123'
        }
      };
      
      // Store mock tokens
      this.storeTokens(mockAuthData.access_token, mockAuthData.refresh_token);
      console.log('✅ Mock backend authentication successful');
      return mockAuthData;
    }
    
    // For authentication, make direct fetch call without auth headers
    const url = `${this.baseUrl}${this.endpoints.auth.google}`;
    
    try {
      console.log(`🌐 Auth API Call: POST ${url}`);
      
      const response = await fetch(url, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({ token: googleIdToken })
      });
      
      if (!response.ok) {
        const error = await response.json().catch(() => ({ error: 'Unknown error' }));
        throw new Error(error.error || error.message || `HTTP ${response.status}`);
      }
      
      const data = await response.json();
      
      if (data.access_token) {
        this.storeTokens(data.access_token, data.refresh_token || data.session?.refresh_token);
        console.log('✅ Backend authentication successful');
        return data;
      }
      
      throw new Error('Backend authentication failed - no access token received');
      
    } catch (error) {
      console.error('❌ Backend auth error:', error);
      throw error;
    }
  }

  /**
   * Refresh authentication token
   */
  async refreshToken() {
    // Try sessionStorage first, then localStorage
    const refreshToken = sessionStorage.getItem('backend_refresh_token') || 
                        localStorage.getItem('backend_refresh_token');
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
    
    // Return mock data in mock mode
    if (this.mockMode) {
      console.log('🔧 Mock mode: Returning mock user profile');
      const mockUser = {
        id: 'mock-user-123',
        name: 'Development User',
        email: 'dev@example.com',
        credits: 15,
        subscription: {
          status: 'none',
          hasActiveSubscription: false
        }
      };
      console.log('✅ Mock user profile loaded:', mockUser.name, `(${mockUser.credits} credits)`);
      return mockUser;
    }
    
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
   * Use a credit for an action (legacy)
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

  /**
   * Deduct credits based on actual API call cost
   */
  async deductCredits(costData) {
    const { cost, provider, model, tokens, action, caller } = costData;
    
    console.log(`💰 Deducting credits for ${provider} API call: $${cost.toFixed(6)}`);
    
    const response = await this.makeAPICall(this.endpoints.credits.deduct, {
      method: 'POST',
      body: JSON.stringify({ 
        cost, 
        provider, 
        model, 
        tokens, 
        action: action || `${provider}_api_call`,
        caller 
      })
    });

    const result = await response.json();
    
    if (result.success) {
      console.log(`✅ Credits deducted: ${result.creditCost.toFixed(4)} credits. Remaining: ${result.remainingCredits}`);
    } else {
      console.error(`❌ Failed to deduct credits:`, result.error);
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
    
    // Return mock data in mock mode
    if (this.mockMode) {
      console.log('🔧 Mock mode: Returning mock subscription status');
      const mockSubscription = {
        status: 'none',
        hasActiveSubscription: false,
        planType: 'free',
        creditsRemaining: 15
      };
      console.log('✅ Mock subscription status loaded:', mockSubscription.status);
      return mockSubscription;
    }
    
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