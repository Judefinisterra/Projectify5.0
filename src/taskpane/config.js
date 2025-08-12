// config.js
// Configuration for both development and production environments
export const CONFIG = {
  // Determine if we're in development or production
  isDevelopment: process.env.NODE_ENV === 'development',
  
  // Base URL for assets and API calls
  get baseUrl() {
    return this.isDevelopment ? 'https://localhost:3002' : '';
  },
  
  // Helper function to get full URL for assets
  getAssetUrl(path) {
    // Remove leading slash if present
    const cleanPath = path.startsWith('/') ? path.slice(1) : path;
    let fullUrl;
    
    if (this.isDevelopment) {
      fullUrl = `${this.baseUrl}/${cleanPath}`;
      console.log(`[CONFIG] Asset URL (dev): ${cleanPath} -> ${fullUrl}`);
    } else {
      // In production, try relative path first
      fullUrl = `./${cleanPath}`;
      console.log(`[CONFIG] Asset URL (prod): ${cleanPath} -> ${fullUrl}`);
    }
    
    return fullUrl;
  },

  // Helper function to get multiple possible asset URLs for fallback
  getAssetUrlsWithFallback(path) {
    const cleanPath = path.startsWith('/') ? path.slice(1) : path;
    
    if (this.isDevelopment) {
      return [`${this.baseUrl}/${cleanPath}`];
    } else {
      // In production, try multiple possible paths
      return [
        `./${cleanPath}`,           // Relative to current page
        `/${cleanPath}`,            // Absolute from root
        `_next/static/${cleanPath}`, // Next.js static folder
        `assets/${cleanPath.replace('assets/', '')}`, // Direct assets folder
      ];
    }
  },
  
  // Helper function to get prompt URL
  getPromptUrl(filename) {
    return this.isDevelopment 
      ? `https://localhost:3002/prompts/${filename}`
      : `./prompts/${filename}`;
  },

  // Backend API Configuration
  // Backend API Configuration
  backend: {
    baseUrl: "https://backend-projectify-74f2tkrkr-judefinisterras-projects.vercel.app",
    timeout: 30000,
    endpoints: {
      health: '/health',  // Add health endpoint
      auth: {
        google: '/auth/google',
        refresh: '/auth/refresh'
      },
      user: {
        profile: '/me',
        credits: '/use-credit', // Legacy
        usage: '/usage'
      },
      credits: {
        use: '/use-credit', // Legacy
        deduct: '/deduct-credits' // New
      },
      subscription: {
        status: '/subscription',
        create: '/subscription/create',
        cancel: '/subscription/cancel',
        checkout: '/api/subscription/checkout'
      }
    }
  },

  // Microsoft Authentication Configuration
  // Subscription API Configuration
  get subscription() {
    return {
      // Your backend API endpoint for checking subscription status (automatically derived from backend.baseUrl)
      apiUrl: `${this.backend.baseUrl}/subscription`,
      
      // Enable mock mode for development (returns fake subscription data)
      mockMode: false, // Disabled - use real backend
      
      // Mock response for development/testing
      mockResponse: {
        status: 'none',
        plan: 'Free',
        hasActiveSubscription: false,
        credits: 15,
        expires: null
      }
    };
  },

  authentication: {
    // To set up Microsoft authentication:
    // 1. Go to Azure Active Directory > App registrations > New registration
    // 2. Set name (e.g., "Projectify Add-in") and select "Accounts in any organizational directory and personal Microsoft accounts"
    // 3. Set redirect URI to your add-in URL (e.g., https://yourdomain.com/src/taskpane/taskpane.html)
    // 4. Copy the Application (client) ID below
    // 5. In Authentication settings, enable "Access tokens" and "ID tokens"
    // 6. In API permissions, ensure "User.Read" is granted
    
    msalConfig: {
      auth: {
        // IMPORTANT: Replace this with your actual Azure App Registration Client ID
        // Get this from: Azure Portal > Azure Active Directory > App registrations > Your App > Overview
        clientId: process.env.NODE_ENV === 'development' 
          ? "your-azure-app-client-id-here" // For development - replace this
          : "your-production-client-id-here", // For production - replace this
        
        authority: "https://login.microsoftonline.com/common", // For multi-tenant apps (personal + work accounts)
        
        // Redirect URI must match exactly what's configured in Azure
        redirectUri: (() => {
          if (process.env.NODE_ENV === 'development') {
            return "https://localhost:3002/src/taskpane/taskpane.html";
          }
          // For production, use actual domain
          return window.location.origin + "/src/taskpane/taskpane.html";
        })()
      },
      cache: {
        cacheLocation: "sessionStorage", // Use sessionStorage for better security in Office Add-ins
        storeAuthStateInCookie: false // Set to true if you have IE11 support requirements
      }
    },
    
    // Login Request Configuration
    loginRequest: {
      scopes: ["User.Read", "profile", "openid", "email"] // Permissions we need
    },
    
    // Token Request Configuration  
    tokenRequest: {
      scopes: ["User.Read"]
    },
    
    // Graph API Configuration
    graphConfig: {
      graphMeUrl: "https://graph.microsoft.com/v1.0/me", // Get user profile
      graphPhotoUrl: "https://graph.microsoft.com/v1.0/me/photo/$value" // Get user photo
    },
    
    // Development helper
    isDevelopmentPlaceholder() {
      return this.msalConfig.auth.clientId.includes('your-azure-app-client-id-here');
    }
  }
};
//dfd
// Log configuration on load (for debugging)
if (CONFIG.isDevelopment) {
  console.log('Config loaded:', {
    isDevelopment: CONFIG.isDevelopment,
    baseUrl: CONFIG.baseUrl
  });
}

