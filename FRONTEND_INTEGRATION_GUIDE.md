# 📋 Excel Addin Frontend Integration Instructions

## 🎯 **Objective**
Integrate the Excel addin frontend with the deployed backend API to enable Google OAuth authentication, credit-based usage, and Stripe subscription management.

## 🔧 **Backend Configuration**

**API Base URL:**
```javascript
const API_BASE_URL = 'https://backend-projectify-mpdqopqjv-judefinisterras-projects.vercel.app';
```

## 🚀 **Implementation Requirements**

### **1. Authentication System**

**Add Google OAuth Integration:**
- Install Google OAuth library: `npm install google-auth-library`
- Create Google sign-in button in your UI
- Implement authentication flow that:
  1. Gets Google ID token from user sign-in
  2. Sends token to `POST ${API_BASE_URL}/auth/google`
  3. Stores returned `access_token` for API calls
  4. Uses token in Authorization header: `Bearer <token>`

**Example Authentication Code:**
```javascript
// Sign in with Google
async function signInWithGoogle(googleIdToken) {
  const response = await fetch(`${API_BASE_URL}/auth/google`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ token: googleIdToken })
  });
  
  const data = await response.json();
  if (data.access_token) {
    localStorage.setItem('access_token', data.access_token);
    localStorage.setItem('refresh_token', data.session.refresh_token);
    return data;
  }
  throw new Error('Authentication failed');
}

// Use stored token for API calls
function getAuthHeaders() {
  const token = localStorage.getItem('access_token');
  return {
    'Content-Type': 'application/json',
    'Authorization': `Bearer ${token}`
  };
}
```

### **2. User Profile Management**

**Fetch User Data:**
- Call `GET ${API_BASE_URL}/me` on app startup
- Display user name, email, credits, and subscription status
- Show subscription status badge (trial, active, cancelled)

**Example User Profile Code:**
```javascript
async function getUserProfile() {
  const response = await fetch(`${API_BASE_URL}/me`, {
    headers: getAuthHeaders()
  });
  
  if (response.status === 401) {
    // Token expired, try refresh or redirect to login
    await refreshToken();
    return getUserProfile();
  }
  
  const user = await response.json();
  updateUI({
    name: user.name,
    email: user.email,
    credits: user.credits,
    hasSubscription: user.hasActiveSubscription,
    canUseFeatures: user.canUseFeatures
  });
  
  return user;
}
```

### **3. Credit System Integration**

**Before Model Building/Updating:**
- Check if user has credits or active subscription
- Call `POST ${API_BASE_URL}/use-credit` when user performs actions
- Handle insufficient credits by showing subscription prompt

**Example Credit Usage:**
```javascript
async function buildModel() {
  try {
    // Show loading state
    showLoading('Building model...');
    
    const response = await fetch(`${API_BASE_URL}/use-credit`, {
      method: 'POST',
      headers: getAuthHeaders(),
      body: JSON.stringify({ action: 'build' })
    });
    
    const result = await response.json();
    
    if (response.status === 402) {
      // Insufficient credits
      showSubscriptionPrompt(result.credits);
      return;
    }
    
    if (result.success) {
      // Update credits display
      updateCreditsDisplay(result.remainingCredits);
      
      // Proceed with actual model building
      await performModelBuild();
      
      showSuccess(`${result.message}. Credits remaining: ${result.remainingCredits}`);
    }
    
  } catch (error) {
    showError('Model build failed: ' + error.message);
  } finally {
    hideLoading();
  }
}

async function updateModel() {
  // Similar to buildModel but with action: 'update'
  const response = await fetch(`${API_BASE_URL}/use-credit`, {
    method: 'POST',
    headers: getAuthHeaders(),
    body: JSON.stringify({ action: 'update' })
  });
  // ... handle response
}
```

### **4. Subscription Management**

**Stripe Checkout Integration:**
- Create "Upgrade to Pro" or "Get Credits" button
- Handle checkout session creation and redirect
- Check subscription status after successful payment

**Example Subscription Code:**
```javascript
async function startSubscription() {
  try {
    const response = await fetch(`${API_BASE_URL}/create-checkout-session`, {
      method: 'POST',
      headers: getAuthHeaders(),
      body: JSON.stringify({
        successUrl: `${window.location.origin}/success`,
        cancelUrl: `${window.location.origin}/cancel`
      })
    });
    
    const { url } = await response.json();
    
    // Redirect to Stripe Checkout
    window.location.href = url;
    
  } catch (error) {
    showError('Failed to start subscription: ' + error.message);
  }
}

async function getSubscriptionStatus() {
  const response = await fetch(`${API_BASE_URL}/subscription`, {
    headers: getAuthHeaders()
  });
  
  const subscription = await response.json();
  
  updateSubscriptionUI({
    status: subscription.status,
    hasActive: subscription.hasActiveSubscription,
    credits: subscription.credits,
    nextBilling: subscription.stripeSubscription?.currentPeriodEnd
  });
  
  return subscription;
}

async function cancelSubscription() {
  if (!confirm('Are you sure you want to cancel your subscription?')) return;
  
  const response = await fetch(`${API_BASE_URL}/cancel-subscription`, {
    method: 'POST',
    headers: getAuthHeaders()
  });
  
  const result = await response.json();
  if (result.success) {
    showSuccess(`Subscription cancelled. Will end on ${new Date(result.cancelAt).toLocaleDateString()}`);
    await getSubscriptionStatus(); // Refresh UI
  }
}
```

### **5. Token Refresh Management**

**Handle Expired Tokens:**
```javascript
async function refreshToken() {
  const refreshToken = localStorage.getItem('refresh_token');
  if (!refreshToken) {
    redirectToLogin();
    return;
  }
  
  try {
    const response = await fetch(`${API_BASE_URL}/auth/refresh`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ refresh_token: refreshToken })
    });
    
    const data = await response.json();
    if (data.access_token) {
      localStorage.setItem('access_token', data.access_token);
      return data.access_token;
    }
  } catch (error) {
    console.error('Token refresh failed:', error);
  }
  
  redirectToLogin();
}
```

### **6. Usage Analytics**

**Track User Activity:**
```javascript
async function getUsageAnalytics() {
  const response = await fetch(`${API_BASE_URL}/usage`, {
    headers: getAuthHeaders()
  });
  
  const usage = await response.json();
  
  displayUsageStats({
    recentActivity: usage.recentActivity,
    totalBuilds: usage.summary.build || 0,
    totalUpdates: usage.summary.update || 0,
    totalActions: usage.totalActions
  });
}
```

### **7. Error Handling**

**Implement Comprehensive Error Handling:**
```javascript
async function makeAPICall(url, options = {}) {
  try {
    const response = await fetch(url, {
      ...options,
      headers: { ...getAuthHeaders(), ...options.headers }
    });
    
    if (response.status === 401) {
      await refreshToken();
      // Retry once with new token
      return fetch(url, {
        ...options,
        headers: { ...getAuthHeaders(), ...options.headers }
      });
    }
    
    if (response.status === 402) {
      // Insufficient credits
      const error = await response.json();
      showSubscriptionPrompt(error.credits);
      throw new Error('Insufficient credits');
    }
    
    if (!response.ok) {
      const error = await response.json();
      throw new Error(error.error || 'API call failed');
    }
    
    return response;
  } catch (error) {
    console.error('API Error:', error);
    throw error;
  }
}
```

### **8. UI Components to Add**

**Required UI Elements:**
1. **Sign-in Button** - Google OAuth integration
2. **User Profile Section** - Name, email, credits display
3. **Credits Counter** - Real-time credit balance
4. **Subscription Status Badge** - trial/active/cancelled
5. **Upgrade Button** - Link to Stripe checkout
6. **Usage Statistics** - Show user activity
7. **Settings Panel** - Manage subscription, view usage

### **9. App Initialization Flow**

**On App Startup:**
```javascript
async function initializeApp() {
  const token = localStorage.getItem('access_token');
  
  if (!token) {
    showSignInScreen();
    return;
  }
  
  try {
    const user = await getUserProfile();
    const subscription = await getSubscriptionStatus();
    
    showMainApp();
    updateUserInterface(user, subscription);
    
    // Check for low credits
    if (user.credits < 2 && !user.hasActiveSubscription) {
      showLowCreditsWarning();
    }
    
  } catch (error) {
    console.error('App initialization failed:', error);
    showSignInScreen();
  }
}

// Call on app load
initializeApp();
```

### **10. Feature Gating**

**Control Access Based on Credits/Subscription:**
```javascript
function canUseFeature(featureType) {
  const user = getCurrentUser();
  return user.canUseFeatures || user.hasActiveSubscription;
}

function enforceFeatureAccess(featureType, callback) {
  if (canUseFeature(featureType)) {
    callback();
  } else {
    showSubscriptionPrompt();
  }
}

// Usage
document.getElementById('buildButton').onclick = () => {
  enforceFeatureAccess('build', buildModel);
};
```

## 🔐 **Security Notes**

1. **Never expose sensitive environment variables** in frontend code
2. **Store tokens securely** in localStorage or sessionStorage
3. **Always handle 401 responses** with token refresh
4. **Validate user input** before sending to backend
5. **Use HTTPS** for all API calls (already configured)

## 🎯 **Testing Checklist**

After implementation, test:
- [ ] Google sign-in flow
- [ ] User profile loading
- [ ] Credit consumption on actions
- [ ] Subscription checkout process
- [ ] Token refresh on expiry
- [ ] Error handling for insufficient credits
- [ ] Subscription cancellation
- [ ] Usage analytics display

## 📊 **Backend API Endpoints**

### Authentication
- `POST /auth/google` - Sign in with Google
- `POST /auth/refresh` - Refresh token

### User Management  
- `GET /me` - Get user profile and credits

### Credits
- `POST /use-credit` - Consume credit for actions

### Subscriptions
- `POST /create-checkout-session` - Start Stripe checkout
- `GET /subscription` - Get subscription status
- `POST /cancel-subscription` - Cancel subscription

### Analytics
- `GET /usage` - Get usage statistics

### Health
- `GET /health` - Health check

This integration will give your Excel addin a complete authentication, subscription, and credit management system! 🚀