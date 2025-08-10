/**
 * Login Module
 * Handles authentication for EBITDAI
 */

export async function initialize() {
    console.log('Initializing login module');
    
    // Set up event listeners
    setupGoogleSignIn();
    setupMicrosoftSignIn();
    setupManualSignIn();
    
    // Check if already signing in (redirect from OAuth)
    checkOAuthCallback();
}

/**
 * Set up Google Sign In
 */
function setupGoogleSignIn() {
    const googleButton = document.getElementById('google-signin-button');
    if (googleButton) {
        googleButton.addEventListener('click', async () => {
            try {
                showLoading();
                
                // Trigger Google OAuth flow
                const clientId = '169155377864-2vru0t1ohv6afpi7nvr1lq7ri3vqt9qp.apps.googleusercontent.com';
                const redirectUri = window.location.origin + '/auth/google/callback.html';
                const scope = 'openid profile email';
                
                const authUrl = `https://accounts.google.com/o/oauth2/v2/auth?` +
                    `client_id=${clientId}` +
                    `&redirect_uri=${encodeURIComponent(redirectUri)}` +
                    `&response_type=token` +
                    `&scope=${encodeURIComponent(scope)}`;
                
                // Store return URL
                localStorage.setItem('auth_return_view', 'client-mode');
                
                // Redirect to Google auth
                window.location.href = authUrl;
                
            } catch (error) {
                console.error('Google sign in error:', error);
                showError('Failed to sign in with Google. Please try again.');
                hideLoading();
            }
        });
    }
}

/**
 * Set up Microsoft Sign In
 */
function setupMicrosoftSignIn() {
    const microsoftButton = document.getElementById('microsoft-signin-button');
    if (microsoftButton) {
        microsoftButton.addEventListener('click', async () => {
            try {
                showLoading();
                
                // Check if MSAL is loaded
                if (typeof msal === 'undefined') {
                    throw new Error('Microsoft authentication library not loaded');
                }
                
                // Initialize MSAL
                const msalConfig = {
                    auth: {
                        clientId: 'fa14d5bf-6c7f-443e-958f-83ac94e1b0e7',
                        authority: 'https://login.microsoftonline.com/common',
                        redirectUri: window.location.origin
                    }
                };
                
                const msalInstance = new msal.PublicClientApplication(msalConfig);
                
                // Login request
                const loginRequest = {
                    scopes: ['user.read', 'openid', 'profile', 'email']
                };
                
                // Try popup first, fall back to redirect
                try {
                    const response = await msalInstance.loginPopup(loginRequest);
                    handleMicrosoftAuthSuccess(response);
                } catch (popupError) {
                    console.log('Popup blocked, using redirect');
                    localStorage.setItem('auth_return_view', 'client-mode');
                    msalInstance.loginRedirect(loginRequest);
                }
                
            } catch (error) {
                console.error('Microsoft sign in error:', error);
                showError('Failed to sign in with Microsoft. Please try again.');
                hideLoading();
            }
        });
    }
}

/**
 * Set up Manual Sign In (API Key)
 */
function setupManualSignIn() {
    const manualButton = document.getElementById('manual-signin-button');
    if (manualButton) {
        manualButton.addEventListener('click', () => {
            showApiKeyDialog();
        });
    }
}

/**
 * Show API Key input dialog
 */
function showApiKeyDialog() {
    const dialog = document.createElement('div');
    dialog.className = 'api-key-dialog';
    dialog.innerHTML = `
        <div class="api-key-content">
            <h3>Enter Your API Key</h3>
            <p>Enter your OpenAI API key to use EBITDAI</p>
            <input type="password" id="api-key-input" class="form-input" placeholder="sk-..." />
            <div class="api-key-actions">
                <button id="api-key-cancel" class="btn btn-secondary">Cancel</button>
                <button id="api-key-submit" class="btn btn-primary">Sign In</button>
            </div>
        </div>
    `;
    
    document.body.appendChild(dialog);
    
    // Focus input
    const input = document.getElementById('api-key-input');
    input.focus();
    
    // Handle cancel
    document.getElementById('api-key-cancel').addEventListener('click', () => {
        dialog.remove();
    });
    
    // Handle submit
    document.getElementById('api-key-submit').addEventListener('click', () => {
        const apiKey = input.value.trim();
        if (apiKey) {
            handleApiKeySignIn(apiKey);
            dialog.remove();
        }
    });
    
    // Handle enter key
    input.addEventListener('keypress', (e) => {
        if (e.key === 'Enter') {
            const apiKey = input.value.trim();
            if (apiKey) {
                handleApiKeySignIn(apiKey);
                dialog.remove();
            }
        }
    });
}

/**
 * Handle API Key sign in
 */
async function handleApiKeySignIn(apiKey) {
    try {
        showLoading();
        
        // Validate API key format
        if (!apiKey.startsWith('sk-')) {
            throw new Error('Invalid API key format');
        }
        
        // Store API key
        localStorage.setItem('user_api_key', apiKey);
        localStorage.setItem('auth_method', 'api_key');
        
        // Create mock user data for API key users
        const userData = {
            name: 'API Key User',
            email: 'api@user.local',
            picture: null,
            credits: 100,
            subscription: 'Free'
        };
        
        localStorage.setItem('user_data', JSON.stringify(userData));
        
        // Show success and redirect
        showSuccess();
        setTimeout(() => {
            window.app.loadView('client-mode');
        }, 1500);
        
    } catch (error) {
        console.error('API key sign in error:', error);
        showError('Invalid API key. Please check and try again.');
        hideLoading();
    }
}

/**
 * Handle Microsoft auth success
 */
function handleMicrosoftAuthSuccess(response) {
    try {
        // Store tokens
        localStorage.setItem('msal_access_token', response.accessToken);
        localStorage.setItem('auth_method', 'microsoft');
        
        // Store user data
        const userData = {
            name: response.account.name,
            email: response.account.username,
            picture: null,
            credits: 100,
            subscription: 'Free'
        };
        
        localStorage.setItem('user_data', JSON.stringify(userData));
        
        // Show success and redirect
        showSuccess();
        setTimeout(() => {
            window.app.loadView('client-mode');
        }, 1500);
        
    } catch (error) {
        console.error('Error handling Microsoft auth:', error);
        showError('Authentication failed. Please try again.');
    }
}

/**
 * Check for OAuth callback
 */
function checkOAuthCallback() {
    // Check for Google OAuth token in URL
    const hash = window.location.hash;
    if (hash && hash.includes('access_token')) {
        const params = new URLSearchParams(hash.substring(1));
        const accessToken = params.get('access_token');
        
        if (accessToken) {
            handleGoogleAuthSuccess(accessToken);
        }
    }
}

/**
 * Handle Google auth success
 */
async function handleGoogleAuthSuccess(accessToken) {
    try {
        showLoading();
        
        // Store token
        localStorage.setItem('google_access_token', accessToken);
        localStorage.setItem('auth_method', 'google');
        
        // Get user info
        const response = await fetch('https://www.googleapis.com/oauth2/v2/userinfo', {
            headers: {
                'Authorization': `Bearer ${accessToken}`
            }
        });
        
        const userInfo = await response.json();
        
        // Store user data
        const userData = {
            name: userInfo.name,
            email: userInfo.email,
            picture: userInfo.picture,
            credits: 100,
            subscription: 'Free'
        };
        
        localStorage.setItem('user_data', JSON.stringify(userData));
        
        // Show success and redirect
        showSuccess();
        setTimeout(() => {
            const returnView = localStorage.getItem('auth_return_view') || 'client-mode';
            localStorage.removeItem('auth_return_view');
            window.app.loadView(returnView);
        }, 1500);
        
    } catch (error) {
        console.error('Error handling Google auth:', error);
        showError('Authentication failed. Please try again.');
        hideLoading();
    }
}

/**
 * Show loading state
 */
function showLoading() {
    const container = document.querySelector('.login-container');
    if (container) {
        const loading = document.createElement('div');
        loading.className = 'login-loading';
        loading.innerHTML = '<div class="login-spinner"></div>';
        container.appendChild(loading);
    }
}

/**
 * Hide loading state
 */
function hideLoading() {
    const loading = document.querySelector('.login-loading');
    if (loading) {
        loading.remove();
    }
}

/**
 * Show error message
 */
function showError(message) {
    const container = document.querySelector('.login-content');
    if (container) {
        let error = container.querySelector('.login-error');
        if (!error) {
            error = document.createElement('div');
            error.className = 'login-error';
            container.insertBefore(error, container.firstChild);
        }
        error.textContent = message;
        error.classList.add('show');
        
        // Hide after 5 seconds
        setTimeout(() => {
            error.classList.remove('show');
        }, 5000);
    }
}

/**
 * Show success state
 */
function showSuccess() {
    const container = document.querySelector('.login-container');
    if (container) {
        container.innerHTML = `
            <div class="login-success">
                <svg class="login-success-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                    <path d="M20 6L9 17l-5-5" stroke-linecap="round" stroke-linejoin="round"/>
                </svg>
                <h2>Welcome to EBITDAI!</h2>
                <p>Redirecting to chat...</p>
            </div>
        `;
    }
}
