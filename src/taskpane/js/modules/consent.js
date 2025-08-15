/**
 * Consent Module
 * Handles data sharing consent for EBITDAI
 */

export async function initialize() {
    console.log('Initializing consent module');
    
    // Set up event listeners
    setupConsentButtons();
    
    // Check if consent was already given
    checkExistingConsent();
}

/**
 * Set up consent button event listeners
 */
function setupConsentButtons() {
    const acceptButton = document.getElementById('consent-accept');
    const declineButton = document.getElementById('consent-decline');
    
    if (acceptButton) {
        acceptButton.addEventListener('click', () => handleConsentResponse(true));
    }
    
    if (declineButton) {
        declineButton.addEventListener('click', () => handleConsentResponse(false));
    }
}

/**
 * Handle consent response
 * @param {boolean} consented - Whether user consented to data sharing
 */
async function handleConsentResponse(consented) {
    try {
        showLoading();
        
        // Store consent decision with timestamp
        const consentData = {
            consented: consented,
            timestamp: new Date().toISOString(),
            version: '1.0' // Track consent version for future updates
        };
        
        localStorage.setItem('data_sharing_consent', JSON.stringify(consentData));
        
        // Log the consent decision (for analytics)
        console.log(`User ${consented ? 'accepted' : 'declined'} data sharing consent`);
        
        // Show appropriate feedback
        if (consented) {
            showConsentSuccess();
        } else {
            showConsentDeclined();
        }
        
        // Proceed to main application after a short delay
        setTimeout(() => {
            window.app.loadView('client-mode');
        }, consented ? 1500 : 2000);
        
    } catch (error) {
        console.error('Error handling consent response:', error);
        showError('Failed to save consent preference. Please try again.');
        hideLoading();
    }
}

/**
 * Check if user has already given consent
 */
function checkExistingConsent() {
    const storedConsent = localStorage.getItem('data_sharing_consent');
    
    if (storedConsent) {
        try {
            const consentData = JSON.parse(storedConsent);
            
            // Check if consent is still valid (e.g., within last 365 days)
            const consentDate = new Date(consentData.timestamp);
            const daysSinceConsent = (new Date() - consentDate) / (1000 * 60 * 60 * 24);
            
            if (daysSinceConsent <= 365 && consentData.version === '1.0') {
                // Consent is still valid, skip to main app
                console.log('Valid consent found, proceeding to main application');
                setTimeout(() => {
                    window.app.loadView('client-mode');
                }, 500);
                return;
            } else {
                // Consent expired or version changed, need new consent
                console.log('Consent expired or version changed, requesting new consent');
                localStorage.removeItem('data_sharing_consent');
            }
        } catch (error) {
            console.error('Error parsing stored consent:', error);
            localStorage.removeItem('data_sharing_consent');
        }
    }
}

/**
 * Get current consent status
 * @returns {Object|null} Consent data or null if no consent given
 */
export function getConsentStatus() {
    const storedConsent = localStorage.getItem('data_sharing_consent');
    
    if (storedConsent) {
        try {
            return JSON.parse(storedConsent);
        } catch (error) {
            console.error('Error parsing consent data:', error);
            return null;
        }
    }
    
    return null;
}

/**
 * Check if user has consented to data sharing
 * @returns {boolean} True if user has consented and consent is valid
 */
export function hasValidConsent() {
    const consentData = getConsentStatus();
    
    if (!consentData) return false;
    
    // Check if consent is still valid
    const consentDate = new Date(consentData.timestamp);
    const daysSinceConsent = (new Date() - consentDate) / (1000 * 60 * 60 * 24);
    
    return consentData.consented && daysSinceConsent <= 365 && consentData.version === '1.0';
}

/**
 * Show loading state
 */
function showLoading() {
    const container = document.querySelector('.consent-container');
    if (container) {
        const loading = document.createElement('div');
        loading.className = 'consent-loading';
        loading.innerHTML = '<div class="consent-spinner"></div>';
        container.appendChild(loading);
    }
}

/**
 * Hide loading state
 */
function hideLoading() {
    const loading = document.querySelector('.consent-loading');
    if (loading) {
        loading.remove();
    }
}

/**
 * Show consent acceptance success
 */
function showConsentSuccess() {
    const container = document.querySelector('.consent-container');
    if (container) {
        container.innerHTML = `
            <div class="consent-success">
                <div class="consent-header">
                    <img src="../../assets/logo-filled.png" alt="EBITDAI" class="consent-logo" />
                    <h1 class="consent-brand">EBITDAI</h1>
                </div>
                <div style="padding: 40px; text-align: center;">
                    <svg style="width: 64px; height: 64px; color: #059669; margin-bottom: 20px;" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                        <path d="M20 6L9 17l-5-5" stroke-linecap="round" stroke-linejoin="round"/>
                    </svg>
                    <h2 style="margin: 0 0 10px; font-size: 24px; color: #111827;">Thank You!</h2>
                    <p style="margin: 0; color: #6b7280; font-size: 16px;">Your consent helps us improve our AI models and provide better financial modeling tools.</p>
                </div>
            </div>
        `;
    }
}

/**
 * Show consent declined message
 */
function showConsentDeclined() {
    const container = document.querySelector('.consent-container');
    if (container) {
        container.innerHTML = `
            <div class="consent-declined">
                <div class="consent-header">
                    <img src="../../assets/logo-filled.png" alt="EBITDAI" class="consent-logo" />
                    <h1 class="consent-brand">EBITDAI</h1>
                </div>
                <div style="padding: 40px; text-align: center;">
                    <svg style="width: 64px; height: 64px; color: #f59e0b; margin-bottom: 20px;" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                        <path d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-2.5L13.732 4c-.77-.833-1.964-.833-2.732 0L3.732 16.5c-.77.833.192 2.5 1.732 2.5z"/>
                    </svg>
                    <h2 style="margin: 0 0 10px; font-size: 24px; color: #111827;">No Problem</h2>
                    <p style="margin: 0 0 15px; color: #6b7280; font-size: 16px;">You can still use EBITDAI. Your data will not be shared with partners.</p>
                    <p style="margin: 0; color: #9ca3af; font-size: 14px;">You can change this preference anytime in settings.</p>
                </div>
            </div>
        `;
    }
}

/**
 * Show error message
 */
function showError(message) {
    const container = document.querySelector('.consent-content');
    if (container) {
        let error = container.querySelector('.consent-error');
        if (!error) {
            error = document.createElement('div');
            error.className = 'consent-error';
            error.style.cssText = `
                background: #fef2f2;
                border: 1px solid #fecaca;
                border-radius: 8px;
                padding: 12px 16px;
                margin-bottom: 20px;
                color: #dc2626;
                font-size: 14px;
            `;
            container.insertBefore(error, container.firstChild);
        }
        error.textContent = message;
        
        // Hide after 5 seconds
        setTimeout(() => {
            error.remove();
        }, 5000);
    }
}