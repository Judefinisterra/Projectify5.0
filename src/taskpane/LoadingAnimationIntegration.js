/**
 * LoadingAnimationIntegration.js
 * 
 * This file demonstrates how to integrate the LoadingAnimation module
 * with the existing taskpane.js code, replacing the current implementation
 * with the reusable module.
 */

import LoadingAnimation from './LoadingAnimation.js';

// Create global loader instances for different modes
let devModeLoader = null;
let clientModeLoader = null;

/**
 * Initialize the loading animations for both developer and client modes
 */
export function initializeLoadingAnimations() {
    // Developer mode loader
    devModeLoader = new LoadingAnimation({
        containerId: 'loading-animation',
        showText: false,
        position: 'absolute',
        spinnerSize: '2rem',
        spinnerBorderWidth: '2px',
        spinnerColor: '#1f2937',
        spinnerTrackColor: '#e5e7eb'
    });

    // Client mode loader
    clientModeLoader = new LoadingAnimation({
        containerId: 'loading-animation-client',
        showText: false,
        spinnerSize: '32px',
        spinnerBorderWidth: '3px',
        spinnerColor: '#000000',
        spinnerTrackColor: '#f0f0f0',
        animationDuration: '0.8s'
    });

    // Initialize loaders if their containers exist
    const devContainer = document.getElementById('chat-log-container');
    if (devContainer) {
        devModeLoader.init(devContainer);
    }

    const clientContainer = document.getElementById('client-chat-container');
    if (clientContainer) {
        clientModeLoader.init(clientContainer);
    }
}

/**
 * Replacement for the existing setButtonLoading function
 * @param {boolean} isLoading - Whether to show or hide the loading animation
 */
export function setButtonLoading(isLoading) {
    console.log(`[setButtonLoading] Called with isLoading: ${isLoading}`);
    
    const sendButton = document.getElementById('send');
    if (sendButton) {
        sendButton.disabled = isLoading;
    } else {
        console.warn("[setButtonLoading] Could not find send button with id='send'");
    }
    
    if (devModeLoader) {
        if (isLoading) {
            devModeLoader.show();
        } else {
            devModeLoader.hide();
        }
    } else {
        console.error("[setButtonLoading] Developer mode loader not initialized");
    }
}

/**
 * Replacement for the existing setButtonLoadingClient function
 * @param {boolean} isLoading - Whether to show or hide the loading animation
 */
export function setButtonLoadingClient(isLoading) {
    console.log(`[setButtonLoadingClient] Called with isLoading: ${isLoading}`);
    
    const sendButton = document.getElementById('send-client');
    if (sendButton) {
        sendButton.disabled = isLoading;
    } else {
        console.warn("[setButtonLoadingClient] Could not find send button with id='send-client'");
    }
    
    if (clientModeLoader) {
        if (isLoading) {
            clientModeLoader.show();
        } else {
            clientModeLoader.hide();
        }
    } else {
        console.error("[setButtonLoadingClient] Client mode loader not initialized");
    }
}

/**
 * Create a loading animation for voice transcription
 * @returns {LoadingAnimation} The transcription loader instance
 */
export function createTranscriptionLoader() {
    const loader = new LoadingAnimation({
        spinnerSize: '24px',
        spinnerBorderWidth: '3px',
        spinnerColor: '#3b82f6',
        spinnerTrackColor: '#e5e7eb',
        animationDuration: '0.8s',
        loadingText: 'Transcribing...',
        textColor: '#6b7280',
        textSize: '14px',
        className: 'loading-animation inline'
    });
    
    return loader;
}

/**
 * Create a loading animation for file processing
 * @param {string} fileName - The name of the file being processed
 * @returns {LoadingAnimation} The file processing loader instance
 */
export function createFileProcessingLoader(fileName) {
    const loader = new LoadingAnimation({
        loadingText: `Processing ${fileName}...`,
        spinnerColor: '#10b981',
        textColor: '#065f46',
        className: 'loading-animation success'
    });
    
    return loader;
}

/**
 * Create a full-page loading overlay
 * @param {string} message - The loading message to display
 * @returns {Object} Object containing the overlay element and loader instance
 */
export function createLoadingOverlay(message = 'Please wait...') {
    // Create overlay container
    const overlay = document.createElement('div');
    overlay.className = 'loading-animation-overlay';
    overlay.style.cssText = `
        position: fixed;
        top: 0;
        left: 0;
        width: 100%;
        height: 100%;
        background-color: rgba(0, 0, 0, 0.5);
        backdrop-filter: blur(4px);
        display: flex;
        align-items: center;
        justify-content: center;
        z-index: 9999;
    `;
    
    // Create inner container for the loader
    const loaderContainer = document.createElement('div');
    loaderContainer.style.cssText = `
        background-color: white;
        padding: 2rem;
        border-radius: 8px;
        box-shadow: 0 10px 25px rgba(0, 0, 0, 0.2);
    `;
    overlay.appendChild(loaderContainer);
    
    // Create loader
    const loader = new LoadingAnimation({
        loadingText: message,
        spinnerSize: '48px',
        position: 'relative'
    });
    
    // Add to DOM
    document.body.appendChild(overlay);
    loader.init(loaderContainer);
    loader.show();
    
    return {
        overlay,
        loader,
        remove() {
            loader.destroy();
            overlay.remove();
        }
    };
}

/**
 * Wrap an async function with loading animation
 * @param {Function} asyncFn - The async function to execute
 * @param {Object} options - Options for the loading animation
 * @returns {Promise} The result of the async function
 */
export async function withLoading(asyncFn, options = {}) {
    const {
        container = '#app-body',
        message = 'Loading...',
        onError = null
    } = options;
    
    const loader = LoadingAnimation.show(container, {
        loadingText: message,
        position: 'absolute'
    });
    
    try {
        const result = await asyncFn();
        loader.hide();
        return result;
    } catch (error) {
        // Update loader to show error state
        loader.updateConfig({
            spinnerColor: '#dc2626',
            loadingText: onError || 'An error occurred',
            className: 'loading-animation error'
        });
        
        // Hide after showing error
        setTimeout(() => loader.hide(), 2000);
        throw error;
    }
}

/**
 * Migration guide for updating taskpane.js
 * 
 * 1. Import the integration module at the top of taskpane.js:
 *    import { initializeLoadingAnimations, setButtonLoading, setButtonLoadingClient } from './LoadingAnimationIntegration.js';
 * 
 * 2. Initialize the loaders after DOM is ready (in Office.onReady or DOMContentLoaded):
 *    initializeLoadingAnimations();
 * 
 * 3. Remove the existing setButtonLoading and setButtonLoadingClient functions
 * 
 * 4. The new functions will be available globally and work exactly the same way
 * 
 * 5. For new loading animations, use the LoadingAnimation class directly:
 *    const customLoader = new LoadingAnimation({ ... });
 *    customLoader.init('#my-container');
 *    customLoader.show();
 */

// Example of how to update the existing code in taskpane.js:
/*
// OLD CODE (Remove this):
function setButtonLoading(isLoading) {
    console.log(`[setButtonLoading] Called with isLoading: ${isLoading}`);
    const sendButton = document.getElementById('send');
    const loadingAnimation = document.getElementById('loading-animation');
    
    if (sendButton) {
        sendButton.disabled = isLoading;
    } else {
        console.warn("[setButtonLoading] Could not find send button with id='send'");
    }
    
    if (loadingAnimation) {
        const newDisplay = isLoading ? 'flex' : 'none';
        console.log(`[setButtonLoading] Found loadingAnimation element. Setting display to: ${newDisplay}`);
        loadingAnimation.style.display = newDisplay;
    } else {
        console.error("[setButtonLoading] Could not find loading animation element with id='loading-animation'");
    }
}

// NEW CODE (Add this to imports):
import { initializeLoadingAnimations, setButtonLoading, setButtonLoadingClient } from './LoadingAnimationIntegration.js';

// In Office.onReady or when DOM is ready:
initializeLoadingAnimations();

// The setButtonLoading function is now imported and will work the same way
*/

export {
    devModeLoader,
    clientModeLoader
};


