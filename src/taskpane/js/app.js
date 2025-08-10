/**
 * Main Application Controller
 * Handles view management and Office.js initialization
 */

// Global app object
window.app = {
    viewLoader: null,
    currentView: null,
    isOfficeInitialized: false
};

// Initialize Office
Office.onReady(async (info) => {
    console.log('Office.js is ready');
    app.isOfficeInitialized = true;
    
    // Initialize view loader
    app.viewLoader = new ViewLoader();
    
    // Determine initial view based on context
    const initialView = determineInitialView();
    
    // Load the initial view
    await app.loadView(initialView);
});

/**
 * Determine which view to load initially
 * @returns {string} View name to load
 */
function determineInitialView() {
    // Check URL parameters
    const urlParams = new URLSearchParams(window.location.search);
    const viewParam = urlParams.get('view');
    
    if (viewParam) {
        return viewParam;
    }
    
    // Check local storage for last used view
    const lastView = localStorage.getItem('lastView');
    if (lastView && lastView !== 'startup-menu') {
        return lastView;
    }
    
    // Check if user is authenticated
    const isAuthenticated = checkAuthentication();
    if (!isAuthenticated) {
        return 'authentication';
    }
    
    // Default to startup menu
    return 'startup-menu';
}

/**
 * Check if user is authenticated
 * @returns {boolean}
 */
function checkAuthentication() {
    // Check for stored auth tokens
    const googleToken = localStorage.getItem('google_access_token');
    const msalToken = localStorage.getItem('msal_access_token');
    
    return !!(googleToken || msalToken);
}

/**
 * Load a specific view
 * @param {string} viewName - Name of the view to load
 */
app.loadView = async function(viewName) {
    try {
        // Show loading state
        showLoadingState();
        
        // Load the view
        await app.viewLoader.loadView(viewName);
        
        // Update current view
        app.currentView = viewName;
        
        // Save to local storage
        localStorage.setItem('lastView', viewName);
        
        // Hide loading state
        hideLoadingState();
        
        // Fire view loaded event
        fireViewLoadedEvent(viewName);
        
    } catch (error) {
        console.error('Error loading view:', error);
        showErrorState(error.message);
    }
};

/**
 * Show loading state
 */
function showLoadingState() {
    const container = document.getElementById('main-container');
    container.classList.add('loading');
    
    // Optionally show a loading spinner
    const loadingHtml = `
        <div class="loading-overlay">
            <div class="spinner"></div>
            <p>Loading...</p>
        </div>
    `;
    
    // Only show spinner if loading takes more than 200ms
    app.loadingTimeout = setTimeout(() => {
        if (!document.querySelector('.loading-overlay')) {
            container.insertAdjacentHTML('beforeend', loadingHtml);
        }
    }, 200);
}

/**
 * Hide loading state
 */
function hideLoadingState() {
    clearTimeout(app.loadingTimeout);
    
    const container = document.getElementById('main-container');
    container.classList.remove('loading');
    
    const loadingOverlay = document.querySelector('.loading-overlay');
    if (loadingOverlay) {
        loadingOverlay.remove();
    }
}

/**
 * Show error state
 * @param {string} message - Error message to display
 */
function showErrorState(message) {
    const container = document.getElementById('main-container');
    container.innerHTML = `
        <div class="error-state">
            <h2>Error Loading View</h2>
            <p>${message}</p>
            <button onclick="app.loadView('startup-menu')">Return to Home</button>
        </div>
    `;
}

/**
 * Fire custom event when view is loaded
 * @param {string} viewName - Name of the loaded view
 */
function fireViewLoadedEvent(viewName) {
    const event = new CustomEvent('viewLoaded', {
        detail: { viewName }
    });
    document.dispatchEvent(event);
}

// Global navigation functions
app.goBack = function() {
    app.loadView('startup-menu');
};

app.signOut = function() {
    // Clear authentication tokens
    localStorage.removeItem('google_access_token');
    localStorage.removeItem('msal_access_token');
    localStorage.removeItem('user_data');
    
    // Load authentication view
    app.loadView('authentication');
};

// Export for use in modules
window.app = app;
