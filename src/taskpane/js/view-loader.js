/**
 * ViewLoader - Handles dynamic loading of HTML views and their associated CSS
 */
class ViewLoader {
    constructor() {
        this.currentView = null;
        this.loadedStyles = new Set();
    }

    /**
     * Load a view and its associated styles
     * @param {string} viewName - Name of the view to load
     * @returns {Promise<void>}
     */
    async loadView(viewName) {
        try {
            // Unload previous view's specific styles
            if (this.currentView && this.currentView !== viewName) {
                this.unloadStyles(this.currentView);
            }

            // Load HTML content
            const htmlPath = `views/${viewName}.html`;
            const response = await fetch(htmlPath);
            const html = await response.text();

            // Update the main container
            const mainContainer = document.getElementById('main-container');
            mainContainer.innerHTML = html;

            // Load view-specific CSS
            await this.loadStyles(viewName);

            // Load view-specific JavaScript module
            await this.loadModule(viewName);

            this.currentView = viewName;
            console.log(`Loaded view: ${viewName}`);
        } catch (error) {
            console.error(`Error loading view ${viewName}:`, error);
        }
    }

    /**
     * Load CSS file for a specific view
     * @param {string} viewName - Name of the view
     * @returns {Promise<void>}
     */
    async loadStyles(viewName) {
        const styleId = `${viewName}-styles`;
        
        // Skip if already loaded
        if (this.loadedStyles.has(styleId)) {
            return;
        }

        return new Promise((resolve, reject) => {
            const link = document.createElement('link');
            link.id = styleId;
            link.rel = 'stylesheet';
            link.href = `styles/${viewName}.css`;
            link.onload = () => {
                this.loadedStyles.add(styleId);
                resolve();
            };
            link.onerror = reject;
            document.head.appendChild(link);
        });
    }

    /**
     * Unload CSS file for a specific view
     * @param {string} viewName - Name of the view
     */
    unloadStyles(viewName) {
        const styleId = `${viewName}-styles`;
        const styleElement = document.getElementById(styleId);
        
        if (styleElement) {
            styleElement.remove();
            this.loadedStyles.delete(styleId);
        }
    }

    /**
     * Load JavaScript module for a specific view
     * @param {string} viewName - Name of the view
     * @returns {Promise<void>}
     */
    async loadModule(viewName) {
        try {
            const module = await import(`./modules/${viewName}.js`);
            if (module.initialize) {
                await module.initialize();
            }
        } catch (error) {
            console.log(`No JavaScript module found for ${viewName} or error loading:`, error);
        }
    }
}

// Export for use in other modules
window.ViewLoader = ViewLoader;
