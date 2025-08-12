/**
 * LoadingAnimation.js - A reusable loading animation module
 * 
 * This module provides a configurable loading animation that can be easily
 * imported and used throughout the application.
 */

class LoadingAnimation {
    constructor(options = {}) {
        // Default configuration
        this.config = {
            containerId: null, // ID of existing container or null to create new
            spinnerSize: '32px',
            spinnerBorderWidth: '3px',
            spinnerColor: '#000000',
            spinnerTrackColor: '#f0f0f0',
            animationDuration: '0.8s',
            showText: true,
            loadingText: 'Loading...',
            textColor: '#6b7280',
            textSize: '14px',
            className: 'loading-animation-container',
            position: 'relative', // 'relative', 'absolute', 'fixed'
            zIndex: 1000,
            ...options
        };

        this.container = null;
        this.styleElement = null;
        this.isShowing = false;
    }

    /**
     * Initialize the loading animation
     * @param {HTMLElement|string} target - Target element or selector where the animation will be inserted
     */
    init(target) {
        // Get target element
        this.targetElement = typeof target === 'string' 
            ? document.querySelector(target) 
            : target;

        if (!this.targetElement) {
            throw new Error('LoadingAnimation: Target element not found');
        }

        // Create or get container
        if (this.config.containerId) {
            this.container = document.getElementById(this.config.containerId);
        }

        if (!this.container) {
            this.container = this.createContainer();
            this.targetElement.appendChild(this.container);
        }

        // Inject styles if not already present
        this.injectStyles();

        // Create animation content
        this.createAnimationContent();

        // Initially hide the animation
        this.hide();
    }

    /**
     * Create the container element
     */
    createContainer() {
        const container = document.createElement('div');
        container.className = this.config.className;
        container.style.cssText = `
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            gap: ${this.config.showText ? '12px' : '0'};
            position: ${this.config.position};
            z-index: ${this.config.zIndex};
        `;
        return container;
    }

    /**
     * Create the animation content (spinner and optional text)
     */
    createAnimationContent() {
        // Clear any existing content
        this.container.innerHTML = '';

        // Create spinner
        const spinner = document.createElement('div');
        spinner.className = 'loading-spinner';
        spinner.style.cssText = `
            width: ${this.config.spinnerSize};
            height: ${this.config.spinnerSize};
            border: ${this.config.spinnerBorderWidth} solid ${this.config.spinnerTrackColor};
            border-top-color: ${this.config.spinnerColor};
            border-radius: 50%;
            animation: loading-spin ${this.config.animationDuration} linear infinite;
        `;
        this.container.appendChild(spinner);

        // Create loading text if enabled
        if (this.config.showText) {
            const text = document.createElement('span');
            text.className = 'loading-text';
            text.textContent = this.config.loadingText;
            text.style.cssText = `
                color: ${this.config.textColor};
                font-size: ${this.config.textSize};
                font-weight: 500;
            `;
            this.container.appendChild(text);
        }
    }

    /**
     * Inject required CSS styles
     */
    injectStyles() {
        const styleId = 'loading-animation-styles';
        
        // Check if styles already exist
        if (document.getElementById(styleId)) {
            return;
        }

        // Create and inject style element
        this.styleElement = document.createElement('style');
        this.styleElement.id = styleId;
        this.styleElement.textContent = `
            @keyframes loading-spin {
                0% { 
                    transform: rotate(0deg); 
                }
                100% { 
                    transform: rotate(360deg); 
                }
            }

            .${this.config.className} {
                transition: opacity 0.3s ease-in-out;
            }

            .${this.config.className}.loading-hidden {
                opacity: 0;
                pointer-events: none;
            }

            .${this.config.className}.loading-visible {
                opacity: 1;
            }
        `;
        document.head.appendChild(this.styleElement);
    }

    /**
     * Show the loading animation
     */
    show() {
        if (!this.container) {
            console.error('LoadingAnimation: Not initialized. Call init() first.');
            return;
        }

        this.container.style.display = 'flex';
        // Force reflow to ensure transition works
        this.container.offsetHeight;
        this.container.classList.remove('loading-hidden');
        this.container.classList.add('loading-visible');
        this.isShowing = true;
    }

    /**
     * Hide the loading animation
     */
    hide() {
        if (!this.container) {
            return;
        }

        this.container.classList.remove('loading-visible');
        this.container.classList.add('loading-hidden');
        
        // Hide after transition completes
        setTimeout(() => {
            if (!this.isShowing) {
                this.container.style.display = 'none';
            }
        }, 300);
        
        this.isShowing = false;
    }

    /**
     * Toggle the loading animation
     */
    toggle() {
        if (this.isShowing) {
            this.hide();
        } else {
            this.show();
        }
    }

    /**
     * Update the loading text
     * @param {string} text - New loading text
     */
    updateText(text) {
        const textElement = this.container?.querySelector('.loading-text');
        if (textElement) {
            textElement.textContent = text;
        }
    }

    /**
     * Update configuration and refresh the animation
     * @param {Object} newConfig - New configuration options
     */
    updateConfig(newConfig) {
        this.config = { ...this.config, ...newConfig };
        if (this.container) {
            this.createAnimationContent();
        }
    }

    /**
     * Destroy the loading animation and clean up
     */
    destroy() {
        if (this.container && this.container.parentNode) {
            this.container.parentNode.removeChild(this.container);
        }
        this.container = null;
        this.targetElement = null;
        this.isShowing = false;
    }

    /**
     * Static factory method for quick creation
     * @param {HTMLElement|string} target - Target element or selector
     * @param {Object} options - Configuration options
     * @returns {LoadingAnimation} - LoadingAnimation instance
     */
    static create(target, options = {}) {
        const loader = new LoadingAnimation(options);
        loader.init(target);
        return loader;
    }

    /**
     * Static method to create and show a loading animation in one call
     * @param {HTMLElement|string} target - Target element or selector
     * @param {Object} options - Configuration options
     * @returns {LoadingAnimation} - LoadingAnimation instance
     */
    static show(target, options = {}) {
        const loader = LoadingAnimation.create(target, options);
        loader.show();
        return loader;
    }
}

// Export for use in other modules
export default LoadingAnimation;

// Also export as named export for flexibility
export { LoadingAnimation };
