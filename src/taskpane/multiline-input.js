/**
 * Dynamic Multiline Input Handler
 * Toggles .multiline class on .chatgpt-input-bar when textarea content wraps
 */

class MultilineInputHandler {
    constructor() {
        this.textarea = null;
        this.inputBar = null;
        this.debounceTimer = null;
        this.isInitialized = false;
    }

    /**
     * Initialize the multiline handler
     */
    init() {
        // Find the textarea and input bar
        this.textarea = document.getElementById('user-input-client');
        this.inputBar = document.querySelector('.chatgpt-input-bar');
        
        if (!this.textarea || !this.inputBar) {
            console.warn('[MultilineInput] Textarea or input bar not found');
            return false;
        }

        // Set up event listeners
        this.setupEventListeners();
        
        // Initial check
        this.checkMultilineState();
        
        this.isInitialized = true;
        console.log('[MultilineInput] Initialized successfully');
        return true;
    }

    /**
     * Set up event listeners with debouncing
     */
    setupEventListeners() {
        // Input event (typing, pasting, etc.)
        this.textarea.addEventListener('input', () => {
            this.debouncedCheck();
        });

        // Focus event (when user clicks in textarea)
        this.textarea.addEventListener('focus', () => {
            this.checkMultilineState();
        });

        // Window resize event
        window.addEventListener('resize', () => {
            this.debouncedCheck();
        });

        // Paste event (for large text)
        this.textarea.addEventListener('paste', () => {
            // Small delay to allow paste content to be processed
            setTimeout(() => {
                this.checkMultilineState();
            }, 10);
        });
    }

    /**
     * Debounced check to prevent excessive calls during rapid typing
     */
    debouncedCheck() {
        clearTimeout(this.debounceTimer);
        this.debounceTimer = setTimeout(() => {
            this.checkMultilineState();
        }, 50); // 50ms delay
    }

    /**
     * Check if textarea content requires multiline layout
     */
    checkMultilineState() {
        if (!this.textarea || !this.inputBar) return;

        // Use the actual single line height from CSS
        const singleLineHeight = 44; // From CSS: min-height: 44px for single line
        const maxHeight = 116; // From CSS: max-height: 116px
        
        // Temporarily reset to get accurate scrollHeight
        this.textarea.style.height = 'auto';
        const scrollHeight = this.textarea.scrollHeight;
        
        // Calculate new height (same logic as original auto-resize)
        const newHeight = Math.min(scrollHeight, maxHeight);
        this.textarea.style.height = newHeight + 'px';
        
        // Determine if we need multiline layout
        const hasWrappedText = scrollHeight > singleLineHeight;
        const hasNewlines = this.textarea.value.includes('\n');
        const needsMultiline = hasWrappedText || hasNewlines;
        
        // Only toggle if state actually changes (prevent flicker)
        const currentlyMultiline = this.inputBar.classList.contains('multiline');
        
        if (needsMultiline && !currentlyMultiline) {
            this.inputBar.classList.add('multiline');
        } else if (!needsMultiline && currentlyMultiline) {
            this.inputBar.classList.remove('multiline');
        }
        
        // Also handle the "expanded" class for compatibility with existing CSS
        if (scrollHeight > singleLineHeight) {
            this.inputBar.classList.add('expanded');
        } else {
            this.inputBar.classList.remove('expanded');
        }
        
        // Handle scrollable state
        if (scrollHeight > maxHeight) {
            this.textarea.classList.add('scrollable');
            this.textarea.style.overflowY = 'auto';
        } else {
            this.textarea.classList.remove('scrollable');
            this.textarea.style.overflowY = 'hidden';
        }

        // Debug logging (reduced to prevent spam)
        if (needsMultiline !== currentlyMultiline) {
            console.log('[MultilineInput] State changed:', {
                scrollHeight,
                newHeight,
                singleLineHeight,
                needsMultiline,
                hasNewlines
            });
        }
    }

    /**
     * Calculate single line height including padding
     */
    getSingleLineHeight() {
        // Create a temporary element to measure single line height
        const temp = document.createElement('textarea');
        temp.style.cssText = getComputedStyle(this.textarea).cssText;
        temp.style.position = 'absolute';
        temp.style.left = '-9999px';
        temp.style.height = 'auto';
        temp.rows = 1;
        temp.value = 'X'; // Single character
        
        document.body.appendChild(temp);
        const height = temp.scrollHeight;
        document.body.removeChild(temp);
        
        return height;
    }

    /**
     * Check for long unbroken words that might force wrapping
     */
    hasLongUnbrokenText() {
        const text = this.textarea.value;
        const containerWidth = this.textarea.clientWidth;
        
        // Rough estimation: if any word is longer than ~80% of container width in characters
        // This is approximate based on average character width
        const avgCharWidth = 8; // pixels, approximate
        const maxCharsPerLine = Math.floor(containerWidth * 0.8 / avgCharWidth);
        
        const words = text.split(/\s+/);
        return words.some(word => word.length > maxCharsPerLine);
    }

    /**
     * Force a state check (useful for external calls)
     */
    forceCheck() {
        if (this.isInitialized) {
            this.checkMultilineState();
        }
    }

    /**
     * Reset to single-line state
     */
    reset() {
        if (this.inputBar) {
            this.inputBar.classList.remove('multiline');
        }
    }

    /**
     * Cleanup event listeners
     */
    destroy() {
        if (this.textarea) {
            this.textarea.removeEventListener('input', this.debouncedCheck);
            this.textarea.removeEventListener('focus', this.checkMultilineState);
            this.textarea.removeEventListener('paste', this.checkMultilineState);
        }
        
        window.removeEventListener('resize', this.debouncedCheck);
        clearTimeout(this.debounceTimer);
        
        this.isInitialized = false;
        console.log('[MultilineInput] Destroyed');
    }
}

// Create global instance (will be initialized by taskpane.js)
const multilineInputHandler = new MultilineInputHandler();

// ES6 export
export default MultilineInputHandler;
export { multilineInputHandler };