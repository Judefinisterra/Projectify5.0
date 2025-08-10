/**
 * Chat Layout Debugger
 * Diagnostic utility to understand and fix the chat positioning issue
 */

class ChatLayoutDebugger {
    constructor() {
        this.isDebugging = true;
        this.headerHeight = '0';
    }

    /**
     * Debug and fix chat layout issues
     */
    debugAndFix() {
        if (!this.isDebugging) return;

        console.log('🔍 CHAT LAYOUT DEBUGGER STARTED');
        
        // Find all relevant elements
        const elements = this.findElements();
        this.logElementInfo(elements);
        
        // Apply nuclear fixes
        this.applyNuclearFixes(elements);
        
        // Monitor for changes
        this.monitorChanges(elements);
    }

    /**
     * Find all relevant layout elements
     */
    findElements() {
        return {
            clientModeView: document.getElementById('client-mode-view'),
            chatContainer: document.getElementById('client-chat-container'),
            chatLog: document.getElementById('chat-log-client'),
            header: document.querySelector('#client-mode-view .simple-header'),
            mainContent: document.querySelector('#client-mode-view .main-content')
        };
    }

    /**
     * Log detailed information about elements
     */
    logElementInfo(elements) {
        console.log('📊 ELEMENT ANALYSIS:');
        
        Object.entries(elements).forEach(([name, element]) => {
            if (element) {
                const computed = window.getComputedStyle(element);
                console.log(`\n${name.toUpperCase()}:`, {
                    element,
                    display: computed.display,
                    position: computed.position,
                    top: computed.top,
                    marginTop: computed.marginTop,
                    paddingTop: computed.paddingTop,
                    zIndex: computed.zIndex,
                    height: computed.height,
                    transform: computed.transform,
                    classes: element.className,
                    inlineStyles: element.style.cssText
                });
            } else {
                console.log(`\n${name.toUpperCase()}: NOT FOUND`);
            }
        });
    }

    /**
     * Apply nuclear layout fixes that override everything
     */
    applyNuclearFixes(elements) {
        console.log('💥 APPLYING NUCLEAR FIXES...');

        // Fix 0: Ensure client mode view is visible
        if (elements.clientModeView) {
            this.setStyles(elements.clientModeView, {
                display: 'flex',
                visibility: 'visible',
                opacity: '1'
            });
            console.log('✅ Client mode view made visible');
        }

        // Fix 1: Ensure header is properly positioned (STICKY within container)
        if (elements.header) {
            this.setStyles(elements.header, {
                position: 'sticky',
                top: '0',
                left: '0',
                right: '0',
                zIndex: '10',
                backgroundColor: 'white',
                borderBottom: '1px solid #eee'
            });
            console.log('✅ Header set to sticky');
        }

        // Fix 2: Force chat container layout and conversation state
        if (elements.chatContainer) {
            // Force conversation-active class
            if (!elements.chatContainer.classList.contains('conversation-active')) {
                elements.chatContainer.classList.add('conversation-active');
                console.log('🔄 Added conversation-active class');
            }
            
            this.setStyles(elements.chatContainer, {
                paddingTop: this.headerHeight,
                boxSizing: 'border-box',
                minHeight: '100vh',
                height: '100%',
                overflowY: 'auto',
                overflowX: 'hidden',
                scrollPaddingTop: this.headerHeight
            });
            console.log('✅ Chat container fixed');
        }

        // Fix 3: Force chat log positioning AND visibility (CONSERVATIVE)
        if (elements.chatLog) {
            this.setStyles(elements.chatLog, {
                display: 'block',
                visibility: 'visible',
                opacity: '1',
                marginTop: '0',
                paddingTop: '1rem',
                minHeight: '200px',
                maxHeight: 'none',
                overflow: 'visible',
                position: 'relative',
                zIndex: '1'
            });
            console.log('✅ Chat log fixed and made visible (conservative)');
        }

        // Fix 4: Adjust main content (MINIMAL)
        if (elements.mainContent) {
            this.setStyles(elements.mainContent, {
                minHeight: '100vh',
                overflow: 'visible'
            });
            console.log('✅ Main content fixed (minimal)');
        }
    }

    /**
     * Set styles with highest priority
     */
    setStyles(element, styles) {
        const toKebab = (prop) => prop.replace(/[A-Z]/g, (m) => '-' + m.toLowerCase());
        Object.entries(styles).forEach(([property, value]) => {
            const cssProp = toKebab(property);
            element.style.setProperty(cssProp, value, 'important');
        });
    }

    /**
     * Monitor for DOM changes that might break the layout
     */
    monitorChanges(elements) {
        console.log('👀 MONITORING FOR LAYOUT CHANGES...');

        // Watch for style changes on chat log
        if (elements.chatLog) {
            const observer = new MutationObserver((mutations) => {
                mutations.forEach((mutation) => {
                    if (mutation.type === 'attributes' && mutation.attributeName === 'style') {
                        console.log('⚠️ Chat log styles changed, reapplying fixes...');
                        this.applyNuclearFixes(elements);
                    }
                });
            });

            observer.observe(elements.chatLog, {
                attributes: true,
                attributeFilter: ['style', 'class']
            });
        }

        // Watch for new chat log creation
        const containerObserver = new MutationObserver((mutations) => {
            mutations.forEach((mutation) => {
                mutation.addedNodes.forEach((node) => {
                    if (node.nodeType === 1 && node.id === 'chat-log-client') {
                        console.log('🆕 New chat log detected, applying fixes...');
                        setTimeout(() => {
                            const newElements = this.findElements();
                            this.applyNuclearFixes(newElements);
                        }, 100);
                    }
                });
            });
        });

        if (elements.chatContainer) {
            containerObserver.observe(elements.chatContainer, {
                childList: true,
                subtree: true
            });
        }
    }

    /**
     * Force fix on demand
     */
    forceFix() {
        console.log('🚀 FORCE FIX TRIGGERED');
        const elements = this.findElements();
        this.applyNuclearFixes(elements);
    }
}

// Create global instance
window.chatDebugger = new ChatLayoutDebugger();

// Disabled auto-debugging - CSS flexbox layout should handle everything
// if (document.readyState === 'loading') {
//     document.addEventListener('DOMContentLoaded', () => {
//         setTimeout(() => window.chatDebugger.debugAndFix(), 1000);
//     });
// } else {
//     setTimeout(() => window.chatDebugger.debugAndFix(), 1000);
// }

// Export for use in other modules
export default ChatLayoutDebugger; 