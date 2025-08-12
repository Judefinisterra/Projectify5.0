/**
 * LoadingAnimationExamples.js
 * 
 * This file demonstrates various ways to use the LoadingAnimation module
 * in your application. Each example shows a different use case and configuration.
 */

import LoadingAnimation from './LoadingAnimation.js';

// ========================================
// EXAMPLE 1: Basic Usage
// ========================================
function basicExample() {
    // Create a basic loading animation
    const loader = new LoadingAnimation();
    loader.init('#chat-log-container');
    
    // Show loading when making an API call
    loader.show();
    
    // Simulate API call
    setTimeout(() => {
        loader.hide();
    }, 3000);
}

// ========================================
// EXAMPLE 2: Custom Configuration
// ========================================
function customConfigExample() {
    const loader = new LoadingAnimation({
        spinnerSize: '48px',
        spinnerBorderWidth: '4px',
        spinnerColor: '#3b82f6',
        spinnerTrackColor: '#e5e7eb',
        animationDuration: '1s',
        loadingText: 'Processing your request...',
        textColor: '#374151',
        textSize: '16px'
    });
    
    loader.init('#my-container');
    loader.show();
}

// ========================================
// EXAMPLE 3: Using Static Methods
// ========================================
function staticMethodExample() {
    // Quick one-liner to show loading
    const loader = LoadingAnimation.show('#content-area', {
        loadingText: 'Loading data...',
        position: 'absolute'
    });
    
    // Hide after operation completes
    fetchData().then(() => {
        loader.hide();
    });
}

// ========================================
// EXAMPLE 4: Inline Loading Animation
// ========================================
function inlineLoadingExample() {
    const button = document.querySelector('#submit-button');
    
    // Create inline loader next to button text
    const loader = new LoadingAnimation({
        className: 'loading-animation inline',
        spinnerSize: '16px',
        showText: false,
        position: 'relative'
    });
    
    // Create a span to hold the loader
    const loaderSpan = document.createElement('span');
    button.appendChild(loaderSpan);
    loader.init(loaderSpan);
    
    button.addEventListener('click', async () => {
        button.disabled = true;
        loader.show();
        
        await performAction();
        
        loader.hide();
        button.disabled = false;
    });
}

// ========================================
// EXAMPLE 5: Full Page Overlay
// ========================================
function fullPageOverlayExample() {
    // Create overlay container
    const overlay = document.createElement('div');
    overlay.className = 'loading-animation-overlay';
    document.body.appendChild(overlay);
    
    // Create loader in overlay
    const loader = new LoadingAnimation({
        loadingText: 'Please wait...',
        spinnerSize: '64px',
        position: 'relative'
    });
    
    loader.init(overlay);
    loader.show();
    
    // Remove overlay after loading
    setTimeout(() => {
        loader.destroy();
        overlay.remove();
    }, 3000);
}

// ========================================
// EXAMPLE 6: Dynamic Text Updates
// ========================================
function dynamicTextExample() {
    const loader = LoadingAnimation.create('#progress-container', {
        loadingText: 'Initializing...'
    });
    
    loader.show();
    
    // Update text as operation progresses
    setTimeout(() => loader.updateText('Connecting to server...'), 1000);
    setTimeout(() => loader.updateText('Fetching data...'), 2000);
    setTimeout(() => loader.updateText('Processing...'), 3000);
    setTimeout(() => loader.updateText('Almost done...'), 4000);
    setTimeout(() => {
        loader.updateText('Complete!');
        setTimeout(() => loader.hide(), 500);
    }, 5000);
}

// ========================================
// EXAMPLE 7: Multiple Loaders
// ========================================
function multipleLoadersExample() {
    // Create different loaders for different sections
    const chatLoader = new LoadingAnimation({
        containerId: 'chat-loading',
        loadingText: 'Loading messages...',
        spinnerColor: '#10b981'
    });
    
    const dataLoader = new LoadingAnimation({
        containerId: 'data-loading',
        loadingText: 'Fetching data...',
        spinnerColor: '#3b82f6'
    });
    
    chatLoader.init('#chat-section');
    dataLoader.init('#data-section');
    
    // Show/hide independently
    chatLoader.show();
    dataLoader.show();
    
    setTimeout(() => chatLoader.hide(), 2000);
    setTimeout(() => dataLoader.hide(), 3500);
}

// ========================================
// EXAMPLE 8: Error and Success States
// ========================================
function stateChangeExample() {
    const loader = new LoadingAnimation({
        loadingText: 'Saving...'
    });
    
    loader.init('#form-container');
    loader.show();
    
    // Simulate save operation
    saveData()
        .then(() => {
            // Change to success state
            loader.updateConfig({
                spinnerColor: '#16a34a',
                loadingText: 'Saved successfully!',
                className: 'loading-animation success'
            });
            
            setTimeout(() => loader.hide(), 1500);
        })
        .catch(() => {
            // Change to error state
            loader.updateConfig({
                spinnerColor: '#dc2626',
                loadingText: 'Error saving data',
                className: 'loading-animation error'
            });
            
            setTimeout(() => loader.hide(), 2000);
        });
}

// ========================================
// EXAMPLE 9: Integration with Existing Code
// ========================================
function integrationExample() {
    // Replace existing loading animation logic
    const loader = new LoadingAnimation({
        containerId: 'loading-animation', // Use existing container ID
        showText: false // Match existing behavior
    });
    
    // Initialize once
    loader.init('#chat-log-container');
    
    // Replace setButtonLoading function
    window.setButtonLoading = function(isLoading) {
        const sendButton = document.getElementById('send');
        if (sendButton) {
            sendButton.disabled = isLoading;
        }
        
        if (isLoading) {
            loader.show();
        } else {
            loader.hide();
        }
    };
}

// ========================================
// EXAMPLE 10: Promise-based Usage
// ========================================
function promiseExample() {
    const loader = new LoadingAnimation({
        loadingText: 'Processing...'
    });
    
    loader.init('#async-container');
    
    // Helper function to wrap async operations with loading
    async function withLoading(asyncFn) {
        loader.show();
        try {
            const result = await asyncFn();
            return result;
        } finally {
            loader.hide();
        }
    }
    
    // Usage
    withLoading(async () => {
        const response = await fetch('/api/data');
        return response.json();
    }).then(data => {
        console.log('Data loaded:', data);
    });
}

// ========================================
// EXAMPLE 11: Custom Animation Types
// ========================================
function customAnimationTypes() {
    // Dots animation
    const dotsLoader = {
        init(container) {
            const dotsHtml = `
                <div class="loading-dots">
                    <span></span>
                    <span></span>
                    <span></span>
                </div>
            `;
            container.innerHTML = dotsHtml;
        }
    };
    
    // Progress bar animation
    const progressLoader = {
        init(container) {
            const progressHtml = `
                <div class="loading-progress">
                    <div class="loading-progress-bar"></div>
                </div>
            `;
            container.innerHTML = progressHtml;
        }
    };
    
    // Pulse animation
    const pulseLoader = {
        init(container) {
            const pulseHtml = `
                <div class="loading-pulse"></div>
            `;
            container.innerHTML = pulseHtml;
        }
    };
}

// ========================================
// EXAMPLE 12: React/Vue Integration
// ========================================

// React Component Example
/*
import React, { useEffect, useRef } from 'react';
import LoadingAnimation from './LoadingAnimation';

function LoadingSpinner({ isLoading, text = 'Loading...' }) {
    const containerRef = useRef(null);
    const loaderRef = useRef(null);
    
    useEffect(() => {
        if (containerRef.current) {
            loaderRef.current = new LoadingAnimation({
                loadingText: text
            });
            loaderRef.current.init(containerRef.current);
        }
        
        return () => {
            if (loaderRef.current) {
                loaderRef.current.destroy();
            }
        };
    }, []);
    
    useEffect(() => {
        if (loaderRef.current) {
            if (isLoading) {
                loaderRef.current.show();
            } else {
                loaderRef.current.hide();
            }
        }
    }, [isLoading]);
    
    useEffect(() => {
        if (loaderRef.current) {
            loaderRef.current.updateText(text);
        }
    }, [text]);
    
    return <div ref={containerRef}></div>;
}
*/

// Vue Component Example
/*
<template>
  <div ref="loaderContainer"></div>
</template>

<script>
import LoadingAnimation from './LoadingAnimation';

export default {
  props: {
    isLoading: Boolean,
    text: {
      type: String,
      default: 'Loading...'
    }
  },
  mounted() {
    this.loader = new LoadingAnimation({
      loadingText: this.text
    });
    this.loader.init(this.$refs.loaderContainer);
  },
  watch: {
    isLoading(newVal) {
      if (newVal) {
        this.loader.show();
      } else {
        this.loader.hide();
      }
    },
    text(newVal) {
      this.loader.updateText(newVal);
    }
  },
  beforeDestroy() {
    if (this.loader) {
      this.loader.destroy();
    }
  }
};
</script>
*/

// ========================================
// Utility Functions
// ========================================

// Mock functions for examples
async function fetchData() {
    return new Promise(resolve => setTimeout(resolve, 2000));
}

async function performAction() {
    return new Promise(resolve => setTimeout(resolve, 1500));
}

async function saveData() {
    return new Promise((resolve, reject) => {
        setTimeout(() => {
            if (Math.random() > 0.5) {
                resolve();
            } else {
                reject(new Error('Save failed'));
            }
        }, 2000);
    });
}

// Export examples for testing
export {
    basicExample,
    customConfigExample,
    staticMethodExample,
    inlineLoadingExample,
    fullPageOverlayExample,
    dynamicTextExample,
    multipleLoadersExample,
    stateChangeExample,
    integrationExample,
    promiseExample,
    customAnimationTypes
};
