# LoadingAnimation Module

A reusable, customizable loading animation module that can be easily integrated into any web application. This module provides a flexible loading spinner with optional text, various styles, and multiple animation types.

## Features

- 🎨 Fully customizable appearance (size, colors, animations)
- 📝 Optional loading text with dynamic updates
- 🎯 Multiple positioning options (relative, absolute, fixed)
- 🎭 Different animation types (spinner, dots, progress bar, pulse)
- 🎨 Theme support (light, dark, success, error states)
- 📱 Responsive and accessible
- 🔧 Easy integration with existing code
- 🚀 Lightweight with no dependencies

## Installation

Simply import the module into your project:

```javascript
import LoadingAnimation from './LoadingAnimation.js';
```

Don't forget to include the CSS file:

```html
<link rel="stylesheet" href="./styles/components/loading-animation.css">
```

## Quick Start

### Basic Usage

```javascript
// Create a loader instance
const loader = new LoadingAnimation();
loader.init('#my-container');

// Show loading
loader.show();

// Hide loading
loader.hide();
```

### One-liner with Static Method

```javascript
// Show loading immediately
const loader = LoadingAnimation.show('#my-container', {
    loadingText: 'Please wait...'
});

// Hide when done
setTimeout(() => loader.hide(), 3000);
```

## Configuration Options

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `containerId` | string | null | ID of existing container element |
| `spinnerSize` | string | '32px' | Size of the spinner |
| `spinnerBorderWidth` | string | '3px' | Border width of the spinner |
| `spinnerColor` | string | '#000000' | Color of the spinning part |
| `spinnerTrackColor` | string | '#f0f0f0' | Color of the spinner track |
| `animationDuration` | string | '0.8s' | Duration of one rotation |
| `showText` | boolean | true | Whether to show loading text |
| `loadingText` | string | 'Loading...' | Text to display |
| `textColor` | string | '#6b7280' | Color of the loading text |
| `textSize` | string | '14px' | Font size of the loading text |
| `className` | string | 'loading-animation-container' | CSS class name for the container |
| `position` | string | 'relative' | CSS position (relative, absolute, fixed) |
| `zIndex` | number | 1000 | z-index value |

## API Methods

### Constructor
```javascript
new LoadingAnimation(options)
```
Creates a new LoadingAnimation instance with the specified options.

### init(target)
```javascript
loader.init('#container-id')
// or
loader.init(document.getElementById('container'))
```
Initializes the loader in the specified target element.

### show()
```javascript
loader.show()
```
Shows the loading animation.

### hide()
```javascript
loader.hide()
```
Hides the loading animation.

### toggle()
```javascript
loader.toggle()
```
Toggles the visibility of the loading animation.

### updateText(text)
```javascript
loader.updateText('Almost done...')
```
Updates the loading text dynamically.

### updateConfig(newConfig)
```javascript
loader.updateConfig({
    spinnerColor: '#16a34a',
    loadingText: 'Success!'
})
```
Updates the configuration and refreshes the animation.

### destroy()
```javascript
loader.destroy()
```
Removes the loader from the DOM and cleans up.

## Static Methods

### LoadingAnimation.create(target, options)
```javascript
const loader = LoadingAnimation.create('#container', {
    loadingText: 'Processing...'
})
```
Factory method to create and initialize a loader.

### LoadingAnimation.show(target, options)
```javascript
const loader = LoadingAnimation.show('#container', {
    position: 'fixed'
})
```
Creates, initializes, and shows a loader in one call.

## Examples

### Full Page Overlay
```javascript
// Create overlay
const overlay = document.createElement('div');
overlay.className = 'loading-animation-overlay';
document.body.appendChild(overlay);

// Add loader
const loader = LoadingAnimation.show(overlay, {
    loadingText: 'Loading application...',
    spinnerSize: '64px'
});

// Remove after loading
setTimeout(() => {
    loader.destroy();
    overlay.remove();
}, 3000);
```

### Inline Button Loading
```javascript
const button = document.querySelector('#submit-btn');
const loader = new LoadingAnimation({
    className: 'loading-animation inline',
    spinnerSize: '16px',
    showText: false
});

// Add loader to button
const span = document.createElement('span');
button.appendChild(span);
loader.init(span);

button.addEventListener('click', async () => {
    button.disabled = true;
    loader.show();
    
    await submitForm();
    
    loader.hide();
    button.disabled = false;
});
```

### Dynamic Progress Updates
```javascript
const loader = LoadingAnimation.create('#progress-area', {
    loadingText: 'Starting...'
});

loader.show();

const steps = [
    'Connecting to server...',
    'Authenticating...',
    'Fetching data...',
    'Processing results...',
    'Complete!'
];

steps.forEach((step, index) => {
    setTimeout(() => {
        loader.updateText(step);
        if (index === steps.length - 1) {
            setTimeout(() => loader.hide(), 500);
        }
    }, index * 1000);
});
```

### Error Handling
```javascript
async function saveData() {
    const loader = LoadingAnimation.show('#form', {
        loadingText: 'Saving...'
    });
    
    try {
        await api.save(data);
        loader.updateConfig({
            spinnerColor: '#16a34a',
            loadingText: 'Saved!',
            className: 'loading-animation success'
        });
        setTimeout(() => loader.hide(), 1500);
    } catch (error) {
        loader.updateConfig({
            spinnerColor: '#dc2626',
            loadingText: 'Error: ' + error.message,
            className: 'loading-animation error'
        });
        setTimeout(() => loader.hide(), 3000);
    }
}
```

## CSS Classes

The module uses these CSS classes that you can customize:

- `.loading-animation` - Main container
- `.loading-spinner` - The spinner element
- `.loading-text` - The text element
- `.loading-hidden` - Hidden state
- `.loading-visible` - Visible state
- `.loading-animation.inline` - Inline variant
- `.loading-animation.success` - Success state
- `.loading-animation.error` - Error state
- `.loading-animation.theme-dark` - Dark theme

## Integration with Existing Code

To replace existing loading animations:

```javascript
// Import the integration helper
import { initializeLoadingAnimations, setButtonLoading } from './LoadingAnimationIntegration.js';

// Initialize on page load
document.addEventListener('DOMContentLoaded', () => {
    initializeLoadingAnimations();
});

// Use the same function names as before
setButtonLoading(true);  // Show loading
setButtonLoading(false); // Hide loading
```

## Browser Support

- Chrome/Edge: Full support
- Firefox: Full support
- Safari: Full support
- IE11: Basic support (with polyfills for ES6 features)

## Performance Tips

1. **Reuse instances**: Create loader instances once and reuse them
2. **Use containers**: Specify `containerId` to avoid creating new DOM elements
3. **Clean up**: Call `destroy()` when loaders are no longer needed
4. **Batch updates**: Use `updateConfig()` for multiple changes at once

## Troubleshooting

### Loader not showing
- Ensure the target element exists in the DOM
- Check if CSS file is loaded
- Verify no CSS conflicts with `display: none`

### Animation not smooth
- Check if other JavaScript is blocking the main thread
- Ensure no CSS transitions conflict with the animation
- Try adjusting `animationDuration`

### Text not updating
- Make sure `showText` is set to `true`
- Check if the text element exists in the DOM
- Use `updateText()` method, not direct DOM manipulation

## License

This module is part of the Projectify application and follows the project's licensing terms.


