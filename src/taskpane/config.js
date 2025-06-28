// config.js
// Configuration for both development and production environments
export const CONFIG = {
  // Determine if we're in development or production
  isDevelopment: process.env.NODE_ENV === 'development',
  
  // Base URL for assets and API calls
  get baseUrl() {
    return this.isDevelopment ? 'https://localhost:3002' : '';
  },
  
  // Helper function to get full URL for assets
  getAssetUrl(path) {
    // Remove leading slash if present
    const cleanPath = path.startsWith('/') ? path.slice(1) : path;
    return `${this.baseUrl}/${cleanPath}`;
  },
  
  // Helper function to get prompt URL
  getPromptUrl(filename) {
    return this.isDevelopment 
      ? `https://localhost:3002/prompts/${filename}`
      : `./prompts/${filename}`;
  }
};
//dfd
// Log configuration on load (for debugging)
if (CONFIG.isDevelopment) {
  console.log('Config loaded:', {
    isDevelopment: CONFIG.isDevelopment,
    baseUrl: CONFIG.baseUrl
  });
}

