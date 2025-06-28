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
    let fullUrl;
    
    if (this.isDevelopment) {
      fullUrl = `${this.baseUrl}/${cleanPath}`;
      console.log(`[CONFIG] Asset URL (dev): ${cleanPath} -> ${fullUrl}`);
    } else {
      // In production, try relative path first
      fullUrl = `./${cleanPath}`;
      console.log(`[CONFIG] Asset URL (prod): ${cleanPath} -> ${fullUrl}`);
    }
    
    return fullUrl;
  },

  // Helper function to get multiple possible asset URLs for fallback
  getAssetUrlsWithFallback(path) {
    const cleanPath = path.startsWith('/') ? path.slice(1) : path;
    
    if (this.isDevelopment) {
      return [`${this.baseUrl}/${cleanPath}`];
    } else {
      // In production, try multiple possible paths
      return [
        `./${cleanPath}`,           // Relative to current page
        `/${cleanPath}`,            // Absolute from root
        `_next/static/${cleanPath}`, // Next.js static folder
        `assets/${cleanPath.replace('assets/', '')}`, // Direct assets folder
      ];
    }
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

