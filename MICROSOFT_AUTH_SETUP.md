# Microsoft Authentication Setup Guide

This guide will help you set up Microsoft authentication for your Projectify Office Add-in using Azure Active Directory.

## Prerequisites

- Azure account (free tier is sufficient)
- Access to Azure Active Directory
- Your Office Add-in deployed and accessible via HTTPS

## Step 1: Create Azure App Registration

1. **Go to Azure Portal**
   - Navigate to [Azure Portal](https://portal.azure.com)
   - Sign in with your Azure account

2. **Access Azure Active Directory**
   - Search for "Azure Active Directory" in the search bar
   - Click on "Azure Active Directory"

3. **Create New App Registration**
   - Click on "App registrations" in the left sidebar
   - Click "New registration"

4. **Configure App Registration**
   - **Name**: Enter a name like "Projectify Office Add-in"
   - **Supported account types**: Select "Accounts in any organizational directory and personal Microsoft accounts"
   - **Redirect URI**: 
     - Platform: Select "Single-page application (SPA)"
     - URI: Enter your add-in URL (e.g., `https://yourdomain.com/src/taskpane/taskpane.html`)
   - Click "Register"

## Step 2: Configure Authentication Settings

1. **Copy Application ID**
   - After registration, copy the "Application (client) ID" from the Overview page
   - You'll need this for the configuration file

2. **Configure Authentication**
   - Go to "Authentication" in the left sidebar
   - Under "Implicit grant and hybrid flows":
     - ✅ Check "Access tokens"
     - ✅ Check "ID tokens"
   - Click "Save"

3. **Add Additional Redirect URIs** (if needed)
   - For development: `https://localhost:3002/src/taskpane/taskpane.html`
   - For production: Your actual domain

## Step 3: Set API Permissions

1. **Go to API Permissions**
   - Click "API permissions" in the left sidebar

2. **Verify Default Permissions**
   - You should see "User.Read" under Microsoft Graph
   - This permission allows reading user profile information

3. **Add Additional Permissions** (optional)
   - Click "Add a permission" if you need additional scopes
   - For basic authentication, "User.Read" is sufficient

## Step 4: Update Configuration

1. **Open Configuration File**
   - Navigate to `src/taskpane/config.js`

2. **Update Client ID**
   ```javascript
   msalConfig: {
     auth: {
       clientId: "YOUR_COPIED_CLIENT_ID_HERE", // Replace with your actual client ID
       authority: "https://login.microsoftonline.com/common",
       redirectUri: window.location.origin + "/src/taskpane/taskpane.html"
     }
   }
   ```

3. **Update Redirect URI** (if needed)
   - Make sure the `redirectUri` matches what you configured in Azure
   - For development, you might need: `"https://localhost:3002/src/taskpane/taskpane.html"`

## Step 5: Test Authentication

1. **Deploy Your Add-in**
   - Make sure your add-in is accessible via HTTPS
   - The authentication will not work with HTTP

2. **Test Sign-In**
   - Open your Office Add-in
   - Navigate to Client Mode
   - Click "Sign in with Microsoft"
   - You should be redirected to Microsoft's login page

3. **Verify User Info**
   - After successful sign-in, you should see:
     - User's name and email
     - User's profile picture (if available)
     - Sign-out button

## Troubleshooting

### Common Issues

1. **"AADSTS50011: The reply URL specified in the request does not match"**
   - Ensure the redirect URI in Azure matches exactly what's in your configuration
   - Include the full path, not just the domain

2. **"The provided request must include a 'scope' input parameter"**
   - This is usually handled automatically by the configuration
   - Ensure `loginRequest.scopes` includes required permissions

3. **"User cancelled the authentication"**
   - This happens when users close the popup or click cancel
   - Normal behavior, no action needed

4. **"Popup window error"**
   - Ensure popups are allowed in the browser
   - Some corporate networks block authentication popups

### Development vs Production

- **Development**: Use `https://localhost:3002` URLs
- **Production**: Use your actual domain
- Both environments need to be registered as redirect URIs in Azure

### Office Add-in Specific Notes

- Office Add-ins run in a sandboxed environment
- Authentication popups work differently than regular web apps
- Session storage is preferred over local storage for security
- The add-in must be served over HTTPS for authentication to work

## Security Best Practices

1. **Never expose client secrets** - Single-page applications don't use client secrets
2. **Use HTTPS everywhere** - Authentication requires secure connections
3. **Validate tokens** - Always verify tokens on your backend if you have one
4. **Implement proper error handling** - Handle authentication failures gracefully
5. **Regular token refresh** - MSAL handles this automatically

## Support

If you encounter issues:
1. Check the browser's developer console for error messages
2. Verify your Azure app registration settings
3. Ensure redirect URIs match exactly
4. Test with different browsers to rule out browser-specific issues

For more information, refer to:
- [Microsoft Authentication Library (MSAL) documentation](https://docs.microsoft.com/en-us/azure/active-directory/develop/msal-overview)
- [Office Add-ins authentication guide](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/auth-with-office-dialog-api) 