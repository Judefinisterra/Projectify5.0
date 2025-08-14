# Authentication Troubleshooting Guide

## Recent Changes Made

I've implemented several improvements to handle authentication issues in your Excel add-in:

### 1. Enhanced Token Refresh Logic
- Added better error handling when tokens expire
- Improved logging to debug token storage issues
- Added fallback checks for different token storage keys

### 2. Automatic Re-authentication Flow
- When authentication fails and no refresh token is available, the app now:
  - Clears all stored tokens
  - Emits an 'auth-expired' event
  - Automatically shows the authentication view
  - Displays an error message to the user

### 3. Debugging Tools
- Added `debugTokens()` function accessible from browser console
- Enhanced logging throughout the authentication process
- Better visibility into what tokens are stored and where

## How to Debug Authentication Issues

### 1. Check Token Status
Open the browser console and run:
```javascript
debugTokens()
```

This will show you:
- Which tokens are stored in sessionStorage
- Which tokens are stored in localStorage
- Whether the app considers the user authenticated

### 2. Check Backend Response
When signing in, look for this log message in the console:
```
📦 Backend auth response received: {...}
```

This shows whether the backend is sending both access and refresh tokens.

### 3. Common Issues and Solutions

#### Issue: "No refresh token available"
**Cause**: The backend didn't send a refresh token during initial authentication
**Solution**: 
- Check the backend API to ensure it's sending refresh_token in the auth response
- The backend should return both `access_token` and `refresh_token` (or `session.refresh_token`)

#### Issue: Authentication works but expires quickly
**Cause**: Access token has short expiration and refresh fails
**Solution**: 
- Ensure refresh token is being properly stored
- Check that the `/auth/refresh` endpoint is working correctly
- Verify refresh token hasn't expired on the backend

#### Issue: User sees authentication screen repeatedly
**Cause**: Tokens are not being persisted properly
**Solution**: 
- Check browser storage settings (cookies/localStorage not blocked)
- Ensure the app domain is trusted in Excel
- Try clearing all storage and re-authenticating

## Testing the Fix

1. Clear all browser storage for your app
2. Sign in with Google
3. Check the console for the backend auth response
4. Run `debugTokens()` to verify tokens are stored
5. Wait for token expiration or manually delete the access token
6. Try to use a feature that requires authentication
7. The app should either refresh the token automatically or show the sign-in screen

## Next Steps

If issues persist:
1. Check the backend API logs to see what's happening during authentication
2. Verify the backend is correctly validating Google ID tokens
3. Ensure the backend is generating and storing refresh tokens properly
4. Check if there are any CORS issues preventing proper API communication

