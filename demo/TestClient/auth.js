// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// <signInSignOutSnippet>
// Create the main MSAL instance
// configuration parameters are located in config.js
const msalClient = new msal.PublicClientApplication(msalConfig);

async function signIn() {
  // Login
  try {
    // Use MSAL to login
    const authResult = await msalClient.loginPopup(msalRequest);
    // Save the account username, needed for token acquisition
    sessionStorage.setItem('msal-userName', authResult.account.username);
    // Refresh home page
    updatePage(Views.home);
  } catch (error) {
    console.log(error);
    updatePage(Views.error, {
      message: 'Error logging in',
      debug: error
    });
  }
}

function signOut() {
  account = null;
  sessionStorage.removeItem('msal-userName');
  msalClient.logout();
}
// </signInSignOutSnippet>

// <getTokenSnippet>
async function getToken() {
  let account = sessionStorage.getItem('msal-userName');
  if (!account){
    throw new Error(
      'User account missing from session. Please sign out and sign in again.');
  }

  try {
    // First, attempt to get the token silently
    const silentRequest = {
      scopes: msalRequest.scopes,
      account: msalClient.getAccountByUsername(account)
    };

    const silentResult = await msalClient.acquireTokenSilent(silentRequest);
    return silentResult.accessToken;
  } catch (silentError) {
    // If silent requests fails with InteractionRequiredAuthError,
    // attempt to get the token interactively
    if (silentError instanceof msal.InteractionRequiredAuthError) {
      const interactiveResult = await msalClient.acquireTokenPopup(msalRequest);
      return interactiveResult.accessToken;
    } else {
      throw silentError;
    }
  }
}
// </getTokenSnippet>
