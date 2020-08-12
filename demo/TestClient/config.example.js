// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// <msalConfigSnippet>
const msalConfig = {
  auth: {
    clientId: 'YOUR_TEST_APP_APP_ID_HERE',
    authority: 'https://login.microsoftonline.com/YOUR_TENANT_ID_HERE'
  }
};

const msalRequest = {
  // Scope of the Azure Function
  scopes: [
    'YOUR_AZURE_FUNCTION_APP_ID_HERE/.default'
  ]
}
// </msalConfigSnippet>
