// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

const msalConfig = {
  auth: {
    clientId: 'YOUR_TEST_APP_CLIENT_ID_HERE',
    authority: 'https://login.microsoftonline.com/YOUR_TENANT_ID_HERE'
  }
};

const msalRequest = {
  // Scope of the Azure Function
  scopes: [
    'YOUR_AZURE_FUNCTION_CLIENT_ID_HERE/.default'
  ]
}
