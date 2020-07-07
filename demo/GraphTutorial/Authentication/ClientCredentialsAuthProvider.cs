// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

// <AuthProviderSnippet>
using Microsoft.Extensions.Logging;
using Microsoft.Identity.Client;
using System;
using System.Threading.Tasks;

namespace GraphTutorial.Authentication
{
    public class ClientCredentialsAuthProvider
    {
        private IConfidentialClientApplication _msalClient;
        private string[] _scopes;
        private ILogger _logger;

        public ClientCredentialsAuthProvider(string appId, string clientSecret, string tenantId, string[] scopes, ILogger logger)
        {
            _scopes = scopes;
            _logger = logger;

            _msalClient = ConfidentialClientApplicationBuilder
                .Create(appId)
                .WithAuthority(AadAuthorityAudience.AzureAdMyOrg, true)
                .WithTenantId(tenantId)
                .WithClientSecret(clientSecret)
                .Build();
        }

        public async Task<string> GetAccessToken()
        {
            try
            {
                // Invoke client credentials flow
                var result = await _msalClient
                  .AcquireTokenForClient(_scopes)
                  .ExecuteAsync();

                _logger.LogInformation($"App-only access token: {result.AccessToken}");

                return result.AccessToken;
            }
            catch (Exception exception)
            {
                _logger.LogError(exception, "Error getting access token via client credentials flow");
                return null;
            }
        }
    }
}
// </AuthProviderSnippet>
