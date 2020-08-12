// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

// <AuthProviderSnippet>
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace GraphTutorial.Authentication
{
    public class ClientCredentialsAuthProvider : IAuthenticationProvider
    {
        private IConfidentialClientApplication _msalClient;
        private string[] _scopes;
        private ILogger _logger;

        public ClientCredentialsAuthProvider(
            string appId,
            string clientSecret,
            string tenantId,
            string[] scopes,
            ILogger logger)
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
                // NOTE: This will return a cached token if a valid one
                // exists
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

        // This is the delegate called by the GraphServiceClient on each
        // request.
        public async Task AuthenticateRequestAsync(HttpRequestMessage requestMessage)
        {
            // Get the current access token
            var token = await GetAccessToken();

            // Add the token in the Authorization header
            requestMessage.Headers.Authorization =
                new AuthenticationHeaderValue("Bearer", token);
        }
    }
}
// </AuthProviderSnippet>
