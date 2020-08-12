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
    public class OnBehalfOfAuthProvider : IAuthenticationProvider
    {
        private IConfidentialClientApplication _msalClient;

        private TokenValidationResult _tokenResult;
        private string[] _scopes;
        private ILogger _logger;

        public OnBehalfOfAuthProvider(
            IConfidentialClientApplication msalClient,
            TokenValidationResult tokenResult,
            string[] scopes,
            ILogger logger)
        {
            _scopes = scopes;
            _logger = logger;

            _tokenResult = tokenResult;
            _msalClient = msalClient;
        }

        public async Task<string> GetAccessToken()
        {
            try
            {
                // First attempt to get token from the cache for this user
                // Check for a matching account in the cache
                var account = await _msalClient.GetAccountAsync(_tokenResult.MsalAccountId);
                if (account != null)
                {
                    // Make a "silent" request for a token. This will
                    // return the cached token if still valid, and will handle
                    // refreshing the token if needed
                    var cacheResult = await _msalClient
                        .AcquireTokenSilent(_scopes, account)
                        .ExecuteAsync();

                    _logger.LogInformation($"User access token: {cacheResult.AccessToken}");
                    return cacheResult.AccessToken;
                }
            }
            catch (MsalUiRequiredException)
            {
                // This exception indicates that a new token
                // can only be obtained by invoking the on-behalf-of
                // flow. "UiRequired" isn't really accurate since the OBO
                // flow doesn't involve UI.
                // Catching the exception so code will continue to the
                // AcquireTokenOnBehalfOf call below.
            }
            catch (Exception exception)
            {
                _logger.LogError(exception, "Error getting access token via on-behalf-of flow");
                return null;
            }

            try
            {
                _logger.LogInformation("Token not found in cache, attempting OBO flow");

                // Use the token sent by the calling client as a
                // user assertion
                var userAssertion = new UserAssertion(_tokenResult.Token);

                // Invoke on-behalf-of flow
                var result = await _msalClient
                .AcquireTokenOnBehalfOf(_scopes, userAssertion)
                .ExecuteAsync();

                _logger.LogInformation($"User access token: {result.AccessToken}");
                return result.AccessToken;
            }
            catch (Exception exception)
            {
                _logger.LogError(exception, "Error getting access token from cache");
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
