// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

// <AuthProviderSnippet>
using Microsoft.Extensions.Logging;
using Microsoft.Identity.Client;
using System;
using System.Threading.Tasks;

namespace GraphTutorial.Authentication
{
    public class OnBehalfOfAuthProvider
    {
        private IConfidentialClientApplication _msalClient;
        private string[] _scopes;
        private ILogger _logger;

        public OnBehalfOfAuthProvider(string appId, string clientSecret, string tenantId, string[] scopes, ILogger logger)
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

        public async Task<string> GetAccessToken(string userToken)
        {
            try
            {
                // Use the token sent by the calling client as a
                // user assertion
                var userAssertion = new UserAssertion(userToken);

                // Invoke on-behalf-of flow
                var result = await _msalClient
                  .AcquireTokenOnBehalfOf(_scopes, userAssertion)
                  .ExecuteAsync();

                _logger.LogInformation($"Access token: {result.AccessToken}");

                return result.AccessToken;
            }
            catch (Exception exception)
            {
                _logger.LogError(exception, "Error getting access token");
                return null;
            }
        }
    }
}
// </AuthProviderSnippet>
