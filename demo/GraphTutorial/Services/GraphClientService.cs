// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.
using GraphTutorial.Authentication;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Identity.Client;
using Microsoft.Graph;

namespace GraphTutorial.Services
{
    // Service added via dependency injection
    // Used to get an authenticated Graph client
    public class GraphClientService : IGraphClientService
    {
        // <UserGraphClientMembers>
        // Configuration
        private IConfiguration _config;

        // Single MSAL client object used for all user-related
        // requests. Making this a "singleton" here because the sample
        // uses the default in-memory token cache.
        private IConfidentialClientApplication _userMsalClient;
        // </UserGraphClientMembers>

        // <AppGraphClientMembers>
        private GraphServiceClient _appGraphClient;
        // </AppGraphClientMembers>

        // <UserGraphClientFunctions>
        public GraphClientService(IConfiguration config)
        {
          _config = config;
        }

        public GraphServiceClient GetUserGraphClient(TokenValidationResult validation, string[] scopes, ILogger logger)
        {
            // Only create the MSAL client once
            if (_userMsalClient == null)
            {
                _userMsalClient = ConfidentialClientApplicationBuilder
                    .Create(_config["apiFunctionId"])
                    .WithAuthority(AadAuthorityAudience.AzureAdMyOrg, true)
                    .WithTenantId(_config["tenantId"])
                    .WithClientSecret(_config["apiFunctionSecret"])
                    .Build();
            }

            // Create a new OBO auth provider for the specific user
            var authProvider = new OnBehalfOfAuthProvider(_userMsalClient, validation, scopes, logger);

            // Return a GraphServiceClient initialized with the auth provider
            return new GraphServiceClient(authProvider);
        }
        // </UserGraphClientFunctions>

        // <AppGraphClientFunctions>
        public GraphServiceClient GetAppGraphClient(ILogger logger)
        {
            if (_appGraphClient == null)
            {
                // Create a client credentials auth provider
                var authProvider = new ClientCredentialsAuthProvider(
                    _config["webHookId"],
                    _config["webHookSecret"],
                    _config["tenantId"],
                    // The https://graph.microsoft.com/.default scope
                    // is required for client credentials. It requests
                    // all of the permissions that are explicitly set on
                    // the app registration
                    new[] { "https://graph.microsoft.com/.default" },
                    logger);

                _appGraphClient = new GraphServiceClient(authProvider);
            }

            return _appGraphClient;
        }
        // </AppGraphClientFunctions>
    }
}
