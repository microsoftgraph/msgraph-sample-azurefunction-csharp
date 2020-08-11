// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

// <TokenValidationSnippet>
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Microsoft.IdentityModel.Protocols;
using Microsoft.IdentityModel.Protocols.OpenIdConnect;
using Microsoft.IdentityModel.Tokens;
using System;
using System.IdentityModel.Tokens.Jwt;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace GraphTutorial.Authentication
{
    public static class TokenValidation
    {
        private static TokenValidationParameters _validationParameters = null;
        public static async Task<string> ValidateAuthorizationHeader(HttpRequest request,
                                                                     string tenantId,
                                                                     string expectedAudience,
                                                                     ILogger log)
        {
            // Check for Authorization header
            if (request.Headers.ContainsKey("authorization"))
            {
                var authHeader = AuthenticationHeaderValue.Parse(request.Headers["authorization"]);

                if (authHeader != null &&
                    authHeader.Scheme.ToLower() == "bearer" &&
                    !string.IsNullOrEmpty(authHeader.Parameter))
                {
                    if (_validationParameters == null)
                    {
                        // Load the tenant-specific OpenID config from Azure
                        var configManager = new ConfigurationManager<OpenIdConnectConfiguration>(
                        $"https://login.microsoftonline.com/{tenantId}/.well-known/openid-configuration",
                        new OpenIdConnectConfigurationRetriever());

                        var config = await configManager.GetConfigurationAsync();

                        _validationParameters = new TokenValidationParameters
                        {
                            // Use signing keys retrieved from Azure
                            IssuerSigningKeys = config.SigningKeys,
                            ValidateAudience = true,
                            // Audience MUST be the app ID for the Web API
                            ValidAudience = expectedAudience,
                            ValidateIssuer = true,
                            // Use the issuer retrieved from Azure
                            ValidIssuer = config.Issuer,
                            ValidateLifetime = true
                        };
                    }

                    var tokenHandler = new JwtSecurityTokenHandler();

                    SecurityToken jwtToken;
                    try
                    {
                        // Validate the token
                        var result = tokenHandler.ValidateToken(authHeader.Parameter,
                            _validationParameters, out jwtToken);

                        // If ValidateToken did not throw an exception, token is valid.
                        // Return the token
                        return authHeader.Parameter;
                    }
                    catch (Exception exception)
                    {
                        log.LogError(exception, "Error validating bearer token");
                    }
                }
            }

            return null;
        }
    }
}
// </TokenValidationSnippet>
