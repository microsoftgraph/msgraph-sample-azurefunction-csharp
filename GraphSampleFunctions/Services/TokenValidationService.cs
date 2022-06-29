// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

using System.Globalization;
using System.IdentityModel.Tokens.Jwt;
using System.Net.Http.Headers;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.IdentityModel.Protocols;
using Microsoft.IdentityModel.Protocols.OpenIdConnect;
using Microsoft.IdentityModel.Tokens;

namespace GraphSampleFunctions.Services
{
    public class TokenValidationService : ITokenValidationService
    {
        private TokenValidationParameters? _validationParameters;
        private readonly IConfiguration _config;

        private readonly ILogger _logger;

        public TokenValidationService(IConfiguration config, ILoggerFactory loggerFactory)
        {
            _config = config;
            _logger = loggerFactory.CreateLogger<TokenValidationService>();
        }

        public async Task<string?> ValidateAuthorizationHeaderAsync(
            Microsoft.Azure.Functions.Worker.Http.HttpRequestData request)
        {
            // The incoming request should have an Authorization header
            if (request.Headers.TryGetValues("authorization", out IEnumerable<string>? authValues))
            {
                var authHeader = AuthenticationHeaderValue.Parse(authValues.ToArray().First());

                // Make sure that the value is "Bearer token-value"
                if (authHeader != null &&
                    string.Compare(authHeader.Scheme, "bearer", true, CultureInfo.InvariantCulture) == 0 &&
                    !string.IsNullOrEmpty(authHeader.Parameter))
                {
                    var validationParameters = await GetTokenValidationParametersAsync();
                    if (validationParameters == null)
                    {
                        return null;
                    }

                    var tokenHandler = new JwtSecurityTokenHandler();
                    try
                    {
                        // Validate the token
                        var result = tokenHandler.ValidateToken(authHeader.Parameter,
                            _validationParameters, out SecurityToken jwtToken);

                        // If ValidateToken did not throw an exception, token is valid.
                        return authHeader.Parameter;
                    }
                    catch (Exception exception)
                    {
                        _logger.LogError(exception, "Error validating bearer token");
                    }
                }
            }

            return null;
        }

        private async Task<TokenValidationParameters?> GetTokenValidationParametersAsync()
        {
            if (_validationParameters == null)
            {
                // Get tenant ID and client ID
                var tenantId = _config["tenantId"];
                var apiClientId = _config["apiClientId"];
                if (string.IsNullOrEmpty(tenantId) ||
                    string.IsNullOrEmpty(apiClientId))
                {
                    _logger.LogError("Required settings missing: 'tenantId' and 'apiClientId'.");
                    return null;
                }

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
                    ValidAudience = apiClientId,
                    ValidateIssuer = true,
                    // Use the issuer retrieved from Azure
                    ValidIssuer = config.Issuer,
                    ValidateLifetime = true
                };
            }

            return _validationParameters;
        }
    }
}
