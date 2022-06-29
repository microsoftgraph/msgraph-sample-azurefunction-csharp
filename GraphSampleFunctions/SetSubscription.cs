// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

using System.Globalization;
using System.Net;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using GraphSampleFunctions.Models;
using GraphSampleFunctions.Services;

namespace GraphSampleFunctions
{
    public class SetSubscription
    {
        public static readonly string ClientState = "GraphSampleFunctionState";
        private readonly ITokenValidationService _tokenValidationService;
        private readonly IGraphClientService _graphClientService;
        private readonly IConfiguration _config;
        private readonly ILogger _logger;

        public SetSubscription(
            ITokenValidationService tokenValidationService,
            IGraphClientService graphClientService,
            IConfiguration config,
            ILoggerFactory loggerFactory)
        {
            _tokenValidationService = tokenValidationService;
            _graphClientService = graphClientService;
            _config = config;
            _logger = loggerFactory.CreateLogger<SetSubscription>();
        }

        [Function("SetSubscription")]
        public async Task<HttpResponseData> RunAsync(
            [HttpTrigger(AuthorizationLevel.Function, "post")] HttpRequestData req)
        {
            _logger.LogInformation("SetSubscription function triggered.");

            // Validate the bearer token
            var validationResult = await _tokenValidationService
                .ValidateAuthorizationHeaderAsync(req);
            if (validationResult == null)
            {
                // If token wasn't returned it isn't valid
                return req.CreateResponse(HttpStatusCode.Unauthorized);
            }

            var graphClient = _graphClientService.GetAppGraphClient();
            if (graphClient == null)
            {
                _logger.LogError("Could not create a Graph client for the app");
                return req.CreateResponse(HttpStatusCode.InternalServerError);
            }

            // Get the POST body
            var payload = graphClient.HttpProvider.Serializer
                .DeserializeObject<SetSubscriptionPayload>(req.Body);
            if (payload == null)
            {
                var response = req.CreateResponse(HttpStatusCode.BadRequest);
                response.WriteString("Invalid request payload");
                return response;
            }

            if (string.Compare(payload.RequestType, "subscribe", true, CultureInfo.InvariantCulture) == 0)
            {
                if (string.IsNullOrEmpty(payload.UserId))
                {
                    var badRequest = req.CreateResponse(HttpStatusCode.BadRequest);
                    badRequest.WriteString("Required field 'userId' missing in payload.");
                    return badRequest;
                }

                // Get ngrok URL if set (for local development)
                var notificationHost = _config["ngrokUrl"] ?? req.Url.Host;

                // Create a new subscription object
                var subscription = new Subscription
                {
                    ChangeType = "created,updated",
                    NotificationUrl = $"{notificationHost}/api/Notify",
                    Resource = $"/users/{payload.UserId}/mailFolders/inbox/messages",
                    ExpirationDateTime = DateTimeOffset.UtcNow.AddDays(2),
                    ClientState = SetSubscription.ClientState
                };

                _logger.LogInformation($"Creating subscription for user ${payload.UserId}");

                // POST /subscriptions
                var createdSubscription = await graphClient.Subscriptions
                    .Request()
                    .AddAsync(subscription);

                var response = req.CreateResponse(HttpStatusCode.OK);
                await response.WriteAsJsonAsync<Subscription>(createdSubscription);
                return response;
            }
            else
            {
                if (string.IsNullOrEmpty(payload.SubscriptionId))
                {
                    var badRequest = req.CreateResponse(HttpStatusCode.BadRequest);
                    badRequest.WriteString("Required field 'subscriptionId' missing in payload.");
                    return badRequest;
                }

                _logger.LogInformation($"Deleting subscription with ID {payload.SubscriptionId}");

                // DELETE /subscriptions/subscriptionId
                await graphClient.Subscriptions[payload.SubscriptionId]
                    .Request()
                    .DeleteAsync();

                return req.CreateResponse(HttpStatusCode.Accepted);
            }
        }
    }
}
