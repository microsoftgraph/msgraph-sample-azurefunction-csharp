// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

// <SetSubscriptionSnippet>

using System;
using System.IO;
using System.Net;
using System.Text.Json;
using System.Threading.Tasks;
using GraphTutorial.Authentication;
using GraphTutorial.Models;
using GraphTutorial.Services;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;

namespace GraphTutorial
{
    public class SetSubscription
    {
        private IConfiguration _config;
        private IGraphClientService _clientService;

        public SetSubscription(IConfiguration config, IGraphClientService clientService)
        {
            _config = config;
            _clientService = clientService;
        }


        [Function("SetSubscription")]
        public async Task<HttpResponseData> Run([HttpTrigger(AuthorizationLevel.Function, "get", "post")] HttpRequestData req,
            FunctionContext executionContext)
        {
            var logger = executionContext.GetLogger("SetSubscription");

            // Check configuration
            if (string.IsNullOrEmpty(_config["webHookId"]) ||
                string.IsNullOrEmpty(_config["webHookSecret"]) ||
                string.IsNullOrEmpty(_config["tenantId"]) ||
                string.IsNullOrEmpty(_config["apiFunctionId"]))
            {
                logger.LogError("Invalid app settings configured");
                return req.CreateResponse(HttpStatusCode.InternalServerError);
            }

            var notificationHost = _config["ngrokUrl"];
            if (string.IsNullOrEmpty(notificationHost))
            {
                notificationHost = req.Url.Host;
            }

            // Validate the bearer token
            var validationResult = await TokenValidation.ValidateAuthorizationHeader(
                req, _config["tenantId"], _config["apiFunctionId"], logger);

            // If token wasn't returned it isn't valid
            if (validationResult == null) {
                return req.CreateResponse(HttpStatusCode.Unauthorized);
            }

            var requestBody = await new StreamReader(req.Body).ReadToEndAsync();

            var jsonOptions = new JsonSerializerOptions
            {
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase
            };

            // Deserialize the JSON payload into a SetSubscriptionPayload object
            var payload = JsonSerializer.Deserialize<SetSubscriptionPayload>(requestBody, jsonOptions);

            if (payload == null)
            {
                var response = req.CreateResponse(HttpStatusCode.BadRequest);
                response.WriteString("Invalid request payload");
                return response;
            }

            // Initialize Graph client
            var graphClient = _clientService.GetAppGraphClient(logger);

            if (payload.RequestType.ToLower() == "subscribe")
            {
                if (string.IsNullOrEmpty(payload.UserId))
                {
                    var response = req.CreateResponse(HttpStatusCode.BadRequest);
                    response.WriteString("Required fields in payload missing");
                    return response;
                }

                // Create a new subscription object
                var subscription = new Subscription
                {
                    ChangeType = "created,updated",
                    NotificationUrl = $"{notificationHost}/api/Notify",
                    Resource = $"/users/{payload.UserId}/mailfolders/inbox/messages",
                    ExpirationDateTime = DateTimeOffset.UtcNow.AddDays(2),
                    ClientState = Notify.ClientState
                };

                // POST /subscriptions
                var createdSubscription = await graphClient.Subscriptions
                    .Request()
                    .AddAsync(subscription);

                var okResponse = req.CreateResponse(HttpStatusCode.OK);
                await okResponse.WriteAsJsonAsync<Subscription>(createdSubscription);
                return okResponse;
            }
            else
            {
                if (string.IsNullOrEmpty(payload.SubscriptionId))
                {
                    var response = req.CreateResponse(HttpStatusCode.BadRequest);
                    response.WriteString("Subscription ID missing in payload");
                    return response;
                }

                // DELETE /subscriptions/subscriptionId
                await graphClient.Subscriptions[payload.SubscriptionId]
                    .Request()
                    .DeleteAsync();

                return req.CreateResponse(HttpStatusCode.Accepted);
            }
        }
    }
}
// </SetSubscriptionSnippet>
