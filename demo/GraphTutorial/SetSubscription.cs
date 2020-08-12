// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

// <SetSubscriptionSnippet>
using GraphTutorial.Authentication;
using GraphTutorial.Models;
using GraphTutorial.Services;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using System;
using System.IO;
using System.Text.Json;
using System.Threading.Tasks;
using System.Web.Http;

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

        [FunctionName("SetSubscription")]
        public async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            // Check configuration
            if (string.IsNullOrEmpty(_config["webHookId"]) ||
                string.IsNullOrEmpty(_config["webHookSecret"]) ||
                string.IsNullOrEmpty(_config["tenantId"]) ||
                string.IsNullOrEmpty(_config["apiFunctionId"]))
            {
                log.LogError("Invalid app settings configured");
                return new InternalServerErrorResult();
            }

            var notificationHost = _config["ngrokUrl"];
            if (string.IsNullOrEmpty(notificationHost))
            {
                notificationHost = req.Host.Value;
            }

            // Validate the bearer token
            var validationResult = await TokenValidation.ValidateAuthorizationHeader(
                req, _config["tenantId"], _config["apiFunctionId"], log);

            // If token wasn't returned it isn't valid
            if (validationResult == null) {
                return new UnauthorizedResult();
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
                return new BadRequestErrorMessageResult("Invalid request payload");
            }

            // Initialize Graph client
            var graphClient = _clientService.GetAppGraphClient(log);

            if (payload.RequestType.ToLower() == "subscribe")
            {
                if (string.IsNullOrEmpty(payload.UserId))
                {
                    return new BadRequestErrorMessageResult("Required fields in payload missing");
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

                return new OkObjectResult(createdSubscription);
            }
            else
            {
                if (string.IsNullOrEmpty(payload.SubscriptionId))
                {
                    return new BadRequestErrorMessageResult("Subscription ID missing in payload");
                }

                // DELETE /subscriptions/subscriptionId
                await graphClient.Subscriptions[payload.SubscriptionId]
                    .Request()
                    .DeleteAsync();

                return new AcceptedResult();
            }
        }
    }
}
// </SetSubscriptionSnippet>
