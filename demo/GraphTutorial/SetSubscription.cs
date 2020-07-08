// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

// <SetSubscriptionSnippet>
using GraphTutorial.Authentication;
using GraphTutorial.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using System;
using System.IO;
using System.Net.Http.Headers;
using System.Text.Json;
using System.Threading.Tasks;
using System.Web.Http;

namespace GraphTutorial
{
    public class SetSubscription
    {
        private IConfiguration _config;

        public SetSubscription(IConfiguration config)
        {
            _config = config;
        }

        [FunctionName("SetSubscription")]
        public async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            // Check configuration
            if (string.IsNullOrEmpty(_config["webHookId"]) ||
                string.IsNullOrEmpty(_config["webHookSecret"]) ||
                string.IsNullOrEmpty(_config["tenantId"]))
            {
                log.LogError("Invalid app settings configured");
                return new InternalServerErrorResult();
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

            // Initialize an auth provider
            var authProvider = new ClientCredentialsAuthProvider(
                _config["webHookId"],
                _config["webHookSecret"],
                _config["tenantId"],
                // The https://graph.microsoft.com/.default scope
                // is required for client credentials. It requests
                // all of the permissions that are explicitly set on
                // the app registration
                new[] { "https://graph.microsoft.com/.default" },
                log);

            var appToken = await authProvider.GetAccessToken();

            // Initialize Graph client
            var graphClient = new GraphServiceClient(new DelegateAuthenticationProvider(
                (requestMessage) => {
                    requestMessage.Headers.Authorization =
                        new AuthenticationHeaderValue("bearer", appToken);
                    return Task.FromResult(0);
                }));

            if (payload.RequestType.ToLower() == "subscribe")
            {
                if (string.IsNullOrEmpty(payload.UserId) ||
                    string.IsNullOrEmpty(payload.NgrokProxy))
                {
                    return new BadRequestErrorMessageResult("Required fields in payload missing");
                }

                // Create a new subscription object
                var subscription = new Subscription
                {
                    ChangeType = "created,updated",
                    NotificationUrl = $"{payload.NgrokProxy}/api/Notify",
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
