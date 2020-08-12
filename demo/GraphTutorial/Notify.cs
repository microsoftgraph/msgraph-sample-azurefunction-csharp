// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

// <NotifySnippet>
using GraphTutorial.Models;
using GraphTutorial.Services;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using System.IO;
using System.Text.Json;
using System.Threading.Tasks;
using System.Web.Http;

namespace GraphTutorial
{
    public class Notify
    {
        public static readonly string ClientState = "GraphTutorialState";
        private IConfiguration _config;
        private IGraphClientService _clientService;

        public Notify(IConfiguration config, IGraphClientService clientService)
        {
            _config = config;
            _clientService = clientService;
        }

        [FunctionName("Notify")]
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

            // Is this a validation request?
            // https://docs.microsoft.com/graph/webhooks#notification-endpoint-validation
            string validationToken = req.Query["validationToken"];
            if (!string.IsNullOrEmpty(validationToken))
            {
                // Because validationToken is a string, OkObjectResult
                // will return a text/plain response body, which is
                // required for validation
                return new OkObjectResult(validationToken);
            }

            // Not a validation request, process the body
            var requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            log.LogInformation($"Change notification payload: {requestBody}");

            var jsonOptions = new JsonSerializerOptions
            {
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase
            };

            // Deserialize the JSON payload into a list of ChangeNotification
            // objects
            var notifications = JsonSerializer.Deserialize<NotificationList>(requestBody, jsonOptions);

            foreach (var notification in notifications.Value)
            {
                if (notification.ClientState == ClientState)
                {
                    // Process each notification
                    await ProcessNotification(notification, log);
                }
                else
                {
                    log.LogInformation($"Notification received with unexpected client state: {notification.ClientState}");
                }
            }

            // Return 202 per docs
            return new AcceptedResult();
        }

        private async Task ProcessNotification(ChangeNotification notification, ILogger log)
        {
            var graphClient = _clientService.GetAppGraphClient(log);

            // The resource field in the notification has the URL to the
            // message, including the user ID and message ID. Since we
            // have the URL, use a MessageRequestBuilder instead of the fluent
            // API
            var msgRequestBuilder = new MessageRequestBuilder(
                $"https://graph.microsoft.com/v1.0/{notification.Resource}",
                graphClient);

            var message = await msgRequestBuilder.Request()
                .Select(m => new
                {
                    m.Subject
                })
                .GetAsync();

            log.LogInformation($"The following message was {notification.ChangeType}:");
            log.LogInformation($"Subject: {message.Subject}, ID: {message.Id}");
        }
    }
}
// </NotifySnippet>
