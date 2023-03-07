// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

using System.Net;
using System.Text.Json;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Me.Messages.Item;
using Microsoft.Kiota.Serialization.Json;
using GraphSampleFunctions.Services;

namespace GraphSampleFunctions
{
    public class Notify
    {
        private readonly IGraphClientService _graphClientService;
        private readonly ILogger _logger;

        public Notify(
            IGraphClientService graphClientService,
            ILoggerFactory loggerFactory)
        {
            _graphClientService = graphClientService;
            _logger = loggerFactory.CreateLogger<Notify>();
        }

        [Function("Notify")]
        public async Task<HttpResponseData> RunAsync(
            [HttpTrigger(AuthorizationLevel.Function, "post")] HttpRequestData req)
        {
            _logger.LogInformation("Notify function triggered.");

            // Is this a validation request?
            // https://learn.microsoft.com/graph/webhooks#notification-endpoint-validation
            if (req.FunctionContext.BindingContext.BindingData
                .TryGetValue("validationToken", out object? validationToken))
            {
                // Return the validation token in a plain text body
                var validationResponse = req.CreateResponse(HttpStatusCode.OK);
                validationResponse.Headers.Add("Content-Type", "text/plain; charset=utf-8");
                validationResponse.WriteString(validationToken?.ToString() ?? string.Empty);
                return validationResponse;
            }

            var graphClient = _graphClientService.GetAppGraphClient();
            if (graphClient == null)
            {
                _logger.LogError("Could not create a Graph client for the app");
                return req.CreateResponse(HttpStatusCode.InternalServerError);
            }

            var payload = JsonDocument.Parse(req.Body);
            var node = new JsonParseNode(payload.RootElement);
            var notifications = node.GetObjectValue<ChangeNotificationCollectionResponse>(
                ChangeNotificationCollectionResponse.CreateFromDiscriminatorValue);
            if (notifications?.Value != null)
            {
                foreach (var notification in notifications?.Value!)
                {
                    await ProcessNotificationAsync(graphClient, notification);
                }
            }

            // Return 202 per docs
            return req.CreateResponse(HttpStatusCode.Accepted);
        }

        private async Task ProcessNotificationAsync(
            GraphServiceClient graphClient,
            ChangeNotification notification)
        {
            // The resource field in the notification has the URL to the
            // message, including the user ID and message ID. Since we
            // have the URL, use a MessageRequestBuilder instead of the fluent
            // API
            var msgRequestBuilder = new MessageItemRequestBuilder(
                $"https://graph.microsoft.com/v1.0/{notification.Resource}",
                graphClient.RequestAdapter);

            var message = await msgRequestBuilder.GetAsync(config =>
            {
                config.QueryParameters.Select = new[] { "subject" };
            });

            _logger.LogInformation($"The following message was {notification.ChangeType}:");
            _logger.LogInformation($"Subject: {message?.Subject}, ID: {message?.Id}");
        }
    }
}
