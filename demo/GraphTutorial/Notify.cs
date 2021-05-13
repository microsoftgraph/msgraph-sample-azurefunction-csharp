using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text.Json;
using System.Threading.Tasks;
using GraphTutorial.Models;
using GraphTutorial.Services;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;

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

        [Function("Notify")]
        public async Task<HttpResponseData> Run([HttpTrigger(AuthorizationLevel.Function, "get", "post")] HttpRequestData req,
            FunctionContext executionContext)
        {
            var logger = executionContext.GetLogger("Notify");

            // Check configuration
            if (string.IsNullOrEmpty(_config["webHookId"]) ||
                string.IsNullOrEmpty(_config["webHookSecret"]) ||
                string.IsNullOrEmpty(_config["tenantId"]))
            {
                logger.LogError("Invalid app settings configured");
                return req.CreateResponse(HttpStatusCode.InternalServerError);
            }

            // Is this a validation request?
            // https://docs.microsoft.com/graph/webhooks#notification-endpoint-validation
            if (executionContext.BindingContext.BindingData
                .TryGetValue("validationToken", out object validationToken))
            {
                // Because validationToken is a string, OkObjectResult
                // will return a text/plain response body, which is
                // required for validation
                var response = req.CreateResponse(HttpStatusCode.OK);
                response.Headers.Add("Content-Type", "text/plain; charset=utf-8");

                response.WriteString(validationToken.ToString());
                return response;
            }

            // Not a validation request, process the body
            var requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            logger.LogInformation($"Change notification payload: {requestBody}");

            var jsonOptions = new JsonSerializerOptions
            {
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase
            };

            // Deserialize the JSON payload into a list of ChangeNotificationPayload
            // objects
            var notifications = JsonSerializer.Deserialize<NotificationList>(requestBody, jsonOptions);

            foreach (var notification in notifications.Value)
            {
                if (notification.ClientState == ClientState)
                {
                    // Process each notification
                    await ProcessNotification(notification, logger);
                }
                else
                {
                    logger.LogInformation($"Notification received with unexpected client state: {notification.ClientState}");
                }
            }

            // Return 202 per docs
            return req.CreateResponse(HttpStatusCode.Accepted);
        }

        private async Task ProcessNotification(ChangeNotificationPayload notification, ILogger log)
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
