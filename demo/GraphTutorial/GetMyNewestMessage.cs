// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

// <GetMyNewestMessageSnippet>
using System.Net;
using System.Threading.Tasks;
using GraphTutorial.Authentication;
using GraphTutorial.Services;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;

namespace GraphTutorial
{
    public class GetMyNewestMessage
    {
        private IConfiguration _config;
        private IGraphClientService _clientService;

        public GetMyNewestMessage(IConfiguration config, IGraphClientService clientService)
        {
            _config = config;
            _clientService = clientService;
        }

        [Function("GetMyNewestMessage")]
        public async Task<HttpResponseData> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get")] HttpRequestData req,
            FunctionContext executionContext)
        {
            var logger = executionContext.GetLogger("GetMyNewestMessage");

            // Check configuration
            if (string.IsNullOrEmpty(_config["apiFunctionId"]) ||
                string.IsNullOrEmpty(_config["apiFunctionSecret"]) ||
                string.IsNullOrEmpty(_config["tenantId"]))
            {
                logger.LogError("Invalid app settings configured");
                return req.CreateResponse(HttpStatusCode.InternalServerError);
            }

            // Validate the bearer token
            var validationResult = await TokenValidation.ValidateAuthorizationHeader(
                req, _config["tenantId"], _config["apiFunctionId"], logger);

            // If token wasn't returned it isn't valid
            if (validationResult == null)
            {
                return req.CreateResponse(HttpStatusCode.Unauthorized);
            }

            // Initialize a Graph client for this user
            var graphClient = _clientService.GetUserGraphClient(validationResult,
                new[] { "https://graph.microsoft.com/.default" }, logger);

            // Get the user's newest message in inbox
            // GET /me/mailfolders/inbox/messages
            var messagePage = await graphClient.Me
                .MailFolders
                .Inbox
                .Messages
                .Request()
                // Limit the fields returned
                .Select(m => new
                {
                    m.From,
                    m.ReceivedDateTime,
                    m.Subject
                })
                // Sort by received time, newest on top
                .OrderBy("receivedDateTime DESC")
                // Only get back one message
                .Top(1)
                .GetAsync();

            if (messagePage.CurrentPage.Count > 0)
            {
                var response = req.CreateResponse(HttpStatusCode.OK);
                // Return the message in the response
                await response.WriteAsJsonAsync<Message>(messagePage.CurrentPage[0]);
                return response;
            }

            return req.CreateResponse(HttpStatusCode.NoContent);
        }
    }
}
// </GetMyNewestMessageSnippet>
