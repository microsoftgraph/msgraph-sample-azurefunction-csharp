// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

// <GetMyNewestMessageSnippet>
using GraphTutorial.Authentication;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using System.Web.Http;

namespace GraphTutorial
{
    public class GetMyNewestMessage
    {
        private IConfiguration _config;

        public GetMyNewestMessage(IConfiguration config)
        {
            _config = config;
        }

        [FunctionName("GetMyNewestMessage")]
        public async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", Route = null)] HttpRequest req,
            ILogger log)
        {
            // Check configuration
            if (string.IsNullOrEmpty(_config["apiFunctionId"]) ||
                string.IsNullOrEmpty(_config["apiFunctionSecret"]) ||
                string.IsNullOrEmpty(_config["tenantId"]))
            {
                log.LogError("Invalid app settings configured");
                return new InternalServerErrorResult();
            }

            // Validate the bearer token
            var bearerToken = await TokenValidation.ValidateAuthorizationHeader(
                req, _config["tenantId"], _config["apiFunctionId"], log);

            // If token wasn't returned it isn't valid
            if (string.IsNullOrEmpty(bearerToken))
            {
                return new UnauthorizedResult();
            }

            // Initialize the auth provider
            var authProvider = new OnBehalfOfAuthProvider(
                _config["apiFunctionId"],
                _config["apiFunctionSecret"],
                _config["tenantId"],
                new[] { "https://graph.microsoft.com/.default" },
                log);

            // Exchange the token sent by client for a Graph-compatible token
            var graphToken = await authProvider.GetAccessToken(bearerToken);

            // Initialize a Graph client that uses the token
            var graphClient = new GraphServiceClient(new DelegateAuthenticationProvider(
                (requestMessage) => {
                    requestMessage.Headers.Authorization =
                        new AuthenticationHeaderValue("bearer", graphToken);
                    return Task.FromResult(0);
                }));

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

            if (messagePage.CurrentPage.Count < 1)
            {
                return new OkObjectResult(null);
            }

            // Return the message in the response
            return new OkObjectResult(messagePage.CurrentPage[0]);
        }
    }
}
// </GetMyNewestMessageSnippet>
