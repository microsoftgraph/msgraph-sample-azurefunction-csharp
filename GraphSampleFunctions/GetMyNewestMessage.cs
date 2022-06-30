// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

using System.Net;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using GraphSampleFunctions.Services;

namespace GraphSampleFunctions
{
    public class GetMyNewestMessage
    {
        private readonly ITokenValidationService _tokenValidationService;
        private readonly IGraphClientService _graphClientService;
        private readonly ILogger _logger;

        public GetMyNewestMessage(
            ITokenValidationService tokenValidationService,
            IGraphClientService graphClientService,
            ILoggerFactory loggerFactory)
        {
            _tokenValidationService = tokenValidationService;
            _graphClientService = graphClientService;
            _logger = loggerFactory.CreateLogger<GetMyNewestMessage>();
        }

        [Function("GetMyNewestMessage")]
        public async Task<HttpResponseData> RunAsync(
            [HttpTrigger(AuthorizationLevel.Function, "get")] HttpRequestData req)
        {
            _logger.LogInformation("GetMyNewMessage function triggered.");

            // Validate the bearer token
            var bearerToken = await _tokenValidationService
                .ValidateAuthorizationHeaderAsync(req);
            if (string.IsNullOrEmpty(bearerToken))
            {
                // If token wasn't returned it isn't valid
                return req.CreateResponse(HttpStatusCode.Unauthorized);
            }

            var graphClient = _graphClientService.GetUserGraphClient(bearerToken);
            if (graphClient == null)
            {
                _logger.LogError("Could not create a Graph client for the user");
                return req.CreateResponse(HttpStatusCode.InternalServerError);
            }

            // Get the user's newest message in inbox
            // GET /me/mailFolders/inbox/messages
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
                // Only get back one (the newest) message
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
