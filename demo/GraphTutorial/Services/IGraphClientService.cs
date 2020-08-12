// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

// <IGraphClientServiceSnippet>
using GraphTutorial.Authentication;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;

namespace GraphTutorial.Services
{
    public interface IGraphClientService
    {
        GraphServiceClient GetUserGraphClient(
          TokenValidationResult validation,
          string[] scopes,
          ILogger logger);

        GraphServiceClient GetAppGraphClient(ILogger logger);
    }
}
// </IGraphClientServiceSnippet>
