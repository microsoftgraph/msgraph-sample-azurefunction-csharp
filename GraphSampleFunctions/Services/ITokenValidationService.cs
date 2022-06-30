// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

using Microsoft.Azure.Functions.Worker.Http;

namespace GraphSampleFunctions.Services
{
    public interface ITokenValidationService
    {
        public Task<string?> ValidateAuthorizationHeaderAsync(
            HttpRequestData request);
    }
}
