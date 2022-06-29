// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

namespace GraphSampleFunctions.Services
{
    public class TokenValidationResult
    {
        // User ID
        public string UserId { get; private set; }

        // The extracted token - used to build user assertion
        // for OBO flow
        public string Token { get; private set; }

        public TokenValidationResult(string userId, string token)
        {
            UserId = userId;
            Token = token;
        }
    }
}
