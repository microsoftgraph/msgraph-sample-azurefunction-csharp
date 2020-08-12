// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

// <TokenValidationResultSnippet>
namespace GraphTutorial.Authentication
{
    public class TokenValidationResult
    {
        // MSAL account ID - used to access the token
        // cache
        public string MsalAccountId { get; private set; }

        // The extracted token - used to build user assertion
        // for OBO flow
        public string Token { get; private set; }

        public TokenValidationResult(string msalAccountId, string token)
        {
            MsalAccountId = msalAccountId;
            Token = token;
        }
    }
}
// </TokenValidationResultSnippet>
