// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

// <SetSubscriptionPayloadSnippet>
namespace GraphTutorial.Models
{
    // Class to represent the payload sent to the
    // SetSubscription function
    public class SetSubscriptionPayload
    {
        // "subscribe" or "unsubscribe"
        public string RequestType { get;set; }
        // If unsubscribing, the subscription to delete
        public string SubscriptionId { get;set; }
        // If subscribing, the user ID to subscribe to
        // Can be object ID of user, or userPrincipalName
        public string UserId { get;set; }
    }
}
// </SetSubscriptionPayloadSnippet>
