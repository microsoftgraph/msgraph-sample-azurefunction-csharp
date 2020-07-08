// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

// <ChangeNotificationSnippet>
using System;

namespace GraphTutorial.Models
{
    // Represents a change notification payload
    // https://docs.microsoft.com/graph/api/resources/changenotification?view=graph-rest-1.0
    public class ChangeNotification
    {
        public string ChangeType { get;set; }
        public string ClientState { get;set; }
        public string Resource { get;set; }
        public ResourceData ResourceData { get;set; }
        public DateTime SubscriptionExpirationDateTime { get;set; }
        public string SubscriptionId { get;set; }
        public string TenantId { get;set; }
    }
}
// </ChangeNotificationSnippet>
