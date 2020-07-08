// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

// <NotificationListSnippet>
namespace GraphTutorial.Models
{
    // Class representing an array of notifications
    // in a notification payload
    public class NotificationList
    {
        public ChangeNotification[] Value { get;set; }
    }
}
// </NotificationListSnippet>
