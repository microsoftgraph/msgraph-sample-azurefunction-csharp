// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// <getLatestMessageSnippet>
async function getLatestMessage() {
  const token = await getToken();
  if (!token) {
    updatePage(Views.error, {
      message: 'Could not retrieve token for user'
    });
    return;
  }

  try {
    const response = await fetch('http://localhost:7071/api/GetMyNewestMessage', {
      headers: {
        Authorization: `Bearer ${token}`
      }
    });

    const message = await response.json();

    updatePage(Views.message, message);
  } catch (error) {
    updatePage(Views.error, {
      message: 'Error getting message',
      debug: error
    });
  }
}
// </getLatestMessageSnippet>

// <createSubscriptionSnippet>
async function createSubscription()
{
  // Get the user to subscribe for
  const userId = document.getElementById('subscribe-user').value;
  if (!userId) {
    updatePage(Views.error, {
      message: 'Please provide a user ID or userPrincipalName'
    });
    return;
  }

  const token = await getToken();
  if (!token) {
    updatePage(Views.error, {
      message: 'Could not retrieve token for user'
    });
    return;
  }

  // Build the JSON payload for the subscribe request
  const payload = {
    requestType: 'subscribe',
    userId: userId
  };

  const response = await fetch('http://localhost:7071/api/SetSubscription', {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token}`
    },
    body: JSON.stringify(payload)
  });

  if (response.ok) {
    // Get the new subscription from the response
    const subscription = await response.json();

    // Add the new subscription to the array of subscriptions
    // in the session
    let existingSubscriptions =
      JSON.parse(sessionStorage.getItem('graph-subscriptions')) || [];

    existingSubscriptions.push({
      userId: userId,
      subscriptionId: subscription.id
    });

    sessionStorage.setItem('graph-subscriptions',
      JSON.stringify(existingSubscriptions));

    // Refresh the subscriptions page to display the new
    // subscription
    updatePage(Views.subscriptions);
    return;
  }

  updatePage(Views.error, {
    message: `Call to SetSubscription returned ${response.status}`,
    debug: response.statusText
  });
}
// </createSubscriptionSnippet>

// <deleteSubscriptionSnippet>
async function deleteSubscription(subscriptionId) {
  const token = await getToken();
  if (!token) {
    updatePage(Views.error, {
      message: 'Could not retrieve token for user'
    });
    return;
  }

  // Build the JSON payload for the unsubscribe request
  const payload = {
    requestType: 'unsubscribe',
    subscriptionId: subscriptionId
  };

  const response = await fetch('http://localhost:7071/api/SetSubscription', {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token}`
    },
    body: JSON.stringify(payload)
  });

  if (response.ok) {
    // Remove the subscription from the array
    let existingSubscriptions =
      JSON.parse(sessionStorage.getItem('graph-subscriptions')) || [];

    const subscriptionIndex = existingSubscriptions.findIndex((item) => {
      return item.subscriptionId === subscriptionId;
    });

    existingSubscriptions.splice(subscriptionIndex, 1);

    sessionStorage.setItem('graph-subscriptions',
      JSON.stringify(existingSubscriptions));

    // Refresh the subscriptions page
    updatePage(Views.subscriptions);
    return;
  }

  updatePage(Views.error, {
    message: `Call to SetSubscription returned ${response.status}`,
    debug: response.statusText
  });
}
// </deleteSubscriptionSnippet>