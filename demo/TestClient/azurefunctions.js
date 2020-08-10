// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// <getLatestMessageSnippet>
async function getLatestMessage() {
  const token = await getToken();
  if (!token) {
    throw new Error('Could not retrieve token for user.');
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
