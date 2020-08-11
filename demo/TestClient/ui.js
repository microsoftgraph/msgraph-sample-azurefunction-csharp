// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// <uiJsSnippet>
// Select DOM elements to work with
const authenticatedNav = document.getElementById('authenticated-nav');
const accountNav = document.getElementById('account-nav');
const mainContainer = document.getElementById('main-container');

const Views = { error: 1, home: 2, message: 3, subscriptions: 4 };

// Helper function to create an element, set class, and add text
function createElement(type, className, text) {
  const element = document.createElement(type);
  element.className = className;

  if (text) {
    const textNode = document.createTextNode(text);
    element.appendChild(textNode);
  }

  return element;
}

// Show the navigation items that should only show if
// the user is signed in
function showAuthenticatedNav(user, view) {
  authenticatedNav.innerHTML = '';

  if (user) {
    // Add message link
    const messageNav = createElement('li', 'nav-item');

    const messageLink = createElement('button',
      `btn btn-link nav-link${view === Views.message ? ' active' : '' }`,
      'Latest Message');
    messageLink.setAttribute('onclick', 'getLatestMessage();');
    messageNav.appendChild(messageLink);

    authenticatedNav.appendChild(messageNav);

    // Add subscriptions link
    const subscriptionNav = createElement('li', 'nav-item');

    const subscriptionLink = createElement('button',
      `btn btn-link nav-link${view === Views.message ? ' active' : '' }`,
      'Subscriptions');
    subscriptionLink.setAttribute('onclick', `updatePage(${Views.subscriptions});`);
    subscriptionNav.appendChild(subscriptionLink);

    authenticatedNav.appendChild(subscriptionNav);
  }
}

// Show the sign in button or the dropdown to sign-out
function showAccountNav(user) {

  accountNav.innerHTML = '';

  if (user) {
    // Show the "signed-in" nav
    accountNav.className = 'nav-item dropdown';

    const dropdown = createElement('a', 'nav-link dropdown-toggle');
    dropdown.setAttribute('data-toggle', 'dropdown');
    dropdown.setAttribute('role', 'button');
    accountNav.appendChild(dropdown);

    const userIcon = createElement('i',
      'far fa-user-circle fa-lg rounded-circle align-self-center');
    userIcon.style.width = '32px';
    dropdown.appendChild(userIcon);

    const menu = createElement('div', 'dropdown-menu dropdown-menu-right');
    dropdown.appendChild(menu);

    const userName = createElement('h5', 'dropdown-item-text mb-0', user);
    menu.appendChild(userName);

    const divider = createElement('div', 'dropdown-divider');
    menu.appendChild(divider);

    const signOutButton = createElement('button', 'dropdown-item', 'Sign out');
    signOutButton.setAttribute('onclick', 'signOut();');
    menu.appendChild(signOutButton);
  } else {
    // Show a "sign in" button
    accountNav.className = 'nav-item';

    const signInButton = createElement('button', 'btn btn-link nav-link', 'Sign in');
    signInButton.setAttribute('onclick', 'signIn();');
    accountNav.appendChild(signInButton);
  }
}

// Renders the home view
function showWelcomeMessage(user) {
  // Create jumbotron
  const jumbotron = createElement('div', 'jumbotron');

  const heading = createElement('h1', null, 'Azure Functions Graph Tutorial Test Client');
  jumbotron.appendChild(heading);

  const lead = createElement('p', 'lead',
    'This sample app is used to test the Azure Functions in the Azure Functions Graph Tutorial');
  jumbotron.appendChild(lead);

  if (user) {
    // Welcome the user by name
    const welcomeMessage = createElement('h4', null, `Welcome ${user}!`);
    jumbotron.appendChild(welcomeMessage);

    const callToAction = createElement('p', null,
      'Use the navigation bar at the top of the page to get started.');
    jumbotron.appendChild(callToAction);
  } else {
    // Show a sign in button in the jumbotron
    const signInButton = createElement('button', 'btn btn-primary btn-large',
      'Click here to sign in');
    signInButton.setAttribute('onclick', 'signIn();')
    jumbotron.appendChild(signInButton);
  }

  mainContainer.innerHTML = '';
  mainContainer.appendChild(jumbotron);
}

// Renders an email message
function showLatestMessage(message) {
  // Show message
  const messageCard = createElement('div', 'card');

  const cardBody = createElement('div', 'card-body');
  messageCard.appendChild(cardBody);

  const subject = createElement('h1', 'card-title', `${message.subject || '(No subject)'}`);
  cardBody.appendChild(subject);

  const fromLine = createElement('div', 'd-flex');
  cardBody.appendChild(fromLine);

  const fromLabel = createElement('div', 'mr-3');
  fromLabel.appendChild(createElement('strong', '', 'From:'));
  fromLine.appendChild(fromLabel);

  fromLine.appendChild(createElement('div', '', message.from.emailAddress.name));

  const receivedLine = createElement('div', 'd-flex');
  cardBody.appendChild(receivedLine);

  const receivedLabel = createElement('div', 'mr-3');
  receivedLabel.appendChild(createElement('strong', '', 'Received:'));
  receivedLine.appendChild(receivedLabel);

  receivedLine.appendChild(createElement('div', '', message.receivedDateTime));

  mainContainer.innerHTML = '';
  mainContainer.appendChild(messageCard);
}

// Renders current subscriptions from the session, and allows the user
// to add new subscriptions
function showSubscriptions() {
  const subscriptions = JSON.parse(sessionStorage.getItem('graph-subscriptions'));

  // Show new subscription form
  const form = createElement('form', 'form-inline mb-3');

  const userInput = createElement('input', 'form-control mb-2 mr-2 flex-grow-1');
  userInput.setAttribute('id', 'subscribe-user');
  userInput.setAttribute('type', 'text');
  userInput.setAttribute('placeholder', 'User to subscribe to (user ID or UPN)');
  form.appendChild(userInput);

  const subscribeButton = createElement('button', 'btn btn-primary mb-2', 'Subscribe');
  subscribeButton.setAttribute('type', 'button');
  subscribeButton.setAttribute('onclick', 'createSubscription();');
  form.appendChild(subscribeButton);

  const card = createElement('div', 'card');

  const cardBody = createElement('div', 'card-body');
  card.appendChild(cardBody);

  cardBody.appendChild(createElement('h2', 'card-title mb-4', 'Existing subscriptions'));

  const subscriptionTable = createElement('table', 'table');
  cardBody.appendChild(subscriptionTable);

  const thead = createElement('thead', '');
  subscriptionTable.appendChild(thead);

  const theadRow = createElement('tr', '');
  thead.appendChild(theadRow);

  theadRow.appendChild(createElement('th', ''));
  theadRow.appendChild(createElement('th', '', 'User'));
  theadRow.appendChild(createElement('th', '', 'Subscription ID'))

  if (subscriptions) {
    // List subscriptions
    for (const subscription of subscriptions) {
      const row = createElement('tr', '');
      subscriptionTable.appendChild(row);

      const deleteButtonCell = createElement('td', '');
      row.appendChild(deleteButtonCell);

      const deleteButton = createElement('button', 'btn btn-sm btn-primary', 'Delete');
      deleteButton.setAttribute('onclick', `deleteSubscription("${subscription.subscriptionId}");`);
      deleteButtonCell.appendChild(deleteButton);

      row.appendChild(createElement('td', '', subscription.userId));
      row.appendChild(createElement('td', '', subscription.subscriptionId));
    }
  }

  mainContainer.innerHTML = '';
  mainContainer.appendChild(form);
  mainContainer.appendChild(card);
}

// Renders an error
function showError(error) {
  const alert = createElement('div', 'alert alert-danger');

  const message = createElement('p', 'mb-3', error.message);
  alert.appendChild(message);

  if (error.debug)
  {
    const pre = createElement('pre', 'alert-pre border bg-light p-2');
    alert.appendChild(pre);

    const code = createElement('code', 'text-break text-wrap',
      JSON.stringify(error.debug, null, 2));
    pre.appendChild(code);
  }

  mainContainer.innerHTML = '';
  mainContainer.appendChild(alert);
}

// Re-renders the page with the selected view
function updatePage(view, data) {
  if (!view) {
    view = Views.home;
  }

  // Get the user name from the session
  const user = sessionStorage.getItem('msal-userName');
  if (!user && view !== Views.error)
  {
    view = Views.home;
  }

  showAccountNav(user);
  showAuthenticatedNav(user, view);

  switch (view) {
    case Views.error:
      showError(data);
      break;
    case Views.home:
      showWelcomeMessage(user);
      break;
    case Views.message:
      showLatestMessage(data);
      break;
    case Views.subscriptions:
      showSubscriptions();
      break;
  }
}

updatePage(Views.home);
// </uiJsSnippet>
