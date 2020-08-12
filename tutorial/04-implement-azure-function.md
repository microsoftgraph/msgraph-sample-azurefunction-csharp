<!-- markdownlint-disable MD002 MD041 -->

In this exercise you will finish implementing the Azure Function `GetMyNewestMessage` and update the test client to call the function.

The Azure Function uses the [on-behalf-of flow](https://docs.microsoft.com/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow). The basic order of events in this flow are:

- The test application uses an interactive auth flow to allow the user to sign in and grant consent. It gets back a token that is scoped to the Azure Function. The token does **NOT** contain any Microsoft Graph scopes.
- The test application invokes the Azure Function, sending its access token in the `Authorization` header.
- The Azure Function validates the token, then exchanges that token for a second access token that contains Microsoft Graph scopes.
- The Azure Function calls Microsoft Graph on the user's behalf using the second access token.

> [!IMPORTANT]
> To avoid storing the application ID and secret in source, you will use the [.NET Secret Manager](https://docs.microsoft.com/aspnet/core/security/app-secrets) to store these values. The Secret Manager is for development purposes only, production apps should use a trusted secret manager for storing secrets.

## Add authentication to the single page application

Start by adding authentication to the SPA. This will allow the application to get an access token granting access to call the Azure Function. Because this is a SPA, it will use the [authorization code flow with PKCE](https://docs.microsoft.com/azure/active-directory/develop/v2-oauth2-auth-code-flow).

1. Create a new file in the **TestClient** directory named **config.js** and add the following code.

    :::code language="javascript" source="../demo/graph-tutorial/config.example.js" id="msalConfigSnippet":::

    Replace `YOUR_TEST_APP_APP_ID_HERE` with the application ID you created in the Azure portal for the **Graph Azure Function Test App**. Replace `YOUR_TENANT_ID_HERE` with the **Directory (tenant) ID** value you copied from the Azure portal. Replace `YOUR_AZURE_FUNCTION_APP_ID_HERE` with the application ID for the **Graph Azure Function**.

    > [!IMPORTANT]
    > If you're using source control such as git, now would be a good time to exclude the **config.js** file from source control to avoid inadvertently leaking your app IDs and tenant ID.

1. Create a new file in the **TestClient** directory named **auth.js** and add the following code.

    :::code language="javascript" source="../demo/TestClient/auth.js" id="signInSignOutSnippet":::

    Consider what this code does.

    - It initializes a `PublicClientApplication` using the values stored in **config.js**.
    - It uses `loginPopup` to sign the user in, using the permission scope for the Azure Function.
    - It stores the user's username in the session.

    > [!IMPORTANT]
    > Since the app uses `loginPopup`, you may need to change your browser's pop-up blocker to allow pop-ups from `http://localhost:8080`.

1. Refresh the page and sign in. The page should update with the user name, indicating that the sign in was successful.

> [!TIP]
> You can parse the access token at [https://jwt.ms](https://jwt.ms) and confirm that the `aud` claim is the app ID for the Azure Function, and that the `scp` claim contains the Azure Function's permission scope, not Microsoft Graph.

## Add authentication to the Azure Function

In this section you'll implement the on-behalf-of flow in the `GetMyNewestMessage` Azure Function to get an access token compatible with Microsoft Graph.

1. Initialize the .NET development secret store by opening your CLI in the directory that contains **GraphTutorial.csproj** and running the following command.

    ```Shell
    dotnet user-secrets init
    ```

1. Add your application ID, secret, and tenant ID to the secret store using the following commands. Replace `YOUR_API_FUNCTION_APP_ID_HERE` with the application ID for the **Graph Azure Function**. Replace `YOUR_API_FUNCTION_APP_SECRET_HERE` with the application secret you created in the Azure portal for the **Graph Azure Function**. Replace `YOUR_TENANT_ID_HERE` with the **Directory (tenant) ID** value you copied from the Azure portal.

    ```Shell
    dotnet user-secrets set apiFunctionId "YOUR_API_FUNCTION_APP_ID_HERE"
    dotnet user-secrets set apiFunctionSecret "YOUR_API_FUNCTION_APP_SECRET_HERE"
    dotnet user-secrets set tenantId "YOUR_TENANT_ID_HERE"
    ```

### Process the incoming bearer token

In this section you'll implement a class to validate and process the bearer token sent from the SPA to the Azure Function.

1. Create a new directory in the **GraphTutorial** directory named **Authentication**.

1. Create a new file named **TokenValidationResult.cs** in the **./GraphTutorial/Authentication** folder, and add the following code.

    :::code language="csharp" source="../demo/GraphTutorial/Authentication/TokenValidationResult.cs" id="TokenValidationResultSnippet":::

1. Create a new file named **TokenValidation.cs** in the **./GraphTutorial/Authentication** folder, and add the following code.

    :::code language="csharp" source="../demo/GraphTutorial/Authentication/TokenValidation.cs" id="TokenValidationSnippet":::

Consider what this code does.

- It ensure there is a bearer token in the `Authorization` header.
- It verifies the signature and issuer from Azure's published OpenID configuration.
- It verifies that the audience (`aud` claim) matches the Azure Function's application ID.
- It parses the token and generates an MSAL account ID, which will be needed to take advantage of token caching.

### Create an on-behalf-of authentication provider

1. Create a new file in the **Authentication** directory named **OnBehalfOfAuthProvider.cs** and add the following code to that file.

    :::code language="csharp" source="../demo/GraphTutorial/Authentication/OnBehalfOfAuthProvider.cs" id="AuthProviderSnippet":::

Take a moment to consider what the code in **OnBehalfOfAuthProvider.cs** does.

- In the `GetAccessToken` function, it first attempts to get a user token from the token cache using `AcquireTokenSilent`. If this fails, it uses the bearer token sent by the test app to the Azure Function to generate a user assertion. It then uses that user assertion to get a Graph-compatible token using `AcquireTokenOnBehalfOf`.
- It implements the `Microsoft.Graph.IAuthenticationProvider` interface, allowing this class to be passed in the constructor of the `GraphServiceClient` to authenticate outgoing requests.

### Implement a Graph client service

In this section you'll implement a service that can be registered for [dependency injection](https://docs.microsoft.com/azure/azure-functions/functions-dotnet-dependency-injection). The service will be used to get an authenticated Graph client.

1. Create a new directory in the **GraphTutorial** directory named **Services**.

1. Create a new file in the **Services** directory named **IGraphClientService.cs** and add the following code to that file.

    :::code language="csharp" source="../demo/GraphTutorial/Services/IGraphClientService.cs" id="IGraphClientServiceSnippet":::

1. Create a new file in the **Services** directory named **GraphClientService.cs** and add the following code to that file.

    ```csharp
    using GraphTutorial.Authentication;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Extensions.Logging;
    using Microsoft.Identity.Client;
    using Microsoft.Graph;

    namespace GraphTutorial.Services
    {
        // Service added via dependency injection
        // Used to get an authenticated Graph client
        public class GraphClientService : IGraphClientService
        {
        }
    }
    ```

1. Add the following properties to the `GraphClientService` class.

    :::code language="csharp" source="../demo/GraphTutorial/Services/GraphClientService.cs" id="UserGraphClientMembers":::

1. Add the following functions to the `GraphClientService` class.

    :::code language="csharp" source="../demo/GraphTutorial/Services/GraphClientService.cs" id="UseGraphClientFunctions":::

1. Add a placeholder implementation for the `GetAppGraphClient` function. You will implement that in later sections.

    ```csharp
    public GraphServiceClient GetAppGraphClient()
    {
        throw new System.NotImplementedException();
    }
    ```

    The `GetUserGraphClient` function takes the results of token validation and builds an authenticated `GraphServiceClient` for the user.

1. Create a new file in the **GraphTutorial** directory named **Startup.cs** and add the following code to that file.

    :::code language="csharp" source="../demo/GraphTutorial/Startup.cs" id="StartupSnippet":::

    This code will enable [dependency injection](https://docs.microsoft.com/azure/azure-functions/functions-dotnet-dependency-injection) in your Azure Functions, exposing the `IConfiguration` object and the `GraphClientService` service.

### Implement GetMyNewestMessage function

1. Open **./GraphTutorial/GetMyNewestMessage.cs** and replace its entire contents with the following.

    :::code language="csharp" source="../demo/GraphTutorial/GetMyNewestMessage.cs" id="GetMyNewestMessageSnippet":::

#### Review the code in GetMyNewestMessage.cs

Take a moment to consider what the code in **GetMyNewestMessage.cs** does.

- In the constructor, it saves the `IConfiguration` object passed in via dependency injection.
- In the `Run` function, it does the following:
  - Validates the required configuration values are present in the `IConfiguration` object.
  - Validates the bearer token and returns a `401` status code if the token is invalid.
  - Gets a Graph client from the `GraphClientService` for the user that made this request.
  - Uses the Microsoft Graph SDK to get the newest message from the user's inbox and returns it as a JSON body in the response.

## Call the Azure Function from the test app

1. Open **auth.js** and add the following function to get an access token.

    :::code language="javascript" source="../demo/TestClient/auth.js" id="getTokenSnippet":::

    Consider what this code does.

    - It first attempts to get an access token silently, without user interaction. Since the user should already be signed in, MSAL should have tokens for the user in its cache.
    - If that fails with an error that indicates the user needs to interact, it attempts to get a token interactively.

1. Create a new file in the **TestClient** directory named **azurefunctions.js** and add the following code.

    :::code language="javascript" source="../demo/TestClient/azurefunctions.js" id="getLatestMessageSnippet":::

1. Change the current directory in your CLI to the **./GraphTutorial** directory and run the following command to start the Azure Function locally.

    ```Shell
    func start
    ```

1. If not already serving the SPA, open a second CLI window and change the current directory to the **./TestClient** directory. Run the following command to run the test application.

    ```Shell
    dotnet serve -h "Cache-Control: no-cache, no-store, must-revalidate"
    ```

1. Open your browser and navigate to `http://localhost:8080`. Sign in and select the **Latest Message** navigation item. The app displays information about the newest message in the user's inbox.
