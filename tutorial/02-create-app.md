<!-- markdownlint-disable MD002 MD041 -->

In this tutorial, you will create a simple Azure Function that implements HTTP trigger functions that call Microsoft Graph. These functions will cover the following scenarios:

- Implements an API to access a user's inbox using [on-behalf-of flow](https://docs.microsoft.com/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow) authentication.
- Implements an API to subscribe and unsubscribe for notifications on a user's inbox, using using [client credentials grant flow](https://docs.microsoft.com/azure/active-directory/develop/v2-oauth2-client-creds-grant-flow) authentication.
- Implements a webhook to receive [change notifications](https://docs.microsoft.com/graph/webhooks) from Microsoft Graph and access data using client credentials grant flow.

You will also create a simple JavaScript single-page application (SPA) to call the APIs implemented in the Azure Function.

## Create Azure Functions project

1. Open your command-line interface (CLI) in a directory where you want to create the project. Run the following command.

    ```Shell
    func init GraphTutorial --dotnet
    ```

1. Change the current directory in your CLI to the **GraphTutorial** directory and run the following commands to create three functions in the project.

    ```Shell
    func new --name GetMyNewestMessage --template "HTTP trigger" --language C#
    func new --name SetSubscription --template "HTTP trigger" --language C#
    func new --name Notify --template "HTTP trigger" --language C#
    ```

1. Run the following command to run the project locally.

    ```Shell
    func start
    ```

1. If everything is working, you will see the following output:

    ```Shell
    Http Functions:

        GetMyNewestMessage: [GET,POST] http://localhost:7071/api/GetMyNewestMessage

        Notify: [GET,POST] http://localhost:7071/api/Notify

        SetSubscription: [GET,POST] http://localhost:7071/api/SetSubscription
    ```

1. Verify that the functions are working correctly by opening your browser and browsing to the function URLs shown in the output. You should see the following message in your browser: `This HTTP triggered function executed successfully. Pass a name in the query string or in the request body for a personalized response.`.

## Create single-page application

1. Open your CLI in a directory where you want to create the project. Create a directory named **TestClient** to hold your HTML and JavaScript files.

1. Create a new file named **index.html** in the **TestClient** directory and add the following code.

    :::code language="html" source="../demo/TestClient/index.html" id="indexSnippet":::

    This defines the basic layout of the app, including a navigation bar. It also adds the following:

    - [Bootstrap](https://getbootstrap.com/) and its supporting JavaScript
    - [FontAwesome](https://fontawesome.com/)
    - [Microsoft Authentication Library for JavaScript (MSAL.js) 2.0](https://github.com/AzureAD/microsoft-authentication-library-for-js/tree/dev/lib/msal-browser)

    > [!TIP]
    > The page includes a favicon, (`<link rel="shortcut icon" href="g-raph.png">`). You can remove this line, or you can download the **g-raph.png** file from [GitHub](https://github.com/microsoftgraph/g-raph).

1. Create a new file named **style.css** in the **TestClient** directory and add the following code.

    :::code language="css" source="../demo/TestClient/style.css":::

1. Create a new file named **ui.js** in the **TestClient** directory and add the following code.

    :::code language="javascript" source="../demo/TestClient/ui.js" id="uiJsSnippet":::

    This code uses JavaScript to render the current page based on the selected view.

### Test the single-page application

> [!NOTE]
> This section includes instructions for using [dotnet-serve](https://github.com/natemcmaster/dotnet-serve) to run a simple testing HTTP server on your development machine. Using this specific tool is not required. You can use any testing server you prefer to serve the **TestClient** directory.

1. Run the following command in your CLI to install **dotnet-serve**.

    ```Shell
    dotnet tool install --global dotnet-serve
    ```

1. Change the current directory in your CLI to the **TestClient** directory and run the following command to start an HTTP server.

    ```Shell
    dotnet serve -h "Cache-Control: no-cache, no-store, must-revalidate"
    ```

1. Open your browser and navigate to `http://localhost:8080`. The page should render, but none of the buttons currently work.

## Add NuGet packages

Before moving on, install some additional NuGet packages that you will use later.

- [Microsoft.Azure.Functions.Extensions](https://www.nuget.org/packages/Microsoft.Azure.Functions.Extensions) to enable dependency injection in the Azure Functions project.
- [Microsoft.Extensions.Configuration.UserSecrets](https://www.nuget.org/packages/Microsoft.Extensions.Configuration.UserSecrets) to read application configuration from the [.NET development secret store](https://docs.microsoft.com/aspnet/core/security/app-secrets).
- [Microsoft.Graph](https://www.nuget.org/packages/Microsoft.Graph/) for making calls to Microsoft Graph.
- [Microsoft.Identity.Client](https://www.nuget.org/packages/Microsoft.Identity.Client/) for authenticating and managing tokens.
- [Microsoft.IdentityModel.Protocols.OpenIdConnect](https://www.nuget.org/packages/Microsoft.IdentityModel.Protocols.OpenIdConnect) for retrieving OpenID configuration for token validation.
- [System.IdentityModel.Tokens.Jwt](https://www.nuget.org/packages/System.IdentityModel.Tokens.Jwt) for validating tokens sent to the web API.

1. Change the current directory in your CLI to the **GraphTutorial** directory and run the following commands.

    ```Shell
    dotnet add package Microsoft.Azure.Functions.Extensions --version 1.0.0
    dotnet add package Microsoft.Extensions.Configuration.UserSecrets --version 3.1.5
    dotnet add package Microsoft.Graph --version 3.8.0
    dotnet add package Microsoft.Identity.Client --version 4.15.0
    dotnet add package Microsoft.IdentityModel.Protocols.OpenIdConnect --version 6.7.1
    dotnet add package System.IdentityModel.Tokens.Jwt --version 6.7.1
    ```
