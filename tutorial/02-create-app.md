<!-- markdownlint-disable MD002 MD041 -->

In this tutorial, you will create a simple Azure Function that implements HTTP trigger functions that call Microsoft Graph. These functions will cover the following scenarios:

- Implements a web API to access a user's inbox using [on-behalf-of flow](https://docs.microsoft.com/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow) authentication.
- Implements a web API to subscribe and unsubscribe for notifications on a user's inbox, using using [client credentials grant flow](https://docs.microsoft.com/azure/active-directory/develop/v2-oauth2-client-creds-grant-flow) authentication.
- Implements a webhook to receive [change notifications](https://docs.microsoft.com/graph/webhooks) from Microsoft Graph and access data using client credentials grant flow.

You will also create a simple command line application to call the web APIs implemented in the Azure Function.

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

## Create command line application

1. Open your CLI in a directory where you want to create the project. Run the following command.

    ```Shell
    dotnet new console -o InvokeAzureFunction
    ```

1. Once the project is created, verify that it works by changing the current directory to the **InvokeAzureFunction** directory and running the following command in your CLI.

    ```Shell
    dotnet run
    ```

    If it works, the app outputs `Hello World!`.

## Add NuGet packages

Before moving on, install some additional NuGet packages that you will use later.

- [Microsoft.Extensions.Configuration.UserSecrets](https://github.com/aspnet/extensions) to read application configuration from the [.NET development secret store](https://docs.microsoft.com/aspnet/core/security/app-secrets).
- [Microsoft.Identity.Client](https://www.nuget.org/packages/Microsoft.Identity.Client/) for authenticating and managing tokens.
- [Microsoft.Graph](https://www.nuget.org/packages/Microsoft.Graph/) for making calls to Microsoft Graph.

1. Change the current directory in your CLI to the **GraphTutorial** directory and run the following commands.

    ```Shell
    dotnet add package Microsoft.Identity.Client --version 4.15.0
    dotnet add package Microsoft.Graph --version 3.8.0
    ```

1. Change the current directory in your CLI to the **InvokeAzureFunction** directory and run the following command.

    ```Shell
    dotnet add package Microsoft.Extensions.Configuration.UserSecrets --version 3.1.5
    dotnet add package Microsoft.Identity.Client --version 4.15.0
    ```

## Design the app

In this section you will create a simple console-based menu.

1. Open **.\InvokeAzureFunction\Program.cs** in a text editor (such as [Visual Studio Code](https://code.visualstudio.com/)) and replace its entire contents with the following code.

    ```csharp
    using Microsoft.Extensions.Configuration;
    using System;
    using System.Threading.Tasks;

    namespace InvokeAzureFunction
    {
        class Program
        {
            static async Task Main(string[] args)
            {
                Console.WriteLine("Azure Function Graph Tutorial\n");

                int choice = -1;

                while (choice != 0) {
                    Console.WriteLine("Please choose one of the following options:");
                    Console.WriteLine("0. Exit");
                    Console.WriteLine("1. Display the newest message in my inbox");
                    Console.WriteLine("2. Subscribe to notifications in a user's inbox");
                    Console.WriteLine("3. Unsubscribe to notifications in a user's inbox");

                    try
                    {
                        choice = int.Parse(Console.ReadLine());
                    }
                    catch (System.FormatException)
                    {
                        // Set to invalid value
                        choice = -1;
                    }

                    switch(choice)
                    {
                        case 0:
                            // Exit the program
                            Console.WriteLine("Goodbye...");
                            break;
                        case 1:
                            // Get signed-in user's newest email message
                            break;
                        case 2:
                            // Subscribe
                            break;
                        case 2:
                            // Unsubscribe
                            break;
                        default:
                            Console.WriteLine("Invalid choice! Please try again.");
                            break;
                    }
                }
            }
        }
    }
    ```

This implements a basic menu and reads the user's choice from the command line.
