<!-- markdownlint-disable MD002 MD041 -->

In this exercise you'll learn about what changes are needed to the sample Azure Function to prepare for [deployment to an Azure Functions app](https://docs.microsoft.com/azure/azure-functions/functions-run-local#publish).

## Update code

Configuration is read from the user secret store, which only applies to your development machine. Before you deploy to Azure, you'll need to change where you store your configuration, and update the code in **Startup.cs** accordingly.

Application secrets should be stored in secure storage, such as [Azure Key Vault](https://docs.microsoft.com/azure/key-vault/general/overview).

## Update CORS setting for Azure Function

In this sample we configured CORS in **local.settings.json** to allow the test application to call the function. You'll need to configure your deployed function to allow any SPA apps that will call it.

## Update app registrations

The  `knownClientApplications` property in the manifest for the **Graph Azure Function** app registration will need to be updated with the application IDs of any apps that will be calling the Azure Function.

## Recreate existing subscriptions

Any subscriptions created using the webhook URL on your local machine or ngrok should be recreated using the production URL of the `Notify` Azure Function.
