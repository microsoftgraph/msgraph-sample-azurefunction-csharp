<!-- markdownlint-disable MD002 MD041 -->

This tutorial teaches you how to build an Azure Function that uses the Microsoft Graph API to retrieve calendar information for a user.

> [!TIP]
> If you prefer to just download the completed tutorial, you can download or clone the [GitHub repository](https://github.com/microsoftgraph/msgraph-training-azurefunction-csharp). See the README file in the **demo** folder for instructions on configuring the app with an app ID and secret.

## Prerequisites

Before you start this tutorial, you should have the following tools installed on your development machine.

- [.NET Core SDK](https://dotnet.microsoft.com/download)
- [Node.js](https://nodejs.org/)
- [Azure Functions Core Tools](https://www.npmjs.com/package/azure-functions-core-tools)
- [Azure CLI](/cli/azure/install-azure-cli)

You should also have either a personal Microsoft account with a mailbox on Outlook.com, or a Microsoft work or school account. If you don't have a Microsoft account, there are a couple of options to get a free account:

- You can [sign up for a new personal Microsoft account](https://signup.live.com/signup?wa=wsignin1.0&rpsnv=12&ct=1454618383&rver=6.4.6456.0&wp=MBI_SSL_SHARED&wreply=https://mail.live.com/default.aspx&id=64855&cbcxt=mai&bk=1454618383&uiflavor=web&uaid=b213a65b4fdc484382b6622b3ecaa547&mkt=E-US&lc=1033&lic=1).
- You can [sign up for the Office 365 Developer Program](https://developer.microsoft.com/office/dev-program) to get a free Office 365 subscription.

> [!NOTE]
> This tutorial was written with the following versions of the above tools. The steps in this guide may work with other versions, but that has not been tested.
>
> - .NET Core SDK 3.1.301
> - Node.js 12.16.1
> - Azure Functions Core Tools 2.7.2628
> - Azure CLI 2.8.0

## Feedback

Please provide any feedback on this tutorial in the [GitHub repository](https://github.com/microsoftgraph/msgraph-training-azurefunction-csharp).
