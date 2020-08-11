// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

using InvokeAzureFunction.Authentication;
using Microsoft.Extensions.Configuration;
using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.Json;
using System.Threading.Tasks;

namespace InvokeAzureFunction
{
    class Program
    {
        static async Task Main(string[] args)
        {
            Console.WriteLine("Azure Function Graph Tutorial\n");

            // Prompt for ngrok URL
            string ngrokProxy = "";
            while (string.IsNullOrEmpty(ngrokProxy))
            {
                Console.Write("Enter https ngrok URL: ");
                ngrokProxy = Console.ReadLine();

                if (!Uri.IsWellFormedUriString(ngrokProxy, UriKind.Absolute))
                {
                    Console.WriteLine("Invalid input, please enter URL in form https://418ead6a47a6.ngrok.io");
                    ngrokProxy = "";
                }
            }

            // <InitializationSnippet>
            var appConfig = LoadAppSettings();

            if (appConfig == null)
            {
                Console.WriteLine("Missing or invalid values in secret store...exiting");
                return;
            }

            var appId = appConfig["appId"];
            var tenantId = appConfig["tenantId"];
            var scopes = new[] { $"{appConfig["apiFunctionId"]}/.default" };

            // Initialize the auth provider with values from appsettings.json
            var authProvider = new DeviceCodeAuthProvider(appId, tenantId, scopes);

            // Request a token to sign in the user
            var accessToken = await authProvider.GetAccessToken();

            Console.WriteLine("Authentication successful!");
            Console.WriteLine($"Token: {accessToken}\n");
            // </InitializationSnippet>

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

                try
                {
                    switch(choice)
                    {
                        case 0:
                            // Exit the program
                            Console.WriteLine("Goodbye...");
                            break;
                        case 1:
                            // Get signed-in user's newest email message
                            await GetNewestMessage(accessToken, ngrokProxy);
                            break;
                        case 2:
                            // Subscribe
                            await CreateSubscription(ngrokProxy);
                            break;
                        case 3:
                            // Unsubscribe
                            await DeleteSubscription(ngrokProxy);
                            break;
                        default:
                            Console.WriteLine("Invalid choice! Please try again.");
                            break;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"\nERROR: {ex.Message}\n");
                }
            }
        }

        // Pretty-print a JSON string using System.Text.Json
        private static string PrettyPrintJson(string uglyJson)
        {
            using var jsonDoc = JsonDocument.Parse(uglyJson);

            var stream = new MemoryStream();
            using (var jsonWriter = new Utf8JsonWriter(stream, new JsonWriterOptions{ Indented = true }))
            {
                jsonDoc.WriteTo(jsonWriter);
            }

            return new System.Text.UTF8Encoding().GetString(stream.ToArray());
        }

        // <GetNewestMessageSnippet>
        private static async Task GetNewestMessage(string token, string ngrokProxy)
        {
            // Do a GET to the Web API
            var request = new HttpRequestMessage(HttpMethod.Get,
                $"{ngrokProxy}/api/GetMyNewestMessage");
            // Add token in Authorization header
            request.Headers.Authorization = new AuthenticationHeaderValue("bearer", token);

            using (var client = new HttpClient())
            {
                var response = await client.SendAsync(request);

                // Read the response body and output as JSON
                var message = await response.Content.ReadAsStringAsync();
                Console.WriteLine($"\nMessage: {PrettyPrintJson(message)}\n");
            }
        }
        // </GetNewestMessageSnippet>

        // <CreateSubscriptionSnippet>
        private static async Task CreateSubscription(string ngrokProxy)
        {
            // Prompt user for user
            Console.Write("User to subscribe for: ");
            var userId = Console.ReadLine();

            var payload = $"{{\"requestType\": \"subscribe\" ,\"userId\": \"{userId}\", \"ngrokProxy\": \"{ngrokProxy}\"}}";

            var request = new HttpRequestMessage(HttpMethod.Post,
                $"{ngrokProxy}/api/SetSubscription");
            request.Content = new StringContent(payload);
            request.Content.Headers.ContentType.MediaType = "application/json";

            using (var client = new HttpClient())
            {
                var response = await client.SendAsync(request);

                var subscription = await response.Content.ReadAsStringAsync();
                Console.WriteLine($"\nSubscription created: {PrettyPrintJson(subscription)}\n");
            }
        }
        // </CreateSubscriptionSnippet>

        // <DeleteSubscriptionSnippet>
        private static async Task DeleteSubscription(string ngrokProxy)
        {
            // Prompt for subscription ID
            Console.Write("Subscription ID: ");
            var subscriptionId = Console.ReadLine();

            var payload = $"{{\"requestType\": \"unsubscribe\" ,\"subscriptionId\": \"{subscriptionId}\"}}";

            var request = new HttpRequestMessage(HttpMethod.Post,
                $"{ngrokProxy}/api/SetSubscription");
            request.Content = new StringContent(payload);
            request.Content.Headers.ContentType.MediaType = "application/json";

            using (var client = new HttpClient())
            {
                var response = await client.SendAsync(request);

                if (response.StatusCode == HttpStatusCode.Accepted)
                {
                    Console.WriteLine("Subscription deleted");
                }
                else
                {
                    Console.WriteLine($"Request returned {response.StatusCode}");
                    var errorMessage = await response.Content.ReadAsStringAsync();
                    if (!string.IsNullOrEmpty(errorMessage))
                    {
                        Console.WriteLine(errorMessage);
                    }
                }
            }
        }
        // </DeleteSubscriptionSnippet>

        // <LoadAppSettingsSnippet>
        private static IConfigurationRoot LoadAppSettings()
        {
            var appConfig = new ConfigurationBuilder()
                .AddUserSecrets<Program>()
                .Build();

            // Check for required settings
            if (string.IsNullOrEmpty(appConfig["appId"]) ||
                string.IsNullOrEmpty(appConfig["tenantId"]) ||
                string.IsNullOrEmpty(appConfig["apiFunctionId"]))
            {
                return null;
            }

            return appConfig;
        }
        // </LoadAppSettingsSnippet>
    }
}
