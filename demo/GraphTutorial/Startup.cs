// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

// <StartupSnippet>
using GraphTutorial.Services;
using Microsoft.Azure.Functions.Extensions.DependencyInjection;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using System.Reflection;

// Register the assembly
[assembly: FunctionsStartup(typeof(GraphTutorial.Startup))]
namespace GraphTutorial
{
    public class Startup : FunctionsStartup
    {
        // Override the Configure method to load configuration values
        // from the .NET user secret store
        public override void Configure(IFunctionsHostBuilder builder)
        {
            var config = new ConfigurationBuilder()
                .AddUserSecrets(Assembly.GetExecutingAssembly(), false)
                .Build();

            // Make the loaded config available via dependency injection
            builder.Services.AddSingleton<IConfiguration>(config);

            // Add the Graph client service
            builder.Services.AddSingleton<IGraphClientService, GraphClientService>();
        }
    }
}
// </StartupSnippet>
