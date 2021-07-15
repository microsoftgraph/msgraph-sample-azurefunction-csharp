// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

// <ProgramSnippet>
using System.Reflection;
using GraphTutorial.Services;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;

namespace GraphTutorial
{
    public class Program
    {
        public static void Main()
        {
            var host = new HostBuilder()
                .ConfigureFunctionsWorkerDefaults()
                .ConfigureAppConfiguration(configuration =>
                {
                    configuration.AddUserSecrets(
                        Assembly.GetExecutingAssembly(), false);
                })
                .ConfigureServices(services =>
                {
                    services.AddSingleton<IGraphClientService, GraphClientService>();
                })
                .Build();

            host.Run();
        }
    }
}
// </ProgramSnippet>
