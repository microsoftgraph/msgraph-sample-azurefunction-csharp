// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

using System.Reflection;
using System.Text.Json;
using Azure.Core.Serialization;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using GraphSampleFunctions.Services;

var host = new HostBuilder()
    .ConfigureFunctionsWorkerDefaults(configureOptions: options =>
    {
        options.Serializer = new JsonObjectSerializer(
            new JsonSerializerOptions
            {
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            }
        );
    })
    .ConfigureAppConfiguration(config => {
        config.AddUserSecrets(Assembly.GetExecutingAssembly(), false);
    })
    .ConfigureServices(services => {
        services.AddSingleton<ITokenValidationService, TokenValidationService>();
        services.AddSingleton<IGraphClientService, GraphClientService>();
    })
    .Build();

host.Run();
