// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

using System.Reflection;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using GraphSampleFunctions.Services;

var host = new HostBuilder()
    .ConfigureFunctionsWorkerDefaults()
    .ConfigureAppConfiguration(config => {
        config.AddUserSecrets(Assembly.GetExecutingAssembly(), false);
    })
    .ConfigureServices(services => {
        services.AddSingleton<ITokenValidationService, TokenValidationService>();
        services.AddSingleton<IGraphClientService, GraphClientService>();
    })
    .Build();

host.Run();
