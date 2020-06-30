// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

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
