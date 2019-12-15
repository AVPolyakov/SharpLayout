using System;
using System.Reflection;
using DbUp;

namespace DatabaseMigration
{
    class Program
    {
        public static string ConnectionString => @"Data Source=(local)\SQL2014;Initial Catalog=SharpLayout;Integrated Security=True";
        
        static int Main()
        {
            var upgrader =
                DeployChanges.To
                    .SqlDatabase(ConnectionString)
                    .WithScriptsEmbeddedInAssembly(Assembly.GetExecutingAssembly())
                    .LogToConsole()
                    .Build();

            var result = upgrader.PerformUpgrade();

            if (!result.Successful)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(result.Error);
                Console.ResetColor();
#if DEBUG
                Console.ReadLine();
#endif                
                return -1;
            }

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Success!");
            Console.ResetColor();
            return 0;
        }
    }
}
