using System;
using System.Threading.Tasks;

namespace GoldRatesExtractor
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static async Task Main(string[] args)
        {
            Console.WriteLine("Gold Rates Extractor Starting...");

            try
            {
                // Create and run the extractor
                var extractor = new GoldRatesExtractor();
                await extractor.StartAsync();
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Fatal error: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
                Console.ResetColor();
            }

            // To keep console window open if running manually
            if (IsRunningInteractively())
            {
                Console.WriteLine("\nPress any key to exit...");
                Console.ReadKey();
            }
        }

        // Check if the program is running interactively (with a console) vs. via task scheduler
        private static bool IsRunningInteractively()
        {
            try
            {
                // This will throw an exception if there's no console
                return Console.WindowHeight > 0;
            }
            catch
            {
                return false;
            }
        }
    }
}