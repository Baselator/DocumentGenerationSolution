using System;
using WordDocumentGenerator.Tests; // Namespace of your test project

class Program
{
    static void Main(string[] args)
    {
        try
        {
            var test = new WordGeneratorTests();
            test.GenerateWordDocument_WithMixedData_ReturnsNonEmptyByteArray();
            Console.WriteLine("Test executed successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Test failed with exception: {ex.Message}");
            Console.WriteLine($"Stack Trace: {ex.StackTrace}");
        }

        Console.WriteLine("Press any key to exit...");
        Console.ReadKey();
    }
}
