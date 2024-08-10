using Xunit;
using WordDocumentGenerator;
using System;
using System.Collections.Generic;
using System.IO;
using QrCodeGeneratorLibrary;

namespace WordDocumentGenerator.Tests
{
    public class WordGeneratorTests
    {

        [Fact]
        public void GenerateWordDocument_WithMixedData_ReturnsNonEmptyByteArray()
        {
            // Arrange
            var generator = new WordGenerator();
            string templatePath = Path.Combine("TestTemplates", "template.docx");
            string fullPath = Path.GetFullPath(templatePath);

            byte[] template = File.ReadAllBytes(fullPath);

            // Generate QR code as byte array
            var qrCodeGenerator = new QrCodeGenerator();
            byte[] qrCodeBytes = qrCodeGenerator.GenerateQrCode("Sample QR Code Data");

            var data = new Dictionary<string, object>
            {
                { "Title", "Certificate of Completion" },
                { "Recipient", "John Doe" },
                { "Date", DateTime.Now.ToString("yyyy-MM-dd") },
                { "Logo", qrCodeBytes },
                { "Employees", new List<Dictionary<string, object>>
                    {
                        new Dictionary<string, object> { { "ID", "001" }, { "Name", "Alice" }, { "Role", "Developer" } },
                        new Dictionary<string, object> { { "ID", "002" }, { "Name", "Bob" }, { "Role", "Designer" } }
                    }
                }
            };

            // Act
            byte[] wordResult = generator.GenerateWordDocument(template, data);
            byte[] pdfResult = generator.ConvertToPdf(wordResult);

            // Assert
            Assert.NotNull(wordResult);
            Assert.True(wordResult.Length > 0);
            Assert.NotNull(pdfResult);
            Assert.True(pdfResult.Length > 0);

            // Save the result for manual verification
            string wordOutputPath = @"C:\Users\basel\Desktop\2Bone1\Project 2\result.docx";
            File.WriteAllBytes(wordOutputPath, wordResult);
            string pdfOutputPath = @"C:\Users\basel\Desktop\2Bone1\Project 2\result.pdf";
            File.WriteAllBytes(pdfOutputPath, pdfResult);
            Console.WriteLine($"Generated Word document saved to: {wordOutputPath}");
            Console.WriteLine($"Generated PDF document saved to: {pdfOutputPath}");

        }
    }
}

