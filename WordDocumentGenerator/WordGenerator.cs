using Aspose.Words;
using MiniSoftware;
using Aspose;
using GemBox.Document;
using System.Collections.Generic;
using System.IO;

namespace WordDocumentGenerator
{
    /// <summary>
    /// Provides functionality to generate Word documents from templates and data.
    /// </summary>
    public class WordGenerator
    {
        private List<string> tempFilePaths = new List<string>();

        /// <summary>
        /// Generates a Word document from a template and data.
        /// </summary>
        /// <param name="template">The Word document template as a byte array.</param>
        /// <param name="data">The data to fill into the template.</param>
        /// <returns>A byte array representing the filled Word document.</returns>
        public byte[] GenerateWordDocument(byte[] template, Dictionary<string, object> data)
        {
            using (var outputStream = new MemoryStream())
            {
                // Process the data to convert byte arrays to MiniWordPicture objects
                var processedData = ProcessImageBytes(data);
                MiniWord.SaveAsByTemplate(outputStream, template, processedData);
                CleanupTempFiles(processedData);
                return outputStream.ToArray();
            }
        }

        /// <summary>
        /// Converts byte arrays in the data dictionary to MiniWordPicture objects.
        /// </summary>
        /// <param name="data">The data dictionary containing image byte arrays.</param>
        /// <returns>A new dictionary with MiniWordPicture objects for images.</returns>
        private Dictionary<string, object> ProcessImageBytes(Dictionary<string, object> data)
        {
            var processedData = new Dictionary<string, object>();
            foreach (var key in data.Keys)
            {
                if (data[key] is byte[] byteArray)
                {
                    string tempFilePath = SaveToTempFile(byteArray);
                    processedData[key] = new MiniWordPicture
                    {
                        Path = tempFilePath,
                        Width = 100,  // Set default width
                        Height = 100  // Set default height
                    };
                }
                else if (data[key] is List<Dictionary<string, object>> list)
                {
                    var processedList = new List<Dictionary<string, object>>();
                    foreach (var item in list)
                    {
                        processedList.Add(ProcessImageBytes(item));
                    }
                    processedData[key] = processedList;
                }
                else
                {
                    processedData[key] = data[key];
                }
            }
            return processedData;
        }

        /// <summary>
        /// Saves a byte array to a temporary file and returns the file path.
        /// </summary>
        /// <param name="bytes">The byte array to save.</param>
        /// <returns>The path to the temporary file.</returns>
        private string SaveToTempFile(byte[] bytes)
        {
            // Get the system's temporary directory
            string tempDir = Path.GetTempPath();

            // Generate a random file name with a .png extension
            string tempFilePath = Path.Combine(tempDir, Path.GetRandomFileName() + ".png");

            // Write the byte array to the file
            File.WriteAllBytes(tempFilePath, bytes);

            return tempFilePath;
        }


        /// <summary>
        /// Deletes the temporary files used in the processed data.
        /// </summary>
        /// <param name="data">The processed data dictionary containing MiniWordPicture objects.</param>
        private void CleanupTempFiles(Dictionary<string, object> data)
        {
            foreach (var key in data.Keys)
            {
                if (data[key] is MiniWordPicture picture)
                {
                    if (File.Exists(picture.Path))
                    {
                        File.Delete(picture.Path);
                    }
                }
                else if (data[key] is List<Dictionary<string, object>> list)
                {
                    foreach (var item in list)
                    {
                        CleanupTempFiles(item);
                    }
                }
            }
        }

        /// <summary>
        /// Converts a Word document byte array to a PDF byte array.
        /// </summary>
        /// <param name="wordBytes">The Word document byte array.</param>
        /// <returns>A byte array representing the PDF document.</returns>
        public byte[] ConvertToPdf(byte[] wordBytes)
        {
            using (var inputStream = new MemoryStream(wordBytes))
            using (var outputStream = new MemoryStream())
            {
                var document = new Document(inputStream);
                document.Save(outputStream, SaveFormat.Pdf);
                return outputStream.ToArray();
            }
        }
    }
}
