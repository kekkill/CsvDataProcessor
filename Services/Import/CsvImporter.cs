using CsvDataProcessor.Data;
using CsvDataProcessor.Domain;
using Microsoft.VisualBasic.FileIO;
using System.Text;

namespace CsvDataProcessor.Services.Import
{
    public class CsvImporter
    {
        public static async Task<(int successCount, int errorCount)> ImportCsvToDatabaseAsync(string filePath, AppDbContext context, IProgress<string> progress)
        {
            int successCount = 0;
            int errorCount = 0;
            int processedLines = 0;
            var peopleBatch = new List<Person>(1000);

            using var reader = new StreamReader(filePath, Encoding.UTF8);
            using var parser = new TextFieldParser(reader)
            {
                Delimiters = new[] { ";" },
                HasFieldsEnclosedInQuotes = true,
                TrimWhiteSpace = true
            };

            bool hasHeader = CheckForHeader(parser);
            if (hasHeader) processedLines++;

            while (!parser.EndOfData)
            {
                try
                {
                    string[] fields = parser.ReadFields();
                    processedLines++;

                    if (fields == null || fields.Length < 6)
                    {
                        errorCount++;
                        continue;
                    }

                    var person = ParsePerson(fields);
                    if (person == null)
                    {
                        errorCount++;
                        continue;
                    }

                    peopleBatch.Add(person);

                    if (peopleBatch.Count >= 1000)
                    {
                        var (batchSuccess, batchError) = await SaveBatchAsync(context, peopleBatch);
                        successCount += batchSuccess;
                        errorCount += batchError;
                        peopleBatch.Clear();

                        progress?.Report($"Обработано: {processedLines:N0} | Успешно: {successCount:N0} | Ошибок: {errorCount:N0}");
                    }
                }
                catch
                {
                    errorCount++;
                }
            }

            if (peopleBatch.Count > 0)
            {
                var (batchSuccess, batchError) = await SaveBatchAsync(context, peopleBatch);
                successCount += batchSuccess;
                errorCount += batchError;
            }

            progress?.Report($"Завершено: {processedLines:N0} строк | Успешно: {successCount:N0} | Ошибок: {errorCount:N0}");
            return (successCount, errorCount);
        }

        private static bool CheckForHeader(TextFieldParser parser)
        {
            if (parser.EndOfData) return false;

            string[] firstLine = parser.ReadFields();
            if (firstLine == null || firstLine.Length == 0) return false;

            string firstField = firstLine[0].Trim().ToLowerInvariant();
            return firstField.Contains("date") ||
                   firstField.Contains("дата") ||
                   firstField.Contains("surname") ||
                   firstField.Contains("страна");
        }

        private static Person ParsePerson(string[] fields)
        {
            if (!DateTime.TryParse(fields[0].Trim(), out DateTime parsedDate))
                return null;

            var person = new Person
            {
                Date = parsedDate.Date,
                FirstName = fields[1].Trim(),
                LastName = fields[2].Trim(),
                SurName = fields.Length > 3 ? fields[3].Trim() : "",
                City = fields.Length > 4 ? fields[4].Trim() : "",
                Country = fields.Length > 5 ? fields[5].Trim() : ""
            };

            return string.IsNullOrWhiteSpace(person.FirstName) ||
                   string.IsNullOrWhiteSpace(person.LastName)
                   ? null
                   : person;
        }

        private static async Task<(int successCount, int errorCount)> SaveBatchAsync(AppDbContext context, List<Person> batch)
        {
            try
            {
                await context.People.AddRangeAsync(batch);
                await context.SaveChangesAsync();
                return (batch.Count, 0);
            }
            catch
            {
                return (0, batch.Count);
            }
            finally
            {
                context.ChangeTracker.Clear();
            }
        }
    }
}