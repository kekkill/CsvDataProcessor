using CsvDataProcess.Data;
using CsvDataProcess.Domain;
using Microsoft.VisualBasic.FileIO;
using System.Text;

namespace CsvDataProcess.Services.Import
{
    public class CsvImporter
    {
        public static async Task<(int successCount, int errorCount)> ImportCsvToDatabaseAsync(string filePath, AppDbContext context, IProgress<int> progress)
        {
            int successCount = 0;
            int errorCount = 0;
            int processedLines = 0;
            int totalLinesEstimate = 0;
            var peopleBatch = new List<Person>();
            string[] firstLineFields = null;
            bool hasHeader = false;

            using (var reader = new StreamReader(filePath, Encoding.UTF8))
            {
                while (await reader.ReadLineAsync() != null)
                {
                    totalLinesEstimate++;
                }
            }

            using (var reader = new StreamReader(filePath, Encoding.UTF8))
            using (var parser = new TextFieldParser(reader))
            {
                parser.Delimiters = new string[] { ";" };
                parser.HasFieldsEnclosedInQuotes = true;
                parser.TrimWhiteSpace = true;

                if (!parser.EndOfData)
                {
                    firstLineFields = parser.ReadFields();
                    processedLines++;

                    if (firstLineFields != null && firstLineFields.Length > 0)
                    {
                        string headerCheck = firstLineFields[0].Trim().ToLower();

                        if (headerCheck.Contains("date") || headerCheck.Contains("дата") ||
                            headerCheck.Contains("firstname") || headerCheck.Contains("имя") ||
                            headerCheck.Contains("lastname") || headerCheck.Contains("фамилия"))
                        {
                            hasHeader = true;
                            firstLineFields = null; 
                        }
                    }
                }

                while (!parser.EndOfData || firstLineFields != null)
                {
                    try
                    {
                        string[] fields;

                        if (firstLineFields != null)
                        {
                            fields = firstLineFields;
                            firstLineFields = null;
                        }
                        else
                        {
                            fields = parser.ReadFields();
                            processedLines++;
                        }

                        if (fields == null || fields.Length < 6)
                        {
                            errorCount++;
                            continue;
                        }

                        DateTime parsedDate;
                        if (!DateTime.TryParse(fields[0].Trim(), out parsedDate))
                        {
                            errorCount++;
                            continue;
                        }

                        var person = new Person
                        {
                            Date = parsedDate.Date,
                            FirstName = fields[1].Trim(),
                            LastName = fields[2].Trim(),
                            SurName = fields.Length > 3 ? fields[3].Trim() : "",
                            City = fields.Length > 4 ? fields[4].Trim() : "",
                            Country = fields.Length > 5 ? fields[5].Trim() : ""
                        };

                        if (string.IsNullOrEmpty(person.FirstName) || string.IsNullOrEmpty(person.LastName))
                        {
                            errorCount++;
                            continue;
                        }

                        peopleBatch.Add(person);
                        successCount++;

                        if (peopleBatch.Count >= 1000)
                        {
                            await context.People.AddRangeAsync(peopleBatch);
                            await context.SaveChangesAsync();
                            peopleBatch.Clear();
                            context.ChangeTracker.Clear(); 
                        }

                        if (progress != null && totalLinesEstimate > 0)
                        {
                            int percent = (int)((double)processedLines / totalLinesEstimate * 100);
                            progress.Report(Math.Min(percent, 100));
                        }
                    }
                    catch
                    {
                        errorCount++;
                    }
                }
            }

            if (peopleBatch.Count > 0)
            {
                await context.People.AddRangeAsync(peopleBatch);
                await context.SaveChangesAsync();
            }

            return (successCount, errorCount);
        }
    }
}