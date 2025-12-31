using ClosedXML.Excel;
using CsvDataProcessor.Data;
using CsvDataProcessor.Services.Database;
using Microsoft.EntityFrameworkCore;
using System.Text;
using System.Xml;

namespace CsvDataProcessor.Services.Export
{
    public class ExportService
    {
        public async Task ExportToExcelAsync(ExportParameters parameters, string filePath)
        {
            await Task.Run(async () =>
            {
                using var workbook = new XLWorkbook();
                var worksheet = workbook.Worksheets.Add("Данные");

                worksheet.Cell(1, 1).Value = "Дата";
                worksheet.Cell(1, 2).Value = "Имя";
                worksheet.Cell(1, 3).Value = "Фамилия";
                worksheet.Cell(1, 4).Value = "Отчество";
                worksheet.Cell(1, 5).Value = "Город";
                worksheet.Cell(1, 6).Value = "Страна";

                var headerRange = worksheet.Range(1, 1, 1, 6);
                headerRange.Style.Font.Bold = true;
                headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;
                headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                int row = 2;
                const int pageSize = 1000;
                int page = 0;

                await using var context = new AppDbContext();
                var personService = new PersonService(context);
                var baseQuery = personService.GetFilteredPeople(
                    parameters.UseStartDate, parameters.StartDate,
                    parameters.UseEndDate, parameters.EndDate,
                    parameters.FirstName, parameters.LastName,
                    parameters.SurName, parameters.City, parameters.Country
                );

                while (true)
                {
                    var batch = await baseQuery
                        .Skip(page * pageSize)
                        .Take(pageSize)
                        .ToListAsync();

                    if (batch.Count == 0) break;

                    foreach (var person in batch)
                    {
                        worksheet.Cell(row, 1).Value = person.Date;
                        worksheet.Cell(row, 1).Style.DateFormat.Format = "dd.MM.yyyy";
                        worksheet.Cell(row, 2).Value = person.FirstName;
                        worksheet.Cell(row, 3).Value = person.LastName;
                        worksheet.Cell(row, 4).Value = person.SurName;
                        worksheet.Cell(row, 5).Value = person.City;
                        worksheet.Cell(row, 6).Value = person.Country;
                        row++;
                    }
                    page++;
                }

                if (row == 2)
                    throw new Exception("Нет данных для экспорта по выбранным фильтрам");

                worksheet.Columns().AdjustToContents();
                if (row > 2) worksheet.Range(1, 1, row - 1, 6).SetAutoFilter();
                worksheet.SheetView.FreezeRows(1);

                workbook.SaveAs(filePath);
            });
        }

        public async Task ExportToXmlAsync(ExportParameters parameters, string filePath)
        {
            await Task.Run(async () =>
            {
                var settings = new XmlWriterSettings
                {
                    Indent = true,
                    Encoding = Encoding.UTF8,
                    IndentChars = "  "
                };

                using var writer = XmlWriter.Create(filePath, settings);
                writer.WriteStartDocument();
                writer.WriteStartElement("TestProgram");

                await using var context = new AppDbContext();
                var personService = new PersonService(context);
                var query = personService.GetFilteredPeople(
                    parameters.UseStartDate, parameters.StartDate,
                    parameters.UseEndDate, parameters.EndDate,
                    parameters.FirstName, parameters.LastName,
                    parameters.SurName, parameters.City, parameters.Country
                );

                int id = 1;
                var asyncEnumerator = query.AsAsyncEnumerable().GetAsyncEnumerator();
                try
                {
                    while (await asyncEnumerator.MoveNextAsync())
                    {
                        var person = asyncEnumerator.Current;
                        writer.WriteStartElement("Record");
                        writer.WriteAttributeString("id", id.ToString());
                        writer.WriteElementString("Date", person.Date.ToString("yyyy-MM-dd"));
                        writer.WriteElementString("FirstName", person.FirstName);
                        writer.WriteElementString("LastName", person.LastName);
                        writer.WriteElementString("SurName", person.SurName);
                        writer.WriteElementString("City", person.City);
                        writer.WriteElementString("Country", person.Country);
                        writer.WriteEndElement();
                        id++;
                    }
                }
                finally
                {
                    await asyncEnumerator.DisposeAsync();
                }

                if (id == 1)
                    throw new Exception("Нет данных для экспорта по выбранным фильтрам");

                writer.WriteEndElement();
                writer.WriteEndDocument();
            });
        }
    }
}