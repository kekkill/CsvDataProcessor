using ClosedXML.Excel;
using CsvDataProcess.Domain;
using Microsoft.EntityFrameworkCore;

namespace CsvDataProcess.Services.Export
{
    public class ExportService
    {
        public async Task ExportToExcelAsync(IQueryable<Person> query, string filePath)
        {
            await Task.Run(() =>
            {
                var people = query.AsNoTracking().ToList();

                if (people.Count == 0)
                {
                    throw new Exception("Нет данных для экспорта по выбранным фильтрам");
                }

                if (people.Count > 1000000)
                {
                    throw new Exception($"Слишком много данных для экспорта в Excel (максимум 1 000 000 записей).\n" +
                                        $"Сейчас выбрано: {people.Count:N0}\n" +
                                        $"Пожалуйста, сузьте фильтры.");
                }

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
                foreach (var person in people)
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

                worksheet.Columns().AdjustToContents();

                if (row > 2) 
                {
                    worksheet.Range(1, 1, row - 1, 6).SetAutoFilter();
                }

                worksheet.SheetView.FreezeRows(1);

                workbook.SaveAs(filePath);
            });
        }

        public async Task ExportToXmlAsync(IQueryable<Person> query, string filePath)
        {
            await Task.Run(() =>
            {
                var people = query.AsNoTracking().ToList();

                if (people.Count == 0)
                {
                    throw new Exception("Нет данных для экспорта по выбранным фильтрам");
                }

                var settings = new System.Xml.XmlWriterSettings
                {
                    Indent = true,
                    Encoding = System.Text.Encoding.UTF8,
                    IndentChars = "  "
                };

                using (System.Xml.XmlWriter writer = System.Xml.XmlWriter.Create(filePath, settings))
                {
                    writer.WriteStartDocument();
                    writer.WriteStartElement("TestProgram");

                    int id = 1;
                    foreach (var person in people)
                    {
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

                    writer.WriteEndElement();
                    writer.WriteEndDocument();
                }
            });
        }
    }
}