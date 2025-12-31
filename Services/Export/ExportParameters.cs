namespace CsvDataProcessor.Services.Export
{
    public class ExportParameters
    {
        public bool UseStartDate { get; set; }
        public DateTime StartDate { get; set; }
        public bool UseEndDate { get; set; }
        public DateTime EndDate { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string SurName { get; set; }
        public string City { get; set; }
        public string Country { get; set; }
    }
}