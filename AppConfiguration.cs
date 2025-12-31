namespace CsvDataProcessor.Configuration
{
    public static class AppConfiguration
    {
        public static string GetDefaultConnectionString()
        {
            return "Server=(localdb)\\mssqllocaldb;Database=CsvDataProcessorDb;Trusted_Connection=true;";
        }
    }
}