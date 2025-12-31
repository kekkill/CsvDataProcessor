using CsvDataProcessor.Domain;
using CsvDataProcessor.Configuration;
using Microsoft.EntityFrameworkCore;

namespace CsvDataProcessor.Data
{
    public class AppDbContext : DbContext
    {
        public DbSet<Person> People { get; set; }

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            string connectionString = AppConfiguration.GetDefaultConnectionString();
            optionsBuilder.UseSqlServer(connectionString);
        }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<Person>().HasIndex(p => new { p.LastName, p.FirstName });
            modelBuilder.Entity<Person>().HasIndex(p => p.City);
            modelBuilder.Entity<Person>().HasIndex(p => p.Country);
            modelBuilder.Entity<Person>().HasIndex(p => p.Date);
        }

        public void InitializeDatabase()
        {
            if (!Database.CanConnect())
            {
                Database.EnsureCreated();
            }
            else
            {
                Database.Migrate();
            }
        }
    }
}