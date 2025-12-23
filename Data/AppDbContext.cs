using CsvDataProcess.Domain;
using Microsoft.EntityFrameworkCore;

namespace CsvDataProcess.Data
{
    public class AppDbContext : DbContext
    {
        public DbSet<Person> People { get; set; }

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            optionsBuilder.UseSqlServer(@"Server=(localdb)\mssqllocaldb;Database=CsvAppDatabase;Trusted_Connection=true;");
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
            Database.EnsureCreated();
        }
    }
}