using CsvDataProcessor.Data;
using CsvDataProcessor.Domain;

namespace CsvDataProcessor.Services.Database
{
    public class PersonService
    {
        private readonly AppDbContext _context;

        public PersonService(AppDbContext context)
        {
            _context = context;
        }

        public IQueryable<Person> GetFilteredPeople(bool useStartDate, DateTime startDate, bool useEndDate, DateTime endDate, string firstName, string lastName, string surName, string city, string country)
        {
            var query = _context.People.AsQueryable();

            if (useStartDate)
            {
                query = query.Where(p => p.Date >= startDate.Date);
            }

            if (useEndDate)
            {
                query = query.Where(p => p.Date <= endDate.Date);
            }

            if (!string.IsNullOrWhiteSpace(firstName))
            {
                string cleanFirstName = firstName.Trim();
                query = query.Where(p => p.FirstName.Contains(cleanFirstName));
            }

            if (!string.IsNullOrWhiteSpace(lastName))
            {
                string cleanLastName = lastName.Trim();
                query = query.Where(p => p.LastName.Contains(cleanLastName));
            }

            if (!string.IsNullOrWhiteSpace(surName))
            {
                string cleanSurName = surName.Trim();
                query = query.Where(p => p.SurName.Contains(cleanSurName));
            }

            if (!string.IsNullOrWhiteSpace(city))
            {
                string cleanCity = city.Trim();
                query = query.Where(p => p.City.Contains(cleanCity));
            }

            if (!string.IsNullOrWhiteSpace(country))
            {
                string cleanCountry = country.Trim();
                query = query.Where(p => p.Country.Contains(cleanCountry));
            }

            return query.OrderBy(p => p.LastName).ThenBy(p => p.FirstName);
        }

        public int GetTotalCount()
        {
            return _context.People.Count();
        }

        public void ClearAllData()
        {
            _context.People.RemoveRange(_context.People);
            _context.SaveChanges();
        }
    }
}