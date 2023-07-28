using ImportData.Models;
using Microsoft.EntityFrameworkCore;

namespace ImportData.Data
{
    public class CountriesAPIDbContext : DbContext
    {
        public CountriesAPIDbContext(DbContextOptions options) : base(options)
        {
        }

        public DbSet<Countries> Countries { get; set; }
    }
}
