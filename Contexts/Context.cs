using EmailReseiver.Models;
using Microsoft.EntityFrameworkCore;
namespace EmailReseiver.Contexts
{
    public class Context: DbContext
    {
        public Context(DbContextOptions<Context> options) : base(options) { }
        public DbSet<ImportData> ImportData { get; set; }
        public DbSet<ImportDataDuplicate> ImportDataDuplicate { get; set; }

    }
}
