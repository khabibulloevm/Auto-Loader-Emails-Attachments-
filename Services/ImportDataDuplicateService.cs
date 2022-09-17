using System;
using System.Threading.Tasks;
using EmailReseiver.Contexts;
using EmailReseiver.Models;
using Microsoft.EntityFrameworkCore;

namespace EmailReseiver.Services
{
    public class ImportDataDuplicateService
    {
        public ImportDataDuplicateService(Context context)
        {
            _context = context;
        }
        public async Task<ImportDataDuplicate?> AddEntry(ImportDataDuplicate entry)
        {
            entry.InsertDate = DateTime.Now;
            await _context.ImportDataDuplicate.AddAsync(entry);
            await _context.SaveChangesAsync();
            return await FindItem(entry.Id);
        }

        public Task<ImportDataDuplicate?> FindItem(Int64 id) =>
            _context.ImportDataDuplicate.AsNoTracking()
                .FirstOrDefaultAsync(i => i.Id == id);

        private readonly Context _context;
    }
}
