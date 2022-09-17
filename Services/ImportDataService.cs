using System;
using System.Threading.Tasks;
using EmailReseiver.Contexts;
using EmailReseiver.Models;
using Microsoft.EntityFrameworkCore;

namespace EmailReseiver.MailServices
{
    public class ImportDataService
    {
        public ImportDataService(Context context)
        {
            _context = context;
        }
        public async Task<ImportData?> AddEntry(ImportData entry)
        {
            entry.InsertDate = DateTime.Now;
            await _context.AddAsync(entry);
            await _context.SaveChangesAsync();
            return await FindItem(entry.Id);
        }
 
        public Task<ImportData?> FindItem(Int64 id) => 
            _context.ImportData.AsNoTracking()
                .FirstOrDefaultAsync(i => i.Id == id);

        public Task<ImportData?> FindItemByRecNum(string recNum) =>
            _context.ImportData.AsNoTracking()
                .FirstOrDefaultAsync(i => i.RecNum == recNum);

        public async Task<bool> IsRecNumExistAsync(string recNum)
        {
            var importData = await FindItemByRecNum(recNum);

            if (importData != null)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private readonly Context _context;
    }
}
