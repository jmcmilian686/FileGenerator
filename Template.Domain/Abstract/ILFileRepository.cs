using System.Linq;
using FileGenerator.Domain.Entities;

namespace FileGenerator.Domain.Abstract
{
    public interface ILFileRepository
    {
        IQueryable<LFile> LFiles { get; }

        void SaveLFile(LFile lfiles);

        LFile DeleteLFile(int id);


    }
}