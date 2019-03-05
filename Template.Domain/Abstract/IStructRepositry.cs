using System.Linq;
using FileGenerator.Domain.Entities;

namespace FileGenerator.Domain.Abstract
{
    public interface IStructRepository
    {
        IQueryable<Struct> Structs { get; }

        void SaveStruct(Struct fileStruct);

        Struct DeleteStruct(int id);


    }
}