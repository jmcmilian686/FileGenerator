using System.Linq;
using FileGenerator.Domain.Entities;

namespace FileGenerator.Domain.Abstract
{
    public interface IStructFieldRepository
    {
        IQueryable<StructField> StructFields { get; }

        void SaveStructField(StructField structField);

        StructField DeleteStructField(int id);


    }
}