using System.Linq;
using FileGenerator.Domain.Entities;

namespace FileGenerator.Domain.Abstract
{
    public interface IFieldsRepository
    {
        IQueryable<Field> Fields { get; }

        void SaveField(Field field);

        Field DeleteField(int id);


    }
}