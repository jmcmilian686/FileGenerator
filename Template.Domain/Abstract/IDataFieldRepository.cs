using System.Linq;
using FileGenerator.Domain.Entities;

namespace FileGenerator.Domain.Abstract
{
    public interface IDataFieldRepository
    {
        IQueryable<DataField> DataFields { get; }

        void SaveDataField(DataField datafield);

        DataField DeleteDataField(int id);


    }
}