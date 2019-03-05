using FileGenerator.Domain.Abstract;
using FileGenerator.Domain.Entities;
using System.Linq;

namespace FileGenerator.Domain.Concrete
{
    public class EFDataFieldRepository : IDataFieldRepository
    {
        //Entity Framework context
        private EFDbContext context = new EFDbContext();

        // Get Method Implementation
        public IQueryable<DataField> DataFields
        {

            get { return context.DataFields; }
        }

        // Save Method Implementation
        public void SaveDataField(DataField datafield)
        {
            if (datafield.ID == 0)
            {
                context.DataFields.Add(datafield);
            }
            else
            {
                DataField datafieldDb = context.DataFields.Find(datafield.ID);
                if (datafieldDb != null)
                {
                    datafieldDb.Data = datafield.Data;
                    datafieldDb.FieldID = datafield.FieldID;
                    datafieldDb.Link_P = datafield.Link_P;
                    datafieldDb.Link_S = datafield.Link_S;
                 }
            }
            context.SaveChanges();
        }

        // Delete Method Implementation
        public DataField DeleteDataField(int id)
        {
            DataField datafieldDb = context.DataFields.Find(id);
            if (datafieldDb != null)
            {
                context.DataFields.Remove(datafieldDb);
            }

            context.SaveChanges();

            return datafieldDb;


        }
    }
}
