using FileGenerator.Domain.Abstract;
using FileGenerator.Domain.Entities;
using System.Linq;

namespace FileGenerator.Domain.Concrete
{
    public class EFStructFIeldRepository : IStructFieldRepository
    {
        //Entity Framework context
        private EFDbContext context = new EFDbContext();

        // Get Method Implementation
        public IQueryable<StructField> StructFields
        {

            get { return context.StructFields; }
        }

        // Save Method Implementation
        public void SaveStructField(StructField structField)
        {
            if (structField.ID == 0)
            {
                context.StructFields.Add(structField);
            }
            else
            {
                StructField structFieldDb = context.StructFields.Find(structField.ID);
                if (structField != null)
                {
                    structFieldDb.Field_Order = structField.Field_Order;
                    structFieldDb.Required = structField.Required;
                    structFieldDb.StructID = structField.StructID;
                    structFieldDb.FieldID = structField.FieldID;




                }
            }
            context.SaveChanges();
        }

        // Delete Method Implementation
        public StructField DeleteStructField(int id)
        {
            StructField structFieldDb = context.StructFields.Find(id);
            if (structFieldDb != null)
            {
                context.StructFields.Remove(structFieldDb);
            }

            context.SaveChanges();

            return structFieldDb;


        }
    }
}
