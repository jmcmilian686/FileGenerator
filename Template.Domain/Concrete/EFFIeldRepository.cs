using FileGenerator.Domain.Abstract;
using FileGenerator.Domain.Entities;
using System.Linq;

namespace FileGenerator.Domain.Concrete
{
    public class EFFIeldRepository : IFieldsRepository
    {
        //Entity Framework context
        private EFDbContext context = new EFDbContext();

        // Get Method Implementation
        public IQueryable<Field> Fields
        {

            get { return context.Fields; }
        }

        // Save Method Implementation
        public void SaveField(Field field)
        {
            if (field.ID == 0)
            {
                context.Fields.Add(field);
            }
            else
            {
                Field fieldDb = context.Fields.Find(field.ID);
                if (fieldDb != null)
                {
                    fieldDb.Field_Name = field.Field_Name;
                    fieldDb.Field_Type = field.Field_Type;
                    fieldDb.Length = field.Length;
                    fieldDb.DefaultV = field.DefaultV;
                    fieldDb.Field_Desc = field.Field_Desc;
                    fieldDb.UniqueV = field.UniqueV;
                    fieldDb.FLevel = field.FLevel;
                   
                    

                }
            }
            context.SaveChanges();
        }

        // Delete Method Implementation
        public Field DeleteField(int id)
        {
            Field fieldDb = context.Fields.Find(id);
            if (fieldDb != null)
            {
                context.Fields.Remove(fieldDb);
            }

            context.SaveChanges();

            return fieldDb;


        }
    }
}
