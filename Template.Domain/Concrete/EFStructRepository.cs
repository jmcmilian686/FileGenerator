using FileGenerator.Domain.Abstract;
using FileGenerator.Domain.Entities;
using System.Linq;

namespace FileGenerator.Domain.Concrete
{
    public class EFStructRepository : IStructRepository
    {
        //Entity Framework context
        private EFDbContext context = new EFDbContext();

        // Get Method Implementation
        public IQueryable<Struct> Structs
        {

            get { return context.Structs; }
        }

        // Save Method Implementation
        public void SaveStruct(Struct struct_f)
        {
            if (struct_f.ID == 0)
            {
                context.Structs.Add(struct_f);
            }
            else {
                Struct fileStDb = context.Structs.Find(struct_f.ID);
                if (fileStDb != null) {
                    fileStDb.St_Name = struct_f.St_Name;
                    fileStDb.St_Description = struct_f.St_Description;
                    fileStDb.LFile_ID = struct_f.LFile_ID;
                    fileStDb.Multiple = struct_f.Multiple;
                }
            }
            context.SaveChanges();
        }

        // Delete Method Implementation
        public Struct DeleteStruct(int id) {
            Struct fileStDb = context.Structs.Find(id);
            if (fileStDb != null)
            {
                context.Structs.Remove(fileStDb);
            }

            context.SaveChanges();
            return fileStDb;

            
        }

    }
}
