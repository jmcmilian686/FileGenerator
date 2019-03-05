using FileGenerator.Domain.Abstract;
using FileGenerator.Domain.Entities;
using System.Linq;

namespace FileGenerator.Domain.Concrete
{
    public class EFLFilesRepository : ILFileRepository
    {
        //Entity Framework context
        private EFDbContext context = new EFDbContext();

        // Get Method Implementation
        public IQueryable<LFile> LFiles
        {

            get { return context.LFiles; }
        }

        // Save Method Implementation
        public void SaveLFile(LFile lfile)
        {
            if (lfile.LFile_ID == 0)
            {
                context.LFiles.Add(lfile);

                

            }
            else
            {
                LFile lfileDb = context.LFiles.Find(lfile.LFile_ID);
                if (lfile != null)
                {
                    lfileDb.Doc_Name = lfile.Doc_Name;
                    lfileDb.Doc_Description = lfile.Doc_Description;
                    lfileDb.Num_Struct = lfile.Num_Struct;
                    
                  



                }
            }
            context.SaveChanges();
        }

        // Delete Method Implementation
        public LFile DeleteLFile(int id)
        {
            LFile LFileDb = context.LFiles.Find(id);
            if (LFileDb != null)
            {
                context.LFiles.Remove(LFileDb);
            }

            context.SaveChanges();

            return LFileDb;


        }
    }
}
