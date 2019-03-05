using System.Data.Entity;
using FileGenerator.Domain.Entities;

namespace FileGenerator.Domain.Concrete
{
    public class EFDbContext : DbContext
    {
        public DbSet<Product> Products { get; set; }
        public DbSet<Struct> Structs { get; set; }
        public DbSet<LFile> LFiles { get; set; }
        public DbSet<Field> Fields { get; set; }
        public DbSet<StructField> StructFields { get; set; }
        public DbSet<DataField> DataFields { get; set; }

    }
}