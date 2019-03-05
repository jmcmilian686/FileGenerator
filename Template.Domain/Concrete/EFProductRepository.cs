using System.Linq;
using FileGenerator.Domain.Abstract;
using FileGenerator.Domain.Concrete;
using FileGenerator.Domain.Entities;

namespace FileGenerator.Domain.Concrete
{
    public class EFProductRepository : IProductRepository
    {
        private EFDbContext context = new EFDbContext();
        public IQueryable<Product> Products
        {
            get { return context.Products; }
        }
    }
}