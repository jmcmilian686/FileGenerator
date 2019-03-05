using System.Linq;
using FileGenerator.Domain.Entities;

namespace FileGenerator.Domain.Abstract
{
    public interface IProductRepository
    {
        IQueryable<Product> Products { get; }
    }
}