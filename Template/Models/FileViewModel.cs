using FileGenerator.Domain.Entities;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace FileGenerator.Models
{
    public class FileViewModel
    {
        public List<FRowViewModel> Rows { get; set; }
        public int Amount { get; set; }
    }
}