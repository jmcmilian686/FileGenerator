using FileGenerator.Domain.Entities;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace FileGenerator.Models
{
    public class StructPosViewModel
    {
        public int ID { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
        public int Order { get; set; }
    }
}