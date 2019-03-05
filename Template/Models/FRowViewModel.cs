using FileGenerator.Domain.Entities;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace FileGenerator.Models
{
    public class FRowViewModel
    {
        public Dictionary<string,string> DataShow { get; set; }
        public int? Level { get; set; }
     }
}