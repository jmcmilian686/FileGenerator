using FileGenerator.Domain.Entities;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace FileGenerator.Models
{
    public class Elements
    {
        public string Name { get; set; }

        public string Val { get; set; }

        public string Type { get; set; }

        public string Percent { get; set; }
    }
}