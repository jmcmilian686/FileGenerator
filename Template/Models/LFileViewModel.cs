using FileGenerator.Domain.Entities;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace FileGenerator.Models
{
    public class LFileViewModel
    {
        public LFile Document { get; set; }

        [Display(Name = "Document Structures")]
        public AddRowViewModel DocStructs { get; set; }
    }
}