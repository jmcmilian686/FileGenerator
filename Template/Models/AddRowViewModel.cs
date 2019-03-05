using FileGenerator.Domain.Entities;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace FileGenerator.Models
{
    public class AddRowViewModel
    {
        public IEnumerable<StructPosViewModel> StructurePos { get; set; }
        public int? Posit { get; set; }
        public int? NewStId { get; set; }
    }
}