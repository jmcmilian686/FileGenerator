using FileGenerator.Domain.Entities;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace FileGenerator.Models
{
    public class FormCreateViewModel
    {
        public int Organization { get; set; }
        public int Dist_Center { get; set; }
        public int Bussiness_Unit { get; set; }
        public int Order_Type { get; set; }
        public int Document { get; set; }
    }
}