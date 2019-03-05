using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace FileGenerator.Domain.Entities
{
    public  class Struct
    {
        [Key]
        public int ID { get; set; }

        [Required]
        [Display(Name = "Structure Name")]
        public string St_Name { get; set; }

        [Display(Name = "Structure Description")]
        public string St_Description { get; set; }
        
        [Display(Name = "Order in Document")]
        public int? Order_In_Doc { get; set; }

        [Display(Name = "Multiple?")]
        public bool Multiple { get; set; }

        [Display(Name = "File")]
        public int? LFile_ID { get; set; }
        public virtual LFile LFile { get; set; }

        ICollection<StructField> StructField { get; set; }
    }
}
