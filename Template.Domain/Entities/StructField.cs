using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace FileGenerator.Domain.Entities
{
    public class StructField
    {
        [Key]
        public int ID { get; set; }

        [Required]
        [Display(Name = "Field Order")]
        public int Field_Order { get; set; }

        [Display(Name = "Is Required?")]
        public bool Required { get; set; }

        [Display(Name = "Struct")]
        public int StructID { get; set; }
        public virtual Struct Struct { get; set; }

        [Display(Name = "Field")]
        public int FieldID { get; set; }
        public virtual Field Field { get; set; }


    }
}
