using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;


namespace FileGenerator.Domain.Entities
{
    public class LFile
    {
        [Key]
        public int LFile_ID { get; set; }

        [Required]
        [Display(Name = "Document Name")]
        public string Doc_Name { get; set; }

        [Display(Name = "Document Description")]
        public string Doc_Description { get; set; }

        [Display (Name ="Number of Structures")]
        public int? Num_Struct { get; set; }

        ICollection<Struct> Struct { get; set; }

    }
}
