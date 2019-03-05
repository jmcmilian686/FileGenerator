using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace FileGenerator.Domain.Entities
{
    public class Field
    {
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int ID { get; set; }

        [Required]
        [Display(Name = "Field Name")]
        public string Field_Name { get; set; }

        [DataType(DataType.MultilineText)]
        [Display(Name = "Field Description")]
        public string Field_Desc { get; set; }

        [Display(Name = "Field Type")]
        public string Field_Type { get; set; }

        [Display(Name = "Field Length")]
        [Range(1, int.MaxValue, ErrorMessage = "Please enter a value bigger than 0")]
        public decimal? Length { get; set; }

        [Display(Name = "Default")]
        public string DefaultV { get; set; }

        [Display(Name = "Field Hierarchy")]
        public int FLevel { get; set; }

        [Display(Name = "Unique?")]
        public bool UniqueV { get; set; }

        ICollection<DataField> DaatField { get; set; }
        ICollection<StructField> StructField { get; set; }
       
        [NotMapped]
        public List<object> Typ {
           
            get
            {
                List<object> types = new List<object>
                {
                    new{ value = "String",text ="String"},
                    new{ value = "Number",text ="Number"},
                    new{ value = "Float",text ="Float"}

                };
                return types;
                }
        }

        

    }
}
