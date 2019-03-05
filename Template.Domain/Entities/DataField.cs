using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace FileGenerator.Domain.Entities
{
    public class DataField
    {
        [Key]
        public int ID { get; set; }

       
        public string Data { get; set; }

        public int? FieldID { get; set; }

        public int? Link_P { get; set; }

        public int? Link_S { get; set; }

        public virtual Field Field { get; set; }

    }
}
