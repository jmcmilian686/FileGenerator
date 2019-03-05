using FileGenerator.Domain.Entities;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using static Template.Controllers.DatafieldController;

namespace FileGenerator.Models
{
    

    public class TableViewModel
    {
       
        public int ID { get; set; }
      
        public List<Elements> DataValue { get; set; }
    }
}