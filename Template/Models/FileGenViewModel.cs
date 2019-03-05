using FileGenerator.Domain.Entities;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using static Template.Controllers.DatafieldController;

namespace FileGenerator.Models
{


    public class FileGenViewModel
    {
        public int DocID { get; set; }
        public int NDocs { get; set; }
        public int NDets { get; set; }
        public int NBatch { get; set; }
        public string FileName { get; set; }
        public bool Max { get; set; }
        public List<Elements> FiltValues { get; set; }
    }
}