using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace ExcelImportExport.Models
{
    [Table("Student")]
    public class ExcelUpload
    {
        [Key]
        public int StudentId { get; set; }
        public string StudenteName { get; set; }
    }
}