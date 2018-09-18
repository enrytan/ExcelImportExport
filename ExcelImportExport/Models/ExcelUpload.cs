using LinqToExcel.Attributes;
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
        [ExcelColumn("Student ID")] //maps the "StudentId" property to the "Student ID" column
        public int StudentId { get; set; }
        [ExcelColumn("Student Name")] //maps the "StudentName" property to the "Student Name" column
        public string StudentName { get; set; }
    }
}